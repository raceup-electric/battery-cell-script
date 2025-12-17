import openpyxl
import time
import easy_scpi as scpi
import csv
import nidaqmx
from nidaqmx.constants import BridgeConfiguration, ADCTimingMode, AcquisitionType, TemperatureUnits, ThermocoupleType
import asyncio


# =========================
# SETTINGS
# =========================
CSV_FILE = "test.csv" #RICORDARSI DI CAMBIARE NOME FILE
LOG_FILE = "test_IR_carica_log.xlsx" #RICORDARSI DI CAMBIARE NOME FILE
COM_PORT = "COM6"
MIN_VOLT = 3.0
MAX_VOLT = 4.3


# EXCEL FILE SETUP
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Log"
ws.append(["Index", "Time (s)", "Voltage (V)", "Setpoint (A)", "Current (A)", "Ah", "Wh", "Weight (kg)", "T1 (°C)", "T2 (°C)", "T3 (°C)", "T4 (°C)", "T5 (°C)", "Tamb (°C)", "T7 (°C)", "T8(°C)"])


# POWER SUPPLY
inst = scpi.Instrument(COM_PORT) #modificare con il com giusto
inst.connect()
inst.write("SYST:CLE")
inst.write("FUNC:MODE FIX")
inst.write("SYST:COMM:SER:BAUD 115200")

inst.write("REM:SENS 1") #mettere 1 se si usano i pin di sensing per la tensione
inst.write("INIT:ACQ")
inst.write("SENS:AHO:RES")
inst.write("SENS:WHO:RES")
inst.write("CURR:SLEW MIN")
inst.write("VOLT:SLEW MIN")

inst.write("FUNC CURR")
inst.write("CURR 0")
inst.write("VOLT:LIM " + str(MAX_VOLT))
inst.write("VOLT:LIM:NEG " + str(MIN_VOLT)) 

# DATA CONTAINER
class SharedData:
    def __init__(self):
        self.value = 0.0
        self.lock = asyncio.Lock()

# FUNZIONI
def set_current(current):
    inst.source.current(current)
    inst.write("OUTP 1")

def read_csv():
    data = []
    try:
        with open(CSV_FILE, newline='') as f:
            reader = csv.reader(f, delimiter=',')
            for line_num, line in enumerate(reader, start=1):
                try:
                    data.append((float(line[0]), float(line[1])))
                except (ValueError, IndexError) as e:
                    print(f"Errore alla riga {line_num}: {e}")
    except FileNotFoundError:
        print(f"File {CSV_FILE} non trovato.")
    except Exception as e:
        print(f"Errore imprevisto durante la lettura: {e}")

    return data

def read_data(start_time, row_index, task, setpoint_current):
    t = time.time() - start_time
    
    v = float(inst.query("FETC:SCAL:VOLT?"))
    i = float(inst.query("FETC:SCAL:CURR?"))
    Ah = float(inst.query("FETC:AHO?"))
    Wh = float(inst.query("FETC:WHO?"))
    
    sample = task.read(1, 0.01)
    sample = [s[0] for s in sample]
    
    voltage = float(sample[0])
    pressure_kg = voltage * 500000
    temp1 = float(sample[1])
    temp2 = float(sample[2])
    temp3 = float(sample[3])
    temp4 = float(sample[4])
    temp5 = float(sample[5])
    temp6 = float(sample[6])
    temp7 = float(sample[7])
    temp8 = float(sample[8])
    
    print(f"\r#{row_index} | t={t:.1f}s | V={v:.3f} V | I={i:.3f} A | Ah={Ah:.2f} | Wh={Wh:.2f} | kg={pressure_kg:.0f} | t1: {temp1:.2f} | t2: {temp2:.2f} | t3: {temp3:.2f} | t4: {temp4:.2f} | t5: {temp5:.2f} | Tamb: {temp6:.2f} | t7: {temp7:.2f} | t8: {temp8:.2f}", end="")
    return [row_index, t, v, setpoint_current, i, Ah, Wh, pressure_kg, temp1, temp2, temp3, temp4, temp5, temp6, temp7, temp8]

async def v_integrity_check(stop_trigger):
    while not stop_trigger.is_set():
        v = float(inst.query("FETC:SCAL:VOLT?"))
        if v > MAX_VOLT or v < MIN_VOLT:
            print(f"\nVoltage limit reached (V = {v:.2f}), stopping step early.", end="")
            stop_trigger.set()

        await asyncio.sleep(0.01)

async def setpoint_handler(stop_trigger, pause_trigger, setpoint_current):
    csv_setpoints = read_csv()
    while not stop_trigger.is_set():

        for line in csv_setpoints:
            if not pause_trigger.is_set():
                await pause_trigger.wait()   # wait until restart

            async with setpoint_current.lock:
                setpoint_time = float(line[0])
                setpoint_current.value = float(line[1])

                print(f"\nApplying I = {setpoint_current.value} A for {setpoint_time} seconds...")
                set_current(setpoint_current.value)

            await asyncio.sleep(setpoint_time)
        
        stop_trigger.set()

async def logger(stop_trigger, task, setpoint_current):
    row_index = 1
    save_counter = 0
    start_time = time.time()

    while not stop_trigger.is_set():

        async with setpoint_current.lock:
            
            try:
                data = read_data(start_time, row_index, task, setpoint_current.value)
            except Exception as e:
                print(f"\nErrore lettura dati: {e}")

            # Write Excel
            ws.append(data)
            row_index += 1
            
        await asyncio.sleep(0.1)

async def user_input(stop_trigger, pause_trigger):
    while not stop_trigger.is_set():
        cmd = await asyncio.to_thread(input, "Comando (p=pause, r=resume, q=quit): ")
        if cmd == "p":
            pause_trigger.clear()
            inst.write("OUTP 0")
            print("Script is having a break...")
        elif cmd == "r":
            pause_trigger.set()
            print("Script is starting again...")
        elif cmd == "q":
            stop_trigger.set()
            print("Script stopped, have a nice day!")

async def main():
    stop_trigger = asyncio.Event()
    pause_trigger = asyncio.Event()
    setpoint_current = SharedData() #shares setpoint among logger and setpoint_handler

    # --- setup sync tasks ---
    print("\n=== STARTING CYCLES ==")  
 
    with nidaqmx.Task() as task:
        
        loadcell = task.ai_channels.add_ai_bridge_chan("cDAQ1Mod1/ai0", min_val = -0.05, max_val = 0.05, bridge_config = BridgeConfiguration.FULL_BRIDGE, nominal_bridge_resistance=350)
        loadcell.ai_adc_timing_mode = ADCTimingMode.BEST_50_HZ_REJECTION
        
        canale1 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai0", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.K)
        canale1.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale2 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai1", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.K)
        canale2.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale3 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai2", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.K)
        canale3.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale4 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai3", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.K)
        canale4.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale5 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai4", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.K)
        canale5.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale6 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai5", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.J)
        canale6.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION

        canale7 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai6", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.K)
        canale7.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION

        canale8 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai7", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.K)
        canale8.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        task.timing.cfg_samp_clk_timing(10.0, sample_mode=AcquisitionType.CONTINUOUS)
        task.start()

        tasks = [
            asyncio.create_task(user_input(stop_trigger, pause_trigger)),
            asyncio.create_task(v_integrity_check(stop_trigger)),
            asyncio.create_task(setpoint_handler(stop_trigger, pause_trigger, setpoint_current)),
            asyncio.create_task(logger(stop_trigger, task, setpoint_current)),
        ]
    
        # ASYNC BLOCK 
        try:
            await stop_trigger.wait()     # Waits until stop event is triggered  

        except KeyboardInterrupt:
            print("\n*** Manual STOP ***")
            stop_trigger.set()
        except Exception as e:
            print(f"\n*** ERROR OCCURRED: {e} ***")
            wb.save(LOG_FILE)
            inst.write("OUTP 0")
            stop_trigger.set()
        finally:
            # Task deletion procedure
            for t in tasks:
                t.cancel()           
            await asyncio.gather(*tasks, return_exceptions=True)
            inst.write("OUTP 0")
            wb.save(LOG_FILE)        
            print(f"\nSaved to {LOG_FILE}")
            print("\n=== DONE ===")

    print("System stopped safely")


# MAIN 
if __name__ == "__main__":
    asyncio.run(main())
