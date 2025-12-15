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
CSV_FILE = "test.csv"
LOG_FILE = "charge_cell_log.xlsx"
COM_PORT = "COM6"
MIN_VOLT = 0.0
MAX_VOLT = 4.3


# EXCEL FILE SETUP
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Log"
ws.append(["Index", "Time (s)", "Voltage (V)", "Current setpoint (A)", "Current (A)", "Ah", "Wh", "Weight (kg)"])


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
    i = abs(float(inst.query("FETC:SCAL:CURR?")))
    Ah = float(inst.query("FETC:AHO?"))
    Wh = float(inst.query("FETC:WHO?"))
    
    sample = task.read(1, 0.01)
    sample = [s[0] for s in sample]
    
    voltage = float(sample[0])
    kg = voltage * 500000
    temp1 = float(sample[1])
    temp2 = float(sample[2])
    temp3 = float(sample[3])
    temp4 = float(sample[4])
    temp5 = float(sample[5])
    pressure_kg = voltage * 500000
    
    print(f"\r#{row_index} | t={t:.1f}s | V={v:.2f} V | I={i:.3f} A | Ah={Ah:.3f} | Wh={Wh:.3f} | kg={pressure_kg:.3f} | temp1={temp1:.3f} | temp2={temp2:.3f}", end="") #DA AGGIUNGERE DATI
    return [row_index, t, v, setpoint_current, i, Ah, Wh, pressure_kg] #DA AGGIUNGERE DATI

async def v_integrity_check(stop_trigger):
    while not stop_trigger.is_set():
        v = float(inst.query("FETC:SCAL:VOLT?"))
        if v >= MAX_VOLT or v <= MIN_VOLT:
            print("\nVoltage limit reached (V = {v:.2f}), stopping step early.")
            stop_trigger.set()

        await asyncio.sleep(0.01)

async def setpoint_handler(stop_trigger, pause_trigger, setpoint_current):
    csv_setpoints = read_csv()
    while not stop_trigger.is_set():
        if not pause_trigger.is_set():
            await pause_trigger.wait()   # wait until restart

        for line in csv_setpoints:
            async with setpoint_current.lock:
                setpoint_current.value = int(line[1])
                setpoint_time = float(line[0])

                print(f"\nApplying I = {setpoint_current.value} A for {setpoint_time} seconds...")
                set_current(setpoint_current.value)

            await asyncio.sleep(setpoint_time)
        
        stop_trigger.set()

async def logger(stop_trigger, pause_trigger, task, setpoint_current):
    row_index = 1
    save_counter = 0
    start_time = time.time()

    while not stop_trigger.is_set():
        if not pause_trigger.is_set():
            await pause_trigger.wait()   # wait until restart

        async with setpoint_current.lock:
            data = read_data(start_time, row_index, task, setpoint_current.value)

            # Write Excel
            ws.append(data)
            row_index += 1

            # Save periodically
            save_counter += 1
            if save_counter >= 200:
                wb.save(LOG_FILE)
                save_counter = 0

        await asyncio.sleep(0.1)

async def user_input(stop_trigger, pause_trigger):
    while not stop_trigger.is_set():
        cmd = await asyncio.to_thread(input, "Comando (p=pause, r=resume, q=quit): ")
        if cmd == "p":
            pause_trigger.clear()
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
        
        canale1 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai0", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.J)
        canale1.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale2 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai1", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.J)
        canale2.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale3 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai2", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.J)
        canale3.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale4 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai3", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.J)
        canale4.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        canale5 = task.ai_channels.add_ai_thrmcpl_chan("cDAQ1Mod2/ai4", min_val = 0.0, max_val = 100.0, units = TemperatureUnits.DEG_C, thermocouple_type = ThermocoupleType.J)
        canale5.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION
        
        task.timing.cfg_samp_clk_timing(10.0, sample_mode=AcquisitionType.CONTINUOUS)
        task.start()

        tasks = [
            asyncio.create_task(user_input(stop_trigger, pause_trigger)),
            asyncio.create_task(v_integrity_check(stop_trigger)),
            asyncio.create_task(setpoint_handler(stop_trigger, pause_trigger, setpoint_current)),
            asyncio.create_task(logger(stop_trigger, pause_trigger, task, setpoint_current)),
        ]
    
        # ASYNC BLOCK 
        try:
            await stop_trigger.wait()     # Waits until stop event is triggered  

        except KeyboardInterrupt:
            print("\n*** Manual STOP ***")
            stop_trigger.set()        
        finally:
            # Task deletion procedure
            for t in tasks:
                t.cancel()           
            await asyncio.gather(*tasks, return_exceptions=True)
            inst.write("OUTP 0")
            inst.close()
            wb.save(LOG_FILE)        
            print(f"\nSaved to {LOG_FILE}")
            print("\n=== DONE ===")

    print("System stopped safely")


# MAIN 
if __name__ == "__main__":
    asyncio.run(main())