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
LOG_FILE = "test_IR_carica_log.xlsx" #RICORDARSI DI CAMBIARE NOME FILE
COM_PORT = "COM6"
MIN_VOLT = 3.0
MAX_VOLT = 4.3
CURRENT  = 17.8
STOP_CURRENT = 0.89

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
inst.source.current(CURRENT)
inst.write("VOLT:LIM " + str(MAX_VOLT))
inst.write("VOLT:LIM:NEG " + str(MIN_VOLT))

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
    
if __name__ == "__main__":
    
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
        
        print("\n=== STARTING ==")  
        inst.write("OUTP 1")
        
        row_index = 1
        start_time = time.time()
        
        try:
            
            while True:
                
                try:
                    data = read_data(start_time, row_index, task, CURRENT)
                except Exception as e:
                    print(f"\nErrore lettura dati: {e}")
                
                ws.append(data)
                row_index += 1
                    
                time.sleep(0.1)
                if data[4] < STOP_CURRENT:
                    print("\nCiclo completato automaticamente.")
                    break

        except KeyboardInterrupt:
            print("\n*** Manual STOP ***")
            inst.write("OUTP 0")
        except Exception as e:
            print(f"\n*** ERROR OCCURRED: {e} ***")
            inst.write("OUTP 0")
        finally:
            inst.write("OUTP 0")
            wb.save(LOG_FILE)        
            print(f"\nSaved to {LOG_FILE}")
            print("\n=== DONE ===")

    print("System stopped safely")
