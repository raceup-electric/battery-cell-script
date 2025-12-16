import nidaqmx
from nidaqmx.constants import BridgeConfiguration, ADCTimingMode, AcquisitionType, TemperatureUnits, ThermocoupleType
import numpy as np
import time

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
    
    
    task.timing.cfg_samp_clk_timing(10.0, sample_mode=AcquisitionType.CONTINUOUS)
    task.start()
        
    try:
        while True:
            sample = task.read(1)
            sample = [s[0] for s in sample]
            voltage = float(sample[0])
            kg = voltage * 500000
            temp1 = float(sample[1])
            temp2 = float(sample[2])
            temp3 = float(sample[3])
            temp4 = float(sample[4])
            temp5 = float(sample[5])
            temp6 = float(sample[6])
            print(f"{kg:.2f} kg | temp1: {temp1:.2f} °C | temp2: {temp2:.2f} °C | temp3: {temp3:.2f} °C | temp4: {temp4:.2f} °C | temp5: {temp5:.2f} °C | temp6: {temp6:.2f} °C")
            
    except KeyboardInterrupt:
        print("Acquisition stopped by user.")
