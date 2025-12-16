# Battery Cell Test Script

## Overview
This project provides an **async Python script** to automate battery cell testing.  
It orchestrates power supply control, voltage integrity checks, setpoint handling, and data logging, while allowing interactive user commands (quit, pause, resume).

---

## External Libraries
- **openpyxl** → Read/write Excel files for configuration or results.  
- **time** → Timing utilities for delays and scheduling.  
- **easy_scpi** → SCPI communication with the programmable power supply (*alimentatore*).  
- **csv** → Read test setpoints and profiles from CSV files.  
- **nidaqmx** → National Instruments DAQ interface for sensor acquisition.  
  - Includes constants for bridge configuration, ADC timing, acquisition type, temperature units, and thermocouple type.  
- **asyncio** → Core async framework for concurrent tasks.

---

## Script Structure

### Async Tasks
The script launches four main tasks:
1. **`user_input(stop_trigger, pause_trigger)`**  
   - Runs in a background thread.  
   - Collects user commands: `quit`, `pause`, `resume`.  
   - Pauses/stops the power supply while continuing logging.

2. **`v_integrity_check(stop_trigger)`**  
   - Async task executed frequently.  
   - Monitors voltage acquisition for reliability.  
   - Stops the test if values go out of bounds.

3. **`setpoint_handler(stop_trigger, pause_trigger, setpoint_current)`**  
   - Async task triggered by CSV-defined states.  
   - Updates current setpoints on the power supply.  
   - Ensures reproducible test profiles.

4. **`logger(stop_trigger, pause_trigger, task, setpoint_current)`**  
   - Async acquisition task (~10 Hz).  
   - Logs voltage, current, and load cell data.  
   - Provides continuous monitoring during test execution.

---

## Workflow
1. Load test parameters from CSV.  
2. Configure power supply via SCPI.  
3. Start async tasks: user input, integrity check, setpoint handler, logger.  
4. Monitor acquisition and log results.  
5. Stop or pause safely via user commands.
