import openpyxl
import time
import easy_scpi as scpi
import csv
import nidaqmx
import os
import sys
from nidaqmx.constants import (
    BridgeConfiguration, ADCTimingMode, AcquisitionType,
    TemperatureUnits, ThermocoupleType, READ_ALL_AVAILABLE
)
import asyncio


# =========================
# SETTINGS
# =========================
CSV_FILE  = "GITT_6C_pt2.csv"          # RICORDARSI DI CAMBIARE NOME FILE (FILE DATI ENTRATA)
LOG_FILE  = "log_scarica_storage.xlsx"  # RICORDARSI DI CAMBIARE NOME FILE (FILE DATI USCITA)
COM_PORT  = "COM3"
MIN_VOLT  = 2.75
MAX_VOLT  = 4.45

RISE_TIME = 0.01
FALL_TIME = 0.01

# =========================
# FILE OVERWRITE PROTECTION
# =========================
def check_log_file(path: str) -> None:
    """
    Se il file di log esiste già, chiede conferma all'utente prima di procedere.
    In caso di risposta negativa termina lo script in modo pulito.
    """
    if os.path.exists(path):
        print(f"\n⚠️  ATTENZIONE: il file '{path}' esiste già e verrebbe sovrascritto!")
        while True:
            answer = input("Vuoi continuare e sovrascriverlo? [s/n]: ").strip().lower()
            if answer in ("s", "si", "sì", "y", "yes"):
                print("Ok, il file verrà sovrascritto.\n")
                break
            elif answer in ("n", "no"):
                print("Script interrotto. Cambia LOG_FILE e riavvia.")
                sys.exit(0)
            else:
                print("Risposta non valida. Digita 's' per sì oppure 'n' per no.")

check_log_file(LOG_FILE)

# =========================
# EXCEL FILE SETUP
# =========================
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Log"
ws.append([
    "Index", "Time (s)", "Voltage (V)", "Setpoint (A)", "Current (A)",
    "Ah", "Wh", "Weight (kg)",
    "T1 (°C)", "T2 (°C)", "T3 (°C)", "T4 (°C)", "T5 (°C)",
    "Tamb (°C)", "T7 (°C)", "T8 (°C)", "T9 (°C)", "T10 (°C)"
])

# =========================
# POWER SUPPLY INIT
# =========================
inst = scpi.Instrument(COM_PORT)
inst.connect()
inst.write("SYST:CLE")
inst.write("FUNC:MODE FIX")
inst.write("SYST:COMM:SER:BAUD 115200")

inst.write("REM:SENS 1")       # usare i pin di sensing per la tensione
inst.write("INIT:ACQ")
inst.write("SENS:AHO:RES")
inst.write("SENS:WHO:RES")
inst.write("VOLT:SLEW MIN")

inst.write(f"CURR:SLEW {RISE_TIME}")
inst.write(f"CURR:SLEW {FALL_TIME}")

inst.write("FUNC CURR")
inst.write("CURR 0")
inst.write("VOLT:LIM "     + str(MAX_VOLT))
inst.write("VOLT:LIM:NEG " + str(MIN_VOLT))

# =========================
# SHARED DATA
# =========================
class SharedData:
    """Contenitore thread-safe per il setpoint di corrente."""
    def __init__(self):
        self.value = 0.0
        self.lock  = asyncio.Lock()

# =========================
# FUNZIONI SINCRONE
# =========================
def set_current(current: float) -> None:
    inst.source.current(current)
    inst.write("OUTP 1")


def read_csv() -> list:
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


def read_daq(task) -> list:
    """
    Legge TUTTI i campioni disponibili nel buffer DAQmx e restituisce
    l'ultimo per ogni canale, così il buffer non va mai in overflow.
    """
    samples = task.read(
        number_of_samples_per_channel=READ_ALL_AVAILABLE,
        timeout=0.5
    )
    # samples è una lista di liste (un elemento per canale)
    return [s[-1] if isinstance(s, list) else s for s in samples]


def read_data(start_time: float, row_index: int,
              daq_sample: list, setpoint_current: float) -> list:
    """Interroga l'alimentatore e compone la riga di log."""
    t   = time.time() - start_time
    v   = float(inst.query("FETC:SCAL:VOLT?"))
    i   = float(inst.query("FETC:SCAL:CURR?"))
    Ah  = float(inst.query("FETC:AHO?"))
    Wh  = float(inst.query("FETC:WHO?"))

    # daq_sample: [load-cell, T1..T10]
    voltage     = float(daq_sample[0])
    pressure_kg = voltage * 500000
    temp1  = float(daq_sample[1])
    temp2  = float(daq_sample[2])
    temp3  = float(daq_sample[3])
    temp4  = float(daq_sample[4])
    temp5  = float(daq_sample[5])
    temp6  = float(daq_sample[6])
    temp7  = float(daq_sample[7])
    temp8  = float(daq_sample[8])
    temp9  = float(daq_sample[9])
    temp10 = float(daq_sample[10])

    print(
        f"\r#{row_index} | t={t:.1f}s | V={v:.3f} V | I={i:.3f} A | "
        f"Ah={Ah:.2f} | Wh={Wh:.2f} | kg={pressure_kg:.0f} | "
        f"t1:{temp1:.2f} t2:{temp2:.2f} t3:{temp3:.2f} t4:{temp4:.2f} "
        f"t5:{temp5:.2f} Tamb:{temp6:.2f} t7:{temp7:.2f} t8:{temp8:.2f} "
        f"t9:{temp9:.2f} t10:{temp10:.2f}",
        end=""
    )

    return [
        row_index, t, v, setpoint_current, i, Ah, Wh, pressure_kg,
        temp1, temp2, temp3, temp4, temp5, temp6, temp7, temp8, temp9, temp10
    ]

# =========================
# COROUTINE ASINCRONE
# =========================
async def v_integrity_check(stop_trigger: asyncio.Event,
                             pause_trigger: asyncio.Event) -> None:
    """Controlla la tensione ogni 10 ms e ferma il ciclo se fuori limite."""
    while not stop_trigger.is_set():
        try:
            v = float(inst.query("FETC:SCAL:VOLT?"))
            if v > MAX_VOLT or v <= MIN_VOLT:
                print(f"\nTensione fuori limite (V={v:.3f} V), step interrotto.", end="")
                pause_trigger.clear()
        except Exception as e:
            print(f"\nErrore v_integrity_check: {e}")
        await asyncio.sleep(0.01)


async def setpoint_handler(stop_trigger:  asyncio.Event,
                            pause_trigger: asyncio.Event,
                            setpoint_current: SharedData) -> None:
    """Scorre il CSV e applica i setpoint di corrente in sequenza."""
    csv_setpoints = read_csv()

    for line in csv_setpoints:
        if stop_trigger.is_set():
            break

        if not pause_trigger.is_set():
            await pause_trigger.wait()   # aspetta il resume

        async with setpoint_current.lock:
            setpoint_time          = float(line[0])
            setpoint_current.value = float(line[1])
            print(f"\nApplico I = {setpoint_current.value} A per {setpoint_time} s...")
            set_current(setpoint_current.value)

        await asyncio.sleep(setpoint_time)

    stop_trigger.set()


async def logger(stop_trigger:     asyncio.Event,
                 task,
                 setpoint_current: SharedData) -> None:
    """
    Legge il DAQ FUORI dal lock (per non bloccare mai lo svuotamento del buffer),
    poi prende il lock solo per le query SCPI e la scrittura Excel.
    """
    row_index  = 1
    start_time = time.time()

    while not stop_trigger.is_set():

        # --- Lettura DAQ: fuori dal lock, in un thread separato ---
        try:
            daq_sample = await asyncio.to_thread(read_daq, task)
        except Exception as e:
            print(f"\nErrore lettura DAQ: {e}")
            await asyncio.sleep(0.05)
            continue

        # --- Query alimentatore + log: dentro il lock ---
        async with setpoint_current.lock:
            try:
                data = read_data(start_time, row_index, daq_sample, setpoint_current.value)
            except Exception as e:
                print(f"\nErrore lettura dati alimentatore: {e}")
                await asyncio.sleep(0.05)
                continue

            ws.append(data)
            row_index += 1

        await asyncio.sleep(0.1)


async def user_input(stop_trigger:  asyncio.Event,
                     pause_trigger: asyncio.Event) -> None:
    """Gestisce i comandi da tastiera in modo non bloccante."""
    while not stop_trigger.is_set():
        cmd = await asyncio.to_thread(
            input, "Comando (p=pause, r=resume, q=quit): "
        )
        if cmd == "p":
            pause_trigger.clear()
            inst.write("OUTP 0")
            print("Script in pausa...")
        elif cmd == "r":
            pause_trigger.set()
            print("Script ripreso.")
        elif cmd == "q":
            stop_trigger.set()
            print("Script fermato. Arrivederci!")

# =========================
# MAIN
# =========================
async def main() -> None:
    stop_trigger     = asyncio.Event()
    pause_trigger    = asyncio.Event()
    pause_trigger.set()                          # parte non in pausa
    setpoint_current = SharedData()

    print("\n=== AVVIO CICLI ===")

    with nidaqmx.Task() as task:

        # --- Load cell ---
        loadcell = task.ai_channels.add_ai_bridge_chan(
            "cDAQ1Mod1/ai0",
            min_val=-.05, max_val=.05,
            bridge_config=BridgeConfiguration.FULL_BRIDGE,
            nominal_bridge_resistance=350
        )
        loadcell.ai_adc_timing_mode = ADCTimingMode.BEST_50_HZ_REJECTION

        # --- Termocoppie ---
        tc_channels = [
            ("cDAQ1Mod2/ai0",  ThermocoupleType.K),
            ("cDAQ1Mod2/ai1",  ThermocoupleType.K),
            ("cDAQ1Mod2/ai2",  ThermocoupleType.K),
            ("cDAQ1Mod2/ai3",  ThermocoupleType.K),
            ("cDAQ1Mod2/ai4",  ThermocoupleType.K),
            ("cDAQ1Mod2/ai5",  ThermocoupleType.K),
            ("cDAQ1Mod2/ai6",  ThermocoupleType.K),
            ("cDAQ1Mod2/ai7",  ThermocoupleType.J),
            ("cDAQ1Mod2/ai8",  ThermocoupleType.J),
            ("cDAQ1Mod2/ai9",  ThermocoupleType.K),
        ]
        for ch, tc_type in tc_channels:
            ch_obj = task.ai_channels.add_ai_thrmcpl_chan(
                ch,
                min_val=0.0, max_val=100.0,
                units=TemperatureUnits.DEG_C,
                thermocouple_type=tc_type
            )
            ch_obj.ai_adc_timing_mode = ADCTimingMode.HIGH_RESOLUTION

        # --- Buffer grande per avere margine in caso di rallentamenti ---
        task.in_stream.input_buf_size = 1000

        task.timing.cfg_samp_clk_timing(
            10.0, sample_mode=AcquisitionType.CONTINUOUS
        )
        task.start()

        tasks = [
            asyncio.create_task(user_input(stop_trigger, pause_trigger)),
            asyncio.create_task(v_integrity_check(stop_trigger, pause_trigger)),
            asyncio.create_task(setpoint_handler(stop_trigger, pause_trigger, setpoint_current)),
            asyncio.create_task(logger(stop_trigger, task, setpoint_current)),
        ]

        try:
            await stop_trigger.wait()

        except KeyboardInterrupt:
            print("\n*** Stop manuale (Ctrl+C) ***")
            stop_trigger.set()
        except Exception as e:
            print(f"\n*** ERRORE: {e} ***")
            stop_trigger.set()
        finally:
            inst.write("OUTP 0")
            for t in tasks:
                t.cancel()
            await asyncio.gather(*tasks, return_exceptions=True)
            wb.save(LOG_FILE)
            print(f"\nDati salvati in '{LOG_FILE}'")

    print("Sistema fermato in sicurezza.")


if __name__ == "__main__":
    asyncio.run(main())
