import pyvisa
import time
import openpyxl

# ---------------- CONFIG ----------------
ADDRESS = "USB0::0x2EC7::0x8500::804690011807170001::INSTR"  # cambia col tuo
DISCHARGE_A = 3.3
STOP_VOLT = 3.0
STOP_CAP = 8
LOG_FILE = "discharge_log_1.xlsx"
LOG_INTERVAL = 0.01     # secondi tra misure
# ----------------------------------------

rm = pyvisa.ResourceManager()
inst = rm.open_resource(ADDRESS)
print("IDN:", inst.query("*IDN?").strip())

inst.write("SENS ON")

# Setup Battery Mode
inst.write("SYST:RUNMode BATT")
inst.write(f"BATT:DISCharge:CURR {DISCHARGE_A}")
inst.write(f"BATT:STOP:VOLT {STOP_VOLT}")
inst.write(f"BATT:STOP:CAP {STOP_CAP}")

# Lettura capacità iniziale
ah0 = inst.query("BATT:CAP?").strip()
print("Capacità iniziale:", ah0, "Ah")

# Crea file Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Discharge"
ws.append(["Index", "Time (s)", "Voltage (V)", "Current (A)", "Capacity (Ah)"])

# Avvia scarica
inst.write("BATT:STAT ON")
print("Scarica avviata... (CTRL+C per fermarla manualmente)")

start_time = time.time()
row_index = 1

try:
    while True:
        state = inst.query("BATT:STAT?").strip()
        v = float(inst.query("MEAS:VOLT?"))
        i = float(inst.query("MEAS:CURR?"))
        ah = float(inst.query("BATT:CAP?"))
        t = time.time() - start_time  # tempo relativo in secondi

        # Scrivi su Excel
        ws.append([row_index, round(t, 2), v, i, ah])

        # Stampa live
        print(f"\r#{row_index} | t={t:.1f}s | V={v:.2f} V | I={i:.3f} A | Ah={ah:.3f}", end="")

        row_index += 1

        if state in ["0", "OFF"]:
            print("\nScarica completata automaticamente.")
            break

        time.sleep(LOG_INTERVAL)

except KeyboardInterrupt:
    print("\nStop manuale richiesto.")
    inst.write("BATT:STAT OFF")

# Lettura finale
ah_final = inst.query("BATT:CAP?").strip()
print("Capacità finale:", ah_final, "Ah")

# Salva Excel
wb.save(LOG_FILE)
print(f"Dati salvati in {LOG_FILE}")

# Chiudi
inst.close()
rm.close()


