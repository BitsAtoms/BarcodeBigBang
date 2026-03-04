from serial.tools import list_ports

for p in list_ports.comports():
    print("----")
    print("Device:", p.device)
    print("Description:", p.description)
    print("HWID:", p.hwid)
    print("VID:", p.vid)
    print("PID:", p.pid)