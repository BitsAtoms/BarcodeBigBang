import os
import json
import time
import re
import ctypes
import xml.etree.ElementTree as ET

import serial
from serial.tools import list_ports


# -----------------------------
# Configuración
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MAPPING_JSON_PATH = os.path.join(BASE_DIR, "barcode_id.json")

RELATIVE_XML_PATH = os.path.join("admira", "conditions", "biomax.xml")

# Si sabes el puerto, ponlo aquí (por ejemplo "COM7").
# Si lo dejas en None, el script intenta autodetectarlo.
SERIAL_PORT = None  # Ej: "COM7"

# Parámetros típicos (ajusta si tu escáner usa otros)
BAUDRATE = 9600
BYTESIZE = 8
PARITY = "N"
STOPBITS = 1

# Tiempo de espera para lecturas del puerto
READ_TIMEOUT_SEC = 1.0


# -----------------------------
# Buscar biomax.xml en cualquier unidad
# -----------------------------
def list_windows_drives():
    buf = ctypes.create_unicode_buffer(256)
    n = ctypes.windll.kernel32.GetLogicalDriveStringsW(len(buf), buf)
    if n == 0:
        return []
    return [d for d in buf.value.split("\x00") if d]

def find_biomax_xml():
    for drive in list_windows_drives():
        candidate = os.path.join(drive, RELATIVE_XML_PATH)
        if os.path.isfile(candidate):
            return candidate
    return None


# -----------------------------
# JSON mapping
# -----------------------------
def load_mapping():
    with open(MAPPING_JSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError("El JSON debe ser un objeto: {\"CODIGO\": \"VALOR\"}")
    return {str(k).strip(): str(v) for k, v in data.items()}


# -----------------------------
# XML saneado (evita basura al inicio y comillas sueltas)
# -----------------------------
def _read_xml_sanitized_bytes(xml_path: str) -> bytes:
    raw = open(xml_path, "rb").read()

    # Quitar BOM UTF-8 si existe
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]

    # Cortar todo lo anterior a <?xml o al primer '<'
    idx_decl = raw.find(b"<?xml")
    if idx_decl != -1:
        raw = raw[idx_decl:]
    else:
        idx_lt = raw.find(b"<")
        if idx_lt > 0:
            raw = raw[idx_lt:]

    text = raw.decode("utf-8", errors="replace")

    # Eliminar comillas sueltas al final de línea:  ...>"\n
    text = re.sub(r'"\s*(\r?\n)', r"\1", text)
    text = re.sub(r'"\s*$', "", text)

    return text.encode("utf-8")


def update_xml(xml_path: str, new_value_text: str, new_tstamp: str):
    xml_bytes = _read_xml_sanitized_bytes(xml_path)
    root = ET.fromstring(xml_bytes)
    tree = ET.ElementTree(root)

    root.set("tstamp", new_tstamp)

    value_elem = root.find("value")
    if value_elem is None:
        raise RuntimeError("No se encontró el nodo <value> en biomax.xml")
    value_elem.text = new_value_text

    tree.write(xml_path, encoding="UTF-8", xml_declaration=True)


def unix_ts_seconds() -> str:
    return str(int(time.time()))


# -----------------------------
# Serial: autodetección del puerto
# -----------------------------
def autodetect_serial_port() -> str:
    """
    Intenta elegir un puerto USB-Serial.
    Prioriza puertos cuyo descriptor/hwid sugiera USB/Serial.
    """
    ports = list(list_ports.comports())
    if not ports:
        raise RuntimeError("No se encontraron puertos COM. ¿Está conectado el escáner en modo USB-COM?")

    # Primero, candidatos con pinta de USB serial
    preferred = []
    others = []
    for p in ports:
        desc = (p.description or "").lower()
        hwid = (p.hwid or "").lower()

        if ("usb" in desc) or ("usb" in hwid) or ("serial" in desc) or ("ch340" in desc) or ("cp210" in desc) or ("ftdi" in desc):
            preferred.append(p)
        else:
            others.append(p)

    chosen = (preferred[0] if preferred else ports[0])
    return chosen.device


def open_serial(port_name: str) -> serial.Serial:
    return serial.Serial(
        port=port_name,
        baudrate=BAUDRATE,
        bytesize=BYTESIZE,
        parity=PARITY,
        stopbits=STOPBITS,
        timeout=READ_TIMEOUT_SEC
    )


def clean_scanned_line(line: str) -> str:
    # Quita CR/LF y espacios
    return line.strip().replace("\x00", "")


def main():
    if not os.path.isfile(MAPPING_JSON_PATH):
        raise FileNotFoundError(f"No existe JSON: {MAPPING_JSON_PATH}")

    xml_path = find_biomax_xml()
    if not xml_path:
        raise FileNotFoundError("No se encontró biomax.xml en \\admira\\conditions\\ en ninguna unidad.")

    port = SERIAL_PORT or autodetect_serial_port()

    # Abrir serie
    ser = open_serial(port)

    # Opcional: limpiar buffers
    try:
        ser.reset_input_buffer()
        ser.reset_output_buffer()
    except Exception:
        pass

    # Loop principal: leer códigos
    while True:
        try:
            raw = ser.readline()  # lee hasta \n o timeout
            if not raw:
                continue

            try:
                line = raw.decode("utf-8", errors="ignore")
            except Exception:
                continue

            code = clean_scanned_line(line)
            if not code:
                continue

            mapping = load_mapping()  # recargar por si editas el JSON sin reiniciar
            if code not in mapping:
                continue

            value = mapping[code]
            update_xml(xml_path, value, unix_ts_seconds())

        except serial.SerialException:
            # Si se desconecta/reconecta el escáner, reintentar
            try:
                ser.close()
            except Exception:
                pass
            time.sleep(1.0)

            # Re-detectar y reabrir
            port = SERIAL_PORT or autodetect_serial_port()
            ser = open_serial(port)

        except Exception:
            # Sin logs (como pediste): ignora errores puntuales y sigue
            time.sleep(0.2)


if __name__ == "__main__":
    main()
