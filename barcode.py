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
SERIAL_PORT = None # Ej: "COM7"

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
    root = ET.fromstring(xml_bytes)      # root = <conditions>
    tree = ET.ElementTree(root)

    # Buscar el condition correcto (id="4"). Si no existe, usa el primero.
    condition = root.find(".//condition[@id='4']")
    if condition is None:
        condition = root.find(".//condition")
    if condition is None:
        raise RuntimeError("No se encontró ningún nodo <condition> en biomax.xml")

    # Actualizar tstamp en <condition>
    condition.set("tstamp", new_tstamp)

    # Actualizar <value> dentro de <condition>
    value_elem = condition.find("value")
    if value_elem is None:
        raise RuntimeError("No se encontró el nodo <value> dentro de <condition> en biomax.xml")
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
    print("=== INICIO BARCODE DEBUG ===")

    if not os.path.isfile(MAPPING_JSON_PATH):
        raise FileNotFoundError(f"No existe JSON: {MAPPING_JSON_PATH}")

    print(f"JSON encontrado en: {MAPPING_JSON_PATH}")

    xml_path = find_biomax_xml()
    if not xml_path:
        raise FileNotFoundError("No se encontró biomax.xml en \\admira\\conditions\\ en ninguna unidad.")

    print(f"XML encontrado en: {xml_path}")

    port = SERIAL_PORT or autodetect_serial_port()
    print(f"Puerto serie seleccionado: {port}")
    print(f"Baudrate: {BAUDRATE}, Bytesize: {BYTESIZE}, Parity: {PARITY}, Stopbits: {STOPBITS}")

    # Abrir serie
    ser = open_serial(port)
    print("Puerto abierto correctamente.")

    # Opcional: limpiar buffers
    try:
        ser.reset_input_buffer()
        ser.reset_output_buffer()
        print("Buffers limpiados.")
    except Exception:
        pass

    print("Esperando datos del escáner...\n")

    # Loop principal: leer códigos
    while True:
        try:
            raw = ser.readline()  # lee hasta \n o timeout

            if not raw:
                continue

            # print(f"[RAW BYTES] {raw}")

            try:
                line = raw.decode("utf-8", errors="ignore")
            except Exception as e:
                print(f"Error decodificando: {e}")
                continue

            # print(f"[DECODED] '{line}'")

            code = clean_scanned_line(line)
            print(f"[CLEAN CODE] '{code}'")

            if not code:
                print("Código vacío después de limpiar.")
                continue

            mapping = load_mapping()

            if code not in mapping:
                print(f"Código '{code}' NO encontrado en JSON.")
                continue

            value = mapping[code]
            print(f"Código válido. Valor asociado: {value}")

            update_xml(xml_path, value, unix_ts_seconds())
            print("XML actualizado correctamente.\n")

        except serial.SerialException as e:
            print(f"Error serial: {e}")
            try:
                ser.close()
            except Exception:
                pass

            time.sleep(1.0)

            print("Reintentando conexión al puerto...")
            port = SERIAL_PORT or autodetect_serial_port()
            ser = open_serial(port)
            print("Puerto reabierto.")

        except KeyboardInterrupt:
            print("\nPrograma detenido manualmente.")
            break

        except Exception as e:
            print(f"Error inesperado: {e}")
            time.sleep(0.2)


if __name__ == "__main__":
    main()
