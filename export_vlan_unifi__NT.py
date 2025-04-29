import requests
import urllib3
from openpyxl import Workbook

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# CONFIGURACIÓN
CONTROLLER_URL = "https://192.168.1.203:8443"
USERNAME = "zabbix"
PASSWORD = "zabbix"
SITE = "default"

session = requests.Session()

try:
    # 1. Login
    login_data = {
        "username": USERNAME,
        "password": PASSWORD
    }
    login_response = session.post(f"{CONTROLLER_URL}/api/login", json=login_data, verify=False)
    login_response.raise_for_status()
    print("✅ Login exitoso en la Controladora UniFi.")

    # 2. Obtener lista de redes
    api_url = f"{CONTROLLER_URL}/api/s/{SITE}/rest/networkconf"
    response = session.get(api_url, verify=False)
    response.raise_for_status()
    networks = response.json()['data']

    # 3. Crear archivo Excel con solo las columnas necesarias
    wb = Workbook()
    ws = wb.active
    ws.title = "VLANs"
    ws.append(["Nombre", "VLAN ID", "Router", "Tipo"])

    for net in networks:
        name = net.get('name', '').strip()
        purpose = net.get('purpose', 'corporate')

        # ❌ Filtrar redes no deseadas (WAN o sin nombre)
        if purpose == 'wan' or 'wan' in name.lower():
            continue
        if not name:
            continue

        vlan_id = net.get('vlan', '1' if not net.get('vlan_enabled') else 'Sin ID')

        # Lógica de tipo de router
        if purpose == 'vlan-only' or (purpose == 'corporate' and not net.get('gateway')):
            router = 'Third-party Gateway'
        elif net.get('gateway'):
            router = 'Controladora'
        else:
            router = '-'

        ws.append([name, vlan_id, router, purpose])

    # 4. Guardar archivo
    output_filename = "unifi_vlans_simplificado.xlsx"
    wb.save(output_filename)
    print(f"✅ Archivo Excel generado exitosamente: {output_filename}")

except requests.HTTPError as e:
    print(f"❌ Error HTTP: {e.response.status_code} - {e.response.reason}")
except Exception as ex:
    print(f"❌ Error general: {ex}")
