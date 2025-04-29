import requests
import urllib3
from openpyxl import Workbook

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

CONTROLLER_URL = "https://192.168.1.203:8443"
USERNAME = "zabbix"
PASSWORD = "zabbix"
SITE = "default"

session = requests.Session()

def infer_vlan_policy(excluded_ids, native_id, all_ids):
    if not excluded_ids:
        return "Permitir todo"
    remaining = all_ids - ({native_id} if native_id else set())
    if excluded_ids.issuperset(remaining):
        return "Bloquear todo"
    return "Personalizado"

try:
    session.post(f"{CONTROLLER_URL}/api/login", json={"username": USERNAME, "password": PASSWORD}, verify=False)

    API_BASE = f"{CONTROLLER_URL}/api/s/{SITE}"

    networks = session.get(f"{API_BASE}/rest/networkconf", verify=False).json()['data']
    networkconf_name_map = {n["_id"]: n.get("name", "Desconocido") for n in networks if "_id" in n}
    all_networkconf_ids = set(networkconf_name_map.keys())

    portconfs = session.get(f"{API_BASE}/rest/portconf", verify=False).json()['data']
    portconf_map = {p["_id"]: p.get("name", "-") for p in portconfs if "_id" in p}

    devices = session.get(f"{API_BASE}/stat/device", verify=False).json()["data"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Puertos"
    ws.append(["Switch", "Port", "Name", "PoE Mode", "Estado", "Profile", "Native VLAN", "Política VLAN Etiquetada"])

    for device in devices:
        if device.get("type") != "usw":
            continue

        switch_name = device.get("name", device.get("mac"))
        port_table = {p["port_idx"]: p for p in device.get("port_table", [])}
        port_overrides = {p["port_idx"]: p for p in device.get("port_overrides", [])}

        for port_idx in sorted(port_table.keys()):
            base = port_table[port_idx]
            override = port_overrides.get(port_idx, {})

            name = override.get("name") or base.get("name", f"Port {port_idx}")
            poe_mode = override.get("poe_mode") or base.get("poe_mode", "off")
            enable = override.get("enabled") if "enabled" in override else base.get("enable", True)

            estado = "Activo" if enable else "Deshabilitado"
            if base.get("portconf_id") is None and not enable:
                estado = "Restringido"

            portconf_id = override.get("portconf_id") or base.get("portconf_id")
            profile_name = portconf_map.get(portconf_id, "-") if portconf_id else "-"

            native_id = override.get("native_networkconf_id") or base.get("native_networkconf_id")
            vlan_name = "-" if profile_name != "-" else networkconf_name_map.get(native_id, "Default" if not native_id else "Desconocido")

            if profile_name != "-":
                vlan_policy = "-"
            else:
                excluded = set(override.get("excluded_networkconf_ids", base.get("excluded_networkconf_ids", [])))
                vlan_policy = infer_vlan_policy(excluded, native_id, all_networkconf_ids)

            ws.append([
                switch_name,
                port_idx,
                name,
                poe_mode,
                estado,
                profile_name,
                vlan_name,
                vlan_policy
            ])

    output_filename = "unifi_switch_ports_vlan_policy_final.xlsx"
    wb.save(output_filename)
    print(f"✅ Exportación completa: {output_filename}")

except Exception as e:
    print(f"❌ Error: {e}")
