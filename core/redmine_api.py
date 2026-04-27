"""
Cliente HTTP para la API REST de Redmine.
Funciones puras que no dependen de la UI ni del estado de la app.
"""
import json
import requests
import pandas as pd


def hdrs(api_key):
    return {"X-Redmine-API-Key": api_key, "Content-Type": "application/json"}


def obtener_proyectos(url, api_key):
    """Trae proyectos paginando para obtener todos, solo proyectos activos."""
    proyectos = []
    offset = 0
    limit = 100
    try:
        while True:
            r = requests.get(
                f"{url}/projects.json?limit={limit}&offset={offset}&status=1",
                headers=hdrs(api_key))
            if r.status_code != 200:
                return None, f"HTTP {r.status_code}"
            data = r.json()
            batch = data.get("projects", [])
            proyectos.extend(batch)
            total = data.get("total_count", 0)
            offset += limit
            if offset >= total or not batch:
                break
        return proyectos, None
    except Exception as e:
        return None, str(e)


def obtener_id_actividad(url, api_key, nombre):
    r = requests.get(f"{url}/enumerations/time_entry_activities.json",
                     headers=hdrs(api_key))
    if r.status_code != 200:
        return None, f"HTTP {r.status_code}"
    for a in r.json().get("time_entry_activities", []):
        if a["name"].lower() == nombre.lower():
            return a["id"], None
    return None, f"Actividad '{nombre}' no encontrada"


def obtener_mi_id(url, api_key):
    r = requests.get(f"{url}/users/current.json", headers=hdrs(api_key))
    if r.status_code == 200:
        return r.json()["user"]["id"], None
    return None, f"HTTP {r.status_code}"


def crear_issue(url, api_key, proyecto_id, titulo, usuario_id):
    payload = {"issue": {
        "project_id": proyecto_id, "subject": titulo,
        "tracker_id": 1, "assigned_to_id": usuario_id,
    }}
    r = requests.post(f"{url}/issues.json", headers=hdrs(api_key),
                      data=json.dumps(payload))
    if r.status_code == 201:
        return r.json()["issue"]["id"], None
    return None, r.text


def cargar_entrada(url, api_key, issue_id, fecha, horas, comentario,
                   act_id, hs_cliente, tipo_hora, ubicacion):
    entry = {"time_entry": {
        "issue_id": int(issue_id), "spent_on": fecha,
        "hours": float(horas), "comments": comentario, "activity_id": act_id,
        "custom_fields": [
            {"id": 3, "value": str(hs_cliente)},
            {"id": 2, "value": tipo_hora},
            {"id": 4, "value": ubicacion},
        ]
    }}
    r = requests.post(f"{url}/time_entries.json", headers=hdrs(api_key),
                      data=json.dumps(entry))
    return r.status_code, r.text


def obtener_titulo_issue(url, api_key, issue_id):
    """Obtiene el título del issue y lo divide en ID_Ticket y Titulo."""
    try:
        r = requests.get(f"{url}/issues/{int(issue_id)}.json",
                         headers=hdrs(api_key))
        if r.status_code != 200:
            return None, None
        subject = r.json().get("issue", {}).get("subject", "")
        for sep in [" - ", " # ", "-", "#"]:
            if sep in subject:
                parts = subject.split(sep, 1)
                return parts[0].strip(), parts[1].strip()
        return "", subject.strip()
    except Exception:
        return None, None


def verificar_duplicado_redmine(url, api_key, issue_id, fecha_str, horas, comentario):
    """Consulta Redmine para ver si ya existe una entrada idéntica."""
    try:
        r = requests.get(
            f"{url}/time_entries.json",
            headers=hdrs(api_key),
            params={"issue_id": issue_id, "limit": 100}
        )
        if r.status_code != 200:
            return False
        entries = r.json().get("time_entries", [])
        for e in entries:
            if (e.get("spent_on") == fecha_str and
                abs(float(e.get("hours", 0)) - float(horas)) < 0.01 and
                e.get("comments", "").strip() == comentario.strip()):
                return True
        return False
    except Exception:
        return False


# ============================================================
#  HELPERS DE TEXTO
# ============================================================

def limpiar(val):
    """Convierte un valor de Excel a string limpio."""
    return "" if pd.isna(val) else str(val).strip()


def armar_titulo_issue(id_ticket, titulo):
    """Título del issue nuevo = ID_Ticket + Titulo."""
    c = limpiar(id_ticket)
    t = limpiar(titulo)
    if c and t: return f"{c} - {t}"
    return c or t


def armar_comentario(comentario, id_ticket=None, titulo=None):
    """Comentario de la entrada de tiempo dedicado.
    Si hay comentario, lo usa. Si no, usa ID_Ticket - Titulo como fallback."""
    c = limpiar(comentario)
    if c:
        return c
    t = limpiar(id_ticket)
    ti = limpiar(titulo)
    if t and ti: return f"{t} - {ti}"
    return t or ti
