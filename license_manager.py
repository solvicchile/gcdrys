import base64
import hashlib
import json
import os
import socket
from datetime import datetime
from urllib import error as urllib_error
from urllib import request as urllib_request


LICENSE_API_URL = os.getenv("RYS_LICENSE_API_URL", "https://licencias.neicon.cl/validar.php").strip()
LICENSE_API_SECRET = os.getenv("RYS_LICENSE_API_SECRET", "").strip()
LICENSE_GRACE_DAYS = int(os.getenv("RYS_LICENSE_GRACE_DAYS") or "3")
TRIAL_DAYS = int(os.getenv("RYS_TRIAL_DAYS") or "15")


def obtener_nombre_equipo():
    try:
        return socket.gethostname()
    except Exception:
        return "Equipo no identificado"


def obtener_usuario_sistema():
    return os.getenv("USERNAME") or os.getenv("USER") or "Usuario Windows no identificado"


def obtener_directorio_app():
    base = os.getenv("APPDATA") or os.path.expanduser("~")
    ruta = os.path.join(base, "GCDRYS")
    os.makedirs(ruta, exist_ok=True)
    return ruta


def obtener_ruta_licencia_local():
    return os.path.join(obtener_directorio_app(), "licencia.json")


def obtener_ruta_evaluacion_local():
    return os.path.join(obtener_directorio_app(), "evaluacion.json")


def codificar_texto_local(texto):
    if not texto:
        return ""
    return base64.urlsafe_b64encode(texto.encode("utf-8")).decode("ascii")


def decodificar_texto_local(texto):
    if not texto:
        return ""
    try:
        return base64.urlsafe_b64decode(texto.encode("ascii")).decode("utf-8")
    except Exception:
        return ""


def obtener_equipo_id():
    base = f"{obtener_nombre_equipo()}|{obtener_usuario_sistema()}"
    return hashlib.sha256(base.encode("utf-8")).hexdigest()[:32]


def enmascarar_licencia(license_key):
    license_key = str(license_key or "").strip()
    if len(license_key) <= 8:
        return "Sin licencia"
    return f"{license_key[:6]}...{license_key[-4:]}"


def hash_licencia(license_key):
    return hashlib.sha256(str(license_key or "").strip().encode("utf-8")).hexdigest()


def cargar_licencia_local():
    ruta = obtener_ruta_licencia_local()
    if not os.path.exists(ruta):
        return {}

    try:
        with open(ruta, "r", encoding="utf-8") as archivo:
            datos = json.load(archivo)
    except Exception:
        return {}

    datos["license_key"] = decodificar_texto_local(datos.get("license_key_local", ""))
    return datos


def guardar_licencia_local(datos):
    licencia = str(datos.get("license_key", "")).strip()
    datos_guardar = dict(datos)
    datos_guardar["license_key_local"] = codificar_texto_local(licencia)
    datos_guardar.pop("license_key", None)
    datos_guardar["ultima_validacion_local"] = datetime.now().isoformat(timespec="seconds")

    with open(obtener_ruta_licencia_local(), "w", encoding="utf-8") as archivo:
        json.dump(datos_guardar, archivo, ensure_ascii=False, indent=2)


def validar_licencia_online(license_key):
    if not LICENSE_API_URL:
        raise RuntimeError("Falta configurar RYS_LICENSE_API_URL")

    if not LICENSE_API_SECRET:
        raise RuntimeError("Falta configurar RYS_LICENSE_API_SECRET")

    payload = json.dumps(
        {
            "license_key": license_key,
            "equipo_id": obtener_equipo_id(),
            "api_secret": LICENSE_API_SECRET,
        }
    ).encode("utf-8")
    req = urllib_request.Request(
        LICENSE_API_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib_request.urlopen(req, timeout=12) as respuesta:
            raw = respuesta.read().decode("utf-8")
    except urllib_error.HTTPError as e:
        raw = e.read().decode("utf-8")
    except urllib_error.URLError as e:
        raise RuntimeError(f"No se pudo conectar al servidor de licencias: {e.reason}")

    try:
        datos = json.loads(raw)
    except json.JSONDecodeError:
        raise RuntimeError("El servidor de licencias respondio un formato no valido")

    if datos.get("ok"):
        datos["license_key"] = license_key
        datos["license_key_hash"] = hash_licencia(license_key)
        guardar_licencia_local(datos)

    return datos


def obtener_estado_licencia():
    datos = cargar_licencia_local()
    if not datos:
        return {
            "ok": False,
            "estado": "sin_licencia",
            "cliente": "Licencia no activada",
            "dias_restantes": None,
            "license_key": "",
        }
    return datos


def calcular_dias_restantes(fecha_expiracion):
    if not fecha_expiracion:
        return None

    try:
        expira = datetime.strptime(str(fecha_expiracion)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None

    return (expira - datetime.now().date()).days


def licencia_local_vigente():
    datos = obtener_estado_licencia()
    if datos.get("estado") != "activa":
        return False, datos

    dias = calcular_dias_restantes(datos.get("fecha_expiracion"))
    if dias is None:
        dias = datos.get("dias_restantes")

    try:
        dias = int(dias)
    except (TypeError, ValueError):
        dias = None

    if dias is not None and dias < 0:
        datos["estado"] = "expirada"
        datos["dias_restantes"] = dias
        return False, datos

    datos["dias_restantes"] = dias
    return True, datos


def cargar_evaluacion_local():
    ruta = obtener_ruta_evaluacion_local()
    if os.path.exists(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as archivo:
                return json.load(archivo)
        except Exception:
            return {}
    return {}


def guardar_evaluacion_local(datos):
    with open(obtener_ruta_evaluacion_local(), "w", encoding="utf-8") as archivo:
        json.dump(datos, archivo, ensure_ascii=False, indent=2)


def obtener_estado_evaluacion():
    datos = cargar_evaluacion_local()
    inicio = datos.get("fecha_inicio")

    if not inicio:
        inicio = datetime.now().date().isoformat()
        datos = {
            "fecha_inicio": inicio,
            "dias_evaluacion": TRIAL_DAYS,
            "equipo_id": obtener_equipo_id(),
        }
        guardar_evaluacion_local(datos)

    try:
        fecha_inicio = datetime.strptime(str(inicio)[:10], "%Y-%m-%d").date()
    except ValueError:
        fecha_inicio = datetime.now().date()

    dias_usados = (datetime.now().date() - fecha_inicio).days
    dias_restantes = TRIAL_DAYS - dias_usados
    expira = fecha_inicio.toordinal() + TRIAL_DAYS

    return {
        "estado": "evaluacion" if dias_restantes >= 0 else "evaluacion_expirada",
        "fecha_inicio": fecha_inicio.isoformat(),
        "fecha_expiracion": datetime.fromordinal(expira).date().isoformat(),
        "dias_restantes": max(dias_restantes, 0),
        "dias_usados": max(dias_usados, 0),
        "dias_evaluacion": TRIAL_DAYS,
    }


def obtener_estado_acceso_software():
    licencia_ok, licencia = licencia_local_vigente()
    if licencia_ok:
        return {
            "permitido": True,
            "modo": "licenciado",
            "mensaje": f"Licencia activa: {licencia.get('cliente', 'Cliente no informado')}",
            "detalle": f"Dias restantes: {licencia.get('dias_restantes', 'No informado')}",
            "licencia": licencia,
        }

    evaluacion = obtener_estado_evaluacion()
    if evaluacion["estado"] == "evaluacion":
        return {
            "permitido": True,
            "modo": "evaluacion",
            "mensaje": f"Version de evaluacion: quedan {evaluacion['dias_restantes']} dias",
            "detalle": f"Expira el {evaluacion['fecha_expiracion']}",
            "evaluacion": evaluacion,
        }

    return {
        "permitido": False,
        "modo": "evaluacion_expirada",
        "mensaje": "Version de evaluacion expirada",
        "detalle": "Ingrese una licencia valida para continuar usando el software.",
        "evaluacion": evaluacion,
    }


def obtener_cliente_licencia():
    return obtener_estado_licencia().get("cliente") or "Licencia no activada"


def resumen_licencia_para_metadatos(app_version):
    datos = obtener_estado_licencia()
    cliente = datos.get("cliente") or "Licencia no activada"
    licencia = enmascarar_licencia(datos.get("license_key", ""))
    return f"Licencia a nombre de: {cliente} | Licencia: {licencia} | Software: GCDRYS-2 {app_version}"
