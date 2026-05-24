import os
import hashlib
import hmac
import json
import subprocess
import tempfile
from urllib.parse import urlencode
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from datetime import datetime

import psycopg2
from openpyxl import Workbook
from PIL import Image, ImageTk

from license_manager import (
    enmascarar_licencia,
    hash_licencia,
    obtener_estado_acceso_software,
    obtener_estado_licencia,
    obtener_nombre_equipo,
    obtener_usuario_sistema,
    resumen_licencia_para_metadatos as resumen_licencia_base,
    validar_licencia_online,
)


# =========================
# CONFIG DB
# =========================
# Configure estas variables en Windows antes de ejecutar:
# setx RYS_DB_HOST "aws-1-us-east-1.pooler.supabase.com"
# setx RYS_DB_NAME "postgres"
# setx RYS_DB_USER "postgres.xxxxxxxxxxxxxxxxxxxx"
# setx RYS_DB_PASSWORD "su_password"
# setx RYS_DB_PORT "6543"
DB_HOST = os.getenv("RYS_DB_HOST", "")
DB_NAME = os.getenv("RYS_DB_NAME", "postgres")
DB_USER = os.getenv("RYS_DB_USER", "")
DB_PASSWORD = os.getenv("RYS_DB_PASSWORD", "")
DB_PORT = int(os.getenv("RYS_DB_PORT") or "6543")

APP_TITLE = "GESTION Y CONTROL DOCUMENTAL RYS-2"
APP_VERSION = "v1"
APP_COPYRIGHT = "Copyright (c) 2026 GCDRYS. Todos los derechos reservados."
EMPRESA_NOMBRE = os.getenv("RYS_EMPRESA_NOMBRE", "RYS")
EMPRESA_SUBTITULO = os.getenv("RYS_EMPRESA_SUBTITULO", "Gestion y Control Documental")
DOC_PREFIX = "RYS"
VALIDATION_BASE_URL = os.getenv(
    "RYS_VALIDATION_BASE_URL",
    "https://solvicchile.github.io/rys/validar_documento.html",
).strip()
PASSWORD_HASH_ITERATIONS = 260000
ROLES = ("Superusuario", "admin", "colaborador", "user", "lector")
ROLES_ADMIN_CREACION = ("admin", "colaborador", "user", "lector")
ESTADOS_DOCUMENTO = (
    "Borrador",
    "En revision",
    "Pendiente aprobacion",
    "Vigente",
    "Obsoleto",
    "Anulado",
)
ESTADOS_PROTEGIDOS = ("Vigente", "Obsoleto", "Anulado")
PERMISOS_CATALOGO = (
    ("consulta.ver", "Consulta", "Ver consulta documental"),
    ("consulta.buscar", "Consulta", "Buscar y filtrar documentos"),
    ("documentos.ficha.ver", "Consulta", "Ver ficha documental"),
    ("documentos.historial.ver", "Consulta", "Ver historial documental"),
    ("dashboard.ver", "Dashboard", "Ver dashboard documental"),
    ("dashboard.exportar", "Dashboard", "Exportar dashboard a PDF"),
    ("dashboard.imprimir", "Dashboard", "Imprimir dashboard"),
    ("documentos.crear", "Documentos", "Crear documento en borrador"),
    ("documentos.codigo.generar", "Documentos", "Generar codigo documental"),
    ("documentos.numero.automatico", "Documentos", "Usar numeracion automatica"),
    ("documentos.confidencial.crear", "Documentos", "Crear documento confidencial"),
    ("documentos.confidencial.ver", "Documentos", "Ver documentos confidenciales"),
    ("documentos.nombre.editar", "Documentos", "Editar nombre de documento"),
    ("documentos.confidencial.editar", "Documentos", "Editar confidencialidad"),
    ("documentos.codigo.editar", "Documentos", "Editar codigo, area, tipo, numero y version"),
    ("documentos.editar.borrador", "Documentos", "Editar documento en Borrador"),
    ("documentos.editar.revision", "Documentos", "Editar documento En revision"),
    ("documentos.editar.pendiente_aprobacion", "Documentos", "Editar documento Pendiente aprobacion"),
    ("documentos.editar.vigente", "Documentos", "Editar documento Vigente"),
    ("documentos.editar.obsoleto", "Documentos", "Editar documento Obsoleto"),
    ("documentos.editar.anulado", "Documentos", "Editar documento Anulado"),
    ("documentos.enviar_revision", "Flujo documental", "Enviar documento a revision"),
    ("documentos.devolver_borrador", "Flujo documental", "Devolver documento a borrador"),
    ("documentos.marcar_revisado", "Flujo documental", "Marcar documento como revisado"),
    ("documentos.enviar_aprobacion", "Flujo documental", "Enviar documento a aprobacion"),
    ("documentos.devolver_revision", "Flujo documental", "Devolver documento a revision"),
    ("documentos.aprobar", "Flujo documental", "Aprobar documento y dejar vigente"),
    ("documentos.version.aprobar", "Flujo documental", "Aprobar nueva version"),
    ("documentos.version.obsoletar", "Flujo documental", "Obsoletar version anterior automaticamente"),
    ("documentos.anular", "Flujo documental", "Anular documento"),
    ("documentos.qr.generar", "Validacion QR", "Generar QR documental"),
    ("documentos.qr.guardar", "Validacion QR", "Guardar o copiar QR"),
    ("reportes.general.exportar", "Reportes", "Exportar reporte documental general"),
    ("reportes.general.imprimir", "Reportes", "Imprimir reporte documental general"),
    ("reportes.consulta.exportar", "Reportes", "Exportar consulta filtrada"),
    ("reportes.consulta.imprimir", "Reportes", "Imprimir consulta filtrada"),
    ("reportes.ficha.exportar", "Reportes", "Exportar ficha documental"),
    ("reportes.ficha.imprimir", "Reportes", "Imprimir ficha documental"),
    ("reportes.auditoria.ver", "Reportes", "Ver auditoria de reportes"),
    ("usuarios.ver", "Usuarios", "Ver listado de usuarios"),
    ("usuarios.administrar", "Usuarios", "Administrar usuarios"),
    ("usuarios.crear", "Usuarios", "Crear usuarios"),
    ("usuarios.editar", "Usuarios", "Editar usuarios"),
    ("usuarios.password", "Usuarios", "Cambiar contrasena de usuario"),
    ("usuarios.activar", "Usuarios", "Activar o desactivar usuarios"),
    ("usuarios.rol", "Usuarios", "Cambiar rol de usuario"),
    ("usuarios.cargo", "Usuarios", "Asignar cargo a usuario"),
    ("config.ver", "Configuracion", "Ver configuracion del sistema"),
    ("config.empresa", "Configuracion", "Configurar empresa, logo y subtitulo"),
    ("config.areas", "Configuracion", "Administrar areas"),
    ("config.tipos", "Configuracion", "Administrar tipos de documento"),
    ("config.cargos", "Configuracion", "Crear, editar y desactivar cargos"),
    ("config.cargos.permisos", "Configuracion", "Asignar permisos a cargos"),
)
PERMISOS_TODOS = tuple(codigo for codigo, _, _ in PERMISOS_CATALOGO)
PERMISOS_ROL_BASE = {
    "Superusuario": PERMISOS_TODOS,
    "admin": (
        "consulta.ver", "consulta.buscar", "documentos.ficha.ver", "documentos.historial.ver",
        "dashboard.ver", "dashboard.exportar", "dashboard.imprimir",
        "documentos.crear", "documentos.codigo.generar", "documentos.numero.automatico",
        "documentos.confidencial.crear", "documentos.confidencial.ver", "documentos.nombre.editar",
        "documentos.confidencial.editar", "documentos.editar.borrador", "documentos.editar.revision",
        "documentos.editar.pendiente_aprobacion", "documentos.enviar_revision",
        "documentos.devolver_borrador", "documentos.marcar_revisado", "documentos.enviar_aprobacion",
        "documentos.devolver_revision", "documentos.aprobar", "documentos.version.aprobar",
        "documentos.version.obsoletar", "documentos.anular", "documentos.qr.generar",
        "documentos.qr.guardar", "reportes.general.exportar", "reportes.general.imprimir",
        "reportes.consulta.exportar", "reportes.consulta.imprimir", "reportes.ficha.exportar",
        "reportes.ficha.imprimir", "reportes.auditoria.ver", "usuarios.ver", "usuarios.administrar",
        "usuarios.crear", "usuarios.editar", "usuarios.password", "usuarios.activar",
        "usuarios.rol", "usuarios.cargo", "config.ver", "config.empresa", "config.areas",
        "config.tipos", "config.cargos",
    ),
    "colaborador": (
        "consulta.ver", "consulta.buscar", "documentos.ficha.ver", "documentos.historial.ver",
        "dashboard.ver", "documentos.crear", "documentos.codigo.generar",
        "documentos.numero.automatico", "documentos.confidencial.crear",
        "documentos.confidencial.ver", "documentos.nombre.editar", "documentos.confidencial.editar",
        "documentos.editar.borrador", "documentos.enviar_revision", "documentos.qr.generar",
        "documentos.qr.guardar", "reportes.general.exportar", "reportes.general.imprimir",
        "reportes.consulta.exportar", "reportes.consulta.imprimir", "reportes.ficha.exportar",
        "reportes.ficha.imprimir",
    ),
    "user": (
        "consulta.ver", "consulta.buscar", "documentos.ficha.ver", "documentos.crear",
        "documentos.codigo.generar", "documentos.numero.automatico", "documentos.editar.borrador",
        "documentos.qr.generar", "documentos.qr.guardar",
    ),
    "lector": (
        "consulta.ver", "consulta.buscar", "documentos.ficha.ver",
    ),
}
PERMISO_EDICION_ESTADO = {
    "Borrador": "documentos.editar.borrador",
    "En revision": "documentos.editar.revision",
    "Pendiente aprobacion": "documentos.editar.pendiente_aprobacion",
    "Vigente": "documentos.editar.vigente",
    "Obsoleto": "documentos.editar.obsoleto",
    "Anulado": "documentos.editar.anulado",
}
PERMISO_TRANSICION_ESTADO = {
    ("Borrador", "En revision"): "documentos.enviar_revision",
    ("Borrador", "Anulado"): "documentos.anular",
    ("En revision", "Borrador"): "documentos.devolver_borrador",
    ("En revision", "Pendiente aprobacion"): "documentos.marcar_revisado",
    ("En revision", "Anulado"): "documentos.anular",
    ("Pendiente aprobacion", "En revision"): "documentos.devolver_revision",
    ("Pendiente aprobacion", "Vigente"): "documentos.aprobar",
    ("Pendiente aprobacion", "Anulado"): "documentos.anular",
    ("Vigente", "Obsoleto"): "documentos.version.obsoletar",
    ("Vigente", "Anulado"): "documentos.anular",
}

TABLAS_REQUERIDAS = {
    "areas": ("codigo", "nombre"),
    "tipos_documento": ("codigo", "nombre"),
    "documentos": ("id", "codigo", "nombre", "area", "tipo", "numero", "version", "estado", "fecha", "creado_por", "confidencial", "licencia_cliente", "licencia_key_hash"),
    "usuarios": ("id", "username", "password", "email", "rol", "activo"),
    "documentos_historial": ("id", "documento_id", "codigo", "campo", "valor_anterior", "valor_nuevo", "fecha"),
    "empresa_config": ("id", "nombre", "subtitulo", "logo_path"),
    "cargos": ("id", "nombre", "rol_base", "activo"),
    "cargo_permisos": ("cargo_id", "permiso", "permitido"),
    "documentos_reportes_log": ("id", "documento_id", "accion", "tipo_reporte", "fecha"),
}

PERFILES_PERMISOS_CARGO = {
    "Consulta completa": (
        "consulta.ver", "consulta.buscar", "documentos.ficha.ver", "documentos.historial.ver",
        "dashboard.ver",
    ),
    "Gestion documental": (
        "consulta.ver", "consulta.buscar", "documentos.ficha.ver", "documentos.historial.ver",
        "documentos.crear", "documentos.codigo.generar", "documentos.numero.automatico",
        "documentos.nombre.editar", "documentos.editar.borrador", "documentos.enviar_revision",
        "documentos.qr.generar", "documentos.qr.guardar",
    ),
    "Reportes y auditoria": (
        "consulta.ver", "consulta.buscar", "documentos.ficha.ver", "dashboard.ver",
        "dashboard.exportar", "dashboard.imprimir", "reportes.general.exportar",
        "reportes.general.imprimir", "reportes.consulta.exportar", "reportes.consulta.imprimir",
        "reportes.ficha.exportar", "reportes.ficha.imprimir", "reportes.auditoria.ver",
    ),
}


# =========================
# CONEXION
# =========================
def conectar():
    if not DB_HOST or not DB_USER or not DB_PASSWORD:
        raise RuntimeError(
            "Faltan variables de entorno de base de datos: "
            "RYS_DB_HOST, RYS_DB_USER y RYS_DB_PASSWORD."
        )

    return psycopg2.connect(
        host=DB_HOST,
        database=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD,
        port=DB_PORT,
        sslmode="require",
    )


def diagnosticar_sistema():
    resultado = {
        "ok": True,
        "lineas": [],
    }

    requeridas = {
        "RYS_DB_HOST": DB_HOST,
        "RYS_DB_USER": DB_USER,
        "RYS_DB_PASSWORD": DB_PASSWORD,
    }
    faltantes = [nombre for nombre, valor in requeridas.items() if not valor]

    if faltantes:
        resultado["ok"] = False
        resultado["lineas"].append("Variables faltantes: " + ", ".join(faltantes))
        return resultado

    resultado["lineas"].append("Variables de entorno: OK")

    try:
        with conectar() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT current_database(), current_user, version()")
                db_nombre, db_usuario, db_version = cur.fetchone()
                resultado["lineas"].append(f"Conexion Supabase/PostgreSQL: OK ({db_nombre} / {db_usuario})")
                resultado["lineas"].append(f"Motor: {str(db_version).split(',')[0]}")

                asegurar_configuracion_sistema(cur)
                asegurar_columna_confidencial(cur)
                asegurar_tabla_auditoria_reportes(cur)

                problemas = []
                for tabla, columnas in TABLAS_REQUERIDAS.items():
                    cur.execute(
                        """
                        SELECT column_name
                        FROM information_schema.columns
                        WHERE table_schema = 'public'
                        AND table_name = %s
                        """,
                        (tabla,),
                    )
                    existentes = {fila[0] for fila in cur.fetchall()}

                    if not existentes:
                        problemas.append(f"{tabla}: tabla no encontrada")
                        continue

                    faltan_columnas = [col for col in columnas if col not in existentes]
                    if faltan_columnas:
                        problemas.append(f"{tabla}: faltan columnas {', '.join(faltan_columnas)}")

                if problemas:
                    resultado["ok"] = False
                    resultado["lineas"].append("Esquema: revisar")
                    resultado["lineas"].extend(f"- {problema}" for problema in problemas)
                else:
                    resultado["lineas"].append("Esquema requerido: OK")
    except Exception as e:
        resultado["ok"] = False
        resultado["lineas"].append(f"Conexion/esquema: ERROR - {e}")

    return resultado


def mostrar_diagnostico_sistema(parent=None):
    diagnostico = diagnosticar_sistema()
    titulo = "Diagnostico del sistema"
    mensaje = "\n".join(diagnostico["lineas"]) or "Sin informacion de diagnostico"

    if diagnostico["ok"]:
        messagebox.showinfo(titulo, mensaje, parent=parent)
    else:
        messagebox.showwarning(titulo, mensaje, parent=parent)


# =========================
# SEGURIDAD
# =========================
def hash_password(password):
    salt = os.urandom(16).hex()
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        bytes.fromhex(salt),
        PASSWORD_HASH_ITERATIONS,
    ).hex()
    return f"pbkdf2_sha256${PASSWORD_HASH_ITERATIONS}${salt}${digest}"


def hash_password_legacy(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verificar_password(password, hash_guardado):
    if not hash_guardado:
        return False, False

    if hash_guardado.startswith("pbkdf2_sha256$"):
        try:
            _, iteraciones, salt, digest = hash_guardado.split("$", 3)
            calculado = hashlib.pbkdf2_hmac(
                "sha256",
                password.encode("utf-8"),
                bytes.fromhex(salt),
                int(iteraciones),
            ).hex()
        except (ValueError, TypeError):
            return False, False

        return hmac.compare_digest(calculado, digest), False

    valido = hmac.compare_digest(hash_password_legacy(password), hash_guardado)
    return valido, valido


def validar_password_plana(password):
    if len(password) < 8:
        return "La contrasena debe tener al menos 8 caracteres"

    return ""


# =========================
# UTILIDADES UI
# =========================
def centrar_ventana(ventana, ancho, alto):
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


def cargar_logo(ruta, tamano):
    if not os.path.exists(ruta):
        return None

    img = Image.open(ruta)
    img = img.resize(tamano)
    return ImageTk.PhotoImage(img)


def cargar_logo_empresa(tamano):
    return cargar_logo(obtener_logo_empresa(), tamano)


def resolver_recurso(nombre_archivo):
    rutas = [
        os.path.join(os.getcwd(), nombre_archivo),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), nombre_archivo),
    ]

    for ruta in rutas:
        if os.path.exists(ruta):
            return ruta

    return ""


def aplicar_icono(ventana):
    if os.path.exists("logo.ico"):
        try:
            ventana.iconbitmap("logo.ico")
        except Exception:
            pass


def extraer_codigo(valor_combo):
    return valor_combo.split(" - ")[0].strip()


def formatear_fecha(valor):
    if isinstance(valor, datetime):
        return valor.strftime("%d-%m-%Y %H:%M")

    return "" if valor is None else str(valor)


def formatear_fila(fila):
    return tuple(formatear_fecha(valor) for valor in fila)


def resumen_licencia_para_metadatos():
    return resumen_licencia_base(APP_VERSION)


def color_estado_documental(estado):
    estado = normalizar_estado(estado)

    if estado == "Vigente":
        return "#1f8f45"

    if estado in ("En revision", "Pendiente aprobacion"):
        return "#d79b00"

    return "#c0392b"


def estado_vigencia_documental(estado):
    return "VIGENTE" if normalizar_estado(estado) == "Vigente" else "NO VIGENTE"


def buscar_ultimo_cambio_estado(historial, estados_objetivo):
    for fecha_h, campo, anterior, nuevo, observacion, usuario in historial:
        if campo == "estado" and normalizar_estado(nuevo) in estados_objetivo:
            return usuario, fecha_h

    return "Pendiente", None


def construir_resumen_validacion(ficha, historial):
    (
        _,
        codigo,
        nombre,
        area,
        area_nombre,
        tipo,
        tipo_nombre,
        numero,
        version,
        estado,
        fecha,
        creado_por,
        confidencial,
    ) = ficha

    estado = normalizar_estado(estado)
    revisado_por, fecha_revision = buscar_ultimo_cambio_estado(
        historial,
        ("Pendiente aprobacion", "En revision"),
    )
    aprobado_por, fecha_aprobacion = buscar_ultimo_cambio_estado(historial, ("Vigente",))

    return {
        "nombre": nombre or "Sin nombre documental",
        "codigo": codigo,
        "estado": estado,
        "vigencia": estado_vigencia_documental(estado),
        "confidencial": bool(confidencial),
        "area": f"{area} - {area_nombre}",
        "tipo": f"{tipo} - {tipo_nombre}",
        "numero": numero,
        "version": version,
        "creado_por": creado_por,
        "fecha_creacion": fecha,
        "revisado_por": revisado_por,
        "fecha_revision": fecha_revision,
        "aprobado_por": aprobado_por,
        "fecha_aprobacion": fecha_aprobacion,
    }


def serializar_resumen_validacion(resumen):
    payload = {
        "sistema": "GCDRYS-2",
        "codigo": resumen["codigo"],
        "nombre": resumen["nombre"],
        "estado": resumen["estado"],
        "vigencia": resumen["vigencia"],
        "confidencial": "SI" if resumen.get("confidencial") else "NO",
        "creado_por": resumen["creado_por"],
        "fecha_creacion": formatear_fecha(resumen["fecha_creacion"]),
        "revisado_por": resumen["revisado_por"],
        "fecha_revision": formatear_fecha(resumen["fecha_revision"]),
        "aprobado_por": resumen["aprobado_por"],
        "fecha_aprobacion": formatear_fecha(resumen["fecha_aprobacion"]),
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def construir_texto_validacion(resumen):
    return "\n".join(
        [
            "GCDRYS-2 | VALIDACION DOCUMENTAL",
            f"Documento: {resumen['nombre']}",
            f"Codigo: {resumen['codigo']}",
            f"Estado: {resumen['estado']} ({resumen['vigencia']})",
            f"Confidencial: {'SI' if resumen.get('confidencial') else 'NO'}",
            f"Creado por: {resumen['creado_por']}",
            f"Fecha creacion: {formatear_fecha(resumen['fecha_creacion'])}",
            f"Revisado por: {resumen['revisado_por']}",
            f"Fecha revision: {formatear_fecha(resumen['fecha_revision']) or 'Pendiente'}",
            f"Aprobado por: {resumen['aprobado_por']}",
            f"Fecha aprobacion: {formatear_fecha(resumen['fecha_aprobacion']) or 'Pendiente'}",
        ]
    )


def construir_url_validacion(resumen):
    if not VALIDATION_BASE_URL:
        return ""

    separador = "&" if "?" in VALIDATION_BASE_URL else "?"
    return f"{VALIDATION_BASE_URL}{separador}{urlencode({'codigo': resumen['codigo']})}"


def asegurar_configuracion_sistema(cur):
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS empresa_config (
            id INTEGER PRIMARY KEY DEFAULT 1,
            nombre TEXT NOT NULL DEFAULT 'RYS',
            subtitulo TEXT NOT NULL DEFAULT 'Gestion y Control Documental',
            logo_path TEXT,
            updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            CONSTRAINT empresa_config_id_unico CHECK (id = 1)
        )
        """
    )
    cur.execute(
        """
        INSERT INTO empresa_config (id, nombre, subtitulo)
        VALUES (1, %s, %s)
        ON CONFLICT (id) DO NOTHING
        """,
        (EMPRESA_NOMBRE, EMPRESA_SUBTITULO),
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS cargos (
            id BIGSERIAL PRIMARY KEY,
            nombre TEXT NOT NULL UNIQUE,
            rol_base TEXT NOT NULL DEFAULT 'user',
            activo BOOLEAN NOT NULL DEFAULT true,
            fecha TIMESTAMPTZ NOT NULL DEFAULT now()
        )
        """
    )
    cur.execute("ALTER TABLE cargos ADD COLUMN IF NOT EXISTS rol_base TEXT NOT NULL DEFAULT 'user'")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS cargo_permisos (
            cargo_id BIGINT NOT NULL REFERENCES cargos(id) ON DELETE CASCADE,
            permiso TEXT NOT NULL,
            permitido BOOLEAN NOT NULL DEFAULT true,
            updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
            PRIMARY KEY (cargo_id, permiso)
        )
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS cargo_permisos_permiso_idx ON cargo_permisos (permiso)")
    cur.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS nombre TEXT")
    cur.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS cargo_id BIGINT REFERENCES cargos(id)")
    cur.execute("CREATE INDEX IF NOT EXISTS usuarios_cargo_id_idx ON usuarios (cargo_id)")


def obtener_config_empresa():
    try:
        with conectar() as conn:
            with conn.cursor() as cur:
                asegurar_configuracion_sistema(cur)
                cur.execute(
                    """
                    SELECT nombre, subtitulo, COALESCE(logo_path, '')
                    FROM empresa_config
                    WHERE id = 1
                    """
                )
                fila = cur.fetchone()
    except Exception:
        fila = None

    if not fila:
        return {
            "nombre": EMPRESA_NOMBRE,
            "subtitulo": EMPRESA_SUBTITULO,
            "logo_path": resolver_recurso("logo.png"),
        }

    logo_path = fila[2] or resolver_recurso("logo.png")
    return {"nombre": fila[0], "subtitulo": fila[1], "logo_path": logo_path}


def guardar_config_empresa(nombre, subtitulo, logo_path):
    nombre = nombre.strip() or EMPRESA_NOMBRE
    subtitulo = subtitulo.strip() or EMPRESA_SUBTITULO
    logo_path = logo_path.strip()

    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_configuracion_sistema(cur)
            cur.execute(
                """
                UPDATE empresa_config
                SET nombre = %s,
                    subtitulo = %s,
                    logo_path = %s,
                    updated_at = now()
                WHERE id = 1
                """,
                (nombre, subtitulo, logo_path),
            )


def obtener_logo_empresa():
    config = obtener_config_empresa()
    return config.get("logo_path") or resolver_recurso("logo.png")


def construir_contenido_qr(resumen):
    return construir_url_validacion(resumen) or construir_texto_validacion(resumen)


def generar_imagen_qr(datos):
    try:
        import qrcode
    except ImportError as exc:
        raise RuntimeError(
            "Falta instalar la dependencia qrcode. Ejecute: pip install qrcode"
        ) from exc

    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=9,
        border=3,
    )
    qr.add_data(datos)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


def ruta_temporal_qr(codigo):
    nombre_archivo = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in codigo)
    return os.path.join(tempfile.gettempdir(), f"GCDRYS_QR_{nombre_archivo}.png")


def copiar_imagen_al_portapapeles(ruta_imagen):
    ruta_segura = ruta_imagen.replace("'", "''")
    comando = (
        "Add-Type -AssemblyName System.Windows.Forms; "
        "Add-Type -AssemblyName System.Drawing; "
        f"$img=[System.Drawing.Image]::FromFile('{ruta_segura}'); "
        "[System.Windows.Forms.Clipboard]::SetImage($img); "
        "$img.Dispose()"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-STA", "-Command", comando],
        check=True,
        capture_output=True,
        text=True,
    )


def abrir_validacion_qr(resumen, parent=None):
    datos_qr = construir_contenido_qr(resumen)

    try:
        imagen_qr = generar_imagen_qr(datos_qr)
    except Exception as e:
        messagebox.showerror("Error QR", str(e))
        return

    ruta_qr = ruta_temporal_qr(resumen["codigo"])
    imagen_qr.save(ruta_qr)

    win = tk.Toplevel(parent)
    win.title("Validacion Documental QR")
    aplicar_icono(win)
    centrar_ventana(win, 540, 640)
    win.configure(bg="white")
    win.minsize(420, 480)

    scroll_area = tk.Frame(win, bg="white")
    scroll_area.pack(side="top", fill="both", expand=True)
    canvas = tk.Canvas(scroll_area, bg="white", highlightthickness=0)
    scrollbar = ttk.Scrollbar(scroll_area, orient="vertical", command=canvas.yview)
    contenido = tk.Frame(canvas, bg="white")
    contenido_id = canvas.create_window((0, 0), window=contenido, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def ajustar_scroll(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfigure(contenido_id, width=canvas.winfo_width())

    contenido.bind("<Configure>", ajustar_scroll)
    canvas.bind("<Configure>", ajustar_scroll)
    canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

    logo_tk = cargar_logo("logo.png", (95, 95))
    if logo_tk:
        logo = tk.Label(contenido, image=logo_tk, bg="white")
        logo.image = logo_tk
        logo.pack(pady=(18, 6))

    color_estado = "#1f8f45" if resumen["vigencia"] == "VIGENTE" else "#c0392b"

    tk.Label(
        contenido,
        text=resumen["vigencia"],
        bg=color_estado,
        fg="white",
        font=("Arial", 18, "bold"),
        width=22,
        pady=8,
    ).pack(pady=(6, 12))

    tk.Label(
        contenido,
        text=resumen["nombre"],
        bg="white",
        fg="#222222",
        font=("Arial", 13, "bold"),
        wraplength=450,
        justify="center",
    ).pack(padx=20, pady=(0, 6))

    tk.Label(
        contenido,
        text=resumen["codigo"],
        bg="white",
        fg="#2d5bd1",
        font=("Arial", 12, "bold"),
    ).pack(pady=(0, 12))

    qr_tk = ImageTk.PhotoImage(imagen_qr.resize((210, 210)))
    qr_label = tk.Label(contenido, image=qr_tk, bg="white")
    qr_label.image = qr_tk
    qr_label.pack(pady=(0, 14))

    datos_frame = tk.Frame(contenido, bg="white")
    datos_frame.pack(fill="x", padx=42, pady=(0, 12))
    datos_frame.grid_columnconfigure(1, weight=1)

    filas = [
        ("Estado", resumen["estado"]),
        ("Creado por", resumen["creado_por"]),
        ("Fecha creacion", formatear_fecha(resumen["fecha_creacion"])),
        ("Revisado por", resumen["revisado_por"]),
        ("Fecha revision", formatear_fecha(resumen["fecha_revision"]) or "Pendiente"),
        ("Aprobado por", resumen["aprobado_por"]),
        ("Fecha aprobacion", formatear_fecha(resumen["fecha_aprobacion"]) or "Pendiente"),
    ]

    for idx, (etiqueta, valor) in enumerate(filas):
        tk.Label(datos_frame, text=etiqueta, bg="white", fg="gray", font=("Arial", 10)).grid(
            row=idx, column=0, sticky="e", padx=(0, 10), pady=3
        )
        tk.Label(datos_frame, text=valor, bg="white", fg="#222222", font=("Arial", 10), wraplength=270).grid(
            row=idx, column=1, sticky="w", pady=3
        )

    estado_copia = tk.StringVar(value="")

    def copiar_qr():
        try:
            copiar_imagen_al_portapapeles(ruta_qr)
            estado_copia.set("QR copiado al portapapeles")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo copiar el QR: {e}")

    def guardar_qr():
        destino = filedialog.asksaveasfilename(
            parent=win,
            title="Guardar QR",
            defaultextension=".png",
            filetypes=[("Imagen PNG", "*.png")],
            initialfile=f"QR_{resumen['codigo']}.png",
        )

        if not destino:
            return

        imagen_qr.save(destino)
        estado_copia.set(f"QR guardado en {destino}")

    acciones = tk.Frame(win, bg="white", bd=1, relief="solid")
    acciones.pack(side="bottom", fill="x")
    botones = tk.Frame(acciones, bg="white")
    botones.pack(pady=(8, 4))

    tk.Button(botones, text="Copiar QR", command=copiar_qr, width=16).grid(row=0, column=0, padx=6)
    tk.Button(botones, text="Guardar PNG", command=guardar_qr, width=16).grid(row=0, column=1, padx=6)
    tk.Button(botones, text="Cerrar", command=win.destroy, width=12).grid(row=0, column=2, padx=6)

    tk.Label(acciones, textvariable=estado_copia, bg="white", fg="green", font=("Arial", 9)).pack(pady=(0, 8))


def normalizar_rol(rol):
    return rol if rol in ROLES else "user"


def es_superusuario(rol):
    return normalizar_rol(rol) == "Superusuario"


def permisos_por_rol(rol):
    return set(PERMISOS_ROL_BASE.get(normalizar_rol(rol), ()))


def obtener_permisos_cargo(cargo_id):
    if not cargo_id:
        return set(), None

    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_configuracion_sistema(cur)
            cur.execute(
                """
                SELECT rol_base
                FROM cargos
                WHERE id = %s
                AND activo = true
                """,
                (cargo_id,),
            )
            fila = cur.fetchone()
            rol_base = fila[0] if fila else None
            cur.execute(
                """
                SELECT permiso
                FROM cargo_permisos
                WHERE cargo_id = %s
                AND permitido = true
                """,
                (cargo_id,),
            )
            permisos = {permiso for (permiso,) in cur.fetchall() if permiso in PERMISOS_TODOS}

    return permisos, normalizar_rol(rol_base) if rol_base else None


def obtener_permisos_usuario(user_id, rol):
    rol = normalizar_rol(rol)

    if es_superusuario(rol):
        return set(PERMISOS_TODOS)

    permisos = permisos_por_rol(rol)

    if not user_id:
        return permisos

    try:
        with conectar() as conn:
            with conn.cursor() as cur:
                asegurar_configuracion_sistema(cur)
                cur.execute(
                    """
                    SELECT cargo_id
                    FROM usuarios
                    WHERE id = %s
                    """,
                    (user_id,),
                )
                fila = cur.fetchone()

        cargo_id = fila[0] if fila else None
        permisos_cargo, rol_cargo = obtener_permisos_cargo(cargo_id)

        if rol_cargo:
            permisos.update(permisos_por_rol(rol_cargo))

        permisos.update(permisos_cargo)
    except Exception:
        pass

    return permisos


def usuario_tiene_permiso(user_id, rol, permiso):
    return permiso in obtener_permisos_usuario(user_id, rol)


def puede_administrar_usuarios(rol, user_id=None):
    return usuario_tiene_permiso(user_id, rol, "usuarios.administrar")


def puede_editar_codigo_documental(rol, user_id=None):
    return usuario_tiene_permiso(user_id, rol, "documentos.codigo.editar")


def puede_crear_documentos(rol, user_id=None):
    return usuario_tiene_permiso(user_id, rol, "documentos.crear")


def puede_editar_documentos(rol, user_id=None):
    permisos = obtener_permisos_usuario(user_id, rol)
    return any(permiso in permisos for permiso in PERMISO_EDICION_ESTADO.values()) or "documentos.codigo.editar" in permisos


def puede_ver_historial(rol, user_id=None):
    return usuario_tiene_permiso(user_id, rol, "documentos.historial.ver")


def normalizar_estado(estado):
    return estado if estado in ESTADOS_DOCUMENTO else "Borrador"


def version_a_numero(version):
    texto = str(version or "").strip().upper()

    if texto.startswith("V"):
        texto = texto[1:].strip()

    try:
        return int(texto)
    except ValueError:
        return None


def obtener_versiones_vigentes_relacionadas(cur, area, tipo, numero, documento_id=None):
    parametros = [area, tipo, numero]
    filtro_id = ""

    if documento_id:
        filtro_id = "AND id <> %s"
        parametros.append(documento_id)

    cur.execute(
        f"""
        SELECT id, codigo, version
        FROM documentos
        WHERE area = %s
        AND tipo = %s
        AND numero = %s
        {filtro_id}
        AND COALESCE(estado, 'Borrador') = 'Vigente'
        ORDER BY fecha DESC, id DESC
        """,
        tuple(parametros),
    )
    return cur.fetchall()


def evaluar_aprobacion_version(area, tipo, numero, version, documento_id=None):
    version_nueva = version_a_numero(version)

    if version_nueva is None:
        return {
            "ok": False,
            "mensaje": "La version debe ser numerica para aprobar el documento.",
            "vigentes": [],
        }

    with conectar() as conn:
        with conn.cursor() as cur:
            vigentes = obtener_versiones_vigentes_relacionadas(cur, area, tipo, numero, documento_id)

    bloqueos = []
    reemplazos = []

    for _, codigo_vigente, version_vigente in vigentes:
        version_actual = version_a_numero(version_vigente)

        if version_actual is not None and version_nueva <= version_actual:
            bloqueos.append(f"{codigo_vigente} (V{version_vigente})")
        else:
            reemplazos.append(f"{codigo_vigente} (V{version_vigente})")

    if bloqueos:
        return {
            "ok": False,
            "mensaje": "Ya existe una version vigente igual o superior: " + ", ".join(bloqueos),
            "vigentes": vigentes,
        }

    mensaje = ""
    if reemplazos:
        mensaje = "Al aprobar, se marcara como obsoleta la version vigente anterior: " + ", ".join(reemplazos)

    return {
        "ok": True,
        "mensaje": mensaje,
        "vigentes": vigentes,
    }


def estados_creacion_permitidos(rol):
    return ("Borrador",)


def puede_editar_metadatos_documento(rol, estado_actual, user_id=None):
    if puede_editar_codigo_documental(rol, user_id):
        return True

    permisos = obtener_permisos_usuario(user_id, rol)
    return (
        "documentos.nombre.editar" in permisos
        and PERMISO_EDICION_ESTADO.get(normalizar_estado(estado_actual)) in permisos
    )


def puede_cambiar_estado_documental(rol, estado_actual, estado_nuevo, user_id=None):
    rol = normalizar_rol(rol)
    estado_actual = normalizar_estado(estado_actual)
    estado_nuevo = normalizar_estado(estado_nuevo)

    if estado_actual == estado_nuevo:
        return True

    permiso = PERMISO_TRANSICION_ESTADO.get((estado_actual, estado_nuevo))

    if permiso:
        return usuario_tiene_permiso(user_id, rol, permiso)

    transiciones = {
        "Superusuario": {
            "Borrador": ("En revision", "Anulado"),
            "En revision": ("Borrador", "Pendiente aprobacion", "Anulado"),
            "Pendiente aprobacion": ("En revision", "Vigente", "Anulado"),
            "Vigente": ("Obsoleto", "Anulado"),
            "Obsoleto": (),
            "Anulado": (),
        },
        "admin": {
            "Borrador": ("En revision", "Anulado"),
            "En revision": ("Borrador", "Pendiente aprobacion", "Anulado"),
            "Pendiente aprobacion": ("En revision", "Vigente", "Anulado"),
            "Vigente": ("Obsoleto", "Anulado"),
            "Obsoleto": (),
            "Anulado": (),
        },
        "colaborador": {
            "Borrador": ("En revision",),
            "En revision": (),
            "Pendiente aprobacion": (),
            "Vigente": (),
            "Obsoleto": (),
            "Anulado": (),
        },
    }

    return estado_nuevo in transiciones.get(rol, {}).get(estado_actual, ())


def estados_edicion_permitidos(rol, estado_actual, user_id=None):
    estado_actual = normalizar_estado(estado_actual)
    return tuple(
        estado
        for estado in ESTADOS_DOCUMENTO
        if estado == estado_actual or puede_cambiar_estado_documental(rol, estado_actual, estado, user_id)
    )


# =========================
# USUARIOS
# =========================
def registrar_usuario(username, password, email, rol="user", nombre="", cargo_id=None):
    username = username.strip()
    email = email.strip()
    nombre = nombre.strip()
    rol = normalizar_rol(rol)

    if not username or not password:
        messagebox.showerror("Error", "Debe completar usuario y contrasena")
        return False

    error_password = validar_password_plana(password)
    if error_password:
        messagebox.showerror("Error", error_password)
        return False

    try:
        with conectar() as conn:
            with conn.cursor() as cur:
                asegurar_configuracion_sistema(cur)
                cur.execute(
                    """
                    INSERT INTO usuarios (username, password, email, rol, activo, nombre, cargo_id)
                    VALUES (%s, %s, %s, %s, true, %s, %s)
                    """,
                    (username, hash_password(password), email, rol, nombre, cargo_id),
                )

        messagebox.showinfo("Exito", "Usuario registrado")
        return True

    except Exception as e:
        messagebox.showerror("Error", str(e))
        return False


def validar_usuario(username, password):
    username = username.strip()

    try:
        with conectar() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT
                        id,
                        username,
                        COALESCE(rol, 'user') AS rol,
                        password
                    FROM usuarios
                    WHERE username = %s
                    AND COALESCE(activo, true) = true
                    """,
                    (username,),
                )
                resultado = cur.fetchone()

                if not resultado:
                    return None

                user_id, usuario, rol, hash_guardado = resultado
                password_ok, requiere_actualizacion = verificar_password(password, hash_guardado)

                if not password_ok:
                    return None

                if requiere_actualizacion:
                    cur.execute(
                        """
                        UPDATE usuarios
                        SET password = %s
                        WHERE id = %s
                        """,
                        (hash_password(password), user_id),
                    )

                return user_id, usuario, rol

    except Exception as e:
        messagebox.showerror("Error de conexion", str(e))
        return None


# =========================
# DOCUMENTOS
# =========================
def codigo_existe_db(codigo):
    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT 1
                FROM documentos
                WHERE codigo = %s
                """,
                (codigo,),
            )
            return cur.fetchone() is not None


def obtener_siguiente_numero(area, tipo):
    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT numero
                FROM documentos
                WHERE area = %s
                AND tipo = %s
                """,
                (area, tipo),
            )
            numeros = cur.fetchall()

    max_num = 0

    for n in numeros:
        try:
            num = int(n[0])
            if num > max_num:
                max_num = num
        except (TypeError, ValueError):
            pass

    return str(max_num + 1).zfill(3)


def asegurar_columna_confidencial(cur):
    cur.execute(
        """
        ALTER TABLE documentos
        ADD COLUMN IF NOT EXISTS confidencial BOOLEAN NOT NULL DEFAULT false
        """
    )
    cur.execute(
        """
        ALTER TABLE documentos
        ADD COLUMN IF NOT EXISTS licencia_cliente TEXT
        """
    )
    cur.execute(
        """
        ALTER TABLE documentos
        ADD COLUMN IF NOT EXISTS licencia_key_hash TEXT
        """
    )


def guardar_en_db(codigo, nombre, area, tipo, numero, version, estado, user_id, confidencial=False):
    estado = normalizar_estado(estado)
    licencia = obtener_estado_licencia()
    licencia_cliente = licencia.get("cliente") or "Licencia no activada"
    licencia_key_hash = licencia.get("license_key_hash") or hash_licencia(licencia.get("license_key", ""))

    try:
        with conectar() as conn:
            with conn.cursor() as cur:
                asegurar_columna_confidencial(cur)
                cur.execute(
                    """
                    INSERT INTO documentos
                    (
                        codigo,
                        nombre,
                        area,
                        tipo,
                        numero,
                        version,
                        estado,
                        creado_por,
                        confidencial,
                        licencia_cliente,
                        licencia_key_hash
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    RETURNING id
                    """,
                    (
                        codigo,
                        nombre,
                        area,
                        tipo,
                        numero,
                        version,
                        estado,
                        user_id,
                        bool(confidencial),
                        licencia_cliente,
                        licencia_key_hash,
                    ),
                )
                documento_id = cur.fetchone()[0]
                registrar_historial(
                    cur,
                    documento_id,
                    codigo,
                    "creacion",
                    "",
                    f"{codigo} | {nombre} | {estado} | Confidencial: {'Si' if confidencial else 'No'} | Licencia: {licencia_cliente}",
                    user_id,
                    "Documento creado",
                )

        messagebox.showinfo("Exito", f"Documento guardado:\n{codigo}\n{nombre}")
        return True

    except Exception as e:
        if getattr(e, "pgcode", "") == "23505":
            messagebox.showerror("Duplicado", "El codigo documental ya existe en la base de datos")
        else:
            messagebox.showerror("Error", str(e))
        return False


def obtener_documentos(area=None, tipo=None, estado=None, texto=None, incluir_id=False):
    columnas = "codigo, nombre, area, tipo, numero, version, COALESCE(estado, 'Borrador') AS estado, fecha, COALESCE(confidencial, false) AS confidencial"

    if incluir_id:
        columnas = "id, " + columnas

    query = """
        SELECT {columnas}
        FROM documentos
        WHERE 1=1
    """.format(columnas=columnas)
    parametros = []

    if area:
        query += " AND area = %s"
        parametros.append(area)

    if tipo:
        query += " AND tipo = %s"
        parametros.append(tipo)

    if estado:
        query += " AND COALESCE(estado, 'Borrador') = %s"
        parametros.append(estado)

    if texto:
        query += " AND (codigo ILIKE %s OR nombre ILIKE %s)"
        parametros.extend([f"%{texto}%", f"%{texto}%"])

    query += " ORDER BY fecha DESC"

    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_columna_confidencial(cur)
            cur.execute(query, tuple(parametros))
            return cur.fetchall()


def obtener_documento_por_id(documento_id):
    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_columna_confidencial(cur)
            cur.execute(
                """
                SELECT id, codigo, nombre, area, tipo, numero, version, COALESCE(estado, 'Borrador') AS estado, fecha, COALESCE(confidencial, false) AS confidencial
                FROM documentos
                WHERE id = %s
                """,
                (documento_id,),
            )
            return cur.fetchone()


def obtener_ficha_documento(documento_id):
    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_columna_confidencial(cur)
            cur.execute(
                """
                SELECT
                    d.id,
                    d.codigo,
                    d.nombre,
                    d.area,
                    COALESCE(a.nombre, d.area) AS area_nombre,
                    d.tipo,
                    COALESCE(t.nombre, d.tipo) AS tipo_nombre,
                    d.numero,
                    d.version,
                    COALESCE(d.estado, 'Borrador') AS estado,
                    d.fecha,
                    COALESCE(u.username, 'usuario desconocido') AS creado_por,
                    COALESCE(d.confidencial, false) AS confidencial
                FROM documentos d
                LEFT JOIN areas a ON a.codigo = d.area
                LEFT JOIN tipos_documento t ON t.codigo = d.tipo
                LEFT JOIN usuarios u ON u.id = d.creado_por
                WHERE d.id = %s
                """,
                (documento_id,),
            )
            return cur.fetchone()


def codigo_existe_en_otro_documento(codigo, documento_id):
    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT 1
                FROM documentos
                WHERE codigo = %s
                AND id <> %s
                """,
                (codigo, documento_id),
            )
            return cur.fetchone() is not None


def registrar_historial(cur, documento_id, codigo, campo, anterior, nuevo, user_id, observacion=""):
    cur.execute(
        """
        INSERT INTO documentos_historial
        (
            documento_id,
            codigo,
            campo,
            valor_anterior,
            valor_nuevo,
            observacion,
            cambiado_por
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        """,
        (
            documento_id,
            codigo,
            campo,
            "" if anterior is None else str(anterior),
            "" if nuevo is None else str(nuevo),
            observacion.strip(),
            user_id,
        ),
    )


def obsoletar_versiones_vigentes_anteriores(cur, datos_documento, documento_id, user_id):
    cur.execute(
        """
        UPDATE documentos
        SET estado = 'Obsoleto'
        WHERE area = %s
        AND tipo = %s
        AND numero = %s
        AND id <> %s
        AND COALESCE(estado, 'Borrador') = 'Vigente'
        RETURNING id, codigo
        """,
        (
            datos_documento["area"],
            datos_documento["tipo"],
            datos_documento["numero"],
            documento_id,
        ),
    )
    documentos_obsoletos = cur.fetchall()

    for doc_obsoleto_id, codigo_obsoleto in documentos_obsoletos:
        registrar_historial(
            cur,
            doc_obsoleto_id,
            codigo_obsoleto,
            "estado",
            "Vigente",
            "Obsoleto",
            user_id,
            f"Obsoletado automaticamente por aprobacion de {datos_documento['codigo']}",
        )

    return documentos_obsoletos


def validar_version_mayor_que_vigente(cur, datos_documento, documento_id):
    version_nueva = version_a_numero(datos_documento["version"])

    if version_nueva is None:
        raise ValueError("La version debe ser numerica para aprobar el documento")

    vigentes = obtener_versiones_vigentes_relacionadas(
        cur,
        datos_documento["area"],
        datos_documento["tipo"],
        datos_documento["numero"],
        documento_id,
    )

    for _, codigo_vigente, version_vigente in vigentes:
        version_actual = version_a_numero(version_vigente)

        if version_actual is not None and version_nueva <= version_actual:
            raise ValueError(
                f"No se puede aprobar la version {datos_documento['version']}. "
                f"Ya existe una version vigente igual o superior: {codigo_vigente} (V{version_vigente})"
            )


def describir_documentos_reemplazados(documentos_obsoletos):
    if not documentos_obsoletos:
        return ""

    codigos = [codigo for _, codigo in documentos_obsoletos]
    return "Reemplaza version vigente anterior: " + ", ".join(codigos)


def actualizar_documento(documento_id, datos_nuevos, user_id, rol):
    actual = obtener_documento_por_id(documento_id)

    if not actual:
        raise ValueError("Documento no encontrado")

    campos = ("id", "codigo", "nombre", "area", "tipo", "numero", "version", "estado", "fecha", "confidencial")
    datos_actuales = dict(zip(campos, actual))

    if not puede_editar_documentos(rol, user_id):
        raise PermissionError("No tiene permisos para editar documentos")

    estado_actual = normalizar_estado(datos_actuales["estado"])
    observacion = str(datos_nuevos.get("observacion", "")).strip()
    puede_editar_metadatos = puede_editar_metadatos_documento(rol, estado_actual, user_id)

    campos_editables = ["estado"]
    permisos_usuario = obtener_permisos_usuario(user_id, rol)

    if puede_editar_metadatos:
        if "documentos.nombre.editar" in permisos_usuario:
            campos_editables.append("nombre")
        if "documentos.confidencial.editar" in permisos_usuario:
            campos_editables.append("confidencial")

    if puede_editar_codigo_documental(rol, user_id):
        campos_editables = ["codigo", "nombre", "area", "tipo", "numero", "version", "estado", "confidencial"]

    cambios = {}

    for campo in campos_editables:
        nuevo = str(datos_nuevos.get(campo, "")).strip()

        if campo == "confidencial":
            nuevo = bool(datos_nuevos.get(campo))

        if campo == "numero":
            nuevo = nuevo.zfill(3)

        if campo == "estado":
            nuevo = normalizar_estado(nuevo)

        if campo in ("codigo", "nombre", "area", "tipo", "numero", "version", "estado") and not nuevo:
            raise ValueError(f"El campo {campo} no puede quedar vacio")

        anterior = bool(datos_actuales[campo]) if campo == "confidencial" else "" if datos_actuales[campo] is None else str(datos_actuales[campo])

        if nuevo != anterior:
            cambios[campo] = (anterior, nuevo)

    if "codigo" in cambios and codigo_existe_en_otro_documento(cambios["codigo"][1], documento_id):
        raise ValueError("El codigo ingresado ya existe en otro documento")

    if "estado" in cambios:
        estado_nuevo = cambios["estado"][1]

        if not puede_cambiar_estado_documental(rol, estado_actual, estado_nuevo, user_id):
            raise PermissionError(f"No tiene permisos para cambiar de {estado_actual} a {estado_nuevo}")

        if not observacion:
            raise ValueError("Debe ingresar una observacion para cambiar el estado documental")

    cambios_metadatos = [campo for campo in cambios if campo != "estado"]

    if cambios_metadatos and not puede_editar_metadatos_documento(rol, estado_actual, user_id):
        raise PermissionError(f"No se pueden editar metadatos de un documento en estado {estado_actual}")

    if not cambios:
        return False

    set_sql = ", ".join([f"{campo} = %s" for campo in cambios.keys()])
    valores = [nuevo for _, nuevo in cambios.values()]
    valores.append(documento_id)

    with conectar() as conn:
        with conn.cursor() as cur:
            datos_version_vigente = None

            if cambios.get("estado", ("", ""))[1] == "Vigente":
                datos_version_vigente = datos_actuales.copy()

                for campo, (_, nuevo) in cambios.items():
                    datos_version_vigente[campo] = nuevo

                validar_version_mayor_que_vigente(cur, datos_version_vigente, documento_id)

            cur.execute(
                f"""
                UPDATE documentos
                SET {set_sql}
                WHERE id = %s
                """,
                tuple(valores),
            )

            codigo_historial = cambios.get("codigo", (datos_actuales["codigo"], datos_actuales["codigo"]))[1]

            for campo, (anterior, nuevo) in cambios.items():
                anterior_historial = "Si" if campo == "confidencial" and anterior else "No" if campo == "confidencial" else anterior
                nuevo_historial = "Si" if campo == "confidencial" and nuevo else "No" if campo == "confidencial" else nuevo
                registrar_historial(
                    cur,
                    documento_id,
                    codigo_historial,
                    campo,
                    anterior_historial,
                    nuevo_historial,
                    user_id,
                    observacion if campo == "estado" else "",
                )

            if datos_version_vigente:
                documentos_obsoletos = obsoletar_versiones_vigentes_anteriores(
                    cur,
                    datos_version_vigente,
                    documento_id,
                    user_id,
                )
                descripcion_reemplazo = describir_documentos_reemplazados(documentos_obsoletos)

                if descripcion_reemplazo:
                    registrar_historial(
                        cur,
                        documento_id,
                        codigo_historial,
                        "version",
                        "",
                        descripcion_reemplazo,
                        user_id,
                        "Version documental aprobada",
                    )

    return True


def obtener_historial_documento(documento_id):
    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT
                    h.fecha,
                    h.campo,
                    h.valor_anterior,
                    h.valor_nuevo,
                    COALESCE(h.observacion, '') AS observacion,
                    COALESCE(u.username, 'usuario desconocido') AS usuario
                FROM documentos_historial h
                LEFT JOIN usuarios u ON u.id = h.cambiado_por
                WHERE h.documento_id = %s
                ORDER BY h.fecha DESC
                """,
                (documento_id,),
            )
            return cur.fetchall()


def asegurar_tabla_auditoria_reportes(cur):
    asegurar_columna_confidencial(cur)
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS documentos_reportes_log (
            id BIGSERIAL PRIMARY KEY,
            documento_id BIGINT REFERENCES documentos(id) ON DELETE CASCADE,
            codigo TEXT,
            accion TEXT NOT NULL,
            tipo_reporte TEXT NOT NULL,
            archivo TEXT,
            equipo TEXT,
            usuario_sistema TEXT,
            usuario_id BIGINT REFERENCES usuarios(id),
            fecha TIMESTAMPTZ NOT NULL DEFAULT now()
        )
        """
    )
    cur.execute(
        """
        ALTER TABLE documentos_reportes_log
        DROP CONSTRAINT IF EXISTS documentos_reportes_log_accion_check
        """
    )
    cur.execute(
        """
        ALTER TABLE documentos_reportes_log
        ADD CONSTRAINT documentos_reportes_log_accion_check
        CHECK (accion IN ('visualizacion', 'exportacion', 'impresion'))
        """
    )
    cur.execute(
        """
        CREATE INDEX IF NOT EXISTS documentos_reportes_log_documento_id_idx
        ON documentos_reportes_log (documento_id)
        """
    )
    cur.execute(
        """
        CREATE INDEX IF NOT EXISTS documentos_reportes_log_fecha_idx
        ON documentos_reportes_log (fecha DESC)
        """
    )


def registrar_evento_reporte(documentos, accion, tipo_reporte, archivo, user_id):
    equipo = obtener_nombre_equipo()
    usuario_sistema = obtener_usuario_sistema()

    if not documentos:
        documentos = [(None, "")]

    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_tabla_auditoria_reportes(cur)

            for documento_id, codigo in documentos:
                cur.execute(
                    """
                    INSERT INTO documentos_reportes_log
                    (
                        documento_id,
                        codigo,
                        accion,
                        tipo_reporte,
                        archivo,
                        equipo,
                        usuario_sistema,
                        usuario_id
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        documento_id,
                        codigo,
                        accion,
                        tipo_reporte,
                        archivo,
                        equipo,
                        usuario_sistema,
                        user_id,
                    ),
                )


def obtener_eventos_reporte_documento(documento_id):
    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_tabla_auditoria_reportes(cur)
            cur.execute(
                """
                SELECT
                    l.fecha,
                    l.accion,
                    l.tipo_reporte,
                    COALESCE(u.username, 'usuario desconocido') AS usuario_app,
                    COALESCE(l.usuario_sistema, '') AS usuario_sistema,
                    COALESCE(l.equipo, '') AS equipo,
                    COALESCE(l.archivo, '') AS archivo
                FROM documentos_reportes_log l
                LEFT JOIN usuarios u ON u.id = l.usuario_id
                WHERE l.documento_id = %s
                ORDER BY l.fecha DESC
                """,
                (documento_id,),
            )
            return cur.fetchall()


def obtener_metricas_dashboard():
    metricas = {
        "total": 0,
        "mes_actual": 0,
        "por_estado": [],
        "por_area": [],
        "por_tipo": [],
        "ultimos_documentos": [],
        "ultimos_cambios": [],
    }

    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM documentos")
            metricas["total"] = cur.fetchone()[0]

            cur.execute(
                """
                SELECT COUNT(*)
                FROM documentos
                WHERE fecha >= date_trunc('month', now())
                """
            )
            metricas["mes_actual"] = cur.fetchone()[0]

            cur.execute(
                """
                SELECT COALESCE(estado, 'Borrador') AS estado, COUNT(*)
                FROM documentos
                GROUP BY COALESCE(estado, 'Borrador')
                ORDER BY COUNT(*) DESC, estado
                """
            )
            metricas["por_estado"] = cur.fetchall()

            cur.execute(
                """
                SELECT d.area, COALESCE(a.nombre, d.area) AS nombre, COUNT(*)
                FROM documentos d
                LEFT JOIN areas a ON a.codigo = d.area
                GROUP BY d.area, COALESCE(a.nombre, d.area)
                ORDER BY COUNT(*) DESC, d.area
                LIMIT 8
                """
            )
            metricas["por_area"] = cur.fetchall()

            cur.execute(
                """
                SELECT d.tipo, COALESCE(t.nombre, d.tipo) AS nombre, COUNT(*)
                FROM documentos d
                LEFT JOIN tipos_documento t ON t.codigo = d.tipo
                GROUP BY d.tipo, COALESCE(t.nombre, d.tipo)
                ORDER BY COUNT(*) DESC, d.tipo
                LIMIT 8
                """
            )
            metricas["por_tipo"] = cur.fetchall()

            cur.execute(
                """
                SELECT codigo, nombre, COALESCE(estado, 'Borrador') AS estado, fecha
                FROM documentos
                ORDER BY fecha DESC
                LIMIT 10
                """
            )
            metricas["ultimos_documentos"] = cur.fetchall()

            cur.execute(
                """
                SELECT h.fecha, h.codigo, h.campo, COALESCE(u.username, 'usuario desconocido') AS usuario
                FROM documentos_historial h
                LEFT JOIN usuarios u ON u.id = h.cambiado_por
                ORDER BY h.fecha DESC
                LIMIT 10
                """
            )
            metricas["ultimos_cambios"] = cur.fetchall()

    return metricas


# =========================
# EXPORTAR EXCEL
# =========================
def exportar_excel():
    try:
        datos = obtener_documentos()

        wb = Workbook()
        ws = wb.active
        ws.title = "Documentos"
        wb.properties.creator = EMPRESA_NOMBRE
        wb.properties.title = f"GCDRYS-2 {APP_VERSION} - Documentos"
        wb.properties.subject = resumen_licencia_para_metadatos()

        ws.append(
            [
                "Codigo",
                "Nombre Documento",
                "Area",
                "Tipo",
                "Numero",
                "Version",
                "Estado",
                "Fecha",
                "Confidencial",
            ]
        )

        for fila in datos:
            fila_excel = list(fila)
            fila_excel[-1] = "Si" if fila_excel[-1] else "No"
            ws.append(fila_excel)

        archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="documentos_exportados.xlsx",
        )

        if archivo:
            wb.save(archivo)
            messagebox.showinfo("Exportado", "Archivo Excel generado correctamente")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================
# EXPORTAR PDF
# =========================
def cargar_reportlab():
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.platypus import (
            Image as PdfImage,
            Paragraph,
            SimpleDocTemplate,
            Spacer,
            Table,
            TableStyle,
        )

        return {
            "colors": colors,
            "TA_CENTER": TA_CENTER,
            "TA_LEFT": TA_LEFT,
            "A4": A4,
            "landscape": landscape,
            "ParagraphStyle": ParagraphStyle,
            "getSampleStyleSheet": getSampleStyleSheet,
            "cm": cm,
            "PdfImage": PdfImage,
            "Paragraph": Paragraph,
            "SimpleDocTemplate": SimpleDocTemplate,
            "Spacer": Spacer,
            "Table": Table,
            "TableStyle": TableStyle,
        }
    except ImportError as exc:
        raise RuntimeError(
            "Falta instalar la libreria reportlab. Ejecute: pip install reportlab"
        ) from exc


def construir_estilos_pdf(rl):
    estilos_base = rl["getSampleStyleSheet"]()
    return {
        "titulo": rl["ParagraphStyle"](
            "TituloCorporativo",
            parent=estilos_base["Title"],
            fontName="Helvetica-Bold",
            fontSize=16,
            leading=20,
            textColor=rl["colors"].HexColor("#1f4e79"),
            alignment=rl["TA_CENTER"],
            spaceAfter=4,
        ),
        "subtitulo": rl["ParagraphStyle"](
            "SubtituloCorporativo",
            parent=estilos_base["Normal"],
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            textColor=rl["colors"].HexColor("#555555"),
            alignment=rl["TA_CENTER"],
        ),
        "seccion": rl["ParagraphStyle"](
            "SeccionCorporativa",
            parent=estilos_base["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=14,
            textColor=rl["colors"].HexColor("#1f4e79"),
            spaceBefore=8,
            spaceAfter=6,
        ),
        "celda": rl["ParagraphStyle"](
            "CeldaTabla",
            parent=estilos_base["Normal"],
            fontName="Helvetica",
            fontSize=7,
            leading=9,
            textColor=rl["colors"].HexColor("#222222"),
            alignment=rl["TA_LEFT"],
        ),
        "celda_centrada": rl["ParagraphStyle"](
            "CeldaTablaCentrada",
            parent=estilos_base["Normal"],
            fontName="Helvetica",
            fontSize=7,
            leading=9,
            textColor=rl["colors"].HexColor("#222222"),
            alignment=rl["TA_CENTER"],
        ),
    }


def encabezado_pie_pdf(canvas, doc, rl):
    canvas.saveState()

    ancho, alto = doc.pagesize
    config_empresa = getattr(doc, "config_empresa", obtener_config_empresa())
    azul = rl["colors"].HexColor("#1f4e79")
    gris = rl["colors"].HexColor("#666666")
    margen_x = doc.leftMargin
    y_header = alto - 1.35 * rl["cm"]

    logo = config_empresa.get("logo_path") or resolver_recurso("logo.png")
    if logo:
        try:
            canvas.drawImage(
                logo,
                margen_x,
                y_header - 0.2 * rl["cm"],
                width=1.05 * rl["cm"],
                height=1.05 * rl["cm"],
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    canvas.setFillColor(azul)
    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawString(margen_x + 1.25 * rl["cm"], y_header + 0.32 * rl["cm"], config_empresa["nombre"])

    canvas.setFillColor(gris)
    canvas.setFont("Helvetica", 7)
    canvas.drawString(margen_x + 1.25 * rl["cm"], y_header, config_empresa["subtitulo"])

    canvas.setStrokeColor(azul)
    canvas.setLineWidth(1)
    canvas.line(margen_x, alto - 1.65 * rl["cm"], ancho - doc.rightMargin, alto - 1.65 * rl["cm"])

    canvas.setFillColor(gris)
    canvas.setFont("Helvetica", 7)
    detalle = getattr(doc, "detalle_generacion", f"Generado: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    canvas.drawString(margen_x, 0.8 * rl["cm"], detalle)
    canvas.setFont("Helvetica", 6.5)
    canvas.drawString(margen_x, 0.48 * rl["cm"], resumen_licencia_para_metadatos())
    canvas.drawRightString(ancho - doc.rightMargin, 0.8 * rl["cm"], f"Pagina {doc.page}")

    canvas.restoreState()


def guardar_pdf(nombre_sugerido):
    return filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
        initialfile=nombre_sugerido,
    )


def texto_pdf(valor):
    from xml.sax.saxutils import escape

    return escape(formatear_fecha(valor))


def construir_detalle_generacion(username):
    return (
        f"Generado: {datetime.now().strftime('%d-%m-%Y %H:%M')} | "
        f"Usuario app: {username or 'No identificado'} | "
        f"PC: {obtener_nombre_equipo()} | "
        f"Usuario Windows: {obtener_usuario_sistema()} | "
        f"{APP_VERSION}"
    )


def preparar_filas_documentos_pdf(datos):
    ids = []
    filas = []

    for fila in datos:
        if len(fila) == 10:
            ids.append((fila[0], fila[1]))
            filas.append(fila[1:])
        elif len(fila) == 9:
            filas.append(fila)
        elif len(fila) == 8:
            filas.append(tuple(fila) + (False,))
        else:
            filas.append(fila)

    return ids, filas


def generar_pdf_documentos(
    datos,
    titulo="Reporte Documental",
    subtitulo="Listado de documentos",
    archivo=None,
    username="",
    mostrar_mensaje=True,
):
    if archivo is None:
        archivo = guardar_pdf("reporte_documental.pdf")

    if not archivo:
        return None, []

    rl = cargar_reportlab()
    estilos = construir_estilos_pdf(rl)
    documentos_ids, filas_pdf = preparar_filas_documentos_pdf(datos)

    doc = rl["SimpleDocTemplate"](
        archivo,
        pagesize=rl["landscape"](rl["A4"]),
        rightMargin=1.1 * rl["cm"],
        leftMargin=1.1 * rl["cm"],
        topMargin=2.2 * rl["cm"],
        bottomMargin=1.4 * rl["cm"],
        title=titulo,
        author=EMPRESA_NOMBRE,
    )
    doc.detalle_generacion = construir_detalle_generacion(username)
    doc.config_empresa = obtener_config_empresa()

    elementos = [
        rl["Paragraph"](texto_pdf(titulo), estilos["titulo"]),
        rl["Paragraph"](texto_pdf(subtitulo), estilos["subtitulo"]),
        rl["Spacer"](1, 0.35 * rl["cm"]),
    ]

    columnas = ["Codigo", "Nombre Documento", "Area", "Tipo", "Numero", "Version", "Estado", "Fecha", "Conf."]
    tabla_datos = [[rl["Paragraph"](texto_pdf(col), estilos["celda_centrada"]) for col in columnas]]

    for fila in filas_pdf:
        fila_formateada = formatear_fila(fila)
        tabla_datos.append(
            [
                rl["Paragraph"](texto_pdf(fila_formateada[0]), estilos["celda"]),
                rl["Paragraph"](texto_pdf(fila_formateada[1]), estilos["celda"]),
                rl["Paragraph"](texto_pdf(fila_formateada[2]), estilos["celda_centrada"]),
                rl["Paragraph"](texto_pdf(fila_formateada[3]), estilos["celda_centrada"]),
                rl["Paragraph"](texto_pdf(fila_formateada[4]), estilos["celda_centrada"]),
                rl["Paragraph"](texto_pdf(fila_formateada[5]), estilos["celda_centrada"]),
                rl["Paragraph"](texto_pdf(fila_formateada[6]), estilos["celda_centrada"]),
                rl["Paragraph"](texto_pdf(fila_formateada[7]), estilos["celda_centrada"]),
                rl["Paragraph"]("Si" if fila[8] else "No", estilos["celda_centrada"]),
            ]
        )

    if len(tabla_datos) == 1:
        tabla_datos.append([rl["Paragraph"]("Sin documentos para mostrar", estilos["celda"])] + [""] * 8)

    tabla = rl["Table"](
        tabla_datos,
        repeatRows=1,
        colWidths=[3.7 * rl["cm"], 6.4 * rl["cm"], 1.5 * rl["cm"], 1.5 * rl["cm"], 1.5 * rl["cm"], 1.5 * rl["cm"], 2.3 * rl["cm"], 2.9 * rl["cm"], 1.2 * rl["cm"]],
    )
    tabla.setStyle(
        rl["TableStyle"](
            [
                ("BACKGROUND", (0, 0), (-1, 0), rl["colors"].HexColor("#1f4e79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), rl["colors"].white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.25, rl["colors"].HexColor("#c9d6e2")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl["colors"].white, rl["colors"].HexColor("#f4f7fb")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elementos.append(tabla)

    doc.build(
        elementos,
        onFirstPage=lambda canvas, doc: encabezado_pie_pdf(canvas, doc, rl),
        onLaterPages=lambda canvas, doc: encabezado_pie_pdf(canvas, doc, rl),
    )
    if mostrar_mensaje:
        messagebox.showinfo("Exportado", "Archivo PDF generado correctamente")
    return archivo, documentos_ids


def tabla_resumen_pdf(rl, titulo, columnas, filas, estilos):
    elementos = [rl["Paragraph"](texto_pdf(titulo), estilos["seccion"])]
    datos = [[rl["Paragraph"](texto_pdf(col), estilos["celda_centrada"]) for col in columnas]]

    for fila in filas:
        datos.append([rl["Paragraph"](texto_pdf(valor), estilos["celda"]) for valor in fila])

    if len(datos) == 1:
        datos.append([rl["Paragraph"]("Sin datos", estilos["celda"])] + [""] * (len(columnas) - 1))

    tabla = rl["Table"](datos, repeatRows=1, hAlign="LEFT")
    tabla.setStyle(
        rl["TableStyle"](
            [
                ("BACKGROUND", (0, 0), (-1, 0), rl["colors"].HexColor("#1f4e79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), rl["colors"].white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.25, rl["colors"].HexColor("#c9d6e2")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl["colors"].white, rl["colors"].HexColor("#f4f7fb")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elementos.append(tabla)
    elementos.append(rl["Spacer"](1, 0.25 * rl["cm"]))
    return elementos


def generar_pdf_ficha_documental(ficha, historial, archivo=None, username=""):
    if archivo is None:
        archivo = guardar_pdf(f"ficha_{ficha[1]}.pdf")

    if not archivo:
        return None

    rl = cargar_reportlab()
    estilos = construir_estilos_pdf(rl)
    doc = rl["SimpleDocTemplate"](
        archivo,
        pagesize=rl["A4"],
        rightMargin=1.4 * rl["cm"],
        leftMargin=1.4 * rl["cm"],
        topMargin=2.2 * rl["cm"],
        bottomMargin=1.4 * rl["cm"],
        title=f"Ficha Documental {ficha[1]}",
        author=EMPRESA_NOMBRE,
    )
    doc.detalle_generacion = construir_detalle_generacion(username)
    doc.config_empresa = obtener_config_empresa()

    (
        _,
        codigo,
        nombre,
        area,
        area_nombre,
        tipo,
        tipo_nombre,
        numero,
        version,
        estado,
        fecha,
        creado_por,
        *_extras,
    ) = ficha
    confidencial = bool(_extras[0]) if _extras else False

    datos_ficha = [
        ("Codigo", codigo),
        ("Nombre Documento", nombre),
        ("Area", f"{area} - {area_nombre}"),
        ("Tipo", f"{tipo} - {tipo_nombre}"),
        ("Numero", numero),
        ("Version", version),
        ("Estado", normalizar_estado(estado)),
        ("Confidencial", "Si" if confidencial else "No"),
        ("Fecha registro", fecha),
        ("Creado por", creado_por),
    ]

    elementos = [
        rl["Paragraph"](texto_pdf("Ficha Documental"), estilos["titulo"]),
        rl["Paragraph"](texto_pdf(codigo), estilos["subtitulo"]),
        rl["Spacer"](1, 0.35 * rl["cm"]),
    ]
    elementos.extend(tabla_resumen_pdf(rl, "Datos del documento", ("Campo", "Valor"), datos_ficha, estilos))
    elementos.extend(
        tabla_resumen_pdf(
            rl,
            "Historial documental",
            ("Fecha", "Campo", "Anterior", "Nuevo", "Observacion", "Usuario"),
            historial,
            estilos,
        )
    )

    doc.build(
        elementos,
        onFirstPage=lambda canvas, doc: encabezado_pie_pdf(canvas, doc, rl),
        onLaterPages=lambda canvas, doc: encabezado_pie_pdf(canvas, doc, rl),
    )
    return archivo


def abrir_pdf_para_imprimir(archivo):
    if not archivo:
        return False

    if not hasattr(os, "startfile"):
        raise RuntimeError("La apertura del PDF para impresion solo esta disponible en Windows.")

    os.startfile(archivo)
    return True


def exportar_dashboard_pdf(metricas=None, user_id=None, username="", imprimir=False):
    try:
        if metricas is None:
            metricas = obtener_metricas_dashboard()

        if imprimir:
            archivo = os.path.join(tempfile.gettempdir(), f"dashboard_documental_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        else:
            archivo = guardar_pdf("dashboard_documental.pdf")

        if not archivo:
            return

        rl = cargar_reportlab()
        estilos = construir_estilos_pdf(rl)
        doc = rl["SimpleDocTemplate"](
            archivo,
            pagesize=rl["A4"],
            rightMargin=1.4 * rl["cm"],
            leftMargin=1.4 * rl["cm"],
            topMargin=2.2 * rl["cm"],
            bottomMargin=1.4 * rl["cm"],
            title="Dashboard Documental",
            author=EMPRESA_NOMBRE,
        )
        doc.detalle_generacion = construir_detalle_generacion(username)
        doc.config_empresa = obtener_config_empresa()

        elementos = [
            rl["Paragraph"](texto_pdf("Dashboard Documental"), estilos["titulo"]),
            rl["Paragraph"](texto_pdf("Resumen ejecutivo del control documental"), estilos["subtitulo"]),
            rl["Spacer"](1, 0.35 * rl["cm"]),
        ]

        resumen = [
            ("Total documentos", metricas["total"]),
            ("Creados este mes", metricas["mes_actual"]),
            ("Estados activos", len(metricas["por_estado"])),
            ("Cambios recientes", len(metricas["ultimos_cambios"])),
        ]
        elementos.extend(tabla_resumen_pdf(rl, "Indicadores generales", ("Indicador", "Valor"), resumen, estilos))
        elementos.extend(tabla_resumen_pdf(rl, "Documentos por estado", ("Estado", "Cantidad"), metricas["por_estado"], estilos))
        elementos.extend(tabla_resumen_pdf(rl, "Documentos por area", ("Codigo", "Area", "Cantidad"), metricas["por_area"], estilos))
        elementos.extend(tabla_resumen_pdf(rl, "Documentos por tipo", ("Codigo", "Tipo", "Cantidad"), metricas["por_tipo"], estilos))
        elementos.extend(tabla_resumen_pdf(rl, "Ultimos documentos", ("Codigo", "Nombre", "Estado", "Fecha"), metricas["ultimos_documentos"], estilos))
        elementos.extend(tabla_resumen_pdf(rl, "Ultimos cambios", ("Fecha", "Codigo", "Campo", "Usuario"), metricas["ultimos_cambios"], estilos))

        doc.build(
            elementos,
            onFirstPage=lambda canvas, doc: encabezado_pie_pdf(canvas, doc, rl),
            onLaterPages=lambda canvas, doc: encabezado_pie_pdf(canvas, doc, rl),
        )

        accion = "impresion" if imprimir else "exportacion"
        if imprimir:
            abrir_pdf_para_imprimir(archivo)
            messagebox.showinfo(
                "Impresion",
                "El dashboard se abrio como PDF. Use la opcion Imprimir del visor para seleccionar impresora.",
            )
        else:
            messagebox.showinfo("Exportado", "Dashboard PDF generado correctamente")

        try:
            registrar_evento_reporte([], accion, "Dashboard documental", archivo, user_id)
        except Exception as auditoria_error:
            messagebox.showwarning(
                "Auditoria pendiente",
                "El PDF fue generado, pero no se pudo registrar la accion.\n"
                f"Verifique la migracion de auditoria.\n\n{auditoria_error}",
            )

    except Exception as e:
        messagebox.showerror("Error", str(e))


def exportar_pdf(user_id=None, username="", imprimir=False):
    try:
        datos = obtener_documentos(incluir_id=True)
        archivo_destino = None
        if imprimir:
            archivo_destino = os.path.join(tempfile.gettempdir(), f"reporte_documental_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

        archivo, documentos_ids = generar_pdf_documentos(
            datos,
            titulo="Reporte Documental",
            subtitulo="Listado general de documentos registrados",
            archivo=archivo_destino,
            username=username,
            mostrar_mensaje=not imprimir,
        )
        if archivo:
            accion = "impresion" if imprimir else "exportacion"
            if imprimir:
                abrir_pdf_para_imprimir(archivo)
                messagebox.showinfo(
                    "Impresion",
                    "El reporte se abrio como PDF. Use la opcion Imprimir del visor para seleccionar impresora.",
                )

            try:
                registrar_evento_reporte(documentos_ids, accion, "Reporte documental general", archivo, user_id)
            except Exception as auditoria_error:
                messagebox.showwarning(
                    "Auditoria pendiente",
                    "El PDF fue generado, pero no se pudo registrar la accion.\n"
                    f"Verifique la migracion de auditoria.\n\n{auditoria_error}",
                )
    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================
# AREAS Y TIPOS
# =========================
def obtener_areas():
    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT codigo, nombre
                FROM areas
                ORDER BY codigo
                """
            )
            datos = cur.fetchall()

    return [f"{x[0]} - {x[1]}" for x in datos]


def obtener_tipos():
    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT codigo, nombre
                FROM tipos_documento
                ORDER BY codigo
                """
            )
            datos = cur.fetchall()

    return [f"{x[0]} - {x[1]}" for x in datos]


def obtener_cargos(activos=True):
    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_configuracion_sistema(cur)
            query = "SELECT id, nombre FROM cargos"
            if activos:
                query += " WHERE activo = true"
            query += " ORDER BY nombre"
            cur.execute(query)
            datos = cur.fetchall()

    return [f"{x[0]} - {x[1]}" for x in datos]


def obtener_cargo_por_id(cargo_id):
    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_configuracion_sistema(cur)
            cur.execute(
                """
                SELECT id, nombre, COALESCE(rol_base, 'user') AS rol_base
                FROM cargos
                WHERE id = %s
                """,
                (cargo_id,),
            )
            return cur.fetchone()


def obtener_permisos_configurados_cargo(cargo_id):
    permisos, _ = obtener_permisos_cargo(cargo_id)
    return permisos


def extraer_id(valor_combo):
    try:
        return int(str(valor_combo).split(" - ")[0].strip())
    except (TypeError, ValueError):
        return None


def guardar_catalogo(tabla, codigo, nombre):
    codigo = codigo.strip().upper()
    nombre = nombre.strip()

    if not codigo or not nombre:
        raise ValueError("Debe completar codigo y nombre")

    if tabla not in ("areas", "tipos_documento"):
        raise ValueError("Catalogo no valido")

    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute(
                f"UPDATE {tabla} SET nombre = %s WHERE codigo = %s",
                (nombre, codigo),
            )
            if cur.rowcount == 0:
                cur.execute(
                    f"INSERT INTO {tabla} (codigo, nombre) VALUES (%s, %s)",
                    (codigo, nombre),
                )


def eliminar_catalogo(tabla, codigo):
    if tabla not in ("areas", "tipos_documento"):
        raise ValueError("Catalogo no valido")

    with conectar() as conn:
        with conn.cursor() as cur:
            cur.execute(f"DELETE FROM {tabla} WHERE codigo = %s", (codigo,))


def guardar_cargo(nombre, cargo_id=None, rol_base="user"):
    nombre = nombre.strip()
    rol_base = normalizar_rol(rol_base)
    if not nombre:
        raise ValueError("Debe ingresar el nombre del cargo")

    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_configuracion_sistema(cur)
            if cargo_id:
                cur.execute(
                    """
                    UPDATE cargos
                    SET nombre = %s,
                        rol_base = %s,
                        activo = true
                    WHERE id = %s
                    """,
                    (nombre, rol_base, cargo_id),
                )
            else:
                cur.execute(
                    """
                    INSERT INTO cargos (nombre, rol_base, activo)
                    VALUES (%s, %s, true)
                    ON CONFLICT (nombre) DO UPDATE
                    SET rol_base = EXCLUDED.rol_base,
                        activo = true
                    """,
                    (nombre, rol_base),
                )


def guardar_permisos_cargo(cargo_id, permisos):
    if not cargo_id:
        raise ValueError("Seleccione un cargo")

    permisos_validos = [permiso for permiso in permisos if permiso in PERMISOS_TODOS]

    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_configuracion_sistema(cur)
            cur.execute("DELETE FROM cargo_permisos WHERE cargo_id = %s", (cargo_id,))

            for permiso in permisos_validos:
                cur.execute(
                    """
                    INSERT INTO cargo_permisos (cargo_id, permiso, permitido, updated_at)
                    VALUES (%s, %s, true, now())
                    ON CONFLICT (cargo_id, permiso) DO UPDATE
                    SET permitido = true,
                        updated_at = now()
                    """,
                    (cargo_id, permiso),
                )


def desactivar_cargo(cargo_id):
    with conectar() as conn:
        with conn.cursor() as cur:
            asegurar_configuracion_sistema(cur)
            cur.execute(
                """
                UPDATE cargos
                SET activo = false
                WHERE id = %s
                """,
                (cargo_id,),
            )


# =========================
# CONSULTA DOCUMENTAL
# =========================
def abrir_consulta_documental(user_id, rol, username=""):
    rol = normalizar_rol(rol)

    if not usuario_tiene_permiso(user_id, rol, "consulta.ver"):
        messagebox.showerror("Permiso denegado", "No tiene permisos para consultar documentos")
        return

    consulta = tk.Toplevel()
    consulta.title("Consulta Documental")
    consulta.configure(bg="#f0f0f0")
    aplicar_icono(consulta)
    centrar_ventana(consulta, 1250, 720)

    filtros_frame = tk.Frame(consulta, bg="white", bd=1, relief="solid")
    filtros_frame.pack(fill="x", padx=15, pady=15)

    area_filtro = tk.StringVar()
    tipo_filtro = tk.StringVar()
    estado_filtro = tk.StringVar()
    texto_filtro = tk.StringVar()

    areas = ["TODAS"] + obtener_areas()
    tipos = ["TODOS"] + obtener_tipos()

    area_filtro.set(areas[0])
    tipo_filtro.set(tipos[0])
    estado_filtro.set("TODOS")

    tk.Label(filtros_frame, text="Buscar", bg="white", font=("Arial", 11)).grid(
        row=0, column=0, padx=(12, 6), pady=12, sticky="e"
    )

    busqueda_entry = tk.Entry(
        filtros_frame,
        textvariable=texto_filtro,
        width=32,
        font=("Arial", 11),
    )
    busqueda_entry.grid(row=0, column=1, padx=6, pady=12, sticky="w")

    tk.Label(filtros_frame, text="Area", bg="white", font=("Arial", 11)).grid(
        row=0, column=2, padx=(20, 6), pady=12, sticky="e"
    )

    area_menu = ttk.Combobox(
        filtros_frame,
        textvariable=area_filtro,
        values=areas,
        width=24,
        state="readonly",
    )
    area_menu.grid(row=0, column=3, padx=6, pady=12, sticky="w")

    tk.Label(filtros_frame, text="Tipo", bg="white", font=("Arial", 11)).grid(
        row=0, column=4, padx=(20, 6), pady=12, sticky="e"
    )

    tipo_menu = ttk.Combobox(
        filtros_frame,
        textvariable=tipo_filtro,
        values=tipos,
        width=24,
        state="readonly",
    )
    tipo_menu.grid(row=0, column=5, padx=6, pady=12, sticky="w")

    tk.Label(filtros_frame, text="Estado", bg="white", font=("Arial", 11)).grid(
        row=1, column=0, padx=(12, 6), pady=(0, 12), sticky="e"
    )

    estado_menu = ttk.Combobox(
        filtros_frame,
        textvariable=estado_filtro,
        values=["TODOS"] + list(ESTADOS_DOCUMENTO),
        width=24,
        state="readonly",
    )
    estado_menu.grid(row=1, column=1, padx=6, pady=(0, 12), sticky="w")

    acciones_frame = tk.Frame(consulta, bg="#f0f0f0")
    acciones_frame.pack(fill="x", padx=15, pady=(0, 10))

    tabla_frame = tk.Frame(consulta, bg="white", bd=1, relief="solid")
    tabla_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))

    columnas = ("id", "codigo", "nombre", "area", "tipo", "numero", "version", "estado", "fecha", "confidencial")

    tabla = ttk.Treeview(tabla_frame, columns=columnas, show="headings")
    tabla.heading("id", text="ID")
    tabla.heading("codigo", text="Codigo")
    tabla.heading("nombre", text="Nombre Documento")
    tabla.heading("area", text="Area")
    tabla.heading("tipo", text="Tipo")
    tabla.heading("numero", text="Numero")
    tabla.heading("version", text="Version")
    tabla.heading("estado", text="Estado")
    tabla.heading("fecha", text="Fecha")
    tabla.heading("confidencial", text="Conf.")

    tabla.column("id", width=0, minwidth=0, stretch=False)
    tabla.column("codigo", width=210, anchor="w")
    tabla.column("nombre", width=390, anchor="w")
    tabla.column("area", width=80, anchor="center")
    tabla.column("tipo", width=80, anchor="center")
    tabla.column("numero", width=80, anchor="center")
    tabla.column("version", width=80, anchor="center")
    tabla.column("estado", width=110, anchor="center")
    tabla.column("fecha", width=160, anchor="center")
    tabla.column("confidencial", width=70, anchor="center")

    scrollbar_y = ttk.Scrollbar(tabla_frame, orient="vertical", command=tabla.yview)
    scrollbar_x = ttk.Scrollbar(tabla_frame, orient="horizontal", command=tabla.xview)
    tabla.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    tabla.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=1, sticky="ns")
    scrollbar_x.grid(row=1, column=0, sticky="ew")

    tabla_frame.grid_rowconfigure(0, weight=1)
    tabla_frame.grid_columnconfigure(0, weight=1)

    def cargar_documentos():
        for item in tabla.get_children():
            tabla.delete(item)

        area = None if area_filtro.get() == "TODAS" else extraer_codigo(area_filtro.get())
        tipo = None if tipo_filtro.get() == "TODOS" else extraer_codigo(tipo_filtro.get())
        estado = None if estado_filtro.get() == "TODOS" else estado_filtro.get()
        texto = texto_filtro.get().strip() or None

        try:
            documentos = obtener_documentos(area=area, tipo=tipo, estado=estado, texto=texto, incluir_id=True)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        for doc in documentos:
            fila = list(doc)
            fila[-1] = "Si" if fila[-1] else "No"
            tabla.insert("", tk.END, values=tuple(fila))

    def obtener_documento_seleccionado():
        seleccionado = tabla.selection()

        if not seleccionado:
            messagebox.showwarning("Seleccion requerida", "Seleccione un documento")
            return None

        valores = tabla.item(seleccionado[0], "values")

        if not valores:
            return None

        return valores

    def abrir_edicion_documento():
        if not puede_editar_documentos(rol, user_id):
            messagebox.showerror("Permiso denegado", "No tiene permisos para editar documentos")
            return

        seleccionado = obtener_documento_seleccionado()

        if not seleccionado:
            return

        documento_id = seleccionado[0]

        try:
            documento = obtener_documento_por_id(documento_id)
            areas_edicion = obtener_areas()
            tipos_edicion = obtener_tipos()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        if not documento:
            messagebox.showerror("Error", "Documento no encontrado")
            return

        _, codigo, nombre, area, tipo, numero, version, estado, fecha, confidencial = documento

        win = tk.Toplevel(consulta)
        win.title("Editar Documento")
        aplicar_icono(win)
        centrar_ventana(win, 600, 620)
        win.configure(bg="white")

        codigo_var = tk.StringVar(value=codigo)
        nombre_var = tk.StringVar(value=nombre)
        area_var = tk.StringVar(value=next((x for x in areas_edicion if x.startswith(f"{area} - ")), area))
        tipo_var = tk.StringVar(value=next((x for x in tipos_edicion if x.startswith(f"{tipo} - ")), tipo))
        numero_var = tk.StringVar(value=numero)
        version_var = tk.StringVar(value=version)
        estado_var = tk.StringVar(value=normalizar_estado(estado))
        observacion_var = tk.StringVar()
        confidencial_var = tk.BooleanVar(value=bool(confidencial))
        estados_disponibles = estados_edicion_permitidos(rol, estado, user_id)

        puede_editar_codigo = puede_editar_codigo_documental(rol, user_id)
        puede_editar_metadatos = puede_editar_metadatos_documento(rol, normalizar_estado(estado), user_id)
        estado_codigo = "normal" if puede_editar_codigo else "readonly"
        estado_combo = "readonly" if puede_editar_codigo else "disabled"
        estado_num_version = "normal" if puede_editar_codigo else "readonly"
        estado_nombre = "normal" if usuario_tiene_permiso(user_id, rol, "documentos.nombre.editar") and puede_editar_metadatos else "readonly"
        estado_confidencial = "normal" if (
            puede_editar_codigo
            or (usuario_tiene_permiso(user_id, rol, "documentos.confidencial.editar") and puede_editar_metadatos)
        ) else "disabled"

        form = tk.Frame(win, bg="white")
        form.pack(fill="both", expand=True, padx=24, pady=20)
        form.grid_columnconfigure(1, weight=1)

        tk.Label(form, text="Codigo", bg="white", font=("Arial", 11)).grid(row=0, column=0, padx=(0, 12), pady=8, sticky="e")
        tk.Entry(form, textvariable=codigo_var, state=estado_codigo, font=("Arial", 11)).grid(row=0, column=1, pady=8, sticky="ew")

        tk.Label(form, text="Nombre Documento", bg="white", font=("Arial", 11)).grid(row=1, column=0, padx=(0, 12), pady=8, sticky="e")
        tk.Entry(form, textvariable=nombre_var, state=estado_nombre, font=("Arial", 11)).grid(row=1, column=1, pady=8, sticky="ew")

        tk.Label(form, text="Area", bg="white", font=("Arial", 11)).grid(row=2, column=0, padx=(0, 12), pady=8, sticky="e")
        ttk.Combobox(form, textvariable=area_var, values=areas_edicion, state=estado_combo, font=("Arial", 10)).grid(row=2, column=1, pady=8, sticky="ew")

        tk.Label(form, text="Tipo", bg="white", font=("Arial", 11)).grid(row=3, column=0, padx=(0, 12), pady=8, sticky="e")
        ttk.Combobox(form, textvariable=tipo_var, values=tipos_edicion, state=estado_combo, font=("Arial", 10)).grid(row=3, column=1, pady=8, sticky="ew")

        tk.Label(form, text="Numero", bg="white", font=("Arial", 11)).grid(row=4, column=0, padx=(0, 12), pady=8, sticky="e")
        tk.Entry(form, textvariable=numero_var, state=estado_num_version, font=("Arial", 11)).grid(row=4, column=1, pady=8, sticky="ew")

        tk.Label(form, text="Version", bg="white", font=("Arial", 11)).grid(row=5, column=0, padx=(0, 12), pady=8, sticky="e")
        tk.Entry(form, textvariable=version_var, state=estado_num_version, font=("Arial", 11)).grid(row=5, column=1, pady=8, sticky="ew")

        tk.Label(form, text="Estado", bg="white", font=("Arial", 11)).grid(row=6, column=0, padx=(0, 12), pady=8, sticky="e")
        ttk.Combobox(form, textvariable=estado_var, values=estados_disponibles, state="readonly", font=("Arial", 10)).grid(row=6, column=1, pady=8, sticky="ew")

        tk.Label(form, text="Observacion estado", bg="white", font=("Arial", 11)).grid(row=7, column=0, padx=(0, 12), pady=8, sticky="e")
        tk.Entry(form, textvariable=observacion_var, font=("Arial", 11)).grid(row=7, column=1, pady=8, sticky="ew")

        tk.Checkbutton(
            form,
            text="Documento confidencial",
            variable=confidencial_var,
            state=estado_confidencial,
            bg="white",
            font=("Arial", 10, "bold"),
        ).grid(row=8, column=0, columnspan=2, pady=(8, 4))

        tk.Label(form, text=f"Fecha registro: {fecha}", bg="white", fg="gray", font=("Arial", 10)).grid(row=9, column=0, columnspan=2, pady=(8, 16))

        if not puede_editar_codigo:
            tk.Label(
                form,
                text="Codigo, area, tipo, numero y version bloqueados. Solo Superusuario puede modificarlos.",
                bg="white",
                fg="#8a5a00",
                font=("Arial", 9),
                wraplength=470,
            ).grid(row=10, column=0, columnspan=2, pady=(0, 12))

        if not puede_editar_metadatos and not puede_editar_codigo:
            tk.Label(
                form,
                text=f"Metadatos bloqueados para documentos en estado {normalizar_estado(estado)}.",
                bg="white",
                fg="#8a5a00",
                font=("Arial", 9),
                wraplength=510,
            ).grid(row=11, column=0, columnspan=2, pady=(0, 12))

        tk.Label(
            form,
            text=f"Estados permitidos desde {normalizar_estado(estado)}: {', '.join(estados_disponibles)}.",
            bg="white",
            fg="#2f3b4a",
            font=("Arial", 9),
            wraplength=510,
        ).grid(row=12, column=0, columnspan=2, pady=(0, 12))

        def guardar_cambios():
            datos_nuevos = {
                "codigo": codigo_var.get(),
                "nombre": nombre_var.get(),
                "area": extraer_codigo(area_var.get()),
                "tipo": extraer_codigo(tipo_var.get()),
                "numero": numero_var.get(),
                "version": version_var.get(),
                "estado": estado_var.get(),
                "observacion": observacion_var.get(),
                "confidencial": confidencial_var.get(),
            }

            if normalizar_estado(estado) != "Vigente" and normalizar_estado(datos_nuevos["estado"]) == "Vigente":
                try:
                    evaluacion = evaluar_aprobacion_version(
                        datos_nuevos["area"],
                        datos_nuevos["tipo"],
                        datos_nuevos["numero"],
                        datos_nuevos["version"],
                        documento_id,
                    )
                except Exception as e:
                    messagebox.showerror("Error", str(e))
                    return

                if not evaluacion["ok"]:
                    messagebox.showerror("Version no valida", evaluacion["mensaje"])
                    return

                if evaluacion["mensaje"] and not messagebox.askyesno(
                    "Confirmar aprobacion de version",
                    evaluacion["mensaje"] + "\n\nDesea continuar?",
                    parent=win,
                ):
                    return

            try:
                hubo_cambios = actualizar_documento(documento_id, datos_nuevos, user_id, rol)
            except Exception as e:
                messagebox.showerror("Error", str(e))
                return

            if hubo_cambios:
                messagebox.showinfo("OK", "Documento actualizado")
                cargar_documentos()
                win.destroy()
            else:
                messagebox.showinfo("Sin cambios", "No hay cambios para guardar")

        botones = tk.Frame(form, bg="white")
        botones.grid(row=13, column=0, columnspan=2, pady=8)

        tk.Button(botones, text="Guardar Cambios", command=guardar_cambios, width=18).grid(row=0, column=0, padx=8)
        tk.Button(botones, text="Cerrar", command=win.destroy, width=14).grid(row=0, column=1, padx=8)

    def abrir_ficha_documental():
        seleccionado = obtener_documento_seleccionado()

        if not seleccionado:
            return

        documento_id = seleccionado[0]

        try:
            ficha = obtener_ficha_documento(documento_id)
            historial = obtener_historial_documento(documento_id)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        if not ficha:
            messagebox.showerror("Error", "Documento no encontrado")
            return

        (
            _,
            codigo,
            nombre,
            area,
            area_nombre,
            tipo,
            tipo_nombre,
            numero,
            version,
            estado,
            fecha,
            creado_por,
            confidencial,
        ) = ficha

        estado = normalizar_estado(estado)
        color_estado = color_estado_documental(estado)
        resumen_qr = construir_resumen_validacion(ficha, historial)

        if confidencial and not usuario_tiene_permiso(user_id, rol, "documentos.confidencial.ver"):
            messagebox.showerror("Permiso denegado", "No tiene permisos para ver documentos confidenciales")
            return

        if confidencial:
            try:
                registrar_evento_reporte(
                    [(documento_id, codigo)],
                    "visualizacion",
                    "Ficha documental confidencial",
                    "Vista en pantalla",
                    user_id,
                )
            except Exception as auditoria_error:
                messagebox.showwarning(
                    "Auditoria pendiente",
                    "La ficha confidencial se abrio, pero no se pudo registrar la visualizacion.\n"
                    f"Verifique la migracion de auditoria.\n\n{auditoria_error}",
                )

        win = tk.Toplevel(consulta)
        win.title("Ficha Documental")
        aplicar_icono(win)
        centrar_ventana(win, 1120, 640)
        win.configure(bg="#f0f0f0")
        win.minsize(760, 500)

        scroll_area = tk.Frame(win, bg="#f0f0f0")
        scroll_area.pack(side="top", fill="both", expand=True)
        canvas = tk.Canvas(scroll_area, bg="#f0f0f0", highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_area, orient="vertical", command=canvas.yview)
        contenido = tk.Frame(canvas, bg="#f0f0f0")
        contenido_id = canvas.create_window((0, 0), window=contenido, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def ajustar_scroll_ficha(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(contenido_id, width=canvas.winfo_width())

        contenido.bind("<Configure>", ajustar_scroll_ficha)
        canvas.bind("<Configure>", ajustar_scroll_ficha)
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

        def cerrar_ficha():
            canvas.unbind_all("<MouseWheel>")
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", cerrar_ficha)

        header = tk.Frame(contenido, bg="white", bd=1, relief="solid")
        header.pack(fill="x", padx=15, pady=15)
        header.grid_columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=codigo,
            bg="white",
            fg="#2d5bd1",
            font=("Arial", 18, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 2))

        tk.Label(
            header,
            text=nombre or "Sin nombre documental",
            bg="white",
            fg="#222222",
            font=("Arial", 13),
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 12))

        tk.Label(
            header,
            text=estado.upper(),
            bg=color_estado,
            fg="white",
            font=("Arial", 12, "bold"),
            width=24,
        ).grid(row=0, column=1, padx=16, pady=(14, 3), sticky="e")

        if confidencial:
            tk.Label(
                header,
                text="CONFIDENCIAL",
                bg="#8e1b1b",
                fg="white",
                font=("Arial", 11, "bold"),
                width=24,
            ).grid(row=1, column=1, padx=16, pady=(3, 14), sticky="e")

        cuerpo = tk.Frame(contenido, bg="#f0f0f0")
        cuerpo.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        cuerpo.grid_columnconfigure(0, weight=1)
        cuerpo.grid_columnconfigure(1, weight=1)
        cuerpo.grid_rowconfigure(1, weight=1)

        datos_panel = tk.Frame(cuerpo, bg="white", bd=1, relief="solid")
        datos_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 12))
        etapas_panel = tk.Frame(cuerpo, bg="white", bd=1, relief="solid")
        etapas_panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 12))

        tk.Label(datos_panel, text="Datos generales", bg="white", fg="#2d5bd1", font=("Arial", 12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 8)
        )

        datos = [
            ("Area", f"{area} - {area_nombre}"),
            ("Tipo", f"{tipo} - {tipo_nombre}"),
            ("Numero", numero),
            ("Version", version),
            ("Confidencial", "Si" if confidencial else "No"),
            ("Creado por", creado_por),
            ("Fecha creacion", formatear_fecha(fecha)),
        ]

        for idx, (label, valor) in enumerate(datos, start=1):
            tk.Label(datos_panel, text=label, bg="white", fg="gray", font=("Arial", 10)).grid(
                row=idx, column=0, sticky="e", padx=(12, 8), pady=4
            )
            tk.Label(datos_panel, text=valor, bg="white", fg="#222222", font=("Arial", 10)).grid(
                row=idx, column=1, sticky="w", padx=(0, 12), pady=4
            )

        tk.Label(etapas_panel, text="Etapas documentales", bg="white", fg="#2d5bd1", font=("Arial", 12, "bold")).pack(
            anchor="w", padx=12, pady=(10, 8)
        )

        etapas = [
            ("Creacion", "Registrado", creado_por, formatear_fecha(fecha), "#1f6fbf"),
        ]

        for item in reversed(historial):
            fecha_h, campo, anterior, nuevo, observacion, usuario = item
            if campo == "estado":
                etapas.append((nuevo, f"Desde: {anterior}", usuario, formatear_fecha(fecha_h), color_estado_documental(nuevo)))

        for titulo, subtitulo, usuario, fecha_etapa, color in etapas:
            fila = tk.Frame(etapas_panel, bg="white")
            fila.pack(fill="x", padx=12, pady=4)

            tk.Label(fila, bg=color, width=2, height=2).pack(side="left", padx=(0, 8))
            texto = tk.Frame(fila, bg="white")
            texto.pack(side="left", fill="x", expand=True)
            tk.Label(texto, text=titulo, bg="white", fg="#222222", font=("Arial", 10, "bold")).pack(anchor="w")
            tk.Label(texto, text=f"{subtitulo} | {usuario} | {fecha_etapa}", bg="white", fg="gray", font=("Arial", 9)).pack(anchor="w")

        historial_panel = tk.Frame(cuerpo, bg="white", bd=1, relief="solid")
        historial_panel.grid(row=1, column=0, columnspan=2, sticky="nsew")
        historial_panel.grid_columnconfigure(0, weight=1)
        historial_panel.grid_rowconfigure(1, weight=1)

        tk.Label(historial_panel, text="Historial completo", bg="white", fg="#2d5bd1", font=("Arial", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 8)
        )

        columnas_ficha = ("fecha", "campo", "anterior", "nuevo", "observacion", "usuario")
        tabla_ficha = ttk.Treeview(historial_panel, columns=columnas_ficha, show="headings", height=8)

        encabezados = {
            "fecha": "Fecha",
            "campo": "Campo",
            "anterior": "Valor Anterior",
            "nuevo": "Valor Nuevo",
            "observacion": "Observacion",
            "usuario": "Usuario",
        }

        anchos = {
            "fecha": 150,
            "campo": 120,
            "anterior": 190,
            "nuevo": 190,
            "observacion": 280,
            "usuario": 140,
        }

        for col in columnas_ficha:
            tabla_ficha.heading(col, text=encabezados[col])
            tabla_ficha.column(col, width=anchos[col], anchor="w")

        scroll_hist = ttk.Scrollbar(historial_panel, orient="vertical", command=tabla_ficha.yview)
        tabla_ficha.configure(yscrollcommand=scroll_hist.set)
        tabla_ficha.grid(row=1, column=0, sticky="nsew", padx=(12, 0), pady=(0, 12))
        scroll_hist.grid(row=1, column=1, sticky="ns", padx=(0, 12), pady=(0, 12))

        for item in historial:
            tabla_ficha.insert("", tk.END, values=formatear_fila(item))

        if not historial:
            tabla_ficha.insert("", tk.END, values=("", "Sin historial", "", "", "", ""))

        def exportar_ficha_pdf():
            try:
                archivo = generar_pdf_ficha_documental(ficha, historial, archivo=None, username=username)
                if not archivo:
                    return

                messagebox.showinfo("Exportado", "Ficha documental PDF generada correctamente")
                try:
                    registrar_evento_reporte(
                        [(documento_id, codigo)],
                        "exportacion",
                        "Ficha documental",
                        archivo,
                        user_id,
                    )
                except Exception as auditoria_error:
                    messagebox.showwarning(
                        "Auditoria pendiente",
                        "La ficha fue generada, pero no se pudo registrar la exportacion.\n"
                        f"Verifique la migracion de auditoria.\n\n{auditoria_error}",
                    )
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def imprimir_ficha_pdf():
            try:
                archivo = os.path.join(tempfile.gettempdir(), f"ficha_{codigo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
                archivo = generar_pdf_ficha_documental(ficha, historial, archivo=archivo, username=username)
                if not archivo:
                    return

                abrir_pdf_para_imprimir(archivo)
                messagebox.showinfo(
                    "Impresion",
                    "La ficha se abrio como PDF. Use la opcion Imprimir del visor para seleccionar impresora.",
                )
                try:
                    registrar_evento_reporte(
                        [(documento_id, codigo)],
                        "impresion",
                        "Ficha documental",
                        archivo,
                        user_id,
                    )
                except Exception as auditoria_error:
                    messagebox.showwarning(
                        "Auditoria pendiente",
                        "La ficha fue enviada a imprimir, pero no se pudo registrar la impresion.\n"
                        f"Verifique la migracion de auditoria.\n\n{auditoria_error}",
                    )
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def abrir_registro_reportes_documento():
            try:
                eventos = obtener_eventos_reporte_documento(documento_id)
            except Exception as e:
                messagebox.showerror("Error", str(e))
                return

            reg = tk.Toplevel(win)
            reg.title("Registro de Exportaciones e Impresiones")
            aplicar_icono(reg)
            centrar_ventana(reg, 1120, 430)

            tk.Label(
                reg,
                text=f"{codigo} | {nombre}",
                font=("Arial", 12, "bold"),
            ).pack(pady=(12, 8))

            columnas_reg = ("fecha", "accion", "reporte", "usuario_app", "usuario_sistema", "equipo", "archivo")
            tabla_reg = ttk.Treeview(reg, columns=columnas_reg, show="headings")
            encabezados_reg = {
                "fecha": "Fecha",
                "accion": "Accion",
                "reporte": "Reporte",
                "usuario_app": "Usuario App",
                "usuario_sistema": "Usuario Windows",
                "equipo": "Equipo",
                "archivo": "Archivo",
            }
            anchos_reg = {
                "fecha": 150,
                "accion": 100,
                "reporte": 170,
                "usuario_app": 130,
                "usuario_sistema": 150,
                "equipo": 150,
                "archivo": 320,
            }

            for col in columnas_reg:
                tabla_reg.heading(col, text=encabezados_reg[col])
                tabla_reg.column(col, width=anchos_reg[col], anchor="w")

            scroll_reg = ttk.Scrollbar(reg, orient="vertical", command=tabla_reg.yview)
            tabla_reg.configure(yscrollcommand=scroll_reg.set)
            tabla_reg.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=(0, 12))
            scroll_reg.pack(side="right", fill="y", padx=(0, 12), pady=(0, 12))

            for item in eventos:
                tabla_reg.insert("", tk.END, values=formatear_fila(item))

            if not eventos:
                tabla_reg.insert("", tk.END, values=("", "Sin registros", "", "", "", "", ""))

        acciones = tk.Frame(win, bg="#f0f0f0", bd=1, relief="solid")
        acciones.pack(side="bottom", fill="x")
        botones = tk.Frame(acciones, bg="#f0f0f0")
        botones.pack(fill="x", padx=15, pady=10)

        if puede_editar_documentos(rol, user_id):
            tk.Button(botones, text="Editar Documento", command=lambda: [cerrar_ficha(), abrir_edicion_documento()], width=18).pack(side="left")

        if usuario_tiene_permiso(user_id, rol, "documentos.qr.generar"):
            tk.Button(
                botones,
                text="Generar QR",
                command=lambda: abrir_validacion_qr(resumen_qr, win),
                width=16,
            ).pack(side="left", padx=(10, 0))

        if usuario_tiene_permiso(user_id, rol, "reportes.ficha.exportar"):
            tk.Button(
                botones,
                text="Exportar Ficha",
                command=exportar_ficha_pdf,
                width=16,
            ).pack(side="left", padx=(10, 0))

        if usuario_tiene_permiso(user_id, rol, "reportes.ficha.imprimir"):
            tk.Button(
                botones,
                text="Imprimir Ficha",
                command=imprimir_ficha_pdf,
                width=16,
            ).pack(side="left", padx=(10, 0))

        if usuario_tiene_permiso(user_id, rol, "reportes.auditoria.ver"):
            tk.Button(
                botones,
                text="Registro PDF/Imp.",
                command=abrir_registro_reportes_documento,
                width=18,
            ).pack(side="left", padx=(10, 0))

        tk.Button(botones, text="Cerrar", command=cerrar_ficha, width=14).pack(side="right")

    def abrir_historial_documento():
        if not puede_ver_historial(rol, user_id):
            messagebox.showerror("Permiso denegado", "No tiene permisos para ver historial")
            return

        seleccionado = obtener_documento_seleccionado()

        if not seleccionado:
            return

        documento_id = seleccionado[0]
        codigo = seleccionado[1]
        nombre = seleccionado[2]

        try:
            historial = obtener_historial_documento(documento_id)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        win = tk.Toplevel(consulta)
        win.title("Historial Documento")
        aplicar_icono(win)
        centrar_ventana(win, 1050, 460)

        tk.Label(
            win,
            text=f"{codigo} | {nombre}",
            font=("Arial", 12, "bold"),
        ).pack(pady=(12, 8))

        columnas_hist = ("fecha", "campo", "anterior", "nuevo", "observacion", "usuario")
        tabla_hist = ttk.Treeview(win, columns=columnas_hist, show="headings")
        tabla_hist.heading("fecha", text="Fecha")
        tabla_hist.heading("campo", text="Campo")
        tabla_hist.heading("anterior", text="Valor Anterior")
        tabla_hist.heading("nuevo", text="Valor Nuevo")
        tabla_hist.heading("observacion", text="Observacion")
        tabla_hist.heading("usuario", text="Usuario")

        tabla_hist.column("fecha", width=150)
        tabla_hist.column("campo", width=120)
        tabla_hist.column("anterior", width=190)
        tabla_hist.column("nuevo", width=190)
        tabla_hist.column("observacion", width=250)
        tabla_hist.column("usuario", width=140)
        tabla_hist.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        for item in historial:
            tabla_hist.insert("", tk.END, values=formatear_fila(item))

        if not historial:
            tabla_hist.insert("", tk.END, values=("", "Sin historial", "", "", "", ""))

    def mostrar_todo():
        area_filtro.set("TODAS")
        tipo_filtro.set("TODOS")
        estado_filtro.set("TODOS")
        texto_filtro.set("")
        cargar_documentos()

    def exportar_pdf_consulta():
        generar_pdf_consulta(imprimir=False)

    def imprimir_pdf_consulta():
        generar_pdf_consulta(imprimir=True)

    def generar_pdf_consulta(imprimir=False):
        area = None if area_filtro.get() == "TODAS" else extraer_codigo(area_filtro.get())
        tipo = None if tipo_filtro.get() == "TODOS" else extraer_codigo(tipo_filtro.get())
        estado = None if estado_filtro.get() == "TODOS" else estado_filtro.get()
        texto = texto_filtro.get().strip() or None

        filtros = []
        if area:
            filtros.append(f"Area: {area}")
        if tipo:
            filtros.append(f"Tipo: {tipo}")
        if estado:
            filtros.append(f"Estado: {estado}")
        if texto:
            filtros.append(f"Busqueda: {texto}")

        subtitulo = "Consulta documental"
        if filtros:
            subtitulo += " | " + " | ".join(filtros)

        try:
            datos = obtener_documentos(area=area, tipo=tipo, estado=estado, texto=texto, incluir_id=True)
            archivo_destino = None
            if imprimir:
                archivo_destino = os.path.join(
                    tempfile.gettempdir(),
                    f"consulta_documental_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                )

            archivo, documentos_ids = generar_pdf_documentos(
                datos,
                titulo="Reporte de Consulta Documental",
                subtitulo=subtitulo,
                archivo=archivo_destino,
                username=username,
                mostrar_mensaje=not imprimir,
            )
            if archivo:
                accion = "impresion" if imprimir else "exportacion"
                if imprimir:
                    abrir_pdf_para_imprimir(archivo)
                    messagebox.showinfo(
                        "Impresion",
                        "El reporte se abrio como PDF. Use la opcion Imprimir del visor para seleccionar impresora.",
                    )

                try:
                    registrar_evento_reporte(documentos_ids, accion, "Reporte de consulta documental", archivo, user_id)
                except Exception as auditoria_error:
                    messagebox.showwarning(
                        "Auditoria pendiente",
                        "El PDF fue generado, pero no se pudo registrar la accion.\n"
                        f"Verifique la migracion de auditoria.\n\n{auditoria_error}",
                    )
        except Exception as e:
            messagebox.showerror("Error", str(e))

    tk.Button(
        filtros_frame,
        text="Filtrar",
        width=14,
        command=cargar_documentos,
    ).grid(row=0, column=6, padx=(24, 6), pady=12)

    tk.Button(
        filtros_frame,
        text="Mostrar Todo",
        width=14,
        command=mostrar_todo,
    ).grid(row=0, column=7, padx=(6, 12), pady=12)

    if puede_editar_documentos(rol, user_id):
        tk.Button(
            acciones_frame,
            text="Editar Documento",
            width=18,
            command=abrir_edicion_documento,
        ).pack(side="left", padx=(0, 10))

    if puede_ver_historial(rol, user_id):
        tk.Button(
            acciones_frame,
            text="Ver Historial",
            width=18,
            command=abrir_historial_documento,
        ).pack(side="left")

    tk.Button(
        acciones_frame,
        text="Ver Ficha",
        width=18,
        command=abrir_ficha_documental,
    ).pack(side="left", padx=(10, 0))

    if usuario_tiene_permiso(user_id, rol, "reportes.consulta.exportar"):
        tk.Button(
            acciones_frame,
            text="Exportar PDF",
            width=18,
            command=exportar_pdf_consulta,
        ).pack(side="right", padx=(10, 0))

    if usuario_tiene_permiso(user_id, rol, "reportes.consulta.imprimir"):
        tk.Button(
            acciones_frame,
            text="Imprimir PDF",
            width=18,
            command=imprimir_pdf_consulta,
        ).pack(side="right")

    busqueda_entry.bind("<Return>", lambda event: cargar_documentos())
    tabla.bind("<Double-1>", lambda event: abrir_ficha_documental())
    cargar_documentos()


# =========================
# DASHBOARD
# =========================
def abrir_dashboard(user_id=None, username="", rol="lector"):
    rol = normalizar_rol(rol)

    if not usuario_tiene_permiso(user_id, rol, "dashboard.ver"):
        messagebox.showerror("Permiso denegado", "No tiene permisos para ver el dashboard")
        return

    dashboard = tk.Toplevel()
    dashboard.title("Dashboard Documental")
    dashboard.configure(bg="#f0f0f0")
    aplicar_icono(dashboard)
    centrar_ventana(dashboard, 1180, 720)

    canvas = tk.Canvas(dashboard, bg="#f0f0f0", highlightthickness=0)
    scroll_y = ttk.Scrollbar(dashboard, orient="vertical", command=canvas.yview)
    dashboard_body = tk.Frame(canvas, bg="#f0f0f0")

    dashboard_body_id = canvas.create_window((0, 0), window=dashboard_body, anchor="nw")
    canvas.configure(yscrollcommand=scroll_y.set)

    canvas.pack(side="left", fill="both", expand=True)
    scroll_y.pack(side="right", fill="y")

    def ajustar_scroll(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def ajustar_ancho(event):
        canvas.itemconfigure(dashboard_body_id, width=event.width)

    def rueda_mouse(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    dashboard_body.bind("<Configure>", ajustar_scroll)
    canvas.bind("<Configure>", ajustar_ancho)
    canvas.bind_all("<MouseWheel>", rueda_mouse)

    def limpiar_rueda_mouse(event):
        if event.widget == dashboard:
            canvas.unbind_all("<MouseWheel>")

    dashboard.bind("<Destroy>", limpiar_rueda_mouse)

    header = tk.Frame(dashboard_body, bg="white", bd=1, relief="solid")
    header.pack(fill="x", padx=15, pady=15)

    tk.Label(
        header,
        text="DASHBOARD DOCUMENTAL",
        font=("Arial", 18, "bold"),
        fg="#2d5bd1",
        bg="white",
    ).pack(side="left", padx=16, pady=14)

    contenido = tk.Frame(dashboard_body, bg="#f0f0f0")
    contenido.pack(fill="both", expand=True, padx=15, pady=(0, 15))
    contenido.grid_columnconfigure(0, weight=1)
    contenido.grid_columnconfigure(1, weight=1)
    contenido.grid_rowconfigure(1, weight=1, minsize=190)
    contenido.grid_rowconfigure(2, weight=2, minsize=230)
    contenido.grid_rowconfigure(3, weight=2, minsize=210)

    try:
        metricas = obtener_metricas_dashboard()
    except Exception as e:
        messagebox.showerror("Error", str(e))
        dashboard.destroy()
        return

    if usuario_tiene_permiso(user_id, rol, "dashboard.exportar"):
        tk.Button(
            header,
            text="Exportar PDF",
            font=("Arial", 10),
            width=16,
            command=lambda: exportar_dashboard_pdf(metricas, user_id, username),
        ).pack(side="right", padx=16, pady=14)

    if usuario_tiene_permiso(user_id, rol, "dashboard.imprimir"):
        tk.Button(
            header,
            text="Imprimir PDF",
            font=("Arial", 10),
            width=16,
            command=lambda: exportar_dashboard_pdf(metricas, user_id, username, imprimir=True),
        ).pack(side="right", padx=(0, 8), pady=14)

    resumen = tk.Frame(contenido, bg="#f0f0f0")
    resumen.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))

    def tarjeta(parent, titulo, valor, columna):
        frame = tk.Frame(parent, bg="white", bd=1, relief="solid", width=250, height=88)
        frame.grid(row=0, column=columna, padx=(0, 12), sticky="ew")
        frame.grid_propagate(False)
        parent.grid_columnconfigure(columna, weight=1)

        tk.Label(frame, text=titulo, bg="white", fg="gray", font=("Arial", 10)).pack(anchor="w", padx=14, pady=(12, 2))
        tk.Label(frame, text=str(valor), bg="white", fg="#222222", font=("Arial", 24, "bold")).pack(anchor="w", padx=14)

    tarjeta(resumen, "Total documentos", metricas["total"], 0)
    tarjeta(resumen, "Creados este mes", metricas["mes_actual"], 1)
    tarjeta(resumen, "Estados activos", len(metricas["por_estado"]), 2)
    tarjeta(resumen, "Cambios recientes", len(metricas["ultimos_cambios"]), 3)

    def panel(parent, titulo, row, column, columnspan=1):
        frame = tk.Frame(parent, bg="white", bd=1, relief="solid")
        frame.grid(
            row=row,
            column=column,
            columnspan=columnspan,
            sticky="nsew",
            padx=(0, 8) if column == 0 and columnspan == 1 else (8, 0) if column == 1 else 0,
            pady=(0, 12),
        )
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        tk.Label(
            frame,
            text=titulo,
            bg="white",
            fg="#2d5bd1",
            font=("Arial", 12, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))

        return frame

    panel_estado = panel(contenido, "Documentos por estado", 1, 0)
    panel_area = panel(contenido, "Documentos por area", 1, 1)
    panel_tipo = panel(contenido, "Documentos por tipo", 2, 0)
    panel_recientes = panel(contenido, "Ultimos documentos", 2, 1)
    panel_cambios = panel(contenido, "Ultimos cambios", 3, 0, columnspan=2)

    def tabla_simple(parent, columnas, anchos, filas):
        contenedor_tabla = tk.Frame(parent, bg="white")
        contenedor_tabla.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        contenedor_tabla.grid_rowconfigure(0, weight=1)
        contenedor_tabla.grid_columnconfigure(0, weight=1)

        tabla = ttk.Treeview(contenedor_tabla, columns=columnas, show="headings", height=7)

        for col, ancho in zip(columnas, anchos):
            tabla.heading(col, text=col)
            tabla.column(col, width=ancho, anchor="w")

        scrollbar = ttk.Scrollbar(contenedor_tabla, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=scrollbar.set)

        tabla.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        for fila in filas:
            tabla.insert("", tk.END, values=formatear_fila(fila))

        if not filas:
            tabla.insert("", tk.END, values=tuple("" for _ in columnas))

        return tabla

    tabla_simple(
        panel_estado,
        ("Estado", "Cantidad"),
        (220, 90),
        metricas["por_estado"],
    )

    tabla_simple(
        panel_area,
        ("Codigo", "Area", "Cantidad"),
        (80, 260, 90),
        metricas["por_area"],
    )

    tabla_simple(
        panel_tipo,
        ("Codigo", "Tipo", "Cantidad"),
        (80, 260, 90),
        metricas["por_tipo"],
    )

    tabla_simple(
        panel_recientes,
        ("Codigo", "Nombre", "Estado", "Fecha"),
        (160, 300, 110, 160),
        metricas["ultimos_documentos"],
    )

    tabla_simple(
        panel_cambios,
        ("Fecha", "Codigo", "Campo", "Usuario"),
        (150, 220, 160, 180),
        metricas["ultimos_cambios"],
    )


# =========================
# PANEL ADMIN
# =========================
def abrir_panel_configuracion(current_user_id=None, current_role="Superusuario"):
    current_role = normalizar_rol(current_role)
    puede_admin_permisos = usuario_tiene_permiso(current_user_id, current_role, "config.cargos.permisos")
    win = tk.Toplevel()
    win.title("Panel de Configuracion")
    aplicar_icono(win)
    centrar_ventana(win, 980, 680)
    win.configure(bg="#f0f0f0")

    notebook = ttk.Notebook(win)
    notebook.pack(fill="both", expand=True, padx=14, pady=14)

    empresa_tab = tk.Frame(notebook, bg="white")
    licencia_tab = tk.Frame(notebook, bg="white")
    areas_tab = tk.Frame(notebook, bg="white")
    tipos_tab = tk.Frame(notebook, bg="white")
    cargos_tab = tk.Frame(notebook, bg="white")
    permisos_tab = tk.Frame(notebook, bg="white")

    notebook.add(empresa_tab, text="Empresa")
    notebook.add(licencia_tab, text="Licencia")
    notebook.add(areas_tab, text="Areas")
    notebook.add(tipos_tab, text="Tipos Documento")
    notebook.add(cargos_tab, text="Cargos")

    if puede_admin_permisos:
        notebook.add(permisos_tab, text="Permisos Cargos")

    def construir_empresa_tab():
        config = obtener_config_empresa()
        nombre_var = tk.StringVar(value=config["nombre"])
        subtitulo_var = tk.StringVar(value=config["subtitulo"])
        logo_var = tk.StringVar(value=config["logo_path"])
        logo_actual_path = config["logo_path"]

        form = tk.Frame(empresa_tab, bg="white")
        form.pack(fill="x", padx=24, pady=24)
        form.grid_columnconfigure(1, weight=1)

        tk.Label(form, text="Nombre empresa", bg="white", font=("Arial", 11)).grid(row=0, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Entry(form, textvariable=nombre_var, font=("Arial", 11)).grid(row=0, column=1, sticky="ew", pady=8)

        tk.Label(form, text="Subtitulo", bg="white", font=("Arial", 11)).grid(row=1, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Entry(form, textvariable=subtitulo_var, font=("Arial", 11)).grid(row=1, column=1, sticky="ew", pady=8)

        tk.Label(form, text="Ruta logo", bg="white", font=("Arial", 11)).grid(row=2, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Entry(form, textvariable=logo_var, font=("Arial", 10)).grid(row=2, column=1, sticky="ew", pady=8)

        previews = tk.Frame(form, bg="white")
        previews.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 6))
        previews.grid_columnconfigure(0, weight=1)
        previews.grid_columnconfigure(1, weight=1)

        def crear_preview(parent, titulo):
            contenedor = tk.Frame(parent, bg="white", bd=1, relief="solid", width=210, height=180)
            contenedor.grid_propagate(False)
            tk.Label(contenedor, text=titulo, bg="white", fg="#2d5bd1", font=("Arial", 10, "bold")).pack(pady=(10, 6))
            imagen_label = tk.Label(contenedor, text="Sin logo", bg="white", fg="gray", width=22, height=7)
            imagen_label.pack(expand=True)
            ruta_label = tk.Label(contenedor, text="", bg="white", fg="#687180", font=("Arial", 8), wraplength=180)
            ruta_label.pack(pady=(4, 10))
            return contenedor, imagen_label, ruta_label

        actual_box, logo_actual_label, ruta_actual_label = crear_preview(previews, "Logo actual")
        nuevo_box, logo_nuevo_label, ruta_nuevo_label = crear_preview(previews, "Logo nuevo")
        actual_box.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        nuevo_box.grid(row=0, column=1, sticky="ew", padx=(10, 0))

        def cargar_preview(label, ruta_label, ruta, texto_sin_logo="Sin logo"):
            ruta = str(ruta or "").strip()

            if not ruta or not os.path.exists(ruta):
                label.configure(image="", text=texto_sin_logo)
                label.image = None
                ruta_label.configure(text="")
                return

            try:
                imagen = Image.open(ruta)
                imagen.thumbnail((120, 100))
                foto = ImageTk.PhotoImage(imagen)
                label.configure(image=foto, text="")
                label.image = foto
                ruta_label.configure(text=os.path.basename(ruta))
            except Exception as e:
                label.configure(image="", text="No se pudo cargar")
                label.image = None
                ruta_label.configure(text=str(e))

        def actualizar_previews(*_):
            logo_nuevo = logo_var.get().strip()
            cargar_preview(logo_actual_label, ruta_actual_label, logo_actual_path)

            if logo_nuevo and logo_nuevo != logo_actual_path:
                cargar_preview(logo_nuevo_label, ruta_nuevo_label, logo_nuevo, "Seleccione logo")
            else:
                cargar_preview(logo_nuevo_label, ruta_nuevo_label, "", "Sin cambio")

        logo_var.trace_add("write", actualizar_previews)

        def buscar_logo():
            archivo = filedialog.askopenfilename(
                title="Seleccionar logo",
                filetypes=[("Imagenes", "*.png;*.jpg;*.jpeg;*.ico"), ("Todos", "*.*")],
            )
            if archivo:
                logo_var.set(archivo)

        def guardar():
            try:
                logo_nuevo = logo_var.get().strip()
                if logo_nuevo and logo_nuevo != logo_actual_path:
                    if not os.path.exists(logo_nuevo):
                        messagebox.showerror("Error", "La ruta del logo nuevo no existe")
                        return

                    if not messagebox.askyesno(
                        "Confirmar cambio de logo",
                        "Se cambiara el logo actual por el logo nuevo mostrado.\n\nDesea guardar el cambio?",
                        parent=win,
                    ):
                        return

                guardar_config_empresa(nombre_var.get(), subtitulo_var.get(), logo_var.get())
                messagebox.showinfo("OK", "Datos de empresa actualizados")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        botones = tk.Frame(form, bg="white")
        botones.grid(row=4, column=0, columnspan=2, pady=16)
        tk.Button(botones, text="Buscar Logo", command=buscar_logo, width=16).grid(row=0, column=0, padx=8)
        tk.Button(botones, text="Guardar", command=guardar, width=16).grid(row=0, column=1, padx=8)
        actualizar_previews()

    def construir_licencia_tab():
        licencia_tab.grid_columnconfigure(0, weight=1)
        licencia = obtener_estado_licencia()

        panel = tk.Frame(licencia_tab, bg="white")
        panel.pack(fill="x", padx=28, pady=28)
        panel.grid_columnconfigure(1, weight=1)

        estado_var = tk.StringVar()
        cliente_var = tk.StringVar()
        vigencia_var = tk.StringVar()
        equipos_var = tk.StringVar()
        licencia_mask_var = tk.StringVar()
        nueva_licencia_var = tk.StringVar()

        def refrescar():
            datos = obtener_estado_licencia()
            estado_var.set(datos.get("estado", "sin_licencia"))
            cliente_var.set(datos.get("cliente", "Licencia no activada"))
            dias = datos.get("dias_restantes")
            fecha = datos.get("fecha_expiracion", "No informada")
            vigencia_var.set(f"{fecha} | Dias restantes: {dias if dias is not None else 'No informado'}")
            max_equipos = datos.get("max_activaciones") or datos.get("equipos_autorizados") or "No informado"
            equipos_var.set(str(max_equipos))
            licencia_mask_var.set(enmascarar_licencia(datos.get("license_key", "")))

        def activar():
            licencia_ingresada = nueva_licencia_var.get().strip()
            if not licencia_ingresada:
                messagebox.showwarning("Licencia", "Ingrese una licencia para activar", parent=win)
                return

            try:
                respuesta = validar_licencia_online(licencia_ingresada)
            except Exception as e:
                messagebox.showerror("Licencia", str(e), parent=win)
                return

            if not respuesta.get("ok"):
                messagebox.showerror("Licencia", respuesta.get("mensaje", "Licencia no valida"), parent=win)
                return

            nueva_licencia_var.set("")
            refrescar()
            messagebox.showinfo("Licencia", "Software activado correctamente", parent=win)

        def validar_actual():
            datos = obtener_estado_licencia()
            licencia_actual = datos.get("license_key", "")
            if not licencia_actual:
                messagebox.showwarning("Licencia", "No hay licencia guardada", parent=win)
                return
            nueva_licencia_var.set(licencia_actual)
            activar()

        tk.Label(panel, text="Estado", bg="white", font=("Arial", 11)).grid(row=0, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Label(panel, textvariable=estado_var, bg="white", fg="#2d5bd1", font=("Arial", 11, "bold")).grid(row=0, column=1, sticky="w", pady=8)

        tk.Label(panel, text="Propietario", bg="white", font=("Arial", 11)).grid(row=1, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Label(panel, textvariable=cliente_var, bg="white", font=("Arial", 11)).grid(row=1, column=1, sticky="w", pady=8)

        tk.Label(panel, text="Vigencia", bg="white", font=("Arial", 11)).grid(row=2, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Label(panel, textvariable=vigencia_var, bg="white", font=("Arial", 11)).grid(row=2, column=1, sticky="w", pady=8)

        tk.Label(panel, text="Equipos autorizados", bg="white", font=("Arial", 11)).grid(row=3, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Label(panel, textvariable=equipos_var, bg="white", font=("Arial", 11)).grid(row=3, column=1, sticky="w", pady=8)

        tk.Label(panel, text="Licencia", bg="white", font=("Arial", 11)).grid(row=4, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Label(panel, textvariable=licencia_mask_var, bg="white", font=("Arial", 11)).grid(row=4, column=1, sticky="w", pady=8)

        tk.Label(panel, text="Nueva licencia", bg="white", font=("Arial", 11)).grid(row=5, column=0, sticky="e", padx=(0, 12), pady=8)
        tk.Entry(panel, textvariable=nueva_licencia_var, show="*", font=("Arial", 11)).grid(row=5, column=1, sticky="ew", pady=8)

        botones = tk.Frame(panel, bg="white")
        botones.grid(row=6, column=0, columnspan=2, pady=16)
        tk.Button(botones, text="Activar / Cambiar", command=activar, width=18).pack(side="left", padx=8)
        tk.Button(botones, text="Validar Vigencia", command=validar_actual, width=18).pack(side="left", padx=8)

        tk.Label(
            panel,
            text="La licencia completa no se muestra en pantalla. La validacion se realiza contra el servidor configurado.",
            bg="white",
            fg="#687180",
            font=("Arial", 9),
            wraplength=620,
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))
        refrescar()

    def construir_catalogo_tab(parent, titulo, tabla, cargar_func):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        tk.Label(parent, text=titulo, bg="white", fg="#2d5bd1", font=("Arial", 14, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(18, 8)
        )

        lista = tk.Listbox(parent, height=18, font=("Arial", 10))
        lista.grid(row=1, column=0, sticky="nsew", padx=(18, 10), pady=8)

        form = tk.Frame(parent, bg="white")
        form.grid(row=1, column=1, sticky="nsew", padx=(10, 18), pady=8)
        form.grid_columnconfigure(1, weight=1)

        codigo_var = tk.StringVar()
        nombre_var = tk.StringVar()

        tk.Label(form, text="Codigo", bg="white").grid(row=0, column=0, sticky="e", padx=(0, 10), pady=8)
        tk.Entry(form, textvariable=codigo_var).grid(row=0, column=1, sticky="ew", pady=8)
        tk.Label(form, text="Nombre", bg="white").grid(row=1, column=0, sticky="e", padx=(0, 10), pady=8)
        tk.Entry(form, textvariable=nombre_var).grid(row=1, column=1, sticky="ew", pady=8)

        def cargar():
            lista.delete(0, tk.END)
            for item in cargar_func():
                lista.insert(tk.END, item)

        def seleccionar(event=None):
            valor = lista.get(tk.ACTIVE)
            if not valor:
                return
            codigo_var.set(extraer_codigo(valor))
            nombre_var.set(valor.split(" - ", 1)[1] if " - " in valor else "")

        def guardar():
            try:
                guardar_catalogo(tabla, codigo_var.get(), nombre_var.get())
                cargar()
                messagebox.showinfo("OK", "Catalogo actualizado")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def eliminar():
            codigo = codigo_var.get().strip()
            if not codigo:
                messagebox.showwarning("Seleccion requerida", "Seleccione un registro")
                return
            if not messagebox.askyesno("Confirmar", f"Eliminar {codigo}?"):
                return
            try:
                eliminar_catalogo(tabla, codigo)
                codigo_var.set("")
                nombre_var.set("")
                cargar()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo eliminar. Puede estar en uso por documentos.\n\n{e}")

        lista.bind("<<ListboxSelect>>", seleccionar)
        tk.Button(form, text="Guardar / Actualizar", command=guardar, width=20).grid(row=2, column=0, columnspan=2, pady=(18, 8))
        tk.Button(form, text="Eliminar", command=eliminar, width=20).grid(row=3, column=0, columnspan=2, pady=8)
        tk.Button(form, text="Nuevo", command=lambda: [codigo_var.set(""), nombre_var.set("")], width=20).grid(row=4, column=0, columnspan=2, pady=8)
        cargar()

    def construir_cargos_tab():
        cargos_tab.grid_columnconfigure(0, weight=1)
        cargos_tab.grid_columnconfigure(1, weight=1)
        cargos_tab.grid_rowconfigure(1, weight=1)

        tk.Label(cargos_tab, text="Cargos", bg="white", fg="#2d5bd1", font=("Arial", 14, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(18, 8)
        )

        lista = tk.Listbox(cargos_tab, height=18, font=("Arial", 10))
        lista.grid(row=1, column=0, sticky="nsew", padx=(18, 10), pady=8)

        form = tk.Frame(cargos_tab, bg="white")
        form.grid(row=1, column=1, sticky="nsew", padx=(10, 18), pady=8)
        form.grid_columnconfigure(1, weight=1)

        cargo_id_var = tk.StringVar()
        nombre_var = tk.StringVar()
        rol_base_var = tk.StringVar(value="user")

        tk.Label(form, text="ID", bg="white").grid(row=0, column=0, sticky="e", padx=(0, 10), pady=8)
        tk.Entry(form, textvariable=cargo_id_var, state="readonly").grid(row=0, column=1, sticky="ew", pady=8)
        tk.Label(form, text="Cargo", bg="white").grid(row=1, column=0, sticky="e", padx=(0, 10), pady=8)
        tk.Entry(form, textvariable=nombre_var).grid(row=1, column=1, sticky="ew", pady=8)
        tk.Label(form, text="Rol base", bg="white").grid(row=2, column=0, sticky="e", padx=(0, 10), pady=8)
        ttk.Combobox(form, textvariable=rol_base_var, values=ROLES, state="readonly").grid(row=2, column=1, sticky="ew", pady=8)

        def cargar():
            lista.delete(0, tk.END)
            for item in obtener_cargos(activos=True):
                lista.insert(tk.END, item)

        def seleccionar(event=None):
            valor = lista.get(tk.ACTIVE)
            if not valor:
                return
            cargo_id_var.set(str(extraer_id(valor) or ""))
            cargo_id = extraer_id(valor)
            nombre_var.set(valor.split(" - ", 1)[1] if " - " in valor else "")

            if cargo_id:
                try:
                    cargo = obtener_cargo_por_id(cargo_id)
                    if cargo:
                        rol_base_var.set(normalizar_rol(cargo[2]))
                except Exception:
                    rol_base_var.set("user")

        def guardar():
            try:
                guardar_cargo(nombre_var.get(), extraer_id(cargo_id_var.get()), rol_base_var.get())
                cargo_id_var.set("")
                nombre_var.set("")
                rol_base_var.set("user")
                cargar()
                messagebox.showinfo("OK", "Cargo actualizado")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def eliminar():
            cargo_id = extraer_id(cargo_id_var.get())
            if not cargo_id:
                messagebox.showwarning("Seleccion requerida", "Seleccione un cargo")
                return
            if not messagebox.askyesno("Confirmar", "Desactivar este cargo?"):
                return
            try:
                desactivar_cargo(cargo_id)
                cargo_id_var.set("")
                nombre_var.set("")
                rol_base_var.set("user")
                cargar()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        lista.bind("<<ListboxSelect>>", seleccionar)
        tk.Button(form, text="Guardar / Actualizar", command=guardar, width=20).grid(row=3, column=0, columnspan=2, pady=(18, 8))
        tk.Button(form, text="Desactivar", command=eliminar, width=20).grid(row=4, column=0, columnspan=2, pady=8)
        tk.Button(form, text="Nuevo", command=lambda: [cargo_id_var.set(""), nombre_var.set(""), rol_base_var.set("user")], width=20).grid(row=5, column=0, columnspan=2, pady=8)
        cargar()

    def construir_permisos_tab():
        permisos_tab.grid_columnconfigure(0, weight=1)
        permisos_tab.grid_columnconfigure(1, weight=2)
        permisos_tab.grid_rowconfigure(1, weight=1)

        tk.Label(
            permisos_tab,
            text="Permisos por cargo",
            bg="white",
            fg="#2d5bd1",
            font=("Arial", 14, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(18, 4))

        tk.Label(
            permisos_tab,
            text="El rol base entrega permisos heredados; aqui se activan permisos puntuales adicionales para el cargo.",
            bg="white",
            fg="#2f3b4a",
            font=("Arial", 9),
        ).grid(row=1, column=0, columnspan=2, sticky="nw", padx=18, pady=(0, 8))

        lista = tk.Listbox(permisos_tab, height=22, font=("Arial", 10))
        lista.grid(row=2, column=0, sticky="nsew", padx=(18, 10), pady=(8, 18))

        panel = tk.Frame(permisos_tab, bg="white")
        panel.grid(row=2, column=1, sticky="nsew", padx=(10, 18), pady=(8, 18))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(3, weight=1)

        cargo_id_var = tk.StringVar()
        cargo_nombre_var = tk.StringVar(value="Seleccione un cargo")
        cargo_rol_var = tk.StringVar(value="")
        permisos_resumen_var = tk.StringVar(value="")
        permisos_vars = {}
        rol_base_actual = tk.StringVar(value="user")

        tk.Label(panel, textvariable=cargo_nombre_var, bg="white", fg="#222222", font=("Arial", 12, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        tk.Label(panel, textvariable=cargo_rol_var, bg="white", fg="gray", font=("Arial", 9)).grid(
            row=1, column=0, sticky="w", pady=(0, 8)
        )
        tk.Label(panel, textvariable=permisos_resumen_var, bg="white", fg="#2f3b4a", font=("Arial", 9)).grid(
            row=2, column=0, sticky="w", pady=(0, 8)
        )

        canvas = tk.Canvas(panel, bg="white", highlightthickness=0)
        scroll = ttk.Scrollbar(panel, orient="vertical", command=canvas.yview)
        contenido = tk.Frame(canvas, bg="white")
        contenido_id = canvas.create_window((0, 0), window=contenido, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.grid(row=3, column=0, sticky="nsew")
        scroll.grid(row=3, column=1, sticky="ns")

        def ajustar_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(contenido_id, width=canvas.winfo_width())

        contenido.bind("<Configure>", ajustar_scroll)
        canvas.bind("<Configure>", ajustar_scroll)

        fila = 0
        categoria_actual = None
        for codigo, categoria, descripcion in PERMISOS_CATALOGO:
            if categoria != categoria_actual:
                categoria_actual = categoria
                tk.Label(
                    contenido,
                    text=categoria,
                    bg="white",
                    fg="#2d5bd1",
                    font=("Arial", 10, "bold"),
                ).grid(row=fila, column=0, sticky="w", pady=(12, 4))
                fila += 1

            var = tk.BooleanVar(value=False)
            permisos_vars[codigo] = var
            tk.Checkbutton(
                contenido,
                text=f"{codigo} - {descripcion}",
                variable=var,
                bg="white",
                anchor="w",
                justify="left",
                wraplength=560,
            ).grid(row=fila, column=0, sticky="w", pady=1)
            fila += 1

        def cargar_cargos():
            lista.delete(0, tk.END)
            for item in obtener_cargos(activos=True):
                lista.insert(tk.END, item)

        def limpiar_permisos():
            for var in permisos_vars.values():
                var.set(False)
            actualizar_resumen_permisos()

        def permisos_seleccionados():
            return {codigo for codigo, var in permisos_vars.items() if var.get()}

        def actualizar_resumen_permisos():
            heredados = permisos_por_rol(rol_base_actual.get())
            adicionales = permisos_seleccionados()
            efectivos = heredados | adicionales
            permisos_resumen_var.set(
                f"Heredados: {len(heredados)} | Adicionales: {len(adicionales)} | Efectivos: {len(efectivos)}"
            )

        def marcar_permisos(permisos):
            permisos = set(permisos)
            for codigo, var in permisos_vars.items():
                var.set(codigo in permisos)
            actualizar_resumen_permisos()

        def cargar_permisos_rol_base():
            marcar_permisos(permisos_por_rol(rol_base_actual.get()))

        def aplicar_perfil(nombre_perfil):
            marcar_permisos(PERFILES_PERMISOS_CARGO[nombre_perfil])

        for var in permisos_vars.values():
            var.trace_add("write", lambda *_: actualizar_resumen_permisos())

        def seleccionar(event=None):
            valor = lista.get(tk.ACTIVE)
            cargo_id = extraer_id(valor)

            if not cargo_id:
                return

            cargo_id_var.set(str(cargo_id))
            limpiar_permisos()

            try:
                cargo = obtener_cargo_por_id(cargo_id)
                permisos = obtener_permisos_configurados_cargo(cargo_id)
            except Exception as e:
                messagebox.showerror("Error", str(e))
                return

            if cargo:
                cargo_nombre_var.set(cargo[1])
                rol_base_actual.set(normalizar_rol(cargo[2]))
                cargo_rol_var.set(f"Rol base heredado: {rol_base_actual.get()}")

            for permiso in permisos:
                if permiso in permisos_vars:
                    permisos_vars[permiso].set(True)

            actualizar_resumen_permisos()

        def guardar():
            cargo_id = extraer_id(cargo_id_var.get())
            if not cargo_id:
                messagebox.showwarning("Seleccion requerida", "Seleccione un cargo")
                return

            permisos = sorted(permisos_seleccionados())

            try:
                guardar_permisos_cargo(cargo_id, permisos)
                messagebox.showinfo("OK", "Permisos del cargo actualizados")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        botones = tk.Frame(panel, bg="white")
        botones.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        tk.Button(botones, text="Guardar Permisos", command=guardar, width=18).pack(side="left")
        tk.Button(botones, text="Limpiar Seleccion", command=limpiar_permisos, width=18).pack(side="left", padx=(10, 0))
        tk.Button(botones, text="Cargar Rol Base", command=cargar_permisos_rol_base, width=18).pack(side="left", padx=(10, 0))

        perfiles = tk.Frame(panel, bg="white")
        perfiles.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        tk.Label(perfiles, text="Perfiles rapidos:", bg="white", fg="#2f3b4a", font=("Arial", 9)).pack(side="left")
        for nombre_perfil in PERFILES_PERMISOS_CARGO:
            tk.Button(
                perfiles,
                text=nombre_perfil,
                command=lambda perfil=nombre_perfil: aplicar_perfil(perfil),
                width=18,
            ).pack(side="left", padx=(8, 0))

        lista.bind("<<ListboxSelect>>", seleccionar)
        cargar_cargos()
        actualizar_resumen_permisos()

    construir_empresa_tab()
    construir_licencia_tab()
    construir_catalogo_tab(areas_tab, "Areas / Departamentos", "areas", obtener_areas)
    construir_catalogo_tab(tipos_tab, "Tipos de Documento", "tipos_documento", obtener_tipos)
    construir_cargos_tab()
    if puede_admin_permisos:
        construir_permisos_tab()


def abrir_panel_admin(current_user_id, current_role):
    if not puede_administrar_usuarios(current_role, current_user_id):
        messagebox.showerror("Permiso denegado", "No tiene permisos para administrar usuarios")
        return

    admin = tk.Toplevel()
    admin.title("Administracion de Usuarios")
    aplicar_icono(admin)
    centrar_ventana(admin, 1180, 560)

    lista = tk.Listbox(admin, width=155, height=14, font=("Consolas", 9))
    lista.pack(pady=10)

    def usuario_seleccionado():
        seleccionado = lista.get(tk.ACTIVE)

        if not seleccionado:
            return None

        partes = [p.strip() for p in seleccionado.split("|")]

        if len(partes) < 7:
            return None

        return {
            "id": partes[0],
            "username": partes[1],
            "nombre": partes[2],
            "email": partes[3],
            "cargo": partes[4],
            "rol": partes[5],
            "activo": partes[6] == "Activo",
        }

    def validar_permiso_sobre_usuario(datos_usuario, accion):
        if not datos_usuario:
            messagebox.showwarning("Seleccion requerida", "Seleccione un usuario")
            return False

        if es_superusuario(datos_usuario["rol"]) and not es_superusuario(current_role):
            messagebox.showerror("Permiso denegado", f"Solo un Superusuario puede {accion} otro Superusuario")
            return False

        return True

    def cargar_usuarios():
        lista.delete(0, tk.END)

        try:
            with conectar() as conn:
                with conn.cursor() as cur:
                    asegurar_configuracion_sistema(cur)
                    cur.execute(
                        """
                        SELECT
                            u.id,
                            u.username,
                            COALESCE(u.nombre, '') AS nombre,
                            COALESCE(u.email, '') AS email,
                            COALESCE(c.nombre, 'Sin cargo') AS cargo,
                            COALESCE(u.rol, 'user') AS rol,
                            COALESCE(u.activo, true) AS activo
                        FROM usuarios u
                        LEFT JOIN cargos c ON c.id = u.cargo_id
                        ORDER BY u.id
                        """
                    )
                    usuarios = cur.fetchall()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        for u in usuarios:
            estado = "Activo" if u[6] else "Inactivo"
            lista.insert(tk.END, f"{u[0]} | {u[1]} | {u[2]} | {u[3]} | {u[4]} | {u[5]} | {estado}")

    def crear_usuario():
        win = tk.Toplevel(admin)
        win.title("Crear Usuario")
        aplicar_icono(win)
        centrar_ventana(win, 430, 520)

        tk.Label(win, text="Nombre completo").pack(pady=(12, 3))
        nombre_entry = tk.Entry(win, width=38)
        nombre_entry.pack()

        tk.Label(win, text="Usuario").pack(pady=(10, 3))
        u = tk.Entry(win, width=38)
        u.pack()

        tk.Label(win, text="Email").pack(pady=(10, 3))
        e = tk.Entry(win, width=38)
        e.pack()

        tk.Label(win, text="Cargo").pack(pady=(10, 3))
        cargos = ["Sin cargo"] + obtener_cargos(activos=True)
        cargo_var = tk.StringVar(value=cargos[0])
        ttk.Combobox(win, textvariable=cargo_var, values=cargos, state="readonly", width=35).pack()

        tk.Label(win, text="Password").pack(pady=(10, 3))
        p = tk.Entry(win, show="*", width=38)
        p.pack()

        tk.Label(win, text="Repetir password").pack(pady=(10, 3))
        p2 = tk.Entry(win, show="*", width=38)
        p2.pack()

        mostrar_var = tk.BooleanVar(value=False)

        def alternar_password():
            show = "" if mostrar_var.get() else "*"
            p.configure(show=show)
            p2.configure(show=show)

        tk.Checkbutton(win, text="Mostrar contrasenas", variable=mostrar_var, command=alternar_password).pack(pady=(6, 0))

        tk.Label(win, text="Rol").pack(pady=(10, 3))
        roles_disponibles = ROLES if es_superusuario(current_role) else ROLES_ADMIN_CREACION
        rol_var = tk.StringVar(value="user")
        ttk.Combobox(
            win,
            textvariable=rol_var,
            values=roles_disponibles,
            state="readonly",
            width=29,
        ).pack()

        def guardar():
            if p.get() != p2.get():
                messagebox.showerror("Error", "Las contrasenas no coinciden")
                return

            cargo_id = None if cargo_var.get() == "Sin cargo" else extraer_id(cargo_var.get())
            if registrar_usuario(u.get(), p.get(), e.get(), rol_var.get(), nombre_entry.get(), cargo_id):
                win.destroy()
                cargar_usuarios()

        tk.Button(win, text="Guardar", command=guardar, width=16).pack(pady=14)

    def cambiar_estado_usuario(activo):
        datos = usuario_seleccionado()

        if not validar_permiso_sobre_usuario(datos, "modificar"):
            return

        if str(current_user_id) == str(datos["id"]):
            messagebox.showerror("Accion no permitida", "No puede desactivar o reactivar su propio usuario activo")
            return

        accion = "reactivar" if activo else "desactivar"

        if messagebox.askyesno("Confirmar", f"Desea {accion} este usuario?"):
            try:
                with conectar() as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            """
                            UPDATE usuarios
                            SET activo = %s
                            WHERE id = %s
                            """,
                            (activo, datos["id"]),
                        )
                cargar_usuarios()
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def cambiar_rol_usuario():
        datos = usuario_seleccionado()

        if not validar_permiso_sobre_usuario(datos, "cambiar el rol de"):
            return

        win = tk.Toplevel(admin)
        win.title("Cambiar Rol")
        aplicar_icono(win)
        centrar_ventana(win, 340, 170)

        tk.Label(win, text=f"Usuario: {datos['username']}").pack(pady=(14, 6))

        roles_disponibles = ROLES if es_superusuario(current_role) else ROLES_ADMIN_CREACION
        rol_var = tk.StringVar(value=normalizar_rol(datos["rol"]))

        ttk.Combobox(
            win,
            textvariable=rol_var,
            values=roles_disponibles,
            state="readonly",
            width=28,
        ).pack(pady=4)

        def guardar():
            nuevo_rol = normalizar_rol(rol_var.get())

            if nuevo_rol == "Superusuario" and not es_superusuario(current_role):
                messagebox.showerror("Permiso denegado", "Solo un Superusuario puede asignar ese rol")
                return

            try:
                with conectar() as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            """
                            UPDATE usuarios
                            SET rol = %s
                            WHERE id = %s
                            """,
                            (nuevo_rol, datos["id"]),
                        )
                messagebox.showinfo("OK", "Rol actualizado")
                win.destroy()
                cargar_usuarios()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        tk.Button(win, text="Guardar", command=guardar, width=16).pack(pady=12)

    def editar_datos_usuario():
        datos = usuario_seleccionado()

        if not validar_permiso_sobre_usuario(datos, "editar"):
            return

        win = tk.Toplevel(admin)
        win.title("Editar Datos Usuario")
        aplicar_icono(win)
        centrar_ventana(win, 430, 300)

        tk.Label(win, text=f"Usuario: {datos['username']}", font=("Arial", 10, "bold")).pack(pady=(12, 8))

        tk.Label(win, text="Nombre completo").pack(pady=(4, 3))
        nombre_var = tk.StringVar(value="" if datos["nombre"] == "" else datos["nombre"])
        tk.Entry(win, textvariable=nombre_var, width=38).pack()

        tk.Label(win, text="Email").pack(pady=(10, 3))
        email_var = tk.StringVar(value=datos["email"])
        tk.Entry(win, textvariable=email_var, width=38).pack()

        tk.Label(win, text="Cargo").pack(pady=(10, 3))
        cargos = ["Sin cargo"] + obtener_cargos(activos=True)
        cargo_actual = next((c for c in cargos if c.endswith(f" - {datos['cargo']}")), cargos[0])
        cargo_var = tk.StringVar(value=cargo_actual)
        ttk.Combobox(win, textvariable=cargo_var, values=cargos, state="readonly", width=35).pack()

        def guardar():
            cargo_id = None if cargo_var.get() == "Sin cargo" else extraer_id(cargo_var.get())
            try:
                with conectar() as conn:
                    with conn.cursor() as cur:
                        asegurar_configuracion_sistema(cur)
                        cur.execute(
                            """
                            UPDATE usuarios
                            SET nombre = %s,
                                email = %s,
                                cargo_id = %s
                            WHERE id = %s
                            """,
                            (nombre_var.get().strip(), email_var.get().strip(), cargo_id, datos["id"]),
                        )
                messagebox.showinfo("OK", "Datos de usuario actualizados")
                win.destroy()
                cargar_usuarios()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        tk.Button(win, text="Guardar", command=guardar, width=16).pack(pady=16)

    def cambiar_password():
        datos = usuario_seleccionado()

        if not validar_permiso_sobre_usuario(datos, "cambiar la contrasena de"):
            return

        win = tk.Toplevel(admin)
        win.title("Nueva Contrasena")
        aplicar_icono(win)
        centrar_ventana(win, 360, 260)

        tk.Label(win, text="Nueva contrasena").pack(pady=(14, 4))
        p = tk.Entry(win, show="*", width=30)
        p.pack()

        tk.Label(win, text="Repetir contrasena").pack(pady=(10, 4))
        p2 = tk.Entry(win, show="*", width=30)
        p2.pack()

        mostrar_var = tk.BooleanVar(value=False)

        def alternar_password():
            show = "" if mostrar_var.get() else "*"
            p.configure(show=show)
            p2.configure(show=show)

        tk.Checkbutton(win, text="Mostrar contrasenas", variable=mostrar_var, command=alternar_password).pack(pady=(8, 0))

        def guardar():
            if not p.get():
                messagebox.showerror("Error", "Debe ingresar una contrasena")
                return

            if p.get() != p2.get():
                messagebox.showerror("Error", "Las contrasenas no coinciden")
                return

            error_password = validar_password_plana(p.get())
            if error_password:
                messagebox.showerror("Error", error_password)
                return

            try:
                with conectar() as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            """
                            UPDATE usuarios
                            SET password = %s
                            WHERE id = %s
                            """,
                            (hash_password(p.get()), datos["id"]),
                        )
                messagebox.showinfo("OK", "Contrasena actualizada")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        tk.Button(win, text="Guardar", command=guardar, width=16).pack(pady=12)

    botones = tk.Frame(admin)
    botones.pack(pady=5)

    tk.Button(botones, text="Crear Usuario", command=crear_usuario, width=18).grid(
        row=0, column=0, padx=5, pady=5
    )
    tk.Button(botones, text="Desactivar Usuario", command=lambda: cambiar_estado_usuario(False), width=18).grid(
        row=0, column=1, padx=5, pady=5
    )
    tk.Button(botones, text="Reactivar Usuario", command=lambda: cambiar_estado_usuario(True), width=18).grid(
        row=0, column=2, padx=5, pady=5
    )
    tk.Button(botones, text="Cambiar Rol", command=cambiar_rol_usuario, width=18).grid(
        row=1, column=0, padx=5, pady=5
    )
    tk.Button(botones, text="Cambiar Contrasena", command=cambiar_password, width=18).grid(
        row=1, column=1, padx=5, pady=5
    )
    tk.Button(botones, text="Editar Datos", command=editar_datos_usuario, width=18).grid(
        row=1, column=2, padx=5, pady=5
    )
    tk.Button(
        botones,
        text="Configuracion",
        command=lambda: abrir_panel_configuracion(current_user_id, current_role),
        width=18,
    ).grid(
        row=2, column=0, padx=5, pady=5
    )
    tk.Button(botones, text="Refrescar", command=cargar_usuarios, width=18).grid(
        row=2, column=1, padx=5, pady=5
    )

    cargar_usuarios()


def mostrar_acerca_de(parent=None):
    licencia = obtener_estado_licencia()
    mensaje = "\n".join(
        [
            f"GCDRYS-2 {APP_VERSION}",
            APP_COPYRIGHT,
            "",
            "Software propietario para gestion y control documental.",
            "Queda prohibida su copia, modificacion, distribucion o uso no autorizado.",
            "",
            f"Licencia a nombre de: {licencia.get('cliente', 'Licencia no activada')}",
            f"Licencia: {enmascarar_licencia(licencia.get('license_key', ''))}",
        ]
    )
    messagebox.showinfo("Acerca de GCDRYS-2", mensaje, parent=parent)


def abrir_activacion_licencia(parent=None, on_success=None):
    win_lic = tk.Toplevel(parent or root)
    win_lic.title("Activar Licencia")
    aplicar_icono(win_lic)
    centrar_ventana(win_lic, 460, 240)
    win_lic.configure(bg="white")
    win_lic.transient(parent or root)

    licencia_var = tk.StringVar()

    tk.Label(
        win_lic,
        text="Activacion de licencia",
        bg="white",
        fg="#2d5bd1",
        font=("Arial", 14, "bold"),
    ).pack(pady=(20, 10))

    tk.Label(
        win_lic,
        text="Ingrese la licencia entregada al cliente.",
        bg="white",
        font=("Arial", 10),
    ).pack(pady=(0, 8))

    entry = tk.Entry(win_lic, textvariable=licencia_var, show="*", width=42, font=("Arial", 11))
    entry.pack(ipady=3, pady=6)
    entry.focus_set()

    def activar():
        licencia_ingresada = licencia_var.get().strip()
        if not licencia_ingresada:
            messagebox.showwarning("Licencia", "Ingrese una licencia para activar", parent=win_lic)
            return

        try:
            respuesta = validar_licencia_online(licencia_ingresada)
        except Exception as e:
            messagebox.showerror("Licencia", str(e), parent=win_lic)
            return

        if not respuesta.get("ok"):
            messagebox.showerror("Licencia", respuesta.get("mensaje", "Licencia no valida"), parent=win_lic)
            return

        messagebox.showinfo("Licencia", "Software activado correctamente", parent=win_lic)
        win_lic.destroy()
        if on_success:
            on_success()

    botones = tk.Frame(win_lic, bg="white")
    botones.pack(pady=16)
    tk.Button(botones, text="Activar", command=activar, width=16).pack(side="left", padx=8)
    tk.Button(botones, text="Cerrar", command=win_lic.destroy, width=12).pack(side="left", padx=8)
    win_lic.bind("<Return>", lambda event: activar())


# =========================
# UI PRINCIPAL
# =========================
def abrir_app(user_id, username, rol):
    rol = normalizar_rol(rol)
    app = tk.Toplevel()
    app.title(APP_TITLE)
    aplicar_icono(app)
    app.configure(bg="#f0f0f0")
    centrar_ventana(app, 1060, 730)

    app.grid_columnconfigure(0, weight=1, minsize=320)
    app.grid_columnconfigure(1, weight=3, minsize=660)
    app.grid_rowconfigure(0, weight=1)

    left_frame = tk.Frame(app, bg="white", bd=1, relief="solid")
    left_frame.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")

    right_frame = tk.Frame(app, bg="white", bd=1, relief="solid")
    right_frame.grid(row=0, column=1, padx=(0, 15), pady=15, sticky="nsew")
    right_frame.grid_columnconfigure(0, weight=0)
    right_frame.grid_columnconfigure(1, weight=1)

    bottom_frame = tk.Frame(app, bg="#d9d9d9", height=30)
    bottom_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
    bottom_frame.grid_propagate(False)

    logo_app_tk = cargar_logo_empresa((160, 160))
    if logo_app_tk:
        logo_app = tk.Label(left_frame, image=logo_app_tk, bg="white")
        logo_app.image = logo_app_tk
        logo_app.pack(pady=(12, 6))

    tk.Label(
        left_frame,
        text="GCDRYS-2",
        font=("Arial", 26, "bold"),
        bg="white",
    ).pack(pady=4)

    tk.Frame(left_frame, bg="#2d89ef", height=2, width=250).pack(pady=8)

    fecha_actual = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    tk.Label(
        left_frame,
        text=f"Usuario activo: {username}",
        font=("Arial", 11),
        bg="white",
    ).pack(pady=3)

    tk.Label(
        left_frame,
        text=f"Rol: {rol}",
        font=("Arial", 11),
        bg="white",
    ).pack(pady=3)

    tk.Label(
        left_frame,
        text=f"Fecha: {fecha_actual}",
        font=("Arial", 11),
        bg="white",
    ).pack(pady=3)

    tk.Label(
        left_frame,
        text="Estado: Conectado a Supabase",
        font=("Arial", 11, "bold"),
        fg="green",
        bg="white",
    ).pack(pady=3)

    tk.Label(
        left_frame,
        text=f"Version del sistema: {APP_VERSION}",
        font=("Arial", 10, "italic"),
        bg="white",
    ).pack(pady=3)

    menu_frame = tk.Frame(left_frame, bg="white")
    menu_frame.pack(fill="x", padx=28, pady=(14, 0))

    def boton_menu(texto, comando):
        tk.Button(
            menu_frame,
            text=texto,
            font=("Arial", 10),
            width=24,
            command=comando,
        ).pack(pady=4)

    if puede_administrar_usuarios(rol, user_id):
        boton_menu("Administrar Usuarios", lambda: abrir_panel_admin(user_id, rol))
        boton_menu("Configuracion Sistema", lambda: abrir_panel_configuracion(user_id, rol))
        boton_menu("Diagnostico Sistema", lambda: mostrar_diagnostico_sistema(app))

    if usuario_tiene_permiso(user_id, rol, "dashboard.ver"):
        boton_menu("Dashboard", lambda: abrir_dashboard(user_id, username, rol))

    if usuario_tiene_permiso(user_id, rol, "consulta.ver"):
        boton_menu("Administrador Documental", lambda: abrir_consulta_documental(user_id, rol, username))

    boton_menu("Acerca de / Copyright", lambda: mostrar_acerca_de(app))

    area_var = tk.StringVar()
    tipo_var = tk.StringVar()
    estado_doc_var = tk.StringVar(value="Borrador")
    resultado_var = tk.StringVar()
    auto_num_var = tk.BooleanVar(value=True)
    confidencial_var = tk.BooleanVar(value=False)

    try:
        areas = obtener_areas()
        tipos = obtener_tipos()
    except Exception as e:
        messagebox.showerror("Error", str(e))
        app.destroy()
        return

    if not areas or not tipos:
        messagebox.showerror("Error", "No hay areas o tipos en la base de datos")
        app.destroy()
        return

    area_var.set(areas[0])
    tipo_var.set(tipos[0])
    estados_creacion = estados_creacion_permitidos(rol)
    estado_doc_var.set(estados_creacion[0])

    tk.Label(
        right_frame,
        text="GENERACION DE CODIGO DOCUMENTAL",
        font=("Arial", 18),
        fg="#2d5bd1",
        bg="white",
    ).grid(row=0, column=0, columnspan=2, pady=(24, 10))

    tk.Frame(right_frame, bg="#2d5bd1", height=2, width=520).grid(
        row=1, column=0, columnspan=2, pady=(0, 20)
    )

    def generar_codigo():
        if not puede_crear_documentos(rol, user_id):
            messagebox.showerror("Permiso denegado", "No tiene permisos para crear documentos")
            return

        area_codigo = extraer_codigo(area_var.get())
        tipo_codigo = extraer_codigo(tipo_var.get())

        if auto_num_var.get():
            try:
                numero = obtener_siguiente_numero(area_codigo, tipo_codigo)
            except Exception as e:
                messagebox.showerror("Error", str(e))
                return

            numero_entry.delete(0, tk.END)
            numero_entry.insert(0, numero)
        else:
            numero = numero_entry.get().strip().zfill(3)
            numero_entry.delete(0, tk.END)
            numero_entry.insert(0, numero)

        version = version_entry.get().strip()

        if not version:
            messagebox.showerror("Error", "Debe ingresar la version")
            return

        codigo = f"{DOC_PREFIX}-{area_codigo}-{tipo_codigo}-{numero}-V{version}"
        resultado_var.set(codigo)

        try:
            if codigo_existe_db(codigo):
                messagebox.showwarning("Duplicado", "El codigo ya existe")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def guardar_codigo():
        if not puede_crear_documentos(rol, user_id):
            messagebox.showerror("Permiso denegado", "No tiene permisos para guardar documentos")
            return

        codigo = resultado_var.get().strip()
        nombre = nombre_entry.get().strip()
        area = extraer_codigo(area_var.get())
        tipo = extraer_codigo(tipo_var.get())
        numero = numero_entry.get().strip()
        version = version_entry.get().strip()
        estado = estado_doc_var.get()
        confidencial = confidencial_var.get()

        if confidencial and not usuario_tiene_permiso(user_id, rol, "documentos.confidencial.crear"):
            messagebox.showerror("Permiso denegado", "No tiene permisos para crear documentos confidenciales")
            return

        if estado not in estados_creacion_permitidos(rol):
            messagebox.showerror("Permiso denegado", "No tiene permisos para crear documentos con ese estado")
            return

        if not nombre:
            messagebox.showerror("Error", "Debe ingresar el nombre del documento")
            return

        if not codigo:
            messagebox.showerror("Error", "Debe generar el codigo antes de guardar")
            return

        if not numero or not version:
            messagebox.showerror("Error", "Debe completar numero y version")
            return

        if guardar_en_db(codigo, nombre, area, tipo, numero, version, estado, user_id, confidencial):
            resultado_var.set("")
            nombre_entry.delete(0, tk.END)
            numero_entry.delete(0, tk.END)
            version_entry.delete(0, tk.END)
            version_entry.insert(0, "1")
            estado_doc_var.set(estados_creacion_permitidos(rol)[0])
            confidencial_var.set(False)

    def cerrar_app():
        root.destroy()

    form = tk.Frame(right_frame, bg="white")
    form.grid(row=2, column=0, columnspan=2, padx=35, pady=5, sticky="ew")
    form.grid_columnconfigure(0, weight=0)
    form.grid_columnconfigure(1, weight=1)

    label_opts = {"font": ("Arial", 12), "bg": "white"}
    entry_opts = {"width": 42, "font": ("Arial", 12)}

    tk.Label(form, text="Nombre Documento", **label_opts).grid(
        row=0, column=0, padx=(0, 18), pady=10, sticky="e"
    )
    nombre_entry = tk.Entry(form, **entry_opts)
    nombre_entry.grid(row=0, column=1, padx=0, pady=10, sticky="ew")

    tk.Label(form, text="Area", **label_opts).grid(
        row=1, column=0, padx=(0, 18), pady=10, sticky="e"
    )
    ttk.Combobox(
        form,
        textvariable=area_var,
        values=areas,
        width=39,
        state="readonly",
        font=("Arial", 11),
    ).grid(row=1, column=1, padx=0, pady=10, sticky="w")

    tk.Label(form, text="Tipo", **label_opts).grid(
        row=2, column=0, padx=(0, 18), pady=10, sticky="e"
    )
    ttk.Combobox(
        form,
        textvariable=tipo_var,
        values=tipos,
        width=39,
        state="readonly",
        font=("Arial", 11),
    ).grid(row=2, column=1, padx=0, pady=10, sticky="w")

    tk.Label(form, text="Numero", **label_opts).grid(
        row=3, column=0, padx=(0, 18), pady=10, sticky="e"
    )
    numero_entry = tk.Entry(form, **entry_opts)
    numero_entry.grid(row=3, column=1, padx=0, pady=10, sticky="ew")

    tk.Label(form, text="Version", **label_opts).grid(
        row=4, column=0, padx=(0, 18), pady=10, sticky="e"
    )
    version_entry = tk.Entry(form, **entry_opts)
    version_entry.grid(row=4, column=1, padx=0, pady=10, sticky="ew")
    version_entry.insert(0, "1")

    tk.Label(form, text="Estado", **label_opts).grid(
        row=5, column=0, padx=(0, 18), pady=10, sticky="e"
    )
    ttk.Combobox(
        form,
        textvariable=estado_doc_var,
        values=estados_creacion,
        width=39,
        state="readonly",
        font=("Arial", 11),
    ).grid(row=5, column=1, padx=0, pady=10, sticky="w")

    tk.Checkbutton(
        form,
        text="Numero Automatico",
        variable=auto_num_var,
        bg="white",
        font=("Arial", 11),
    ).grid(row=6, column=0, columnspan=2, pady=12)

    tk.Checkbutton(
        form,
        text="Documento confidencial",
        variable=confidencial_var,
        state="normal" if usuario_tiene_permiso(user_id, rol, "documentos.confidencial.crear") else "disabled",
        bg="white",
        fg="#8e1b1b",
        font=("Arial", 11, "bold"),
    ).grid(row=7, column=0, columnspan=2, pady=(0, 12))

    tk.Button(
        form,
        text="Generar Codigo",
        command=generar_codigo,
        width=22,
        height=2,
    ).grid(row=8, column=0, columnspan=2, pady=(4, 14))

    tk.Entry(
        form,
        textvariable=resultado_var,
        width=48,
        justify="center",
        font=("Arial", 12, "bold"),
        state="readonly",
    ).grid(row=9, column=0, columnspan=2, pady=10, sticky="ew")

    botones_frame = tk.Frame(form, bg="white")
    botones_frame.grid(row=10, column=0, columnspan=2, pady=20)

    tk.Button(
        botones_frame,
        text="Guardar",
        command=guardar_codigo,
        width=16,
        height=2,
    ).grid(row=0, column=0, padx=12)

    tk.Button(
        botones_frame,
        text="Exportar Excel",
        command=exportar_excel,
        width=16,
        height=2,
    ).grid(row=0, column=1, padx=12)

    if usuario_tiene_permiso(user_id, rol, "reportes.general.exportar"):
        tk.Button(
            botones_frame,
            text="Exportar PDF",
            command=lambda: exportar_pdf(user_id, username),
            width=16,
            height=2,
        ).grid(row=0, column=2, padx=12)

    if usuario_tiene_permiso(user_id, rol, "reportes.general.imprimir"):
        tk.Button(
            botones_frame,
            text="Imprimir PDF",
            command=lambda: exportar_pdf(user_id, username, imprimir=True),
            width=16,
            height=2,
        ).grid(row=0, column=3, padx=12)

    tk.Button(
        app,
        text="Salir",
        command=cerrar_app,
        bg="darkred",
        fg="white",
        width=12,
        height=2,
    ).place(relx=0.91, rely=0.89, anchor="center")

    tk.Label(
        bottom_frame,
        text="Conectado a Supabase",
        bg="#d9d9d9",
        fg="green",
        font=("Arial", 10),
    ).pack(side="left", padx=20)

    tk.Label(
        bottom_frame,
        text=f"Usuario: {username} | Rol: {rol}",
        bg="#d9d9d9",
        font=("Arial", 10),
    ).pack(side="left", padx=40)

    tk.Label(
        bottom_frame,
        text=f"GCDRYS-2 | {APP_VERSION}",
        bg="#d9d9d9",
        font=("Arial", 10),
    ).pack(side="right", padx=20)


# =========================
# LOGIN UI
# =========================
root = tk.Tk()
root.title("Login GCDRYS")
root.configure(bg="white")
aplicar_icono(root)
centrar_ventana(root, 440, 680)
root.minsize(420, 650)

logo_login_tk = cargar_logo_empresa((140, 140))
if logo_login_tk:
    logo_login = tk.Label(root, image=logo_login_tk, bg="white")
    logo_login.image = logo_login_tk
    logo_login.pack(pady=(22, 12))

tk.Label(
    root,
    text="GCDRYS-2",
    font=("Arial", 24, "bold"),
    bg="white",
    fg="#2d5bd1",
).pack(pady=10)

estado_acceso_var = tk.StringVar()
detalle_acceso_var = tk.StringVar()


def refrescar_estado_acceso_login():
    estado = obtener_estado_acceso_software()
    estado_acceso_var.set(estado["mensaje"])
    detalle_acceso_var.set(estado["detalle"])
    color = "#1f8f45" if estado["modo"] == "licenciado" else "#d79b00" if estado["permitido"] else "#c0392b"
    estado_acceso_label.configure(fg=color)


estado_acceso_label = tk.Label(
    root,
    textvariable=estado_acceso_var,
    bg="white",
    fg="#1f8f45",
    font=("Arial", 10, "bold"),
    wraplength=360,
)
estado_acceso_label.pack(pady=(0, 3))

tk.Label(
    root,
    textvariable=detalle_acceso_var,
    bg="white",
    fg="#687180",
    font=("Arial", 9),
    wraplength=360,
).pack(pady=(0, 8))

tk.Label(root, text="Usuario", bg="white", font=("Arial", 11)).pack(pady=(16, 5))

user_entry = tk.Entry(root, width=30, font=("Arial", 11))
user_entry.pack(ipady=3)

tk.Label(root, text="Contrasena", bg="white", font=("Arial", 11)).pack(pady=(16, 5))

pass_entry = tk.Entry(root, show="*", width=30, font=("Arial", 11))
pass_entry.pack(ipady=3)


def login():
    acceso = obtener_estado_acceso_software()

    if not acceso["permitido"]:
        refrescar_estado_acceso_login()
        messagebox.showerror(
            "Licencia requerida",
            acceso["mensaje"] + "\n\n" + acceso["detalle"],
        )
        abrir_activacion_licencia(root, refrescar_estado_acceso_login)
        return

    user = user_entry.get()
    password = pass_entry.get()

    resultado = validar_usuario(user, password)

    if resultado:
        messagebox.showinfo("OK", "Bienvenido")
        root.withdraw()
        abrir_app(resultado[0], resultado[1], resultado[2])
    else:
        messagebox.showerror("Error", "Credenciales incorrectas")


def abrir_registro():
    reg = tk.Toplevel()
    reg.title("Registro")
    aplicar_icono(reg)
    centrar_ventana(reg, 350, 300)

    tk.Label(reg, text="Usuario").pack(pady=5)
    u = tk.Entry(reg, width=30)
    u.pack()

    tk.Label(reg, text="Password").pack(pady=5)
    p = tk.Entry(reg, show="*", width=30)
    p.pack()

    tk.Label(reg, text="Email").pack(pady=5)
    e = tk.Entry(reg, width=30)
    e.pack()

    tk.Button(
        reg,
        text="Registrar",
        command=lambda: reg.destroy() if registrar_usuario(u.get(), p.get(), e.get(), "user") else None,
        width=18,
    ).pack(pady=20)


tk.Button(
    root,
    text="Login",
    command=login,
    width=20,
    height=2,
    bg="#2d89ef",
    fg="white",
).pack(pady=25)

tk.Button(
    root,
    text="Activar Licencia",
    command=lambda: abrir_activacion_licencia(root, refrescar_estado_acceso_login),
    width=20,
).pack(pady=(0, 10))

tk.Button(root, text="Registrar", command=abrir_registro, width=20).pack(pady=(0, 18))

tk.Button(root, text="Diagnostico Sistema", command=lambda: mostrar_diagnostico_sistema(root), width=20).pack(pady=(0, 18))

tk.Label(
    root,
    text="GESTION Y CONTROL DOCUMENTAL",
    bg="white",
    fg="gray",
).pack(pady=(4, 18))

root.bind("<Return>", lambda event: login())
refrescar_estado_acceso_login()
root.mainloop()



