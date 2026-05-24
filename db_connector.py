import os
import hmac
import hashlib
import psycopg2
from psycopg2 import sql, extras

# Configuración desde variables de entorno
DB_HOST = os.getenv("RYS_DB_HOST", "aws-1-us-east-1.pooler.supabase.com")
DB_NAME = os.getenv("RYS_DB_NAME", "postgres")
DB_USER = os.getenv("RYS_DB_USER", "postgres.xgajgsranipwivnspsot")
DB_PASSWORD = os.getenv("RYS_DB_PASSWORD", "Vjd2iGyQ3dxJmMNT")
DB_PORT = int(os.getenv("RYS_DB_PORT") or "6543")
PASSWORD_HASH_ITERATIONS = 100000 # Valor por defecto usado en GCDRYS-2

class SupabaseConnector:
    def __init__(self):
        self.conn_params = {
            "host": DB_HOST,
            "database": DB_NAME,
            "user": DB_USER,
            "password": DB_PASSWORD,
            "port": DB_PORT,
            "sslmode": "require"
        }

    def conectar(self):
        if not DB_HOST or not DB_USER or not DB_PASSWORD:
            raise RuntimeError(
                "Faltan variables de entorno de base de datos. "
                "Asegúrate de configurar RYS_DB_HOST, RYS_DB_USER y RYS_DB_PASSWORD."
            )
        return psycopg2.connect(**self.conn_params)

    def validar_credenciales(self, username, password):
        """
        Valida usuario y contraseña replicando la lógica de GCDRYS-2.py.
        Retorna: (user_id, username, rol) o None si falla.
        """
        username = username.strip()
        try:
            with self.conectar() as conn:
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        SELECT id, username, COALESCE(rol, 'user') AS rol, password
                        FROM usuarios
                        WHERE username = %s AND COALESCE(activo, true) = true
                        """,
                        (username,),
                    )
                    resultado = cur.fetchone()

                    if not resultado:
                        return None

                    user_id, usuario, rol, hash_guardado = resultado
                    
                    # Verificación de password (PBKDF2 o Legacy)
                    if self._verificar_password(password, hash_guardado):
                        # Si era legacy y validó, actualizamos a PBKDF2
                        if hash_guardado and not hash_guardado.startswith("pbkdf2_sha256$"):
                             cur.execute(
                                "UPDATE usuarios SET password = %s WHERE id = %s",
                                (self._hash_password(password), user_id)
                            )
                        return user_id, usuario, rol
                    else:
                        return None

        except Exception as e:
            print(f"Error en validación: {e}")
            return None

    def _hash_password(self, password):
        salt = os.urandom(16).hex()
        digest = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            bytes.fromhex(salt),
            PASSWORD_HASH_ITERATIONS,
        ).hex()
        return f"pbkdf2_sha256${PASSWORD_HASH_ITERATIONS}${salt}${digest}"

    def _verificar_password(self, password, hash_guardado):
        if not hash_guardado:
            return False

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
                return False
            return hmac.compare_digest(calculado, digest)
        
        # Legacy SHA256
        calculado_legacy = hashlib.sha256(password.encode("utf-8")).hexdigest()
        return hmac.compare_digest(calculado_legacy, hash_guardado)

    # ... (resto de métodos: obtener_areas, diagnosticar, etc.)
    def obtener_areas(self):
        try:
            with self.conectar() as conn:
                with conn.cursor(cursor_factory=extras.RealDictCursor) as cur:
                    cur.execute("SELECT id, nombre FROM areas ORDER BY nombre ASC")
                    return cur.fetchall()
        except Exception as e:
            print(f"Error al obtener áreas: {e}")
            return []

    def diagnosticar(self):
        try:
            with self.conectar() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT current_database(), current_user")
                    db, user = cur.fetchone()
                    return {"ok": True, "msg": f"Conectado a {db} como {user}"}
        except Exception as e:
            return {"ok": False, "msg": str(e)}
