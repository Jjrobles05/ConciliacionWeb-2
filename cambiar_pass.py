import hashlib
import secrets
import sqlite3
import streamlit as st

# Configura aquí el usuario y la nueva contraseña
USUARIO = "angie mejía p"
NUEVA_CLAVE = "Angie1697."

def conectar():
    if "TURSO_DATABASE_URL" in st.secrets and "TURSO_AUTH_TOKEN" in st.secrets:
        url = str(st.secrets["TURSO_DATABASE_URL"]).strip()
        token = str(st.secrets["TURSO_AUTH_TOKEN"]).strip()
        if url.startswith("libsql://"):
            url = "https://" + url[len("libsql://"): ]
        import turso_serverless
        return turso_serverless.connect(url, auth_token=token)
    return sqlite3.connect("conciliaciones.db", check_same_thread=False)

salt = secrets.token_hex(16)
password_hash = hashlib.pbkdf2_hmac(
    "sha256", NUEVA_CLAVE.encode("utf-8"), salt.encode("utf-8"), 200_000
).hex()

conn = conectar()
c = conn.cursor()
c.execute("""
    UPDATE usuarios 
    SET password_hash = ?, salt = ? 
    WHERE LOWER(usuario) = LOWER(?)
""", (password_hash, salt, USUARIO))
conn.commit()

if c.rowcount > 0:
    print(f"¡Contraseña actualizada con éxito para el usuario '{USUARIO}'!")
else:
    print(f"No se encontró el usuario '{USUARIO}'.")
conn.close()
