import sqlite3
from pathlib import Path

# Ruta de la base de datos
DB_PATH = Path("data") / "sistema.db"

# Crear carpeta data si no existe
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

# Conexión
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

# Activar claves foráneas
cur.execute("PRAGMA foreign_keys = ON;")

# =========================
# TABLA USUARIOS
# =========================
cur.execute("""
CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre_completo TEXT NOT NULL,
    correo TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    rol TEXT NOT NULL CHECK (rol IN ('COORDINADOR', 'ASISTENTE')),
    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0,1)),
    fecha_creacion TEXT NOT NULL,
    ultimo_acceso TEXT
);
""")

# =========================
# TABLA SOLICITUDES
# =========================
cur.execute("""
CREATE TABLE IF NOT EXISTS solicitudes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,

    -- Datos del formulario
    marca_temporal TEXT NOT NULL,
    correo TEXT NOT NULL,
    tipo_documento TEXT NOT NULL CHECK (tipo_documento IN ('CERT','CONST')),
    nombres TEXT NOT NULL,
    apellidos TEXT NOT NULL,
    documento TEXT NOT NULL,
    fecha_inicio TEXT NOT NULL,
    fecha_fin TEXT NOT NULL,
    universidad TEXT NOT NULL,
    codigo_alumno TEXT NOT NULL,
    facultad TEXT NOT NULL,
    carrera TEXT NOT NULL,
    ciclo TEXT,

    -- Datos internos del sistema
    horas_totales INTEGER,
    estado TEXT NOT NULL DEFAULT 'RECIBIDO'
        CHECK (estado IN ('RECIBIDO','PENDIENTE','OBSERVADO','REVISADO','EMITIDO','ANULADO')),
    observaciones TEXT,
    codigo_documento TEXT,
    revisado_por INTEGER,
    fecha_revision TEXT,
    emitido_por INTEGER,
    fecha_emision TEXT,
    ruta_pdf TEXT,

    UNIQUE (marca_temporal, documento, tipo_documento),

    FOREIGN KEY (revisado_por) REFERENCES usuarios(id),
    FOREIGN KEY (emitido_por) REFERENCES usuarios(id)
);
""")

# =========================
# TABLA HISTORIAL
# =========================
cur.execute("""
CREATE TABLE IF NOT EXISTS historial (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    solicitud_id INTEGER NOT NULL,
    fecha TEXT NOT NULL,
    usuario_id INTEGER NOT NULL,
    accion TEXT NOT NULL,
    detalle TEXT,

    FOREIGN KEY (solicitud_id) REFERENCES solicitudes(id) ON DELETE CASCADE,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);
""")

conn.commit()
conn.close()

print("Base de datos sistema.db creada correctamente en la carpeta data.")