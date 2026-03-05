from flask import Flask, render_template, request, redirect, url_for, flash, session
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3
from pathlib import Path
from datetime import datetime
import pandas as pd
import os
from werkzeug.utils import secure_filename
from flask import send_file, abort
from docx import Document
import comtypes.client
import uuid
import pythoncom
import time


app = Flask(__name__)
app.secret_key = "CAMBIA_ESTA_CLAVE"

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "sistema.db"
ARCHIVOS_DIR = BASE_DIR / "archivos"
PLANTILLAS_DIR = ARCHIVOS_DIR / "plantillas"
PLANT_CERT_DIR = PLANTILLAS_DIR / "certificados"
PLANT_CONST_DIR = PLANTILLAS_DIR / "constancias"

ALLOWED_TEMPLATE_EXT = {".docx"}

# URL CSV publicada
SHEETS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQG-icmEF9C0D4GOAig3dBVXCl0fN0qXS-6-o1yVM8KY-6MkpbHDF22CizH7elirxudmA5Anh-wSM0C/pub?output=csv"


def db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def login_required(view_func):
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            flash("Inicia sesión para continuar", "info")
            return redirect(url_for("login"))
        return view_func(*args, **kwargs)
    wrapper.__name__ = view_func.__name__
    return wrapper


def usuario_actual():
    if not session.get("user_id"):
        return None
    conn = db()
    u = conn.execute("SELECT * FROM usuarios WHERE id = ?", (session["user_id"],)).fetchone()
    conn.close()
    return u


def s(v):
    if v is None:
        return ""
    return str(v).strip()

def upper(v):
    return s(v).upper()

def parse_form_datetime_to_iso(value):
    # Espera "d/m/yyyy HH:MM:SS" o "dd/mm/yyyy HH:MM:SS"
    txt = s(value)
    if not txt:
        return ""
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
        try:
            dt = datetime.strptime(txt, fmt)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    return ""


def ensure_solicitudes_schema(conn):
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS solicitudes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sheet_uid TEXT UNIQUE,
        marca_temporal TEXT,
        marca_dt TEXT,
        correo_solicitante TEXT,
        correo TEXT,
        tipo_documento TEXT,
        nombres TEXT,
        apellidos TEXT,
        documento TEXT,
        fecha_inicio TEXT,
        fecha_fin TEXT,
        universidad TEXT,
        codigo_alumno TEXT,
        facultad TEXT,
        carrera TEXT,
        ciclo TEXT,
        cargo TEXT,
        actividades TEXT,
        horas_totales INTEGER,
        estado TEXT,
        observaciones TEXT,
        fecha_revision TEXT,
        revisado_por TEXT,
        fecha_emision TEXT,
        emitido_por TEXT,
        codigo_documento TEXT,
        ruta_pdf TEXT,
        creado_en TEXT,
        actualizado_en TEXT
    )
    """)

    # 2 Detecta columnas existentes
    existentes = {r["name"] for r in conn.execute("PRAGMA table_info(solicitudes)").fetchall()}

    # 3 Agrega columnas faltantes de manera segura
    # Nota importante, en SQLite no puedes añadir NOT NULL sin DEFAULT
    columnas = {
        "sheet_uid": "TEXT",
        "marca_temporal": "TEXT",
        "marca_dt": "TEXT",
        "correo_solicitante": "TEXT",
        "correo": "TEXT",
        "tipo_documento": "TEXT",
        "nombres": "TEXT",
        "apellidos": "TEXT",
        "documento": "TEXT",
        "fecha_inicio": "TEXT",
        "fecha_fin": "TEXT",
        "universidad": "TEXT",
        "codigo_alumno": "TEXT",
        "facultad": "TEXT",
        "carrera": "TEXT",
        "ciclo": "TEXT",
        "cargo": "TEXT",
        "actividades": "TEXT",
        "horas_totales": "INTEGER",
        "estado": "TEXT",
        "observaciones": "TEXT",
        "fecha_revision": "TEXT",
        "revisado_por": "TEXT",
        "fecha_emision": "TEXT",
        "emitido_por": "TEXT",
        "codigo_documento": "TEXT",
        "ruta_pdf": "TEXT",
        "creado_en": "TEXT",
        "actualizado_en": "TEXT",
    }

    for col, tipo in columnas.items():
        if col not in existentes:
            cur.execute(f"ALTER TABLE solicitudes ADD COLUMN {col} {tipo}")

    # 4 Backfill para que no te vuelvan a salir errores por valores vacíos
    # Copia correo_solicitante a correo si correo está vacío
    cur.execute("""
        UPDATE solicitudes
        SET correo = correo_solicitante
        WHERE (correo IS NULL OR TRIM(correo) = '')
          AND (correo_solicitante IS NOT NULL AND TRIM(correo_solicitante) <> '')
    """)

    # Si actualizado_en está vacío, lo igualamos a creado_en
    cur.execute("""
        UPDATE solicitudes
        SET actualizado_en = creado_en
        WHERE (actualizado_en IS NULL OR TRIM(actualizado_en) = '')
          AND (creado_en IS NOT NULL AND TRIM(creado_en) <> '')
    """)

    # 5 Índices
    cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_solicitudes_sheet_uid ON solicitudes(sheet_uid)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_solicitudes_estado ON solicitudes(estado)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_solicitudes_documento ON solicitudes(documento)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_solicitudes_marca_dt ON solicitudes(marca_dt)")

    conn.commit()

def ensure_historial_schema(conn):
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS historial_solicitud (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        solicitud_id INTEGER NOT NULL,
        ficha TEXT NOT NULL,
        usuario TEXT NOT NULL,
        accion TEXT NOT NULL,
        detalle TEXT,
        creado_en TEXT NOT NULL,
        FOREIGN KEY (solicitud_id) REFERENCES solicitudes(id) ON DELETE CASCADE
    )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_hist_sol_id ON historial_solicitud(solicitud_id)")
    conn.commit()

def ensure_usuarios_schema(conn):
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre_completo TEXT NOT NULL,
        correo TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        rol TEXT NOT NULL,
        activo INTEGER NOT NULL DEFAULT 1,
        fecha_creacion TEXT NOT NULL,
        ultimo_acceso TEXT
    )
    """)
    cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_usuarios_correo ON usuarios(correo)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_usuarios_rol ON usuarios(rol)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_usuarios_activo ON usuarios(activo)")
    conn.commit()

def ensure_config_schema(conn):
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS configuracion (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        ruta_salida TEXT NOT NULL DEFAULT 'archivos/emitidos/',
        correo_emisor TEXT NOT NULL DEFAULT '',
        envio_correo INTEGER NOT NULL DEFAULT 0,
        actualizado_en TEXT NOT NULL
    )
    """)

    row = cur.execute("SELECT id FROM configuracion WHERE id = 1").fetchone()
    if not row:
        cur.execute("""
            INSERT INTO configuracion (id, ruta_salida, correo_emisor, envio_correo, actualizado_en)
            VALUES (1, 'archivos/emitidos/', '', 0, ?)
        """, (ahora(),))
    conn.commit()


def get_config(conn):
    ensure_config_schema(conn)
    return conn.execute("SELECT * FROM configuracion WHERE id = 1").fetchone()


@app.get("/")
def root():
    if session.get("user_id"):
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.get("/login")
def login():
    return render_template("login.html")


@app.post("/iniciar-sesion")
def iniciar_sesion():
    correo = (request.form.get("correo") or "").strip().lower()
    contrasena = (request.form.get("contrasena") or "").strip()

    if not correo or not contrasena:
        flash("Completa correo y contraseña", "error")
        return redirect(url_for("login"))

    conn = db()
    u = conn.execute("SELECT * FROM usuarios WHERE correo = ?", (correo,)).fetchone()

    if not u:
        conn.close()
        flash("Credenciales incorrectas", "error")
        return redirect(url_for("login"))

    if int(u["activo"]) != 1:
        conn.close()
        flash("Usuario inactivo. Contacta al coordinador", "error")
        return redirect(url_for("login"))

    if not check_password_hash(u["password_hash"], contrasena):
        conn.close()
        flash("Credenciales incorrectas", "error")
        return redirect(url_for("login"))

    conn.execute("UPDATE usuarios SET ultimo_acceso = ? WHERE id = ?", (ahora(), u["id"]))
    conn.commit()
    conn.close()

    session["user_id"] = u["id"]
    session["rol"] = u["rol"]
    session["nombre"] = u["nombre_completo"]

    flash("Sesión iniciada", "success")
    return redirect(url_for("dashboard"))


@app.get("/registro")
def registro():
    return render_template("registro.html")


@app.post("/registrarse")
def registrarse():
    nombre = (request.form.get("nombre") or "").strip()
    correo = (request.form.get("correo") or "").strip().lower()
    contrasena = (request.form.get("contrasena") or "").strip()
    confirmar = (request.form.get("confirmar") or "").strip()

    if not nombre or not correo or not contrasena or not confirmar:
        flash("Completa todos los campos", "error")
        return redirect(url_for("registro"))

    if contrasena != confirmar:
        flash("Las contraseñas no coinciden", "error")
        return redirect(url_for("registro"))

    conn = db()

    existe = conn.execute("SELECT 1 FROM usuarios WHERE correo = ?", (correo,)).fetchone()
    if existe:
        conn.close()
        flash("Ese correo ya está registrado", "error")
        return redirect(url_for("registro"))

    total = conn.execute("SELECT COUNT(*) AS n FROM usuarios").fetchone()["n"]
    rol = "COORDINADOR" if total == 0 else "ASISTENTE"

    password_hash = generate_password_hash(contrasena)

    conn.execute("""
        INSERT INTO usuarios (nombre_completo, correo, password_hash, rol, activo, fecha_creacion)
        VALUES (?, ?, ?, ?, 1, ?)
    """, (nombre, correo, password_hash, rol, ahora()))
    conn.commit()
    conn.close()

    flash(f"Registro exitoso. Rol asignado {rol}", "success")
    return redirect(url_for("login"))


@app.get("/logout")
def logout():
    session.clear()
    flash("Sesión cerrada", "info")
    return redirect(url_for("login"))


@app.get("/dashboard")
@login_required
def dashboard():
    conn = db()

    # por si aún no existen
    ensure_solicitudes_schema(conn)
    ensure_historial_schema(conn)

    def count_estado(e):
        r = conn.execute("SELECT COUNT(*) AS n FROM solicitudes WHERE estado = ?", (e,)).fetchone()
        return int(r["n"]) if r else 0

    kpi_recibido  = count_estado("RECIBIDO")
    kpi_pendiente = count_estado("PENDIENTE")
    kpi_observado = count_estado("OBSERVADO")
    kpi_revisado  = count_estado("REVISADO")
    kpi_emitido   = count_estado("EMITIDO")
    kpi_anulado   = count_estado("ANULADO")

    # últimas 10 operaciones desde el historial
    ultimas10 = conn.execute("""
        SELECT
            h.id AS id,
            s.tipo_documento AS tipo_documento,
            s.documento AS documento,
            (s.nombres || ' ' || s.apellidos) AS nombre_completo,
            s.estado AS estado,
            h.usuario AS usuario,
            h.creado_en AS fecha
        FROM historial_solicitud h
        JOIN solicitudes s ON s.id = h.solicitud_id
        ORDER BY h.id DESC
        LIMIT 10
    """).fetchall()

    alert_observadas = kpi_observado
    alert_plantillas = 0
    alert_correos_fallidos = 0

    conn.close()

    return render_template(
        "dashboard.html",
        active="dashboard",
        kpi_recibido=kpi_recibido,
        kpi_pendiente=kpi_pendiente,
        kpi_observado=kpi_observado,
        kpi_revisado=kpi_revisado,
        kpi_emitido=kpi_emitido,
        kpi_anulado=kpi_anulado,
        ultimas10=ultimas10,
        alert_observadas=alert_observadas,
        alert_plantillas=alert_plantillas,
        alert_correos_fallidos=alert_correos_fallidos,
    )

def add_historial(conn, solicitud_id, ficha, usuario, accion, detalle=None):
    conn.execute("""
        INSERT INTO historial_solicitud (solicitud_id, ficha, usuario, accion, detalle, creado_en)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (solicitud_id, ficha, usuario, accion, detalle, ahora()))


@app.get("/solicitudes")
@login_required
def solicitudes():
    estado = (request.args.get("estado") or "").strip().upper()
    tipo = (request.args.get("tipo") or "").strip().upper()
    dni = (request.args.get("dni") or "").strip()
    desde = (request.args.get("desde") or "").strip()  # YYYY-MM-DD
    hasta = (request.args.get("hasta") or "").strip()  # YYYY-MM-DD

    conn = db()
    ensure_solicitudes_schema(conn)
    ensure_historial_schema(conn)

    where = []
    params = []

    if estado:
        where.append("estado = ?")
        params.append(estado)

    if tipo:
        where.append("tipo_documento = ?")
        params.append(tipo)

    if dni:
        where.append("documento LIKE ?")
        params.append(f"%{dni}%")

    # filtro por fecha usando marca_dt
    # desde y hasta vienen como YYYY-MM-DD
    if desde:
        where.append("(marca_dt >= ? OR (marca_dt IS NULL AND creado_en >= ?))")
        params.append(desde + " 00:00:00")
        params.append(desde + " 00:00:00")

    if hasta:
        where.append("(marca_dt <= ? OR (marca_dt IS NULL AND creado_en <= ?))")
        params.append(hasta + " 23:59:59")
        params.append(hasta + " 23:59:59")

    sql = "SELECT * FROM solicitudes"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY id DESC LIMIT 500"

    rows = conn.execute(sql, params).fetchall()
    total = conn.execute("SELECT COUNT(*) AS n FROM solicitudes").fetchone()["n"]

    conn.close()

    return render_template(
        "solicitudes.html",
        active="solicitudes",
        solicitudes=rows,
        total=total,
        estados=["RECIBIDO","PENDIENTE","OBSERVADO","REVISADO","EMITIDO","ANULADO"],
        filtros={"estado": estado, "tipo": tipo, "dni": dni, "desde": desde, "hasta": hasta}
    )

@app.get("/solicitudes/<int:sid>")
@login_required
def solicitudes_detalle(sid):
    conn = db()
    ensure_solicitudes_schema(conn)
    ensure_historial_schema(conn)

    s = get_solicitud_por_id(conn, sid)
    if not s:
        conn.close()
        flash("Solicitud no encontrada", "error")
        return redirect(url_for("solicitudes"))

    historial = get_historial(conn, sid)
    conn.close()

    return render_template(
        "detalle.html",
        active="solicitudes",
        s=s,
        historial=historial
    )

def get_solicitud_por_id(conn, sid):
    return conn.execute("SELECT * FROM solicitudes WHERE id = ?", (sid,)).fetchone()

def get_historial(conn, sid):
    return conn.execute("""
        SELECT * FROM historial_solicitud
        WHERE solicitud_id = ?
        ORDER BY id DESC
    """, (sid,)).fetchall()

@app.post("/solicitudes/<int:sid>/guardar")
@login_required
def solicitudes_guardar(sid):
    horas = (request.form.get("horas_totales") or "").strip()
    observ = (request.form.get("observaciones") or "").strip()

    conn = db()
    ensure_solicitudes_schema(conn)
    ensure_historial_schema(conn)

    s = get_solicitud_por_id(conn, sid)
    if not s:
        conn.close()
        flash("Solicitud no encontrada", "error")
        return redirect(url_for("solicitudes"))

    cambios = []
    if (s["horas_totales"] or "") != horas:
        cambios.append(f"Horas {s['horas_totales'] or ''} -> {horas}")
    if (s["observaciones"] or "") != observ:
        cambios.append("Observaciones actualizadas")

    conn.execute("""
        UPDATE solicitudes
        SET horas_totales = ?, observaciones = ?, actualizado_en = ?
        WHERE id = ?
    """, (horas, observ, ahora(), sid))

    if cambios:
        add_historial(
            conn,
            sid,
            s["marca_temporal"] or "",
            session.get("nombre", "USUARIO"),
            "GUARDAR",
            " | ".join(cambios)
        )

    conn.commit()
    conn.close()

    flash("Cambios guardados", "success")
    return redirect(url_for("solicitudes_detalle", sid=sid))

@app.post("/solicitudes/<int:sid>/estado/observado")
@login_required
def solicitudes_marcar_observado(sid):
    conn = db()
    ensure_solicitudes_schema(conn)
    ensure_historial_schema(conn)

    s = get_solicitud_por_id(conn, sid)
    if not s:
        conn.close()
        flash("Solicitud no encontrada", "error")
        return redirect(url_for("solicitudes"))

    conn.execute("""
        UPDATE solicitudes
        SET estado = 'OBSERVADO',
            actualizado_en = ?
        WHERE id = ?
    """, (ahora(), sid))

    add_historial(conn, sid, s["marca_temporal"] or "", session.get("nombre","USUARIO"), "ESTADO", "Marcado como OBSERVADO")
    conn.commit()
    conn.close()

    flash("Estado actualizado a OBSERVADO", "success")
    return redirect(url_for("solicitudes_detalle", sid=sid))

@app.post("/solicitudes/<int:sid>/estado/revisado")
@login_required
def solicitudes_marcar_revisado(sid):
    conn = db()
    ensure_solicitudes_schema(conn)
    ensure_historial_schema(conn)

    s = get_solicitud_por_id(conn, sid)
    if not s:
        conn.close()
        flash("Solicitud no encontrada", "error")
        return redirect(url_for("solicitudes"))

    conn.execute("""
        UPDATE solicitudes
        SET estado = 'REVISADO',
            revisado_por = ?,
            fecha_revision = ?,
            actualizado_en = ?
        WHERE id = ?
    """, (session.get("nombre","USUARIO"), ahora(), ahora(), sid))

    add_historial(conn, sid, s["marca_temporal"] or "", session.get("nombre","USUARIO"), "ESTADO", "Marcado como REVISADO")
    conn.commit()
    conn.close()

    flash("Estado actualizado a REVISADO", "success")
    return redirect(url_for("solicitudes_detalle", sid=sid))

@app.post("/solicitudes/<int:sid>/emitir")
@login_required
def solicitudes_emitir(sid):
    conn = db()
    ensure_solicitudes_schema(conn)
    ensure_historial_schema(conn)
    config = get_config(conn) 

    s = get_solicitud_por_id(conn, sid)
    if not s:
        conn.close()
        flash("Solicitud no encontrada", "error")
        return redirect(url_for("solicitudes"))

    horas = str(s["horas_totales"] or "").strip()
    if not horas:
        conn.close()
        flash("Antes de emitir, registra las horas totales", "error")
        return redirect(url_for("solicitudes_detalle", sid=sid))

    tipo_doc = s["tipo_documento"]
    carrera = s["carrera"]
    plantilla = conn.execute("""
        SELECT ruta_docx FROM plantillas 
        WHERE tipo_documento = ? AND carrera = ? AND activo = 1
    """, (tipo_doc, carrera)).fetchone()

    if not plantilla:
        conn.close()
        flash(f"No hay plantilla activa para {tipo_doc} de {carrera}", "error")
        return redirect(url_for("solicitudes_detalle", sid=sid))

    ruta_plantilla_abs = (BASE_DIR / plantilla["ruta_docx"]).resolve()
    
    try:
        doc = Document(ruta_plantilla_abs)
        
        
        reemplazos = {
            "{{NOMBRE_COMPLETO}}": f"{s['nombres']} {s['apellidos']}",
            "{{DNI}}": str(s["documento"] or ""),
            "{{UNIVERSIDAD}}": str(s["universidad"] or ""),
            "{{FACULTAD}}": str(s["facultad"] or ""),
            "{{CARRERA}}": str(s["carrera"] or ""),
            "{{CODIGO}}": str(s["codigo_alumno"] or ""),
            "{{FECHA_INICIO}}": formatear_fecha_latam(s["fecha_inicio"]), 
            "{{FECHA_FIN}}": formatear_fecha_latam(s["fecha_fin"]),
            "{{CARGO}}": str(s["cargo"] or "ASISTENTE").upper(),
            "{{HORAS_TOTALES}}": str(horas),
            "{{ACTIVIDADES}}": str(s["actividades"] or ""),
            "{{FECHA_EMISION}}": datetime.now().strftime("%d/%m/%Y")
        }

       
        def reemplazar_texto_en_parrafo(parrafo, reemplazos):
            texto_parrafo = parrafo.text
            for llave, valor in reemplazos.items():
                if llave in texto_parrafo:
                    texto_parrafo = texto_parrafo.replace(llave, valor)
            
           
            if parrafo.text != texto_parrafo:
               
                parrafo.clear()
                parrafo.add_run(texto_parrafo)

       
        for parrafo in doc.paragraphs:
            reemplazar_texto_en_parrafo(parrafo, reemplazos)
                
        
        for tabla in doc.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    for parrafo in celda.paragraphs:
                        reemplazar_texto_en_parrafo(parrafo, reemplazos)
        
        nombre_archivo = f"{tipo_doc}_{s['documento']}_{uuid.uuid4().hex[:6]}.docx"
        ruta_salida_dir = BASE_DIR / config["ruta_salida"]
        ruta_salida_dir.mkdir(parents=True, exist_ok=True)
        
        ruta_archivo_final = ruta_salida_dir / nombre_archivo
        doc.save(ruta_archivo_final)
        
        ruta_bd = str(ruta_archivo_final.relative_to(BASE_DIR)).replace("\\", "/")
        
    except Exception as e:
        conn.close()
        flash(f"Error al generar el documento: {e}", "error")
        return redirect(url_for("solicitudes_detalle", sid=sid))
        
    codigo_doc = f"{tipo_doc}-{s['documento']}-{datetime.now().strftime('%y%m%d')}"

    conn.execute("""
        UPDATE solicitudes
        SET estado = 'EMITIDO',
            emitido_por = ?,
            fecha_emision = ?,
            actualizado_en = ?,
            codigo_documento = ?,
            ruta_pdf = ?
        WHERE id = ?
    """, (session.get("nombre","USUARIO"), ahora(), ahora(), codigo_doc, ruta_bd, sid))

    add_historial(conn, sid, s["marca_temporal"] or "", session.get("nombre","USUARIO"), "EMISION", "Documento generado y EMITIDO")
    conn.commit()
    conn.close()

    flash("Solicitud marcada como EMITIDO y documento generado", "success")
    return redirect(url_for("solicitudes_detalle", sid=sid))

@app.post("/solicitudes/<int:sid>/anular")
@login_required
def solicitudes_anular(sid):
    conn = db()
    ensure_solicitudes_schema(conn)
    ensure_historial_schema(conn)

    s = get_solicitud_por_id(conn, sid)
    if not s:
        conn.close()
        flash("Solicitud no encontrada", "error")
        return redirect(url_for("solicitudes"))

    conn.execute("""
        UPDATE solicitudes
        SET estado = 'ANULADO',
            actualizado_en = ?
        WHERE id = ?
    """, (ahora(), sid))

    add_historial(conn, sid, s["marca_temporal"] or "", session.get("nombre","USUARIO"), "ESTADO", "Solicitud ANULADA")
    conn.commit()
    conn.close()

    flash("Solicitud anulada", "success")
    return redirect(url_for("solicitudes_detalle", sid=sid))

def formatear_fecha_latam(valor):
    txt = s(valor)
    if not txt:
        return ""
    
   
    formatos = ["%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y", "%Y-%m-%d"]
    
    for fmt in formatos:
        try:
            dt = datetime.strptime(txt, fmt)
            return dt.strftime("%d/%m/%Y") 
        except ValueError:
            continue
            
    return txt 



@app.post("/solicitudes/sincronizar")
@login_required
def solicitudes_sincronizar():
    try:
        url_con_cache_breaker = f"{SHEETS_CSV_URL}&cache_buster={int(time.time())}"
        df = pd.read_csv(url_con_cache_breaker)
    
        df.columns = df.columns.str.strip()
        
        
    except Exception as e:
        flash(f"Error al conectar con Google Sheets: {e}", "error")
        return redirect(url_for("solicitudes"))

    conn = db()
    ensure_solicitudes_schema(conn)
    cur = conn.cursor()

    nuevos = duplicados = errores = 0

    for _, row in df.iterrows():

        marca         = s(row.get("Marca temporal"))
        tipo_raw      = upper(row.get("Seleccione lo que desea solicitar"))
        nombres       = upper(row.get("NOMBRES"))
        apellidos     = upper(row.get("APELLIDOS"))
        documento     = s(row.get("N° DOCUMENTO"))
        f_inicio      = formatear_fecha_latam(row.get("Fecha de inicio (dd/mm/yyyy)"))
        f_fin         = formatear_fecha_latam(row.get("Fecha de fin (dd/mm/yyyy)"))
        uni           = upper(row.get("NOMBRE DE LA UNIVERSIDAD O INSTITUTO"))
        cod_alumno    = s(row.get("CODIGO DE ALUMNO"))
        facultad      = upper(row.get("FACULTAD"))
        carrera       = upper(row.get("CARRERA"))
        ciclo         = s(row.get("CICLO"))
        cargo         = upper(row.get("CARGO"))
        actividades   = s(row.get("ACTIVIDADES"))
        
        correo = s(row.get("CORREO ELECTRONICO") or row.get("CORREO ELECTRÓNICO") or row.get("Dirección de correo electrónico") or "")

        if not documento or not marca:
            continue

        tipo = "CONST" if "CONSTANCIA" in tipo_raw else "CERT"
        sheet_uid = f"{marca}|{documento}|{tipo}"

        try:
            cur.execute("""
                INSERT INTO solicitudes (
                    sheet_uid, marca_temporal, correo, correo_solicitante, tipo_documento,
                    nombres, apellidos, documento,
                    fecha_inicio, fecha_fin,
                    universidad, codigo_alumno, facultad, carrera, ciclo,cargo,
                    actividades, estado, creado_en, actualizado_en
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'RECIBIDO', ?, ?)
            """, (
                sheet_uid, marca, correo, correo, tipo,
                nombres, apellidos, documento,
                f_inicio, f_fin,
                uni, cod_alumno, facultad, carrera, ciclo,cargo,
                actividades, ahora(), ahora()
            ))
            nuevos += 1
        except sqlite3.IntegrityError:
            duplicados += 1
        except Exception as e:
            print(f"ERROR EN FILA {documento}: {e}") 
            errores += 1

    conn.commit()
    conn.close()

    if errores > 0:
        flash(f"Sincronización: {nuevos} nuevos. Hubo {errores} errores.", "warning")
    else:
        flash(f"Sincronización exitosa: {nuevos} nuevos, {duplicados} duplicados.", "success")
        
    return redirect(url_for("solicitudes"))

def ensure_plantillas_schema(conn):
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS plantillas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo_documento TEXT NOT NULL CHECK (tipo_documento IN ('CERT','CONST')),
        carrera TEXT NOT NULL,
        archivo_nombre TEXT NOT NULL,
        ruta_docx TEXT NOT NULL,
        activo INTEGER NOT NULL DEFAULT 1,
        creado_en TEXT NOT NULL,
        actualizado_en TEXT NOT NULL
    )
    """)

    cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_plantillas_tipo_carrera ON plantillas(tipo_documento, carrera)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_plantillas_activo ON plantillas(activo)")
    conn.commit()

@app.get("/plantillas")
@login_required
def plantillas():
    conn = db()
    ensure_plantillas_schema(conn)

    rows = conn.execute("""
        SELECT *
        FROM plantillas
        ORDER BY actualizado_en DESC, id DESC
    """).fetchall()

    total = conn.execute("SELECT COUNT(*) AS n FROM plantillas").fetchone()["n"]
    conn.close()

    return render_template(
        "plantillas.html",
        active="plantillas",
        plantillas=rows,
        total_plantillas=total
    )

@app.post("/plantillas/subir")
@login_required
def plantillas_subir():
    tipo = (request.form.get("tipo_documento") or request.form.get("tipo") or "").strip().upper()
    carrera = (request.form.get("carrera") or "").strip().upper()
    f = request.files.get("archivo")

    if tipo not in {"CERT", "CONST"}:
        flash("Tipo inválido", "error")
        return redirect(url_for("plantillas"))

    if not carrera:
        flash("Carrera es obligatoria", "error")
        return redirect(url_for("plantillas"))

    if not f or not f.filename:
        flash("Debes seleccionar un archivo docx", "error")
        return redirect(url_for("plantillas"))

    filename = secure_filename(f.filename)
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED_TEMPLATE_EXT:
        flash("Solo se permite .docx", "error")
        return redirect(url_for("plantillas"))

    # carpeta destino según tipo
    PLANT_CERT_DIR.mkdir(parents=True, exist_ok=True)
    PLANT_CONST_DIR.mkdir(parents=True, exist_ok=True)

    dest_dir = PLANT_CERT_DIR if tipo == "CERT" else PLANT_CONST_DIR

    # nombre final del archivo
    # ejemplo CERT_INGENIERIA_DE_SISTEMAS.docx
    safe_carrera = "_".join(carrera.split())
    final_name = f"{tipo}_{safe_carrera}.docx"
    dest_path = dest_dir / final_name

    # guardar archivo
    try:
        f.save(dest_path)
    except Exception as e:
        flash(f"No se pudo guardar el archivo. Detalle {e}", "error")
        return redirect(url_for("plantillas"))

    # guardar en BD con upsert
    conn = db()
    ensure_plantillas_schema(conn)

    ahora_txt = ahora()
    ruta_rel = str(dest_path.relative_to(BASE_DIR)).replace("\\", "/")

    existente = conn.execute("""
        SELECT id FROM plantillas
        WHERE tipo_documento = ? AND carrera = ?
    """, (tipo, carrera)).fetchone()

    if existente:
        conn.execute("""
            UPDATE plantillas
            SET archivo_nombre = ?,
                ruta_docx = ?,
                activo = 1,
                actualizado_en = ?
            WHERE id = ?
        """, (final_name, ruta_rel, ahora_txt, existente["id"]))
        conn.commit()
        conn.close()
        flash("Plantilla actualizada y activada", "success")
        return redirect(url_for("plantillas"))

    conn.execute("""
        INSERT INTO plantillas (tipo_documento, carrera, archivo_nombre, ruta_docx, activo, creado_en, actualizado_en)
        VALUES (?, ?, ?, ?, 1, ?, ?)
    """, (tipo, carrera, final_name, ruta_rel, ahora_txt, ahora_txt))
    conn.commit()
    conn.close()

    flash("Plantilla subida correctamente", "success")
    return redirect(url_for("plantillas"))

@app.get("/usuarios")
@login_required
def usuarios():
    q = (request.args.get("q") or "").strip().lower()
    rol = (request.args.get("rol") or "").strip().upper()
    activo = (request.args.get("activo") or "").strip()  # "1" "0" o ""

    conn = db()
    ensure_usuarios_schema(conn)

    where = []
    params = []

    if q:
        where.append("(LOWER(nombre_completo) LIKE ? OR LOWER(correo) LIKE ?)")
        params.extend([f"%{q}%", f"%{q}%"])

    if rol:
        where.append("rol = ?")
        params.append(rol)

    if activo in ("0", "1"):
        where.append("activo = ?")
        params.append(int(activo))

    sql = "SELECT id, nombre_completo, correo, rol, activo, ultimo_acceso, fecha_creacion FROM usuarios"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY id DESC LIMIT 500"

    rows = conn.execute(sql, params).fetchall()
    conn.close()

    return render_template(
        "usuarios.html",
        active="usuarios",
        usuarios=rows,
        filtros={"q": q, "rol": rol, "activo": activo},
        roles=["COORDINADOR", "ASISTENTE"],
    )

@app.post("/usuarios/crear")
@login_required
def usuarios_crear():
    if (session.get("rol") or "").upper() != "COORDINADOR":
        flash("No tienes permiso para crear usuarios", "error")
        return redirect(url_for("usuarios"))

    nombre = (request.form.get("nombre") or "").strip()
    correo = (request.form.get("correo") or "").strip().lower()
    rol = (request.form.get("rol") or "").strip().upper()
    contrasena = (request.form.get("contrasena") or "").strip()

    if not nombre or not correo or not rol or not contrasena:
        flash("Completa todos los campos", "error")
        return redirect(url_for("usuarios"))

    if rol not in ("COORDINADOR", "ASISTENTE"):
        flash("Rol inválido", "error")
        return redirect(url_for("usuarios"))

    conn = db()
    ensure_usuarios_schema(conn)

    existe = conn.execute("SELECT 1 FROM usuarios WHERE correo = ?", (correo,)).fetchone()
    if existe:
        conn.close()
        flash("Ese correo ya está registrado", "error")
        return redirect(url_for("usuarios"))

    password_hash = generate_password_hash(contrasena)

    conn.execute("""
        INSERT INTO usuarios (nombre_completo, correo, password_hash, rol, activo, fecha_creacion, ultimo_acceso)
        VALUES (?, ?, ?, ?, 1, ?, NULL)
    """, (nombre, correo, password_hash, rol, ahora()))
    conn.commit()
    conn.close()

    flash("Usuario creado", "success")
    return redirect(url_for("usuarios"))

@app.post("/usuarios/<int:uid>/toggle")
@login_required
def usuarios_toggle(uid):
    if (session.get("rol") or "").upper() != "COORDINADOR":
        flash("No tienes permiso", "error")
        return redirect(url_for("usuarios"))

    if session.get("user_id") == uid:
        flash("No puedes modificar tu propio usuario", "error")
        return redirect(url_for("usuarios"))

    conn = db()
    ensure_usuarios_schema(conn)

    u = conn.execute("SELECT id, activo FROM usuarios WHERE id = ?", (uid,)).fetchone()
    if not u:
        conn.close()
        flash("Usuario no encontrado", "error")
        return redirect(url_for("usuarios"))

    nuevo = 0 if int(u["activo"]) == 1 else 1
    conn.execute("UPDATE usuarios SET activo = ? WHERE id = ?", (nuevo, uid))
    conn.commit()
    conn.close()

    flash("Estado actualizado", "success")
    return redirect(url_for("usuarios"))

@app.post("/usuarios/<int:user_id>/cambiar-rol")
@login_required
def usuarios_cambiar_rol(user_id):
    if (session.get("rol") or "").upper() != "COORDINADOR":
        flash("No tienes permiso", "error")
        return redirect(url_for("usuarios"))

    nuevo_rol = (request.form.get("rol") or "").strip().upper()
    if nuevo_rol not in ("COORDINADOR", "ASISTENTE"):
        flash("Rol inválido", "error")
        return redirect(url_for("usuarios"))

    conn = db()
    ensure_usuarios_schema(conn)

    u = conn.execute("SELECT id FROM usuarios WHERE id = ?", (user_id,)).fetchone()
    if not u:
        conn.close()
        flash("Usuario no encontrado", "error")
        return redirect(url_for("usuarios"))

    conn.execute("UPDATE usuarios SET rol = ? WHERE id = ?", (nuevo_rol, user_id))
    conn.commit()
    conn.close()

    flash("Rol actualizado", "success")
    return redirect(url_for("usuarios"))

@app.post("/usuarios/<int:user_id>/eliminar")
@login_required
def usuarios_eliminar(user_id):
    if (session.get("rol") or "").upper() != "COORDINADOR":
        flash("No tienes permiso", "error")
        return redirect(url_for("usuarios"))

    if session.get("user_id") == user_id:
        flash("No puedes eliminarte a ti mismo", "error")
        return redirect(url_for("usuarios"))

    conn = db()
    ensure_usuarios_schema(conn)

    conn.execute("DELETE FROM usuarios WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()

    flash("Usuario eliminado", "success")
    return redirect(url_for("usuarios"))

@app.post("/plantillas/<int:pid>/toggle")
@login_required
def plantillas_toggle(pid):
    conn = db()
    ensure_plantillas_schema(conn)

    row = conn.execute("SELECT id, activo FROM plantillas WHERE id = ?", (pid,)).fetchone()
    if not row:
        conn.close()
        flash("Plantilla no encontrada", "error")
        return redirect(url_for("plantillas"))

    nuevo = 0 if int(row["activo"]) == 1 else 1

    conn.execute("""
        UPDATE plantillas
        SET activo = ?, actualizado_en = ?
        WHERE id = ?
    """, (nuevo, ahora(), pid))
    conn.commit()
    conn.close()

    flash("Plantilla actualizada", "success")
    return redirect(url_for("plantillas"))

@app.get("/plantillas/<int:pid>/descargar")
@login_required
def plantillas_descargar(pid):
    conn = db()
    ensure_plantillas_schema(conn)

    p = conn.execute("SELECT * FROM plantillas WHERE id = ?", (pid,)).fetchone()
    conn.close()

    if not p:
        flash("Plantilla no encontrada", "error")
        return redirect(url_for("plantillas"))

    ruta_rel = p["ruta_docx"]
    ruta_abs = (BASE_DIR / ruta_rel).resolve()

    if not str(ruta_abs).startswith(str(BASE_DIR.resolve())):
        abort(403)

    if not ruta_abs.exists():
        flash("El archivo no existe en disco", "error")
        return redirect(url_for("plantillas"))

    return send_file(ruta_abs, as_attachment=True, download_name=p["archivo_nombre"])

@app.get("/reportes")
@login_required
def reportes():
    tipo = (request.args.get("tipo") or "").strip().upper()
    estado = (request.args.get("estado") or "").strip().upper()
    desde = (request.args.get("desde") or "").strip()
    hasta = (request.args.get("hasta") or "").strip()

    conn = db()
    ensure_solicitudes_schema(conn)

    where = []
    params = []
    
    if tipo:
        where.append("tipo_documento = ?")
        params.append(tipo)
    if estado:
        where.append("estado = ?")
        params.append(estado)
    if desde:
        where.append("fecha_emision >= ?")
        params.append(desde + " 00:00:00")
    if hasta:
        where.append("fecha_emision <= ?")
        params.append(hasta + " 23:59:59")

    sql = """
        SELECT 
            id,
            codigo_documento AS codigo,
            (nombres || ' ' || apellidos) AS nombre_completo,
            documento,
            tipo_documento AS tipo,
            estado,
            fecha_emision,
            emitido_por,
            ruta_pdf
        FROM solicitudes
    """
    where.append("estado IN ('EMITIDO', 'ANULADO')")

    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY fecha_emision DESC LIMIT 500"

    rows = conn.execute(sql, params).fetchall()
    total = len(rows)
    conn.close()

    return render_template(
        "reportes.html",
        active="reportes",
        reportes=rows,
        total=total,
        filtros={"tipo": tipo, "estado": estado, "desde": desde, "hasta": hasta},
    )


@app.get("/configuracion")
@login_required
def configuracion():
    conn = db()
    u = conn.execute("SELECT * FROM usuarios WHERE id = ?", (session["user_id"],)).fetchone()
    config = get_config(conn)
    conn.close()

    return render_template(
        "configuracion.html",
        active="configuracion",
        usuario=u,
        config=config
    )

@app.post("/configuracion/guardar")
@login_required
def configuracion_guardar():
    ruta_salida = (request.form.get("ruta_salida") or "").strip()
    correo_emisor = (request.form.get("correo_emisor") or "").strip()
    envio_correo = 1 if request.form.get("envio_correo") == "1" else 0

    if not ruta_salida:
        flash("La ruta de salida no puede estar vacía", "error")
        return redirect(url_for("configuracion"))

    # normaliza slash final
    ruta_salida = ruta_salida.replace("\\", "/")
    if not ruta_salida.endswith("/"):
        ruta_salida += "/"

    conn = db()
    ensure_config_schema(conn)
    conn.execute("""
        UPDATE configuracion
        SET ruta_salida = ?, correo_emisor = ?, envio_correo = ?, actualizado_en = ?
        WHERE id = 1
    """, (ruta_salida, correo_emisor, envio_correo, ahora()))
    conn.commit()
    conn.close()

    flash("Configuración guardada", "success")
    return redirect(url_for("configuracion"))

@app.post("/configuracion/cambiar-password")
@login_required
def configuracion_cambiar_password():
    actual = (request.form.get("password_actual") or "").strip()
    nueva = (request.form.get("password_nueva") or "").strip()
    confirmar = (request.form.get("password_confirmar") or "").strip()

    if not actual or not nueva or not confirmar:
        flash("Completa todos los campos de contraseña", "error")
        return redirect(url_for("configuracion"))

    if nueva != confirmar:
        flash("La nueva contraseña y su confirmación no coinciden", "error")
        return redirect(url_for("configuracion"))

    if len(nueva) < 6:
        flash("La nueva contraseña debe tener al menos 6 caracteres", "error")
        return redirect(url_for("configuracion"))

    conn = db()
    u = conn.execute("SELECT * FROM usuarios WHERE id = ?", (session["user_id"],)).fetchone()

    if not u or not check_password_hash(u["password_hash"], actual):
        conn.close()
        flash("La contraseña actual es incorrecta", "error")
        return redirect(url_for("configuracion"))

    conn.execute("""
        UPDATE usuarios
        SET password_hash = ?
        WHERE id = ?
    """, (generate_password_hash(nueva), session["user_id"]))
    conn.commit()
    conn.close()

    flash("Contraseña actualizada", "success")
    return redirect(url_for("configuracion"))

@app.get("/documento/<int:doc_id>/ver")
@login_required
def ver_pdf(doc_id):
    conn = db()
    s = get_solicitud_por_id(conn, doc_id)
    conn.close()
    
    if not s or not s["ruta_pdf"]:
        flash("Documento no encontrado", "error")
        return redirect(url_for("reportes"))
        
    ruta_abs = (BASE_DIR / s["ruta_pdf"]).resolve()
    return send_file(ruta_abs, mimetype='application/pdf', as_attachment=False)

@app.get("/documento/<int:doc_id>/descargar")
@login_required
def descargar_doc(doc_id):
    conn = db()
    s = get_solicitud_por_id(conn, doc_id)
    conn.close()
    
    if not s or not s["ruta_pdf"]:
        flash("Documento no encontrado", "error")
        return redirect(url_for("reportes"))
        
    ruta_abs = (BASE_DIR / s["ruta_pdf"]).resolve()
    nombre_descarga = Path(s["ruta_pdf"]).name
    return send_file(ruta_abs, as_attachment=True, download_name=nombre_descarga)

if __name__ == "__main__":
    app.run(debug=True)

 