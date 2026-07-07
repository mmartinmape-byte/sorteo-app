from flask import Flask, request, jsonify, render_template, send_file
from datetime import datetime
import os, re, random

app = Flask(__name__)

DATABASE_URL = os.environ.get('DATABASE_URL', '')
USE_PG = bool(DATABASE_URL and 'postgres' in DATABASE_URL)
ADMIN_KEY = os.environ.get('ADMIN_KEY', 'admin123')
NUMERO_INICIAL = 123  # primer número que se asigna; los siguientes son correlativos


# ── DB helpers ────────────────────────────────────────────────────────────────

def get_conn():
    if USE_PG:
        import psycopg2
        conn = psycopg2.connect(DATABASE_URL)
        return conn, '%s', True
    else:
        import sqlite3
        conn = sqlite3.connect('sorteo.db')
        conn.row_factory = sqlite3.Row
        return conn, '?', False

def q(sql, params=None, fetch='all'):
    conn, ph, pg = get_conn()
    sql = sql.replace('?', ph)
    try:
        if pg:
            from psycopg2.extras import RealDictCursor
            cur = conn.cursor(cursor_factory=RealDictCursor)
        else:
            cur = conn.cursor()
        cur.execute(sql, params or ())
        if fetch == 'all':
            result = [dict(r) for r in cur.fetchall()]
        elif fetch == 'one':
            r = cur.fetchone()
            result = dict(r) if r else None
        else:
            result = None
        conn.commit()
        return result
    finally:
        conn.close()

def run(sql, params=None):
    q(sql, params, fetch=None)


# ── Schema ────────────────────────────────────────────────────────────────────

def init_db():
    pk = 'SERIAL PRIMARY KEY' if USE_PG else 'INTEGER PRIMARY KEY AUTOINCREMENT'
    run(f"""CREATE TABLE IF NOT EXISTS participantes (
        id {pk},
        numero INTEGER UNIQUE NOT NULL,
        nombre TEXT NOT NULL,
        apellido TEXT NOT NULL,
        telefono TEXT NOT NULL,
        email TEXT UNIQUE NOT NULL,
        creado TEXT NOT NULL
    )""")
    run(f"""CREATE TABLE IF NOT EXISTS sorteos (
        id {pk},
        numero_ganador INTEGER NOT NULL,
        nombre TEXT NOT NULL,
        apellido TEXT NOT NULL,
        email TEXT NOT NULL,
        fecha TEXT NOT NULL
    )""")

init_db()


# ── Páginas ───────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/admin')
def admin():
    if request.args.get('clave') != ADMIN_KEY:
        return render_template('admin_login.html'), 401
    return render_template('admin.html', clave=ADMIN_KEY)


# ── API ───────────────────────────────────────────────────────────────────────

EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')

@app.route('/api/registrar', methods=['POST'])
def registrar():
    d = request.get_json(silent=True) or {}
    nombre = (d.get('nombre') or '').strip()
    apellido = (d.get('apellido') or '').strip()
    telefono = (d.get('telefono') or '').strip()
    email = (d.get('email') or '').strip().lower()

    if not nombre or not apellido or not telefono or not email:
        return jsonify(error='Todos los campos son obligatorios'), 400
    if not EMAIL_RE.match(email):
        return jsonify(error='El email no es válido'), 400
    if not re.match(r'^[\d\s\+\-\(\)]{6,20}$', telefono):
        return jsonify(error='El teléfono no es válido'), 400

    existe = q("SELECT numero FROM participantes WHERE email = ?", (email,), fetch='one')
    if existe:
        return jsonify(error='Este email ya está registrado con el número %d' % existe['numero']), 409

    # Reintenta por si dos registros simultáneos calculan el mismo número
    for _ in range(5):
        r = q("SELECT COALESCE(MAX(numero), 0) + 1 AS n FROM participantes", fetch='one')
        numero = max(r['n'], NUMERO_INICIAL)
        try:
            run("""INSERT INTO participantes (numero, nombre, apellido, telefono, email, creado)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (numero, nombre, apellido, telefono, email,
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            return jsonify(numero=numero, nombre=nombre, apellido=apellido)
        except Exception as e:
            if 'email' in str(e).lower():
                return jsonify(error='Este email ya está registrado'), 409
            continue
    return jsonify(error='No se pudo registrar, intentá de nuevo'), 500


def check_admin():
    return request.args.get('clave') == ADMIN_KEY

@app.route('/api/participantes')
def participantes():
    if not check_admin():
        return jsonify(error='No autorizado'), 401
    return jsonify(q("SELECT * FROM participantes ORDER BY numero"))

@app.route('/api/sortear', methods=['POST'])
def sortear():
    if not check_admin():
        return jsonify(error='No autorizado'), 401
    todos = q("SELECT * FROM participantes")
    if not todos:
        return jsonify(error='No hay participantes registrados'), 400

    excluir_ganadores = (request.get_json(silent=True) or {}).get('excluir_ganadores', False)
    if excluir_ganadores:
        ya_ganaron = {s['numero_ganador'] for s in q("SELECT numero_ganador FROM sorteos")}
        candidatos = [p for p in todos if p['numero'] not in ya_ganaron]
        if not candidatos:
            return jsonify(error='Todos los participantes ya ganaron algún sorteo'), 400
    else:
        candidatos = todos

    ganador = random.SystemRandom().choice(candidatos)
    run("""INSERT INTO sorteos (numero_ganador, nombre, apellido, email, fecha)
           VALUES (?, ?, ?, ?, ?)""",
        (ganador['numero'], ganador['nombre'], ganador['apellido'], ganador['email'],
         datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    return jsonify(ganador)

@app.route('/api/sorteos')
def sorteos():
    if not check_admin():
        return jsonify(error='No autorizado'), 401
    return jsonify(q("SELECT * FROM sorteos ORDER BY id DESC"))

@app.route('/api/exportar')
def exportar():
    if not check_admin():
        return jsonify(error='No autorizado'), 401
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from io import BytesIO
    wb = Workbook()
    ws = wb.active
    ws.title = 'Participantes'
    ws.append(['Número', 'Nombre', 'Apellido', 'Teléfono', 'Email', 'Registrado'])
    for c in ws[1]:
        c.font = Font(bold=True)
    for p in q("SELECT * FROM participantes ORDER BY numero"):
        ws.append([p['numero'], p['nombre'], p['apellido'], p['telefono'], p['email'], p['creado']])
    for col, ancho in zip('ABCDEF', (10, 18, 18, 18, 32, 20)):
        ws.column_dimensions[col].width = ancho
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    nombre = 'participantes_%s.xlsx' % datetime.now().strftime('%Y-%m-%d')
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/api/participantes/<int:pid>', methods=['DELETE'])
def borrar_participante(pid):
    if not check_admin():
        return jsonify(error='No autorizado'), 401
    run("DELETE FROM participantes WHERE id = ?", (pid,))
    return jsonify(ok=True)

@app.route('/api/sorteos/<int:sid>', methods=['DELETE'])
def borrar_sorteo(sid):
    if not check_admin():
        return jsonify(error='No autorizado'), 401
    run("DELETE FROM sorteos WHERE id = ?", (sid,))
    return jsonify(ok=True)


if __name__ == '__main__':
    app.run(debug=True, port=5002)
