#prueba
# app.py
import csv
import os
import sqlite3
from flask import Flask, flash, g, redirect, render_template, request, url_for

app = Flask(__name__)
app.config['DATABASE'] = 'laboratorio.db'
app.secret_key = os.environ.get('SECRET_KEY', 'supersecretkey')

CATALOGS = {
    'estados': {
        'table': 'EstadosCatalogo',
        'column': 'estado',
        'default': 'PEND. DE REVISION',
        'title': 'Estados de reparación'
    },
    'equipos': {
        'table': 'EquiposCatalogo',
        'column': 'equipo',
        'default': 'DESCONOCIDO',
        'title': 'Equipos'
    },
    'salas': {
        'table': 'SalasCatalogo',
        'column': 'sala',
        'default': 'DESCONOCIDA',
        'title': 'Salas'
    },
    'familias': {
        'table': 'FamiliasCatalogo',
        'column': 'familia',
        'default': '',
        'title': 'Familias'
    },
    'tecnicos': {
        'table': 'Tecnicos',
        'column': None,
        'default': None,
        'title': 'Técnicos'
    }
}

INDEX_FILTERS = [
    ('PEND_REPARACION', 'Pend. reparación'),
    ('EN_REPARACION', 'En reparación'),
    ('REPARADOS', 'Reparados'),
    ('PENDIENTES', 'Pendientes'),
    ('SIN_REPARACION', 'Sin reparación'),
    ('TODAS', 'Todas')
]

def ensure_support_tables(db):
    db.executescript(
        '''
        CREATE TABLE IF NOT EXISTS Reparaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            sala TEXT NOT NULL,
            uid TEXT,
            npu TEXT,
            parte TEXT,
            familia TEXT,
            equipo TEXT NOT NULL,
            urgente TEXT CHECK(urgente IN ('SI', 'NO')) DEFAULT 'NO',
            tecnico_id INTEGER,
            estado TEXT DEFAULT 'PEND. DE REVISION',
            observaciones TEXT,
            FOREIGN KEY (tecnico_id) REFERENCES Tecnicos (id)
        );

        CREATE TABLE IF NOT EXISTS Tecnicos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        );

        CREATE TABLE IF NOT EXISTS EstadosCatalogo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        );

        CREATE TABLE IF NOT EXISTS EquiposCatalogo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            valor TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS SalasCatalogo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        );

        CREATE TABLE IF NOT EXISTS FamiliasCatalogo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        );
        '''
    )
    columnas_equipos = [row['name'] for row in db.execute("PRAGMA table_info('EquiposCatalogo')").fetchall()]
    if 'valor' not in columnas_equipos:
        db.execute("ALTER TABLE EquiposCatalogo ADD COLUMN valor TEXT DEFAULT ''")
    columnas_tecnicos = [row['name'] for row in db.execute("PRAGMA table_info('Tecnicos')").fetchall()]
    if 'activo' not in columnas_tecnicos:
        db.execute("ALTER TABLE Tecnicos ADD COLUMN activo INTEGER NOT NULL DEFAULT 1")

def ensure_catalog_value(db, table, value, valor=''):
    nombre = (value or '').strip()
    if nombre:
        if table == 'EquiposCatalogo':
            db.execute(f'INSERT OR IGNORE INTO {table} (nombre, valor) VALUES (?, ?)', (nombre, (valor or '').strip()))
            if valor not in (None, ''):
                db.execute(f'UPDATE {table} SET valor = ? WHERE nombre = ?', ((valor or '').strip(), nombre))
        elif table == 'Tecnicos':
            db.execute(f'INSERT OR IGNORE INTO {table} (nombre) VALUES (?)', (nombre,))
            db.execute(f'UPDATE {table} SET activo = 1 WHERE nombre = ?', (nombre,))
        else:
            db.execute(f'INSERT OR IGNORE INTO {table} (nombre) VALUES (?)', (nombre,))

def sync_catalogs_from_reparaciones(db):
    for tipo in ('estados', 'salas', 'familias'):
        config = CATALOGS[tipo]
        ensure_catalog_value(db, config['table'], config['default'])
        rows = db.execute(
            f'''
            SELECT DISTINCT {config["column"]} AS nombre
            FROM Reparaciones
            WHERE {config["column"]} IS NOT NULL AND TRIM({config["column"]}) <> ''
            '''
        ).fetchall()
        for row in rows:
            ensure_catalog_value(db, config['table'], row['nombre'])

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(app.config['DATABASE'])
        g.db.row_factory = sqlite3.Row
        ensure_support_tables(g.db)
        g.db.commit()
    return g.db

def init_db():
    with app.app_context():
        db = get_db()
        ensure_support_tables(db)
        if os.path.exists('schema.sql'):
            with app.open_resource('schema.sql', mode='r') as f:
                db.cursor().executescript(f.read())
        sync_catalogs_from_reparaciones(db)
        db.commit()

def get_catalog_items(db, tipo):
    table = CATALOGS[tipo]['table']
    if tipo == 'equipos':
        return db.execute(f'SELECT id, nombre, valor FROM {table} ORDER BY nombre').fetchall()
    if tipo == 'tecnicos':
        return db.execute(
            '''
            SELECT
                t.id,
                t.nombre,
                COALESCE(t.activo, 1) as activo,
                SUM(CASE WHEN UPPER(COALESCE(r.estado, '')) = 'EN REPARACION' THEN 1 ELSE 0 END) AS rep_en_su_poder,
                SUM(CASE WHEN UPPER(COALESCE(r.estado, '')) LIKE 'REPARADO%' THEN 1 ELSE 0 END) AS reparados
            FROM Tecnicos t
            LEFT JOIN Reparaciones r ON r.tecnico_id = t.id
            GROUP BY t.id, t.nombre, t.activo
            ORDER BY t.activo DESC, t.nombre ASC
            '''
        ).fetchall()
    return db.execute(f'SELECT id, nombre FROM {table} ORDER BY nombre').fetchall()

def get_reparacion_form_options(db):
    return {
        'tecnicos': db.execute("SELECT id, nombre FROM Tecnicos WHERE COALESCE(activo, 1) = 1 ORDER BY nombre").fetchall(),
        'estados': get_catalog_items(db, 'estados'),
        'equipos': get_catalog_items(db, 'equipos'),
        'salas': get_catalog_items(db, 'salas'),
        'familias': get_catalog_items(db, 'familias')
    }

def normalize_catalog_value(tipo, nombre):
    value = (nombre or '').strip()
    if tipo == 'estados':
        return value.upper()
    return value


def normalize_price_value(value):
    digits = ''.join(ch for ch in str(value or '') if ch.isdigit())
    if not digits:
        return ''
    return str(int(digits))


def format_price(value):
    normalized = normalize_price_value(value)
    if not normalized:
        return ''
    return f"$ {int(normalized):,}".replace(',', '.')


def build_equipo_nombre(nombre, lab):
    base_name = (nombre or '').strip()
    lab_code = ''.join(str(lab or '').strip().upper().split())

    if not base_name:
        return ''
    if not lab_code:
        return base_name
    if not lab_code.startswith('LAB'):
        lab_code = f'LAB{lab_code}'
    return f'{base_name} ({lab_code})'


def config_redirect(tipo):
    return redirect(url_for('configuracion') + f'#panel-{tipo}')

@app.teardown_appcontext
def close_connection(exception):
    db = g.pop('db', None)
    if db is not None:
        db.close()

@app.route('/')
def index():
    db = get_db()
    estados_validos = {estado for estado, _ in INDEX_FILTERS}

    estado_actual = request.args.get('estado', 'PEND_REPARACION', type=str).upper()
    if estado_actual not in estados_validos:
        estado_actual = 'PEND_REPARACION'

    busqueda = request.args.get('q', '', type=str).strip()
    page = request.args.get('page', 1, type=int)
    per_page = 100
    if page < 1:
        page = 1

    where_clauses = []
    query_params = []

    if estado_actual == 'PEND_REPARACION':
        where_clauses.append('UPPER(r.estado) = ?')
        query_params.append('PEND. DE REVISION')
    elif estado_actual == 'EN_REPARACION':
        where_clauses.append('UPPER(r.estado) = ?')
        query_params.append('EN REPARACION')
    elif estado_actual == 'REPARADOS':
        where_clauses.append('UPPER(r.estado) LIKE ?')
        query_params.append('REPARADO%')
    elif estado_actual == 'SIN_REPARACION':
        where_clauses.append('UPPER(r.estado) LIKE ?')
        query_params.append('SIN REPARACION%')
    elif estado_actual == 'PENDIENTES':
        where_clauses.append("UPPER(r.estado) NOT LIKE 'REPARADO%'")
        where_clauses.append("UPPER(r.estado) NOT LIKE 'SIN REPARACION%'")

    if busqueda:
        termino = f'%{busqueda}%'
        where_clauses.append('(' + ' OR '.join([
            'r.fecha LIKE ?',
            'r.sala LIKE ?',
            'r.uid LIKE ?',
            'r.npu LIKE ?',
            'r.parte LIKE ?',
            'r.familia LIKE ?',
            'r.equipo LIKE ?',
            'r.estado LIKE ?',
            'r.observaciones LIKE ?',
            't.nombre LIKE ?'
        ]) + ')')
        query_params.extend([termino] * 10)

    where_sql = ''
    if where_clauses:
        where_sql = ' WHERE ' + ' AND '.join(where_clauses)

    count_query = '''
        SELECT COUNT(*) AS total
        FROM Reparaciones r
        LEFT JOIN Tecnicos t ON r.tecnico_id = t.id
    ''' + where_sql
    total_reparaciones = db.execute(count_query, query_params).fetchone()['total']
    total_pages = max(1, (total_reparaciones + per_page - 1) // per_page)

    if page > total_pages:
        page = total_pages

    offset = (page - 1) * per_page
    query = '''
        SELECT r.*, t.nombre AS tecnico_nombre
        FROM Reparaciones r
        LEFT JOIN Tecnicos t ON r.tecnico_id = t.id
    ''' + where_sql + '''
        ORDER BY r.id DESC
        LIMIT ? OFFSET ?
    '''
    params = list(query_params)
    params.extend([per_page, offset])
    reparaciones = db.execute(query, params).fetchall()

    return render_template(
        'index.html',
        reparaciones=reparaciones,
        page=page,
        total_pages=total_pages,
        total_reparaciones=total_reparaciones,
        estado_actual=estado_actual,
        tabs_estado=INDEX_FILTERS,
        busqueda=busqueda
    )

@app.route('/cargar-csv', methods=['POST'])
def cargar_csv():
    csv_path = 'reparaciones.csv'
    if not os.path.exists(csv_path):
        flash('El archivo reparaciones.csv no existe.', 'danger')
        return redirect(url_for('index'))

    db = get_db()
    tecnicos_cache = {}
    count = 0

    try:
        # Limpieza de tablas
        for table in ['Reparaciones', 'Tecnicos', 'EstadosCatalogo', 'EquiposCatalogo', 'SalasCatalogo', 'FamiliasCatalogo']:
            db.execute(f'DELETE FROM {table}')
        db.execute("DELETE FROM sqlite_sequence WHERE name IN ('Reparaciones', 'Tecnicos', 'EstadosCatalogo', 'EquiposCatalogo', 'SalasCatalogo', 'FamiliasCatalogo')")

        ensure_catalog_value(db, CATALOGS['estados']['table'], CATALOGS['estados']['default'])
        ensure_catalog_value(db, CATALOGS['equipos']['table'], CATALOGS['equipos']['default'])
        ensure_catalog_value(db, CATALOGS['salas']['table'], CATALOGS['salas']['default'])
        ensure_catalog_value(db, CATALOGS['familias']['table'], CATALOGS['familias']['default'])

        with open(csv_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            
            # Normalizar los encabezados del CSV a minúsculas
            if reader.fieldnames:
                reader.fieldnames = [name.strip().lower() for name in reader.fieldnames if name]

            for row in reader:
                tecnico_nombre = (row.get('tecnico', 'SIN ASIGNAR') or 'SIN ASIGNAR').strip()
                tecnico_nombre_up = tecnico_nombre.upper()
                tecnico_id = None

                if tecnico_nombre_up != 'SIN ASIGNAR':
                    if tecnico_nombre_up not in tecnicos_cache:
                        cur = db.execute('INSERT INTO Tecnicos (nombre) VALUES (?)', (tecnico_nombre,))
                        tecnico_id = cur.lastrowid
                        tecnicos_cache[tecnico_nombre_up] = tecnico_id
                    else:
                        tecnico_id = tecnicos_cache[tecnico_nombre_up]

                urgente = (row.get('urgente', 'NO') or 'NO').strip().upper()
                if urgente not in ('SI', 'NO'): urgente = 'NO'

                fecha = row.get('fecha', '') or '1/1/2000'
                sala = (row.get('sala', '') or CATALOGS['salas']['default']).strip()
                equipo = (row.get('equipo', '') or CATALOGS['equipos']['default']).strip()
                familia = (row.get('familia', '') or '').strip()
                estado = (row.get('estados de reparacion', CATALOGS['estados']['default']) or CATALOGS['estados']['default']).strip().upper()

                db.execute(
                    '''
                    INSERT INTO Reparaciones (
                        fecha, sala, uid, npu, parte, familia, equipo, urgente, tecnico_id, estado, observaciones
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''',
                    (
                        fecha, 
                        sala, 
                        row.get('uid', ''),
                        row.get('npu', ''), 
                        row.get('parte', ''), 
                        familia, 
                        equipo, 
                        urgente, 
                        tecnico_id, 
                        estado, 
                        row.get('observaciones', '')
                    )
                )
                count += 1

        sync_catalogs_from_reparaciones(db)
        db.commit()
        flash(f'Carga exitosa: {count} registros. Identificadores mapeados correctamente.', 'success')
    except Exception as e:
        db.rollback()
        flash(f'Error durante la importación: {str(e)}', 'danger')

    return redirect(url_for('index'))

@app.route('/reparacion/nueva', methods=['GET', 'POST'])
def nueva_reparacion():
    db = get_db()
    if request.method == 'POST':
        sala = normalize_catalog_value('salas', request.form['sala']) or CATALOGS['salas']['default']
        equipo = normalize_catalog_value('equipos', request.form['equipo']) or CATALOGS['equipos']['default']
        familia = normalize_catalog_value('familias', request.form['familia']) or CATALOGS['familias']['default']
        estado = normalize_catalog_value('estados', request.form['estado']) or CATALOGS['estados']['default']
        tecnico_id = request.form.get('tecnico_id') or None

        ensure_catalog_value(db, CATALOGS['salas']['table'], sala)
        ensure_catalog_value(db, CATALOGS['equipos']['table'], equipo)
        ensure_catalog_value(db, CATALOGS['familias']['table'], familia)
        ensure_catalog_value(db, CATALOGS['estados']['table'], estado)

        db.execute(
            '''
            INSERT INTO Reparaciones (fecha, sala, uid, npu, parte, familia, equipo, urgente, tecnico_id, estado, observaciones)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''',
            (
                request.form['fecha'],
                sala,
                request.form['uid'],
                request.form['npu'],
                request.form.get('parte', ''),
                request.form['familia'],
                equipo,
                request.form['urgente'],
                tecnico_id,
                estado,
                request.form['observaciones']
            )
        )
        db.commit()
        return redirect(url_for('index'))

    return render_template('reparacion.html', reparacion=None, **get_reparacion_form_options(db))

@app.route('/reparacion/editar/<int:id>', methods=['GET', 'POST'])
def editar_reparacion(id):
    db = get_db()
    if request.method == 'POST':
        sala = normalize_catalog_value('salas', request.form['sala']) or CATALOGS['salas']['default']
        equipo = normalize_catalog_value('equipos', request.form['equipo']) or CATALOGS['equipos']['default']
        familia = normalize_catalog_value('familias', request.form['familia']) or CATALOGS['familias']['default']
        estado = normalize_catalog_value('estados', request.form['estado']) or CATALOGS['estados']['default']
        tecnico_id = request.form.get('tecnico_id') or None

        ensure_catalog_value(db, CATALOGS['salas']['table'], sala)
        ensure_catalog_value(db, CATALOGS['equipos']['table'], equipo)
        ensure_catalog_value(db, CATALOGS['familias']['table'], familia)
        ensure_catalog_value(db, CATALOGS['estados']['table'], estado)

        db.execute(
            '''
            UPDATE Reparaciones SET
            fecha=?, sala=?, uid=?, npu=?, parte=?, familia=?, equipo=?, urgente=?, tecnico_id=?, estado=?, observaciones=?
            WHERE id=?
            ''',
            (
                request.form['fecha'],
                sala,
                request.form['uid'],
                request.form['npu'],
                request.form.get('parte', ''),
                request.form['familia'],
                equipo,
                request.form['urgente'],
                tecnico_id,
                estado,
                request.form['observaciones'],
                id
            )
        )
        db.commit()
        return redirect(url_for('index'))

    reparacion = db.execute('SELECT * FROM Reparaciones WHERE id = ?', (id,)).fetchone()
    return render_template('reparacion.html', reparacion=reparacion, **get_reparacion_form_options(db))

@app.route('/reparacion/eliminar/<int:id>')
def eliminar_reparacion(id):
    db = get_db()
    db.execute('DELETE FROM Reparaciones WHERE id = ?', (id,))
    db.commit()
    return redirect(url_for('index'))

@app.route('/api/reparaciones/verificar-activo', methods=['GET'])
def verificar_activo():
    uid = request.args.get('uid', '').strip()
    npu = request.args.get('npu', '').strip()
    
    if not uid and not npu:
        return jsonify({'existe': False})
    
    db = get_db()
    query = "SELECT id, sala, equipo, estado FROM Reparaciones WHERE (UPPER(estado) NOT LIKE 'REPARADO%' AND UPPER(estado) <> 'ENTREGADO')"
    params = []
    
    if uid and npu:
        query += " AND (uid = ? OR npu = ?)"
        params.extend([uid, npu])
    elif uid:
        query += " AND uid = ?"
        params.append(uid)
    else:
        query += " AND npu = ?"
        params.append(npu)
        
    result = db.execute(query, params).fetchone()
    
    if result:
        return jsonify({
            'existe': True,
            'reparacion': {
                'id': result['id'],
                'sala': result['sala'],
                'equipo': result['equipo'],
                'estado': result['estado']
            }
        })
    
    return jsonify({'existe': False})

@app.route('/configuracion')
def configuracion():
    db = get_db()
    return render_template(
        'configuracion.html',
        catalogos={tipo: get_catalog_items(db, tipo) for tipo in CATALOGS},
        catalog_config=CATALOGS,
        format_price=format_price
    )

@app.route('/configuracion/<tipo>/agregar', methods=['POST'])
def agregar_catalogo(tipo):
    if tipo not in CATALOGS:
        return config_redirect(tipo)

    db = get_db()
    if tipo == 'equipos':
        nombre = normalize_catalog_value(
            tipo,
            build_equipo_nombre(
                request.form.get('nombre_equipo', ''),
                request.form.get('lab', '')
            )
        )
        valor = normalize_price_value(request.form.get('valor', ''))
    else:
        nombre = normalize_catalog_value(tipo, request.form.get('nombre', ''))
        valor = request.form.get('valor', '').strip()
    if not nombre:
        flash('El nombre no puede estar vacío.', 'warning')
        return config_redirect(tipo)

    try:
        ensure_catalog_value(db, CATALOGS[tipo]['table'], nombre, valor)
        db.commit()
        if tipo == 'equipos':
            flash('Se agregó un equipo nuevo', 'success')
        else:
            flash(f'{CATALOGS[tipo]["title"]}: valor agregado.', 'success')
    except sqlite3.IntegrityError:
        db.rollback()
        flash('Ese valor ya existe.', 'warning')

    return config_redirect(tipo)

@app.route('/configuracion/<tipo>/editar/<int:item_id>', methods=['POST'])
def editar_catalogo(tipo, item_id):
    if tipo not in CATALOGS:
        return config_redirect(tipo)

    db = get_db()
    config = CATALOGS[tipo]
    actual = db.execute(f'SELECT id, nombre FROM {config["table"]} WHERE id = ?', (item_id,)).fetchone()
    if not actual:
        flash('No se encontró el registro.', 'warning')
        return config_redirect(tipo)

    nuevo_nombre = normalize_catalog_value(tipo, request.form.get('nombre', ''))
    nuevo_valor = normalize_price_value(request.form.get('valor', '')) if tipo == 'equipos' else request.form.get('valor', '').strip()
    if not nuevo_nombre:
        flash('El nombre no puede estar vacío.', 'warning')
        return config_redirect(tipo)

    try:
        if tipo == 'equipos':
            db.execute(f'UPDATE {config["table"]} SET nombre = ?, valor = ? WHERE id = ?', (nuevo_nombre, nuevo_valor, item_id))
        else:
            db.execute(f'UPDATE {config["table"]} SET nombre = ? WHERE id = ?', (nuevo_nombre, item_id))
        if config['column']:
            db.execute(
                f'UPDATE Reparaciones SET {config["column"]} = ? WHERE {config["column"]} = ?',
                (nuevo_nombre, actual['nombre'])
            )
        db.commit()
        flash(f'{config["title"]}: valor actualizado.', 'success')
    except sqlite3.IntegrityError:
        db.rollback()
        flash('Ese valor ya existe.', 'warning')

    return config_redirect(tipo)

@app.route('/configuracion/<tipo>/eliminar/<int:item_id>', methods=['POST'])
def eliminar_catalogo(tipo, item_id):
    if tipo not in CATALOGS:
        return config_redirect(tipo)

    db = get_db()
    config = CATALOGS[tipo]
    actual = db.execute(f'SELECT id, nombre FROM {config["table"]} WHERE id = ?', (item_id,)).fetchone()
    if not actual:
        flash('No se encontró el registro.', 'warning')
        return config_redirect(tipo)

    if tipo == 'tecnicos':
        db.execute('UPDATE Tecnicos SET activo = 0 WHERE id = ?', (item_id,))
        db.commit()
        flash('Técnico eliminado.', 'success')
        return config_redirect(tipo)

    if actual['nombre'] == config['default']:
        flash('No se puede eliminar el valor por defecto.', 'warning')
        return config_redirect(tipo)

    ensure_catalog_value(db, config['table'], config['default'])
    db.execute(
        f'UPDATE Reparaciones SET {config["column"]} = ? WHERE {config["column"]} = ?',
        (config['default'], actual['nombre'])
    )
    db.execute(f'DELETE FROM {config["table"]} WHERE id = ?', (item_id,))
    db.commit()
    flash(f'{config["title"]}: valor eliminado.', 'success')
    return config_redirect(tipo)

@app.route('/configuracion/tecnicos/desactivar/<int:item_id>', methods=['POST'])
def desactivar_tecnico(item_id):
    db = get_db()
    db.execute('UPDATE Tecnicos SET activo = 0 WHERE id = ?', (item_id,))
    db.commit()
    flash('Técnico desactivado.', 'success')
    return config_redirect('tecnicos')

@app.route('/configuracion/tecnicos/eliminar/<int:item_id>', methods=['POST'])
def eliminar_tecnico_permanente(item_id):
    db = get_db()
    # Para evitar romper integridad referencial, podrías verificar si tiene reparacionesssss
    db.execute('DELETE FROM Tecnicos WHERE id = ?', (item_id,))
    db.commit()
    flash('Técnico eliminado permanentemente.', 'success')
    return config_redirect('tecnicos')

@app.route('/configuracion/tecnicos/reactivar/<int:item_id>', methods=['POST'])
def reactivar_tecnico(item_id):
    db = get_db()
    db.execute('UPDATE Tecnicos SET activo = 1 WHERE id = ?', (item_id,))
    db.commit()
    flash('Técnico reactivado.', 'success')
    return config_redirect('tecnicos')

@app.route('/configuracion/salas/sincronizar', methods=['POST'])
def sincronizar_salas():
    db = get_db()
    try:
        config = CATALOGS['salas']
        ensure_catalog_value(db, config['table'], config['default'])
        rows = db.execute(
            f'''
            SELECT DISTINCT {config["column"]} AS nombre
            FROM Reparaciones
            WHERE {config["column"]} IS NOT NULL AND TRIM({config["column"]}) <> ''
            '''
        ).fetchall()
        
        nuevas = 0
        for row in rows:
            nombre = row['nombre'].strip()
            if nombre:
                exists = db.execute(f"SELECT 1 FROM {config['table']} WHERE nombre = ?", (nombre,)).fetchone()
                if not exists:
                    db.execute(f"INSERT INTO {config['table']} (nombre) VALUES (?)", (nombre,))
                    nuevas += 1
        
        db.commit()
        if nuevas > 0:
            flash(f'Sincronización completada. Se añadieron {nuevas} nuevas salas.', 'success')
        else:
            flash('No se encontraron salas nuevas para añadir.', 'info')
    except Exception as e:
        db.rollback()
        flash(f'Error al sincronizar salas: {str(e)}', 'danger')
    
    return redirect(url_for('configuracion') + '#panel-salas')

@app.route('/configuracion/equipos/ajuste-global', methods=['POST'])
def ajuste_global_precios():
    db = get_db()
    tipo_ajuste = request.form.get('tipo_ajuste')  # 'porcentaje' o 'fijo'
    valor_ajuste = request.form.get('valor_ajuste')
    operacion = request.form.get('operacion')  # 'aumento' o 'descuento'

    try:
        valor = float(valor_ajuste)
        if valor < 0: raise ValueError()
    except:
        flash('Valor de ajuste inválido.', 'warning')
        return redirect(url_for('configuracion') + '#panel-equipos')

    equipos = db.execute('SELECT id, valor FROM EquiposCatalogo').fetchall()
    for eq in equipos:
        try:
            precio_actual = int(normalize_price_value(eq['valor']) or 0)
            nuevo_precio = precio_actual
            
            if tipo_ajuste == 'porcentaje':
                modificador = (precio_actual * (valor / 100))
                nuevo_precio = precio_actual + modificador if operacion == 'aumento' else precio_actual - modificador
            else:
                nuevo_precio = precio_actual + valor if operacion == 'aumento' else precio_actual - valor
            
            nuevo_precio = max(0, int(nuevo_precio))
            db.execute('UPDATE EquiposCatalogo SET valor = ? WHERE id = ?', (str(nuevo_precio), eq['id']))
        except:
            continue
    
    db.commit()
    flash('Precios actualizados globalmente.', 'success')
    return redirect(url_for('configuracion') + '#panel-equipos')

@app.route('/configuracion/equipos/exportar', methods=['GET'])
def exportar_equipos_csv():
    db = get_db()
    equipos = db.execute('SELECT nombre, valor FROM EquiposCatalogo ORDER BY nombre').fetchall()
    
    def generate():
        data = [['equipo', 'precio']]
        for eq in equipos:
            data.append([eq['nombre'], eq['valor']])
        
        import io
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerows(data)
        yield output.getvalue()

    from flask import Response
    return Response(
        generate(),
        mimetype='text/csv',
        headers={"Content-disposition": "attachment; filename=equipos_catalogo.csv"}
    )

@app.route('/configuracion/equipos/acciones-masivas', methods=['POST'])
def acciones_masivas_equipos():
    db = get_db()
    ids = request.form.getlist('ids[]')
    accion = request.form.get('accion')

    if not ids:
        flash('No se seleccionaron equipos.', 'warning')
        return redirect(url_for('configuracion') + '#panel-equipos')

    try:
        if accion == 'eliminar':
            placeholders = ','.join(['?'] * len(ids))
            db.execute(f'DELETE FROM EquiposCatalogo WHERE id IN ({placeholders})', ids)
            db.commit()
            flash(f'Se eliminaron {len(ids)} equipos.', 'success')
    except Exception as e:
        db.rollback()
        flash(f'Error al procesar acción masiva: {str(e)}', 'danger')

    return redirect(url_for('configuracion') + '#panel-equipos')

@app.route('/configuracion/equipos/cargar-csv', methods=['POST'])
def cargar_equipos_csv():
    db = get_db()
    archivo = request.files.get('archivo_csv')
    if not archivo or not archivo.filename:
        flash('Seleccioná un archivo CSV para equipos.', 'warning')
        return redirect(url_for('configuracion') + '#panel-equipos')

    cargados = 0
    try:
        contenido = archivo.stream.read().decode('utf-8-sig')
        lector = csv.DictReader(contenido.splitlines())
        if not lector.fieldnames:
            raise ValueError('El CSV no tiene encabezados.')

        campos = {campo.strip().lower(): campo for campo in lector.fieldnames if campo}
        campo_nombre = campos.get('equipo') or campos.get('nombre')
        campo_valor = campos.get('precio') or campos.get('valor')

        if not campo_nombre:
            raise ValueError('El CSV debe incluir una columna "equipo" o "nombre".')

        db.execute('DELETE FROM EquiposCatalogo')
        db.execute("DELETE FROM sqlite_sequence WHERE name = 'EquiposCatalogo'")

        for row in lector:
            nombre = normalize_catalog_value('equipos', row.get(campo_nombre, ''))
            valor = normalize_price_value(row.get(campo_valor, '')) if campo_valor else ''
            if not nombre:
                continue
            ensure_catalog_value(db, CATALOGS['equipos']['table'], nombre, valor)
            cargados += 1

        db.commit()
        flash(f'Se cargaron {cargados} equipos desde el CSV.', 'success')
    except Exception as e:
        db.rollback()
        flash(f'Error al cargar equipos desde CSV: {str(e)}', 'danger')

    return redirect(url_for('configuracion') + '#panel-equipos')

@app.route('/tecnicos')
def tecnicos():
    return redirect(url_for('configuracion') + '#panel-tecnicos')

@app.route('/tecnico/eliminar/<int:id>')
def eliminar_tecnico(id):
    db = get_db()
    db.execute('UPDATE Reparaciones SET tecnico_id = NULL WHERE tecnico_id = ?', (id,))
    db.execute('DELETE FROM Tecnicos WHERE id = ?', (id,))
    db.commit()
    return redirect(url_for('configuracion') + '#panel-tecnicos')

if __name__ == '__main__':
    if not os.path.exists(app.config['DATABASE']):
        init_db()
    app.run(debug=True)
