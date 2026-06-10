import sqlite3
import os

def init_db():
    """Inicializa la base de datos con tablas de docentes y estudiantes"""
    os.makedirs('database', exist_ok=True)
    conn = sqlite3.connect('database/academico.db')
    cursor = conn.cursor()
    
    # Tabla docentes (con nuevos campos)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS docentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_docente TEXT NOT NULL,
            asignatura TEXT NOT NULL,
            codigo TEXT NOT NULL,
            dia_semana TEXT NOT NULL,
            horario_inicio TEXT NOT NULL,
            horario_fin TEXT NOT NULL,
            link_clase TEXT,
            fecha TEXT NOT NULL,
            duracion TEXT,
            link_grabacion TEXT,
            asistencia_estudiantes INTEGER DEFAULT 0,
            asistencia_reportados INTEGER DEFAULT 0,
            asistencia_docente TEXT,
            dependencia TEXT,
            accion TEXT,
            estado TEXT DEFAULT 'Pendiente',
            evidencia TEXT,
            observaciones TEXT,
            fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Tabla estudiantes
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS estudiantes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_estudiante TEXT NOT NULL,
            numero_documento TEXT UNIQUE NOT NULL,
            programa TEXT,
            correo_personal TEXT,
            correo_institucional TEXT,
            telefono TEXT,
            dependencia TEXT,
            accion TEXT,
            estado TEXT DEFAULT 'Pendiente',
            evidencia TEXT,
            asistencia TEXT,
            observaciones TEXT,
            fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Verificar y agregar columnas faltantes en docentes
    try:
        cursor.execute("PRAGMA table_info(docentes)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'dependencia' not in columns:
            cursor.execute("ALTER TABLE docentes ADD COLUMN dependencia TEXT")
        if 'accion' not in columns:
            cursor.execute("ALTER TABLE docentes ADD COLUMN accion TEXT")
        if 'estado' not in columns:
            cursor.execute("ALTER TABLE docentes ADD COLUMN estado TEXT DEFAULT 'Pendiente'")
        if 'evidencia' not in columns:
            cursor.execute("ALTER TABLE docentes ADD COLUMN evidencia TEXT")
    except Exception as e:
        print(f"Error al verificar/agregar columnas en docentes: {e}")
    
    # Verificar y agregar columnas faltantes en estudiantes
    try:
        cursor.execute("PRAGMA table_info(estudiantes)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'dependencia' not in columns:
            cursor.execute("ALTER TABLE estudiantes ADD COLUMN dependencia TEXT")
        if 'accion' not in columns:
            cursor.execute("ALTER TABLE estudiantes ADD COLUMN accion TEXT")
        if 'estado' not in columns:
            cursor.execute("ALTER TABLE estudiantes ADD COLUMN estado TEXT DEFAULT 'Pendiente'")
        if 'evidencia' not in columns:
            cursor.execute("ALTER TABLE estudiantes ADD COLUMN evidencia TEXT")
    except Exception as e:
        print(f"Error al verificar/agregar columnas en estudiantes: {e}")
    
    conn.commit()
    conn.close()

def get_db_connection():
    """Obtiene una conexión a la base de datos"""
    conn = sqlite3.connect('database/academico.db')
    conn.row_factory = sqlite3.Row
    return conn