import os
import sqlite3

DB_PATH = os.path.join('data', 'formulario_inicial.db')

def get_connection():
    """Devuelve una conexión a SQLite."""
    os.makedirs('data', exist_ok=True)
    return sqlite3.connect(DB_PATH)

def init_db():
    """Crea la base si no existe."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute('''CREATE TABLE IF NOT EXISTS formularios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        municipio TEXT, institucion TEXT, sede TEXT,
        complemento_alimenticio TEXT, fecha_inicial TEXT, fecha_final TEXT,
        ciclo_menu TEXT, raciones_preescolar INTEGER,
        raciones_primaria_baja INTEGER, raciones_primaria_alta INTEGER,
        raciones_secundaria INTEGER, raciones_myc INTEGER,
        cupos_totales INTEGER, dias_atendidos INTEGER,
        responsable TEXT, observaciones TEXT
    )''')

    conn.commit()
    conn.close()