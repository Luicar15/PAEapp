import os
import sqlite3
import pandas as pd

DB_PATH = 'formulario_inicial.db'
EXCEL_PATH = 'datos_base_completo_final_v2.xlsx'

TABLES_SHEETS = {
    'municipios': 'municipios',
    'zonas': 'zonas',
    'instituciones': 'instituciones',
    'sedes': 'sedes',
    'menus': 'menus'
}

def create_tables(conn):
    cursor = conn.cursor()

    # Forzar limpieza de tablas para evitar conflictos de esquema
    tables = ['municipios', 'zonas', 'instituciones', 'sedes', 'menus']
    for t in tables:
        cursor.execute(f"DROP TABLE IF EXISTS {t}")

    # Crear tablas desde cero
    cursor.execute('''
        CREATE TABLE municipios (
            id INTEGER PRIMARY KEY,
            nombre TEXT NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE zonas (
            id INTEGER PRIMARY KEY,
            nombre TEXT NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE instituciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            municipio_id INTEGER,
            zona_id INTEGER,
            FOREIGN KEY (municipio_id) REFERENCES municipios(id),
            FOREIGN KEY (zona_id) REFERENCES zonas(id)
        )
    ''')
    cursor.execute('''
        CREATE TABLE sedes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            institucion_id INTEGER,
            FOREIGN KEY (institucion_id) REFERENCES instituciones(id)
        )
    ''')
    cursor.execute('''
        CREATE TABLE menus (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            menu TEXT,
            grupo TEXT,
            componente TEXT,
            preparacion TEXT,
            categoria TEXT,
            alimento TEXT,
            codigo TEXT,
            cantidad_por_racion REAL,
            unidad TEXT
        )
    ''')

    conn.commit()

def sanitize_dataframe(df, table):
    df.columns = [str(col).strip().lower().replace(' ', '_') for col in df.columns]
    df = df.dropna(how='all')
    if 'id' in df.columns:
        df = df.drop_duplicates(subset=['id'], keep='first')

    def clean_int_column(col):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            df[col] = df[col].apply(lambda x: int(x) if pd.notna(x) else None)

    if table == 'instituciones':
        for col in ['id', 'municipio_id', 'zona_id']:
            clean_int_column(col)
    elif table == 'sedes':
        for col in ['id', 'institucion_id']:
            clean_int_column(col)
    elif table in ['municipios', 'zonas']:
        clean_int_column('id')

    # Para mantener decimales exactos en cantidad_por_racion (solo en menus)
    if table == 'menus' and 'cantidad_por_racion' in df.columns:
        df['cantidad_por_racion'] = pd.to_numeric(df['cantidad_por_racion'], errors='coerce')

    return df

def load_excel_to_sqlite():
    if not os.path.exists(EXCEL_PATH):
        print(f"ERROR: No se encontró el archivo {EXCEL_PATH}.")
        return

    conn = sqlite3.connect(DB_PATH)
    create_tables(conn)

    for table, sheet in TABLES_SHEETS.items():
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
        df = sanitize_dataframe(df, table)

        # Si hay duplicados en IDs, eliminarlos para que SQLite autogenere
        if 'id' in df.columns and df['id'].duplicated().any():
            print(f"⚠ Aviso: IDs duplicados detectados en '{table}'. Se eliminará la columna id para autogenerar.")
            df = df.drop(columns=['id'])

        # Limpiar tabla antes de insertar
        conn.execute(f"DELETE FROM {table}")
        conn.commit()

        df.to_sql(table, conn, if_exists='append', index=False)
        print(f"Hoja '{sheet}' importada correctamente en la tabla '{table}'.")

    conn.close()
    print("¡Todos los datos fueron cargados exitosamente!")

if __name__ == '__main__':
    load_excel_to_sqlite()