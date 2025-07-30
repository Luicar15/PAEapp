import os
import pandas as pd
import sqlite3
import uuid
from flask import render_template, request, redirect, flash

DB_FILE = 'formulario_inicial.db'

# Columnas requeridas en el Excel cargado
COLUMNAS_REQUERIDAS = [
    "Municipio",
    "Institucion Educativa (IED)",
    "Sede",
    "Complemento Alimenticio",
    "Fecha Inicial",
    "Fecha Final",
    "Ciclo de Menú (Semana)",
    "Grado 0 (Raciones)",
    "Grado 1,2,3 (Raciones)",
    "Grado 4,5 (Raciones)",
    "Grado 6,7,8,9 (Raciones)",
    "Grado 10,11 (Raciones)",
    "Cupos Totales",
    "Días Atendidos",
    "Responsable",
    "Observaciones"
]

def cargar_excel_masivo():
    if request.method == 'POST':
        archivo = request.files.get('archivo_excel')
        if not archivo or archivo.filename == '':
            flash('Debe seleccionar un archivo Excel válido.', 'danger')
            return redirect(request.url)

        # Guardar archivo temporalmente
        ruta_temp = os.path.join('data', 'uploads', archivo.filename)
        os.makedirs(os.path.dirname(ruta_temp), exist_ok=True)
        archivo.save(ruta_temp)

        try:
            df = pd.read_excel(ruta_temp)

            # Validar columnas
            if not all(col in df.columns for col in COLUMNAS_REQUERIDAS):
                flash('El archivo cargado no tiene el formato correcto.', 'danger')
                return redirect(request.url)

            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()

            errores = []
            cargados = 0

            for _, row in df.iterrows():
                # Validar IED y Sede en la base precargada
                cursor.execute('''
                    SELECT COUNT(*) FROM ied_sedes
                    WHERE municipio=? AND institucion=? AND sede=?
                ''', (row['Municipio'], row['Institucion Educativa (IED)'], row['Sede']))
                if cursor.fetchone()[0] == 0:
                    errores.append(f"No existe IED/Sede: {row['Institucion Educativa (IED)']} - {row['Sede']}")
                    continue

                # Validar Ciclo y Complemento en ciclos_menu
                cursor.execute('''
                    SELECT COUNT(*) FROM ciclos_menu
                    WHERE ciclo=? AND complemento=?
                ''', (row['Ciclo de Menú (Semana)'], row['Complemento Alimenticio']))
                if cursor.fetchone()[0] == 0:
                    errores.append(f"No existe ciclo/complemento para {row['Institucion Educativa (IED)']}")
                    continue

                # Insertar registro válido en formularios con UUID generado
                cursor.execute('''
                    INSERT INTO formularios (
                        id_sede, municipio, institucion, sede, complemento_alimenticio,
                        fecha_inicial, fecha_final, ciclo_menu, raciones_preescolar,
                        raciones_primaria_baja, raciones_primaria_alta, raciones_secundaria,
                        raciones_myc, cupos_totales, dias_atendidos, responsable,
                        observaciones
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    str(uuid.uuid4()), row['Municipio'], row['Institucion Educativa (IED)'], row['Sede'],
                    row['Complemento Alimenticio'], row['Fecha Inicial'], row['Fecha Final'],
                    row['Ciclo de Menú (Semana)'], row['Grado 0 (Raciones)'], row['Grado 1,2,3 (Raciones)'],
                    row['Grado 4,5 (Raciones)'], row['Grado 6,7,8,9 (Raciones)'], row['Grado 10,11 (Raciones)'],
                    row['Cupos Totales'], row['Días Atendidos'], row['Responsable'], row['Observaciones']
                ))
                cargados += 1

            conn.commit()
            conn.close()

            # Vista previa de los primeros 10 registros
            preview = df.head(10).to_html(classes='table table-striped', index=False)
            return render_template('cargue_masivo_preview.html', preview=preview, errores=errores, cargados=cargados)

        except Exception as e:
            flash(f'Error al procesar el archivo: {str(e)}', 'danger')
            return redirect(request.url)

    return render_template('cargue_masivo.html')