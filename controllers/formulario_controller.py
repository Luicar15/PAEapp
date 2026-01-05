import os
import threading
from uuid import uuid4
import pandas as pd
from flask import render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
from utils.generar_kardex import generar_kardex_consolidado
from utils.generar_remision import generar_remision_consolidado

OUTPUT_DIR = os.path.join("data", "outputs")
UPLOAD_DIR = os.path.join("data", "uploads")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

DEFAULT_PROGRESS = {"procesadas": 0, "total": 0, "completado": False, "pdf_final": ""}
ALLOWED_EXTENSIONS = {"xls", "xlsx"}
MAX_UPLOAD_SIZE = 10 * 1024 * 1024  # 10 MB

progreso_kardex = {}
progreso_remision = {}
archivo_cargado = {}


def _get_session_id():
    if "session_id" not in session:
        session["session_id"] = uuid4().hex
    return session["session_id"]


def _get_progress(store, session_id):
    if session_id not in store:
        store[session_id] = DEFAULT_PROGRESS.copy()
    return store[session_id]


def _get_upload_dir(session_id):
    carpeta = os.path.join(UPLOAD_DIR, session_id)
    os.makedirs(carpeta, exist_ok=True)
    return carpeta


def _get_output_dir(session_id):
    carpeta = os.path.join(OUTPUT_DIR, session_id)
    os.makedirs(carpeta, exist_ok=True)
    return carpeta


def _get_uploaded_file(session_id):
    return archivo_cargado.get(session_id)


def _validate_file(archivo):
    filename = secure_filename(archivo.filename or "")
    if not filename:
        return False, "Nombre de archivo inválido."

    if "." not in filename or filename.rsplit(".", 1)[1].lower() not in ALLOWED_EXTENSIONS:
        return False, "Extensión de archivo no permitida. Solo se aceptan .xls o .xlsx."

    archivo.stream.seek(0, os.SEEK_END)
    size = archivo.stream.tell()
    archivo.stream.seek(0)
    if size > MAX_UPLOAD_SIZE:
        return False, "El archivo excede el tamaño máximo permitido de 10 MB."

    return True, filename


def formulario():
    session_id = _get_session_id()
    progreso_actual_kardex = _get_progress(progreso_kardex, session_id)
    progreso_actual_remision = _get_progress(progreso_remision, session_id)
    return render_template("formulario.html",
                           progreso_kardex=progreso_actual_kardex,
                           progreso_remision=progreso_actual_remision)


def cargar_archivo():
    archivo = request.files.get('archivo_excel')
    if not archivo:
        return jsonify({"error": "No se seleccionó archivo"}), 400

    valido, filename = _validate_file(archivo)
    if not valido:
        return jsonify({"error": filename}), 400

    session_id = _get_session_id()
    carpeta_sesion = _get_upload_dir(session_id)
    extension = filename.rsplit(".", 1)[1].lower()
    nombre_final = f"{uuid4().hex}.{extension}"
    ruta = os.path.join(carpeta_sesion, nombre_final)

    archivo.save(ruta)
    archivo_cargado[session_id] = ruta
    progreso_kardex[session_id] = DEFAULT_PROGRESS.copy()
    progreso_remision[session_id] = DEFAULT_PROGRESS.copy()
    return jsonify({"status": "cargado"})


def iniciar_kardex(session_id, ruta_archivo, carpeta_salida):
    progreso_actual = _get_progress(progreso_kardex, session_id)
    if not ruta_archivo:
        progreso_actual.update({"completado": False})
        return
    df = pd.read_excel(ruta_archivo)
    progreso_actual.update({"total": len(df), "procesadas": 0, "completado": False, "pdf_final": ""})
    pdf, _ = generar_kardex_consolidado(df, progreso_actual, carpeta_salida)
    progreso_actual.update({"pdf_final": pdf, "completado": True})


def iniciar_remision(session_id, ruta_archivo, carpeta_salida):
    progreso_actual = _get_progress(progreso_remision, session_id)
    if not ruta_archivo:
        progreso_actual.update({"completado": False})
        return
    df = pd.read_excel(ruta_archivo)
    progreso_actual.update({"total": len(df), "procesadas": 0, "completado": False, "pdf_final": ""})
    pdf, _ = generar_remision_consolidado(df, progreso_actual, carpeta_salida)
    progreso_actual.update({"pdf_final": pdf, "completado": True})


def procesar_kardex():
    session_id = _get_session_id()
    ruta_archivo = _get_uploaded_file(session_id)
    if not ruta_archivo or not os.path.exists(ruta_archivo):
        return jsonify({"error": "No hay un archivo cargado para procesar"}), 400
    carpeta_salida = _get_output_dir(session_id)
    threading.Thread(target=iniciar_kardex, args=(session_id, ruta_archivo, carpeta_salida), daemon=True).start()
    return jsonify({"status": "kardex_iniciado"})


def procesar_remision():
    session_id = _get_session_id()
    ruta_archivo = _get_uploaded_file(session_id)
    if not ruta_archivo or not os.path.exists(ruta_archivo):
        return jsonify({"error": "No hay un archivo cargado para procesar"}), 400
    carpeta_salida = _get_output_dir(session_id)
    threading.Thread(target=iniciar_remision, args=(session_id, ruta_archivo, carpeta_salida), daemon=True).start()
    return jsonify({"status": "remision_iniciado"})


def progreso_kardex_view():
    session_id = _get_session_id()
    return jsonify(_get_progress(progreso_kardex, session_id))


def progreso_remision_view():
    session_id = _get_session_id()
    return jsonify(_get_progress(progreso_remision, session_id))


def descargar_kardex():
    session_id = _get_session_id()
    progreso_actual = _get_progress(progreso_kardex, session_id)
    if not progreso_actual["completado"] or not os.path.exists(progreso_actual["pdf_final"]):
        return jsonify({"error": "PDF no disponible"}), 404
    return send_file(progreso_actual["pdf_final"], as_attachment=True)


def descargar_remision():
    session_id = _get_session_id()
    progreso_actual = _get_progress(progreso_remision, session_id)
    if not progreso_actual["completado"] or not os.path.exists(progreso_actual["pdf_final"]):
        return jsonify({"error": "PDF no disponible"}), 404
    return send_file(progreso_actual["pdf_final"], as_attachment=True)
