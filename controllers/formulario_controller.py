import os
import threading
import pandas as pd
from flask import render_template, request, jsonify, send_file
from utils.generar_kardex import generar_kardex_consolidado
from utils.generar_remision import generar_remision_consolidado

OUTPUT_DIR = os.path.join("data", "outputs")
UPLOAD_DIR = os.path.join("data", "uploads")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Variables de progreso separadas
progreso_kardex = {"procesadas": 0, "total": 0, "completado": False, "pdf_final": ""}
progreso_remision = {"procesadas": 0, "total": 0, "completado": False, "pdf_final": ""}

archivo_cargado = {"ruta": ""}  # Excel cargado solo una vez

def formulario():
    return render_template("formulario.html",
                           progreso_kardex=progreso_kardex,
                           progreso_remision=progreso_remision)

def cargar_archivo():
    archivo = request.files.get('archivo_excel')
    if not archivo:
        return jsonify({"error": "No se seleccionó archivo"}), 400
    ruta = os.path.join(UPLOAD_DIR, archivo.filename)
    archivo.save(ruta)
    archivo_cargado["ruta"] = ruta
    return jsonify({"status": "cargado"})

def iniciar_kardex():
    global progreso_kardex
    if not archivo_cargado["ruta"]:
        return
    df = pd.read_excel(archivo_cargado["ruta"])
    progreso_kardex.update({"total": len(df), "procesadas": 0, "completado": False, "pdf_final": ""})
    pdf, excels = generar_kardex_consolidado(df, progreso_kardex, OUTPUT_DIR)
    progreso_kardex.update({"pdf_final": pdf, "completado": True})

def iniciar_remision():
    global progreso_remision
    if not archivo_cargado["ruta"]:
        return
    df = pd.read_excel(archivo_cargado["ruta"])
    progreso_remision.update({"total": len(df), "procesadas": 0, "completado": False, "pdf_final": ""})
    pdf, excels = generar_remision_consolidado(df, progreso_remision, OUTPUT_DIR)
    progreso_remision.update({"pdf_final": pdf, "completado": True})

def procesar_kardex():
    threading.Thread(target=iniciar_kardex).start()
    return jsonify({"status": "kardex_iniciado"})

def procesar_remision():
    threading.Thread(target=iniciar_remision).start()
    return jsonify({"status": "remision_iniciado"})

def progreso_kardex_view():
    return jsonify(progreso_kardex)

def progreso_remision_view():
    return jsonify(progreso_remision)

def descargar_kardex():
    if not progreso_kardex["completado"] or not os.path.exists(progreso_kardex["pdf_final"]):
        return jsonify({"error": "PDF no disponible"}), 404
    return send_file(progreso_kardex["pdf_final"], as_attachment=True)

def descargar_remision():
    if not progreso_remision["completado"] or not os.path.exists(progreso_remision["pdf_final"]):
        return jsonify({"error": "PDF no disponible"}), 404
    return send_file(progreso_remision["pdf_final"], as_attachment=True)