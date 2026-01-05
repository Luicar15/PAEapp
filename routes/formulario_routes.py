from flask import Blueprint
from controllers.formulario_controller import (
    formulario, cargar_archivo, procesar_kardex, procesar_remision,
    progreso_kardex_view, progreso_remision_view,
    descargar_kardex, descargar_remision
)
from utils.auth import require_auth

formulario_bp = Blueprint("formulario", __name__)

@formulario_bp.route("/formulario", methods=["GET"])
@require_auth
def formulario_view():
    return formulario()

@formulario_bp.route("/formulario/cargar", methods=["POST"])
@require_auth
def cargar_archivo_view():
    return cargar_archivo()

@formulario_bp.route("/formulario/procesar_kardex", methods=["POST"])
@require_auth
def procesar_kardex_view():
    return procesar_kardex()

@formulario_bp.route("/formulario/procesar_remision", methods=["POST"])
@require_auth
def procesar_remision_view():
    return procesar_remision()

@formulario_bp.route("/formulario/progreso_kardex", methods=["GET"])
@require_auth
def progreso_kardex_route():
    return progreso_kardex_view()

@formulario_bp.route("/formulario/progreso_remision", methods=["GET"])
@require_auth
def progreso_remision_route():
    return progreso_remision_view()

@formulario_bp.route("/formulario/descargar_kardex", methods=["GET"])
@require_auth
def descargar_kardex_route():
    return descargar_kardex()

@formulario_bp.route("/formulario/descargar_remision", methods=["GET"])
@require_auth
def descargar_remision_route():
    return descargar_remision()
