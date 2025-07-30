from flask import Blueprint
from controllers import main_controller

main = Blueprint('main', __name__)

@main.route('/')
def index():
    return main_controller.index()

@main.route('/inventario')
def inventario():
    return main_controller.inventario()

@main.route('/formulario')
def formulario():
    return main_controller.formulario()

@main.route('/vista-preliminar')
def vista_preliminar():
    return main_controller.vista_preliminar()

@main.route('/orden-compra')
def orden_compra():
    return main_controller.orden_compra()

@main.route('/categorias')
def categorias():
    return main_controller.categorias()

@main.route('/unidades-medida')
def unidades_medida():
    return main_controller.unidades_medida()

@main.route('/instituciones')
def instituciones():
    return main_controller.instituciones()

@main.route('/sedes')
def sedes():
    return main_controller.sedes()

@main.route('/ciclos-menu')
def ciclos_menu():
    return main_controller.ciclos_menu()

@main.route('/menus-etarios')
def menus_etarios():
    return main_controller.menus_etarios()

@main.route('/parametros')
def parametros():
    return main_controller.parametros()

@main.route('/archivos-base')
def archivos_base():
    return main_controller.archivos_base()

@main.route('/usuarios')
def usuarios():
    return main_controller.usuarios()

@main.route('/formulario-inicial')
def formulario_inicial():
    return main_controller.formulario_inicial()

@main.route('/formulario-manual')
def formulario_manual():
    return main_controller.formulario_manual()