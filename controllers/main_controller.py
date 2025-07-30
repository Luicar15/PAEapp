from flask import render_template

def index():
    return render_template('index.html')

def inventario():
    return render_template('inventario.html')

def formulario():
    return render_template('formulario.html')

def vista_preliminar():
    return render_template('vista_preliminar.html')

def orden_compra():
    return render_template('orden_compra.html')

def categorias():
    return render_template('categorias.html')

def unidades_medida():
    return render_template('unidades_medida.html')

def instituciones():
    return render_template('instituciones.html')

def sedes():
    return render_template('sedes.html')

def ciclos_menu():
    return render_template('ciclos_menu.html')

def menus_etarios():
    return render_template('menus_etarios.html')

def parametros():
    return render_template('parametros.html')

def archivos_base():
    return render_template('archivos_base.html')

def usuarios():
    return render_template('usuarios.html')

def formulario_inicial():
    return render_template('formulario_inicial.html')

def formulario_manual():
    return render_template('formulario_manual.html')  # Lo crearemos después