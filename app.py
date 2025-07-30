from flask import Flask
from routes.main_routes import main
from routes.formulario_routes import formulario_bp

app = Flask(__name__)
app.secret_key = "supersecret"  # Necesario para mensajes flash

# Registrar rutas
app.register_blueprint(main)
app.register_blueprint(formulario_bp)

if __name__ == '__main__':
    # Ya no verificamos ni cargamos datos en SQLite,
    # porque toda la información se procesa directamente desde Excel
    print("Aplicación iniciada. Usando Excel como fuente principal de datos.")
    app.run(debug=True, use_reloader=False)