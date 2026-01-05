import os
from flask import Flask
from routes.main_routes import main
from routes.formulario_routes import formulario_bp

ENVIRONMENT = os.getenv("FLASK_ENV", os.getenv("ENV", "development")).lower()
IS_PRODUCTION = ENVIRONMENT == "production"

secret_key = os.getenv("SECRET_KEY")
if IS_PRODUCTION and not secret_key:
    raise RuntimeError("SECRET_KEY debe estar configurada en el entorno de producción.")

app = Flask(__name__)
app.secret_key = secret_key or "development-secret-key"  # Necesario para mensajes flash

# Registrar rutas
app.register_blueprint(main)
app.register_blueprint(formulario_bp)

if __name__ == '__main__':
    # Ya no verificamos ni cargamos datos en SQLite,
    # porque toda la información se procesa directamente desde Excel
    print("Aplicación iniciada. Usando Excel como fuente principal de datos.")
    app.run(debug=not IS_PRODUCTION, use_reloader=False)
