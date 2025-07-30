from flask import Blueprint
from controllers.cargue_masivo_controller import cargar_excel_masivo

cargue_masivo_bp = Blueprint('cargue_masivo_bp', __name__)

cargue_masivo_bp.route('/cargue-masivo', methods=['GET', 'POST'])(cargar_excel_masivo)