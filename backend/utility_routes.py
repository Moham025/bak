# utility_routes.py
from flask import Blueprint, request, jsonify, send_file
import openpyxl
import io
from openpyxl.utils.exceptions import InvalidFileException

# Note: Les fonctions comme 'traiter_fichier_excel_core' ont été déplacées ici
# pour garder ce module autonome.

utility_bp = Blueprint('utility', __name__)

@utility_bp.route('/get-sheet-names', methods=['POST'])
def get_sheet_names_route():
    if 'excel_file' not in request.files:
        return jsonify({"error": "Aucun fichier ('excel_file') envoyé"}), 400
    file = request.files['excel_file']
    if not file.filename:
        return jsonify({"error": "Aucun fichier sélectionné"}), 400
    
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(file.read()), read_only=True)
        return jsonify({"sheet_names": workbook.sheetnames}), 200
    except InvalidFileException:
        return jsonify({"error": "Fichier Excel invalide ou corrompu."}), 400
    except Exception as e:
        return jsonify({"error": f"Erreur serveur: {str(e)}"}), 500

# Vous pouvez ajouter ici la route /process-excel et sa logique si nécessaire.
# Pour l'instant, elle est omise pour se concentrer sur la nouvelle architecture.
