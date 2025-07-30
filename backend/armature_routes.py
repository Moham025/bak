# armature_routes.py
from flask import Blueprint, request, jsonify, send_file
from combineArm import process_armature_csvs # On suppose que cette fonction est dans combineArm.py

armature_bp = Blueprint('armature', __name__)

@armature_bp.route('/combine-armatures', methods=['POST'])
def combine_armatures_route():
    if not request.files:
        return jsonify({"error": "Aucun fichier envoyé."}), 400

    uploaded_files = request.files.getlist('csv_files')
    if not uploaded_files:
        return jsonify({"error": "Aucun fichier trouvé avec la clé 'csv_files'."}), 400

    files_data = [{'name': f.filename, 'bytes': f.read()} for f in uploaded_files if f and f.filename]
    
    if not files_data:
        return jsonify({"error": "Aucuns fichiers valides trouvés dans la requête."}), 400

    try:
        output_excel_io, output_filename = process_armature_csvs(files_data)
        
        if output_excel_io and output_filename:
            return send_file(
                output_excel_io,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=output_filename
            )
        else:
            error_message = output_filename or "Erreur lors de la combinaison"
            return jsonify({"error": error_message}), 500
    except Exception as e:
        return jsonify({"error": f"Erreur serveur critique: {str(e)}"}), 500
