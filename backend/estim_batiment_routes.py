# estim_batiment_routes.py
from flask import Blueprint, request, jsonify, send_file
from estim_engine import process_estim_batiment

# Création d'un "Blueprint" pour regrouper les routes liées à l'estimation
estim_batiment_bp = Blueprint('estim_batiment', __name__)

@estim_batiment_bp.route('/estim-batiment', methods=['POST'])
def estim_batiment_route():
    """
    Endpoint pour traiter les fichiers Excel d'estimation de bâtiment.
    """
    if 'excel_file' not in request.files:
        return jsonify({"error": "Fichier Excel requis avec la clé 'excel_file'."}), 400

    uploaded_file = request.files['excel_file']
    
    if not uploaded_file or not uploaded_file.filename:
        return jsonify({"error": "Fichier Excel valide requis."}), 400

    print(f"Fichier reçu pour estimation bâtiment: {uploaded_file.filename}")

    try:
        file_bytes = uploaded_file.read()
        if not file_bytes:
            return jsonify({"error": "Le fichier envoyé est vide."}), 400

        # Appel de la logique métier centralisée
        output_excel_io, output_filename = process_estim_batiment(file_bytes)
        
        if output_excel_io and output_filename:
            print(f"Envoi du fichier d'estimation: {output_filename}")
            return send_file(
                output_excel_io,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=output_filename
            )
        else:
            error_message = output_filename or "Erreur inconnue lors du traitement"
            return jsonify({"error": error_message}), 500
            
    except Exception as e:
        print(f"Exception critique dans la route /estim-batiment: {e}")
        return jsonify({"error": f"Erreur serveur critique: {str(e)}"}), 500
