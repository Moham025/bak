# app.py
from flask import Flask, jsonify, request
from flask_cors import CORS
import os

# Import des "Blueprints" (groupes de routes) depuis les nouveaux fichiers
from estim_batiment_routes import estim_batiment_bp
from armature_routes import armature_bp
from utility_routes import utility_bp

# --- Configuration de l'application Flask ---
app = Flask(__name__)
CORS(app)

# --- Enregistrement des Blueprints ---
# Chaque blueprint contient les routes pour une fonctionnalité spécifique
app.register_blueprint(estim_batiment_bp)
app.register_blueprint(armature_bp)
app.register_blueprint(utility_bp)

# --- Route Principale (Health Check) ---
@app.route('/')
def health_check():
    """
    Route principale pour vérifier que le service est en ligne.
    Affiche une page HTML simple pour les navigateurs.
    """
    user_agent = request.headers.get('User-Agent', '').lower()
    
    if 'go-http-client' in user_agent or request.headers.get('Accept', '').startswith('application/json'):
        return jsonify({"status": "healthy", "service": "Excel Processing API"}), 200
    
    html = """
    <!DOCTYPE html>
    <html>
    <head><title>Excel Processing API</title></head>
    <body>
        <h1>Excel Processing API</h1>
        <p>Status: <strong style="color: green;">Healthy</strong></p>
        <p>Les routes sont maintenant organisées en Blueprints.</p>
    </body>
    </html>
    """
    return html, 200, {'Content-Type': 'text/html'}

# --- Lancement de l'application ---
if __name__ == '__main__':
    # Configuration pour le déploiement
    port = int(os.environ.get('PORT', 5000))
    host = '0.0.0.0'
    # Le mode debug ne devrait pas être activé en production
    debug = os.environ.get('FLASK_ENV') != 'production'
    
    print(f"Démarrage du serveur Flask sur {host}:{port}")
    app.run(debug=debug, host=host, port=port)
