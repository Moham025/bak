#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de test pour l'endpoint EstimBatiment
Teste l'upload d'un fichier Excel et la génération du devis
"""

import requests
import os
import sys

def test_estim_batiment_endpoint():
    """
    Teste l'endpoint /estim-batiment avec le fichier EstimType.xlsx
    """
    
    # URL de l'endpoint
    url = "http://192.168.11.107:5000/estim-batiment"
    
    # Chemin vers le fichier de test
    test_file_path = "backend/EstimBatiment/EstimType.xlsx"
    
    # Vérifier que le fichier existe
    if not os.path.exists(test_file_path):
        print(f"❌ Erreur: Le fichier de test '{test_file_path}' n'existe pas.")
        return False
    
    print(f"📁 Fichier de test trouvé: {test_file_path}")
    print(f"📊 Taille du fichier: {os.path.getsize(test_file_path)} bytes")
    
    try:
        # Préparer le fichier pour l'upload
        with open(test_file_path, 'rb') as file:
            files = {'excel_file': (os.path.basename(test_file_path), file, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            
            print(f"🚀 Envoi de la requête à {url}...")
            
            # Envoyer la requête POST
            response = requests.post(url, files=files, timeout=60)
            
            print(f"📡 Réponse reçue - Status: {response.status_code}")
            
            if response.status_code == 200:
                # Succès - sauvegarder le fichier résultat
                output_filename = "test_estimation_result.xlsx"
                
                # Récupérer le nom du fichier depuis l'en-tête Content-Disposition
                content_disposition = response.headers.get('content-disposition', '')
                if 'filename=' in content_disposition:
                    import re
                    match = re.search(r'filename="?([^"]+)"?', content_disposition)
                    if match:
                        output_filename = match.group(1)
                
                with open(output_filename, 'wb') as output_file:
                    output_file.write(response.content)
                
                print(f"✅ Succès! Fichier résultat sauvegardé: {output_filename}")
                print(f"📊 Taille du fichier résultat: {len(response.content)} bytes")
                return True
                
            else:
                # Erreur
                print(f"❌ Erreur {response.status_code}")
                try:
                    error_data = response.json()
                    print(f"💬 Message d'erreur: {error_data.get('error', 'Erreur inconnue')}")
                except:
                    print(f"💬 Réponse brute: {response.text[:500]}...")
                return False
                
    except requests.exceptions.ConnectionError:
        print("❌ Erreur: Impossible de se connecter au serveur.")
        print("💡 Assurez-vous que le serveur Flask est démarré avec 'python main.py'")
        return False
    except requests.exceptions.Timeout:
        print("❌ Erreur: Timeout de la requête (60s)")
        return False
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        return False

def check_server_health():
    """
    Vérifie si le serveur Flask est accessible
    """
    try:
        response = requests.get("http://192.168.11.107:5000", timeout=5)
        print(f"✅ Serveur accessible - Status: {response.status_code}")
        return True
    except:
        print("❌ Serveur non accessible sur http://192.168.11.107:5000")
        return False

if __name__ == "__main__":
    print("🔧 Test de l'endpoint EstimBatiment")
    print("=" * 50)
    
    # Vérifier la santé du serveur
    print("1. Vérification de l'accessibilité du serveur...")
    if not check_server_health():
        print("\n💡 Pour démarrer le serveur:")
        print("   python main.py")
        sys.exit(1)
    
    print("\n2. Test de l'endpoint EstimBatiment...")
    success = test_estim_batiment_endpoint()
    
    if success:
        print("\n🎉 Test réussi!")
    else:
        print("\n💥 Test échoué!")
        sys.exit(1) 