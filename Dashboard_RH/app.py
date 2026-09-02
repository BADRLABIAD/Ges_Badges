#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==========================================================================
    DASHBOARD RH 2026 - Serveur Web HTTPS
==========================================================================
    Application Flask pour servir le Dashboard RH
    Serveur: 10.222.36.9
    Port: 5001
    URL: https://dashboardrh.sonasid.ma:5001 ou https://10.222.36.9:5001
==========================================================================
"""

from flask import Flask, render_template, send_from_directory, jsonify, request
import os
import sys
import json
from datetime import datetime
from pathlib import Path

# Configuration
APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / 'data'
cert_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cert.pem')
PORT = 5001
HOST = '0.0.0.0'

# La génération des données (EFFECTIF/MS/MOUVEMENT + FORMATION) vit dans un seul
# endroit : generate_data.py. On l'importe ici pour que /api/refresh et
# /api/upload restent alignés avec le script de mise à jour mensuelle.
sys.path.insert(0, str(APP_DIR))
from generate_data import generate_data as run_generate_data

# Initialisation Flask
app = Flask(__name__, 
            static_folder='static',
            template_folder='templates')

app.config['SECRET_KEY'] = 'dashboard-rh-2026-secret-key'
app.config['JSON_AS_ASCII'] = False

# ============================================================================
# ROUTES PRINCIPALES
# ============================================================================

@app.route('/')
def index():
    """Page principale du dashboard"""
    return send_from_directory('.', 'index.html')

@app.route('/data/<path:filename>')
def serve_data(filename):
    """Servir les fichiers de données JSON"""
    return send_from_directory('data', filename)

@app.route('/api/data')
def get_data():
    """API pour récupérer les données JSON"""
    try:
        data_file = DATA_DIR / 'data.json'
        if data_file.exists():
            with open(data_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return jsonify(data)
        else:
            return jsonify({'error': 'Fichier data.json non trouvé'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/refresh', methods=['POST'])
def refresh_data():
    """API pour régénérer les données depuis Excel"""
    try:
        result = generate_data_from_excel()
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e), 'success': False}), 500

@app.route('/api/upload', methods=['POST'])
def upload_excel():
    """API pour uploader un nouveau fichier Excel"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Aucun fichier fourni', 'success': False}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Aucun fichier sélectionné', 'success': False}), 400
        
        if file and file.filename.endswith(('.xlsx', '.xls')):
            # Sauvegarder le fichier
            filepath = DATA_DIR / 'TDB_COURANT.xlsx'
            file.save(filepath)
            
            # Régénérer les données
            result = generate_data_from_excel()
            result['message'] = f'Fichier {file.filename} uploadé et traité avec succès'
            return jsonify(result)
        else:
            return jsonify({'error': 'Format de fichier non supporté. Utilisez .xlsx ou .xls', 'success': False}), 400
            
    except Exception as e:
        return jsonify({'error': str(e), 'success': False}), 500

@app.route('/api/status')
def status():
    """API pour vérifier le statut du serveur"""
    data_file = DATA_DIR / 'data.json'
    excel_file = DATA_DIR / 'TDB_COURANT.xlsx'
    formation_file = DATA_DIR / 'FORMATION.xlsx'
    key_file = APP_DIR / 'key.pem'

    status_info = {
        'server': 'running',
        'timestamp': datetime.now().isoformat(),
        'data_file_exists': data_file.exists(),
        'excel_file_exists': excel_file.exists(),
        'formation_file_exists': formation_file.exists(),
        'ssl_enabled': Path(cert_file).exists() and key_file.exists()
    }

    if data_file.exists():
        with open(data_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            status_info['last_update'] = data.get('meta', {}).get('generated_at', 'N/A')
            status_info['total_fte'] = data.get('effectif', {}).get('total_fte', 0)

    return jsonify(status_info)

# ============================================================================
# BUDGETS (partagés entre tous les postes, stockés côté serveur)
# ============================================================================
BUDGET_FILE = DATA_DIR / 'budgets.json'
BUDGET_EDIT_PASSWORD = 'Sonasid2026*'
BUDGET_KEYS = {'masse_salariale', 'fte', 'formation', 'social', 'societal'}

@app.route('/api/budgets')
def get_budgets():
    """Lire les budgets alloués (cibles saisies par un administrateur)"""
    if not BUDGET_FILE.exists():
        return jsonify({})
    try:
        with open(BUDGET_FILE, 'r', encoding='utf-8') as f:
            return jsonify(json.load(f))
    except Exception:
        return jsonify({})

@app.route('/api/budgets', methods=['POST'])
def save_budgets():
    """Enregistrer les budgets alloués. Protégé par mot de passe côté serveur
    (pas seulement côté navigateur) pour que /api/budgets ne soit pas
    modifiable par n'importe qui connaissant juste l'URL."""
    payload = request.get_json(silent=True) or {}
    if payload.get('password') != BUDGET_EDIT_PASSWORD:
        return jsonify({'success': False, 'error': 'Mot de passe incorrect'}), 403

    values = payload.get('budgets', {})
    if not isinstance(values, dict):
        return jsonify({'success': False, 'error': 'Format invalide'}), 400

    cleaned = {}
    for key, value in values.items():
        if key not in BUDGET_KEYS:
            continue
        if value is None or value == '':
            continue
        try:
            cleaned[key] = float(value)
        except (TypeError, ValueError):
            continue

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(BUDGET_FILE, 'w', encoding='utf-8') as f:
        json.dump(cleaned, f, ensure_ascii=False, indent=2)

    return jsonify({'success': True, 'budgets': cleaned})

# ============================================================================
# GÉNÉRATION DES DONNÉES
# ============================================================================
# La logique de calcul (EFFECTIF/MS/MOUVEMENT/FORMATION) est centralisée dans
# generate_data.py pour que le serveur web et le script de mise à jour
# mensuelle produisent toujours le même data.json.

def generate_data_from_excel():
    """Régénérer data.json depuis TDB_COURANT.xlsx (+ FORMATION.xlsx si présent)"""
    excel_file = DATA_DIR / 'TDB_COURANT.xlsx'
    output_file = DATA_DIR / 'data.json'

    if not excel_file.exists():
        return {'success': False, 'error': f'Fichier Excel non trouvé: {excel_file}'}

    try:
        run_generate_data(excel_path=excel_file, output_path=output_file)
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        return {
            'success': True,
            'message': 'Données générées avec succès',
            'timestamp': datetime.now().isoformat(),
            'stats': {
                'total_fte': data['effectif']['total_fte'],
                'masse_salariale_ytd': data['masse_salariale']['ytd'],
                'recrutements': data['mouvements']['nb_recrutements'],
                'departs': data['mouvements']['nb_departs'],
                'formation_actions': data['formation']['kpis']['nb_actions'] if data.get('formation') else 0
            }
        }

    except Exception as e:
        return {'success': False, 'error': str(e)}

# ============================================================================
# DÉMARRAGE DU SERVEUR
# ============================================================================

def main():
    """Fonction principale de démarrage"""
    if os.path.exists(cert_file):
        print("\n" + "=" * 70)
        print("🚀 Démarrage du serveur Dashboard RH 2026")
        print("=" * 70)
        print(f"\n✅ Certificat trouvé: {cert_file}")
        print("\n🔒 HTTPS ACTIVÉ - Démarrage du serveur...")
        print("\n" + "=" * 70)
        print("📍 Accédez à l'application:")
        print("   🔐 https://dashboardrh.sonasid.ma:5001")
        print("   ou")
        print("   🔐 https://10.222.36.9:5001")
        print("\n" + "=" * 70 + "\n")
        
        try:
            # Utiliser le certificat PEM
            app.run(
                debug=True,
                host='0.0.0.0',
                port=5001,
                ssl_context=(cert_file, cert_file)
            )
        except Exception as e:
            print(f"\n❌ Erreur SSL: {e}")
            print("Redémarrage en HTTP...")
            app.run(debug=True, host='0.0.0.0', port=5001)
    else:
        print("\n" + "=" * 70)
        print("❌ Erreur: Fichier cert.pem non trouvé!")
        print("=" * 70)
        print(f"\n📁 Fichier attendu: {os.path.abspath(cert_file)}")
        print("\n📋 Solutions:")
        print("  1. Créez cert.pem avec: python generate_cert.py")
        print("  2. Vérifiez que cert.pem est dans le dossier courant")
        print("\n🌐 Démarrage en HTTP (non sécurisé)...")
        print("=" * 70 + "\n")
        app.run(debug=True, host='0.0.0.0', port=5001)

if __name__ == '__main__':
    main()
