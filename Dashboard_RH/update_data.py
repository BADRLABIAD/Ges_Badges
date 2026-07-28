#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================
SCRIPT DE MISE À JOUR DES DONNÉES RH
============================================

Ce script convertit le fichier Excel TDB_2026.xlsx en fichier JSON
pour alimenter le Dashboard RH.

UTILISATION:
    python update_data.py [chemin_fichier_excel]

EXEMPLE:
    python update_data.py TDB_2026.xlsx

STRUCTURE ATTENDUE DU FICHIER EXCEL:
    - Feuille "EFFECTIF": Liste des collaborateurs
    - Feuille "MOUVEMENT": Recrutements et départs
    - Feuille "MS": Masse salariale

SORTIE:
    - data/data.json: Fichier JSON avec toutes les données
"""

import pandas as pd
import json
import sys
import os
from datetime import datetime

def parse_date(date_value):
    """Convertit une valeur de date en string ISO format"""
    if pd.isna(date_value):
        return None
    if isinstance(date_value, str):
        return date_value
    try:
        return pd.to_datetime(date_value).strftime('%Y-%m-%d')
    except:
        return None

def calculate_age(birth_date):
    """Calcule l'âge à partir de la date de naissance"""
    if pd.isna(birth_date):
        return 0
    try:
        birth = pd.to_datetime(birth_date)
        today = datetime.now()
        return int((today - birth).days / 365.25)
    except:
        return 0

def calculate_anciennete(entry_date):
    """Calcule l'ancienneté à partir de la date d'entrée"""
    if pd.isna(entry_date):
        return 0
    try:
        entry = pd.to_datetime(entry_date)
        today = datetime.now()
        return int((today - entry).days / 365.25)
    except:
        return 0

def extract_effectif(df):
    """Extrait les données des effectifs"""
    effectif_data = []
    
    for _, row in df.iterrows():
        effectif_data.append({
            'matricule': str(row.get('MATRICULE', '')),
            'nom': str(row.get('NOM', '')),
            'prenom': str(row.get('PRENOM', '')),
            'etb': str(row.get('ETB_LIB', '')).strip(),
            'classification': str(row.get('CLASSIFICATION_LIB', '')).strip(),
            'fonction': str(row.get('FONCTION_LIB', '')),
            'age': calculate_age(row.get('DT_NAISSANCE')),
            'anciennete': calculate_anciennete(row.get('DEB_ADMINISTRATION')),
            'contrat': str(row.get('TYPE_CONTRAT', '')),
            'civilite': str(row.get('CIVILITE', '')),
            'dateNaissance': parse_date(row.get('DT_NAISSANCE')),
            'dateEntree': parse_date(row.get('DEB_ADMINISTRATION')),
            'situationFamiliale': str(row.get('SITUAT_FAM', '')) if pd.notna(row.get('SITUAT_FAM')) else None
        })
    
    return effectif_data

def extract_masse_salariale(df):
    """Extrait les données de masse salariale"""
    ms_data = []
    
    months = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 
              'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
    months_keys = ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin',
                   'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre']
    
    for _, row in df.iterrows():
        ms_entry = {
            'compte': str(row.get('COMPTE_IMPUTATION', '')),
            'intitule': str(row.get('INTITULE_COMPLET', '')),
            'etb': str(row.get('etb_name', '')),
            'classification': str(row.get('Classification', ''))
        }
        
        for i, month in enumerate(months):
            value = row.get(month, 0)
            ms_entry[months_keys[i]] = float(value) if pd.notna(value) else 0
        
        ms_data.append(ms_entry)
    
    return ms_data

def extract_mouvements(df_raw):
    """Extrait les recrutements et départs depuis la feuille MOUVEMENT"""
    recrutements = []
    departs = []
    
    # Convertir en liste pour parsing
    data = df_raw.values.tolist()
    
    mode = None  # 'recrutement' ou 'depart'
    
    for row in data:
        row_str = ' '.join([str(x) for x in row if pd.notna(x)])
        
        if 'recrutement' in row_str.lower():
            mode = 'recrutement'
            continue
        elif 'départ' in row_str.lower():
            mode = 'depart'
            continue
        elif 'Total' in row_str:
            continue
        
        # Vérifier si c'est une ligne de données (commence par un matricule numérique)
        if mode and len(row) >= 6:
            matricule = row[1] if len(row) > 1 else None
            if pd.notna(matricule) and str(matricule).replace('.0', '').isdigit():
                try:
                    if mode == 'recrutement':
                        recrutements.append({
                            'matricule': str(int(float(str(matricule).replace('.0', '')))),
                            'nom': str(row[2]) if pd.notna(row[2]) else '',
                            'prenom': str(row[3]) if pd.notna(row[3]) else '',
                            'site': str(row[4]) if pd.notna(row[4]) else '',
                            'college': str(row[5]) if pd.notna(row[5]) else '',
                            'date': parse_date(row[6]) if len(row) > 6 else None,
                            'fonction': str(row[7]) if len(row) > 7 and pd.notna(row[7]) else '',
                            'mois': 'Janvier'  # À ajuster selon le mois
                        })
                    elif mode == 'depart':
                        departs.append({
                            'matricule': str(int(float(str(matricule).replace('.0', '')))),
                            'nom': str(row[2]) if pd.notna(row[2]) else '',
                            'prenom': str(row[3]) if pd.notna(row[3]) else '',
                            'site': str(row[4]) if pd.notna(row[4]) else '',
                            'college': str(row[5]) if pd.notna(row[5]) else '',
                            'date': parse_date(row[6]) if len(row) > 6 and pd.notna(row[6]) else None,
                            'motif': str(row[7]) if len(row) > 7 and pd.notna(row[7]) else 'Non précisé',
                            'mois': 'Janvier'
                        })
                except Exception as e:
                    print(f"  ⚠ Erreur parsing ligne: {e}")
                    continue
    
    return recrutements, departs

def detect_current_month(df_ms):
    """Détecte le mois courant basé sur les données MS non nulles"""
    months = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 
              'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
    
    for i, month in enumerate(reversed(months)):
        if month in df_ms.columns:
            total = df_ms[month].sum()
            if total > 0:
                return month, len(months) - i
    
    return 'Janvier', 1

def main(excel_path):
    """Fonction principale"""
    print("=" * 60)
    print("  MISE À JOUR DES DONNÉES DASHBOARD RH")
    print("=" * 60)
    print(f"\n📂 Fichier source: {excel_path}")
    
    # Vérifier le fichier
    if not os.path.exists(excel_path):
        print(f"❌ ERREUR: Le fichier '{excel_path}' n'existe pas!")
        sys.exit(1)
    
    # Charger les données
    print("\n🔄 Chargement des données...")
    try:
        xlsx = pd.ExcelFile(excel_path)
        print(f"   Feuilles trouvées: {xlsx.sheet_names}")
    except Exception as e:
        print(f"❌ ERREUR: Impossible de lire le fichier Excel: {e}")
        sys.exit(1)
    
    # Extraire les données de chaque feuille
    print("\n📊 Extraction des données...")
    
    # EFFECTIF
    if 'EFFECTIF' in xlsx.sheet_names:
        df_effectif = pd.read_excel(excel_path, sheet_name='EFFECTIF')
        effectif_data = extract_effectif(df_effectif)
        print(f"   ✓ Effectif: {len(effectif_data)} collaborateurs")
    else:
        print("   ⚠ Feuille EFFECTIF non trouvée!")
        effectif_data = []
    
    # MASSE SALARIALE
    if 'MS' in xlsx.sheet_names:
        df_ms = pd.read_excel(excel_path, sheet_name='MS')
        ms_data = extract_masse_salariale(df_ms)
        mois_courant, mois_actif = detect_current_month(df_ms)
        print(f"   ✓ Masse Salariale: {len(ms_data)} lignes")
        print(f"   ✓ Mois courant détecté: {mois_courant}")
    else:
        print("   ⚠ Feuille MS non trouvée!")
        ms_data = []
        mois_courant, mois_actif = 'Janvier', 1
    
    # MOUVEMENTS
    if 'MOUVEMENT' in xlsx.sheet_names:
        df_mouvement = pd.read_excel(excel_path, sheet_name='MOUVEMENT', header=None)
        recrutements, departs = extract_mouvements(df_mouvement)
        print(f"   ✓ Recrutements: {len(recrutements)}")
        print(f"   ✓ Départs: {len(departs)}")
    else:
        print("   ⚠ Feuille MOUVEMENT non trouvée!")
        recrutements, departs = [], []
    
    # Configuration
    config = {
        "annee": 2026,
        "moisCourant": mois_courant,
        "moisActif": mois_actif,
        "devise": "MAD",
        "entreprise": "Votre Entreprise",
        "lastUpdate": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    # Assembler les données
    data = {
        "config": config,
        "effectif": effectif_data,
        "masseSalariale": ms_data,
        "recrutements": recrutements,
        "departs": departs
    }
    
    # Sauvegarder le JSON
    output_path = 'data/data.json'
    os.makedirs('data', exist_ok=True)
    
    print(f"\n💾 Sauvegarde vers {output_path}...")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print("\n" + "=" * 60)
    print("  ✅ MISE À JOUR TERMINÉE AVEC SUCCÈS!")
    print("=" * 60)
    print(f"\n📈 Résumé:")
    print(f"   - Effectif: {len(effectif_data)} collaborateurs")
    print(f"   - Masse Salariale: {len(ms_data)} lignes")
    print(f"   - Recrutements: {len(recrutements)}")
    print(f"   - Départs: {len(departs)}")
    print(f"   - Mois actif: {mois_courant} {config['annee']}")
    print(f"\n📁 Fichier généré: {os.path.abspath(output_path)}")
    print("\n💡 Le dashboard sera mis à jour automatiquement lors du")
    print("   prochain rafraîchissement de la page.")

if __name__ == '__main__':
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excel_file = 'TDB_2026.xlsx'
    
    main(excel_file)
