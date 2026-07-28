#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de génération des données JSON pour le Dashboard RH
==========================================================

Point d'entrée utilisé pour la mise à jour mensuelle (voir README).
La logique de calcul (EFFECTIF/MS/MOUVEMENT/FORMATION) vit dans
generate_data.py à la racine du projet : ce script ne fait que l'appeler,
pour que le cron mensuel et le serveur Flask restent alignés.

INSTRUCTIONS:
1. Déposer le fichier Excel dans le dossier 'data/' (TDB_COURANT.xlsx)
   et le fichier de formation (FORMATION.xlsx) le cas échéant.
2. Exécuter ce script: python3 scripts/generate_data.py
3. Le fichier data/data.json sera mis à jour automatiquement
"""

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT_DIR))

from generate_data import generate_data

if __name__ == '__main__':
    excel_arg = sys.argv[1] if len(sys.argv) > 1 else None
    success = generate_data(excel_path=excel_arg)
    sys.exit(0 if success else 1)
