#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard RH - Générateur de données v11
Génère TOUS les indicateurs nécessaires pour le dashboard
"""

import pandas as pd
import json
from datetime import datetime
from pathlib import Path
import numpy as np
import sys

class NumpyEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (np.integer, np.int64, np.int32)):
            return int(obj)
        elif isinstance(obj, (np.floating, np.float64, np.float32)):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif pd.isna(obj):
            return None
        return super().default(obj)

FORMATION_COLUMNS = [
    'n_action', 'domaine', 'type_action', 'thematique', 'date_debut', 'date_fin',
    'nb_jours', 'nb_heures', 'matricule', 'nom', 'prenom', 'classification', 'site',
    'departement', 'statut', 'cabinet', 'formateur_interne', 'cout_formation',
    'cout_logistique', 'eval_chaud', 'eval_formateur', 'taux_satisfaction',
    'eval_froid', 'resultat'
]

def process_formation(formation_path, fte_total=0, ms_ytd=0):
    """Traiter les données du volet Formation (fichier TB/FORMATION.xlsx)"""
    formation_path = Path(formation_path)
    if not formation_path.exists():
        return None

    df = pd.read_excel(formation_path, sheet_name=0)
    df.columns = FORMATION_COLUMNS[:len(df.columns)]

    # Les colonnes "action" ne sont renseignées que sur la 1ère ligne de chaque
    # formation (cellules fusionnées à l'export) : on les propage vers le bas.
    action_cols = ['n_action', 'domaine', 'type_action', 'thematique', 'date_debut',
                   'date_fin', 'nb_jours', 'cabinet', 'formateur_interne',
                   'cout_formation', 'cout_logistique']
    df[action_cols] = df[action_cols].ffill()
    df = df[df['n_action'].notna()].copy()

    df['date_debut'] = pd.to_datetime(df['date_debut'], errors='coerce')
    df['date_fin'] = pd.to_datetime(df['date_fin'], errors='coerce')
    df['nom'] = df['nom'].fillna('').astype(str).str.strip()
    df['prenom'] = df['prenom'].fillna('').astype(str).str.strip()
    for col in ['domaine', 'type_action', 'classification', 'site', 'departement', 'statut', 'eval_chaud']:
        df[col] = df[col].fillna('Non renseigné').astype(str).str.strip().replace('', 'Non renseigné')
    df['nb_heures'] = pd.to_numeric(df['nb_heures'], errors='coerce').fillna(0)

    actions_df = df.drop_duplicates('n_action').copy()

    nb_actions = int(df['n_action'].nunique())
    nb_participations = int(len(df))
    presents = int((df['statut'] == 'Présent').sum())
    absents = int((df['statut'] == 'Absent').sum())
    base_presence = presents + absents
    taux_presence = round(presents / base_presence * 100, 1) if base_presence > 0 else None

    participants_distincts = df[df['nom'] != ''][['nom', 'prenom']].drop_duplicates()
    nb_participants_distincts = int(len(participants_distincts))
    taux_couverture = round(nb_participants_distincts / fte_total * 100, 1) if fte_total > 0 else None

    heures_stagiaires = int(df['nb_heures'].sum())
    nb_jours_total = round(float(actions_df['nb_jours'].fillna(0).sum()), 1)
    cout_formation_total = round(float(actions_df['cout_formation'].fillna(0).sum()), 2)
    cout_logistique_total = round(float(actions_df['cout_logistique'].fillna(0).sum()), 2)
    cout_total = round(cout_formation_total + cout_logistique_total, 2)
    cout_moyen_participant = round(cout_total / nb_participations, 2) if nb_participations > 0 else 0
    taux_investissement_ms = round(cout_total / ms_ytd * 100, 3) if ms_ytd > 0 else None

    def group_stats(dim):
        g = df.groupby(dim, dropna=False)
        stats = {}
        for key, sub in g:
            stats[key] = {
                'nb_actions': int(sub['n_action'].nunique()),
                'nb_participations': int(len(sub)),
                'nb_heures': int(sub['nb_heures'].sum())
            }
        return stats

    par_domaine = group_stats('domaine')
    par_site = group_stats('site')
    par_departement = group_stats('departement')
    par_classification = group_stats('classification')
    par_type_action = group_stats('type_action')
    par_statut = {k: int(v) for k, v in df['statut'].value_counts().to_dict().items()}
    evaluation_chaud = {k: int(v) for k, v in df['eval_chaud'].value_counts().to_dict().items()}

    # Évolution mensuelle (nombre d'actions démarrées et heures-stagiaires par mois)
    actions_df['mois_key'] = actions_df['date_debut'].dt.to_period('M').astype(str)
    df['mois_key'] = df['date_debut'].dt.to_period('M').astype(str)
    mois_valides = sorted(set(actions_df['mois_key'].dropna()) | set(df['mois_key'].dropna()))
    nb_actions_par_mois = actions_df.groupby('mois_key')['n_action'].nunique().to_dict()
    heures_par_mois = df.groupby('mois_key')['nb_heures'].sum().to_dict()
    evolution_mensuelle = {
        'labels': mois_valides,
        'nb_actions': [int(nb_actions_par_mois.get(m, 0)) for m in mois_valides],
        'heures': [int(heures_par_mois.get(m, 0)) for m in mois_valides]
    }

    actions = []
    for _, row in actions_df.iterrows():
        nb_part_action = int(len(df[df['n_action'] == row['n_action']]))
        actions.append({
            'n_action': int(row['n_action']),
            'domaine': row['domaine'],
            'type_action': row['type_action'],
            'thematique': row['thematique'],
            'date_debut': row['date_debut'].strftime('%Y-%m-%d') if pd.notna(row['date_debut']) else '',
            'date_fin': row['date_fin'].strftime('%Y-%m-%d') if pd.notna(row['date_fin']) else '',
            'nb_jours': float(row['nb_jours']) if pd.notna(row['nb_jours']) else 0,
            'nb_heures': float(row['nb_heures']) if pd.notna(row['nb_heures']) else 0,
            'cabinet': str(row['cabinet']) if pd.notna(row['cabinet']) else '',
            'cout_formation': float(row['cout_formation']) if pd.notna(row['cout_formation']) else 0,
            'cout_logistique': float(row['cout_logistique']) if pd.notna(row['cout_logistique']) else 0,
            'nb_participants': nb_part_action
        })
    actions.sort(key=lambda a: a['date_debut'] or '', reverse=True)

    participations = []
    for _, row in df.iterrows():
        participations.append({
            'n_action': int(row['n_action']),
            'domaine': row['domaine'],
            'thematique': row['thematique'],
            'date_debut': row['date_debut'].strftime('%Y-%m-%d') if pd.notna(row['date_debut']) else '',
            'nom': row['nom'],
            'prenom': row['prenom'],
            'classification': row['classification'],
            'site': row['site'],
            'departement': row['departement'],
            'statut': row['statut'],
            'eval_chaud': row['eval_chaud']
        })

    return {
        'meta': {
            'source_file': formation_path.name,
            'periode_min': mois_valides[0] if mois_valides else None,
            'periode_max': mois_valides[-1] if mois_valides else None
        },
        'kpis': {
            'nb_actions': nb_actions,
            'nb_participations': nb_participations,
            'nb_participants_distincts': nb_participants_distincts,
            'taux_couverture': taux_couverture,
            'presents': presents,
            'absents': absents,
            'taux_presence': taux_presence,
            'nb_jours_total': nb_jours_total,
            'heures_stagiaires': heures_stagiaires,
            'cout_formation_total': cout_formation_total,
            'cout_logistique_total': cout_logistique_total,
            'cout_total': cout_total,
            'cout_moyen_participant': cout_moyen_participant,
            'taux_investissement_ms': taux_investissement_ms
        },
        'par_domaine': par_domaine,
        'par_site': par_site,
        'par_departement': par_departement,
        'par_classification': par_classification,
        'par_type_action': par_type_action,
        'par_statut': par_statut,
        'evaluation_chaud': evaluation_chaud,
        'evolution_mensuelle': evolution_mensuelle,
        'actions': actions,
        'participations': participations
    }

def build_social_placeholder():
    """Le volet Social attend un fichier de données dédié (absentéisme, AT/MP,
    discipline, dialogue social, œuvres sociales...). Tant qu'il n'est pas fourni,
    on expose un état vide explicite plutôt que d'inventer des indicateurs."""
    return {
        'available': False,
        'message': "Le fichier de données Social n'a pas encore été fourni. "
                   "Cette page affichera automatiquement les indicateurs sociaux "
                   "(absentéisme, santé-sécurité, discipline, dialogue social, "
                   "œuvres sociales...) dès sa réception."
    }

def generate_data(excel_path=None, output_path=None, formation_path=None):
    """Générer les données JSON depuis le fichier Excel"""

    # Chemins par défaut
    if excel_path is None:
        excel_path = Path(__file__).parent / 'data' / 'TDB_COURANT.xlsx'
    else:
        excel_path = Path(excel_path)

    if output_path is None:
        output_path = Path(__file__).parent / 'data' / 'data.json'
    else:
        output_path = Path(output_path)

    if formation_path is None:
        formation_path = Path(__file__).parent / 'data' / 'FORMATION.xlsx'
    else:
        formation_path = Path(formation_path)
    
    if not excel_path.exists():
        print(f"❌ Erreur: Fichier non trouvé: {excel_path}")
        return False
    
    print(f"📖 Lecture de {excel_path}...")
    
    # ============ CHARGEMENT DES DONNÉES ============
    try:
        df_effectif = pd.read_excel(excel_path, sheet_name='EFFECTIF')
        df_ms = pd.read_excel(excel_path, sheet_name='MS')
        try:
            df_mouvement = pd.read_excel(excel_path, sheet_name='MOUVEMENT')
        except:
            df_mouvement = pd.DataFrame()
            print("   ⚠️ Onglet MOUVEMENT non trouvé")
    except Exception as e:
        print(f"❌ Erreur lecture Excel: {e}")
        return False
    
    today = datetime.now()
    
    # ============ PÉRIODES DISPONIBLES ============
    periodes = []
    if 'ANNEE_PAIE' in df_effectif.columns and 'NO_PAIE' in df_effectif.columns:
        periodes_df = df_effectif.groupby(['ANNEE_PAIE', 'NO_PAIE']).size().reset_index(name='count')
        for _, row in periodes_df.iterrows():
            periodes.append({
                'annee': int(row['ANNEE_PAIE']),
                'mois': int(row['NO_PAIE']),
                'count': int(row['count'])
            })
        periodes.sort(key=lambda x: (x['annee'], x['mois']), reverse=True)
        derniere_annee = periodes[0]['annee']
        dernier_mois = periodes[0]['mois']
    else:
        derniere_annee = today.year
        dernier_mois = today.month
    
    print(f"📅 Périodes: {periodes}")
    print(f"📅 Période courante: {dernier_mois}/{derniere_annee}")
    
    # ============ TRAITEMENT EFFECTIF ============
    df_effectif['MDP'] = df_effectif['MDP'].fillna('').astype(str).str.lower().str.strip()
    
    # Filtrer par période courante
    df_periode = df_effectif[
        (df_effectif['ANNEE_PAIE'] == derniere_annee) & 
        (df_effectif['NO_PAIE'] == dernier_mois)
    ].copy()
    
    # Exclure les sortants (chèque)
    sortants_cheque = df_periode[df_periode['MDP'] == 'cheque'].copy()
    df_actif = df_periode[df_periode['MDP'] != 'cheque'].copy()
    
    print(f"👥 Période {dernier_mois}/{derniere_annee}: Total={len(df_periode)}, Sortants={len(sortants_cheque)}, Actifs={len(df_actif)}")
    
    # Calculer âge et ancienneté
    df_actif['DT_NAISSANCE'] = pd.to_datetime(df_actif['DT_NAISSANCE'], errors='coerce')
    df_actif['DEB_ADMINISTRATION'] = pd.to_datetime(df_actif['DEB_ADMINISTRATION'], errors='coerce')
    df_actif['Age'] = ((today - df_actif['DT_NAISSANCE']).dt.days / 365.25).fillna(0).astype(int)
    df_actif['Anciennete'] = ((today - df_actif['DEB_ADMINISTRATION']).dt.days / 365.25).fillna(0).astype(int)
    df_actif['Sexe'] = df_actif['CIVILITE'].apply(lambda x: 'Homme' if x == 'Mr' else 'Femme')
    
    fte = len(df_actif)
    
    # Stats de base
    effectif_stats = {
        'total_fte': fte,
        'hommes': len(df_actif[df_actif['Sexe'] == 'Homme']),
        'femmes': len(df_actif[df_actif['Sexe'] == 'Femme']),
        'age_moyen': round(df_actif['Age'].mean(), 1) if fte > 0 else 0,
        'age_min': int(df_actif['Age'].min()) if fte > 0 else 0,
        'age_max': int(df_actif['Age'].max()) if fte > 0 else 0,
        'anciennete_moyenne': round(df_actif['Anciennete'].mean(), 1) if fte > 0 else 0,
        'anciennete_min': int(df_actif['Anciennete'].min()) if fte > 0 else 0,
        'anciennete_max': int(df_actif['Anciennete'].max()) if fte > 0 else 0,
        'sortants_cheque': len(sortants_cheque),
        'annee_paie': derniere_annee,
        'mois_paie': dernier_mois
    }
    
    # Par établissement
    if fte > 0 and 'etb_name' in df_actif.columns:
        by_etb = df_actif.groupby('etb_name').agg({'MATRICULE': 'count', 'Age': 'mean', 'Anciennete': 'mean'}).round(1)
        effectif_stats['par_etablissement'] = {
            etb: {'count': int(row['MATRICULE']), 'age_moyen': round(row['Age'], 1), 'anciennete_moyenne': round(row['Anciennete'], 1)}
            for etb, row in by_etb.iterrows()
        }
    else:
        effectif_stats['par_etablissement'] = {}
    
    # Par classification
    if fte > 0 and 'CLASSIFICATION_LIB' in df_actif.columns:
        by_class = df_actif.groupby('CLASSIFICATION_LIB').agg({'MATRICULE': 'count', 'Age': 'mean', 'Anciennete': 'mean'}).round(1)
        effectif_stats['par_classification'] = {
            cls: {'count': int(row['MATRICULE']), 'age_moyen': round(row['Age'], 1), 'anciennete_moyenne': round(row['Anciennete'], 1)}
            for cls, row in by_class.iterrows()
        }
    else:
        effectif_stats['par_classification'] = {}
    
    # Par contrat
    effectif_stats['par_contrat'] = df_actif['TYPE_CONTRAT'].value_counts().to_dict() if fte > 0 else {}
    
    # Par section (GRADE_LIB)
    if 'GRADE_LIB' in df_actif.columns and fte > 0:
        by_grade = df_actif.groupby('GRADE_LIB').agg({'MATRICULE': 'count', 'Age': 'mean', 'Anciennete': 'mean'}).round(1)
        effectif_stats['par_section'] = {
            grade: {'count': int(row['MATRICULE']), 'age_moyen': round(row['Age'], 1), 'anciennete_moyenne': round(row['Anciennete'], 1)}
            for grade, row in by_grade.iterrows() if pd.notna(grade)
        }
    else:
        effectif_stats['par_section'] = {}
    
    # Situation familiale
    sit_fam_map = {'M': 'Marié(e)', 'C': 'Célibataire', 'D': 'Divorcé(e)', 'V': 'Veuf/Veuve'}
    if fte > 0 and 'SITUAT_FAM' in df_actif.columns:
        effectif_stats['par_situation_familiale'] = {
            sit_fam_map.get(k, k): int(v) 
            for k, v in df_actif['SITUAT_FAM'].value_counts().to_dict().items()
        }
    else:
        effectif_stats['par_situation_familiale'] = {}
    
    # ============ PYRAMIDES ============
    if fte > 0:
        # Pyramide des âges
        age_bins = [0, 25, 30, 35, 40, 45, 50, 55, 60, 100]
        age_labels = ['<25', '25-29', '30-34', '35-39', '40-44', '45-49', '50-54', '55-59', '60+']
        df_actif['Tranche_Age'] = pd.cut(df_actif['Age'], bins=age_bins, labels=age_labels)
        pyramide = df_actif.groupby(['Tranche_Age', 'Sexe'], observed=True).size().unstack(fill_value=0)
        
        hommes_pyramid = [int(pyramide.get('Homme', pd.Series()).get(l, 0)) for l in age_labels]
        femmes_pyramid = [int(pyramide.get('Femme', pd.Series()).get(l, 0)) for l in age_labels]
        
        effectif_stats['pyramide_age'] = {
            'labels': age_labels,
            'hommes': hommes_pyramid,
            'femmes': femmes_pyramid
        }
        
        # Calcul indices démographiques
        # Jeunes (<35 ans) = tranches <25, 25-29, 30-34 (indices 0, 1, 2)
        jeunes = sum(hommes_pyramid[:3]) + sum(femmes_pyramid[:3])
        # Seniors (>50 ans) = tranches 50-54, 55-59, 60+ (indices 6, 7, 8)
        seniors = sum(hommes_pyramid[6:]) + sum(femmes_pyramid[6:])
        
        # Pyramide ancienneté
        anc_bins = [0, 2, 5, 10, 15, 20, 25, 30, 100]
        anc_labels = ['<2', '2-4', '5-9', '10-14', '15-19', '20-24', '25-29', '30+']
        df_actif['Tranche_Anc'] = pd.cut(df_actif['Anciennete'], bins=anc_bins, labels=anc_labels)
        pyr_anc = df_actif.groupby('Tranche_Anc', observed=True).size()
        anc_values = [int(pyr_anc.get(l, 0)) for l in anc_labels]
        
        effectif_stats['pyramide_anciennete'] = {
            'labels': anc_labels,
            'values': anc_values
        }
        
        # Stabilité (ancienneté >= 5 ans) = indices 2 à 7 (5-9, 10-14, 15-19, 20-24, 25-29, 30+)
        stables = sum(anc_values[2:])
    else:
        effectif_stats['pyramide_age'] = {'labels': [], 'hommes': [], 'femmes': []}
        effectif_stats['pyramide_anciennete'] = {'labels': [], 'values': []}
        jeunes = 0
        seniors = 0
        stables = 0
    
    # Liste des sortants
    sortants_liste = []
    for _, row in sortants_cheque.iterrows():
        sortants_liste.append({
            'matricule': str(row['MATRICULE']),
            'nom': str(row['NOM']) if pd.notna(row['NOM']) else '',
            'prenom': str(row['PRENOM']) if pd.notna(row['PRENOM']) else '',
            'etablissement': str(row['etb_name']) if pd.notna(row.get('etb_name')) else ''
        })
    effectif_stats['sortants_liste'] = sortants_liste
    
    # ============ TOUS LES EMPLOYÉS (toutes périodes) ============
    df_all = df_effectif[df_effectif['MDP'] != 'cheque'].copy()
    df_all['DT_NAISSANCE'] = pd.to_datetime(df_all['DT_NAISSANCE'], errors='coerce')
    df_all['DEB_ADMINISTRATION'] = pd.to_datetime(df_all['DEB_ADMINISTRATION'], errors='coerce')
    df_all['Age'] = ((today - df_all['DT_NAISSANCE']).dt.days / 365.25).fillna(0).astype(int)
    df_all['Anciennete'] = ((today - df_all['DEB_ADMINISTRATION']).dt.days / 365.25).fillna(0).astype(int)
    df_all['Sexe'] = df_all['CIVILITE'].apply(lambda x: 'Homme' if x == 'Mr' else 'Femme')
    
    employes = []
    for _, row in df_all.iterrows():
        employes.append({
            'matricule': str(row['MATRICULE']),
            'nom': str(row['NOM']) if pd.notna(row['NOM']) else '',
            'prenom': str(row['PRENOM']) if pd.notna(row['PRENOM']) else '',
            'etablissement': str(row['etb_name']) if pd.notna(row.get('etb_name')) else '',
            'classification': str(row['CLASSIFICATION_LIB']) if pd.notna(row.get('CLASSIFICATION_LIB')) else '',
            'fonction': str(row['FONCTION_LIB']) if pd.notna(row.get('FONCTION_LIB')) else '',
            'section': str(row['GRADE_LIB']) if pd.notna(row.get('GRADE_LIB')) else '',
            'service': str(row['HIERARCHIE_LIB']) if pd.notna(row.get('HIERARCHIE_LIB')) else '',
            'type_contrat': str(row['TYPE_CONTRAT']) if pd.notna(row['TYPE_CONTRAT']) else '',
            'sexe': row['Sexe'],
            'age': int(row['Age']),
            'anciennete': int(row['Anciennete']),
            'date_naissance': row['DT_NAISSANCE'].strftime('%Y-%m-%d') if pd.notna(row['DT_NAISSANCE']) else '',
            'date_embauche': row['DEB_ADMINISTRATION'].strftime('%Y-%m-%d') if pd.notna(row['DEB_ADMINISTRATION']) else '',
            'situation_familiale': sit_fam_map.get(row.get('SITUAT_FAM'), '') if pd.notna(row.get('SITUAT_FAM')) else '',
            'annee_paie': int(row['ANNEE_PAIE']) if pd.notna(row.get('ANNEE_PAIE')) else derniere_annee,
            'mois_paie': int(row['NO_PAIE']) if pd.notna(row.get('NO_PAIE')) else dernier_mois
        })
    
    # ============ MASSE SALARIALE ============
    mois_cols = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
    for col in mois_cols:
        if col in df_ms.columns:
            df_ms[col] = pd.to_numeric(df_ms[col], errors='coerce').fillna(0)
    
    totaux_mois = {m: round(float(df_ms[m].sum()), 2) for m in mois_cols if m in df_ms.columns}
    ytd = sum(totaux_mois.values())
    
    mois_courant = mois_cols[dernier_mois - 1] if dernier_mois <= 12 else 'Janvier'
    montant_courant = totaux_mois.get(mois_courant, 0)
    
    by_etb_ms = {}
    by_class_ms = {}
    if mois_courant in df_ms.columns:
        if 'etb_name' in df_ms.columns:
            by_etb_ms = df_ms.groupby('etb_name')[mois_courant].sum().to_dict()
        if 'Classification' in df_ms.columns:
            by_class_ms = df_ms.groupby('Classification')[mois_courant].sum().to_dict()
    
    ms_stats = {
        'ytd': round(ytd, 2),
        'mois_courant': mois_courant,
        'montant_mois_courant': round(montant_courant, 2),
        'totaux_mois': totaux_mois,
        'par_etablissement': {k: round(float(v), 2) for k, v in by_etb_ms.items()},
        'par_classification': {k: round(float(v), 2) for k, v in by_class_ms.items()},
        'budget_annuel': 245000000,
        'detail': []
    }
    
    # ============ MOUVEMENTS ============
    mvt_stats = {
        'nb_recrutements': 0,
        'nb_departs': 0,
        'solde': 0,
        'motifs_depart': {'Démission': 0, 'Décès': 0, 'Fin période essai': 0, 'Retraite': 0, 'Autre': 0},
        'recrutements': [],
        'departs': [],
        'recrutements_par_site': {},
        'par_motif': {'Démission': 0, 'Décès': 0, 'Fin période essai': 0, 'Retraite': 0, 'Autre': 0}
    }
    
    try:
        # Lire l'onglet MOUVEMENT avec la structure spécifique Sonasid
        df_mvt_raw = pd.read_excel(excel_path, sheet_name='MOUVEMENT', header=None)
        
        # Trouver les sections recrutements et départs
        rec_header_idx = None
        dep_header_idx = None
        
        for idx, row in df_mvt_raw.iterrows():
            cell1 = str(row[1]).lower() if pd.notna(row[1]) else ''
            # Chercher la ligne header "Matricule"
            if cell1 == 'matricule':
                if rec_header_idx is None:
                    rec_header_idx = idx
                else:
                    dep_header_idx = idx
        
        # Parser les recrutements (entre rec_header_idx et dep_header_idx)
        recrutements_list = []
        recrutements_par_site = {}
        
        if rec_header_idx is not None:
            start_idx = rec_header_idx + 1
            end_idx = dep_header_idx if dep_header_idx else len(df_mvt_raw)
            
            for idx in range(start_idx, end_idx):
                row = df_mvt_raw.iloc[idx]
                matricule = row[1]
                
                # Ignorer les lignes vides ou "Total"
                if pd.isna(matricule) or str(matricule).lower() == 'total' or str(matricule).lower() == 'nan':
                    continue
                    
                # Vérifier que c'est un matricule valide (nombre)
                try:
                    mat_int = int(float(matricule))
                except:
                    continue
                
                nom = str(row[2]) if pd.notna(row[2]) else ''
                prenom = str(row[3]) if pd.notna(row[3]) else ''
                site = str(row[4]) if pd.notna(row[4]) else ''
                college = str(row[5]) if pd.notna(row[5]) else ''
                date_rec = row[6]
                fonction = str(row[7]) if pd.notna(row[7]) else ''
                
                # Formater la date
                date_str = ''
                if pd.notna(date_rec):
                    try:
                        date_str = pd.to_datetime(date_rec).strftime('%Y-%m-%d')
                    except:
                        date_str = str(date_rec)
                
                recrutements_list.append({
                    'matricule': str(mat_int),
                    'nom': nom,
                    'prenom': prenom,
                    'etablissement': site,
                    'classification': college,
                    'date': date_str,
                    'fonction': fonction
                })
                
                # Compter par site
                if site:
                    recrutements_par_site[site] = recrutements_par_site.get(site, 0) + 1
        
        # Parser les départs (après dep_header_idx)
        departs_list = []
        motifs_depart = {'Démission': 0, 'Décès': 0, 'Fin période essai': 0, 'Retraite': 0, 'Autre': 0}
        
        if dep_header_idx is not None:
            for idx in range(dep_header_idx + 1, len(df_mvt_raw)):
                row = df_mvt_raw.iloc[idx]
                matricule = row[1]
                
                # Ignorer les lignes vides ou "Total"
                if pd.isna(matricule) or str(matricule).lower() == 'total' or str(matricule).lower() == 'nan':
                    continue
                
                # Vérifier que c'est un matricule valide (nombre)
                try:
                    mat_int = int(float(matricule))
                except:
                    continue
                
                nom = str(row[2]) if pd.notna(row[2]) else ''
                prenom = str(row[3]) if pd.notna(row[3]) else ''
                site = str(row[4]) if pd.notna(row[4]) else ''
                college = str(row[5]) if pd.notna(row[5]) else ''
                date_dep = row[6]
                motif_raw = str(row[7]).lower() if pd.notna(row[7]) else ''
                
                # Formater la date
                date_str = ''
                if pd.notna(date_dep):
                    try:
                        date_str = pd.to_datetime(date_dep).strftime('%Y-%m-%d')
                    except:
                        date_str = str(date_dep)
                
                # Normaliser le motif
                if 'démission' in motif_raw or 'demission' in motif_raw:
                    motif = 'Démission'
                elif 'décès' in motif_raw or 'deces' in motif_raw or 'décés' in motif_raw:
                    motif = 'Décès'
                elif 'essai' in motif_raw or 'éssai' in motif_raw:
                    motif = 'Fin période essai'
                elif 'retraite' in motif_raw:
                    motif = 'Retraite'
                else:
                    motif = 'Autre'
                
                motifs_depart[motif] = motifs_depart.get(motif, 0) + 1
                
                departs_list.append({
                    'matricule': str(mat_int),
                    'nom': nom,
                    'prenom': prenom,
                    'etablissement': site,
                    'classification': college,
                    'date': date_str,
                    'motif': motif
                })
        
        mvt_stats = {
            'nb_recrutements': len(recrutements_list),
            'nb_departs': len(departs_list),
            'solde': len(recrutements_list) - len(departs_list),
            'motifs_depart': motifs_depart,
            'recrutements': recrutements_list,
            'departs': departs_list,
            'recrutements_par_site': recrutements_par_site,
            'par_motif': motifs_depart
        }
        
        print(f"📋 Mouvements: {len(recrutements_list)} recrutements, {len(departs_list)} départs")
        print(f"   Motifs: {motifs_depart}")
        
    except Exception as e:
        print(f"   ⚠️ Erreur lecture MOUVEMENT: {e}")
        import traceback
        traceback.print_exc()
    
    # ============ CALCUL DE TOUS LES INDICATEURS ============
    nb_rec = mvt_stats['nb_recrutements']
    nb_dep = mvt_stats['nb_departs']
    nb_demissions = mvt_stats['motifs_depart'].get('Démission', 0)
    nb_cadres = effectif_stats['par_classification'].get('Cadre', {}).get('count', 0)
    nb_cdi = effectif_stats['par_contrat'].get('CDI', 0)
    
    # Calculs avec protection contre division par zéro
    taux_feminisation = round((effectif_stats['femmes'] / fte) * 100, 1) if fte > 0 else 0
    taux_cdi = round((nb_cdi / fte) * 100, 1) if fte > 0 else 0
    turnover = round(((nb_rec + nb_dep) / 2 / fte) * 100, 2) if fte > 0 else 0
    taux_demission = round((nb_demissions / fte) * 100, 2) if fte > 0 else 0
    taux_depart = round((nb_dep / fte) * 100, 2) if fte > 0 else 0
    taux_retention = round(100 - taux_depart, 2)
    taux_reussite_essai = 100.0 if nb_rec == 0 else round(((nb_rec - mvt_stats['motifs_depart'].get('Fin période essai', 0)) / nb_rec) * 100, 1) if nb_rec > 0 else 100.0
    taux_croissance = round((mvt_stats['solde'] / fte) * 100, 2) if fte > 0 else 0
    cout_moyen = round(montant_courant / fte, 2) if fte > 0 else 0
    
    # Indices démographiques
    indice_jeunesse = round((jeunes / fte) * 100, 1) if fte > 0 else 0
    indice_vieillissement = round((seniors / fte) * 100, 1) if fte > 0 else 0
    indice_stabilite = round((stables / fte) * 100, 1) if fte > 0 else 0
    ratio_encadrement = round((nb_cadres / fte) * 100, 1) if fte > 0 else 0
    
    indicators = {
        # Indicateurs de rotation
        'turnover': turnover,
        'taux_demission': taux_demission,
        'taux_retention': taux_retention,
        'taux_reussite_essai': taux_reussite_essai,
        'taux_croissance': taux_croissance,
        
        # Indicateurs démographiques
        'taux_feminisation': taux_feminisation,
        'age_moyen': effectif_stats['age_moyen'],
        'indice_jeunesse': indice_jeunesse,
        'indice_vieillissement': indice_vieillissement,
        
        # Indicateurs de stabilité et structure
        'indice_stabilite': indice_stabilite,
        'anciennete_moyenne': effectif_stats['anciennete_moyenne'],
        'taux_cdi': taux_cdi,
        'ratio_encadrement': ratio_encadrement,
        
        # Indicateurs financiers
        'cout_moyen_mensuel': cout_moyen,
        'cout_par_etp': cout_moyen,
        'masse_salariale_par_fte': cout_moyen,
        'budget_annuel': ms_stats['budget_annuel'],
        'taux_realisation_budget': round((ytd / ms_stats['budget_annuel']) * 100, 1) if ms_stats['budget_annuel'] > 0 else 0,
        
        # Ratios par classification
        'ratio_cadres': ratio_encadrement,
        'ratio_maitrise': round((effectif_stats['par_classification'].get('Maîtrise', {}).get('count', 0) / fte) * 100, 1) if fte > 0 else 0,
        'ratio_employe': round((effectif_stats['par_classification'].get('Employé', {}).get('count', 0) / fte) * 100, 1) if fte > 0 else 0
    }
    
    # ============ FORMATION ============
    formation_stats = process_formation(formation_path, fte_total=fte, ms_ytd=ytd)
    if formation_stats:
        print(f"🎓 Formation: {formation_stats['kpis']['nb_actions']} actions, "
              f"{formation_stats['kpis']['nb_participations']} participations")
    else:
        print(f"   ⚠️ Fichier Formation non trouvé: {formation_path} (page Formation vide)")

    # ============ SOCIAL (en attente de données) ============
    social_stats = build_social_placeholder()

    # ============ ASSEMBLAGE FINAL ============
    mois_noms = ['', 'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']

    output_data = {
        'meta': {
            'generated_at': datetime.now().isoformat(),
            'source_file': excel_path.name,
            'periode': f"{mois_noms[dernier_mois]} {derniere_annee}",
            'periodes_disponibles': periodes
        },
        'effectif': effectif_stats,
        'employes': employes,
        'masse_salariale': ms_stats,
        'mouvements': mvt_stats,
        'indicateurs': indicators,
        'formation': formation_stats,
        'social': social_stats
    }

    # Sauvegarder
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2, cls=NumpyEncoder)

    print(f"\n✅ Données générées avec succès!")
    print(f"   📁 Fichier: {output_path}")
    print(f"   📅 Période: {mois_noms[dernier_mois]} {derniere_annee}")
    print(f"   👥 Effectif FTE: {fte}")
    print(f"   📊 Sections: {len(effectif_stats['par_section'])}")
    print(f"\n   === INDICATEURS ===")
    for k, v in sorted(indicators.items()):
        print(f"   {k}: {v}")

    return True

if __name__ == '__main__':
    if len(sys.argv) > 1:
        generate_data(sys.argv[1])
    else:
        generate_data()
