#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard RH - Générateur de données v11
Génère TOUS les indicateurs nécessaires pour le dashboard
"""

import pandas as pd
import json
import re
import unicodedata
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

def strip_accents(text):
    """Retire les accents pour des comparaisons de texte robustes aux fautes
    de saisie ('Retraité' / 'retraite', 'éssai' / 'essai', ...)."""
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')

def classify_motif_depart(motif_raw):
    """Normaliser le motif de départ. On ne force jamais tout ce qui n'est
    pas reconnu dans un fourre-tout 'Autre' : un motif qui ne correspond à
    aucun des cas connus est affiché tel quel (texte d'origine, nettoyé)."""
    original = str(motif_raw).strip() if motif_raw is not None else ''
    if not original:
        return 'Non renseigné'
    norm = strip_accents(original).lower()
    if 'demission' in norm:
        return 'Démission'
    if 'deces' in norm or 'décés' in norm:
        return 'Décès'
    if 'essai' in norm:
        return 'Fin période essai'
    if 'volontaire' in norm:
        return 'Départ volontaire'
    if 'retraite' in norm:
        return 'Retraite'
    if 'licenciement' in norm:
        return 'Licenciement'
    if 'abandon' in norm:
        return 'Abandon de poste'
    if 'mutation' in norm:
        return 'Mutation'
    if 'fin de contrat' in norm or 'fin contrat' in norm:
        return 'Fin de contrat'
    # Motif réel non reconnu : on garde le libellé d'origine plutôt que de
    # le masquer sous "Autre".
    return original[:1].upper() + original[1:]

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

    # Les colonnes d'identité de l'action (domaine, thématique, cabinet, dates)
    # ne sont renseignées que sur la 1ère ligne de chaque formation (cellules
    # fusionnées à l'export) : on les propage vers le bas.
    #
    # nb_jours / cout_formation / cout_logistique restent volontairement EN
    # DEHORS de cette propagation : une même action peut avoir été livrée en
    # plusieurs sous-sessions (cohortes différentes, coûts différents), et
    # chacune inscrit sa propre valeur sur une ligne du bloc — parfois
    # identique à la 1ère (répétée par erreur de saisie), parfois différente
    # (ex. Habilitation Electrique B0/H0 : 25200, 18000, 9000, 18000 MAD sur
    # 4 lignes distinctes). Les prendre en compte une seule fois (comme le
    # faisait l'ancienne version) sous-évaluait le coût réel total.
    action_cols = ['n_action', 'domaine', 'type_action', 'thematique', 'date_debut',
                   'date_fin', 'cabinet', 'formateur_interne']
    df[action_cols] = df[action_cols].ffill()
    df = df[df['n_action'].notna()].copy()
    df['nb_jours'] = pd.to_numeric(df['nb_jours'], errors='coerce')
    df['cout_formation'] = pd.to_numeric(df['cout_formation'], errors='coerce')
    df['cout_logistique'] = pd.to_numeric(df['cout_logistique'], errors='coerce')

    df['date_debut'] = pd.to_datetime(df['date_debut'], errors='coerce')
    df['date_fin'] = pd.to_datetime(df['date_fin'], errors='coerce')
    df['nom'] = df['nom'].fillna('').astype(str).str.strip()
    df['prenom'] = df['prenom'].fillna('').astype(str).str.strip()
    for col in ['domaine', 'type_action', 'classification', 'site', 'departement', 'statut', 'eval_chaud']:
        df[col] = df[col].fillna('Non renseigné').astype(str).str.strip().replace('', 'Non renseigné')
    df['nb_heures'] = pd.to_numeric(df['nb_heures'], errors='coerce').fillna(0)

    # Le même cabinet peut être saisi avec une casse différente d'une action à
    # l'autre (ex. "MAE Academy" / "MAE ACADEMY") : on regroupe sur une clé
    # normalisée tout en gardant un libellé d'affichage lisible.
    df['cabinet'] = df['cabinet'].fillna('Non renseigné').astype(str).str.strip().replace('', 'Non renseigné')
    df['cabinet_key'] = df['cabinet'].str.upper()
    cabinet_display = {}
    for key, label in zip(df['cabinet_key'], df['cabinet']):
        cabinet_display.setdefault(key, label)

    # Une action peut compter plusieurs dizaines de participants : si une saisie
    # ultérieure porte une date différente de la 1ère ligne (erreur de saisie),
    # elle ne doit pas faire "sortir" ces participants de leur action. Le N°
    # Action fait foi : on aligne tous les participants sur la date de leur
    # action (1ère ligne rencontrée pour ce N°).
    actions_df = df.drop_duplicates('n_action').copy()
    date_debut_par_action = actions_df.set_index('n_action')['date_debut']
    date_fin_par_action = actions_df.set_index('n_action')['date_fin']
    df['date_debut'] = df['n_action'].map(date_debut_par_action)
    df['date_fin'] = df['n_action'].map(date_fin_par_action)
    df['annee'] = df['date_debut'].dt.year
    actions_df['annee'] = actions_df['date_debut'].dt.year

    # Coût et durée réels d'une action = somme de toutes ses sous-sessions
    # (voir remarque ci-dessus), pas seulement la 1ère ligne du bloc.
    action_totals = df.groupby('n_action')[['nb_jours', 'cout_formation', 'cout_logistique']].sum(min_count=0)
    actions_df = actions_df.set_index('n_action')
    actions_df[['nb_jours', 'cout_formation', 'cout_logistique']] = action_totals
    actions_df = actions_df.reset_index()

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

    # Synthèse par cabinet de formation (coûts déduits au niveau action, pas participation)
    par_cabinet = {}
    for key, sub_actions in actions_df.groupby('cabinet_key'):
        nb_part_cabinet = int(len(df[df['cabinet_key'] == key]))
        par_cabinet[cabinet_display[key]] = {
            'nb_actions': int(len(sub_actions)),
            'nb_participations': nb_part_cabinet,
            'nb_jours': round(float(sub_actions['nb_jours'].fillna(0).sum()), 1),
            'cout_total': round(float((sub_actions['cout_formation'].fillna(0) + sub_actions['cout_logistique'].fillna(0)).sum()), 2)
        }

    annees_disponibles = sorted({int(a) for a in df['annee'].dropna().unique()})

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
            'annee': int(row['annee']) if pd.notna(row['annee']) else None,
            'nb_jours': float(row['nb_jours']) if pd.notna(row['nb_jours']) else 0,
            'nb_heures': float(row['nb_heures']) if pd.notna(row['nb_heures']) else 0,
            'cabinet': cabinet_display[row['cabinet_key']],
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
            'annee': int(row['annee']) if pd.notna(row['annee']) else None,
            'nom': row['nom'],
            'prenom': row['prenom'],
            'classification': row['classification'],
            'site': row['site'],
            'departement': row['departement'],
            'cabinet': cabinet_display[row['cabinet_key']],
            'statut': row['statut'],
            'eval_chaud': row['eval_chaud']
        })

    return {
        'meta': {
            'source_file': formation_path.name,
            'periode_min': mois_valides[0] if mois_valides else None,
            'periode_max': mois_valides[-1] if mois_valides else None,
            'annees_disponibles': annees_disponibles
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
        'par_cabinet': par_cabinet,
        'par_statut': par_statut,
        'evaluation_chaud': evaluation_chaud,
        'evolution_mensuelle': evolution_mensuelle,
        'actions': actions,
        'participations': participations
    }

def build_social_placeholder(label='sociales'):
    """Le volet Social/Sociétal attend un fichier de données dédié. Tant qu'il
    n'est pas fourni, on expose un état vide explicite plutôt que d'inventer
    des indicateurs."""
    return {
        'available': False,
        'message': f"Le fichier de données des actions {label} n'a pas encore été "
                   "fourni. Cette page affichera automatiquement les indicateurs "
                   "dès sa réception."
    }

FR_MONTHS = {
    'janvier': 1, 'fevrier': 2, 'février': 2, 'mars': 3, 'avril': 4, 'mai': 5,
    'juin': 6, 'juillet': 7, 'aout': 8, 'août': 8, 'septembre': 9,
    'octobre': 10, 'novembre': 11, 'decembre': 12, 'décembre': 12
}

def parse_periode_fr(raw):
    """Parse une période FR libre ('Fevrier 2026', 'Juin - Juillet 2026') en
    (année, mois_debut, mois_fin, libellé_normalisé)."""
    text = str(raw).strip().lower().replace('\n', ' ')
    year_match = re.search(r'(\d{4})', text)
    annee = int(year_match.group(1)) if year_match else None
    months_part = re.sub(r'\d{4}', '', text)
    found = []
    for name, num in FR_MONTHS.items():
        idx = re.search(rf'\b{name}\b', months_part)
        if idx:
            found.append((idx.start(), num, name.capitalize()))
    found.sort(key=lambda x: x[0])
    mois_debut = found[0][1] if found else None
    mois_fin = found[-1][1] if found else mois_debut
    if not found:
        libelle = str(raw).strip()
    elif len(found) == 1 or found[0][2] == found[-1][2]:
        libelle = f"{found[0][2]} {annee}" if annee else found[0][2]
    else:
        libelle = f"{found[0][2]} - {found[-1][2]} {annee}" if annee else f"{found[0][2]} - {found[-1][2]}"
    return annee, mois_debut, mois_fin, libelle

def split_multi(val):
    return [p.strip() for p in str(val).split(',') if p.strip()]

def parse_beneficiaires(val):
    try:
        return float(val), None
    except (ValueError, TypeError):
        return None, str(val).strip()

def build_social_stats(df, source_name):
    """Calculer les indicateurs (KPIs, répartitions, évolution) pour un sous-
    ensemble d'actions (Social ou Sociétal)."""
    actions = []
    par_region = {}
    par_site = {}
    mensuel = {}
    total_budget = 0.0
    total_beneficiaires = 0.0
    nb_actions_beneficiaires_non_numerique = 0

    for _, row in df.iterrows():
        annee, mois_debut, mois_fin, periode_libelle = parse_periode_fr(row['periode'])
        beneficiaires, beneficiaires_texte = parse_beneficiaires(row['beneficiaires_raw'])
        budget = float(row['budget']) if pd.notna(row['budget']) else 0.0
        regions = split_multi(row['region_raw']) if pd.notna(row['region_raw']) else []
        sites = split_multi(row['site_raw']) if pd.notna(row['site_raw']) else []

        total_budget += budget
        if beneficiaires is not None:
            total_beneficiaires += beneficiaires
        else:
            nb_actions_beneficiaires_non_numerique += 1

        for r in regions:
            par_region.setdefault(r, {'nb_actions': 0, 'budget': 0.0})
            par_region[r]['nb_actions'] += 1
            par_region[r]['budget'] += budget
        for s in sites:
            par_site.setdefault(s, {'nb_actions': 0, 'budget': 0.0})
            par_site[s]['nb_actions'] += 1
            par_site[s]['budget'] += budget

        if annee and mois_debut:
            key = f"{annee}-{mois_debut:02d}"
            mensuel.setdefault(key, {'nb_actions': 0, 'budget': 0.0})
            mensuel[key]['nb_actions'] += 1
            mensuel[key]['budget'] += budget

        action_text = str(row['action']).strip()
        actions.append({
            'titre': action_text[:90] + ('…' if len(action_text) > 90 else ''),
            'description': action_text,
            'periode': periode_libelle,
            'annee': annee,
            'beneficiaires': beneficiaires,
            'beneficiaires_texte': beneficiaires_texte,
            'budget': budget,
            'regions': regions,
            'sites': sites
        })

    annees_disponibles = sorted({a['annee'] for a in actions if a['annee']})
    mois_keys = sorted(mensuel.keys())

    return {
        'available': True,
        'meta': {
            'source_file': source_name,
            'annees_disponibles': annees_disponibles
        },
        'kpis': {
            'nb_actions': len(actions),
            'total_beneficiaires': int(total_beneficiaires),
            'nb_actions_beneficiaires_non_numerique': nb_actions_beneficiaires_non_numerique,
            'budget_total': round(total_budget, 2),
            'budget_moyen_action': round(total_budget / len(actions), 2) if actions else 0,
            'nb_regions': len(par_region),
            'nb_sites': len(par_site)
        },
        'par_region': par_region,
        'par_site': par_site,
        'evolution_mensuelle': {
            'labels': mois_keys,
            'nb_actions': [mensuel[k]['nb_actions'] for k in mois_keys],
            'budget': [round(mensuel[k]['budget'], 2) for k in mois_keys]
        },
        'actions': actions
    }

def classify_social_action(type_action_raw):
    """La colonne 'Type Action' vaut 'Social' ou 'Sociétales' selon les
    fichiers : on normalise sur ces deux catégories (par défaut Social si la
    valeur est absente ou inattendue)."""
    text = str(type_action_raw).strip().lower()
    if 'sociét' in text or 'societ' in text:
        return 'societal'
    return 'social'

def pick_column(raw_columns, keywords):
    """Trouve puis retire de la liste la première colonne dont l'en-tête
    contient un des mots-clés (recherche insensible à la casse). Retourne
    None si aucune colonne ne correspond."""
    for col in list(raw_columns):
        name = str(col).strip().lower()
        if any(kw in name for kw in keywords):
            raw_columns.remove(col)
            return col
    return None

def process_social(social_path):
    """Traiter le fichier des Actions sociales et sociétales et le scinder en
    deux volets distincts (Social / Sociétal) d'après la colonne 'Type Action'.

    Les en-têtes exacts varient d'un export à l'autre (colonne 'Type Action'
    parfois absente, ordre des colonnes différent) : on identifie chaque rôle
    par mot-clé dans l'en-tête plutôt que par position, pour ne pas décaler
    silencieusement les valeurs si une colonne manque ou change de place."""
    social_path = Path(social_path)
    if not social_path.exists():
        return {
            'social': build_social_placeholder('sociales'),
            'societal': build_social_placeholder('sociétales')
        }

    raw = pd.read_excel(social_path, sheet_name=0)
    available = list(raw.columns)
    col_type = pick_column(available, ['type'])
    col_action = pick_column(available, ['action'])
    col_periode = pick_column(available, ['date', 'période', 'periode', 'réalisation', 'realisation'])
    col_benef = pick_column(available, ['bénéfic', 'benefic'])
    col_budget = pick_column(available, ['budget'])
    col_region = pick_column(available, ['région', 'region'])
    col_site = pick_column(available, ['site'])

    if col_action is None:
        raise ValueError(
            f"Colonne 'Actions' introuvable dans {social_path.name} "
            f"(en-têtes lus : {list(raw.columns)})"
        )

    df = pd.DataFrame({
        'action': raw[col_action],
        'type_action': raw[col_type] if col_type is not None else '',
        'periode': raw[col_periode] if col_periode is not None else '',
        'beneficiaires_raw': raw[col_benef] if col_benef is not None else None,
        'budget': raw[col_budget] if col_budget is not None else 0,
        'region_raw': raw[col_region] if col_region is not None else None,
        'site_raw': raw[col_site] if col_site is not None else None,
    })
    df = df[df['action'].notna()].copy()

    df['category'] = df['type_action'].apply(classify_social_action)

    return {
        'social': build_social_stats(df[df['category'] == 'social'], social_path.name),
        'societal': build_social_stats(df[df['category'] == 'societal'], social_path.name)
    }

def generate_data(excel_path=None, output_path=None, formation_path=None, social_path=None):
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

    if social_path is None:
        social_path = Path(__file__).parent / 'data' / 'SOCIAL.xlsx'
    else:
        social_path = Path(social_path)

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
                    'site': site,
                    'college': college,
                    'date': date_str,
                    'fonction': fonction
                })

                # Compter par site
                if site:
                    recrutements_par_site[site] = recrutements_par_site.get(site, 0) + 1

        # Parser les départs (après dep_header_idx)
        departs_list = []
        motifs_depart = {}
        
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
                motif_raw = row[7] if pd.notna(row[7]) else ''

                # Formater la date
                date_str = ''
                if pd.notna(date_dep):
                    try:
                        date_str = pd.to_datetime(date_dep).strftime('%Y-%m-%d')
                    except:
                        date_str = str(date_dep)

                motif = classify_motif_depart(motif_raw)
                motifs_depart[motif] = motifs_depart.get(motif, 0) + 1

                departs_list.append({
                    'matricule': str(mat_int),
                    'nom': nom,
                    'prenom': prenom,
                    'site': site,
                    'college': college,
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

    # ============ SOCIAL / SOCIÉTAL ============
    social_split = process_social(social_path)
    social_stats = social_split['social']
    societal_stats = social_split['societal']
    if social_stats.get('available'):
        print(f"🤝 Social: {social_stats['kpis']['nb_actions']} actions, "
              f"{social_stats['kpis']['budget_total']:,.0f} MAD de budget")
        print(f"🌍 Sociétal: {societal_stats['kpis']['nb_actions']} actions, "
              f"{societal_stats['kpis']['budget_total']:,.0f} MAD de budget")
    else:
        print(f"   ⚠️ Fichier Social non trouvé: {social_path} (pages Social/Sociétal vides)")

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
        'social': social_stats,
        'societal': societal_stats
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
