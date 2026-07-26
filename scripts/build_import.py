#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the SQL Server injection script + review report from the consolidated
fiches de poste workbook (which now also carries the target table structures
and referentials, as provided by the user)."""
import html
import re
import sys
import unicodedata

import openpyxl

SRC = sys.argv[1] if len(sys.argv) > 1 else "data/Fiches_de_poste_Sonasid_consolidées.xlsx"
COMPETENCE_CSV = sys.argv[2] if len(sys.argv) > 2 else "data/Import_Cometence.csv"
OUT_SQL = sys.argv[3] if len(sys.argv) > 3 else "sql/injection_fiches_de_poste.sql"
OUT_REVIEW = sys.argv[4] if len(sys.argv) > 4 else "data/import_review.xlsx"

CREATED_BY = "IMPORT_FDP"  # dedicated import account -- change if needed

# Known already-in-DB postes (from the ACH_OPEX example export) -> forced POSTE code
KNOWN_DB_POSTE_OVERRIDE = {
    "chef d exploitation achats opex": "ACH_OPEX",
}

DIRECTION_ABBR = {
    "Direction Achats": "ACH",
    "Direction Commerciale & Marketing": "DCM",
    "Direction Financière": "FIN",
    "Direction des Ressources Humaines": "RH",
    "Direction Systèmes d'Information": "SI",
    "Direction Industrielle (Sites)": "IND",
    "Direction Industrielle & SST": "SST",
    "Direction Affaires Juridiques, Contrôle Interne et Risque Management": "JUR",
    "Direction Audit": "AUD",
    "Direction Co-produits": "COPR",
    "Direction SI et stratégie": "STRAT",
    "Direction support": "SUPP",
    "Direction Technique": "TECH",
    "Direction Générale": "DG",
}

# Only confident/unambiguous semantic matches -- rest left NULL (reported)
FAMILLE_METIER_MAP = {
    "Direction Achats": "Ach",
    "Direction Commerciale & Marketing": "COM",
    "Direction Systèmes d'Information": "INF",
    "Direction Industrielle (Sites)": "PRO",
    "Direction Générale": "DIR",
}

STOPWORDS = {"de", "du", "des", "la", "le", "les", "en", "et", "à", "au", "aux",
             "d", "l", "un", "une", "pour", "sur", "dans"}


def strip_accents(s):
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def norm_match(s):
    """Loose normalization for competency-label matching (case/accents/space-insensitive) --
    approximates SQL Server's default case-insensitive collation (CI_AS)."""
    s = strip_accents(str(s or "")).lower()
    s = re.sub(r"\s+", " ", s).strip()
    s = s.rstrip(".")
    return s


def norm_key(s):
    s = strip_accents(str(s or "")).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return s


def slugify_title(title, max_words=3, max_len=24):
    words = re.findall(r"[A-Za-zÀ-ÿ0-9]+", title)
    kept = []
    for w in words:
        wl = strip_accents(w).lower()
        if wl in STOPWORDS:
            continue
        kept.append(strip_accents(w).upper())
        if len(kept) >= max_words:
            break
    slug = "_".join(kept)[:max_len].rstrip("_")
    return slug or "POSTE"


def normalize_existing_code(code):
    code = strip_accents(str(code)).upper()
    code = re.sub(r"[^A-Z0-9]+", "_", code).strip("_")
    return code


def sql_str(v):
    """Render a python value as a T-SQL literal (NULL, N'escaped', or number)."""
    if v is None or v == "":
        return "NULL"
    if isinstance(v, (int, float)):
        return str(v)
    s = str(v).replace("'", "''")
    return f"N'{s}'"


def missions_to_html(text):
    if not text:
        return None
    lines = [l.strip() for l in str(text).splitlines() if l.strip()]
    parts = []
    for l in lines:
        l = l.lstrip("•-*").strip()
        parts.append(f"<p>{html.escape(l)}</p>")
    return "".join(parts)


def map_niveau_etude(text):
    if not text:
        return None, "vide"
    t = norm_key(text)
    if "doctorat" in t:
        return "D", None
    if "ingenieur" in t:
        return "I", None
    if "master" in t:
        return "Ma", None
    if "grande ecole" in t or "grandes ecoles" in t:
        return "G", None
    if "licence" in t:
        return "L", None
    if re.search(r"\bmaitrise\b", t) and "bac" not in t:
        return "M", None
    if re.search(r"\bbac\b", t) and "+" not in text and re.search(r"\bbac\s*\d", t) is None:
        return "B", None
    if "secondaire" in t:
        return "S", None
    if "primaire" in t:
        return "P", None
    return None, f"ambigu/non reconnu: {text!r}"


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["Fiches de poste"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    fiches = [dict(zip(header, r)) for r in rows[1:]]

    ws_comp = wb["CODE_COMPETENCE"]
    comp_rows = list(ws_comp.iter_rows(values_only=True))
    comp_header = comp_rows[0]
    existing_codes = set(str(r[0]) for r in comp_rows[1:] if r[0] is not None)
    existing_intitules = set(str(r[1]).strip() for r in comp_rows[1:] if r[1])
    existing_intitules_norm = {norm_match(x) for x in existing_intitules}

    # Import_Cometence.csv (competency referential) -- decode cp1252
    csv_path = COMPETENCE_CSV
    csv_text = open(csv_path, "rb").read().decode("cp1252")
    csv_rows = []
    for line in csv_text.splitlines():
        if not line.strip():
            continue
        parts = line.split(";")
        if len(parts) < 4:
            continue
        code, intitule, typ, definition = parts[0].strip(), parts[1].strip(), parts[2].strip(), parts[3].strip()
        csv_rows.append((code, intitule, typ, definition))
    missing_comp = [r for r in csv_rows if r[0] not in existing_codes]
    for code, intitule, typ, definition in missing_comp:
        existing_intitules.add(intitule)
        existing_intitules_norm.add(norm_match(intitule))

    used_codes = set(existing_codes)
    assigned = []  # (fiche, poste_code, source_of_code)
    unmapped_metier = []
    unmapped_niveau = []
    unmatched_competences = []  # (title, comp_label)

    for f in fiches:
        title = f["Intitulé du poste"] or ""
        direction = f["Direction / Entité"] or ""
        key = norm_key(title)

        if key in KNOWN_DB_POSTE_OVERRIDE:
            poste_code = KNOWN_DB_POSTE_OVERRIDE[key]
            source = "existant en base (ACH_OPEX)"
        elif f.get("Code poste"):
            poste_code = normalize_existing_code(f["Code poste"])
            source = "Code poste source"
        else:
            abbr = DIRECTION_ABBR.get(direction, "GEN")
            poste_code = f"{abbr}_{slugify_title(title)}"
            source = "généré (direction + intitulé)"

        base_code = poste_code
        n = 2
        while poste_code in used_codes and key not in KNOWN_DB_POSTE_OVERRIDE:
            poste_code = f"{base_code}_{n}"
            n += 1
        used_codes.add(poste_code)

        metier_code = FAMILLE_METIER_MAP.get(direction)
        if not metier_code:
            unmapped_metier.append((title, direction))

        niveau_code, niveau_issue = map_niveau_etude(f.get("Niveau de diplôme"))
        if niveau_issue:
            unmapped_niveau.append((title, f.get("Niveau de diplôme"), niveau_issue))

        competences = f.get("Compétences requises")
        if competences:
            for comp_line in [l.strip() for l in str(competences).splitlines() if l.strip()]:
                m = re.match(r"^(.*?):\s*(.+?)(?:\s*\((.+?)\))?$", comp_line)
                if not m:
                    continue
                intitule_comp = m.group(2).strip()
                if norm_match(intitule_comp) not in existing_intitules_norm:
                    unmatched_competences.append((title, intitule_comp))

        assigned.append({
            "fiche": f,
            "poste_code": poste_code,
            "code_source": source,
            "metier_code": metier_code,
            "niveau_code": niveau_code,
            "already_in_db": key in KNOWN_DB_POSTE_OVERRIDE,
        })

    # ---------------------------------------------------------------
    # Review workbook
    # ---------------------------------------------------------------
    rwb = openpyxl.Workbook()
    rws = rwb.active
    rws.title = "POSTE codes"
    rws.append(["POSTE (code généré)", "Origine du code", "Déjà en base ?", "Intitulé", "Direction",
                "METIER (FAMILLE_METIER)", "NIVEAU_ETUDE mappé", "Niveau de diplôme (source)"])
    for a in assigned:
        f = a["fiche"]
        rws.append([a["poste_code"], a["code_source"], "OUI" if a["already_in_db"] else "",
                    f["Intitulé du poste"], f["Direction / Entité"],
                    a["metier_code"] or "", a["niveau_code"] or "", f.get("Niveau de diplôme") or ""])
    for i, w in enumerate([26, 28, 12, 40, 45, 22, 16, 40], start=1):
        rws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    rws2 = rwb.create_sheet("METIER non mappés")
    rws2.append(["Intitulé du poste", "Direction / Entité (sans code FAMILLE_METIER)"])
    seen_dirs = {}
    for title, direction in unmapped_metier:
        seen_dirs.setdefault(direction, 0)
        seen_dirs[direction] += 1
        rws2.append([title, direction])
    rws2.column_dimensions["A"].width = 45
    rws2.column_dimensions["B"].width = 50

    rws3 = rwb.create_sheet("Directions sans code (résumé)")
    rws3.append(["Direction / Entité", "Nb postes concernés"])
    for d, c in sorted(seen_dirs.items(), key=lambda x: -x[1]):
        rws3.append([d, c])
    rws3.column_dimensions["A"].width = 50

    rws4 = rwb.create_sheet("NIVEAU_ETUDE non mappés")
    rws4.append(["Intitulé du poste", "Niveau de diplôme (texte source)", "Raison"])
    for title, txt, reason in unmapped_niveau:
        rws4.append([title, txt, reason])
    rws4.column_dimensions["A"].width = 45
    rws4.column_dimensions["B"].width = 55
    rws4.column_dimensions["C"].width = 35

    rws5 = rwb.create_sheet("CODE_COMPETENCE manquants")
    rws5.append(["Code", "Intitulé", "Type", "Définition"])
    for r in missing_comp:
        rws5.append(list(r))
    rws5.column_dimensions["A"].width = 16
    rws5.column_dimensions["B"].width = 45
    rws5.column_dimensions["D"].width = 70

    rws6 = rwb.create_sheet("Compétences non matchées")
    rws6.append(["Intitulé du poste", "Compétence (texte fiche)",
                  "Trouvée dans CODE_COMPETENCE.INTITULE ?"])
    for title, comp_label in unmatched_competences:
        rws6.append([title, comp_label, "NON -> ligne POSTE_CRITERES non générée"])
    rws6.column_dimensions["A"].width = 45
    rws6.column_dimensions["B"].width = 55
    rws6.column_dimensions["C"].width = 40

    rwb.save(OUT_REVIEW)

    # ---------------------------------------------------------------
    # SQL script
    # ---------------------------------------------------------------
    lines = []
    lines.append("-- =====================================================================")
    lines.append("-- Script d'injection : Fiches de poste Sonasid -> SQL Server")
    lines.append("-- Genere automatiquement -- A REVOIR avant execution en production.")
    lines.append(f"-- Compte utilise pour CREATED_BY/MODIFIED_BY : {CREATED_BY}")
    lines.append("-- Le script est idempotent (IF NOT EXISTS) : le rejouer ne duplique rien.")
    lines.append("-- =====================================================================")
    lines.append("SET NOCOUNT ON;")
    lines.append("SET XACT_ABORT ON;")
    lines.append("BEGIN TRANSACTION;")
    lines.append("")

    # 1) CODE_COMPETENCE -- only the codes missing from the referential
    lines.append("-- ---------------------------------------------------------------------")
    lines.append(f"-- 1) CODE_COMPETENCE : {len(missing_comp)} competences du referentiel absentes de la table")
    lines.append("-- ---------------------------------------------------------------------")
    for code, intitule, typ, definition in missing_comp:
        nature = typ if typ.strip().lstrip("-").isdigit() else "NULL"
        lines.append(
            "IF NOT EXISTS (SELECT 1 FROM dbo.CODE_COMPETENCE WHERE CODE_COMPETENCE = "
            f"{sql_str(code)})\n"
            "BEGIN\n"
            "    INSERT INTO dbo.CODE_COMPETENCE (CODE_COMPETENCE, INTITULE, NATURE_COMPETENCE, DESCRIPTION, "
            "CREATED_BY, MODIFIED_BY, CREATED_DATE, MODIFIED_DATE)\n"
            f"    VALUES ({sql_str(code)}, {sql_str(intitule)}, {nature}, {sql_str(definition)}, "
            f"{sql_str(CREATED_BY)}, {sql_str(CREATED_BY)}, GETDATE(), GETDATE());\n"
            "END;"
        )
    lines.append("")

    # 2) POSTE_DEFINITION / POSTE_OBJECTIFS / POSTE_CRITERES per fiche
    n_skipped = 0
    n_inserted = 0
    for a in assigned:
        f = a["fiche"]
        poste = a["poste_code"]
        if a["already_in_db"]:
            n_skipped += 1
            lines.append(f"-- {poste} : deja present en base (exemple ACH_OPEX fourni) -> ignore")
            lines.append("")
            continue
        n_inserted += 1

        intitule = f["Intitulé du poste"]
        finalite = f.get("Finalité du poste")
        dimension = f.get("Dimension & enjeux")
        missions_html = missions_to_html(f.get("Missions principales"))
        metier = a["metier_code"]
        niveau = a["niveau_code"]
        teletravail = "NULL"
        psh = "NULL"
        inter_interne = f.get("Interlocuteurs internes")
        inter_externe = f.get("Interlocuteurs externes")
        sup_directe = f.get("Supervision directe (effectif)")
        sup_indirecte = f.get("Supervision indirecte (effectif)")
        exp_globale = f.get("Expérience globale")
        exp_specifique = f.get("Expérience spécifique")
        filiere_evol = f.get("Filière(s) d'évolution")

        lines.append(f"-- ---- POSTE {poste} : {intitule} ----")
        lines.append(f"IF NOT EXISTS (SELECT 1 FROM dbo.POSTE_DEFINITION WHERE POSTE = {sql_str(poste)})")
        lines.append("BEGIN")
        lines.append(
            "    INSERT INTO dbo.POSTE_DEFINITION (POSTE, INTITULE, INTITULE_COMPLET, OBSERVATIONS, DIMENSION, "
            "MISSION_PRINCIPALE, METIER, NIVEAU_ETUDE, TELETRAVAIL, PSH, INTER_INTERNE, INTER_EXTERNE, "
            "MANAGEMENT_DIRECTE, Management_indi, Expérience_globale, Expérience_spéci, Filière_Evolution, "
            "FERMETURE, DT_CREATION, CREATED_BY, MODIFIED_BY, CREATED_DATE, MODIFIED_DATE)"
        )
        lines.append(
            f"    VALUES ({sql_str(poste)}, {sql_str(intitule)}, {sql_str(intitule)}, {sql_str(finalite)}, "
            f"{sql_str(dimension)}, {sql_str(missions_html)}, {sql_str(metier)}, {sql_str(niveau)}, "
            f"{teletravail}, {psh}, {sql_str(inter_interne)}, {sql_str(inter_externe)}, "
            f"{sql_str(sup_directe)}, {sql_str(sup_indirecte)}, {sql_str(exp_globale)}, {sql_str(exp_specifique)}, "
            f"{sql_str(filiere_evol)}, 0, GETDATE(), {sql_str(CREATED_BY)}, {sql_str(CREATED_BY)}, "
            "GETDATE(), GETDATE());"
        )
        lines.append("END;")

        # POSTE_OBJECTIFS -- one row per KPI line
        kpis = f.get("Indicateurs de performance (KPIs)")
        if kpis:
            for kpi_line in [l.strip().lstrip("•-*").strip() for l in str(kpis).splitlines() if l.strip()]:
                lines.append(
                    f"IF NOT EXISTS (SELECT 1 FROM dbo.POSTE_OBJECTIFS WHERE POSTE = {sql_str(poste)} "
                    f"AND RESULTAT_ATTENDU = {sql_str(kpi_line)})"
                )
                lines.append("BEGIN")
                lines.append(
                    "    INSERT INTO dbo.POSTE_OBJECTIFS (POSTE, OBJECTIF, MOYENS_PREVUS, RESULTAT_ATTENDU, "
                    "CREATED_BY, MODIFIED_BY, CREATED_DATE, MODIFIED_DATE)"
                )
                lines.append(
                    f"    VALUES ({sql_str(poste)}, 2, {sql_str(kpi_line)}, {sql_str(kpi_line)}, "
                    f"{sql_str(CREATED_BY)}, {sql_str(CREATED_BY)}, GETDATE(), GETDATE());"
                )
                lines.append("END;")

        # POSTE_CRITERES -- one row per competence, matched by label against CODE_COMPETENCE.INTITULE
        competences = f.get("Compétences requises")
        if competences:
            for comp_line in [l.strip() for l in str(competences).splitlines() if l.strip()]:
                m = re.match(r"^(.*?):\s*(.+?)(?:\s*\((.+?)\))?$", comp_line)
                if not m:
                    continue
                intitule_comp = m.group(2).strip()
                niveau_requis = m.group(3).strip() if m.group(3) else None
                deg = niveau_requis if niveau_requis and niveau_requis.isdigit() else "NULL"
                lines.append(
                    "    -- competence: recherche par intitule (a verifier / completer le code exact si connu)"
                )
                lines.append(
                    f"    IF EXISTS (SELECT 1 FROM dbo.CODE_COMPETENCE WHERE INTITULE = {sql_str(intitule_comp)})"
                )
                lines.append("    BEGIN")
                lines.append(
                    f"        IF NOT EXISTS (SELECT 1 FROM dbo.POSTE_CRITERES pc JOIN dbo.CODE_COMPETENCE cc "
                    f"ON pc.CODE_COMPETENCE = cc.CODE_COMPETENCE WHERE pc.POSTE = {sql_str(poste)} "
                    f"AND cc.INTITULE = {sql_str(intitule_comp)})"
                )
                lines.append("        BEGIN")
                lines.append(
                    "            INSERT INTO dbo.POSTE_CRITERES (POSTE, CODE_COMPETENCE, DEGRE_MAITRISE, "
                    "DESCRIPTION, CREATED_BY, MODIFIED_BY, CREATED_DATE, MODIFIED_DATE)"
                )
                lines.append(
                    f"            SELECT {sql_str(poste)}, cc.CODE_COMPETENCE, {deg}, cc.DESCRIPTION, "
                    f"{sql_str(CREATED_BY)}, {sql_str(CREATED_BY)}, GETDATE(), GETDATE() "
                    f"FROM dbo.CODE_COMPETENCE cc WHERE cc.INTITULE = {sql_str(intitule_comp)};"
                )
                lines.append("        END;")
                lines.append("    END;")
        lines.append("")

    lines.append("COMMIT TRANSACTION;")

    with open(OUT_SQL, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))

    print(f"Postes traites: {len(assigned)} | deja en base (ignores): {n_skipped} | a inserer: {n_inserted}")
    print(f"CODE_COMPETENCE manquants a inserer: {len(missing_comp)}")
    print(f"METIER non mappe: {len(unmapped_metier)} postes sur {len(seen_dirs)} directions")
    print(f"NIVEAU_ETUDE non mappe/ambigu: {len(unmapped_niveau)} postes")
    print(f"Competences non matchees (texte fiche vs CODE_COMPETENCE.INTITULE): {len(unmatched_competences)}")
    print(f"\nSQL: {OUT_SQL}")
    print(f"Review: {OUT_REVIEW}")


if __name__ == "__main__":
    main()
