#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Consolidate 'fiches de poste' (xlsx + docx) into one Excel file."""
import glob
import hashlib
import os
import re
import unicodedata
import sys

import openpyxl
import docx

ROOT = sys.argv[1] if len(sys.argv) > 1 else "."
OUT = sys.argv[2] if len(sys.argv) > 2 else "consolidated.xlsx"

DIRECTION_MAP = {
    "DP ACHATS": "Direction Achats",
    "DP AUDIT": "Direction Audit",
    "DP FINANCE": "Direction Finances",
    "DP MG": "Direction Moyens Généraux",
    "DP RH": "Direction RH",
    "DP JURIDIQUE RM & CI": "Direction Affaires Juridiques, Contrôle Interne et Risque Management",
    "DP DCM": "Direction Commerciale & Marketing",
    "DP IT Supply Chain & STRAT": "Direction SI, Supply Chain & Stratégie",
}


def norm_label(s):
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-zA-Z0-9]+", " ", s).strip().lower()
    return s


DIRECTION_SYNONYMS = {
    "daf": "Direction Financière",
    "direction finances": "Direction Financière",
    "direction financiere": "Direction Financière",
    "direction achats logistique": "Direction Achats",
    "direction des achats": "Direction Achats",
    "direction industrielle": "Direction Industrielle (Sites)",
    "direction juridique": "Direction Affaires Juridiques, Contrôle Interne et Risque Management",
    "direction strategie": "Direction SI et stratégie",
    "direction des ressources humaines et rse": "Direction des Ressources Humaines",
}


def normalize_direction(name):
    if not name:
        return name
    key = norm_label(name)
    return DIRECTION_SYNONYMS.get(key, name)


def guess_direction(path):
    parts = path.split(os.sep)
    for p in reversed(parts[:-1]):
        if p.startswith("Direction "):
            return p
    for p in parts:
        if p in DIRECTION_MAP:
            return DIRECTION_MAP[p]
    # fallback: parent folder name
    return parts[-2] if len(parts) >= 2 else ""


# ---------------------------------------------------------------------------
# XLSX extraction
# ---------------------------------------------------------------------------

FICHE_MARKER = "fiche de poste"

XLSX_LABELS = {
    "libelle": ["libelle du poste"],
    "code": ["code poste"],
    "version": ["version"],
    "entite": ["entite pole de rattachement", "entite pole de rattachement"],
    "metier_poste": ["metier du poste"],
    "sous_famille": ["sous famille de rattachement"],
    "n1": ["n 1 hierarchique"],
    "teletravail": ["eligibilite au teletravail"],
    "psh": ["eligibilite d une psh", "eligibilite dune psh"],
    "finalite_hdr": ["finalite du poste"],
    "dimension_hdr": ["dimension enjeux"],
    "missions_hdr": ["missions principales"],
    "kpi_hdr": ["indicateurs cles de performance"],
    "interlocuteurs_hdr": ["principaux interlocuteurs du poste"],
    "interne": ["interne"],
    "externe": ["externe"],
    "management_hdr": ["management"],
    "sup_directe": ["supervision directe"],
    "sup_indirecte": ["supervision indirecte"],
    "formation_hdr": ["formation experience requises"],
    "niveau_diplome": ["niveau de diplome"],
    "exp_globale": ["experience globale"],
    "exp_specifique": ["experience specifique"],
    "competences_hdr": ["competences requises"],
    "famille_comp_col": ["famille de competences"],
    "intitule_comp_col": ["intitule competence"],
    "niveau_requis_col": ["niveau requis"],
    "filiere_evolution_hdr": ["filiere s d evolution", "filieres d evolution"],
}


def find_fiche_sheets(wb):
    sheets = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        found = False
        for row in ws.iter_rows(min_row=1, max_row=3):
            for c in row:
                if c.value and FICHE_MARKER in norm_label(c.value):
                    found = True
                    break
            if found:
                break
        if found:
            sheets.append(sn)
    return sheets


def extract_xlsx_sheet(ws):
    grid = {}
    max_row = min(ws.max_row, 100)
    max_col = min(ws.max_column, 30)
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for c in row:
            if c.value not in (None, ""):
                grid[(c.row, c.column)] = c.value

    label_pos = {}  # key -> list of (row, col)
    label_cells = set()
    for (r, c), v in grid.items():
        nl = norm_label(v)
        for key, labels in XLSX_LABELS.items():
            for lbl in labels:
                if nl == lbl or nl.startswith(lbl + " "):
                    label_pos.setdefault(key, []).append((r, c))
                    label_cells.add((r, c))
                    break

    label_rows = set(r for positions in label_pos.values() for (r, c) in positions)

    def first(key):
        return label_pos[key][0] if key in label_pos and label_pos[key] else None

    def value_right(row, col, max_search=10):
        for cc in range(col + 1, col + max_search):
            if (row, cc) in label_cells:
                break
            v = grid.get((row, cc))
            if v not in (None, ""):
                return str(v).strip()
        return ""

    def value_below(row, col_start=1, col_end=16, max_rows=6):
        collected = []
        for r in range(row + 1, row + 1 + max_rows):
            if r in label_rows:
                break
            rowvals = []
            for c in range(col_start, col_end):
                v = grid.get((r, c))
                if v not in (None, ""):
                    rowvals.append(str(v).strip())
            if rowvals:
                collected.append(" ".join(rowvals))
        return "\n".join(collected).strip()

    def value_below_col(row, col, max_rows=6):
        collected = []
        for r in range(row + 1, row + 1 + max_rows):
            if r in label_rows:
                break
            v = grid.get((r, col))
            if v not in (None, ""):
                collected.append(str(v).strip())
        return "\n".join(collected).strip()

    data = {}

    p = first("libelle")
    data["libelle"] = value_right(*p) if p else ""

    p = first("code")
    data["code"] = value_right(*p) if p else ""

    p = first("version")
    if p:
        r, c = p
        # 'Version : 0' often sits in the same cell as label -> parse after ':'
        raw = grid.get((r, c), "")
        m = re.search(r":\s*(.+)$", str(raw))
        data["version"] = m.group(1).strip() if m else value_right(r, c)
    else:
        data["version"] = ""

    p = first("entite")
    data["entite"] = value_right(*p) if p else ""

    p = first("metier_poste")
    data["metier_poste"] = value_right(*p) if p else ""

    p = first("n1")
    data["n1"] = value_right(*p) if p else ""

    p = first("finalite_hdr")
    data["finalite"] = value_below(*p) if p else ""

    p = first("dimension_hdr")
    data["dimension"] = value_below(*p) if p else ""

    p = first("missions_hdr")
    data["missions"] = value_below_col(p[0], p[1]) if p else ""

    p = first("kpi_hdr")
    data["kpis"] = value_below_col(p[0], p[1]) if p else ""

    p = first("interne")
    data["interlocuteurs_internes"] = value_right(*p) if p else ""
    if not data["interlocuteurs_internes"] and p:
        data["interlocuteurs_internes"] = value_below_col(*p)

    p = first("externe")
    data["interlocuteurs_externes"] = value_right(*p) if p else ""
    if not data["interlocuteurs_externes"] and p:
        data["interlocuteurs_externes"] = value_below_col(*p)

    p = first("sup_directe")
    data["supervision_directe"] = value_right(*p) if p else ""

    p = first("sup_indirecte")
    data["supervision_indirecte"] = value_right(*p) if p else ""

    p = first("niveau_diplome")
    data["niveau_diplome"] = value_right(*p) if p else ""

    p = first("exp_globale")
    data["exp_globale"] = value_right(*p) if p else ""

    p = first("exp_specifique")
    data["exp_specifique"] = value_right(*p) if p else ""

    p = first("filiere_evolution_hdr")
    data["filiere_evolution"] = value_below(*p) if p else ""

    # Compétences table
    competences = []
    p = first("famille_comp_col")
    if p:
        hdr_row, famille_col = p
        intitule_col = None
        niveau_col = None
        ip = first("intitule_comp_col")
        if ip:
            intitule_col = ip[1]
        nv = first("niveau_requis_col")
        if nv:
            niveau_col = nv[1]
        r = hdr_row + 1
        blanks = 0
        while r < hdr_row + 60 and blanks < 3:
            famille = grid.get((r, famille_col))
            intitule = grid.get((r, intitule_col)) if intitule_col else None
            # some rows put the competency name in column right after famille if intitule blank
            if not intitule:
                for cc in range(famille_col + 1, famille_col + 4):
                    v = grid.get((r, cc))
                    if v:
                        intitule = v
                        break
            niveau = grid.get((r, niveau_col)) if niveau_col else None
            if famille or intitule:
                blanks = 0
                label = f"{str(famille).strip() if famille else ''}: {str(intitule).strip() if intitule else ''}".strip(": ")
                if niveau:
                    label += f" ({str(niveau).strip()})"
                competences.append(label)
            else:
                blanks += 1
            r += 1
    data["competences"] = "\n".join(competences)

    return data


def process_xlsx(path):
    results = []
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        print(f"  ERREUR ouverture {path}: {e}")
        return results
    fiche_sheets = find_fiche_sheets(wb)
    for sn in fiche_sheets:
        ws = wb[sn]
        try:
            data = extract_xlsx_sheet(ws)
        except Exception as e:
            print(f"  ERREUR extraction {path} [{sn}]: {e}")
            continue
        row = {
            "Intitulé du poste": data.get("libelle") or sn,
            "Direction / Entité": normalize_direction(data.get("entite")) or guess_direction(path),
            "Métier du poste": data.get("metier_poste", ""),
            "Code poste": data.get("code", ""),
            "Version": data.get("version", ""),
            "N+1 hiérarchique": data.get("n1", ""),
            "Finalité du poste": data.get("finalite", ""),
            "Dimension & enjeux": data.get("dimension", ""),
            "Missions principales": data.get("missions", ""),
            "Indicateurs de performance (KPIs)": data.get("kpis", ""),
            "Interlocuteurs internes": data.get("interlocuteurs_internes", ""),
            "Interlocuteurs externes": data.get("interlocuteurs_externes", ""),
            "Supervision directe (effectif)": data.get("supervision_directe", ""),
            "Supervision indirecte (effectif)": data.get("supervision_indirecte", ""),
            "Niveau de diplôme": data.get("niveau_diplome", ""),
            "Expérience globale": data.get("exp_globale", ""),
            "Expérience spécifique": data.get("exp_specifique", ""),
            "Compétences requises": data.get("competences", ""),
            "Filière(s) d'évolution": data.get("filiere_evolution", ""),
            "Type source": "Excel",
            "Fichier source": path,
            "Onglet source": sn,
        }
        results.append(row)
    return results


# ---------------------------------------------------------------------------
# DOCX extraction
# ---------------------------------------------------------------------------


def norm_up(s):
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper().strip()


def parse_comp_rows(rows):
    out = []
    for r in rows:
        cells = [c.strip() for c in r]
        if not any(cells):
            continue
        cat = cells[0] if len(cells) > 0 else ""
        comp = cells[1] if len(cells) > 1 else ""
        niv = cells[2] if len(cells) > 2 else ""
        if not comp and not cat:
            continue
        label = f"{cat}: {comp}".strip(": ")
        if niv:
            label += f" ({niv})"
        out.append(label)
    return out


def process_docx(path):
    try:
        d = docx.Document(path)
    except Exception as e:
        print(f"  ERREUR ouverture {path}: {e}")
        return []

    result = {
        "code": "",
        "version": "",
        "finalite": "",
        "missions": "",
        "kpis": "",
        "formation_exp": "",
        "filiere_evolution": "",
        "competences": [],
        "libelle_override": "",
        "entite_override": "",
        "n1_override": "",
    }

    tables = []
    for t in d.tables:
        tables.append([[c.text.strip() for c in r.cells] for r in t.rows])

    pending = None
    i = 0
    while i < len(tables):
        rows = tables[i]
        row0 = rows[0]
        joined0 = norm_up(" | ".join(row0))

        if pending == "finalite":
            result["finalite"] = " ".join(x for x in row0 if x).strip()
            pending = None
            i += 1
            continue
        if pending == "matrice":
            result["competences"] = parse_comp_rows(rows[1:])
            pending = None
            i += 1
            continue

        if "CODE" in joined0 and ("VER" in joined0):
            m = re.search(r"CODE\s*:?\s*(.*)", norm_up(row0[1]) if len(row0) > 1 else "")
            result["code"] = row0[1].split(":", 1)[-1].strip() if len(row0) > 1 and ":" in row0[1] else (row0[1] if len(row0) > 1 else "")
            result["version"] = row0[2].split(":", 1)[-1].strip() if len(row0) > 2 and ":" in row0[2] else (row0[2] if len(row0) > 2 else "")
            i += 1
            continue

        if "FINALITE DU POSTE" in joined0:
            if len(rows) > 1 and any(x for x in rows[1]):
                result["finalite"] = " ".join(x for x in rows[1] if x).strip()
            else:
                pending = "finalite"
            i += 1
            continue

        if "MISSIONS PRINCIPALES" in joined0 and len(row0) >= 2:
            if len(rows) > 1:
                result["missions"] = rows[1][0] if len(rows[1]) > 0 else ""
                result["kpis"] = rows[1][1] if len(rows[1]) > 1 else ""
            i += 1
            continue

        if "FORMATION" in joined0 and len(row0) >= 2:
            if len(rows) > 1:
                result["formation_exp"] = rows[1][0] if len(rows[1]) > 0 else ""
                result["filiere_evolution"] = rows[1][1] if len(rows[1]) > 1 else ""
            i += 1
            continue

        if "MATRICE DES COMPETENCES" in joined0:
            if len(rows) > 2 and "COMPETENCE" in norm_up(" ".join(rows[1])):
                result["competences"] = parse_comp_rows(rows[2:])
            elif len(rows) > 1 and "CATEGORIE" in norm_up(" ".join(rows[1])):
                result["competences"] = parse_comp_rows(rows[2:])
            else:
                pending = "matrice"
            i += 1
            continue

        if "CATEGORIE" in joined0 and "COMPETENCE" in joined0 and "NIV" in joined0:
            result["competences"] = parse_comp_rows(rows[1:])
            i += 1
            continue

        # Alternate simple template: key/value table (INTITULÉ DU POSTE / DIRECTION / RATTACHEMENT)
        if "INTITULE DU POSTE" in joined0:
            for r in rows:
                if len(r) < 2:
                    continue
                k = norm_up(r[0])
                v = r[1].strip()
                if "INTITULE DU POSTE" in k:
                    result["libelle_override"] = v
                elif k == "DIRECTION":
                    result["entite_override"] = v
                elif "RATTACHEMENT" in k:
                    result["n1_override"] = v
            i += 1
            continue

        # Alternate simple template: competency table without CATEGORIE column
        # (category given as a full repeated-value row, e.g. ['MÉTIER','MÉTIER','MÉTIER'])
        if "COMPETENCE" in joined0 and "NIVEAU" in joined0 and "DEFINITION" in joined0 and "CATEGORIE" not in joined0:
            current_cat = ""
            for r in rows[1:]:
                cells = [c.strip() for c in r]
                if not any(cells):
                    continue
                if len(set(c for c in cells if c)) == 1:
                    current_cat = cells[0]
                    continue
                comp = cells[0] if len(cells) > 0 else ""
                niv = cells[1] if len(cells) > 1 else ""
                if not comp:
                    continue
                label = f"{current_cat}: {comp}".strip(": ")
                if niv:
                    label += f" ({niv})"
                result["competences"].append(label)
            i += 1
            continue

        i += 1

    libelle = os.path.splitext(os.path.basename(path))[0]
    libelle = re.sub(r"^DP[_ ]?SONASID[_ ]?", "", libelle, flags=re.I)
    libelle = libelle.replace("_", " ").strip()
    if result["libelle_override"]:
        libelle = result["libelle_override"].title()

    row = {
        "Intitulé du poste": libelle,
        "Direction / Entité": normalize_direction(result["entite_override"]) or normalize_direction(guess_direction(path)),
        "Métier du poste": "",
        "Code poste": result["code"],
        "Version": result["version"],
        "N+1 hiérarchique": result["n1_override"],
        "Finalité du poste": result["finalite"],
        "Dimension & enjeux": "",
        "Missions principales": result["missions"],
        "Indicateurs de performance (KPIs)": result["kpis"],
        "Interlocuteurs internes": "",
        "Interlocuteurs externes": "",
        "Supervision directe (effectif)": "",
        "Supervision indirecte (effectif)": "",
        "Niveau de diplôme": result["formation_exp"],
        "Expérience globale": "",
        "Expérience spécifique": "",
        "Compétences requises": "\n".join(result["competences"]),
        "Filière(s) d'évolution": result["filiere_evolution"],
        "Type source": "Word",
        "Fichier source": path,
        "Onglet source": "",
    }
    return [row]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

COLUMNS = [
    "Intitulé du poste",
    "Direction / Entité",
    "Métier du poste",
    "Code poste",
    "Version",
    "N+1 hiérarchique",
    "Finalité du poste",
    "Dimension & enjeux",
    "Missions principales",
    "Indicateurs de performance (KPIs)",
    "Interlocuteurs internes",
    "Interlocuteurs externes",
    "Supervision directe (effectif)",
    "Supervision indirecte (effectif)",
    "Niveau de diplôme",
    "Expérience globale",
    "Expérience spécifique",
    "Compétences requises",
    "Filière(s) d'évolution",
    "Type source",
    "Fichier source",
    "Onglet source",
]


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read())
    return h.hexdigest()


def main():
    all_rows = []
    seen_hashes = {}
    xlsx_files = sorted(glob.glob(os.path.join(ROOT, "**", "*.xlsx"), recursive=True))
    docx_files = sorted(glob.glob(os.path.join(ROOT, "**", "*.docx"), recursive=True))
    docx_files = [p for p in docx_files if not os.path.basename(p).startswith("~$")]
    xlsx_files = [p for p in xlsx_files if not os.path.basename(p).startswith("~$")]

    print(f"{len(xlsx_files)} fichiers xlsx, {len(docx_files)} fichiers docx")

    dup_count = 0
    for p in xlsx_files:
        h = file_hash(p)
        if h in seen_hashes:
            dup_count += 1
            print(f"  DUPLICATA (identique) ignoré: {p}  (== {seen_hashes[h]})")
            continue
        seen_hashes[h] = p
        rows = process_xlsx(p)
        all_rows.extend(rows)

    for p in docx_files:
        h = file_hash(p)
        if h in seen_hashes:
            dup_count += 1
            print(f"  DUPLICATA (identique) ignoré: {p}  (== {seen_hashes[h]})")
            continue
        seen_hashes[h] = p
        rows = process_docx(p)
        all_rows.extend(rows)

    print(f"\n{len(all_rows)} fiches de poste extraites ({dup_count} doublons de fichiers ignorés)")

    # dedupe by (normalized libelle, direction) - keep the richer record
    def richness(r):
        return sum(len(str(v)) for v in r.values())

    by_key = {}
    order = []
    for r in all_rows:
        key = norm_label(r["Intitulé du poste"])
        if key not in by_key:
            by_key[key] = r
            order.append(key)
        else:
            prev = by_key[key]
            if norm_label(prev["Direction / Entité"]) != norm_label(r["Direction / Entité"]):
                print(f"  ATTENTION: même intitulé '{r['Intitulé du poste']}' avec directions différentes: "
                      f"'{prev['Direction / Entité']}' ({prev['Fichier source']}) vs "
                      f"'{r['Direction / Entité']}' ({r['Fichier source']})")
            if richness(r) > richness(prev):
                by_key[key] = r

    final_rows = [by_key[k] for k in order]
    print(f"{len(final_rows)} fiches après déduplication par (intitulé, direction)")

    # write excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fiches de poste"
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="4472C4")
        cell.alignment = openpyxl.styles.Alignment(vertical="center", wrap_text=True)
    for r in sorted(final_rows, key=lambda x: (x["Direction / Entité"], x["Intitulé du poste"])):
        ws.append([r.get(c, "") for c in COLUMNS])

    widths = {
        "Intitulé du poste": 32,
        "Direction / Entité": 26,
        "Métier du poste": 14,
        "Code poste": 14,
        "Version": 10,
        "N+1 hiérarchique": 22,
        "Finalité du poste": 40,
        "Dimension & enjeux": 40,
        "Missions principales": 55,
        "Indicateurs de performance (KPIs)": 45,
        "Interlocuteurs internes": 30,
        "Interlocuteurs externes": 30,
        "Supervision directe (effectif)": 12,
        "Supervision indirecte (effectif)": 12,
        "Niveau de diplôme": 30,
        "Expérience globale": 20,
        "Expérience spécifique": 30,
        "Compétences requises": 60,
        "Filière(s) d'évolution": 30,
        "Type source": 10,
        "Fichier source": 50,
        "Onglet source": 18,
    }
    for i, col in enumerate(COLUMNS, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = widths.get(col, 20)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"

    wb.save(OUT)
    print(f"\nFichier écrit: {OUT}")


if __name__ == "__main__":
    main()
