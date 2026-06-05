"""
importer_bilan.py — Import automatique bilan WinBooks (XLS ou PDF)
Fidunot Expertise SRL

Formats supportés :
  - XLS WinBooks (3 onglets Actif / Passif / Résultats)
  - PDF bilan provisoire WinBooks (extraction positionnelle)

Retourne dict {champ_calculateur: (valeur_N, valeur_NM1)}.
"""

import re, os
from collections import defaultdict

# ─── Mapping PCMN ─────────────────────────────────────────────────────────────

PCMN_MAP = {
    "21/28":  "immo_nettes",
    "3":      "stocks",
    "40/41":  "creances",
    "54/58":  "disponibilites",
    "10/15":  "fp",
    "13":     "reserves",
    "42/48":  "dettes_ct",
    "17/49":  "_total_dettes",
    "70":     "ca",
    "61":     "_achats_raw",
    "630":    "_amort_raw",
    "65":     "_interets_raw",
    "65/66B": "_interets_raw",
}

LABEL_MAP = {
    "montant total de l'actif":     "total_bilan",
    "montant total du passif":      "total_bilan",
    "bénéfice d'exploitation":      "_ebit",
    "benefice d'exploitation":      "_ebit",
    "perte d'exploitation":         "_ebit_perte",
    "bénéfice de l'exercice avant": "_res_avant_impots",
    "bénéfice de l'exercice":       "_res_net_candidat",
    "benefice de l'exercice":       "_res_net_candidat",
    "perte de l'exercice":          "_res_perte",
}

FOURN_PCMN = {"44", "440/4", "440/9"}


def _to_float(v):
    if v is None: return None
    if isinstance(v, (int, float)):
        return None if (v != v) else float(v)
    s = (str(v).strip()
         .replace("\xa0", "").replace("\u202f", "")
         .replace("(", "-").replace(")", ""))
    # Format belge : 64.054,90 → 64054.90
    if re.match(r"^-?\d{1,3}(\.\d{3})+(,\d{1,2})?$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _finalize(raw):
    result = {}
    for k, v in raw.items():
        if not k.startswith("_"):
            result[k] = v

    for rk, dk in [("_achats_raw","achats"), ("_amort_raw","amort"), ("_interets_raw","interets")]:
        if rk in raw:
            n, nm1 = raw[rk]
            result[dk] = (abs(n) if n is not None else None,
                           abs(nm1) if nm1 is not None else None)

    if "_total_dettes" in raw and "dettes_ct" in result:
        tdn, tdnm1   = raw["_total_dettes"]
        dctn, dctnm1 = result["dettes_ct"]
        result["dettes_lt"] = (
            round(tdn - dctn, 2)     if (tdn   is not None and dctn   is not None) else None,
            round(tdnm1 - dctnm1, 2) if (tdnm1 is not None and dctnm1 is not None) else None,
        )

    if "_ebit" in raw:
        n, nm1 = raw["_ebit"]
        pn = pnm1 = 0
        if "_ebit_perte" in raw:
            pn, pnm1 = raw["_ebit_perte"]
            pn = pn or 0; pnm1 = pnm1 or 0
        result["ebit"] = ((n or 0) - pn, (nm1 or 0) - pnm1)

    if "_res_net_candidat" in raw:
        n, nm1 = raw["_res_net_candidat"]
        pn = pnm1 = 0
        if "_res_perte" in raw:
            pn, pnm1 = raw["_res_perte"]
            pn = pn or 0; pnm1 = pnm1 or 0
        result["res_net"] = ((n or 0) - pn, (nm1 or 0) - pnm1)

    return result


# ─── Parser XLS ───────────────────────────────────────────────────────────────

def parse_xls(filepath):
    """
    WinBooks XLS — 3 onglets, 8 colonnes.
    N    = premier non-NaN parmi cols 2, 3
    N-1  = premier non-NaN parmi cols 6, 7
    """
    try:
        import pandas as pd
    except ImportError:
        raise ImportError("Installer : pip install pandas xlrd")

    ext = os.path.splitext(filepath)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    xls = pd.ExcelFile(filepath, engine=engine)
    raw = {}

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        df = df.where(df.notna(), None)

        for _, row in df.iterrows():
            row = list(row)
            if len(row) < 3: continue

            label = str(row[0]).strip() if row[0] is not None else ""
            pcmn  = re.sub(r"\.0$", "", str(row[1]).strip()) if row[1] is not None else ""

            vn   = next((_to_float(row[c]) for c in (2, 3)
                          if c < len(row) and _to_float(row[c]) is not None), None)
            vnm1 = next((_to_float(row[c]) for c in (6, 7)
                          if c < len(row) and _to_float(row[c]) is not None), None)

            if pcmn in PCMN_MAP and PCMN_MAP[pcmn] not in raw:
                raw[PCMN_MAP[pcmn]] = (vn, vnm1)

            if pcmn in FOURN_PCMN and "dettes_fourn" not in raw and vn is not None:
                raw["dettes_fourn"] = (vn, vnm1)

            label_low = label.lower()
            for pattern, dest in LABEL_MAP.items():
                if label_low.startswith(pattern) and dest not in raw:
                    raw[dest] = (vn, vnm1)
                    break

    return _finalize(raw)


# ─── Parser PDF (extraction positionnelle) ────────────────────────────────────

def parse_pdf(filepath):
    """
    WinBooks PDF — extraction par position X des mots.
    Structure du document (largeur ~595pt) :
      label   : x < 340
      PCMN    : 330 < x < 380
      N       : 380 < x < 490
      N-1     : x > 490
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("Installer : pip install pdfplumber")

    # Seuils X (ajustables si la mise en page diffère)
    X_PCMN_MIN  = 330
    X_PCMN_MAX  = 385
    X_N_MIN     = 385
    X_N_MAX     = 495
    X_NM1_MIN   = 495

    raw = {}

    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=5, y_tolerance=3)
            if not words: continue

            # Grouper par y (tolérance 4pt)
            by_y = defaultdict(list)
            for w in words:
                y = round(float(w["top"]) / 4) * 4
                by_y[y].append(w)

            # Traiter chaque ligne
            for y in sorted(by_y.keys()):
                ws = sorted(by_y[y], key=lambda w: w["x0"])

                label_parts, pcmn_parts, n_parts, nm1_parts = [], [], [], []
                for w in ws:
                    x = float(w["x0"])
                    t = w["text"]
                    if x > X_NM1_MIN:
                        nm1_parts.append(t)
                    elif x > X_N_MIN:
                        n_parts.append(t)
                    elif x > X_PCMN_MIN:
                        pcmn_parts.append(t)
                    else:
                        label_parts.append(t)

                label = " ".join(label_parts).strip()
                pcmn  = " ".join(pcmn_parts).strip()
                vn    = _to_float(" ".join(n_parts))
                vnm1  = _to_float(" ".join(nm1_parts))

                # Match PCMN
                if pcmn in PCMN_MAP and PCMN_MAP[pcmn] not in raw:
                    raw[PCMN_MAP[pcmn]] = (vn, vnm1)

                if pcmn in FOURN_PCMN and "dettes_fourn" not in raw and vn is not None:
                    raw["dettes_fourn"] = (vn, vnm1)

                # Match libellé
                label_low = label.lower()
                for pattern, dest in LABEL_MAP.items():
                    if label_low.startswith(pattern) and dest not in raw:
                        raw[dest] = (vn, vnm1)
                        break

    return _finalize(raw)


# ─── Point d'entrée ──────────────────────────────────────────────────────────

def importer_bilan(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xls", ".xlsx", ".xlsm"):
        return parse_xls(filepath)
    elif ext == ".pdf":
        return parse_pdf(filepath)
    else:
        raise ValueError(f"Format non supporté : {ext}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage : python importer_bilan.py <fichier.xls|fichier.pdf>")
        sys.exit(1)
    data = importer_bilan(sys.argv[1])
    print(f"\n{'Champ':<25} {'N':>12}  {'N-1':>12}")
    print("-" * 52)
    for k, (n, nm1) in sorted(data.items()):
        print(f"  {k:<23} {str(round(n,2) if n else '—'):>12}  {str(round(nm1,2) if nm1 else '—'):>12}")
