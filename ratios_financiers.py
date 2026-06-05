#!/usr/bin/env python3
"""
Calculateur de Ratios Financiers — PME Belges
Fidunot Expertise SRL — Usage confidentiel local
Données traitées uniquement en local, aucune connexion réseau.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
from pathlib import Path
try:
    from importer_bilan import importer_bilan as _import_bilan_fn
    IMPORT_DISPONIBLE = True
except ImportError:
    IMPORT_DISPONIBLE = False

# Chargement du fichier .env (optionnel — fonctionne sans si absent)
try:
    from dotenv import load_dotenv
    _env_path = Path(__file__).parent / ".env"
    load_dotenv(dotenv_path=_env_path, override=False)
except ImportError:
    pass  # python-dotenv non installe : valeurs par defaut utilisees

def _cfg(key, default=""):
    """Lit une variable d'environnement avec valeur par defaut."""
    return os.environ.get(key, default).strip()

# ─── Constantes (alimentees par .env) ────────────────────────────────────────

CABINET        = _cfg("CABINET_NOM",          "Fidunot Expertise SRL")
CABINET_VILLE  = _cfg("CABINET_VILLE",         "Wavre")
APP_TITLE      = "Calculateur de Ratios Financiers — PME Belges"
VERSION        = "1.0"

TYPE_DEFAUT    = _cfg("TYPE_CLIENT_DEFAUT",    "Services")
EXERCICE_N     = _cfg("EXERCICE_N",            str(datetime.now().year))
EXERCICE_NM1   = _cfg("EXERCICE_NM1",          str(datetime.now().year - 1))
EXPORT_DOSSIER = _cfg("EXPORT_DOSSIER",        "")
TVA_TAUX       = float(_cfg("TVA_TAUX",        "0.21"))

CLIENT_TYPES = ["Services", "Industrie / Commerce", "Holding"]

COL_NAVY    = "#0d1e35"
COL_NAVY2   = "#1a3452"
COL_LIGHT   = "#f4f6f9"
COL_WHITE   = "#ffffff"
COL_VERT    = "#1a6b3a"
COL_ORANGE  = "#8a5a00"
COL_ROUGE   = "#8b1a1a"
BG_VERT     = "#e8f7ee"
BG_ORANGE   = "#fff4e0"
BG_ROUGE    = "#fde8e8"

CAT_COLORS = {
    "Rentabilité":    "#1a6b5e",
    "Liquidité":      "#1e4d8c",
    "Solvabilité":    "#7b2d00",
    "Activité & BFR": "#5c3d8f",
    "Structure":      "#2d6040",
}

# ─── Seuils numériques ────────────────────────────────────────────────────────
# Format : ('higher_better', seuil_vert, seuil_orange)
#          ('lower_better',  seuil_vert, seuil_orange)
#          ('range',         vert_min, vert_max, orange_min, orange_max)

THRESHOLDS = {
    "Services": {
        "marge_brute":        ("higher_better", 40, 20),
        "marge_ebitda":       ("higher_better", 15, 8),
        "marge_nette":        ("higher_better", 8, 3),
        "roe":                ("higher_better", 12, 5),
        "roa":                ("higher_better", 8, 3),
        "liq_generale":       ("higher_better", 1.5, 1.0),
        "liq_reduite":        ("higher_better", 1.0, 0.7),
        "liq_immediate":      ("higher_better", 0.3, 0.1),
        "endettement_global": ("lower_better",  50,  70),
        "gearing":            ("lower_better",  0.5, 1.5),
        "couv_interets":      ("higher_better", 5,   2),
        "cap_remboursement":  ("lower_better",  2.5, 4),
        "dso":                ("lower_better",  30,  60),
        "dpo":                ("range",         30, 60, 15, 90),
        "rotation_stocks":    ("lower_better",  30,  60),
        "bfr_jours":          ("lower_better",  30,  60),
        "autonomie":          ("higher_better", 35,  20),
        "couv_immo":          ("higher_better", 1.0, 0.5),
        "frn":                ("higher_better", 0.01, -0.01),
    },
    "Industrie / Commerce": {
        "marge_brute":        ("higher_better", 25, 10),
        "marge_ebitda":       ("higher_better", 12, 5),
        "marge_nette":        ("higher_better", 5,  2),
        "roe":                ("higher_better", 10, 4),
        "roa":                ("higher_better", 5,  2),
        "liq_generale":       ("higher_better", 1.5, 1.0),
        "liq_reduite":        ("higher_better", 0.8, 0.5),
        "liq_immediate":      ("higher_better", 0.2, 0.1),
        "endettement_global": ("lower_better",  55,  75),
        "gearing":            ("lower_better",  0.8, 2.0),
        "couv_interets":      ("higher_better", 4,   2),
        "cap_remboursement":  ("lower_better",  3,   5),
        "dso":                ("lower_better",  45,  75),
        "dpo":                ("range",         30, 60, 15, 90),
        "rotation_stocks":    ("lower_better",  45,  90),
        "bfr_jours":          ("lower_better",  45,  75),
        "autonomie":          ("higher_better", 30,  15),
        "couv_immo":          ("higher_better", 0.8, 0.4),
        "frn":                ("higher_better", 0.01, -0.01),
    },
    "Holding": {
        "marge_brute":        ("higher_better", 60, 30),
        "marge_ebitda":       ("higher_better", 40, 20),
        "marge_nette":        ("higher_better", 20, 8),
        "roe":                ("higher_better", 8,  3),
        "roa":                ("higher_better", 5,  2),
        "liq_generale":       ("higher_better", 1.2, 0.8),
        "liq_reduite":        ("higher_better", 1.0, 0.7),
        "liq_immediate":      ("higher_better", 0.3, 0.1),
        "endettement_global": ("lower_better",  60,  80),
        "gearing":            ("lower_better",  1.0, 2.5),
        "couv_interets":      ("higher_better", 3,   1.5),
        "cap_remboursement":  ("lower_better",  4,   7),
        "dso":                ("lower_better",  30,  60),
        "dpo":                ("range",         20, 45, 10, 75),
        "rotation_stocks":    None,
        "bfr_jours":          ("lower_better",  15,  30),
        "autonomie":          ("higher_better", 25,  10),
        "couv_immo":          ("higher_better", 1.0, 0.5),
        "frn":                ("higher_better", 0.01, -0.01),
    },
}

# (Libellé, Unité, Catégorie)
RATIO_META = {
    "marge_brute":        ("Marge brute",              "%",    "Rentabilité"),
    "marge_ebitda":       ("Marge EBITDA",             "%",    "Rentabilité"),
    "marge_nette":        ("Marge nette",              "%",    "Rentabilité"),
    "roe":                ("ROE",                      "%",    "Rentabilité"),
    "roa":                ("ROA",                      "%",    "Rentabilité"),
    "liq_generale":       ("Liquidité générale",       "×",    "Liquidité"),
    "liq_reduite":        ("Liquidité réduite",        "×",    "Liquidité"),
    "liq_immediate":      ("Liquidité immédiate",      "×",    "Liquidité"),
    "endettement_global": ("Taux d'endettement",       "%",    "Solvabilité"),
    "gearing":            ("Gearing",                  "×",    "Solvabilité"),
    "couv_interets":      ("Couverture des intérêts",  "×",    "Solvabilité"),
    "cap_remboursement":  ("Capacité remboursement",   " ans", "Solvabilité"),
    "dso":                ("Délai clients (DSO)",       " j",   "Activité & BFR"),
    "dpo":                ("Délai fournisseurs (DPO)", " j",   "Activité & BFR"),
    "rotation_stocks":    ("Rotation des stocks",      " j",   "Activité & BFR"),
    "bfr_jours":          ("BFR en jours de CA",       " j",   "Activité & BFR"),
    "frn":                ("Fonds de roulement net",   " €",   "Structure"),
    "autonomie":          ("Autonomie financière",     "%",    "Structure"),
    "couv_immo":          ("Couverture des immo.",     "×",    "Structure"),
}

CHAMPS_SAISIE = [
    ("Compte de résultats", [
        ("ca",           "Chiffre d'affaires (70)",              "€"),
        ("achats",       "Achats & services directs (60+61)",    "€"),
        ("ebit",         "EBIT / Résultat d'exploitation",       "€"),
        ("amort",        "Amortissements & dépréciations (630+660)", "€"),
        ("interets",     "Charges d'intérêts (650)",             "€"),
        ("res_net",      "Résultat net",                         "€"),
    ]),
    ("Bilan — Actif", [
        ("total_bilan",     "Total bilan",                       "€"),
        ("immo_nettes",     "Immobilisations nettes (20-28)",    "€"),
        ("stocks",          "Stocks (3)",                        "€"),
        ("creances",        "Créances clients HTVA (40)",        "€"),
        ("disponibilites",  "Disponibilités (54-58)",            "€"),
    ]),
    ("Bilan — Passif", [
        ("fp",           "Fonds propres (10-15)",                "€"),
        ("dettes_lt",    "Dettes financières LT (17)",           "€"),
        ("dettes_ct",    "Dettes à court terme (total)",         "€"),
        ("dettes_fourn", "Dettes fournisseurs (44)",             "€"),
    ]),
    ("Données complémentaires", [
        ("camv",     "Coût d'achat des marchandises vendues (60)", "€"),
        ("reserves", "Réserves & bénéfices reportés (13+14)",      "€"),
    ]),
]

# ─── Calculs financiers ────────────────────────────────────────────────────────

def _div(a, b):
    if b is None or b == 0:
        return None
    return a / b

def _pct(a, b):
    r = _div(a, b)
    return r * 100 if r is not None else None

def calculate_ratios(d):
    ca    = d.get("ca")    or 0
    ach   = d.get("achats") or 0
    ebit  = d.get("ebit")  or 0
    amort = d.get("amort") or 0
    int_  = d.get("interets") or 0
    rnet  = d.get("res_net") or 0
    tb    = d.get("total_bilan") or 0
    immo  = d.get("immo_nettes") or 0
    stk   = d.get("stocks") or 0
    cre   = d.get("creances") or 0
    dispo = d.get("disponibilites") or 0
    fp    = d.get("fp")    or 0
    dlt   = d.get("dettes_lt") or 0
    dct   = d.get("dettes_ct") or 0
    dfou  = d.get("dettes_fourn") or 0
    camv  = d.get("camv")  or 0

    ebitda   = ebit + amort
    actcirc  = stk + cre + dispo
    dnettes  = dlt - dispo
    tdettes  = dct + dlt
    caperm   = fp + dlt
    frn_val  = caperm - immo
    ca_ttc   = ca  * (1 + TVA_TAUX)
    ach_ttc  = ach * (1 + TVA_TAUX)

    r = {}
    r["marge_brute"]        = _pct(ca - ach, ca)
    r["marge_ebitda"]       = _pct(ebitda, ca)
    r["marge_nette"]        = _pct(rnet, ca)
    r["roe"]                = _pct(rnet, fp)
    r["roa"]                = _pct(ebit, tb)
    r["liq_generale"]       = _div(actcirc, dct)
    r["liq_reduite"]        = _div(actcirc - stk, dct)
    r["liq_immediate"]      = _div(dispo, dct)
    r["endettement_global"] = _pct(tdettes, tb)
    r["gearing"]            = _div(dnettes, fp)
    r["couv_interets"]      = _div(ebit, int_)
    r["cap_remboursement"]  = _div(dnettes, ebitda)
    r["dso"]                = (_div(cre, ca_ttc) * 365)  if ca_ttc else None
    r["dpo"]                = (_div(dfou, ach_ttc) * 365) if ach_ttc else None
    r["rotation_stocks"]    = (_div(stk, camv) * 365)    if camv  else None
    bfr = stk + cre - dfou
    r["bfr_jours"]          = (_div(bfr, ca_ttc) * 365)  if ca_ttc else None
    r["frn"]                = frn_val if (fp or dlt or immo) else None
    r["autonomie"]          = _pct(fp, tb)
    r["couv_immo"]          = _div(fp, immo)
    return r

def eval_ratio(key, value, ct):
    if value is None:
        return None
    t = THRESHOLDS.get(ct, {}).get(key)
    if not t:
        return None
    if t[0] == "higher_better":
        return "vert" if value >= t[1] else ("orange" if value >= t[2] else "rouge")
    if t[0] == "lower_better":
        return "vert" if value <= t[1] else ("orange" if value <= t[2] else "rouge")
    if t[0] == "range":
        if t[1] <= value <= t[2]: return "vert"
        if t[3] <= value <= t[4]: return "orange"
        return "rouge"
    return None

def get_trend(key, vn, vnm1, ct):
    if vn is None or vnm1 is None:
        return None
    diff = vn - vnm1
    if abs(diff) < 1e-9:
        return "stable"
    t = THRESHOLDS.get(ct, {}).get(key)
    if not t:
        return None
    if t[0] == "range":
        order = {"vert": 0, "orange": 1, "rouge": 2, None: 1}
        cn  = eval_ratio(key, vn,   ct)
        cnm1= eval_ratio(key, vnm1, ct)
        d   = order[cn] - order[cnm1]
        return "amelioration" if d < 0 else ("degradation" if d > 0 else "stable")
    better = diff > 0 if t[0] == "higher_better" else diff < 0
    return "amelioration" if better else "degradation"

def calc_altman(vals):
    try:
        x = [float(str(vals.get(f"x{i}", 0) or 0).replace(",", "."))
             for i in range(1, 6)]
        z = 0.717*x[0] + 0.847*x[1] + 3.107*x[2] + 0.420*x[3] + 0.998*x[4]
        if z > 2.9:    zone = ("vert",   "Zone sûre")
        elif z < 1.23: zone = ("rouge",  "Zone de détresse")
        else:          zone = ("orange", "Zone grise")
        return z, zone
    except (ValueError, TypeError, ZeroDivisionError):
        return None, None

# ─── Interface graphique ──────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} — {CABINET} — {CABINET_VILLE}")
        self.root.geometry("1080x800")
        self.root.minsize(900, 650)
        self.root.configure(bg=COL_NAVY)

        self.v_nom  = tk.StringVar()
        self.v_type = tk.StringVar(value=TYPE_DEFAUT)
        self.v_n    = tk.StringVar(value=EXERCICE_N)
        self.v_nm1  = tk.StringVar(value=EXERCICE_NM1)

        self.flds_n   = {}
        self.flds_nm1 = {}
        self.alt_flds = {}
        self.rlabels  = {}

        self.ratios_n   = {}
        self.ratios_nm1 = {}

        self._build()

    # ── Helpers de style ─────────────────────────────────────────────────────

    def _label(self, parent, text, **kw):
        return tk.Label(parent, text=text, bg=kw.pop("bg", COL_LIGHT),
                        font=kw.pop("font", ("Arial", 9)), **kw)

    def _entry(self, parent, var, width=14):
        e = tk.Entry(parent, textvariable=var, width=width,
                     font=("Courier", 10), justify="right",
                     relief="flat", bd=1,
                     highlightthickness=1, highlightbackground="#c8d4e0")
        e.bind("<FocusOut>", lambda _e: self._recalc())
        e.bind("<Return>",   lambda _e: self._recalc())
        return e

    def _scrollable(self, parent):
        canvas = tk.Canvas(parent, bg=COL_LIGHT, highlightthickness=0)
        sb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=COL_LIGHT)
        cid = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cid, width=e.width))
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Enter>", lambda _e, c=canvas: c.bind_all(
            "<MouseWheel>",
            lambda e, c=c: c.yview_scroll(-1 * int(e.delta / 120), "units")))
        canvas.bind("<Leave>", lambda _e, c=canvas: c.unbind_all("<MouseWheel>"))
        return inner

    # ── Construction globale ─────────────────────────────────────────────────

    def _build(self):
        self._build_header()
        self._build_client_bar()
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True)
        self._tab_saisie()
        self._tab_ratios()
        self._tab_altman()
        self._build_footer()

    def _build_header(self):
        f = tk.Frame(self.root, bg=COL_NAVY, pady=10, padx=20)
        f.pack(fill="x")
        tk.Label(f, text=f"{CABINET}  —  {CABINET_VILLE}", font=("Arial", 8), fg="#7a9abf",
                 bg=COL_NAVY).pack(anchor="w")
        tk.Label(f, text="Calculateur de Ratios Financiers — PME Belges",
                 font=("Arial", 15, "bold"), fg="white", bg=COL_NAVY).pack(anchor="w")
        tk.Label(f,
                 text="Traitement 100 % local — aucune donnée transmise à l'extérieur",
                 font=("Arial", 8), fg="#3a6a8a", bg=COL_NAVY).pack(anchor="w")

    def _build_client_bar(self):
        f = tk.Frame(self.root, bg=COL_NAVY2, padx=14, pady=7)
        f.pack(fill="x")

        def lbl(text):
            return tk.Label(f, text=text, fg="#7a9abf", bg=COL_NAVY2,
                            font=("Arial", 9))

        lbl("Client :").grid(row=0, column=0, padx=(0, 4), sticky="e")
        tk.Entry(f, textvariable=self.v_nom, width=22,
                 font=("Arial", 10), relief="flat",
                 highlightthickness=1, highlightbackground="#2d5a8a"
                 ).grid(row=0, column=1, padx=(0, 14))

        lbl("Type :").grid(row=0, column=2, padx=(0, 4), sticky="e")
        cb = ttk.Combobox(f, textvariable=self.v_type, values=CLIENT_TYPES,
                          width=19, state="readonly")
        cb.grid(row=0, column=3, padx=(0, 14))
        cb.bind("<<ComboboxSelected>>", lambda _: self._recalc())

        lbl("Exercice N :").grid(row=0, column=4, padx=(0, 4), sticky="e")
        tk.Entry(f, textvariable=self.v_n, width=7,
                 font=("Arial", 10), relief="flat",
                 highlightthickness=1, highlightbackground="#2d5a8a"
                 ).grid(row=0, column=5, padx=(0, 14))

        lbl("Exercice N-1 :").grid(row=0, column=6, padx=(0, 4), sticky="e")
        tk.Entry(f, textvariable=self.v_nm1, width=7,
                 font=("Arial", 10), relief="flat",
                 highlightthickness=1, highlightbackground="#2d5a8a"
                 ).grid(row=0, column=7)

    # ── Onglet Saisie ────────────────────────────────────────────────────────

    def _tab_saisie(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  Saisie des données  ")
        inner = self._scrollable(tab)

        # En-tête des colonnes
        hdr = tk.Frame(inner, bg="#2d4a6a", pady=6, padx=16)
        hdr.pack(fill="x", pady=(0, 2))
        tk.Label(hdr, text="Poste comptable  (PCMN)", width=44, anchor="w",
                 font=("Arial", 9, "bold"), fg="white", bg="#2d4a6a").grid(row=0, column=0)
        lbl_n = tk.Label(hdr, text=f"N ({self.v_n.get()})", width=16, anchor="center",
                          font=("Arial", 9, "bold"), fg="#6fcf97", bg="#2d4a6a")
        lbl_n.grid(row=0, column=1)
        lbl_nm1 = tk.Label(hdr, text=f"N-1 ({self.v_nm1.get()})", width=16, anchor="center",
                            font=("Arial", 9, "bold"), fg="#f2c94c", bg="#2d4a6a")
        lbl_nm1.grid(row=0, column=2)
        self.v_n.trace_add("write",
            lambda *_: lbl_n.config(text=f"N ({self.v_n.get()})"))
        self.v_nm1.trace_add("write",
            lambda *_: lbl_nm1.config(text=f"N-1 ({self.v_nm1.get()})"))

        for section, fields in CHAMPS_SAISIE:
            # Titre de section
            sh = tk.Frame(inner, bg="#dce8f4", pady=5, padx=16)
            sh.pack(fill="x", pady=(8, 0))
            tk.Label(sh, text=section, font=("Arial", 10, "bold"),
                     fg="#1a3452", bg="#dce8f4").pack(anchor="w")

            for key, label, unit in fields:
                rf = tk.Frame(inner, bg=COL_WHITE, pady=4, padx=16)
                rf.pack(fill="x", pady=1)

                tk.Label(rf, text=f"{label}   ({unit})", width=50, anchor="w",
                         font=("Arial", 9), bg=COL_WHITE,
                         fg="#2a3a4a").grid(row=0, column=0, sticky="w")

                vn = tk.StringVar(); self.flds_n[key]   = vn
                self._entry(rf, vn).grid(row=0, column=1, padx=8)

                vnm1 = tk.StringVar(); self.flds_nm1[key] = vnm1
                self._entry(rf, vnm1).grid(row=0, column=2, padx=8)

    # ── Onglet Ratios ────────────────────────────────────────────────────────

    def _tab_ratios(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  Ratios calculés  ")
        self._ratios_inner = self._scrollable(tab)
        self._build_ratios_table()

    def _build_ratios_table(self):
        inner = self._ratios_inner
        for w in inner.winfo_children():
            w.destroy()
        self.rlabels = {}

        hdr = tk.Frame(inner, bg="#2d4a6a", pady=6, padx=16)
        hdr.pack(fill="x", pady=(0, 2))
        cols = [
            ("Ratio",      26, "white"),
            ("Valeur N",   12, "#6fcf97"),
            ("Statut N",   12, "#6fcf97"),
            ("Valeur N-1", 12, "#f2c94c"),
            ("Statut N-1", 12, "#f2c94c"),
            ("Tendance",    8, "white"),
        ]
        for i, (h, w, fg) in enumerate(cols):
            tk.Label(hdr, text=h, width=w, font=("Arial", 9, "bold"),
                     fg=fg, bg="#2d4a6a", anchor="center").grid(row=0, column=i, padx=4)

        current_cat = None
        for key, (label, unite, cat) in RATIO_META.items():
            if cat != current_cat:
                current_cat = cat
                ch = tk.Frame(inner, bg=CAT_COLORS[cat], pady=4, padx=16)
                ch.pack(fill="x", pady=(6, 0))
                tk.Label(ch, text=cat, font=("Arial", 10, "bold"),
                         fg="white", bg=CAT_COLORS[cat]).pack(anchor="w")

            rf = tk.Frame(inner, bg=COL_WHITE, pady=3, padx=16)
            rf.pack(fill="x", pady=1)

            tk.Label(rf, text=label, width=26, anchor="w",
                     font=("Arial", 9), bg=COL_WHITE,
                     fg="#1a2533").grid(row=0, column=0, padx=4, sticky="w")

            lv_n   = tk.Label(rf, text="—", width=12, font=("Courier", 9),
                               bg=COL_WHITE, fg="#aaa")
            lv_n.grid(row=0, column=1, padx=4)

            ls_n   = tk.Label(rf, text="", width=12, font=("Arial", 9, "bold"),
                               bg=COL_WHITE)
            ls_n.grid(row=0, column=2, padx=4)

            lv_nm1 = tk.Label(rf, text="—", width=12, font=("Courier", 9),
                               bg=COL_WHITE, fg="#aaa")
            lv_nm1.grid(row=0, column=3, padx=4)

            ls_nm1 = tk.Label(rf, text="", width=12, font=("Arial", 9, "bold"),
                               bg=COL_WHITE)
            ls_nm1.grid(row=0, column=4, padx=4)

            lt = tk.Label(rf, text="", width=8, font=("Arial", 13),
                           bg=COL_WHITE)
            lt.grid(row=0, column=5, padx=4)

            self.rlabels[key] = (lv_n, ls_n, lv_nm1, ls_nm1, lt)

    # ── Onglet Altman ────────────────────────────────────────────────────────

    def _tab_altman(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  Altman Z'  ")
        inner = self._scrollable(tab)

        # Header
        hf = tk.Frame(inner, bg="#3d2800", pady=12, padx=16)
        hf.pack(fill="x")
        tk.Label(hf, text="Score Z' d'Altman — PME non cotée (Altman, 1983)",
                 font=("Arial", 13, "bold"), fg="white", bg="#3d2800").pack(anchor="w")
        tk.Label(hf, text="Z' = 0,717×X1 + 0,847×X2 + 3,107×X3 + 0,420×X4 + 0,998×X5",
                 font=("Courier", 10), fg="#f5cc7a", bg="#3d2800").pack(anchor="w", pady=(4, 0))
        zf = tk.Frame(hf, bg="#3d2800")
        zf.pack(anchor="w", pady=(8, 0))
        for zt, zc in [("Zone sûre   Z' > 2,9", COL_VERT),
                        ("Zone grise  1,23 ≤ Z' ≤ 2,9", COL_ORANGE),
                        ("Zone de détresse  Z' < 1,23", COL_ROUGE)]:
            tk.Label(zf, text=zt, font=("Arial", 9, "bold"),
                     fg="white", bg=zc, padx=10, pady=3).pack(side="left", padx=6)

        # Variables
        altman_def = [
            ("x1", "0,717", "X1 — FRN / Total actif",
             "(Actif circulant − Dettes CT) / Total bilan",
             "→ Auto-calculable depuis les données saisies"),
            ("x2", "0,847", "X2 — Réserves cumulées / Total actif",
             "(13 + 14) / Total bilan",
             "→ Saisir manuellement (réserves + bénéfices reportés)"),
            ("x3", "3,107", "X3 — EBIT / Total actif",
             "Résultat d'exploitation / Total bilan",
             "→ Auto-calculable depuis les données saisies"),
            ("x4", "0,420", "X4 — Fonds propres / Total dettes",
             "(10 à 15) / (17 + 42 à 48)",
             "→ Auto-calculable depuis les données saisies"),
            ("x5", "0,998", "X5 — CA / Total actif",
             "70 / Total bilan",
             "→ Auto-calculable depuis les données saisies"),
        ]

        for key, coef, label, pcmn, hint in altman_def:
            rf = tk.Frame(inner, bg=COL_WHITE, padx=16, pady=9)
            rf.pack(fill="x", pady=1)
            tk.Label(rf, text=f"×{coef}", font=("Arial", 10, "bold"),
                     fg="white", bg=COL_NAVY, padx=6, pady=3,
                     width=6).grid(row=0, column=0, rowspan=3, padx=(0, 14), sticky="n")
            tk.Label(rf, text=label, font=("Arial", 10, "bold"),
                     fg="#1a2533", bg=COL_WHITE).grid(row=0, column=1, sticky="w")
            tk.Label(rf, text=pcmn, font=("Courier", 9),
                     fg="#8a9ab0", bg=COL_WHITE).grid(row=1, column=1, sticky="w")
            tk.Label(rf, text=hint, font=("Arial", 8, "italic"),
                     fg="#4a8a5a", bg=COL_WHITE).grid(row=2, column=1, sticky="w")
            var = tk.StringVar()
            self.alt_flds[key] = var
            e = tk.Entry(rf, textvariable=var, width=14,
                         font=("Courier", 10), justify="right",
                         relief="flat", highlightthickness=1,
                         highlightbackground="#c8d4e0")
            e.grid(row=0, column=2, rowspan=3, padx=20, sticky="n", pady=6)
            e.bind("<FocusOut>", lambda _: self._compute_altman())
            e.bind("<Return>",   lambda _: self._compute_altman())

        # Bouton auto-remplir
        bf = tk.Frame(inner, bg=COL_LIGHT, padx=16, pady=6)
        bf.pack(fill="x")
        tk.Button(bf, text="⟳ Auto-remplir X1, X3, X4, X5 depuis les données saisies",
                  font=("Arial", 9, "bold"), bg="#2d8cff", fg="white",
                  relief="flat", padx=12, pady=6, cursor="hand2",
                  command=self._autofill_altman).pack(side="left")
        tk.Label(bf, text="X2 doit être saisi manuellement.",
                 font=("Arial", 8, "italic"), fg="#8a9ab0", bg=COL_LIGHT).pack(
                     side="left", padx=10)

        # Zone résultat
        self._altman_res = tk.Frame(inner, bg=COL_LIGHT, padx=16, pady=12)
        self._altman_res.pack(fill="x", pady=4)
        tk.Label(self._altman_res,
                 text="Renseignez les 5 variables pour calculer le score Z'",
                 font=("Arial", 11), fg="#8a9ab0", bg=COL_LIGHT).pack()

        # Avertissement
        wf = tk.Frame(inner, bg=BG_ORANGE, padx=16, pady=10)
        wf.pack(fill="x", pady=4)
        tk.Label(wf, text="Précautions d'usage", font=("Arial", 9, "bold"),
                 fg=COL_ORANGE, bg=BG_ORANGE).pack(anchor="w")
        tk.Label(wf, font=("Arial", 9), fg="#5a4000", bg=BG_ORANGE,
                 wraplength=750, justify="left",
                 text=(
                     "Le Z' est un outil d'alerte précoce, pas un verdict. "
                     "Calibré sur PME non cotées, il est à croiser avec l'analyse qualitative, "
                     "le secteur et l'évolution tendancielle. "
                     "Les spécificités belges (holding, immobilier, profession libérale) "
                     "peuvent déformer certains ratios. Ne jamais conclure sur un exercice isolé."
                 )).pack(anchor="w")

    # ── Footer ───────────────────────────────────────────────────────────────

    def _build_footer(self):
        f = tk.Frame(self.root, bg=COL_NAVY, pady=8, padx=14)
        f.pack(fill="x", side="bottom")
        bs = dict(font=("Arial", 10, "bold"), relief="flat",
                  padx=14, pady=7, cursor="hand2")
        tk.Button(f, text="📂  Importer bilan", bg="#5c3d8f", fg="white",
                  command=self._import_bilan, **bs).pack(side="left", padx=6)
        tk.Button(f, text="⟳  Calculer", bg="#2d8cff", fg="white",
                  command=self._recalc, **bs).pack(side="left", padx=6)
        tk.Button(f, text="📊  Exporter Excel", bg="#1a6b3a", fg="white",
                  command=self._export, **bs).pack(side="left", padx=6)
        tk.Button(f, text="🗑  Réinitialiser", bg="#5a3800", fg="white",
                  command=self._reset, **bs).pack(side="left", padx=6)
        tk.Label(f,
                 text="Données traitées localement — aucune connexion réseau",
                 font=("Arial", 8), fg="#3a6a8a", bg=COL_NAVY).pack(
                     side="right", padx=14)

    # ── Logique ───────────────────────────────────────────────────────────────

    def _fval(self, var):
        v = var.get().strip().replace(" ", "").replace(",", ".")
        if not v:
            return None
        try:
            return float(v)
        except ValueError:
            return None

    def _read(self, fdict):
        return {k: self._fval(v) for k, v in fdict.items()}

    def _recalc(self):
        ct = self.v_type.get()
        dn  = self._read(self.flds_n)
        dnm1= self._read(self.flds_nm1)
        self.ratios_n   = calculate_ratios(dn)
        self.ratios_nm1 = calculate_ratios(dnm1)

        S_COLORS = {
            "vert":   (BG_VERT,   COL_VERT),
            "orange": (BG_ORANGE, COL_ORANGE),
            "rouge":  (BG_ROUGE,  COL_ROUGE),
        }
        S_TXT = {"vert": "✓ OK", "orange": "⚠ Attention", "rouge": "✗ Alerte"}
        T_SYM = {"amelioration": ("↑", COL_VERT),
                  "degradation":  ("↓", COL_ROUGE),
                  "stable":       ("→", "#8a9ab0")}

        for key, (lv_n, ls_n, lv_nm1, ls_nm1, lt) in self.rlabels.items():
            unite = RATIO_META[key][1]
            vn    = self.ratios_n.get(key)
            vnm1  = self.ratios_nm1.get(key)
            sn    = eval_ratio(key, vn,   ct)
            snm1  = eval_ratio(key, vnm1, ct)
            trend = get_trend(key, vn, vnm1, ct)

            lv_n.config(text=f"{vn:.2f}{unite}" if vn is not None else "—",
                         fg="#1a2533" if vn is not None else "#ccc")
            if sn:
                bg, fg = S_COLORS[sn]
                ls_n.config(text=S_TXT[sn], bg=bg, fg=fg)
            else:
                ls_n.config(text="", bg=COL_WHITE, fg="#ccc")

            lv_nm1.config(text=f"{vnm1:.2f}{unite}" if vnm1 is not None else "—",
                           fg="#1a2533" if vnm1 is not None else "#ccc")
            if snm1:
                bg, fg = S_COLORS[snm1]
                ls_nm1.config(text=S_TXT[snm1], bg=bg, fg=fg)
            else:
                ls_nm1.config(text="", bg=COL_WHITE, fg="#ccc")

            if trend:
                sym, col = T_SYM.get(trend, ("", "#ccc"))
                lt.config(text=sym, fg=col)
            else:
                lt.config(text="", fg="#ccc")

        self._compute_altman()

    def _autofill_altman(self):
        """Remplit X1, X3, X4, X5 depuis les données saisies (N)."""
        d = self._read(self.flds_n)
        tb   = d.get("total_bilan") or 0
        stk  = d.get("stocks")  or 0
        cre  = d.get("creances") or 0
        dispo= d.get("disponibilites") or 0
        dct  = d.get("dettes_ct") or 0
        ebit = d.get("ebit") or 0
        fp   = d.get("fp") or 0
        dlt  = d.get("dettes_lt") or 0
        ca   = d.get("ca") or 0
        reserves = d.get("reserves") or 0

        def fmt(v):
            return f"{v:.4f}".replace(".", ",") if v else ""

        if tb:
            frn = (stk + cre + dispo) - dct
            self.alt_flds["x1"].set(fmt(_div(frn, tb) or 0))
            self.alt_flds["x3"].set(fmt(_div(ebit, tb) or 0))
            tdettes = dct + dlt
            self.alt_flds["x4"].set(fmt(_div(fp, tdettes) or 0))
            self.alt_flds["x5"].set(fmt(_div(ca, tb) or 0))
        self._compute_altman()

    def _compute_altman(self):
        vals = {k: v.get() for k, v in self.alt_flds.items()}
        z, zone = calc_altman(vals)

        for w in self._altman_res.winfo_children():
            w.destroy()

        if z is None:
            tk.Label(self._altman_res,
                     text="Renseignez les 5 variables pour calculer le score Z'",
                     font=("Arial", 11), fg="#8a9ab0", bg=COL_LIGHT).pack()
            return

        color_key, zone_label = zone
        zmap = {"vert":   (BG_VERT,   COL_VERT),
                "orange": (BG_ORANGE, COL_ORANGE),
                "rouge":  (BG_ROUGE,  COL_ROUGE)}
        zmsg = {
            "vert":   "Risque de défaillance faible à horizon 2 ans.",
            "orange": "Zone d'incertitude — suivi renforcé recommandé.",
            "rouge":  "Signal de détresse financière — action urgente requise.",
        }
        bg, fg = zmap[color_key]
        rf = tk.Frame(self._altman_res, bg=bg, padx=20, pady=14)
        rf.pack(fill="x")
        tk.Label(rf, text=f"Score Z'  =  {z:.3f}",
                 font=("Arial", 22, "bold"), fg=fg, bg=bg).pack(side="left", padx=14)
        rr = tk.Frame(rf, bg=bg); rr.pack(side="left", padx=16)
        tk.Label(rr, text=zone_label, font=("Arial", 14, "bold"),
                 fg="white", bg=fg, padx=12, pady=4).pack(anchor="w")
        tk.Label(rr, text=zmsg[color_key], font=("Arial", 10),
                 fg=fg, bg=bg).pack(anchor="w", pady=4)

    def _import_bilan(self):
        """Importe les données depuis un bilan WinBooks XLS ou PDF."""
        if not IMPORT_DISPONIBLE:
            messagebox.showerror(
                "Module manquant",
                "Le fichier importer_bilan.py est introuvable.\n"
                "Vérifier qu'il se trouve dans le même dossier que l'application.")
            return

        filepath = filedialog.askopenfilename(
            title="Sélectionner un bilan WinBooks",
            filetypes=[
                ("Bilan WinBooks", "*.xls *.xlsx *.pdf"),
                ("Excel", "*.xls *.xlsx"),
                ("PDF",   "*.pdf"),
                ("Tous",  "*.*"),
            ]
        )
        if not filepath:
            return

        try:
            data = _import_bilan_fn(filepath)
        except Exception as ex:
            messagebox.showerror("Erreur d'import", str(ex))
            return

        # Correspondance champ → StringVar N et N-1
        importes_n, importes_nm1 = [], []
        for champ, (vn, vnm1) in data.items():
            if champ in self.flds_n:
                if vn is not None:
                    self.flds_n[champ].set(str(round(vn, 2)).replace(".", ","))
                    importes_n.append(champ)
                if vnm1 is not None:
                    self.flds_nm1[champ].set(str(round(vnm1, 2)).replace(".", ","))
                    importes_nm1.append(champ)

        # Nom du client depuis le nom de fichier
        nom_fichier = os.path.splitext(os.path.basename(filepath))[0]
        if not self.v_nom.get():
            self.v_nom.set(nom_fichier)

        self._recalc()
        self.nb.select(1)  # Passer aux ratios

        messagebox.showinfo(
            "Import réussi",
            f"Données importées depuis :\n{os.path.basename(filepath)}\n\n"
            f"Champs remplis N   : {len(importes_n)}\n"
            f"Champs remplis N-1 : {len(importes_nm1)}\n\n"
            f"Vérifier les données avant export."
        )

    def _reset(self):
        if not messagebox.askyesno("Réinitialisation",
                                   "Effacer toutes les données saisies ?"):
            return
        for v in self.flds_n.values():   v.set("")
        for v in self.flds_nm1.values(): v.set("")
        for v in self.alt_flds.values(): v.set("")
        self.v_nom.set("")
        self._recalc()
        self.nb.select(0)

    # ── Export Excel ─────────────────────────────────────────────────────────

    def _export(self):
        if not self.ratios_n:
            self._recalc()
        nom = self.v_nom.get().strip() or "Client"
        n   = self.v_n.get().strip()   or "N"
        nm1 = self.v_nm1.get().strip() or "N-1"
        ct  = self.v_type.get()
        ts  = datetime.now().strftime("%Y%m%d_%H%M")
        default = f"Ratios_{nom.replace(' ','_')}_{n}_{ts}.xlsx"

        init_dir = EXPORT_DOSSIER if EXPORT_DOSSIER and os.path.isdir(EXPORT_DOSSIER) else None
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Classeur Excel", "*.xlsx")],
            initialfile=default,
            initialdir=init_dir,
            title="Enregistrer l'analyse")
        if not path:
            return

        try:
            wb = openpyxl.Workbook()
            self._ws_analyse(wb, nom, ct, n, nm1)
            self._ws_donnees(wb, n, nm1)
            del wb["Sheet"]
            wb.save(path)
            messagebox.showinfo("Export réussi",
                                f"Fichier enregistré :\n{path}")
        except Exception as ex:
            messagebox.showerror("Erreur", str(ex))

    @staticmethod
    def _sc(cell, bold=False, sz=9, fg="000000", bg=None,
             align="left", wrap=False, border=False):
        cell.font      = Font(name="Arial", bold=bold, size=sz, color=fg)
        cell.alignment = Alignment(horizontal=align, vertical="center",
                                   wrap_text=wrap)
        if bg:
            cell.fill = PatternFill("solid", start_color=bg)
        if border:
            t = Side(style="thin", color="DDDDDD")
            cell.border = Border(bottom=t)

    def _ws_analyse(self, wb, nom, ct, n, nm1):
        ws = wb.create_sheet("Ratios — Analyse")
        ws.freeze_panes = "A5"

        # Titre
        ws.merge_cells("A1:G1")
        ws["A1"] = f"{CABINET}  —  Analyse financière  —  {nom}"
        self._sc(ws["A1"], bold=True, sz=13, fg="FFFFFF", bg="0D1E35", align="center")
        ws.row_dimensions[1].height = 22

        ws.merge_cells("A2:G2")
        ws["A2"] = (f"Type : {ct}   |   N : {n}   |   N-1 : {nm1}   |   "
                    f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}")
        self._sc(ws["A2"], sz=9, fg="7A9ABF", bg="1A3452", align="center")
        ws.row_dimensions[2].height = 15

        # En-têtes colonnes
        hdrs = ["Catégorie", "Ratio",
                f"Valeur N ({n})", f"Statut N",
                f"Valeur N-1 ({nm1})", "Statut N-1",
                "Tendance"]
        for c, h in enumerate(hdrs, 1):
            cell = ws.cell(row=4, column=c, value=h)
            self._sc(cell, bold=True, sz=9, fg="FFFFFF", bg="2D4A6A", align="center")
        ws.row_dimensions[4].height = 16

        S_COLORS_XL = {"vert":   ("E8F7EE", "1A6B3A"),
                        "orange": ("FFF4E0", "8A5A00"),
                        "rouge":  ("FDE8E8", "8B1A1A")}
        S_TXT = {"vert": "✓ OK", "orange": "⚠ Attention", "rouge": "✗ Alerte"}
        T_SYM = {"amelioration": "↑", "degradation": "↓", "stable": "→"}
        T_COL = {"amelioration": "1A6B3A", "degradation": "8B1A1A", "stable": "8A9AB0"}
        CAT_HEX = {
            "Rentabilité":    "1A6B5E",
            "Liquidité":      "1E4D8C",
            "Solvabilité":    "7B2D00",
            "Activité & BFR": "5C3D8F",
            "Structure":      "2D6040",
        }

        row = 5
        cur_cat = None
        for key, (label, unite, cat) in RATIO_META.items():
            if cat != cur_cat:
                cur_cat = cat
                ws.merge_cells(f"A{row}:G{row}")
                c = ws.cell(row=row, column=1, value=cat)
                self._sc(c, bold=True, sz=10, fg="FFFFFF",
                          bg=CAT_HEX.get(cat, "333333"))
                ws.row_dimensions[row].height = 15
                row += 1

            vn   = self.ratios_n.get(key)
            vnm1 = self.ratios_nm1.get(key)
            sn   = eval_ratio(key, vn,   ct)
            snm1 = eval_ratio(key, vnm1, ct)
            trend= get_trend(key, vn, vnm1, ct)

            def fmt(v, u): return f"{v:.2f}{u}" if v is not None else "—"

            ws.cell(row=row, column=1, value=cat)
            self._sc(ws.cell(row=row, column=1), sz=8, fg="8A9AB0", border=True)

            ws.cell(row=row, column=2, value=label)
            self._sc(ws.cell(row=row, column=2), bold=True, sz=9, border=True)

            ws.cell(row=row, column=3, value=fmt(vn, unite))
            self._sc(ws.cell(row=row, column=3), sz=9, align="center", border=True)

            s_val = S_TXT.get(sn, "—") if sn else "—"
            ws.cell(row=row, column=4, value=s_val)
            if sn:
                bg, fg = S_COLORS_XL[sn]
                self._sc(ws.cell(row=row, column=4), bold=True, sz=9,
                          fg=fg, bg=bg, align="center", border=True)
            else:
                self._sc(ws.cell(row=row, column=4), sz=9,
                          align="center", border=True)

            ws.cell(row=row, column=5, value=fmt(vnm1, unite))
            self._sc(ws.cell(row=row, column=5), sz=9, align="center", border=True)

            s_val2 = S_TXT.get(snm1, "—") if snm1 else "—"
            ws.cell(row=row, column=6, value=s_val2)
            if snm1:
                bg, fg = S_COLORS_XL[snm1]
                self._sc(ws.cell(row=row, column=6), bold=True, sz=9,
                          fg=fg, bg=bg, align="center", border=True)
            else:
                self._sc(ws.cell(row=row, column=6), sz=9,
                          align="center", border=True)

            t_sym = T_SYM.get(trend, "") if trend else ""
            ws.cell(row=row, column=7, value=t_sym)
            if trend:
                self._sc(ws.cell(row=row, column=7), bold=True, sz=13,
                          fg=T_COL.get(trend, "000000"), align="center", border=True)
            else:
                self._sc(ws.cell(row=row, column=7), sz=9,
                          align="center", border=True)

            ws.row_dimensions[row].height = 14
            row += 1

        # Altman
        row += 1
        ws.merge_cells(f"A{row}:G{row}")
        self._sc(ws.cell(row=row, column=1,
                          value="Score Z' d'Altman — PME non cotée"),
                  bold=True, sz=10, fg="FFFFFF", bg="3D2800")
        ws.row_dimensions[row].height = 15
        row += 1

        z, zone = calc_altman({k: v.get() for k, v in self.alt_flds.items()})
        if z is not None:
            zk, zl = zone
            zbg, zfg = {"vert": ("E8F7EE","1A6B3A"), "orange": ("FFF4E0","8A5A00"),
                         "rouge": ("FDE8E8","8B1A1A")}[zk]
            ws.merge_cells(f"A{row}:C{row}")
            self._sc(ws.cell(row=row, column=1, value=f"Score Z' = {z:.3f}"),
                      bold=True, sz=13, fg=zfg, bg=zbg, align="center")
            ws.merge_cells(f"D{row}:G{row}")
            self._sc(ws.cell(row=row, column=4, value=zl),
                      bold=True, sz=11, fg="FFFFFF", bg=zfg, align="center")
            ws.row_dimensions[row].height = 20

        for col, w in [(1,16),(2,28),(3,14),(4,14),(5,14),(6,14),(7,10)]:
            ws.column_dimensions[get_column_letter(col)].width = w

    def _ws_donnees(self, wb, n, nm1):
        ws = wb.create_sheet("Données saisies")

        for c, (h, fg) in enumerate([
            ("Poste comptable  (PCMN)", "FFFFFF"),
            (f"N ({n})", "6FCF97"),
            (f"N-1 ({nm1})", "F2C94C"),
        ], 1):
            cell = ws.cell(row=1, column=c, value=h)
            self._sc(cell, bold=True, sz=9, fg=fg, bg="2D4A6A", align="center")
        ws.row_dimensions[1].height = 16

        labels = {
            "ca":            "Chiffre d'affaires (70)",
            "achats":        "Achats & services directs (60+61)",
            "ebit":          "EBIT / Résultat d'exploitation",
            "amort":         "Amortissements & dépréciations (630+660)",
            "interets":      "Charges d'intérêts (650)",
            "res_net":       "Résultat net",
            "total_bilan":   "Total bilan",
            "immo_nettes":   "Immobilisations nettes (20-28)",
            "stocks":        "Stocks (3)",
            "creances":      "Créances clients HTVA (40)",
            "disponibilites":"Disponibilités (54-58)",
            "fp":            "Fonds propres (10-15)",
            "dettes_lt":     "Dettes financières LT (17)",
            "dettes_ct":     "Dettes CT (total)",
            "dettes_fourn":  "Dettes fournisseurs (44)",
            "camv":          "CAMV (60)",
            "reserves":      "Réserves & bénéfices reportés (13+14)",
        }

        for r, (key, lbl) in enumerate(labels.items(), 2):
            ws.cell(row=r, column=1, value=lbl).font = Font(name="Arial", size=9)
            for col, fdict in [(2, self.flds_n), (3, self.flds_nm1)]:
                val = self._fval(fdict.get(key, tk.StringVar()))
                c   = ws.cell(row=r, column=col,
                               value=val if val is not None else "")
                c.font      = Font(name="Arial", size=9)
                c.alignment = Alignment(horizontal="right")
                if val is not None:
                    c.number_format = "#,##0.00"

        ws.column_dimensions["A"].width = 36
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 18


# ─── Lancement ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.state("zoomed")  # Maximisé sous Windows
    except Exception:
        root.geometry("1080x800")
    App(root)
    root.mainloop()
