"""
pdf_merger.py — Fusionneur & Diviseur PDF Desktop
Application locale 100% offline de gestion de fichiers PDF.
© Fidunot Expertise
"""

from __future__ import annotations

import logging
import os
import subprocess
import sys
import threading
from pathlib import Path
from typing import Callable, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ── Dépendance obligatoire ────────────────────────────────────────────────────
try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.errors import PdfReadError
except ImportError:
    print("ERREUR : la bibliothèque 'pypdf' est requise.")
    print("Installez-la avec : pip install pypdf")
    sys.exit(1)

# ── Dépendance optionnelle (drag & drop) ──────────────────────────────────────
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

APP_TITLE       = "Fusionneur & Diviseur PDF — Fidunot Expertise"
WIN_WIDTH       = 700
WIN_HEIGHT      = 530
WIN_MIN_W       = 600
WIN_MIN_H       = 480

C_BG            = "#F5F5F5"
C_HEADER_BG     = "#1565C0"
C_HEADER_FG     = "#FFFFFF"
C_BTN_PRI       = "#1565C0"
C_BTN_PRI_FG    = "#FFFFFF"
C_BTN_SEC       = "#ECEFF1"
C_BTN_SEC_FG    = "#263238"
C_BTN_MERGE     = "#2E7D32"
C_BTN_MERGE_FG  = "#FFFFFF"
C_BTN_MERGE_DIS = "#90A4AE"
C_BTN_SPLIT     = "#6A1B9A"
C_BTN_SPLIT_FG  = "#FFFFFF"
C_BTN_SPLIT_DIS = "#90A4AE"
C_LIST_BG       = "#FFFFFF"
C_LIST_SEL      = "#BBDEFB"
C_BORDER        = "#B0BEC5"
C_DEST_BG       = "#E3F2FD"
C_INFO_BG       = "#F3E5F5"
C_STATUS_BG     = "#ECEFF1"
C_STATUS_OK     = "#1B5E20"
C_STATUS_ERR    = "#B71C1C"
C_STATUS_INFO   = "#0D47A1"

F_MAIN          = ("Segoe UI", 10)
F_BOLD          = ("Segoe UI", 10, "bold")
F_SMALL         = ("Segoe UI", 9)
F_HEADER        = ("Segoe UI", 13, "bold")

PAD             = 8
LOG_FILE        = "pdf_merger.log"


# ══════════════════════════════════════════════════════════════════════════════
# LOGGING
# ══════════════════════════════════════════════════════════════════════════════

def _setup_logging() -> logging.Logger:
    """Configure le fichier de log dans le répertoire de l'exécutable."""
    base_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)-8s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.FileHandler(base_dir / LOG_FILE, encoding="utf-8")],
    )
    return logging.getLogger("pdf_merger")


logger = _setup_logging()


# ══════════════════════════════════════════════════════════════════════════════
# UTILITAIRE : PARSING DES PLAGES DE PAGES
# ══════════════════════════════════════════════════════════════════════════════

def parse_page_ranges(text: str, total: int) -> List[Tuple[int, int]]:
    """
    Convertit une chaîne de plages (ex: '1-3, 5, 7-10') en liste de tuples (début, fin).

    Les numéros sont 1-indexés et bornés à [1, total].
    Lève ValueError avec un message explicite en cas d'erreur de syntaxe.
    """
    result: List[Tuple[int, int]] = []
    for raw in text.split(","):
        part = raw.strip()
        if not part:
            continue
        if "-" in part:
            halves = part.split("-", 1)
            try:
                a, b = int(halves[0].strip()), int(halves[1].strip())
            except ValueError:
                raise ValueError(f"Plage invalide : « {part} »")
            if a < 1 or b < a:
                raise ValueError(f"Plage invalide : « {part} » (début ≥ 1, fin ≥ début)")
            if a > total:
                raise ValueError(f"Page {a} hors limites (le PDF a {total} pages)")
            result.append((a, min(b, total)))
        else:
            try:
                n = int(part)
            except ValueError:
                raise ValueError(f"Numéro de page invalide : « {part} »")
            if n < 1 or n > total:
                raise ValueError(f"Page {n} hors limites (le PDF a {total} pages)")
            result.append((n, n))
    return result


# ══════════════════════════════════════════════════════════════════════════════
# MERGE WORKER
# ══════════════════════════════════════════════════════════════════════════════

class MergeWorker(threading.Thread):
    """
    Thread de fusion PDF.

    Ne manipule jamais de widgets directement ; communique exclusivement
    via des callbacks planifiés dans le thread UI via root.after().
    """

    def __init__(
        self,
        files: List[str],
        output_path: str,
        on_progress: Callable[[int, int, str], None],
        on_done: Callable[[int, int, List[str]], None],
        on_error: Callable[[str], None],
    ) -> None:
        """
        Paramètres
        ----------
        files        : Chemins PDF ordonnés à fusionner.
        output_path  : Chemin absolu du fichier de sortie.
        on_progress  : (current, total, label).
        on_done      : (total_pages, file_size_bytes, warnings).
        on_error     : (message) — erreur fatale, fusion abandonnée.
        """
        super().__init__(daemon=True, name="MergeWorker")
        self.files       = files
        self.output_path = output_path
        self.on_progress = on_progress
        self.on_done     = on_done
        self.on_error    = on_error

    def run(self) -> None:
        """Fusionne les PDF un par un ; isole chaque fichier dans son try/except."""
        writer          = PdfWriter()
        warnings: List[str] = []
        total           = len(self.files)
        output_resolved = Path(self.output_path).resolve()

        for idx, path in enumerate(self.files, start=1):
            fname = Path(path).name
            self.on_progress(idx - 1, total, fname)
            try:
                resolved = Path(path).resolve()
                if resolved == output_resolved:
                    warnings.append(f"Ignoré (identique au fichier de sortie) : {fname}")
                    logger.warning("Source == destination, ignoré : %s", path)
                    continue
                if not resolved.exists():
                    raise FileNotFoundError(path)
                reader = PdfReader(str(resolved))
                if reader.is_encrypted:
                    warnings.append(f"PDF protégé par mot de passe ignoré : {fname}")
                    logger.warning("PDF chiffré ignoré : %s", path)
                    continue
                pages_added = 0
                for page in reader.pages:
                    writer.add_page(page)
                    pages_added += 1
                logger.info("Ajouté : %s  (%d pages)", path, pages_added)
            except FileNotFoundError:
                warnings.append(f"Fichier introuvable : {fname}")
                logger.error("Introuvable : %s", path)
            except PdfReadError as exc:
                warnings.append(f"PDF illisible : {fname}  ({exc})")
                logger.error("PdfReadError sur %s : %s", path, exc)
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"Erreur inattendue sur {fname} : {exc}")
                logger.exception("Erreur inattendue sur %s", path)

        self.on_progress(total, total, "Écriture du fichier de sortie…")

        if len(writer.pages) == 0:
            msg = "Aucun contenu valide à fusionner : tous les fichiers ont été ignorés."
            if warnings:
                msg += "\n\nDétails :\n• " + "\n• ".join(warnings)
            logger.error(msg)
            self.on_error(msg)
            return

        try:
            with open(self.output_path, "wb") as fp:
                writer.write(fp)
            file_size = Path(self.output_path).stat().st_size
            logger.info("Fusion → %s  (%d pages, %d o)", self.output_path, len(writer.pages), file_size)
            self.on_done(len(writer.pages), file_size, warnings)
        except PermissionError:
            msg = (
                f"Impossible d'écrire le fichier :\n{self.output_path}\n\n"
                "Vérifiez que le fichier n'est pas ouvert dans un autre programme."
            )
            logger.error("PermissionError : %s", self.output_path)
            self.on_error(msg)
        except OSError as exc:
            logger.error("OSError écriture : %s", exc)
            self.on_error(f"Erreur système lors de l'écriture :\n{exc}")


# ══════════════════════════════════════════════════════════════════════════════
# SPLIT WORKER
# ══════════════════════════════════════════════════════════════════════════════

class SplitWorker(threading.Thread):
    """
    Thread de division PDF.

    Mode 0 : une page par fichier.
    Mode 1 : plages personnalisées définies par l'utilisateur.
    """

    def __init__(
        self,
        source: str,
        output_dir: str,
        prefix: str,
        mode: int,
        ranges: List[Tuple[int, int]],
        on_progress: Callable[[int, int, str], None],
        on_done: Callable[[int, List[str]], None],
        on_error: Callable[[str], None],
    ) -> None:
        """
        Paramètres
        ----------
        source      : Chemin du PDF à diviser.
        output_dir  : Dossier de destination.
        prefix      : Préfixe des fichiers produits.
        mode        : 0 = page par page, 1 = plages.
        ranges      : Liste de (début, fin) 1-indexée (utilisée si mode==1).
        on_progress : (current, total, label).
        on_done     : (n_files, filenames).
        on_error    : (message).
        """
        super().__init__(daemon=True, name="SplitWorker")
        self.source     = source
        self.output_dir = Path(output_dir)
        self.prefix     = prefix or "page_"
        self.mode       = mode
        self.ranges     = ranges
        self.on_progress = on_progress
        self.on_done    = on_done
        self.on_error   = on_error

    def _write_pdf(self, writer: PdfWriter, filename: str) -> None:
        """Écrit un PdfWriter dans output_dir/filename."""
        with open(self.output_dir / filename, "wb") as fp:
            writer.write(fp)

    def run(self) -> None:
        """Exécute la division selon le mode sélectionné."""
        try:
            reader = PdfReader(self.source)
            if reader.is_encrypted:
                self.on_error("Le PDF source est protégé par mot de passe.")
                return
            total_pages = len(reader.pages)
        except Exception as exc:  # noqa: BLE001
            self.on_error(f"Impossible de lire le fichier source :\n{exc}")
            return

        created: List[str] = []

        if self.mode == 0:
            # ── Une page par fichier ─────────────────────────────────────────
            pad = len(str(total_pages))
            for i, page in enumerate(reader.pages, start=1):
                fname = f"{self.prefix}{str(i).zfill(pad)}.pdf"
                self.on_progress(i - 1, total_pages, fname)
                writer = PdfWriter()
                writer.add_page(page)
                try:
                    self._write_pdf(writer, fname)
                    created.append(fname)
                    logger.info("Split page %d → %s", i, fname)
                except OSError as exc:
                    self.on_error(f"Erreur écriture {fname} :\n{exc}")
                    return
            self.on_progress(total_pages, total_pages, "Terminé")

        else:
            # ── Plages personnalisées ─────────────────────────────────────────
            n = len(self.ranges)
            pad = len(str(n))
            for idx, (start, end) in enumerate(self.ranges, start=1):
                label = f"p{start}" if start == end else f"p{start}-{end}"
                fname = f"{self.prefix}{str(idx).zfill(pad)}_{label}.pdf"
                self.on_progress(idx - 1, n, fname)
                writer = PdfWriter()
                for page_num in range(start - 1, end):  # convertit en 0-index
                    writer.add_page(reader.pages[page_num])
                try:
                    self._write_pdf(writer, fname)
                    created.append(fname)
                    logger.info("Split plage %d-%d → %s", start, end, fname)
                except OSError as exc:
                    self.on_error(f"Erreur écriture {fname} :\n{exc}")
                    return
            self.on_progress(n, n, "Terminé")

        logger.info("Division réussie : %d fichier(s) dans %s", len(created), self.output_dir)
        self.on_done(len(created), created)


# ══════════════════════════════════════════════════════════════════════════════
# APPLICATION PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

class App:
    """Fenêtre principale — deux onglets : Fusionner et Diviser."""

    def __init__(self, root: tk.Tk) -> None:
        """Initialise l'état interne, construit l'UI et enregistre le DnD."""
        self.root = root

        # État fusion
        self._pdf_paths: List[str] = []
        self._output_path: Optional[str] = None
        self._merging = False

        # État division
        self._split_source: Optional[str] = None
        self._split_source_pages: int = 0
        self._split_output_dir: Optional[str] = None
        self._splitting = False

        self._setup_window()
        self._configure_styles()
        self._build_ui()
        self._setup_dnd()
        self._refresh_buttons()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Fenêtre & styles ──────────────────────────────────────────────────────

    def _setup_window(self) -> None:
        """Configure titre, dimensions, icône et position centrée."""
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        self.root.minsize(WIN_MIN_W, WIN_MIN_H)
        self.root.configure(bg=C_BG)

        base_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
        icon_path = base_dir / "icon.ico"
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - WIN_WIDTH)  // 2
        y = (self.root.winfo_screenheight() - WIN_HEIGHT) // 2
        self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+{x}+{y}")

    def _configure_styles(self) -> None:
        """Configure les styles ttk une seule fois avant la construction des widgets."""
        s = ttk.Style()
        s.theme_use("default")
        s.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor=C_BORDER, background=C_BTN_MERGE, thickness=14,
        )
        s.configure(
            "Purple.Horizontal.TProgressbar",
            troughcolor=C_BORDER, background=C_BTN_SPLIT, thickness=14,
        )
        s.configure("TNotebook",    background=C_BG, borderwidth=0)
        s.configure("TNotebook.Tab", font=F_MAIN, padding=[14, 6])
        s.map("TNotebook.Tab", background=[("selected", C_BG)])

    # ── Construction de l'interface ───────────────────────────────────────────

    def _build_ui(self) -> None:
        """Assemble les blocs : en-tête, notebook à deux onglets, barre de statut."""
        self._build_header()
        self._build_notebook()
        self._build_statusbar()

    def _build_header(self) -> None:
        """Bandeau titre coloré en haut de la fenêtre."""
        hdr = tk.Frame(self.root, bg=C_HEADER_BG, pady=10)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=APP_TITLE, bg=C_HEADER_BG, fg=C_HEADER_FG, font=F_HEADER).pack(
            side=tk.LEFT, padx=PAD * 2
        )
        if DND_AVAILABLE:
            tk.Label(
                hdr, text="✦ Glisser-déposer activé",
                bg=C_HEADER_BG, fg="#90CAF9", font=("Segoe UI", 8),
            ).pack(side=tk.RIGHT, padx=PAD * 2)

    def _build_notebook(self) -> None:
        """Crée le Notebook et ses deux onglets."""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_merge = tk.Frame(self.notebook, bg=C_BG)
        self.tab_split = tk.Frame(self.notebook, bg=C_BG)

        self.notebook.add(self.tab_merge, text="  📄  Fusionner  ")
        self.notebook.add(self.tab_split, text="  ✂  Diviser  ")

        self._build_merge_tab()
        self._build_split_tab()

    def _build_merge_tab(self) -> None:
        """Construit le contenu de l'onglet Fusionner."""
        p = self.tab_merge
        self._build_toolbar(p)
        self._build_list_area(p)
        self._build_destination_row(p)
        self._build_merge_row(p)

    def _build_toolbar(self, parent: tk.Frame) -> None:
        """Boutons de gestion de la liste de fichiers."""
        frame = tk.Frame(parent, bg=C_BG, pady=6, padx=PAD)
        frame.pack(fill=tk.X)

        self.btn_add = self._make_btn(frame, "➕  Ajouter fichiers", self._add_files,
                                      bg=C_BTN_PRI, fg=C_BTN_PRI_FG, bold=True)
        self.btn_add.pack(side=tk.LEFT, padx=(0, PAD))

        for label, cmd, attr in [
            ("⬆  Monter",    self._move_up,         "btn_up"),
            ("⬇  Descendre", self._move_down,       "btn_down"),
            ("🗑  Supprimer", self._remove_selected, "btn_remove"),
            ("✕  Vider",     self._clear_list,      "btn_clear"),
        ]:
            btn = self._make_btn(frame, label, cmd)
            btn.pack(side=tk.LEFT, padx=(0, 4))
            setattr(self, attr, btn)

    def _build_list_area(self, parent: tk.Frame) -> None:
        """Listbox scrollable pour les fichiers PDF à fusionner."""
        outer = tk.Frame(parent, bg=C_BG)
        outer.pack(fill=tk.BOTH, expand=True, padx=PAD, pady=(0, 4))

        tk.Label(
            outer,
            text="Fichiers PDF à fusionner (double-clic → chemin complet) :",
            bg=C_BG, fg="#546E7A", font=F_SMALL,
        ).pack(anchor=tk.W, pady=(0, 3))

        border = tk.Frame(outer, bg=C_BORDER, bd=1, relief=tk.FLAT)
        border.pack(fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(border, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(
            border, yscrollcommand=vsb.set,
            bg=C_LIST_BG, selectbackground=C_LIST_SEL, selectforeground="#000000",
            font=F_MAIN, relief=tk.FLAT, bd=0, activestyle="none", highlightthickness=0,
        )
        vsb.config(command=self.listbox.yview)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        self.listbox.bind("<Double-Button-1>",  self._on_double_click)
        self.listbox.bind("<Delete>",           lambda _: self._remove_selected())

    def _build_destination_row(self, parent: tk.Frame) -> None:
        """Ligne de sélection du fichier PDF de sortie (fusion)."""
        row = tk.Frame(parent, bg=C_DEST_BG, pady=6, padx=PAD)
        row.pack(fill=tk.X, padx=PAD, pady=(2, 4))

        tk.Label(row, text="Destination :", bg=C_DEST_BG, font=F_BOLD).pack(side=tk.LEFT)
        self.dest_var = tk.StringVar(value="— Non définie —")
        tk.Label(row, textvariable=self.dest_var, bg=C_DEST_BG, fg="#1A237E",
                 font=F_MAIN, anchor=tk.W).pack(side=tk.LEFT, padx=(6, 0), fill=tk.X, expand=True)
        self.btn_dest = self._make_btn(row, "📂  Choisir destination", self._choose_destination,
                                       bg=C_BTN_PRI, fg=C_BTN_PRI_FG)
        self.btn_dest.pack(side=tk.RIGHT)

    def _build_merge_row(self, parent: tk.Frame) -> None:
        """Barre de progression et bouton Fusionner."""
        row = tk.Frame(parent, bg=C_BG, pady=4, padx=PAD)
        row.pack(fill=tk.X, padx=PAD)

        self.merge_progress = ttk.Progressbar(
            row, orient=tk.HORIZONTAL, mode="determinate",
            style="Green.Horizontal.TProgressbar",
        )
        self.merge_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, PAD))

        self.btn_merge = self._make_btn(row, "⚡  Fusionner", self._start_merge,
                                        bg=C_BTN_MERGE, fg=C_BTN_MERGE_FG, bold=True)
        self.btn_merge.pack(side=tk.RIGHT)

    def _build_split_tab(self) -> None:
        """Construit le contenu de l'onglet Diviser."""
        p = self.tab_split

        # ── Source ───────────────────────────────────────────────────────────
        src_row = tk.Frame(p, bg=C_BG, pady=6, padx=PAD)
        src_row.pack(fill=tk.X)
        tk.Label(src_row, text="Fichier PDF source :", bg=C_BG, font=F_BOLD).pack(side=tk.LEFT)
        self.split_src_var = tk.StringVar(value="— Non sélectionné —")
        tk.Label(src_row, textvariable=self.split_src_var, bg=C_BG, fg="#1A237E",
                 font=F_MAIN, anchor=tk.W).pack(side=tk.LEFT, padx=(6, 0), fill=tk.X, expand=True)
        self.btn_split_src = self._make_btn(src_row, "📂  Choisir", self._choose_split_source,
                                            bg=C_BTN_PRI, fg=C_BTN_PRI_FG)
        self.btn_split_src.pack(side=tk.RIGHT)

        # ── Infos PDF ─────────────────────────────────────────────────────────
        info_row = tk.Frame(p, bg=C_INFO_BG, pady=4, padx=PAD)
        info_row.pack(fill=tk.X, padx=PAD, pady=(0, 6))
        self.split_info_var = tk.StringVar(value="Sélectionnez un fichier PDF source pour commencer.")
        tk.Label(info_row, textvariable=self.split_info_var, bg=C_INFO_BG,
                 fg="#4A148C", font=F_SMALL).pack(anchor=tk.W)

        # ── Mode de découpe ───────────────────────────────────────────────────
        mode_frame = tk.LabelFrame(p, text=" Mode de découpe ", bg=C_BG,
                                   font=F_BOLD, padx=PAD, pady=PAD)
        mode_frame.pack(fill=tk.X, padx=PAD, pady=(0, 6))

        self.split_mode = tk.IntVar(value=0)

        tk.Radiobutton(
            mode_frame,
            text="Une page par fichier  (crée N fichiers PDF d'une page chacun)",
            variable=self.split_mode, value=0,
            bg=C_BG, font=F_MAIN, command=self._on_split_mode_change,
        ).pack(anchor=tk.W, pady=(0, 6))

        range_row = tk.Frame(mode_frame, bg=C_BG)
        range_row.pack(fill=tk.X, anchor=tk.W)

        tk.Radiobutton(
            range_row, text="Par plages :",
            variable=self.split_mode, value=1,
            bg=C_BG, font=F_MAIN, command=self._on_split_mode_change,
        ).pack(side=tk.LEFT)

        self.split_ranges_var = tk.StringVar(value="")
        self.split_ranges_entry = tk.Entry(
            range_row, textvariable=self.split_ranges_var,
            font=F_MAIN, width=32, state=tk.DISABLED,
            relief=tk.FLAT, highlightthickness=1,
            highlightbackground=C_BORDER, highlightcolor=C_BTN_PRI,
        )
        self.split_ranges_entry.pack(side=tk.LEFT, padx=(8, 0), ipady=3)

        tk.Label(
            mode_frame,
            text="  Format : « 1-3, 4-6, 7 »  — pages numérotées à partir de 1",
            bg=C_BG, fg="#78909C", font=("Segoe UI", 8),
        ).pack(anchor=tk.W, pady=(4, 0))

        # ── Préfixe + dossier de sortie ───────────────────────────────────────
        out_row = tk.Frame(p, bg=C_BG, pady=4, padx=PAD)
        out_row.pack(fill=tk.X)

        tk.Label(out_row, text="Préfixe :", bg=C_BG, font=F_BOLD).pack(side=tk.LEFT)
        self.split_prefix_var = tk.StringVar(value="page_")
        tk.Entry(
            out_row, textvariable=self.split_prefix_var,
            font=F_MAIN, width=14, relief=tk.FLAT,
            highlightthickness=1, highlightbackground=C_BORDER,
        ).pack(side=tk.LEFT, padx=(4, PAD * 3), ipady=3)

        tk.Label(out_row, text="Dossier de sortie :", bg=C_BG, font=F_BOLD).pack(side=tk.LEFT)
        self.split_dir_var = tk.StringVar(value="— Non défini —")
        tk.Label(out_row, textvariable=self.split_dir_var, bg=C_BG, fg="#1A237E",
                 font=F_MAIN, anchor=tk.W).pack(side=tk.LEFT, padx=(6, 0), fill=tk.X, expand=True)
        self.btn_split_dir = self._make_btn(out_row, "📁  Choisir dossier", self._choose_split_dir,
                                            bg=C_BTN_PRI, fg=C_BTN_PRI_FG)
        self.btn_split_dir.pack(side=tk.RIGHT)

        # ── Progression + bouton ──────────────────────────────────────────────
        action_row = tk.Frame(p, bg=C_BG, pady=8, padx=PAD)
        action_row.pack(fill=tk.X, padx=PAD)

        self.split_progress = ttk.Progressbar(
            action_row, orient=tk.HORIZONTAL, mode="determinate",
            style="Purple.Horizontal.TProgressbar",
        )
        self.split_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, PAD))

        self.btn_split = self._make_btn(action_row, "✂  Diviser", self._start_split,
                                        bg=C_BTN_SPLIT, fg=C_BTN_SPLIT_FG, bold=True)
        self.btn_split.config(state=tk.DISABLED, bg=C_BTN_SPLIT_DIS)
        self.btn_split.pack(side=tk.RIGHT)

    def _build_statusbar(self) -> None:
        """Barre de statut partagée fixée en bas de la fenêtre."""
        self.status_var = tk.StringVar(value="Prêt — Ajoutez des fichiers ou choisissez un PDF à diviser.")
        self._statusbar = tk.Label(
            self.root, textvariable=self.status_var,
            bg=C_STATUS_BG, fg=C_STATUS_INFO, font=F_SMALL,
            anchor=tk.W, padx=PAD, pady=4,
        )
        self._statusbar.pack(fill=tk.X, side=tk.BOTTOM)

    # ── Helpers UI ────────────────────────────────────────────────────────────

    @staticmethod
    def _make_btn(parent, text: str, command,
                  bg: str = C_BTN_SEC, fg: str = C_BTN_SEC_FG,
                  bold: bool = False) -> tk.Button:
        """Crée un bouton avec le style standard de l'application."""
        return tk.Button(
            parent, text=text, command=command,
            bg=bg, fg=fg,
            font=F_BOLD if bold else F_MAIN,
            relief=tk.FLAT, padx=10, pady=5, cursor="hand2",
            activebackground=bg, activeforeground=fg,
        )

    def _set_status(self, msg: str, color: str = C_STATUS_INFO) -> None:
        """Met à jour le texte et la couleur de la barre de statut."""
        self.status_var.set(msg)
        self._statusbar.config(fg=color)

    def _refresh_buttons(self) -> None:
        """Recalcule l'état activé/désactivé des boutons de l'onglet Fusionner."""
        has_files = bool(self._pdf_paths)
        has_dest  = self._output_path is not None
        has_sel   = bool(self.listbox.curselection())

        self.btn_remove.config(state=tk.NORMAL if has_sel   else tk.DISABLED)
        self.btn_up    .config(state=tk.NORMAL if has_sel   else tk.DISABLED)
        self.btn_down  .config(state=tk.NORMAL if has_sel   else tk.DISABLED)
        self.btn_clear .config(state=tk.NORMAL if has_files else tk.DISABLED)

        can_merge = has_files and has_dest
        self.btn_merge.config(
            state=tk.NORMAL if can_merge else tk.DISABLED,
            bg=C_BTN_MERGE if can_merge else C_BTN_MERGE_DIS,
        )

    def _refresh_split_button(self) -> None:
        """Active le bouton Diviser uniquement si source et dossier sont définis."""
        can_split = bool(self._split_source and self._split_output_dir)
        self.btn_split.config(
            state=tk.NORMAL if can_split else tk.DISABLED,
            bg=C_BTN_SPLIT if can_split else C_BTN_SPLIT_DIS,
        )

    def _lock_ui(self, lock: bool) -> None:
        """Verrouille l'onglet Fusionner et désactive l'onglet Diviser pendant la fusion."""
        self._merging = lock
        state = tk.DISABLED if lock else tk.NORMAL
        for btn in (self.btn_add, self.btn_remove, self.btn_up,
                    self.btn_down, self.btn_clear, self.btn_dest):
            btn.config(state=state)
        if lock:
            self.btn_merge.config(state=tk.DISABLED, bg=C_BTN_MERGE_DIS)
        else:
            self._refresh_buttons()
        try:
            self.notebook.tab(1, state="disabled" if lock else "normal")
        except Exception:
            pass

    def _lock_split_ui(self, lock: bool) -> None:
        """Verrouille l'onglet Diviser et désactive l'onglet Fusionner pendant la division."""
        self._splitting = lock
        state = tk.DISABLED if lock else tk.NORMAL
        for btn in (self.btn_split_src, self.btn_split_dir):
            btn.config(state=state)
        # ranges_entry : activée seulement si non verrouillé ET mode plages
        if not lock and self.split_mode.get() == 1:
            self.split_ranges_entry.config(state=tk.NORMAL)
        else:
            self.split_ranges_entry.config(state=tk.DISABLED)
        if lock:
            self.btn_split.config(state=tk.DISABLED, bg=C_BTN_SPLIT_DIS)
        else:
            self._refresh_split_button()
        try:
            self.notebook.tab(0, state="disabled" if lock else "normal")
        except Exception:
            pass

    @staticmethod
    def _fmt_size(n: int) -> str:
        """Convertit des octets en chaîne lisible (Ko / Mo)."""
        return f"{n / 1_048_576:.1f} Mo" if n >= 1_048_576 else f"{n / 1024:.1f} Ko"

    @staticmethod
    def _short_path(path: str, max_len: int = 58) -> str:
        """Tronque un chemin trop long pour l'affichage dans les labels."""
        if len(path) <= max_len:
            return path
        p = Path(path)
        candidate = str(Path(*p.parts[:1]) / "…" / p.parent.name / p.name)
        return candidate if len(candidate) <= max_len else "…" + path[-(max_len - 3):]

    @staticmethod
    def _open_folder(path: str) -> None:
        """Ouvre le dossier dans l'explorateur (Windows / macOS / Linux)."""
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            logger.warning("Impossible d'ouvrir le dossier : %s", exc)

    # ── Actions liste (onglet Fusionner) ──────────────────────────────────────

    def _add_files(self) -> None:
        """Ouvre le sélecteur et ajoute les PDF choisis à la liste."""
        paths = filedialog.askopenfilenames(
            title="Sélectionner des fichiers PDF",
            filetypes=[("Fichiers PDF", "*.pdf"), ("Tous les fichiers", "*.*")],
        )
        added = 0
        for p in paths:
            if p not in self._pdf_paths:
                self._pdf_paths.append(p)
                self.listbox.insert(tk.END, Path(p).name)
                added += 1
        if added:
            self._set_status(f"{added} fichier(s) ajouté(s) — {len(self._pdf_paths)} au total.")
        self._refresh_buttons()

    def _remove_selected(self) -> None:
        """Supprime le ou les éléments sélectionnés de la liste."""
        sel = list(self.listbox.curselection())
        if not sel:
            return
        for i in reversed(sel):
            self.listbox.delete(i)
            del self._pdf_paths[i]
        self._set_status(f"{len(sel)} fichier(s) retiré(s) — {len(self._pdf_paths)} restant(s).")
        self._refresh_buttons()

    def _move_up(self) -> None:
        """Monte l'élément sélectionné d'une position."""
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0:
            return
        i = sel[0]
        self._swap_items(i, i - 1)
        self._reselect(i - 1)

    def _move_down(self) -> None:
        """Descend l'élément sélectionné d'une position."""
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self._pdf_paths) - 1:
            return
        i = sel[0]
        self._swap_items(i, i + 1)
        self._reselect(i + 1)

    def _swap_items(self, i: int, j: int) -> None:
        """Échange deux entrées dans le tableau interne et dans la Listbox."""
        self._pdf_paths[i], self._pdf_paths[j] = self._pdf_paths[j], self._pdf_paths[i]
        a, b = self.listbox.get(i), self.listbox.get(j)
        self.listbox.delete(i); self.listbox.insert(i, b)
        self.listbox.delete(j); self.listbox.insert(j, a)

    def _reselect(self, idx: int) -> None:
        """Repositionne la sélection sur l'index donné."""
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(idx)
        self.listbox.see(idx)

    def _clear_list(self) -> None:
        """Vide entièrement la liste après confirmation."""
        if not self._pdf_paths:
            return
        if messagebox.askyesno("Vider la liste",
                                f"Supprimer les {len(self._pdf_paths)} fichier(s) de la liste ?"):
            self.listbox.delete(0, tk.END)
            self._pdf_paths.clear()
            self._set_status("Liste vidée.")
            self._refresh_buttons()

    def _choose_destination(self) -> None:
        """Boîte de dialogue pour choisir le fichier PDF fusionné de sortie."""
        path = filedialog.asksaveasfilename(
            title="Enregistrer le PDF fusionné sous…",
            defaultextension=".pdf",
            filetypes=[("Fichier PDF", "*.pdf")],
        )
        if path:
            self._output_path = path
            self.dest_var.set(self._short_path(path))
            self._set_status(f"Destination : {Path(path).name}")
            self._refresh_buttons()

    # ── Actions onglet Diviser ────────────────────────────────────────────────

    def _choose_split_source(self) -> None:
        """Sélectionne et valide le PDF source à diviser."""
        path = filedialog.askopenfilename(
            title="Sélectionner le PDF à diviser",
            filetypes=[("Fichiers PDF", "*.pdf"), ("Tous les fichiers", "*.*")],
        )
        if not path:
            return
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                messagebox.showwarning(
                    "PDF protégé",
                    "Ce PDF est protégé par mot de passe et ne peut pas être divisé."
                )
                return
            n_pages = len(reader.pages)
        except Exception as exc:
            messagebox.showerror("Erreur", f"Impossible de lire le PDF :\n{exc}")
            return

        self._split_source       = path
        self._split_source_pages = n_pages
        self.split_src_var.set(self._short_path(path))
        self.split_info_var.set(
            f"Fichier : {Path(path).name}   |   Pages : {n_pages}"
        )
        self.split_ranges_var.set(f"1-{n_pages}")

        # Suggère automatiquement le même dossier que le source
        if not self._split_output_dir:
            suggested = str(Path(path).parent)
            self._split_output_dir = suggested
            self.split_dir_var.set(self._short_path(suggested))

        self._set_status(f"Source : {Path(path).name}  ({n_pages} pages)")
        self._refresh_split_button()

    def _choose_split_dir(self) -> None:
        """Sélectionne le dossier de sortie pour les fichiers divisés."""
        path = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if path:
            self._split_output_dir = path
            self.split_dir_var.set(self._short_path(path))
            self._set_status(f"Dossier de sortie : {path}")
            self._refresh_split_button()

    def _on_split_mode_change(self) -> None:
        """Active ou désactive le champ de plages selon le mode sélectionné."""
        if self.split_mode.get() == 1 and not self._splitting:
            self.split_ranges_entry.config(state=tk.NORMAL)
        else:
            self.split_ranges_entry.config(state=tk.DISABLED)

    # ── Drag & Drop ───────────────────────────────────────────────────────────

    def _setup_dnd(self) -> None:
        """Enregistre la Listbox comme cible de dépôt (si tkinterdnd2 disponible)."""
        if not DND_AVAILABLE:
            return
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop(self, event) -> None:
        """Traite les fichiers PDF glissés-déposés sur la Listbox."""
        try:
            paths = self.root.tk.splitlist(event.data)
        except Exception:
            return
        added = 0
        for raw in paths:
            p = raw.strip("{}")
            if p.lower().endswith(".pdf") and p not in self._pdf_paths:
                self._pdf_paths.append(p)
                self.listbox.insert(tk.END, Path(p).name)
                added += 1
        if added:
            self._set_status(f"{added} fichier(s) déposé(s) — {len(self._pdf_paths)} au total.")
            self._refresh_buttons()

    # ── Événements Listbox ────────────────────────────────────────────────────

    def _on_select(self, _event=None) -> None:
        """Affiche le chemin complet du fichier sélectionné dans la statusbar."""
        sel = self.listbox.curselection()
        if sel:
            self._set_status(self._pdf_paths[sel[0]])
        self._refresh_buttons()

    def _on_double_click(self, _event=None) -> None:
        """Affiche le chemin complet dans une boîte de dialogue."""
        sel = self.listbox.curselection()
        if sel:
            messagebox.showinfo("Chemin complet", self._pdf_paths[sel[0]])

    # ── Fusion ────────────────────────────────────────────────────────────────

    def _start_merge(self) -> None:
        """Valide les prérequis et lance MergeWorker dans un thread séparé."""
        if not self._pdf_paths:
            messagebox.showwarning("Liste vide", "Ajoutez au moins un fichier PDF.")
            return
        if not self._output_path:
            messagebox.showwarning("Destination manquante", "Choisissez le fichier de sortie.")
            return
        if Path(self._output_path).exists():
            if not messagebox.askyesno(
                "Fichier existant",
                f"Le fichier « {Path(self._output_path).name} » existe déjà.\nL'écraser ?"
            ):
                return

        self.merge_progress.config(maximum=len(self._pdf_paths), value=0)
        self._lock_ui(True)
        self._set_status("Fusion en cours…")

        MergeWorker(
            files=list(self._pdf_paths),
            output_path=self._output_path,
            on_progress=self._cb_merge_progress,
            on_done=self._cb_merge_done,
            on_error=self._cb_merge_error,
        ).start()

    def _cb_merge_progress(self, current: int, total: int, label: str) -> None:
        """Planifie la mise à jour de progression de fusion dans le thread UI."""
        self.root.after(0, self._ui_merge_progress, current, total, label)

    def _ui_merge_progress(self, current: int, total: int, label: str) -> None:
        """Met à jour progressbar et statusbar (thread UI)."""
        self.merge_progress["value"] = current
        self._set_status(f"[{current}/{total}] {label}")

    def _cb_merge_done(self, total_pages: int, file_size: int, warnings: List[str]) -> None:
        """Planifie l'affichage du succès de fusion dans le thread UI."""
        self.root.after(0, self._ui_merge_done, total_pages, file_size, warnings)

    def _ui_merge_done(self, total_pages: int, file_size: int, warnings: List[str]) -> None:
        """Affiche le résultat de fusion et déverrouille l'UI."""
        self.merge_progress["value"] = self.merge_progress["maximum"]
        size_str = self._fmt_size(file_size)
        warn_txt = ""
        if warnings:
            warn_txt = f"\n\n⚠ Avertissements ({len(warnings)}) :\n• " + "\n• ".join(warnings)
        msg = (
            f"Fusion terminée avec succès !\n\n"
            f"• Fichiers sources  : {len(self._pdf_paths)}\n"
            f"• Pages totales     : {total_pages}\n"
            f"• Taille du fichier : {size_str}\n\n"
            f"Enregistré dans :\n{self._output_path}"
            f"{warn_txt}"
        )
        self._set_status(f"Fusion réussie — {total_pages} pages, {size_str}  ✓", C_STATUS_OK)
        messagebox.showinfo("Fusion réussie", msg)
        self._lock_ui(False)

    def _cb_merge_error(self, message: str) -> None:
        """Planifie l'affichage d'une erreur fatale de fusion dans le thread UI."""
        self.root.after(0, self._ui_merge_error, message)

    def _ui_merge_error(self, message: str) -> None:
        """Affiche l'erreur fatale et déverrouille l'UI de fusion."""
        self.merge_progress["value"] = 0
        self._set_status("Erreur — fusion annulée.", C_STATUS_ERR)
        messagebox.showerror("Erreur de fusion", message)
        self._lock_ui(False)

    # ── Division ──────────────────────────────────────────────────────────────

    def _start_split(self) -> None:
        """Valide les prérequis et lance SplitWorker dans un thread séparé."""
        if not self._split_source:
            messagebox.showwarning("Source manquante", "Choisissez un fichier PDF source.")
            return
        if not self._split_output_dir:
            messagebox.showwarning("Dossier manquant", "Choisissez un dossier de sortie.")
            return

        mode        = self.split_mode.get()
        prefix      = self.split_prefix_var.get().strip() or "page_"
        parsed_ranges: List[Tuple[int, int]] = []

        if mode == 1:
            ranges_text = self.split_ranges_var.get().strip()
            if not ranges_text:
                messagebox.showwarning("Plages vides", "Saisissez au moins une plage de pages.")
                return
            try:
                parsed_ranges = parse_page_ranges(ranges_text, self._split_source_pages)
            except ValueError as exc:
                messagebox.showerror("Plages invalides", str(exc))
                return
            if not parsed_ranges:
                messagebox.showwarning("Plages vides", "Aucune plage valide trouvée.")
                return
            n_tasks = len(parsed_ranges)
        else:
            n_tasks = self._split_source_pages

        self.split_progress.config(maximum=n_tasks, value=0)
        self._lock_split_ui(True)
        self._set_status("Division en cours…")

        SplitWorker(
            source=self._split_source,
            output_dir=self._split_output_dir,
            prefix=prefix,
            mode=mode,
            ranges=parsed_ranges,
            on_progress=self._cb_split_progress,
            on_done=self._cb_split_done,
            on_error=self._cb_split_error,
        ).start()

    def _cb_split_progress(self, current: int, total: int, label: str) -> None:
        """Planifie la mise à jour de progression de division dans le thread UI."""
        self.root.after(0, self._ui_split_progress, current, total, label)

    def _ui_split_progress(self, current: int, total: int, label: str) -> None:
        """Met à jour la progressbar et la statusbar de division (thread UI)."""
        self.split_progress["value"] = current
        self._set_status(f"[{current}/{total}] Création de {label}")

    def _cb_split_done(self, n_files: int, filenames: List[str]) -> None:
        """Planifie l'affichage du succès de division dans le thread UI."""
        self.root.after(0, self._ui_split_done, n_files, filenames)

    def _ui_split_done(self, n_files: int, filenames: List[str]) -> None:
        """Affiche le résultat de division et propose d'ouvrir le dossier."""
        self.split_progress["value"] = self.split_progress["maximum"]
        preview = "\n".join(f"  • {f}" for f in filenames[:10])
        if len(filenames) > 10:
            preview += f"\n  … et {len(filenames) - 10} autre(s)"
        msg = (
            f"Division réussie !\n\n"
            f"• Fichiers créés : {n_files}\n"
            f"• Dossier        : {self._split_output_dir}\n\n"
            f"{preview}\n\nOuvrir le dossier de sortie ?"
        )
        self._set_status(f"Division réussie — {n_files} fichier(s) créé(s)  ✓", C_STATUS_OK)
        if messagebox.askyesno("Division réussie", msg):
            self._open_folder(self._split_output_dir)
        self._lock_split_ui(False)

    def _cb_split_error(self, message: str) -> None:
        """Planifie l'affichage d'une erreur fatale de division dans le thread UI."""
        self.root.after(0, self._ui_split_error, message)

    def _ui_split_error(self, message: str) -> None:
        """Affiche l'erreur fatale et déverrouille l'UI de division."""
        self.split_progress["value"] = 0
        self._set_status("Erreur — division annulée.", C_STATUS_ERR)
        messagebox.showerror("Erreur de division", message)
        self._lock_split_ui(False)

    # ── Fermeture ─────────────────────────────────────────────────────────────

    def _on_close(self) -> None:
        """Demande confirmation si une opération est en cours avant de quitter."""
        if self._merging or self._splitting:
            op = "fusion" if self._merging else "division"
            if not messagebox.askyesno(
                "Opération en cours",
                f"Une {op} est actuellement en cours.\nQuitter quand même ?"
            ):
                return
        logger.info("Application fermée par l'utilisateur.")
        self.root.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    """Lance l'application (avec support DnD si tkinterdnd2 est installé)."""
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
