"""
pdf_merger.py — Fusionneur PDF Desktop
Application locale 100% offline de fusion de fichiers PDF.
© Fidunot Expertise
"""

from __future__ import annotations

import logging
import sys
import threading
from pathlib import Path
from typing import Callable, List, Optional

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

APP_TITLE          = "Fusionneur PDF — Fidunot Expertise"
WIN_WIDTH          = 700
WIN_HEIGHT         = 500
WIN_MIN_W          = 600
WIN_MIN_H          = 440

C_BG               = "#F5F5F5"
C_HEADER_BG        = "#1565C0"
C_HEADER_FG        = "#FFFFFF"
C_BTN_PRI          = "#1565C0"
C_BTN_PRI_FG       = "#FFFFFF"
C_BTN_SEC          = "#ECEFF1"
C_BTN_SEC_FG       = "#263238"
C_BTN_MERGE        = "#2E7D32"
C_BTN_MERGE_FG     = "#FFFFFF"
C_BTN_MERGE_DIS    = "#90A4AE"
C_LIST_BG          = "#FFFFFF"
C_LIST_SEL         = "#BBDEFB"
C_BORDER           = "#B0BEC5"
C_DEST_BG          = "#E3F2FD"
C_STATUS_BG        = "#ECEFF1"
C_STATUS_OK        = "#1B5E20"
C_STATUS_ERR       = "#B71C1C"
C_STATUS_INFO      = "#0D47A1"

F_MAIN             = ("Segoe UI", 10)
F_BOLD             = ("Segoe UI", 10, "bold")
F_SMALL            = ("Segoe UI", 9)
F_HEADER           = ("Segoe UI", 13, "bold")

PAD                = 8
LOG_FILE           = "pdf_merger.log"


# ══════════════════════════════════════════════════════════════════════════════
# LOGGING
# ══════════════════════════════════════════════════════════════════════════════

def _setup_logging() -> logging.Logger:
    """Configure le fichier de log dans le répertoire de l'exécutable."""
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
    log_path = base_dir / LOG_FILE
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)-8s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.FileHandler(log_path, encoding="utf-8")],
    )
    return logging.getLogger("pdf_merger")


logger = _setup_logging()


# ══════════════════════════════════════════════════════════════════════════════
# MERGE WORKER
# ══════════════════════════════════════════════════════════════════════════════

class MergeWorker(threading.Thread):
    """
    Thread de fusion PDF.

    Communique avec l'UI exclusivement via des callbacks ; ne manipule
    jamais de widgets directement pour garantir la thread-safety.
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
        on_progress  : (current, total, label) — avancement fichier par fichier.
        on_done      : (total_pages, file_size_bytes, warnings) — succès final.
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

                # Empêcher d'écraser une source par la destination
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
            total_pages = len(writer.pages)
            file_size   = Path(self.output_path).stat().st_size
            logger.info(
                "Fusion réussie → %s  (%d pages, %d o)",
                self.output_path, total_pages, file_size,
            )
            self.on_done(total_pages, file_size, warnings)

        except PermissionError:
            msg = (
                f"Impossible d'écrire le fichier :\n{self.output_path}\n\n"
                "Vérifiez que le fichier n'est pas ouvert dans un autre programme."
            )
            logger.error("PermissionError écriture : %s", self.output_path)
            self.on_error(msg)

        except OSError as exc:
            logger.error("OSError écriture : %s", exc)
            self.on_error(f"Erreur système lors de l'écriture du fichier :\n{exc}")


# ══════════════════════════════════════════════════════════════════════════════
# APPLICATION PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════

class App:
    """Fenêtre principale du Fusionneur PDF Fidunot."""

    def __init__(self, root: tk.Tk) -> None:
        """Initialise l'état interne, construit l'UI et enregistre le DnD."""
        self.root          = root
        self._pdf_paths: List[str]     = []
        self._output_path: Optional[str] = None
        self._merging      = False

        self._setup_window()
        self._build_ui()
        self._setup_dnd()
        self._refresh_buttons()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Fenêtre ───────────────────────────────────────────────────────────────

    def _setup_window(self) -> None:
        """Configure titre, dimensions, icône et position centrée."""
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        self.root.minsize(WIN_MIN_W, WIN_MIN_H)
        self.root.configure(bg=C_BG)

        # Icône .ico optionnelle
        if getattr(sys, "frozen", False):
            icon_dir = Path(sys.executable).parent
        else:
            icon_dir = Path(__file__).parent
        icon_path = icon_dir / "icon.ico"
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

        # Centrage
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - WIN_WIDTH)  // 2
        y = (self.root.winfo_screenheight() - WIN_HEIGHT) // 2
        self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+{x}+{y}")

    # ── Construction de l'interface ───────────────────────────────────────────

    def _build_ui(self) -> None:
        """Assemble les blocs de l'interface de haut en bas."""
        self._build_header()
        self._build_toolbar()
        self._build_list_area()
        self._build_destination_row()
        self._build_merge_row()
        self._build_statusbar()

    def _build_header(self) -> None:
        """Bandeau titre coloré."""
        hdr = tk.Frame(self.root, bg=C_HEADER_BG, pady=10)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr, text=APP_TITLE,
            bg=C_HEADER_BG, fg=C_HEADER_FG, font=F_HEADER,
        ).pack(side=tk.LEFT, padx=PAD * 2)
        if DND_AVAILABLE:
            tk.Label(
                hdr, text="✦ Glisser-déposer activé",
                bg=C_HEADER_BG, fg="#90CAF9", font=("Segoe UI", 8),
            ).pack(side=tk.RIGHT, padx=PAD * 2)

    def _build_toolbar(self) -> None:
        """Boutons de gestion de la liste de fichiers."""
        frame = tk.Frame(self.root, bg=C_BG, pady=6, padx=PAD)
        frame.pack(fill=tk.X)

        self.btn_add = self._make_btn(
            frame, "➕  Ajouter fichiers", self._add_files,
            bg=C_BTN_PRI, fg=C_BTN_PRI_FG, bold=True,
        )
        self.btn_add.pack(side=tk.LEFT, padx=(0, PAD))

        btn_specs = [
            ("⬆  Monter",    self._move_up,         "btn_up"),
            ("⬇  Descendre", self._move_down,       "btn_down"),
            ("🗑  Supprimer", self._remove_selected, "btn_remove"),
            ("✕  Vider",     self._clear_list,      "btn_clear"),
        ]
        for label, cmd, attr in btn_specs:
            btn = self._make_btn(frame, label, cmd)
            btn.pack(side=tk.LEFT, padx=(0, 4))
            setattr(self, attr, btn)

    def _build_list_area(self) -> None:
        """Listbox scrollable pour les fichiers PDF."""
        outer = tk.Frame(self.root, bg=C_BG)
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
            border,
            yscrollcommand=vsb.set,
            bg=C_LIST_BG,
            selectbackground=C_LIST_SEL,
            selectforeground="#000000",
            font=F_MAIN,
            relief=tk.FLAT,
            bd=0,
            activestyle="none",
            highlightthickness=0,
        )
        vsb.config(command=self.listbox.yview)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        self.listbox.bind("<Double-Button-1>",  self._on_double_click)
        self.listbox.bind("<Delete>",           lambda _: self._remove_selected())

    def _build_destination_row(self) -> None:
        """Ligne de sélection du fichier PDF de sortie."""
        row = tk.Frame(self.root, bg=C_DEST_BG, pady=6, padx=PAD)
        row.pack(fill=tk.X, padx=PAD, pady=(2, 4))

        tk.Label(row, text="Destination :", bg=C_DEST_BG, font=F_BOLD).pack(side=tk.LEFT)

        self.dest_var = tk.StringVar(value="— Non définie —")
        tk.Label(
            row,
            textvariable=self.dest_var,
            bg=C_DEST_BG, fg="#1A237E", font=F_MAIN, anchor=tk.W,
        ).pack(side=tk.LEFT, padx=(6, 0), fill=tk.X, expand=True)

        self.btn_dest = self._make_btn(
            row, "📂  Choisir destination", self._choose_destination,
            bg=C_BTN_PRI, fg=C_BTN_PRI_FG,
        )
        self.btn_dest.pack(side=tk.RIGHT)

    def _build_merge_row(self) -> None:
        """Barre de progression et bouton principal Fusionner."""
        row = tk.Frame(self.root, bg=C_BG, pady=4, padx=PAD)
        row.pack(fill=tk.X, padx=PAD)

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Merge.Horizontal.TProgressbar",
            troughcolor=C_BORDER,
            background=C_BTN_MERGE,
            thickness=14,
        )
        self.progress = ttk.Progressbar(
            row,
            orient=tk.HORIZONTAL,
            mode="determinate",
            style="Merge.Horizontal.TProgressbar",
        )
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, PAD))

        self.btn_merge = self._make_btn(
            row, "⚡  Fusionner", self._start_merge,
            bg=C_BTN_MERGE, fg=C_BTN_MERGE_FG, bold=True,
        )
        self.btn_merge.pack(side=tk.RIGHT)

    def _build_statusbar(self) -> None:
        """Barre de statut fixe en bas de la fenêtre."""
        self.status_var = tk.StringVar(
            value="Prêt — Ajoutez des fichiers PDF pour commencer."
        )
        self._statusbar = tk.Label(
            self.root,
            textvariable=self.status_var,
            bg=C_STATUS_BG, fg=C_STATUS_INFO,
            font=F_SMALL, anchor=tk.W,
            padx=PAD, pady=4,
        )
        self._statusbar.pack(fill=tk.X, side=tk.BOTTOM)

    # ── Helpers UI ────────────────────────────────────────────────────────────

    @staticmethod
    def _make_btn(
        parent,
        text: str,
        command,
        bg: str = C_BTN_SEC,
        fg: str = C_BTN_SEC_FG,
        bold: bool = False,
    ) -> tk.Button:
        """Crée un bouton selon le style standard de l'application."""
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg, fg=fg,
            font=F_BOLD if bold else F_MAIN,
            relief=tk.FLAT,
            padx=10, pady=5,
            cursor="hand2",
            activebackground=bg,
            activeforeground=fg,
        )

    def _set_status(self, msg: str, color: str = C_STATUS_INFO) -> None:
        """Met à jour le texte et la couleur de la barre de statut."""
        self.status_var.set(msg)
        self._statusbar.config(fg=color)

    def _refresh_buttons(self) -> None:
        """Recalcule l'état activé/désactivé de chaque bouton selon l'état courant."""
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

    def _lock_ui(self, lock: bool) -> None:
        """Verrouille / déverrouille tous les boutons pendant la fusion."""
        self._merging = lock
        state = tk.DISABLED if lock else tk.NORMAL
        for btn in (self.btn_add, self.btn_remove, self.btn_up,
                    self.btn_down, self.btn_clear, self.btn_dest):
            btn.config(state=state)
        if lock:
            self.btn_merge.config(state=tk.DISABLED, bg=C_BTN_MERGE_DIS)
        else:
            self._refresh_buttons()

    @staticmethod
    def _fmt_size(n: int) -> str:
        """Convertit des octets en chaîne lisible (Ko / Mo)."""
        if n >= 1_048_576:
            return f"{n / 1_048_576:.1f} Mo"
        return f"{n / 1024:.1f} Ko"

    @staticmethod
    def _short_path(path: str, max_len: int = 60) -> str:
        """Tronque un chemin trop long pour l'affichage dans la destination."""
        if len(path) <= max_len:
            return path
        p = Path(path)
        candidate = str(Path(*p.parts[:1]) / "…" / p.parent.name / p.name)
        return candidate if len(candidate) <= max_len else "…" + path[-(max_len - 3):]

    # ── Actions liste ─────────────────────────────────────────────────────────

    def _add_files(self) -> None:
        """Ouvre le sélecteur de fichiers et ajoute les PDF choisis."""
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
            self._set_status(
                f"{added} fichier(s) ajouté(s) — {len(self._pdf_paths)} au total."
            )
        self._refresh_buttons()

    def _remove_selected(self) -> None:
        """Supprime le ou les éléments actuellement sélectionnés dans la liste."""
        sel = list(self.listbox.curselection())
        if not sel:
            return
        for i in reversed(sel):
            self.listbox.delete(i)
            del self._pdf_paths[i]
        self._set_status(
            f"{len(sel)} fichier(s) retiré(s) — {len(self._pdf_paths)} restant(s)."
        )
        self._refresh_buttons()

    def _move_up(self) -> None:
        """Monte l'élément sélectionné d'une position dans la liste."""
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0:
            return
        i = sel[0]
        self._swap_items(i, i - 1)
        self._reselect(i - 1)

    def _move_down(self) -> None:
        """Descend l'élément sélectionné d'une position dans la liste."""
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
        self.listbox.delete(i)
        self.listbox.insert(i, b)
        self.listbox.delete(j)
        self.listbox.insert(j, a)

    def _reselect(self, idx: int) -> None:
        """Repositionne la sélection sur l'index donné et le rend visible."""
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(idx)
        self.listbox.see(idx)

    def _clear_list(self) -> None:
        """Vide entièrement la liste après confirmation de l'utilisateur."""
        if not self._pdf_paths:
            return
        if messagebox.askyesno(
            "Vider la liste",
            f"Supprimer les {len(self._pdf_paths)} fichier(s) de la liste ?",
        ):
            self.listbox.delete(0, tk.END)
            self._pdf_paths.clear()
            self._set_status("Liste vidée.")
            self._refresh_buttons()

    # ── Destination ───────────────────────────────────────────────────────────

    def _choose_destination(self) -> None:
        """Boîte de dialogue pour choisir l'emplacement du PDF fusionné."""
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
            self._set_status(
                f"{added} fichier(s) déposé(s) — {len(self._pdf_paths)} au total."
            )
            self._refresh_buttons()

    # ── Événements Listbox ────────────────────────────────────────────────────

    def _on_select(self, _event=None) -> None:
        """Affiche le chemin complet de l'entrée sélectionnée dans la statusbar."""
        sel = self.listbox.curselection()
        if sel:
            self._set_status(self._pdf_paths[sel[0]])
        self._refresh_buttons()

    def _on_double_click(self, _event=None) -> None:
        """Affiche le chemin complet dans une boîte de dialogue au double-clic."""
        sel = self.listbox.curselection()
        if sel:
            messagebox.showinfo("Chemin complet", self._pdf_paths[sel[0]])

    # ── Fusion ────────────────────────────────────────────────────────────────

    def _start_merge(self) -> None:
        """Valide les prérequis, puis lance MergeWorker dans un thread séparé."""
        if not self._pdf_paths:
            messagebox.showwarning("Liste vide", "Ajoutez au moins un fichier PDF.")
            return
        if not self._output_path:
            messagebox.showwarning(
                "Destination manquante", "Choisissez le fichier de sortie."
            )
            return

        if Path(self._output_path).exists():
            if not messagebox.askyesno(
                "Fichier existant",
                f"Le fichier « {Path(self._output_path).name} » existe déjà.\n"
                "L'écraser ?",
            ):
                return

        self.progress.config(maximum=len(self._pdf_paths), value=0)
        self._lock_ui(True)
        self._set_status("Fusion en cours…")

        MergeWorker(
            files=list(self._pdf_paths),
            output_path=self._output_path,
            on_progress=self._cb_progress,
            on_done=self._cb_done,
            on_error=self._cb_error,
        ).start()

    # ── Callbacks MergeWorker → UI (toujours via root.after) ─────────────────

    def _cb_progress(self, current: int, total: int, label: str) -> None:
        """Planifie une mise à jour de progression dans le thread UI."""
        self.root.after(0, self._ui_progress, current, total, label)

    def _ui_progress(self, current: int, total: int, label: str) -> None:
        """Met à jour la progressbar et la statusbar (thread UI)."""
        self.progress["value"] = current
        self._set_status(f"[{current}/{total}] {label}")

    def _cb_done(self, total_pages: int, file_size: int, warnings: List[str]) -> None:
        """Planifie l'affichage du résultat de succès dans le thread UI."""
        self.root.after(0, self._ui_done, total_pages, file_size, warnings)

    def _ui_done(self, total_pages: int, file_size: int, warnings: List[str]) -> None:
        """Affiche la boîte de succès et déverrouille l'interface."""
        self.progress["value"] = self.progress["maximum"]
        size_str  = self._fmt_size(file_size)
        warn_txt  = ""
        if warnings:
            warn_txt = (
                f"\n\n⚠ Avertissements ({len(warnings)}) :\n• "
                + "\n• ".join(warnings)
            )
        msg = (
            f"Fusion terminée avec succès !\n\n"
            f"• Fichiers sources  : {len(self._pdf_paths)}\n"
            f"• Pages totales     : {total_pages}\n"
            f"• Taille du fichier : {size_str}\n\n"
            f"Enregistré dans :\n{self._output_path}"
            f"{warn_txt}"
        )
        self._set_status(
            f"Fusion réussie — {total_pages} pages, {size_str}  ✓", C_STATUS_OK
        )
        messagebox.showinfo("Fusion réussie", msg)
        self._lock_ui(False)

    def _cb_error(self, message: str) -> None:
        """Planifie l'affichage d'une erreur fatale dans le thread UI."""
        self.root.after(0, self._ui_error, message)

    def _ui_error(self, message: str) -> None:
        """Affiche l'erreur fatale, remet la progressbar à zéro et déverrouille l'UI."""
        self.progress["value"] = 0
        self._set_status("Erreur — fusion annulée.", C_STATUS_ERR)
        messagebox.showerror("Erreur de fusion", message)
        self._lock_ui(False)

    # ── Fermeture ─────────────────────────────────────────────────────────────

    def _on_close(self) -> None:
        """Demande confirmation si une fusion est en cours avant de quitter."""
        if self._merging:
            if not messagebox.askyesno(
                "Fusion en cours",
                "Une fusion est actuellement en cours.\nQuitter quand même ?",
            ):
                return
        logger.info("Application fermée par l'utilisateur.")
        self.root.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# POINT D'ENTRÉE
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    """Point d'entrée principal : crée la fenêtre Tk (avec DnD si disponible)."""
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
