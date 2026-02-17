from __future__ import annotations

import os
import sys
import shutil
import tempfile
import traceback
from dataclasses import dataclass
from typing import Optional

import pandas as pd
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QFontMetrics, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFileDialog, QComboBox, QLineEdit, QMessageBox, QCheckBox,
    QTableView, QProgressDialog, QGroupBox, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QHeaderView
)

from app.version import __app_name__, __version__, GITHUB_REPO_SLUG, INSTALLER_ASSET_NAME, INSTALLER_SHA256_ASSET_NAME
from app.updater import GitHubReleaseUpdater
from app.qt_models import DataFrameModel

from app.engine.controllo_fatture_2026 import crea_report_excel as crea_report_excel_2026
from app.engine.controllo_fatture_2025 import crea_report_excel as crea_report_excel_2025


def prepare_preview_df(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()

    drop_cols = [
        c for c in [
            "Country Tariff", "Dest Code Tariff", "Zone Tariff",
            "Nazione tariffa", "Dest Code Tariff", "Zona tariffa",
            "Country_tariff", "Dest_code_tariff", "Zone_tariff",
        ]
        if c in df2.columns
    ]
    if drop_cols:
        df2 = df2.drop(columns=drop_cols, errors="ignore")

    if "Numero DT/FT" in df2.columns:
        import re

        def _fmt_ddt(x):
            if x is None:
                return ""
            try:
                if pd.isna(x):
                    return ""
            except Exception:
                pass

            s = str(x).strip()
            m = re.match(r"^(DT|FT)\s*(.*)$", s, flags=re.IGNORECASE)
            typ = None
            rest = s
            if m:
                typ = m.group(1).upper()
                rest = m.group(2).strip()

            digits = re.sub(r"\D", "", rest)
            if digits:
                digits6 = digits.zfill(6)[-6:]
                return f"{typ} {digits6}".strip() if typ else digits6
            return s

        df2["Numero DT/FT"] = df2["Numero DT/FT"].apply(_fmt_ddt)

    return df2

def resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

@dataclass
class ReportResult:
    msg: str
    report_path: str


class Worker(QThread):
    finished_ok = Signal(object)     # ReportResult
    finished_err = Signal(str)       # error text

    def __init__(self, year: str, pdf_paths: list[str], france_xlsx_path: Optional[str]):
        super().__init__()
        self.year = year
        self.pdf_paths = pdf_paths
        self.france_xlsx_path = france_xlsx_path

    def run(self):
        try:
            crea = crea_report_excel_2026 if self.year == "2026" else crea_report_excel_2025
            tmpdir = tempfile.mkdtemp(prefix="vainieri_report_")
            out_path = os.path.join(tmpdir, "report_controllo_fattura.xlsx")
            msg = crea(self.pdf_paths, out_path, france_xlsx_path=self.france_xlsx_path)
            self.finished_ok.emit(ReportResult(msg=msg, report_path=out_path))
        except Exception:
            self.finished_err.emit(traceback.format_exc())


class MainWindow(QMainWindow):
    PDF_COLS = 3
    PDF_VISIBLE_ROWS = 3
    PDF_ROW_HEIGHT = 26

    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon(resource_path("assets/icon.ico")))
        self.setWindowTitle(f"{__app_name__}  —  v{__version__}")

        self._report_path: Optional[str] = None
        self._df_full: Optional[pd.DataFrame] = None

        # Lista PDF interna (renderizzata nella griglia)
        self._pdf_paths: list[str] = []

        # Menu
        m_file = self.menuBar().addMenu("File")
        act_exit = QAction("Esci", self)
        act_exit.triggered.connect(self.close)
        m_file.addAction(act_exit)

        m_help = self.menuBar().addMenu("Aiuto")
        act_update = QAction("Controlla aggiornamenti…", self)
        act_update.triggered.connect(self.on_check_updates)
        m_help.addAction(act_update)

        # Menubar più sottile (~1/3): font e padding ridotti
        self.menuBar().setStyleSheet("""
            QMenuBar { font-size: 10px; padding: 1px; }
            QMenuBar::item { padding: 2px 6px; margin: 0px; }
            QMenu { font-size: 10px; }
            QMenu::item { padding: 4px 10px; }
        """)

        # Central UI
        root = QWidget()
        self.setCentralWidget(root)
        main = QVBoxLayout(root)

        title = QLabel("📄 Controllo Fatture Vainieri")
        title.setStyleSheet("font-size: 20px; font-weight: 700;")
        main.addWidget(title)

        subtitle = QLabel("Carica una o più fatture PDF e genera automaticamente un unico report Excel.")
        subtitle.setStyleSheet("color: #444;")
        main.addWidget(subtitle)

        # Year
        row_year = QHBoxLayout()
        row_year.addWidget(QLabel("Anno tariffario:"))
        self.cmb_year = QComboBox()
        self.cmb_year.addItems(["2026", "2025"])
        self.cmb_year.setCurrentText("2026")
        row_year.addWidget(self.cmb_year)
        row_year.addStretch(1)
        main.addLayout(row_year)

        # PDFs
        grp_pdf = QGroupBox("Fatture PDF")
        lay_pdf = QVBoxLayout(grp_pdf)

        row_btn = QHBoxLayout()
        self.btn_add_pdf = QPushButton("Aggiungi PDF…")
        self.btn_add_pdf.clicked.connect(self.on_add_pdfs)
        self.btn_remove_pdf = QPushButton("Rimuovi selezionato")
        self.btn_remove_pdf.clicked.connect(self.on_remove_selected_pdf)
        self.btn_clear_pdf = QPushButton("Svuota lista")
        self.btn_clear_pdf.clicked.connect(self.on_clear_pdfs)
        row_btn.addWidget(self.btn_add_pdf)
        row_btn.addWidget(self.btn_remove_pdf)
        row_btn.addWidget(self.btn_clear_pdf)
        row_btn.addStretch(1)
        lay_pdf.addLayout(row_btn)

        # Griglia PDF: 3 colonne, altezza = 3 righe visibili
        self.tbl_pdfs = QTableWidget()
        self.tbl_pdfs.setColumnCount(self.PDF_COLS)
        self.tbl_pdfs.setRowCount(self.PDF_VISIBLE_ROWS)
        self.tbl_pdfs.horizontalHeader().setVisible(False)
        self.tbl_pdfs.verticalHeader().setVisible(False)
        self.tbl_pdfs.setShowGrid(True)
        self.tbl_pdfs.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.tbl_pdfs.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tbl_pdfs.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_pdfs.setWordWrap(False)

        # Colonne che riempiono sempre la larghezza
        self.tbl_pdfs.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_pdfs.horizontalHeader().setMinimumSectionSize(50)

        # No scroll orizzontale (testo elided con "…")
        self.tbl_pdfs.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        for r in range(self.PDF_VISIBLE_ROWS):
            self.tbl_pdfs.setRowHeight(r, self.PDF_ROW_HEIGHT)

        # Altezza fissa: 3 righe visibili (+ bordi)
        self.tbl_pdfs.setFixedHeight(self.PDF_VISIBLE_ROWS * self.PDF_ROW_HEIGHT + 6)

        lay_pdf.addWidget(self.tbl_pdfs)
        main.addWidget(grp_pdf)

        # France excel optional
        grp_fr = QGroupBox("Controllo volumi (opzionale)")
        lay_fr = QHBoxLayout(grp_fr)
        self.txt_france = QLineEdit()
        self.txt_france.setPlaceholderText("Carica l'Excel per il controllo dei volumi (solo spedizioni Francia)…")
        self.txt_france.setReadOnly(True)
        self.btn_browse_fr = QPushButton("Sfoglia…")
        self.btn_browse_fr.clicked.connect(self.on_pick_france_excel)
        self.btn_clear_fr = QPushButton("Pulisci")
        self.btn_clear_fr.clicked.connect(lambda: self.txt_france.setText(""))
        lay_fr.addWidget(self.txt_france, 1)
        lay_fr.addWidget(self.btn_browse_fr)
        lay_fr.addWidget(self.btn_clear_fr)
        main.addWidget(grp_fr)

        info = QLabel("ℹ️ Il confronto con l'excel sarà fatto solo per le spedizioni in Francia.")
        info.setStyleSheet("color: #2a6;")
        main.addWidget(info)

        # Generate
        row_gen = QHBoxLayout()
        self.btn_generate = QPushButton("Genera report")
        self.btn_generate.clicked.connect(self.on_generate)
        self.lbl_status = QLabel("")
        row_gen.addWidget(self.btn_generate)
        row_gen.addWidget(self.lbl_status, 1)
        main.addLayout(row_gen)

        # Legend + filter
        row_opts = QHBoxLayout()
        self.chk_only_errors = QCheckBox("Mostra solo righe con errori")
        self.chk_only_errors.stateChanged.connect(self.apply_filter)
        row_opts.addWidget(self.chk_only_errors)

        legend = QLabel("Legenda:  🟦 volume ≤ 0,3 m³   🟨 Nota / riga non riconosciuta   🟥 Errore")
        legend.setStyleSheet("color:#333;")
        row_opts.addWidget(legend, 1)
        main.addLayout(row_opts)

        # Table preview
        self.table = QTableView()
        self.model = DataFrameModel(pd.DataFrame())
        self.table.setModel(self.model)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        main.addWidget(self.table, 1)

        # Save button
        row_save = QHBoxLayout()
        self.btn_save = QPushButton("⬇️ Salva report Excel…")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self.on_save_report)
        row_save.addWidget(self.btn_save)
        row_save.addStretch(1)
        main.addLayout(row_save)

        self._worker: Optional[Worker] = None
        self._progress: Optional[QProgressDialog] = None

        # Init render
        self.render_pdf_grid()

    # --- Resize: ricalcolo ellissi dei nomi PDF quando cambia larghezza
    def resizeEvent(self, event):
        super().resizeEvent(event)
        # aggiornamento leggero: ricalcola testi elided con nuova larghezza
        if self._pdf_paths:
            self.render_pdf_grid()

    # ---------------- PDF GRID ----------------

    def _elide_filename(self, text: str, col_width_px: int) -> str:
        # padding interno cella (stima) per evitare overflow
        max_px = max(30, col_width_px - 18)
        fm = QFontMetrics(self.tbl_pdfs.font())
        return fm.elidedText(text, Qt.ElideRight, max_px)

    def render_pdf_grid(self):
        n = len(self._pdf_paths)
        rows_needed = max(self.PDF_VISIBLE_ROWS, (n + self.PDF_COLS - 1) // self.PDF_COLS)
        self.tbl_pdfs.setRowCount(rows_needed)

        for r in range(rows_needed):
            self.tbl_pdfs.setRowHeight(r, self.PDF_ROW_HEIGHT)

        self.tbl_pdfs.clearContents()

        # calcola una larghezza media colonna (lo stretch la imposterà in base al viewport)
        viewport_w = max(1, self.tbl_pdfs.viewport().width())
        approx_col_w = max(1, viewport_w // self.PDF_COLS)

        for idx, path in enumerate(self._pdf_paths):
            r = idx // self.PDF_COLS
            c = idx % self.PDF_COLS

            base = os.path.basename(path)
            shown = self._elide_filename(base, approx_col_w)

            it = QTableWidgetItem(shown)
            it.setToolTip(path)  # path completo
            it.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            it.setData(Qt.UserRole, idx)
            self.tbl_pdfs.setItem(r, c, it)

    def on_add_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Seleziona PDF", "", "PDF (*.pdf)")
        if not paths:
            return

        existing = set(self._pdf_paths)
        for p in paths:
            if p not in existing:
                self._pdf_paths.append(p)
                existing.add(p)

        self.render_pdf_grid()

    def on_remove_selected_pdf(self):
        items = self.tbl_pdfs.selectedItems()
        if not items:
            return

        idx = items[0].data(Qt.UserRole)
        if idx is None:
            return

        try:
            idx = int(idx)
            if 0 <= idx < len(self._pdf_paths):
                self._pdf_paths.pop(idx)
        except Exception:
            pass

        self.render_pdf_grid()

    def on_clear_pdfs(self):
        self._pdf_paths = []
        self.render_pdf_grid()

    # ---------------- France Excel ----------------

    def on_pick_france_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Seleziona Excel", "", "Excel (*.xlsx)")
        if path:
            self.txt_france.setText(path)

    # ---------------- Report generation ----------------

    def on_generate(self):
        pdf_paths = list(self._pdf_paths)
        if not pdf_paths:
            QMessageBox.warning(self, "Manca input", "Carica almeno un PDF.")
            return

        year = self.cmb_year.currentText()
        fr = self.txt_france.text().strip() or None

        self.btn_generate.setEnabled(False)
        self.lbl_status.setText("Analisi in corso…")

        self._progress = QProgressDialog("Generazione report…", None, 0, 0, self)
        self._progress.setWindowTitle("Attendere")
        self._progress.setWindowModality(Qt.WindowModal)
        self._progress.show()

        self._worker = Worker(year, pdf_paths, fr)
        self._worker.finished_ok.connect(self.on_generated_ok)
        self._worker.finished_err.connect(self.on_generated_err)
        self._worker.start()

    def on_generated_ok(self, result: ReportResult):
        if self._progress:
            self._progress.close()
            self._progress = None

        self.btn_generate.setEnabled(True)
        self.lbl_status.setText("Report generato con successo!")
        self._report_path = result.report_path
        self.btn_save.setEnabled(True)

        QMessageBox.information(self, "OK", result.msg)

        try:
            df_raw = pd.read_excel(
                self._report_path,
                sheet_name="Controllo",
                engine="openpyxl",
                header=3,
            )
            df_prev = prepare_preview_df(df_raw).head(100)
            self._df_full = df_prev
            self.apply_filter()
        except Exception as e:
            QMessageBox.warning(self, "Anteprima non disponibile", f"Impossibile mostrare l’anteprima: {e}")

    def on_generated_err(self, err: str):
        if self._progress:
            self._progress.close()
            self._progress = None

        self.btn_generate.setEnabled(True)
        self.lbl_status.setText("Errore")
        QMessageBox.critical(self, "Errore", err)

    def apply_filter(self):
        if self._df_full is None:
            self.model.set_df(pd.DataFrame())
            return

        df = self._df_full.copy()
        if self.chk_only_errors.isChecked():
            err = df["Errori"].fillna("") if "Errori" in df.columns else pd.Series([""] * len(df))
            err_fr = df["Errori confronto volume"].fillna("") if "Errori confronto volume" in df.columns else pd.Series([""] * len(df))
            df = df[(err != "") | (err_fr != "")]

        self.model.set_df(df)
        self.table.resizeColumnsToContents()

    def on_save_report(self):
        if not self._report_path or not os.path.exists(self._report_path):
            QMessageBox.warning(self, "Nessun report", "Genera prima un report.")
            return

        out, _ = QFileDialog.getSaveFileName(
            self,
            "Salva report Excel",
            "report_controllo_fattura.xlsx",
            "Excel (*.xlsx)"
        )
        if not out:
            return

        try:
            shutil.copyfile(self._report_path, out)
            QMessageBox.information(self, "Salvato", f"Report salvato in:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    # ---------------- Update ----------------

    def on_check_updates(self):
        updater = GitHubReleaseUpdater(
            repo_slug=GITHUB_REPO_SLUG,
            installer_asset_name=INSTALLER_ASSET_NAME,
            sha256_asset_name=INSTALLER_SHA256_ASSET_NAME,
        )

        try:
            info = updater.check(__version__)
        except Exception as e:
            QMessageBox.warning(self, "Update", f"Impossibile controllare aggiornamenti.\n\n{e}")
            return

        if not info:
            QMessageBox.information(self, "Update", "Sei già all’ultima versione.")
            return

        msg = (
            f"È disponibile un aggiornamento: {info.latest_tag}\n\n"
            f"Note release:\n{info.notes[:1500]}\n\n"
            f"Vuoi scaricare e installare ora?"
        )
        if QMessageBox.question(self, "Aggiornamento disponibile", msg) != QMessageBox.Yes:
            return

        prog = QProgressDialog("Download aggiornamento…", "Annulla", 0, 100, self)
        prog.setWindowTitle("Aggiornamento")
        prog.setWindowModality(Qt.WindowModal)
        prog.show()

        cancelled = {"v": False}

        def progress_cb(stage: str, done: int, total: int):
            if prog.wasCanceled():
                cancelled["v"] = True
                return
            prog.setLabelText(stage)
            if total > 0:
                pct = int(done * 100 / total)
                prog.setValue(max(0, min(100, pct)))
            else:
                prog.setValue(0)

        try:
            installer_path = updater.download_and_verify(info, progress_cb=progress_cb)
            if cancelled["v"]:
                QMessageBox.information(self, "Aggiornamento", "Operazione annullata.")
                return

            prog.close()
            QMessageBox.information(self, "Aggiornamento", "Avvio installer… l’app verrà chiusa.")
            updater.run_installer(installer_path, silent=False)
            QApplication.quit()

        except Exception as e:
            prog.close()
            QMessageBox.critical(self, "Aggiornamento", f"Errore durante l’aggiornamento:\n\n{e}")


def main():
    app = QApplication([])
    app.setWindowIcon(QIcon(resource_path("assets/icon.ico")))
    w = MainWindow()
    w.resize(1100, 700)
    w.show()
    app.exec()


if __name__ == "__main__":
    main()
