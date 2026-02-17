from __future__ import annotations

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QColor


def _cell_str(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


class DataFrameModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self._df = df

    def set_df(self, df: pd.DataFrame):
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df.columns)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None

        r = index.row()
        c = index.column()
        col_name = self._df.columns[c]
        val = self._df.iat[r, c]

        # display
        if role == Qt.DisplayRole:
            if isinstance(val, (int, float)) and not pd.isna(val):
                return f"{val:.2f}"
            return _cell_str(val)

        # coloring: replica Streamlit
        if role == Qt.BackgroundRole:
            row = self._df.iloc[r]
            col_err = "Errori" if "Errori" in self._df.columns else None
            col_err_fr = "Errori confronto volume" if "Errori confronto volume" in self._df.columns else None
            col_note = "Note" if "Note" in self._df.columns else None
            col_vol = "Volume" if "Volume" in self._df.columns else None

            err = _cell_str(row.get(col_err, "")) if col_err else ""
            err_fr = _cell_str(row.get(col_err_fr, "")) if col_err_fr else ""
            note = _cell_str(row.get(col_note, "")) if col_note else ""

            # rosso
            if err or (err_fr and err_fr not in {"non è una spedizione in Francia", "DDT non trovato nel file excel"}):
                return QColor("#FFC7CE")

            # giallo
            if note or err_fr in {"non è una spedizione in Francia", "DDT non trovato nel file excel"}:
                return QColor("#FFEB9C")

            # azzurro
            if col_vol:
                try:
                    v = row.get(col_vol, None)
                    if v is not None and (not pd.isna(v)) and float(v) <= 0.3:
                        return QColor("#C6E0FF")
                except Exception:
                    pass

        return None
