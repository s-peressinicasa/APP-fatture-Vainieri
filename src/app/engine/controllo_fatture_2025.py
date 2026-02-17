import math
import re
import pdfplumber
import pandas as pd
from collections import defaultdict
from pathlib import Path
import openpyxl
from typing import Optional, Tuple, Dict, List, Union, Sequence

# ==============================
# 1) PERCORSO FILE
# ==============================
BASE_DIR = Path(__file__).resolve().parent

#PDF_PATH = r"P:/Stefano/Controllo_fatture_vainieri/FT_  4099_2025-11-24_C001410.pdf"
#TARIFFARIO_PATH = r"P:/Stefano/Controllo_fatture_vainieri/progetto_fatture/prezzi_vainieri.xlsx"
#REPORT_PATH = r"P:/Stefano/Controllo_fatture_vainieri/report_controllo_fattura.xlsx"
PDF_PATH = BASE_DIR / "FT_  4099_2025-11-24_C001410.pdf"   # opzionale
TARIFFARIO_PATH = BASE_DIR / "prezzi_vainieri_2025.xlsx"
REPORT_PATH = BASE_DIR / "report_controllo_fattura.xlsx"   # opzionale

# ==============================
# 2) UTILITIES
# ==============================

def parse_float_eu(s: str) -> float:
    """Converte stringhe tipo '1.234,56' in float Python 1234.56"""
    s = s.replace(".", "").replace(",", ".")
    return float(s)

def qta_is_one(q) -> bool:
    """Ritorna True se q è ~1.0 (con piccola tolleranza)"""
    try:
        return abs(float(q) - 1.0) < 1e-6
    except Exception:
        return False


def normalize_pdf_dt(dt_ft_type: Optional[str], dt_ft_num: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """Normalizza il numero DDT letto dal PDF.

    Regole richieste:
    - Se al posto di DT c'è FT: errore "numero DDT non presente nella fattura"
    - DT viene convertito in 6 cifre (es. 000132)
    """
    if not dt_ft_type or not dt_ft_num:
        return None, "numero DDT non presente nella fattura"

    typ = str(dt_ft_type).strip().upper()
    if typ == "FT":
        return None, "numero DDT non presente nella fattura"
    if typ != "DT":
        return None, "numero DDT non presente nella fattura"

    digits = re.sub(r"\D", "", str(dt_ft_num))
    if not digits:
        return None, "numero DDT non presente nella fattura"

    return digits.zfill(6)[-6:], None


def normalize_pdf_ft(dt_ft_type: Optional[str], dt_ft_num: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """Normalizza il numero FT letto dal PDF.

    - Se al posto di FT c'è DT: errore "numero FT non presente nella fattura"
    - FT viene convertito in 6 cifre (es. 000132)
    """
    if not dt_ft_type or not dt_ft_num:
        return None, "numero FT non presente nella fattura"

    typ = str(dt_ft_type).strip().upper()
    if typ == "DT":
        return None, "numero FT non presente nella fattura"
    if typ != "FT":
        return None, "numero FT non presente nella fattura"

    digits = re.sub(r"\D", "", str(dt_ft_num))
    if not digits:
        return None, "numero FT non presente nella fattura"

    return digits.zfill(6)[-6:], None


def format_dt_ft(dt_ft_type: Optional[str], dt_ft_num: Optional[str]) -> str:
    """Ritorna una stringa del tipo 'DT 000123' o 'FT 000123' per il report."""
    if not dt_ft_type or not dt_ft_num:
        return ""
    typ = str(dt_ft_type).strip().upper()
    digits = re.sub(r"\D", "", str(dt_ft_num))
    if not digits:
        return typ
    digits6 = digits.zfill(6)[-6:]
    return f"{typ} {digits6}".strip()



def normalize_excel_ddt(value) -> Optional[str]:
    """Normalizza il DDT dell'excel prendendo SOLO le ultime 6 cifre."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    digits = re.sub(r"\D", "", str(value))
    if not digits:
        return None
    return digits[-6:].zfill(6)


def load_france_excel(france_xlsx_path: str) -> Tuple[
    Dict[str, float], Dict[str, str], Dict[str, str], Dict[str, str],
    Dict[str, float], Dict[str, str], Dict[str, str], Dict[str, str],
]:
    """Carica l'excel volumi (ex 'FATTURE FRANCIA') e ritorna 2 set di mappe:

    1) Mappe per DDT (colonna 'DDT'):
       - ddt_vol_map: ddt6 -> volume (float arrotondato a 1 decimale)
       - ddt_err_map: ddt6 -> stringa errori (volumi diversi, volume mancante, CAU/Cliente incoerenti, ecc.)
       - ddt_cau_map: ddt6 -> CAU
       - ddt_cliente_map: ddt6 -> Cliente

    2) Mappe per Fattura (colonna 'Fattura'):
       - ft_vol_map: ft6 -> volume
       - ft_err_map: ft6 -> stringa errori
       - ft_cau_map: ft6 -> CAU
       - ft_cliente_map: ft6 -> Cliente

    Note:
    - Header: default riga 8 (header=7). Se non trova le colonne, prova a individuare automaticamente la riga header.
    - Colonne attese: almeno 'Volume' e 'CAU' e una tra 'DDT'/'Fattura'. 'Cliente' è opzionale.
    """

    def _read_with_header_row(header_row_zero_based: int) -> pd.DataFrame:
        return pd.read_excel(france_xlsx_path, header=header_row_zero_based, engine="openpyxl")

    # 1) Lettura standard: header alla riga 8
    df = _read_with_header_row(7)

    # 2) Fallback: trova automaticamente la riga header se mancano colonne base
    def _has_some_key_cols(frame: pd.DataFrame) -> bool:
        cols = [str(c).strip().lower() for c in frame.columns]
        return ("volume" in cols) and (("ddt" in cols) or ("fattura" in cols))

    if not _has_some_key_cols(df):
        preview = pd.read_excel(france_xlsx_path, header=None, nrows=25, engine="openpyxl")
        header_idx = None
        for i in range(len(preview)):
            row = preview.iloc[i].tolist()
            row_norm = [
                str(x).strip().lower()
                for x in row
                if x is not None and not (isinstance(x, float) and pd.isna(x))
            ]
            if ("volume" in row_norm) and (("ddt" in row_norm) or ("fattura" in row_norm)):
                header_idx = i
                break
        if header_idx is not None:
            df = _read_with_header_row(header_idx)

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    # Mappa colonne normalizzate -> nome originale
    norm_cols = {str(c).strip().lower(): c for c in df.columns}

    # DDT (opzionale se presente solo Fattura)
    ddt_col = norm_cols.get("ddt")

    # Fattura (opzionale)
    fattura_col = None
    if "fattura" in norm_cols:
        fattura_col = norm_cols["fattura"]
    else:
        for c in df.columns:
            if "fattura" in str(c).strip().lower():
                fattura_col = c
                break

    # Volume
    vol_col = None
    if "volume" in norm_cols:
        vol_col = norm_cols["volume"]
    else:
        for key in ["vol", "mc", "m3", "volume m3", "volume (m^3)"]:
            if key in norm_cols:
                vol_col = norm_cols[key]
                break
    if vol_col is None:
        for c in df.columns:
            if "vol" in str(c).strip().lower():
                vol_col = c
                break
    if vol_col is None:
        raise ValueError("Colonna volume non trovata nel file excel")

    # CAU (obbligatoria)
    cau_col = None
    if "cau" in norm_cols:
        cau_col = norm_cols["cau"]
    else:
        for c in df.columns:
            if "cau" in str(c).strip().lower():
                cau_col = c
                break
    if cau_col is None:
        raise ValueError("Colonna 'CAU' non trovata nel file excel")

    # Cliente (opzionale)
    cliente_col = None
    if "cliente" in norm_cols:
        cliente_col = norm_cols["cliente"]
    else:
        for c in df.columns:
            if "cliente" in str(c).strip().lower():
                cliente_col = c
                break

    if ddt_col is None and fattura_col is None:
        raise ValueError("Colonna 'DDT' o 'Fattura' non trovata nel file excel")

    # Tieni solo righe che hanno almeno uno tra DDT e Fattura (evita righe di note)
    keep_cols = [c for c in [ddt_col, fattura_col] if c is not None]
    df = df.dropna(subset=keep_cols, how="all")

    def _to_float(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        if isinstance(x, str):
            s = x.strip()
            if not s:
                return None
            s = s.replace(" ", "")
            try:
                return parse_float_eu(s)
            except Exception:
                try:
                    return float(s.replace(",", "."))
                except Exception:
                    return None
        try:
            return float(x)
        except Exception:
            return None

    def _to_str(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        s = str(x).strip()
        return s if s else None

    df["_vol"] = df[vol_col].apply(_to_float)
    df["_cau"] = df[cau_col].apply(_to_str)
    df["_cliente"] = df[cliente_col].apply(_to_str) if cliente_col else None

    if ddt_col:
        df["_ddt6"] = df[ddt_col].apply(normalize_excel_ddt)
    else:
        df["_ddt6"] = None

    if fattura_col:
        df["_ft6"] = df[fattura_col].apply(normalize_excel_ddt)
    else:
        df["_ft6"] = None

    def _build_maps(key_field: str, kind_label: str):
        vol_map: Dict[str, float] = {}
        err_map: Dict[str, str] = {}
        cau_map: Dict[str, str] = {}
        cliente_map: Dict[str, str] = {}

        def _add_err(k: str, msg: str):
            if not k:
                return
            if k in err_map and err_map[k]:
                if msg not in err_map[k]:
                    err_map[k] = f"{err_map[k]} | {msg}"
            else:
                err_map[k] = msg

        df_k = df.dropna(subset=[key_field]).copy()
        df_k = df_k[df_k[key_field].astype(str).str.strip() != ""]
        if df_k.empty:
            return vol_map, err_map, cau_map, cliente_map

        for k, grp in df_k.groupby(key_field):
            # CAU
            causals = [c for c in grp["_cau"].tolist() if c is not None]
            if causals:
                uniq = []
                for c in causals:
                    if c not in uniq:
                        uniq.append(c)
                cau_map[k] = uniq[0]
                if len(uniq) > 1:
                    _add_err(k, f"errore CAU in file excel: causali diverse per lo stesso {kind_label}")
            else:
                cau_map[k] = ""

            # Cliente
            if cliente_col:
                clients = [c for c in grp["_cliente"].tolist() if c is not None]
                if clients:
                    uniq = []
                    for c in clients:
                        if c not in uniq:
                            uniq.append(c)
                    cliente_map[k] = uniq[0]
                    if len(uniq) > 1:
                        _add_err(k, f"errore Cliente in file excel: clienti diversi per lo stesso {kind_label}")
                else:
                    cliente_map[k] = ""
            else:
                cliente_map[k] = ""

            # Volume
            vols = [v for v in grp["_vol"].tolist() if v is not None]
            if not vols:
                _add_err(k, "errore volume in file excel: volume mancante")
                continue

            vols_1 = [round(float(v), 1) for v in vols]
            if any(v != vols_1[0] for v in vols_1[1:]):
                _add_err(k, f"errore volume in file excel: volumi diversi per lo stesso {kind_label}")
                continue

            vol_map[k] = vols_1[0]

        return vol_map, err_map, cau_map, cliente_map

    ddt_vol_map, ddt_err_map, ddt_cau_map, ddt_cliente_map = _build_maps("_ddt6", "ddt")
    ft_vol_map, ft_err_map, ft_cau_map, ft_cliente_map = _build_maps("_ft6", "fattura")

    return (
        ddt_vol_map, ddt_err_map, ddt_cau_map, ddt_cliente_map,
        ft_vol_map, ft_err_map, ft_cau_map, ft_cliente_map,
    )

# ==============================
# 3) CODIFICA TARIFFARIO
# ==============================

xls = pd.ExcelFile(TARIFFARIO_PATH)

def _parse_sheet(xls, candidates):
    """Parse first existing sheet among candidates."""
    for name in candidates:
        if name in xls.sheet_names:
            return xls.parse(name)
    raise ValueError(f"Worksheet named {candidates!r} not found. Available: {xls.sheet_names}")

fr = _parse_sheet(xls, ["FR", "Francia"])
uk = _parse_sheet(xls, ["UK", "GB", "Regno Unito"])
de = _parse_sheet(xls, ["DE", "Germania"])
ie = _parse_sheet(xls, ["IE", "Irlanda"])
ch = _parse_sheet(xls, ["CH", "Svizzera"])
be = _parse_sheet(xls, ["BE", "Belgio"])

# mappa CAP francese -> zona (A/B/C)
dept_to_zone = {}
for _, row in fr.iterrows():
    zone = row["Zona"]
    depts = [d.strip() for d in str(row["codici postali"]).split("-") if d.strip()]
    for d in depts:
        dept_to_zone[d] = zone

# mappa CAP UK -> zona (A/B/C/D)
region_to_zone = {}
for _, row in uk.iterrows():
    zone = row["Zona"]
    regs = [r.strip() for r in str(row["codici postali"]).split("-") if r.strip()]
    for r in regs:
        region_to_zone[r] = zone

# mappa CAP Germania (2 cifre) -> zona (A/B/C/D) dal foglio DE
de_dept_to_zone = {}
for _, row in de.iterrows():
    zone = row["Zona"] if "Zona" in de.columns else str(row["zona"]).upper()
    codes_str = str(row["codici postali"]).replace(" ", "-")
    codes = [c.strip() for c in codes_str.split("-") if c.strip()]
    for c in codes:
        de_dept_to_zone[c] = zone

# dizionari tariffa[zona][fascia_volume] = prezzo €/m3
fr_tariff = {
    row["Zona"]: {
        "0-5": row["da 0 a 5 m^3"],
        "5-10": row["da 5,01 a 10 m^3"],
        "10-15": row["da 10,01 a 15 m^3"],
        "15+": row["da 15,01 m^3"],
    }
    for _, row in fr.iterrows()
}

uk_tariff = {
    row["Zona"]: {
        "0-5": row["da 0 a 5 m^3"],
        "5-10": row["da 5,01 a 10 m^3"],
        "10-15": row["da 10,01 a 15 m^3"],
        "15+": row["da 15,01 m^3"],
    }
    for _, row in uk.iterrows()
}

de_tariff = {
    (row["Zona"] if "Zona" in de.columns else str(row["zona"]).upper()): {
        "0-10": row["da 0 10 m^3"],
        "10+": row["da 10,01 m^3"],
    }
    for _, row in de.iterrows()
}




# Tariffe "Tutto il Paese" (zona ALL)
def _safe_num(v):
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    try:
        return float(v)
    except Exception:
        return None

# Belgio / Svizzera: stessa logica a fasce di volume
be_row = be.iloc[0] if len(be.index) else {}
ch_row = ch.iloc[0] if len(ch.index) else {}
be_tariff = {
    "ALL": {
        "0-5": _safe_num(be_row.get("da 0 a 5 m^3")),
        "5-10": _safe_num(be_row.get("da 5,01 a 10 m^3")),
        "10-15": _safe_num(be_row.get("da 10,01 a 15 m^3")),
        "15+": _safe_num(be_row.get("da 15,01 m^3")),
    }
}
ch_tariff = {
    "ALL": {
        "0-5": _safe_num(ch_row.get("da 0 a 5 m^3")),
        "5-10": _safe_num(ch_row.get("da 5,01 a 10 m^3")),
        "10-15": _safe_num(ch_row.get("da 10,01 a 15 m^3")),
        "15+": _safe_num(ch_row.get("da 15,01 m^3")),
    }
}

# Irlanda: tariffa unica (€/mc)
ie_row = ie.iloc[0] if len(ie.index) else {}
ie_rate = _safe_num(ie_row.get("Tariffa"))


SPECIAL_CLIENT_NAME = "ERCOL FURNITURE LIMITED"
SPECIAL_RATE_OVER_15 = 121.5  # €/mc
def select_tariff(country: str, zone: str, volume: float, cliente: Optional[str] = None, scarico: Optional[str] = None):
    """Restituisce il prezzo €/m3 dalla tabella, in base a paese/zona e volume."""

    c = str(country).strip().upper() if country is not None else ""
    z_raw = str(zone).strip() if zone is not None else None
    z_up = z_raw.upper() if z_raw is not None else None

    # -------------------------
    # Regola speciale ERCOL (2025):
    # sopra 15mc tariffa fissa €/mc = 121.5.
    # Nella pratica il campo "Cliente" non è sempre disponibile (es. spedizioni UK),
    # quindi la regola viene attivata anche se lo scarico contiene "F. EDMONDSON & SONS".
    # -------------------------
    try:
        v = float(volume)
    except Exception:
        v = None

    cli = str(cliente).strip().upper() if cliente is not None else ""
    sca = str(scarico).strip().upper() if scarico is not None else ""
    if c == "UK" and v is not None and v > 15:
        if (cli == str(SPECIAL_CLIENT_NAME).strip().upper()) or ("ERCOL" in cli) or ("EDMONDSON" in sca and "SONS" in sca):
            return float(SPECIAL_RATE_OVER_15)

    # per BE/CH/IE la zona è sempre ALL (tariffa unica paese)
    z_key = z_up
    if c in {"BE", "CH", "IE"} and (z_key is None or z_key == ""):
        z_key = "ALL"

    # per gli altri paesi, se manca la zona non posso calcolare
    if z_key is None:
        return None

    # Se volume non numerico, non posso determinare la fascia
    if v is None:
        return None

    # FRANCIA
    if c == "FR":
        table = fr_tariff
        if v <= 5:
            band = "0-5"
        elif v <= 10:
            band = "5-10"
        elif v <= 15:
            band = "10-15"
        else:
            band = "15+"

    # UK
    elif c == "UK":
        table = uk_tariff
        if v <= 5:
            band = "0-5"
        elif v <= 10:
            band = "5-10"
        elif v <= 15:
            band = "10-15"
        else:
            band = "15+"

    # GERMANIA (solo 0–10 e >10)
    elif c == "DE":
        table = de_tariff
        if v <= 10:
            band = "0-10"
        else:
            band = "10+"

    # BELGIO
    elif c == "BE":
        table = be_tariff
        if v <= 5:
            band = "0-5"
        elif v <= 10:
            band = "5-10"
        elif v <= 15:
            band = "10-15"
        else:
            band = "15+"

    # SVIZZERA
    elif c == "CH":
        table = ch_tariff
        if v <= 5:
            band = "0-5"
        elif v <= 10:
            band = "5-10"
        elif v <= 15:
            band = "10-15"
        else:
            band = "15+"

    # IRLANDA: tariffa unica (€/mc)
    elif c == "IE":
        return ie_rate

    else:
        return None

    # Lookup robusto della zona (es. "Corsica" vs "CORSICA")
    for zk in [z_key, z_raw, (z_raw.title() if z_raw else None)]:
        if not zk:
            continue
        try:
            price = table[zk][band]
            if price is None:
                return None
            try:
                if pd.isna(price):
                    return None
            except Exception:
                pass
            return float(price)
        except Exception:
            continue

    return None
def get_destination_info(scarico: str):
    """
    Ricava:
      - country: 'FR','UK','DE','BE','IE','CH'
      - dest_code: dipartimento (FR), codice regione (UK), CAP a 2 cifre (DE). Vuoto per paesi con tariffa unica.
      - zone: zona tariffaria (oppure 'ALL' per paesi con tariffa "Tutto il Paese")
    dall'indirizzo di scarico.

    Nota: per i paesi con tariffa unica (BE/CH/IE) non serve il codice postale;
    basta la sigla paese dopo il trattino (es. "... - BE").
    """
    if scarico is None or (isinstance(scarico, float) and pd.isna(scarico)):
        return None, None, None

    s = str(scarico).replace("\u2019", "'").upper()

    # prima identifico il paese dalla sigla dopo il trattino
    countries = re.findall(r"-\s*(FR|GB|UK|DE|BE|IE|CH)\b", s)
    if not countries:
        return None, None, None

    country = countries[-1]
    if country == "GB":
        country = "UK"

    # paesi con tariffa unica ("Tutto il Paese"): non richiedo codice
    if country in {"BE", "CH", "IE"}:
        return country, "", "ALL"

    # per FR/UK/DE serve il codice tra parentesi
    if country == "FR":
        m = re.search(r"\((\d{2}|2A|2B)\)\s*-\s*FR\b", s)
        if not m:
            return "FR", None, None
        dept = m.group(1)
        zone = dept_to_zone.get(dept)
        return "FR", dept, zone

    if country == "UK":
        m = re.search(r"\(([A-Z]{1,2})\)\s*-\s*(GB|UK)\b", s)
        if not m:
            return "UK", None, None
        reg = m.group(1)
        # Irlanda del Nord: i CAP iniziano con "BT" -> tariffa Irlanda (IE)
        if str(reg).strip().upper().startswith("BT"):
            return "IE", "", "ALL"
        zone = region_to_zone.get(reg)
        return "UK", reg, zone

    if country == "DE":
        m = re.search(r"\((\d{2})\)\s*-\s*DE\b", s)
        if not m:
            return "DE", None, None
        cap2 = m.group(1)
        zone = de_dept_to_zone.get(cap2)
        return "DE", cap2, zone

    return None, None, None
def parse_shipments(pdf_path: str) -> pd.DataFrame:
    """
    Estrae le spedizioni dal PDF Vainieri.

    Ogni riga del DataFrame corrisponde ad una spedizione (blocco Carico/Scarico)
    e contiene:
    - dati del trasporto (riga TRASPORTO)
    - prezzo PRENOTAZIONE SPEDIZIONE (se presente) usato per la logica di accorpamento
    """
    lines: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.splitlines())

    # trova intestazione tabella
    start = None
    for idx, line in enumerate(lines):
        if line.strip().startswith("DATA NS. RIF. DESCRIZIONE"):
            start = idx
            break
    if start is None:
        raise ValueError("Intestazione 'DATA NS. RIF. DESCRIZIONE' non trovata")

    data = []
    current = None

    pending_date: Optional[str] = None
    last_seen_date: Optional[str] = None

    def flush_current():
        nonlocal current
        if current is not None:
            data.append(current)
        current = None

    _DATE_PAT = r"\d{2}/\d{2}/\d{2}"

    def is_date_only(line: str) -> bool:
        return bool(re.match(rf"^{_DATE_PAT}$", line.strip()))

    def is_new_carico_with_date(line: str) -> bool:
        return bool(re.match(rf"^{_DATE_PAT}\s+\S+\s+Carico:", line))

    def is_new_carico_with_ref(line: str) -> bool:
        # alcuni PDF non ripetono la data sulla riga: es. '10680/SH Carico: ...'
        return bool(re.match(r"^\S+\s+Carico:", line))

    def is_new_carico(line: str) -> bool:
        return is_new_carico_with_date(line) or is_new_carico_with_ref(line)

    def is_block_boundary(line: str) -> bool:
        line = line.strip()
        if is_new_carico(line):
            return True
        if line.startswith("DT ") or line.startswith("FT "):
            return True
        if (
            line.startswith("TRASPORTO")
            or line.startswith("FUEL TAX")
            or line.startswith("PRENOTAZIONE")
            or line.startswith("DOGANA")
        ):
            return True
        if line.startswith("STORNO") or line.startswith("RECUPERO"):
            return True
        if line.startswith("COD. IVA"):
            return True
        return False

    i = start + 1
    while i < len(lines):
        line = lines[i].strip()

        # data su riga isolata (a volte la colonna DATA viene estratta separata)
        if is_date_only(line):
            flush_current()
            pending_date = line
            last_seen_date = line
            i += 1
            continue

        # fine tabella di una pagina
        if line.startswith("COD. IVA IMPONIBILE"):
            flush_current()
            current = None
            i += 1
            continue

        # NUOVA SPEDIZIONE (DATA + NS.RIF + Carico)
        if is_new_carico(line):
            flush_current()
            m = re.match(r"^(\d{2}/\d{2}/\d{2})\s+(\S+)\s+Carico:(.*)$", line)
            if m:
                date, ns_rif, rest = m.groups()
                last_seen_date = date
                pending_date = None
            else:
                m2 = re.match(r"^(\S+)\s+Carico:(.*)$", line)
                if not m2:
                    raise ValueError(f"Riga 'Carico' non riconosciuta: {line}")
                ns_rif, rest = m2.groups()
                date = pending_date or last_seen_date or ""
                pending_date = None
            current = {
                "data": date,
                "ns_rif": ns_rif,
                "carico": rest.strip(),
                "scarico": "",
                "dt_ft_type": None,
                "dt_ft_num": None,
                "trasporto_volume": None,
                "trasporto_qta": None,
                "trasporto_pu": None,
                "trasporto_tot": None,
                "trasporto_cod_iva": None,
                # PRENOTAZIONE SPEDIZIONE (servizio)
                "prenotazione_qta": None,
                "prenotazione_pu": None,
                "prenotazione_tot": None,
                "prenotazione_cod_iva": None,
                "note": "",
            }
            i += 1
            continue

        if current is None:
            i += 1
            continue

        # eventuale seconda riga di Carico (es. riga solo "IT")
        if (
            current["scarico"] == ""
            and current["carico"]
            and not line.startswith("Scarico:")
            and not is_block_boundary(line)
        ):
            current["carico"] += " " + line
            i += 1
            continue

        # BLOCCO SCARICO (può essere su più righe)
        if line.startswith("Scarico:"):
            dest = line
            j = i + 1
            while j < len(lines) and not is_block_boundary(lines[j].strip()):
                dest += " " + lines[j].strip()
                j += 1
            current["scarico"] = dest
            i = j
            continue

        # DT / FT
        if line.startswith("DT ") or line.startswith("FT "):
            typ, num = line.split(maxsplit=1)
            current["dt_ft_type"] = typ
            current["dt_ft_num"] = num.strip()
            i += 1
            continue

        # RIGA TRASPORTO (volume, q.tà, prezzi, IVA)
        if line.startswith("TRASPORTO"):
            m = re.match(
                r"^TRASPORTO\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+(\S+)$",
                line,
            )

            if m:
                vol, qta, pu, tot, cod = m.groups()
                current["trasporto_volume"] = parse_float_eu(vol)
                current["trasporto_qta"] = parse_float_eu(qta)
                current["trasporto_pu"] = parse_float_eu(pu)
                current["trasporto_tot"] = parse_float_eu(tot)
                current["trasporto_cod_iva"] = cod
            else:
                # Caso non riconosciuto (es. "TRASPORTO C/SERVIZIO")
                # Non blocco l'elaborazione: segno la riga come "non controllabile"
                current["trasporto_volume"] = None
                current["trasporto_qta"] = None
                current["trasporto_pu"] = None
                current["trasporto_tot"] = None
                current["trasporto_cod_iva"] = None
                current["note"] = f"Riga TRASPORTO non riconosciuta: {line}"

            i += 1
            continue

        # PRENOTAZIONE SPEDIZIONE (serve per accorpamento)
        if line.startswith("PRENOTAZIONE"):
            # Esempio: "PRENOTAZIONE SPEDIZIONE 1,000 0,67 0,67 E8C"
            m = re.match(
                r"^PRENOTAZIONE(?:\s+SPEDIZIONE)?\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+(\S+)$",
                line,
            )
            if m:
                qta, pu, tot, cod = m.groups()
                current["prenotazione_qta"] = parse_float_eu(qta)
                current["prenotazione_pu"] = parse_float_eu(pu)
                current["prenotazione_tot"] = parse_float_eu(tot)
                current["prenotazione_cod_iva"] = cod
            else:
                # Non blocco, ma segnalo nelle note
                note = current.get("note", "")
                extra = f"Riga PRENOTAZIONE non riconosciuta: {line}"
                current["note"] = f"{note}; {extra}" if note else extra
            i += 1
            continue

        # altre righe (Fuel tax, dogana...) non servono per il controllo
        i += 1

    flush_current()
    df = pd.DataFrame(data)

    # Info di destinazione (FR/UK/DE, dip/regione/CAP, zona)
    df[["country", "dest_code", "zone"]] = df["scarico"].apply(
        lambda s: pd.Series(get_destination_info(s))
    )

    # Info tariffario: di default come lo scarico, ma se lo scarico contiene 'peressini'
    # allora la zona tariffaria va calcolata sul CARICO (stesse regole di parsing).
    def _tariff_addr(row):
        scar = str(row.get('scarico') or '')
        if re.search(r'peressini', scar, flags=re.IGNORECASE):
            return row.get('carico') or scar
        return scar

    df[["country_tariff", "dest_code_tariff", "zone_tariff"]] = df.apply(
        lambda r: pd.Series(get_destination_info(_tariff_addr(r))), axis=1
    )

    # ==========================
    # NUOVA LOGICA DI ACCORPO (richiesta):
    #
    # Si usa il prezzo unitario della riga "PRENOTAZIONE SPEDIZIONE" per capire
    # quante fatture compongono la spedizione (costo totale prenotazione=2€ ripartito):
    #   2,00 -> 1 fattura (NON accorpare)
    #   1,00 -> 2 fatture
    #   0,66 o 0,67 -> 3 fatture
    #   0,50 -> 4 fatture
    #
    # Per capire quali fatture accorpare: stesso SCARICO (normalizzato) e righe con lo stesso scarico (anche non consecutive).
    #
    # Se il numero di fatture trovate non coincide con quello atteso dal prezzo,
    # viene scritto un errore in colonna `errori` (tramite `accorpamento_error`).
    # ==========================
    df = df.reset_index(drop=True)
    n = len(df)

    def _norm_scarico(s: str) -> str:
        if s is None or (isinstance(s, float) and pd.isna(s)):
            return ""
        s = str(s).replace("\u2019", "'")  # apostrofo “curly”
        s = re.sub(r"\s+", " ", s).strip()
        if s.lower().startswith("scarico:"):
            s = s[len("scarico:"):].strip()
        return s.upper()

    def _expected_n_from_pren(price: Optional[float]) -> Optional[int]:
        if price is None or (isinstance(price, float) and pd.isna(price)):
            return None
        try:
            p = float(price)
        except Exception:
            return None

        def close(a: float, b: float, tol: float = 0.02) -> bool:
            return abs(a - b) <= tol

        # mapping esplicito
        if close(p, 2.0):
            return 1
        if close(p, 1.0):
            return 2
        if close(p, 0.5):
            return 4
        if close(p, 0.66) or close(p, 0.67):
            return 3

        # fallback generico: prezzo ~ 2/n
        if p <= 0:
            return None
        n_guess = int(round(2.0 / p))
        if n_guess < 1 or n_guess > 10:
            return None
        if abs(p - (2.0 / n_guess)) <= 0.03:
            return n_guess
        return None


    def _fmt_price_it(val: Optional[float]) -> str:
        """Formatta il prezzo in stile italiano per i messaggi (es. 1 -> '1', 0.5 -> '0,5')."""
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""
        try:
            p = float(val)
        except Exception:
            return ""
        s = f"{p:.2f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")

    scarico_key = df["scarico"].apply(_norm_scarico)

    pren_price = df["prenotazione_pu"].copy() if "prenotazione_pu" in df.columns else pd.Series([None] * n)
    if "prenotazione_tot" in df.columns:
        pren_price = pren_price.fillna(df["prenotazione_tot"])

    exp_n_raw = pren_price.apply(_expected_n_from_pren)

    group_ids = [""] * n
    accorp_errors = [""] * n
    accorp_atteso: List[Optional[int]] = [None] * n

    def _append_err(idx: int, msg: str):
        if not msg:
            return
        if accorp_errors[idx]:
            if msg not in accorp_errors[idx]:
                accorp_errors[idx] += " | " + msg
        else:
            accorp_errors[idx] = msg

    # --- Raggruppamento NON per vicinanza ---
    # L'accorpamento deve cercare in tutto il PDF le righe con lo stesso scarico.
    # Usiamo la combinazione (scarico_normalizzato, n_atteso) e poi spezzettiamo
    # in blocchi da n_atteso (es. 6 righe con n_atteso=3 => 2 gruppi da 3).
    idx_by_keyexp: Dict[Tuple[str, int], List[int]] = defaultdict(list)
    single_idxs: List[int] = []

    for idx in range(n):
        price_i = pren_price.iloc[idx]
        exp = exp_n_raw.iloc[idx]

        # exp_n_raw può essere NaN (float) se il prezzo prenotazione manca/non è interpretabile
        if exp is None or (isinstance(exp, float) and pd.isna(exp)):
            exp = None
        else:
            try:
                exp = int(exp)
            except Exception:
                exp = None

        # se prezzo mancante/non riconosciuto: non accorpo ma segnalo
        if exp is None:
            exp = 1
            if price_i is None or (isinstance(price_i, float) and pd.isna(price_i)):
                _append_err(idx, "Errore prezzo prenotazione spedizione; prezzo non trovato quindi impossibile determinare spedizioni da unire")
            else:
                p_str = _fmt_price_it(price_i)
                _append_err(idx, f"Errore prezzo prenotazione spedizione; trovato prezzo={p_str} non gestito quindi impossibile determinare spedizioni da unire")

        accorp_atteso[idx] = exp

        if exp <= 1:
            single_idxs.append(idx)
        else:
            idx_by_keyexp[(scarico_key.iloc[idx], exp)].append(idx)

    gid = 0

    # 1) Spedizioni singole: group id univoco per riga
    for idx in single_idxs:
        group_ids[idx] = f"G{gid}"
        gid += 1

    # 2) Spedizioni accorpate: raggruppo per scarico in tutto il PDF
    #    e poi divido in chunk da n_atteso
    for (key, exp), idxs in sorted(idx_by_keyexp.items(), key=lambda kv: min(kv[1]) if kv[1] else 10**9):
        idxs_sorted = sorted(idxs)

        # Caso 1: numero righe = atteso -> un solo gruppo
        if len(idxs_sorted) == exp:
            gid_str = f"G{gid}"
            for k in idxs_sorted:
                group_ids[k] = gid_str
            gid += 1
            continue

        # Caso 2: multipli esatti -> spezza in gruppi da exp
        if exp > 0 and (len(idxs_sorted) % exp == 0):
            for c in range(len(idxs_sorted) // exp):
                chunk = idxs_sorted[c * exp : (c + 1) * exp]
                gid_str = f"G{gid}"
                for k in chunk:
                    group_ids[k] = gid_str
                gid += 1
            continue

        # Caso 3: mismatch -> raggruppo tutto insieme e segnalo errore
        gid_str = f"G{gid}"
        gid += 1

        # prezzo "rappresentativo" per messaggio (prendo il primo disponibile)
        p_val = None
        for k in idxs_sorted:
            v = pren_price.iloc[k]
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                p_val = float(v)
                break
        p_str = _fmt_price_it(p_val)

        msg = f"Errore prezzo prenotazione spedizione; trovato prezzo={p_str} quindi attese {exp} spedizioni da unire, trovata/e {len(idxs_sorted)}"
        for k in idxs_sorted:
            group_ids[k] = gid_str
            _append_err(k, msg)

    df["group_id"] = group_ids
    df["accorpamento_atteso"] = accorp_atteso
    df["accorpamento_error"] = accorp_errors

    return df
# ==============================
# 5) CONTROLLO PREZZI
# ==============================

def check_invoice(df_ship: pd.DataFrame, tolerance: float = 0.01):
    """
    Esegue tutti i controlli richiesti.
    Ritorna una lista di dizionari con gli errori trovati.
    """
    errors = []

    for gid, grp in df_ship.groupby("group_id"):
        grp = grp.sort_values("ns_rif")
        # Salta righe senza importi/volumi (es. TRASPORTO C/SERVIZIO)
        grp = grp.dropna(subset=["trasporto_volume", "trasporto_tot", "trasporto_qta"])
        if grp.empty:
            continue
        country_dest = grp.iloc[0]["country"]
        dest_code_dest = grp.iloc[0]["dest_code"]
        zone_dest = grp.iloc[0]["zone"]

        # Per i controlli tariffari uso i campi *_tariff se presenti
        country = grp.iloc[0].get("country_tariff") or country_dest
        dest_code = grp.iloc[0].get("dest_code_tariff") or dest_code_dest
        zone = grp.iloc[0].get("zone_tariff") or zone_dest

        # somma volumi “grezza”
        vol_total_raw = grp["trasporto_volume"].sum()
        trans_total = grp["trasporto_tot"].sum()

        # volume per fascia tariffaria: quello grezzo
        cliente = grp.iloc[0].get("cliente", "")
        tariff = select_tariff(country, zone, vol_total_raw, cliente=cliente, scarico=grp.iloc[0].get("scarico", ""))

        # -------------------------
        # CASO SPEDIZIONI ACCORPATE
        # -------------------------
        if len(grp) > 1:
            if vol_total_raw < 1:
                vol_total_round = 1.0
            else:
                # epsilon per evitare che 1.20000000002 diventi 1.3
                vol_total_round = math.ceil(vol_total_raw * 10 - 1e-9) / 10.0

            if tariff is None:
                errors.append(
                    {
                        "tipo_errore": "tariffa_mancante",
                        "messaggio": "Tariffa mancante per spedizione accorpata",
                        "group_id": gid,
                        "country": country,
                        "dest_code": dest_code,
                        "zone": zone,
                        "volume_totale": vol_total_raw,
                        "volume_totale_arrotondato": vol_total_round,
                    }
                )
                continue

            prezzo_m3_fatt = trans_total / vol_total_round

            if abs(prezzo_m3_fatt - tariff) > tolerance:
                errors.append(
                    {
                        "tipo_errore": "accorpata_prezzo_m3_errato",
                        "messaggio": f"Tariffa €/m³ errata: atteso {tariff}, trovato {round(prezzo_m3_fatt, 4)}",
                        "tariffa_attesa": tariff,
                        "tariffa_trovata": round(prezzo_m3_fatt, 4),
                        "group_id": gid,
                        "country": country,
                        "dest_code": dest_code,
                        "zone": zone,
                        "ns_rif": list(grp["ns_rif"]),
                        "volume_totale": vol_total_raw,
                        "volume_totale_arrotondato": vol_total_round,
                        "trasporto_totale": trans_total,
                        "prezzo_m3_fatturato": round(prezzo_m3_fatt, 4),
                        "prezzo_m3_tariffario": tariff,
                    }
                )

        # -------------------------
        # CASO SPEDIZIONE SINGOLA
        # -------------------------
        else:
            row = grp.iloc[0]
            vol = row["trasporto_volume"]
            qta = row["trasporto_qta"]

            base_info = {
                "group_id": gid,
                "country": country,
                "dest_code": dest_code,
                "zone": zone,
                "ns_rif": row["ns_rif"],
                "dt_ft_num": row["dt_ft_num"],
                "volume": vol,
                "qta": qta,
                "trasporto_tot": row["trasporto_tot"],
            }

            # 1) volume > 1 m3
            if vol > 1.0:
                vol_arrotondato = math.ceil(vol * 10 - 1e-9) / 10.0

                if abs(qta - vol_arrotondato) > tolerance:
                    e = base_info.copy()
                    e.update(
                        {
                            "tipo_errore": "volume>1_qta_errata",
                            "messaggio": f"Volume arrotondato errato: atteso {vol_arrotondato}",
                            "volume_arrotondato_atteso": vol_arrotondato,
                        }
                    )
                    errors.append(e)

                tariff = select_tariff(country, zone, vol, cliente=row.get("cliente", ""), scarico=row.get("scarico", ""))
                if tariff is None:
                    e = base_info.copy()
                    e.update(
                        {
                            "tipo_errore": "tariffa_mancante",
                            "messaggio": "Tariffa mancante",
                        }
                    )
                    errors.append(e)
                else:
                    prezzo_m3_fatt = row["trasporto_tot"] / vol_arrotondato
                    if abs(prezzo_m3_fatt - tariff) > tolerance:
                        e = base_info.copy()
                        e.update(
                            {
                                "tipo_errore": "volume>1_prezzo_m3_errato",
                                "messaggio": f"Tariffa €/m³ errata: atteso {tariff}, trovato {round(prezzo_m3_fatt, 4)}",
                                "tariffa_attesa": tariff,
                                "tariffa_trovata": round(prezzo_m3_fatt, 4),
                                "volume_arrotondato": vol_arrotondato,
                                "prezzo_m3_fatturato": round(prezzo_m3_fatt, 4),
                                "prezzo_m3_tariffario": tariff,
                            }
                        )
                        errors.append(e)

            # 2) volume < 1 m3
            else:
                vol_arrotondato = math.ceil(vol * 10 - 1e-9) / 10.0

                if abs(qta - vol_arrotondato) > tolerance:
                    e = base_info.copy()
                    e.update(
                        {
                            "tipo_errore": "volume<1_qta_errata",
                            "messaggio": f"Volume arrotondato errato: atteso {vol_arrotondato}",
                            "volume_arrotondato_atteso": vol_arrotondato,
                        }
                    )
                    errors.append(e)

                # PREZZO: minimo fatturabile 1 m³
                tariff = select_tariff(country, zone, 1.0, cliente=row.get("cliente", ""), scarico=row.get("scarico", ""))

                if tariff is None:
                    e = base_info.copy()
                    e.update(
                        {
                            "tipo_errore": "tariffa_mancante",
                            "messaggio": "Tariffa mancante",
                        }
                    )
                    errors.append(e)
                else:
                    prezzo_atteso = tariff
                    if abs(row["trasporto_tot"] - prezzo_atteso) > tolerance:
                        e = base_info.copy()
                        e.update(
                            {
                                "tipo_errore": "volume<1_prezzo_errato",
                                "messaggio": f"Prezzo totale errato: atteso {prezzo_atteso}, trovato {row['trasporto_tot']}",
                                "prezzo_atteso": prezzo_atteso,
                                "prezzo_trovato": row["trasporto_tot"],
                                "volume_arrotondato": vol_arrotondato,
                                "prezzo_fatturato": row["trasporto_tot"],
                                "prezzo_m3_tariffario": tariff,
                            }
                        )
                        errors.append(e)

    return errors

# ==============================
# 6) MAIN + CREAZIONE EXCEL
# ==============================
def crea_report_excel(
    pdf_path: Union[str, Sequence[str]],
    report_path: str,
    tolerance: float = 0.01,
    france_xlsx_path: Optional[str] = None,
):
    """Legge una o più fatture PDF, controlla i prezzi e crea il file Excel di report.

    - Se `pdf_path` è una stringa: elabora un solo PDF.
    - Se `pdf_path` è una lista/tupla di stringhe: elabora tutti i PDF e unisce i risultati
      in un unico report.

    Se `france_xlsx_path` è fornito, effettua anche il confronto volumi con l'excel
    (solo per le spedizioni Francia).
    """

    # --------------------------
    # SUPPORTO MULTI-PDF
    # --------------------------
    if isinstance(pdf_path, (list, tuple)):
        pdf_paths = [str(p) for p in pdf_path]
    else:
        pdf_paths = [str(pdf_path)]

    if not pdf_paths:
        raise ValueError("Nessun PDF fornito")

    dfs: List[pd.DataFrame] = []
    multi = len(pdf_paths) > 1

    for i_pdf, p in enumerate(pdf_paths, start=1):
        df_tmp = parse_shipments(p)

        # Se ho più PDF, aggiungo una colonna per distinguere le righe
        # e prefisso il group_id per evitare collisioni tra file diversi
        if multi:
            df_tmp["pdf"] = Path(p).name
            if "group_id" in df_tmp.columns:
                df_tmp["group_id"] = df_tmp["group_id"].apply(lambda g: f"P{i_pdf}_{g}")

        dfs.append(df_tmp)

    df_ship = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]

    # ==========================
    # Preparazione colonne tecniche (DDT6) + caricamento excel volumi (se presente)
    # ==========================
    df_ship = df_ship.copy()

        # colonne tecniche:
    # - ddt6: ultime 6 cifre del DDT (solo se DT)
    # - ft6 : ultime 6 cifre della Fattura (solo se FT)
    # - dt_ft_show: valore da mostrare nel report ('DT 000123' / 'FT 000123')
    ddt6_list: List[str] = []
    ft6_list: List[str] = []
    dtft_show_list: List[str] = []

    for _, row in df_ship.iterrows():
        typ = row.get("dt_ft_type")
        num = row.get("dt_ft_num")
        dtft_show_list.append(format_dt_ft(typ, num))

        ddt6, _ = normalize_pdf_dt(typ, num)
        ft6, _ = normalize_pdf_ft(typ, num)
        ddt6_list.append(ddt6 or "")
        ft6_list.append(ft6 or "")

    df_ship["ddt6"] = ddt6_list
    df_ship["ft6"] = ft6_list
    df_ship["dt_ft_show"] = dtft_show_list

    # Colonne "cliente" e "causale" (da excel, se disponibile)
    if "cliente" not in df_ship.columns:
        df_ship["cliente"] = ""
    if "causale" not in df_ship.columns:
        df_ship["causale"] = ""

    excel_load_error = None
    ddt_excel_vol_map: Dict[str, float] = {}
    ddt_excel_err_map: Dict[str, str] = {}
    ddt_excel_cau_map: Dict[str, str] = {}
    ddt_excel_cliente_map: Dict[str, str] = {}

    ft_excel_vol_map: Dict[str, float] = {}
    ft_excel_err_map: Dict[str, str] = {}
    ft_excel_cau_map: Dict[str, str] = {}
    ft_excel_cliente_map: Dict[str, str] = {}

    if france_xlsx_path:
        try:
            (
                ddt_excel_vol_map, ddt_excel_err_map, ddt_excel_cau_map, ddt_excel_cliente_map,
                ft_excel_vol_map, ft_excel_err_map, ft_excel_cau_map, ft_excel_cliente_map,
            ) = load_france_excel(france_xlsx_path)

            # Riempio subito: serve anche per la regola speciale cliente nel controllo tariffe
            df_ship["causale"] = ""
            df_ship["cliente"] = ""

            _typ = df_ship["dt_ft_type"].fillna("").astype(str).str.upper()
            mask_dt = _typ == "DT"
            mask_ft = _typ == "FT"

            df_ship.loc[mask_dt, "causale"] = df_ship.loc[mask_dt, "ddt6"].map(ddt_excel_cau_map).fillna("").astype(str)
            df_ship.loc[mask_dt, "cliente"] = df_ship.loc[mask_dt, "ddt6"].map(ddt_excel_cliente_map).fillna("").astype(str)

            df_ship.loc[mask_ft, "causale"] = df_ship.loc[mask_ft, "ft6"].map(ft_excel_cau_map).fillna("").astype(str)
            df_ship.loc[mask_ft, "cliente"] = df_ship.loc[mask_ft, "ft6"].map(ft_excel_cliente_map).fillna("").astype(str)
        except Exception as e:
            excel_load_error = str(e)

    # Controllo tariffe
    errori = check_invoice(df_ship, tolerance=tolerance)

    # costruzione struttura errori per riga (controlli prezzo esistenti)
    errs_per_row = defaultdict(list)
    for e in errori:
        gid = e.get("group_id")
        ns_val = e.get("ns_rif")

        if isinstance(ns_val, list):
            ns_list = ns_val
        elif ns_val is None:
            ns_list = list(df_ship.loc[df_ship["group_id"] == gid, "ns_rif"])
        else:
            ns_list = [ns_val]

        for ns in ns_list:
            mask = (df_ship["group_id"] == gid) & (df_ship["ns_rif"] == ns)
            for idx in df_ship.index[mask]:
                errs_per_row[idx].append(e)

    
    # Aggiungo eventuali errori di accorpamento (nuova logica: PRENOTAZIONE SPEDIZIONE)
    if "accorpamento_error" in df_ship.columns:
        for idx, err_msg in df_ship["accorpamento_error"].items():
            if isinstance(err_msg, str) and err_msg.strip():
                errs_per_row[idx].append(
                    {
                        "tipo_errore": "errore_accorpamento",
                        "messaggio": err_msg.strip(),
                        "group_id": df_ship.loc[idx, "group_id"] if "group_id" in df_ship.columns else None,
                        "ns_rif": df_ship.loc[idx, "ns_rif"] if "ns_rif" in df_ship.columns else None,
                    }
                )
# ==========================
    # 6A) CONFRONTO VOLUMI CON EXCEL (SOLO FRANCIA)
    # ==========================
    errs_fr_per_row: Dict[int, List[str]] = defaultdict(list)

    df_out = df_ship.copy()

    # Se l'excel non è stato caricato correttamente, segnalo l'errore alle sole righe FR
    if france_xlsx_path and excel_load_error:
        for idx, row in df_out.iterrows():
            if row.get("country") == "FR":
                errs_fr_per_row[idx].append(f"file excel non compatibile: {excel_load_error}")

    if france_xlsx_path and not excel_load_error:
        # 1) Validazione riga-per-riga + indice per somma volumi
        pdf_idx_by_ddt: Dict[str, List[int]] = defaultdict(list)
        pdf_idx_by_ft: Dict[str, List[int]] = defaultdict(list)

        for idx, row in df_out.iterrows():
            country = row.get("country")

            # confronto volumi solo per FR
            if country != "FR":
                if country:
                    errs_fr_per_row[idx].append("non è una spedizione in Francia")
                continue

            typ = str(row.get("dt_ft_type") or "").strip().upper()

            # --- match per DDT (DT) ---
            if typ == "DT":
                ddt6 = row.get("ddt6") or ""
                if not ddt6:
                    ddt6, err = normalize_pdf_dt(row.get("dt_ft_type"), row.get("dt_ft_num"))
                    if err:
                        errs_fr_per_row[idx].append(err)
                        continue

                if ddt6 in ddt_excel_err_map:
                    errs_fr_per_row[idx].append(ddt_excel_err_map[ddt6])
                    continue

                if ddt6 not in ddt_excel_vol_map:
                    errs_fr_per_row[idx].append("DDT non trovato nel file excel")
                    continue

                pdf_idx_by_ddt[ddt6].append(idx)
                continue

            # --- match per Fattura (FT) ---
            if typ == "FT":
                ft6 = row.get("ft6") or ""
                if not ft6:
                    ft6, err = normalize_pdf_ft(row.get("dt_ft_type"), row.get("dt_ft_num"))
                    if err:
                        errs_fr_per_row[idx].append(err)
                        continue

                if ft6 in ft_excel_err_map:
                    errs_fr_per_row[idx].append(ft_excel_err_map[ft6])
                    continue

                if ft6 not in ft_excel_vol_map:
                    errs_fr_per_row[idx].append("Fattura non trovata nel file excel")
                    continue

                pdf_idx_by_ft[ft6].append(idx)
                continue

            errs_fr_per_row[idx].append("Numero DT/FT non presente nella fattura")

        # 2) Confronto volumi (match non 1:1)
        for ddt6, idxs in pdf_idx_by_ddt.items():
            vol_pdf = df_out.loc[idxs, "trasporto_volume"].dropna().sum()
            vol_xls = ddt_excel_vol_map.get(ddt6)
            if vol_xls is None:
                continue
            if round(float(vol_pdf), 1) != round(float(vol_xls), 1):
                msg = (
                    "volume diverso tra fattura e file excel "
                    f"(PDF={round(float(vol_pdf), 1)} / Excel={round(float(vol_xls), 1)})"
                )
                for i in idxs:
                    errs_fr_per_row[i].append(msg)

        for ft6, idxs in pdf_idx_by_ft.items():
            vol_pdf = df_out.loc[idxs, "trasporto_volume"].dropna().sum()
            vol_xls = ft_excel_vol_map.get(ft6)
            if vol_xls is None:
                continue
            if round(float(vol_pdf), 1) != round(float(vol_xls), 1):
                msg = (
                    "volume diverso tra fattura e file excel "
                    f"(PDF={round(float(vol_pdf), 1)} / Excel={round(float(vol_xls), 1)})"
                )
                for i in idxs:
                    errs_fr_per_row[i].append(msg)



# 6B) COLONNE ERRORI + EXPORT
    # ==========================
    df_out["errori"] = df_out.index.map(
        lambda i: "; ".join(
            err.get("messaggio", err.get("tipo_errore", ""))
            for err in errs_per_row[i]
        ) if i in errs_per_row else ""
    )
    df_out["has_error"] = df_out["errori"].apply(lambda x: x != "")

    # nuova colonna errori (non sovrascrive quella esistente)
    df_out["Errori confronto volume"] = df_out.index.map(
        lambda i: "; ".join(errs_fr_per_row[i]) if i in errs_fr_per_row else ""
    )
    df_out["has_error_volume_francia"] = df_out["Errori confronto volume"].apply(lambda x: x != "")

    any_err = df_out["has_error"].any() or df_out["has_error_volume_francia"].any()
    if not any_err:
        msg = "Nessun errore: tutte le spedizioni risultano fatturate correttamente."
    else:
        msg = "Risultano degli errori, controllare le fatture evidenziate."

    if excel_load_error:
        msg = f"Errore lettura excel volumi: {excel_load_error}. {msg}"

    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        sheet_name = "Controllo"

        # esportiamo senza le colonne tecniche
        drop_final_cols = [
            "has_error",
            "has_error_volume_francia",
            "group_id",              # Group
            "ddt6",                  # Ddt6
            "ft6",                   # Ft6
            "dt_ft_num",            # numero grezzo (solo debug)
            "dt_ft_type",            # DT/FT (prima del rename)
            "trasporto_cod_iva",     # Trasporto Cod Iva
            "carico"                 # Carico
            "country_tariff",        # tariff debug
            "dest_code_tariff",      # tariff debug
            "zone_tariff",           # tariff debug
        ]

        # Colonne tecniche non necessarie nel report finale
        drop_final_cols += [
            "prenotazione_qta",
            "prenotazione_pu",
            "prenotazione_tot",
            "prenotazione_cod_iva",
            "accorpamento_atteso",
            "accorpamento_error",
        ]

        df_export = df_out.drop(columns=drop_final_cols, errors="ignore")

        # Metto la Causale subito dopo il Numero DT/FT (se presente)
        if "causale" in df_export.columns:
            base_cols = [c for c in ["pdf", "data", "ns_rif", "cliente", "scarico", "dt_ft_show", "causale"] if c in df_export.columns]
            remaining = [c for c in df_export.columns if c not in base_cols]
            df_export = df_export[base_cols + remaining]

        # 1) rinomine “hard” (DEVONO matchare i nomi originali)
        rename_map = {
            "pdf": "PDF",
            "dt_ft_show": "Numero DT/FT",
            "country": "Nazione",
            "zone": "Zona",
            "trasporto_pu": "Tariffa",
            "trasporto_volume": "Volume",
            "trasporto_qta": "Volume arrotondato",
            "trasporto_tot": "Importo fatturato"
        }
        df_export = df_export.rename(columns=rename_map)

        # 2) prettify del resto
        def prettify_col(col: str) -> str:
            s = str(col).strip()
            if not s:
                return s

            if "_" not in s:
                # Data -> Data, note -> Note, errori -> Errori
                return s[:1].upper() + s[1:]

            s = s.replace("_", " ")
            s = re.sub(r"\s+", " ", s)
            return " ".join(w if w.isupper() else w.capitalize() for w in s.split(" "))

        
        df_export = df_export.rename(columns={c: prettify_col(c) for c in df_export.columns})


        startrow = 3  # (0-based) -> header alla riga 4
        df_export.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Imposta larghezze colonne in modo robusto (per nome colonna),
        # così l'ordine può cambiare senza rompere la formattazione.
        def _set_col_width(col_name: str, width: int):
            if col_name in df_export.columns:
                cidx = int(df_export.columns.get_loc(col_name))
                worksheet.set_column(cidx, cidx, width)

        _set_col_width("PDF", 20)
        _set_col_width("Data", 12)
        _set_col_width("Ns Rif", 14)
        _set_col_width("Cliente", 28)
        _set_col_width("Scarico", 55)
        _set_col_width("Numero DT/FT", 14)
        _set_col_width("Causale", 10)
        _set_col_width("Tariffa", 10)
        _set_col_width("Volume", 8)
        _set_col_width("Volume arrotondato", 16)
        _set_col_width("Importo fatturato", 16)
        _set_col_width("Note", 30)
        _set_col_width("Errori", 55)
        _set_col_width("Errori confronto volume", 55)



        msg_format = workbook.add_format({"bold": True})
        legend_title_format = workbook.add_format({"bold": True})

        error_format = workbook.add_format({"bg_color": "#FFC7CE"})   # rosso chiaro
        note_format  = workbook.add_format({"bg_color": "#FFEB9C"})   # giallo chiaro
        blue_format  = workbook.add_format({"bg_color": "#C6E0FF"})   # azzurro chiaro

        from xlsxwriter.utility import xl_col_to_name
        last_col_idx = df_export.shape[1] - 1
        last_col_letter = xl_col_to_name(last_col_idx)

        worksheet.merge_range(f"A1:{last_col_letter}1", msg, msg_format)

        worksheet.write("A2", "Legenda:", legend_title_format)
        worksheet.write("B2", "Errore (rosso)", error_format)
        worksheet.write("C2", "Nota / riga non riconosciuta (giallo)", note_format)
        worksheet.write("D2", "Volume ≤ 0,3 (blu)", blue_format)

        first_data_row = startrow + 1
        for i in range(len(df_out)):
            row_idx = first_data_row + i

            has_err = bool(df_out.loc[i, "has_error"]) if "has_error" in df_out.columns else False
            has_err_fr = bool(df_out.loc[i, "has_error_volume_francia"]) if "has_error_volume_francia" in df_out.columns else False
            note = df_out.loc[i, "note"] if "note" in df_out.columns else ""
            vol = df_out.loc[i, "trasporto_volume"] if "trasporto_volume" in df_out.columns else None

            err_fr_text = df_out.loc[i, "Errori confronto volume"] if "Errori confronto volume" in df_out.columns else ""
            yellow_fr_only = (not has_err) and isinstance(err_fr_text, str) and err_fr_text.strip() in {
                "non è una spedizione in Francia",
                "DDT non trovato nel file excel",
            }

            # Evidenzia in rosso solo gli errori "bloccanti".
            # Se la riga è "non Francia" o "DDT non trovato", mostriamo il messaggio ma NON coloriamo in rosso (giallo),
            # a meno che ci siano altri errori che hanno precedenza.
            if has_err or (has_err_fr and not yellow_fr_only):
                worksheet.set_row(row_idx, cell_format=error_format)
                continue

            if yellow_fr_only:
                worksheet.set_row(row_idx, cell_format=note_format)
                continue

            if isinstance(note, str) and note.strip():
                worksheet.set_row(row_idx, cell_format=note_format)
                continue

            if pd.notna(vol) and vol <= 0.3:
                worksheet.set_row(row_idx, cell_format=blue_format)

    return msg
