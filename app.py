import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import csv
import re

st.set_page_config(
    page_title="Cettire Statement → CLIARTFATT CSV",
    page_icon="📄",
    layout="centered",
)

# Intestazioni IDENTICHE al CSV di importazione, compresi gli spazi iniziali.
HEADERS = [
    "TIPO_CF", " COD_CLI", " COD_CLI_XMAG", " RAG_SOCIALE", " PARTITA_IVA",
    " COD_FISCALE", " NAZIONE", " INDIRIZZO", " CAP", " CITTA", " PROVINCIA",
    " MAIL", " CELLULARE", " TELEFONO1", " FAX", " PEC", " INSTRAD_FATTELETT",
    " COD_SDI", " ALI_IVA", " COD_ART", " HSCODE", " DESCR_ART",
    " DESCR_ART_ESTESA", " UNITA_DI_MISURA", " COD_CLIENTE_DOCUMENTO",
    " COD_DOC", " SEZIONALE", " VALUTA", " DATA_DOC", " NUM_DOC",
    " PROGRESSIVO_RIGA", " INDICATORE_TIPORIGA", " COD_ART_DOC",
    " DESCRIZIONE_RIGA", " UM", " QUANTITA", " IVA", " PREZZO_1",
    " SCONTO1", " SCONTO2", " MAGGIORAZIONE1", " MAGGIORAZIONE2",
    " NOTE_MOVIMENTO", " COSTI_SPEDIZIONE", " EXSTRASCONTO"
]

# Configurazione dei quattro clienti/fogli Cettire.
SHEET_CONFIG = {
    "Non EU": {
        "cod_cli": "89746",
        "price_column": "Total EUR",
        "divide_by_vat": False,
    },
    "EU": {
        "cod_cli": "89747",
        "price_column": "Gross EUR",
        "divide_by_vat": False,
    },
    "Netherlands": {
        "cod_cli": "89766",
        "price_column": "Total EUR",
        "divide_by_vat": True,
    },
    "Germany": {
        "cod_cli": "89765",
        "price_column": "Total EUR",
        "divide_by_vat": True,
    },
}

# Valori fissi. I campi anagrafici richiesti restano vuoti.
# I campi non citati vengono mantenuti come nel tracciato usato per Poizon.
DEFAULTS = {
    "TIPO_CF": "0",
    " COD_CLI_XMAG": "1040",
    " RAG_SOCIALE": "",
    " PARTITA_IVA": "",
    " COD_FISCALE": "",
    " NAZIONE": "",
    " INDIRIZZO": "",
    " CAP": "",
    " CITTA": "",
    " PROVINCIA": "",
    " MAIL": "",
    " CELLULARE": "",
    " TELEFONO1": "",
    " FAX": "",
    " PEC": "",
    " INSTRAD_FATTELETT": "",
    " COD_SDI": "",
    " ALI_IVA": "0",
    " HSCODE": "",
    " UNITA_DI_MISURA": "PZ",
    " COD_DOC": "FA",
    " SEZIONALE": "V1",
    " VALUTA": "EURO",
    " INDICATORE_TIPORIGA": "0",
    " UM": "PZ",
    " QUANTITA": "1",
    " IVA": "",
    " SCONTO1": "0",
    " SCONTO2": "0",
    " MAGGIORAZIONE1": "0",
    " MAGGIORAZIONE2": "0",
    " COSTI_SPEDIZIONE": "0",
    " EXSTRASCONTO": "0",
}


def clean_excel_value(value):
    """Converte un valore Excel in testo senza aggiungere .0 ai codici numerici."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def format_excel_date(value):
    """Converte la colonna Date nel formato YYYY-MM-DD per COD_ART/COD_ART_DOC."""
    if pd.isna(value):
        return ""

    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d")

    raw = str(value).strip()

    # Se pandas restituisce una data con orario, toglie la parte temporale.
    try:
        parsed = pd.to_datetime(raw, errors="raise")
        return parsed.strftime("%Y-%m-%d")
    except Exception:
        return raw


def parse_decimal(value):
    """Legge correttamente numeri sia con virgola sia con punto decimale."""
    if pd.isna(value) or str(value).strip() == "":
        raise ValueError("Prezzo vuoto")

    raw = str(value).strip().replace("€", "").replace(" ", "")

    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    else:
        raw = raw.replace(",", ".")

    try:
        return Decimal(raw)
    except InvalidOperation:
        raise ValueError(f"Prezzo non valido: {value}")


def format_price(value, divide_by_vat=False):
    """
    Restituisce PREZZO_1 con 2 decimali e virgola.

    Non EU: Total EUR
    EU: Gross EUR
    Netherlands: Total EUR / 1,22
    Germany: Total EUR / 1,22
    """
    price = parse_decimal(value)

    if divide_by_vat:
        price = price / Decimal("1.22")

    price = price.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(price, ".2f").replace(".", ",")


def extract_statement_id(filename):
    """
    Estrae la cifra tra l'ultimo underscore e .xlsx/.xlsm.
    Esempio:
    260908. Statement for The Dope Factory SRL (Dope Factory)_1789021404.xlsx
    -> 1789021404
    """
    name = Path(filename).name
    match = re.search(r"_(\d+)\.(?:xlsx|xlsm)$", name, flags=re.IGNORECASE)

    if not match:
        raise ValueError(
            "Impossibile ricavare NOTE_MOVIMENTO dal nome del file. "
            "Il file deve terminare, ad esempio, con _1789021404.xlsx"
        )

    return match.group(1)


def read_cettire_sheet(excel_bytes, sheet_name, price_column):
    """
    Legge un foglio Cettire cercando automaticamente la riga che contiene
    le intestazioni Date e Reference, senza dipendere dal numero di riga.
    """
    raw = pd.read_excel(
        BytesIO(excel_bytes),
        sheet_name=sheet_name,
        header=None,
        dtype=object,
    )

    header_idx = None

    for idx, row in raw.iterrows():
        first = clean_excel_value(row.iloc[0]) if len(row) > 0 else ""
        second = clean_excel_value(row.iloc[1]) if len(row) > 1 else ""

        if first == "Date" and second == "Reference":
            header_idx = idx
            break

    if header_idx is None:
        raise ValueError(
            f'Nel foglio "{sheet_name}" non trovo la riga con Date e Reference.'
        )

    columns = [clean_excel_value(v) for v in raw.iloc[header_idx].tolist()]
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = columns

    required = ["Date", "Reference", price_column]
    missing = [column for column in required if column not in df.columns]

    if missing:
        raise ValueError(
            f'Nel foglio "{sheet_name}" mancano queste colonne: '
            + ", ".join(missing)
        )

    # Tiene solo le vere righe movimento.
    df = df[
        df["Date"].notna()
        & df["Reference"].notna()
        & df[price_column].notna()
    ].copy()

    # Esclude anche eventuali stringhe vuote.
    df = df[
        df["Date"].apply(lambda x: clean_excel_value(x) != "")
        & df["Reference"].apply(lambda x: clean_excel_value(x) != "")
    ].copy()

    return df


def read_all_sheets(excel_bytes):
    """Legge i quattro fogli richiesti e restituisce dati + conteggi."""
    sheets = {}

    for sheet_name, config in SHEET_CONFIG.items():
        sheets[sheet_name] = read_cettire_sheet(
            excel_bytes,
            sheet_name,
            config["price_column"],
        )

    return sheets


def validate_data_doc(data_doc_text):
    """Valida DATA_DOC nel formato GG/MM/AAAA e la normalizza."""
    raw = data_doc_text.strip()

    if not raw:
        raise ValueError("Inserisci DATA_DOC.")

    try:
        parsed = datetime.strptime(raw, "%d/%m/%Y")
    except ValueError:
        raise ValueError("DATA_DOC deve essere nel formato GG/MM/AAAA, es. 16/09/2026.")

    return parsed.strftime("%d/%m/%Y")


def build_csv(excel_bytes, source_filename, num_doc, data_doc):
    statement_id = extract_statement_id(source_filename)
    data_doc = validate_data_doc(data_doc)
    sheets = read_all_sheets(excel_bytes)

    output_rows = []
    progressivo = 1

    for sheet_name, config in SHEET_CONFIG.items():
        df = sheets[sheet_name]
        cod_cli = config["cod_cli"]
        price_column = config["price_column"]

        for _, src in df.iterrows():
            row = {header: DEFAULTS.get(header, "") for header in HEADERS}

            article_date = format_excel_date(src["Date"])
            reference = clean_excel_value(src["Reference"])

            row[" COD_CLI"] = cod_cli
            row[" COD_CLIENTE_DOCUMENTO"] = cod_cli
            row[" COD_ART"] = article_date
            row[" COD_ART_DOC"] = article_date
            row[" DESCR_ART"] = reference
            row[" DESCR_ART_ESTESA"] = reference
            row[" DESCRIZIONE_RIGA"] = reference
            row[" DATA_DOC"] = data_doc
            row[" NUM_DOC"] = str(num_doc).strip()
            row[" PROGRESSIVO_RIGA"] = str(progressivo)
            row[" PREZZO_1"] = format_price(
                src[price_column],
                divide_by_vat=config["divide_by_vat"],
            )
            row[" NOTE_MOVIMENTO"] = statement_id

            output_rows.append(row)
            progressivo += 1

    if not output_rows:
        raise ValueError("Non ci sono righe da esportare nei quattro fogli Cettire.")

    sio = StringIO(newline="")
    writer = csv.DictWriter(
        sio,
        fieldnames=HEADERS,
        delimiter=";",
        quotechar='"',
        quoting=csv.QUOTE_MINIMAL,
        lineterminator="\r\n",
        extrasaction="ignore",
    )
    writer.writeheader()
    writer.writerows(output_rows)

    # UTF-8 senza BOM, come il CSV di esempio.
    return sio.getvalue().encode("utf-8"), output_rows, statement_id, sheets


st.title("Cettire Statement → CLIARTFATT.CSV")
st.caption(
    "Carica lo statement Cettire, inserisci NUM_DOC e DATA_DOC e genera il CSV pronto per l'importazione."
)

uploaded_file = st.file_uploader(
    "File Excel Cettire",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
)

num_doc = st.text_input(
    "NUM_DOC",
    placeholder="Es. 123",
    help="Numero documento da riportare su tutte le righe del CSV.",
)

data_doc = st.text_input(
    "DATA_DOC",
    placeholder="GG/MM/AAAA - es. 16/09/2026",
    help="Data documento da riportare su tutte le righe del CSV.",
)

if uploaded_file is not None:
    excel_bytes = uploaded_file.getvalue()

    try:
        statement_preview = extract_statement_id(uploaded_file.name)
        sheets_preview = read_all_sheets(excel_bytes)
        total_rows = sum(len(df) for df in sheets_preview.values())

        details = " · ".join(
            f"{sheet}: {len(sheets_preview[sheet])}"
            for sheet in SHEET_CONFIG
        )

        st.success(
            f"File letto correttamente: {total_rows} righe totali. "
            f"{details}. NOTE_MOVIMENTO: {statement_preview}"
        )

    except Exception as e:
        st.error(str(e))
        st.stop()

    if st.button("Genera CSV", type="primary", use_container_width=True):
        if not num_doc.strip():
            st.error("Inserisci NUM_DOC prima di generare il CSV.")
        elif not data_doc.strip():
            st.error("Inserisci DATA_DOC prima di generare il CSV.")
        else:
            try:
                csv_bytes, output_rows, statement_id, sheets = build_csv(
                    excel_bytes=excel_bytes,
                    source_filename=uploaded_file.name,
                    num_doc=num_doc,
                    data_doc=data_doc,
                )

                st.success(
                    f"CSV generato: {len(output_rows)} righe. "
                    f"NOTE_MOVIMENTO: {statement_id}"
                )

                preview_df = pd.DataFrame(output_rows, columns=HEADERS)
                st.dataframe(
                    preview_df,
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    "Scarica CLIARTFATT.csv",
                    data=csv_bytes,
                    file_name="CLIARTFATT.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Errore durante la generazione: {e}")
