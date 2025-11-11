import streamlit as st
import pandas as pd
import json
import re
from datetime import datetime, date

st.set_page_config(page_title="IV3 Excel → JSON Converter", layout="centered")
st.title("📘 IV3 Excel → JSON Converter")

# ========== INPUT SECTION (like VBA "Start" sheet) ==========
uploaded_file = st.file_uploader("📂 Upload je Iv3 Excel bestand", type=["xlsx"])

if not uploaded_file:
    st.info("Upload eerst een .xlsx bestand om verder te gaan.")

xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

st.header("🧩 Input parameters")
col1, col2 = st.columns(2)

with col1:
    fin_pakket = st.text_input(
        "Welk pakket gebruikt u voor uw financiële administratie: (bijvoorbeeld Sap of Coda of Key2Financiën)",
        value=""
    )
    export_softw = st.text_input(
        "Export software (BI-software of gewoon het pakket) (bijvoorbeeld Cognos of Coda):",
        value=""
    )
    export_name = st.text_input("Naam exportbestand (zonder extensie):", "Iv3_export_2026")
    

with col2:
    details_openbaar = st.radio("Details openbaar ja of nee?", ["Ja", "Nee"], index=1)
    keer_duizend = st.radio("Keer 1000?", ["Ja", "Nee"], index=1)

run_button = st.button("🚀 Start conversie")


# ========== UTILITIES ==========
def read_sheet(xls, name, header=None):
    try:
        return pd.read_excel(xls, sheet_name=name, header=header, engine="openpyxl")
    except Exception:
        return pd.DataFrame()

def header_to_verslagperiode(h, boekjaar):
    h = str(h).lower()
    m = re.search(r'(\d{4})', h)
    year = int(m.group(1)) if m else boekjaar
    if "rekening" in h or "rek" in h:
        return f"Rek_{year}"
    if "begroting" in h or "beg" in h:
        return f"Beg_{year}"
    return f"Rek_{year}"

# ========== PARSERS ==========
def parse_matrix(xls, sheet_name, keer_duizend_factor):
    df = read_sheet(xls, sheet_name, header=None)
    if df.empty:
        return []
    df = df.fillna("")
    cat_row = None
    for i in range(0, 10):
        row = df.iloc[i].astype(str).tolist()
        if sum(1 for v in row if re.match(r'^\d+(\.\d+)+$', v.strip())) >= 3:
            cat_row = i
            break
    if cat_row is None:
        return []
    cat_map = {j: str(df.iat[cat_row, j]).strip()
               for j in range(df.shape[1])
               if re.match(r'^\d+(\.\d+)+$', str(df.iat[cat_row, j]).strip())}
    taak_col = None
    for j in range(df.shape[1]):
        codes = sum(1 for i in range(cat_row+1, min(cat_row+20, len(df)))
                    if re.match(r'^\d+(\.\d+)*$', str(df.iat[i, j]).strip()))
        if codes >= 3:
            taak_col = j
            break
    if taak_col is None:
        return []
    out = []
    for i in range(cat_row+1, len(df)):
        taakveld = str(df.iat[i, taak_col]).strip()
        if not re.match(r'^\d+(\.\d+)*$', taakveld):
            continue
        for j, cat in cat_map.items():
            val = str(df.iat[i, j]).strip()
            if not val:
                continue
            try:
                bedrag = float(val.replace(",", ".")) * keer_duizend_factor
                if bedrag != 0:
                    out.append({"taakveld": taakveld, "categorie": cat, "bedrag": int(round(bedrag, 0))})
            except:
                continue
    return out

def parse_balansstanden(xls, keer_duizend_factor):
    df = read_sheet(xls, "7.Balansstanden", header=None)
    if df.empty:
        return []
    df = df.fillna("")
    code_col = None
    col_1jan, col_ultimo = None, None
    for i in range(min(10, len(df))):
        row = [str(x).lower() for x in df.iloc[i].tolist()]
        for j, v in enumerate(row):
            if v == "code":
                code_col = j
            if "1 januari" in v:
                col_1jan = j
            if "ultimo" in v:
                col_ultimo = j
        if code_col is not None and (col_1jan or col_ultimo):
            header_row = i
            break
    out = []
    if code_col is None:
        return out
    for i in range(header_row+1, len(df)):
        code = str(df.iat[i, code_col]).strip()
        if not re.match(r'^[A-Z]\d{3,4}$', code):
            continue
        for col, label in [(col_1jan, "1 januari"), (col_ultimo, "ultimo")]:
            if col is None:
                continue
            val = str(df.iat[i, col]).strip()
            if not val:
                continue
            try:
                bedrag = float(val.replace(",", ".")) * keer_duizend_factor
                out.append({"balanscode": code, "standper": label, "bedrag": int(round(bedrag, 0))})
            except:
                continue
    return out

def parse_kengetallen(xls, boekjaar):
    df = read_sheet(xls, "11.Financiële kengetallen", header=None)
    if df.empty:
        return []
    df = df.fillna("")
    header_row = None
    for i in range(0, 30):
        if "verloop van de kengetallen" in str(df.iloc[i].tolist()).lower():
            header_row = i+1
            break
    if header_row is None:
        return []
    headers = [str(x).strip() for x in df.iloc[header_row].tolist()]
    col_headers = {j: h for j, h in enumerate(headers) if h}

    name_to_code = {
        "Netto schuldquote": "fk.1",
        "Netto schuldquote gecorrigeerd": "fk.2",
        "Solvabiliteitsratio": "fk.3",
        "Structurele exploitatieruimte": "fk.4",
        "Grondexploitatie": "fk.5",
        "Belastingcapaciteit": "fk.6"
    }

    out = []
    for i in range(header_row + 1, len(df)):
        label = str(df.iat[i, 0]).strip()
        if not label:
            continue
        code = None
        for key, fk in name_to_code.items():
            if key.lower() in label.lower():
                code = fk
        if not code:
            continue
        for j, header in col_headers.items():
            if j == 0:
                continue
            val = str(df.iat[i, j]).strip()
            # Skip only truly blank cells (not "0" or "0.0")
            if val == "" or val.lower() == "nan" or not val:
                continue
            try:
                numeric = float(val.replace(",", "."))
                value_str = str(int(numeric)) if numeric.is_integer() else str(numeric)
            except:
                value_str = val
            verslagperiode = header_to_verslagperiode(header, boekjaar)
            out.append({"kengetal": code, "verslagperiode": verslagperiode, "waarde": value_str})
    return out

# ========== EXECUTION ==========
if run_button:
    keer_duizend_factor = 1000 if keer_duizend == "Ja" else 1

    info_df = read_sheet(xls, "4.Informatie", header=None).fillna("")
    info_map = {}
    for _, row in info_df.iterrows():
        if len(row) >= 3 and str(row[1]).strip() and str(row[2]).strip():
            info_map[str(row[1]).strip()] = str(row[2]).strip()

    def get_info(keys, default=""):
        for key in keys:
            if key in info_map:
                return info_map[key]
        return default

    boekjaar = int(get_info(["Jaar"], datetime.now().year))
    periode = int(get_info(["Periode"], 5))
    datum_raw = get_info(["Datum"], "")
    try:
        parsed_date = pd.to_datetime(datum_raw).date() if datum_raw else date.today()
    except Exception:
        parsed_date = date.today()
    datum_iso = parsed_date.isoformat() if parsed_date else datetime.now().isoformat()
    
    datum_iso = datum_iso + "T00:00:00+02:00"
    

    metadata = {
        "overheidslaag": get_info(["Overheidslaag"], "Gemeente"),
        "overheidsnummer": get_info(["Nummer"], ""),
        "overheidsnaam": get_info(["Naam"], ""),
        "boekjaar": boekjaar,
        "periode": periode,
        "status": get_info(["Status"], "Realisatie"),
        "datum": datum_iso,
        "details_openbaar": details_openbaar.lower() == "ja",
        "financieel_pakket": fin_pakket,
        "export_software": export_softw
    }

    contact = {
        "naam": get_info(["Naam:", "Contactpersoon"], ""),
        "telefoon": get_info(["Telefoon:", "Telefoon"], ""),
        "email": get_info(["E-mail:", "Email"], "")
    }

    lasten = parse_matrix(xls, "5.Verdelingsmatrix lasten", keer_duizend_factor)
    baten = parse_matrix(xls, "6.Verdelingsmatrix baten", keer_duizend_factor)
    balans_standen = parse_balansstanden(xls, keer_duizend_factor)
    kengetallen = parse_kengetallen(xls, boekjaar)

    data = {
        "lasten": lasten,
        "balans_lasten": [],
        "baten": baten,
        "balans_baten": [],
        "balans_standen": balans_standen,
        "kengetallen": kengetallen,
        "beleidsindicatoren": []
    }

    output = {
        "metadata": metadata,
        "contact": contact,
        "data": data
    }

    json_str = json.dumps(output, indent=2, ensure_ascii=False)
    st.success("✅ Conversie voltooid!")

    st.download_button(
        label="💾 Download JSON",
        data=json_str,
        file_name=f"{export_name}.json",
        mime="application/json"
    )

    st.write("### JSON Preview")
    st.json(output)
