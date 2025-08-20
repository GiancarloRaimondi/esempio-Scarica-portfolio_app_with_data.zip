import streamlit as st
import pandas as pd
import numpy as np
import io, os, re
from datetime import datetime
import matplotlib.pyplot as plt
from fpdf import FPDF

st.set_page_config(page_title="Analisi Portafoglio – Preset Universali", layout="wide")
st.title("🧠 Analisi Portafoglio – Preset Universali")
st.caption("Import XLSX 'Categoria prodotto'. Merge automatico con i tuoi file universali Fideuram.")

# ----------------------------- Lettura file Merassi -----------------------------
UPPER_SECTIONS = set(["FONDI","GESTIONI","OBBLIGAZIONI","TITOLI","CASH","LIQUIDITA","LIQUIDITÀ"])

def guess_header_row(df0: pd.DataFrame) -> int:
    for i in range(min(10, len(df0))):
        row = df0.iloc[i].astype(str).str.strip().tolist()
        if ("Nome Prodotto" in row) and ("Codice ISIN" in row):
            return i
    return 1

def load_any_sheet(path: str) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    name = None
    for s in xl.sheet_names:
        if "ategoria" in s.lower():
            name = s
            break
    if name is None:
        name = xl.sheet_names[0]
    df0 = xl.parse(name, header=None)
    hdr = guess_header_row(df0)
    headers = df0.iloc[hdr].tolist()
    base_headers = ["Alert","Nome Prodotto","Nome contratto / dossier","SGR / Emittente","Codice ISIN","Controvalore di fine periodo (€)","Peso sul Totale (%)"]
    for j,h in enumerate(headers):
        if pd.isna(h) and j < len(base_headers):
            headers[j] = base_headers[j]
    if len(headers) < len(base_headers):
        headers = headers + base_headers[len(headers):]
    df = df0.iloc[hdr+1:].copy()
    df.columns = headers[:df.shape[1]]
    return df

def add_section(df: pd.DataFrame) -> pd.DataFrame:
    sec, curr = [], None
    for _, val in df["Nome Prodotto"].items():
        if isinstance(val, str) and val.strip().isupper() and len(val.strip())<=30:
            txt = val.strip()
            if txt in UPPER_SECTIONS or len(txt.split())<=3:
                curr = txt
        sec.append(curr)
    df["Sezione"] = sec
    return df

def classify_tipo(row) -> str:
    sec = str(row.get("Sezione","") or "").upper()
    name = str(row.get("Nome Prodotto","") or "").upper()
    issuer = str(row.get("SGR / Emittente","") or "").upper()
    isin = str(row.get("Codice ISIN","") or "")
    if sec == "FONDI" or "FUND" in name or "FONDITALIA" in name:
        return "Fondo"
    if any(k in name for k in ["BTP","BOND","NOTE","OBBLIG","TLX","XS","EUROBOND"]) or issuer in ["REPUBLIC OF ITALY","MIN FIN"]:
        return "Obbligazione"
    if re.match(r"^[A-Z]{2}", isin) and any(k in name for k in [" INC"," CORP"," PLC"," SPA"]) and len(isin)==12:
        return "Azione"
    if sec in ["GESTIONI","GESTIONE"] or "GP" in name or "GESTIONE" in name:
        return "Gestione"
    return "Titolo"

def normalize(core: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame()
    out["ISIN"] = core["Codice ISIN"].astype(str).str.strip()
    out["Strumento"] = core["Nome Prodotto"].astype(str).str.strip()
    out["Tipo"] = core.apply(classify_tipo, axis=1)
    out["Controvalore"] = pd.to_numeric(core.get("Controvalore di fine periodo (€)"), errors="coerce").fillna(0.0)
    peso_raw = pd.to_numeric(core.get("Peso sul Totale (%)"), errors="coerce")
    if peso_raw.notna().sum()>0:
        s = peso_raw.fillna(0).sum()
        peso = np.where((s>0) & (abs(s-100)>1), peso_raw*100.0/s, peso_raw)
    else:
        tot = out["Controvalore"].sum()
        peso = np.where(tot>0, out["Controvalore"]/tot*100.0, 0.0)
    out["Peso_%"] = pd.to_numeric(peso, errors="coerce").fillna(0.0).round(6)
    for c, v in [("Categoria",""),("Valuta",""),("Area",""),("Settore",""),("Rating",0),("Quartile",0),("Costo_annuo_%",0.0),("Margine_bps",0.0),("Note","")]:
        out[c] = v
    return out

# ----------------------------- Merge file universali -----------------------------
PRESET_FILES = [
    "Fondi asset.xlsx",
    "Fondi Valuta.xlsx",
    "Fondi.xlsx",
    "Gestione e fogli valuta.xlsx",
    "Gestioni e Fogli asset allocation.xlsx",
    "Gestioni e Fogli.xlsx",
    "Povvigioni Fondi.xlsx",
    "Provvigioni Fondi.xlsx",
    "Provvigioni Gestioni e Fogli.xlsx",
]
ALIASES = {
    "ISIN": ["ISIN","Codice ISIN","isin","Isin"],
    "Categoria": ["Categoria","Categoria Prodotto","Sottocategoria","Categoria Morningstar","Classe","Style","Asset Class"],
    "Valuta": ["Valuta","Currency","CCY","Divisa"],
    "Area": ["Area","Regione","Paese","Country","Area Geografica"],
    "Settore": ["Settore","Sector","Industry","Industria"],
    "Rating": ["Rating","Rating MS","Stelle","Stars","Morningstar Rating"],
    "Quartile": ["Quartile","Quartili","Quartile MS","Quartile Morningstar"],
    "Costo_annuo_%": ["Costo_annuo_%","TER_%","OCF_%","Commissioni_annue_%","Expense Ratio %"],
    "Margine_bps": ["Margine_bps","Provvigioni bps","Retro_bps","bps","Retrocession_bps","Retrocessioni_bps"],
}

def normalize_columns(u: pd.DataFrame) -> pd.DataFrame:
    cols = {}
    for tgt, aliases in ALIASES.items():
        for a in aliases:
            if a in u.columns:
                cols[tgt] = a; break
    if "ISIN" not in cols:
        for c in u.columns:
            if u[c].astype(str).str.match(r"^[A-Z]{2}[A-Z0-9]{10}$", na=False).any():
                cols["ISIN"] = c; break
    keep = [k for k in ["ISIN","Categoria","Valuta","Area","Settore","Rating","Quartile","Costo_annuo_%","Margine_bps"] if k in cols]
    out = pd.DataFrame()
    for k in keep: out[k] = u[cols[k]]
    if "ISIN" in out: out["ISIN"] = out["ISIN"].astype(str).str.strip()
    for c in ["Rating","Quartile"]: 
        if c in out: out[c] = pd.to_numeric(out[c], errors="coerce")
    for c in ["Costo_annuo_%","Margine_bps"]:
        if c in out: out[c] = pd.to_numeric(out[c], errors="coerce")
    return out

def merge_universali(df: pd.DataFrame, base_dir: str):
    uni_dir = os.path.join(base_dir, "universali")
    if not os.path.isdir(uni_dir): return df, []
    used = []
    for fname in PRESET_FILES:
        path = os.path.join(uni_dir, fname)
        if not os.path.exists(path): continue
        try:
            u = pd.read_excel(path)
            u_norm = normalize_columns(u)
            if "ISIN" in u_norm.columns and len(u_norm.columns)>=2:
                df = df.merge(u_norm, on="ISIN", how="left", suffixes=("","_u"))
                for c in ["Categoria","Valuta","Area","Settore","Rating","Quartile","Costo_annuo_%","Margine_bps"]:
                    if c+"_u" in df.columns:
                        df[c] = df[c].combine_first(df[c+"_u"])
                        df.drop(columns=[c+"_u"], inplace=True)
                used.append(fname)
        except Exception:
            continue
    df["Rating"] = pd.to_numeric(df["Rating"], errors="coerce").fillna(0).astype(int)
    df["Quartile"] = pd.to_numeric(df["Quartile"], errors="coerce").fillna(0).astype(int)
    df["Costo_annuo_%"] = pd.to_numeric(df["Costo_annuo_%"], errors="coerce").fillna(0.0)
    df["Margine_bps"] = pd.to_numeric(df["Margine_bps"], errors="coerce").fillna(0.0)
    return df, used

# ----------------------------- Export PDF con FPDF2 -----------------------------
def export_pdf_fpdf(metrics: dict, flags: dict) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", size=14)
    pdf.cell(0, 8, "Analisi Portafoglio – Sintesi", ln=1)
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, f"Data: {datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=1)
    pdf.cell(0, 6, f"Totale controvalore: € {metrics['totale']:,.2f}", ln=1)
    pdf.cell(0, 6, f"Margine annuo stimato: € {metrics['margine_eur']:,.2f}", ln=1)
    pdf.cell(0, 6, f"Costi annui stimati: € {metrics['costi_eur']:,.2f}", ln=1)
    pdf.ln(3)
    pdf.set_font("Helvetica", size=12)
    pdf.cell(0, 7, "Bandierine di revisione", ln=1)
    pdf.set_font("Helvetica", size=10)
    for k, dff in flags.items():
        pdf.multi_cell(0, 6, f"- {k}: {len(dff)} strumenti")
    out = io.BytesIO()
    pdf.output(out)
    return out.getvalue()

# ----------------------------- UI -----------------------------
with st.sidebar:
    st.header("Input")
    src = st.radio("Sorgente", ["Carica XLSX","Usa sample opzionale"])
    uploaded = None
    if src == "Carica XLSX":
        uploaded = st.file_uploader("Carica file Excel (report Categoria prodotto)", type=["xlsx","xls","xlsm"])
    else:
        st.info("Puoi mettere un sample in sample_data/Portafoglio.xlsx se vuoi.")
    uni_dir = "universali"

if src == "Usa sample opzionale":
    path = "sample_data/Portafoglio.xlsx"
    if not os.path.exists(path): st.stop()
else:
    if uploaded is None:
        st.info("Carica il file XLSX per procedere.")
        st.stop()
    path = uploaded

raw = load_any_sheet(path)
raw = add_section(raw)
core_cols = [c for c in ["Nome Prodotto","Nome contratto / dossier","SGR / Emittente","Codice ISIN","Controvalore di fine periodo (€)","Peso sul Totale (%)","Sezione"] if c in raw.columns]
core = raw[core_cols].copy()
core = core[core["Codice ISIN"].notna()]
if core.empty:
    st.error("Nessun ISIN rilevato. Controlla il file.")
    st.stop()

norm = normalize(core)
norm, used_files = merge_universali(norm, ".")

df = norm.copy()
df["Peso_%"] = pd.to_numeric(df["Peso_%"], errors="coerce").fillna(0.0)
df["Controvalore"] = pd.to_numeric(df["Controvalore"], errors="coerce").fillna(0.0)

totale = df["Controvalore"].sum()
margine_eur = (df["Controvalore"] * (df["Margine_bps"]/10000.0)).sum()
costi_eur = (df["Controvalore"] * (df["Costo_annuo_%"]/100.0)).sum()
peso_ok = np.isclose(df["Peso_%"].sum(), 100.0, atol=1.0)

flags = {
    "Rating basso (<=2) in fondi": df[(df["Tipo"]=="Fondo") & (df["Rating"]<=2)],
    "Quartile 4 in fondi": df[(df["Tipo"]=="Fondo") & (df["Quartile"]==4)],
    "Costo alto (>=1.80%)": df[(df["Costo_annuo_%"]>=1.80)],
    "Margine nullo (<5 bps)": df[(df["Margine_bps"]<5)],
}
flags_text = [f"{k}: {len(v)} strumenti" for k,v in flags.items()]
metrics = {"totale": totale,"margine_eur": margine_eur,"costi_eur": costi_eur,"peso_ok": peso_ok,"flags_text": flags_text}

c1, c2, c3, c4 = st.columns(4)
c1.metric("Totale", f"€ {totale:,.0f}")
c2.metric("Margine annuo", f"€ {margine_eur:,.0f}")
c3.metric("Costi annui", f"€ {costi_eur:,.0f}")
c4.metric("Pesi ≈ 100%", "OK" if peso_ok else "No")

st.subheader("📄 Posizioni normalizzate + Universali")
st.dataframe(df, use_container_width=True)

st.subheader("🔗 File universali utilizzati")
if used_files:
    for f in used_files: st.write(f"- {f}")
else:
    st.write("Nessun file universale trovato in `universali/`.")

st.subheader("📊 Allocazioni")
by_tipo = df.groupby("Tipo")["Controvalore"].sum().sort_values(ascending=False)
by_area = df.groupby("Area")["Controvalore"].sum().sort_values(ascending=False)
by_settore = df.groupby("Settore")["Controvalore"].sum().sort_values(ascending=False)

g1 = plt.figure(); by_tipo.plot(kind="bar"); plt.title("Per Tipo"); plt.ylabel("Controvalore"); st.pyplot(g1)
g2 = plt.figure(); by_area.plot(kind="bar"); plt.title("Per Area"); plt.ylabel("Controvalore"); st.pyplot(g2)
g3 = plt.figure(); by_settore.head(12).plot(kind="bar"); plt.title("Top Settori"); plt.ylabel("Controvalore"); st.pyplot(g3)

st.subheader("⬇️ Export")
if st.button("Esporta Excel"):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Posizioni")
        by_tipo.to_frame("Controvalore").to_excel(writer, sheet_name="Alloc_Tipo")
        by_area.to_frame("Controvalore").to_excel(writer, sheet_name="Alloc_Area")
        by_settore.to_frame("Controvalore").to_excel(writer, sheet_name="Alloc_Settore")
    st.download_button("Download Analisi.xlsx", data=bio.getvalue(),
                       file_name="Analisi_Preset.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if st.button("Esporta PDF"):
    pdf_bytes = export_pdf_fpdf(metrics, flags)
    st.download_button("Download Report.pdf", data=pdf_bytes,
                       file_name="Report_Preset.pdf", mime="application/pdf")

st.caption("Inserisci in `universali/` i tuoi Excel reali (Fondi/Foglio/Provvigioni). Merge per ISIN con alias colonne.")
