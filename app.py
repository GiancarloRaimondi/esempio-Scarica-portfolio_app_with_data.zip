import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import matplotlib.pyplot as plt

st.set_page_config(page_title="Analisi Portafoglio – Raimondi", layout="wide")

st.title("🧠 Analisi Portafoglio Clienti – Versione Base")
st.caption("Upload XLSX o usa i dati di esempio. Output: sintesi, allocazioni, margine, bandierine di qualità, export PDF/Excel.")

# -----------------------------
# Utility
# -----------------------------
REQUIRED_COLS = ["Tipo","ISIN","Strumento","Categoria","Valuta","Area","Settore","Peso_%","Controvalore","Rating","Quartile","Costo_annuo_%","Margine_bps"]

def load_xlsx(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    return df

def validate_schema(df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    return missing

def compute_metrics(df: pd.DataFrame):
    df = df.copy()
    df["Peso_%"] = pd.to_numeric(df["Peso_%"], errors="coerce").fillna(0.0)
    df["Controvalore"] = pd.to_numeric(df["Controvalore"], errors="coerce").fillna(0.0)
    df["Rating"] = pd.to_numeric(df["Rating"], errors="coerce").fillna(0).astype(int)
    df["Quartile"] = pd.to_numeric(df["Quartile"], errors="coerce").fillna(0).astype(int)
    df["Costo_annuo_%"] = pd.to_numeric(df["Costo_annuo_%"], errors="coerce").fillna(0.0)
    df["Margine_bps"] = pd.to_numeric(df["Margine_bps"], errors="coerce").fillna(0.0)

    totale = df["Controvalore"].sum()
    peso_ok = np.isclose(df["Peso_%"].sum(), 100.0, atol=1.0)

    # Margine annuo in EUR: controvalore * bps/10000
    margine_eur = (df["Controvalore"] * (df["Margine_bps"] / 10000.0)).sum()

    # Costi annuali per fondi/gestioni
    costi_eur = (df["Controvalore"] * (df["Costo_annuo_%"] / 100.0)).sum()

    # Allocazioni
    by_tipo = df.groupby("Tipo")["Controvalore"].sum().sort_values(ascending=False)
    by_area = df.groupby("Area")["Controvalore"].sum().sort_values(ascending=False)
    by_categoria = df.groupby("Categoria")["Controvalore"].sum().sort_values(ascending=False)

    # Flag qualità semplici
    flags = {}
    flags["Fondi_quartile4"] = df[(df["Tipo"]=="Fondo") & (df["Quartile"]==4)]
    flags["Fondi_rating_basso"] = df[(df["Tipo"]=="Fondo") & (df["Rating"]<=2)]
    flags["Fondi_costo_alto"] = df[(df["Tipo"]=="Fondo") & (df["Costo_annuo_%"]>=1.80)]
    flags["Gest_multi_da_rivedere"] = df[(df["Tipo"]=="Gestione") & (df["Rating"]<=3)]

    return {
        "df": df,
        "totale": totale,
        "peso_ok": peso_ok,
        "margine_eur": margine_eur,
        "costi_eur": costi_eur,
        "by_tipo": by_tipo,
        "by_area": by_area,
        "by_categoria": by_categoria,
        "flags": flags
    }

def fig_bar(series: pd.Series, title: str):
    fig, ax = plt.subplots()
    series.plot(kind="bar", ax=ax)
    ax.set_title(title)
    ax.set_xlabel("")
    ax.set_ylabel("Controvalore")
    fig.tight_layout()
    return fig

def export_excel(metrics: dict) -> bytes:
    df = metrics["df"]
    with pd.ExcelWriter(io.BytesIO(), engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Posizioni")
        metrics["by_tipo"].to_frame("Controvalore").to_excel(writer, sheet_name="Alloc_Tipo")
        metrics["by_area"].to_frame("Controvalore").to_excel(writer, sheet_name="Alloc_Area")
        metrics["by_categoria"].to_frame("Controvalore").to_excel(writer, sheet_name="Alloc_Categoria")
        wb = writer.book
        writer._save()
        data = writer._writer.getvalue()
    return data

def export_pdf(metrics: dict) -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4

    df = metrics["df"]
    totale = metrics["totale"]
    margine = metrics["margine_eur"]
    costi = metrics["costi_eur"]

    y = h - 2*cm
    c.setFont("Helvetica-Bold", 14)
    c.drawString(2*cm, y, "Analisi Portafoglio – Sintesi")
    y -= 0.8*cm
    c.setFont("Helvetica", 10)
    c.drawString(2*cm, y, f"Data: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    y -= 0.6*cm
    c.drawString(2*cm, y, f"Totale controvalore: € {totale:,.2f}")
    y -= 0.6*cm
    c.drawString(2*cm, y, f"Margine annuo stimato: € {margine:,.2f}")
    y -= 0.6*cm
    c.drawString(2*cm, y, f"Costi annui stimati: € {costi:,.2f}")
    y -= 1.0*cm

    c.setFont("Helvetica-Bold", 12)
    c.drawString(2*cm, y, "Bandierine di revisione")
    y -= 0.6*cm
    c.setFont("Helvetica", 10)
    flags = metrics["flags"]
    for key, df_flag in flags.items():
        c.drawString(2.2*cm, y, f"- {key}: {len(df_flag)} strumenti")
        y -= 0.5*cm
        if y < 2*cm:
            c.showPage(); y = h - 2*cm; c.setFont("Helvetica", 10)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.read()

# -----------------------------
# UI
# -----------------------------
with st.sidebar:
    st.header("Input")
    mode = st.radio("Sorgente dati", ["Usa dati di esempio","Carica XLSX"])
    uploaded = None
    if mode == "Carica XLSX":
        uploaded = st.file_uploader("Carica file Excel", type=["xlsx"])

if mode == "Usa dati di esempio":
    df = pd.read_excel("sample_data/sample_master.xlsx")
else:
    if uploaded is None:
        st.info("Carica il file XLSX per procedere.")
        st.stop()
    df = load_xlsx(uploaded)

missing = validate_schema(df)
if missing:
    st.error(f"Colonne mancanti: {missing}")
    st.stop()

metrics = compute_metrics(df)

# KPI
c1, c2, c3, c4 = st.columns(4)
c1.metric("Totale", f"€ {metrics['totale']:,.0f}")
c2.metric("Margine annuo", f"€ {metrics['margine_eur']:,.0f}")
c3.metric("Costi annui", f"€ {metrics['costi_eur']:,.0f}")
c4.metric("Pesi ≈ 100%", "OK" if metrics["peso_ok"] else "No")

# Tabelle
st.subheader("📄 Posizioni")
st.dataframe(metrics["df"], use_container_width=True)

# Grafici
st.subheader("📊 Allocazioni")
g1 = fig_bar(metrics["by_tipo"], "Allocazione per Tipo")
st.pyplot(g1)
g2 = fig_bar(metrics["by_area"], "Allocazione per Area")
st.pyplot(g2)
g3 = fig_bar(metrics["by_categoria"].head(10), "Top Categorie")
st.pyplot(g3)

# Bandierine qualità
st.subheader("🚩 Bandierine")
cols = st.columns(4)
for i, (k, dff) in enumerate(metrics["flags"].items()):
    with cols[i % 4]:
        st.write(f"**{k}**: {len(dff)}")
        if len(dff) > 0:
            st.dataframe(dff[["Tipo","ISIN","Strumento","Categoria","Costo_annuo_%","Rating","Quartile","Margine_bps","Controvalore"]], use_container_width=True)

# Export
st.subheader("⬇️ Export")
colA, colB = st.columns(2)
with colA:
    if st.button("Esporta Excel"):
        xls_bytes = export_excel(metrics)
        st.download_button("Download Analisi.xlsx", data=xls_bytes, file_name="Analisi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with colB:
    if st.button("Esporta PDF"):
        pdf_bytes = export_pdf(metrics)
        st.download_button("Download Report.pdf", data=pdf_bytes, file_name="Report.pdf", mime="application/pdf")

st.caption("Versione base offline. Per integrazione rating/quotazioni reali serve collegamento a fonti esterne.")
