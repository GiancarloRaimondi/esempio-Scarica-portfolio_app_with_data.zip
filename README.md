# Portfolio App con dati di esempio

App Streamlit per analisi portafogli base con export PDF/Excel.

## Avvio locale
```bash
pip install -r requirements.txt
streamlit run app.py
```

## File inclusi
- `app.py`: app Streamlit pronta all'uso
- `sample_data/sample_master.xlsx`: dati di esempio
- `requirements.txt`: dipendenze minime
- `.streamlit/config.toml`: layout pulito

## Schema dati atteso (XLSX)
Colonne minime:
- Tipo (Fondo/Azione/Obbligazione/Cash/Gestione)
- ISIN
- Strumento
- Categoria
- Valuta
- Area
- Settore
- Peso_%
- Controvalore
- Rating (1-5)
- Quartile (1-4, Morningstar-like)
- Costo_annuo_%
- Margine_bps (retro)
- Note (opzionale)

## Deployment su Streamlit Cloud
1. Carica questo pacchetto su un repo GitHub
2. Crea una nuova App su Streamlit Cloud puntando a `app.py`
3. Imposta Python 3.11+
