# Estrattore evidenziati e commenti (.docx) – Streamlit

## Avvio
```bash
pip install -r requirements.txt
streamlit run app.py
```
Carica uno o più file .docx. L'app mostra:
- **Evidenziati**: testo evidenziato, colore, contesto (frase) e paragrafo completo
- **Commenti**: ID, autore, data, testo commentato e testo del commento

Filtri disponibili in sidebar: per file, tipo, colore evidenziazione, autore, testo libero, intervallo date.  
Esporta i risultati filtrati in Excel/CSV.
