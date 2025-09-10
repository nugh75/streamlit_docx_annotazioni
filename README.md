# Estrattore evidenziati e commenti (.docx)

Sono fornite due interfacce:

- Streamlit (rapida da usare, completa di filtri ed export)
- Web app moderna (Vite + React) con backend FastAPI

## Requisiti

- Python 3.10+
- Node 18+

## Setup ambiente Python

```bash
pip install -r requirements.txt
```

## Avvio Streamlit

```bash
streamlit run app.py
```

Carica uno o più file .docx. L'app mostra:

- Evidenziati: testo unito per colore, colore, contesto (frase) e paragrafo completo
- Commenti (esplosi per più coppie): ID, autore, data, testo commentato e testo del commento
- Scrematura: vista semplificata per export, con mappa Colore → Categoria, soggetto intervistato e tipo documento

Filtri disponibili in sidebar: per file, tipo, colore evidenziazione, autore, codice/etichetta, tipo documento, testo libero, intervallo date.

Export: Excel globale; CSV/XLSX separati per Interviste e Focus group dalla Scrematura.

## Web App (Vite + React) + Backend FastAPI

### Backend

```bash
# da root progetto
source .venv/bin/activate
uvicorn backend.main:app --host 127.0.0.1 --port 8000 --reload
```

API docs: <http://127.0.0.1:8000/docs>

### Frontend

```bash
cd frontend
npm install
npm run dev
```

Apri <http://localhost:5173> e carica un file .docx. Puoi:

- Vedere il documento a sinistra con testo e highlights inline
- Vedere elenco evidenziati a destra
- Mappare colore → categoria e vedere semplici statistiche per categoria (conteggi)

### Note di design e migliorie

- Merge dei run contigui con stesso colore per testo evidenziato più pulito
- Mappatura colore → macro-categoria modificabile
- Statistiche basilari per categoria
- API separate per future integrazioni (DB, auth, salvataggi)
