# 📊 Ads Analyzer

Web app che analizza uno **screenshot** o un **report** di Meta Ads e ti dice cosa funziona, cosa no, **dove intervenire e come** — in ordine di priorità.

Fai l'upload di uno screenshot del Gestione Inserzioni (oppure di un export CSV/Excel, o incolli i numeri a mano), aggiungi il contesto (città, obiettivo, budget, da quanto è attiva, settore, ticket medio) e ricevi un'analisi operativa generata da Claude (con lettura dell'immagine via vision).

## Come si avvia

Serve [Node.js](https://nodejs.org) 18+ (già presente su questo PC) e una chiave API Anthropic.

```powershell
# 1. Installa le dipendenze
npm install

# 2. Crea il file .env con la tua chiave
copy .env.example .env
#    poi apri .env e incolla la chiave: ANTHROPIC_API_KEY=sk-ant-...

# 3. Avvia
npm start
```

Apri il browser su **http://localhost:3000**.

## Come funziona

- **Frontend** (`public/index.html`): form con i dati di contesto + upload/drag&drop dello screenshot o del report.
- **Backend** (`server.js`): riceve i dati, manda screenshot (vision) o report a Claude con un prompt che incorpora i benchmark di Meta Ads e la logica del funnel, e restituisce un'analisi in Markdown.
- **Modello**: `claude-opus-4-8` con thinking adaptive.

## Input supportati

| Tipo | Come |
|------|------|
| Screenshot | PNG / JPG / WebP del Gestione Inserzioni (letto via vision) |
| Report file | CSV / TSV / TXT / JSON esportato da Meta |
| Testo | Numeri incollati a mano nel campo "report" |

> I numeri esatti (CSV) danno un'analisi più affidabile dello screenshot.

## Cosa restituisce

Sintesi → Cosa funziona 🟢 → Cosa non funziona 🔴 → Da tenere d'occhio 🟡 → **Cosa fare in ordine di priorità** (azioni concrete su creatività / targeting / budget / offerta / landing) → eventuali dati mancanti.

## Note

- La chiave API non viene mai inviata al browser: resta lato server.
- Pensato come prototipo indipendente, ma facilmente integrabile in seguito (es. dentro Brods, con lo stesso stack Supabase + Stripe e i crediti/abbonamento).
