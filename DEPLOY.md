# Mettere online Meta Ads Clinic (guida passo-passo)

Obiettivo: avere **un link pubblico** che chiunque apre e usa, e poterlo anche
incollare nella tua landing **GoHighLevel** sul tuo dominio.

Ricorda lo schema:
- **Vetrina** = la pagina (link Render e/o la tua landing GoHighLevel).
- **Motore** = il server (`server.js`) che custodisce la chiave API. Va online su Render.
- L'utente finale non vede mai il motore: apre solo la vetrina.

---

## PARTE 1 — La chiave API (una volta sola)

1. Vai su **https://console.anthropic.com** → crea un account.
2. Aggiungi un po' di credito (Billing).
3. Vai in **API Keys** → **Create Key** → copia la chiave (`sk-ant-...`).
   Tienila da parte: la incolli su Render al passo 2.4. **Non metterla mai nel widget.**

---

## PARTE 2 — Il motore online su Render (gratis, ~10 min)

Render ha bisogno del codice su **GitHub**. Se non usi git, va benissimo caricare i file dal browser.

### 2.1 Metti il codice su GitHub (dal browser)
1. Crea un account su **https://github.com**.
2. **New repository** → nome es. `meta-ads-clinic` → **Create**.
3. Nella pagina del repo: **Add file → Upload files** e trascina TUTTI questi file/cartelle
   (NON caricare `node_modules` né `.env`):
   - `server.js`
   - `package.json`
   - `package-lock.json`
   - la cartella `public/` (con dentro `index.html`)
   - `.gitignore`
4. **Commit changes**.

### 2.2 Crea il servizio su Render
1. Vai su **https://render.com** → registrati (puoi usare "Sign in with GitHub").
2. **New + → Web Service** → collega il repo `meta-ads-clinic`.
3. Render rileva Node da solo. Controlla:
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Instance Type:** Free
4. **Environment → Add Environment Variable:**
   - Key: `ANTHROPIC_API_KEY` → Value: la tua chiave `sk-ant-...`
   - (facoltativo) Key: `ALLOWED_ORIGIN` → Value: il tuo dominio GHL, es. `https://tuodominio.com`
5. **Create Web Service** e aspetta il deploy (qualche minuto).
6. Otterrai un URL tipo **`https://meta-ads-clinic.onrender.com`**.

✅ Aprendo quell'URL vedi già il widget completo e funzionante: **è già un link condivisibile.**

> Nota: nel piano Free, dopo un po' di inattività il servizio "si addormenta" e la
> prima richiesta dopo la pausa è più lenta (~30s). Per uso continuo si passa a un piano a pagamento.

---

## PARTE 3 — Dentro GoHighLevel (sul tuo dominio)

1. Apri il file **`meta-ads-clinic-GHL-embed.html`**.
2. Trova in fondo la riga:
   ```js
   var CONFIG = { backendUrl: "https://IL-TUO-BACKEND.com" };
   ```
   e sostituisci con l'URL Render del passo 2.6, es.:
   ```js
   var CONFIG = { backendUrl: "https://meta-ads-clinic.onrender.com" };
   ```
3. In GoHighLevel: apri la tua **landing/funnel** → aggiungi un elemento
   **Custom HTML / Custom Code** → incolla **tutto** il contenuto del file.
4. Salva e pubblica. La pagina è sul **tuo dominio** GHL.

✅ Ora l'utente apre il tuo dominio GHL e usa lo strumento. Identico per lui.

---

## Riepilogo dei link che ottieni
- `https://...onrender.com` → link autonomo, pronto subito.
- `https://tuodominio.com/...` → la stessa cosa, dentro il tuo GoHighLevel/brand.

## Da valutare più avanti
- **Protezione**: il link è pubblico e ogni analisi consuma il tuo credito API.
  Possiamo aggiungere una password, un limite di usi, o i crediti (come in Brods).
- **xlsx**: ora legge CSV; volendo aggiungiamo il supporto Excel binario.
