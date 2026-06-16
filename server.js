import express from "express";
import multer from "multer";
import Anthropic from "@anthropic-ai/sdk";
import ExcelJS from "exceljs";
import "dotenv/config";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));

const app = express();
const upload = multer({ limits: { fileSize: 15 * 1024 * 1024 } }); // max 15 MB
const client = new Anthropic(); // legge ANTHROPIC_API_KEY dall'ambiente

// CORS: il widget gira su un dominio GoHighLevel diverso da questo backend.
// In produzione, sostituisci "*" con il/i tuo/i dominio/i GHL per più sicurezza.
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "*";
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
  res.header("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.use(express.static(join(__dirname, "public")));

// Formati immagine accettati dall'API vision
const IMAGE_TYPES = {
  "image/png": "image/png",
  "image/jpeg": "image/jpeg",
  "image/jpg": "image/jpeg",
  "image/webp": "image/webp",
  "image/gif": "image/gif",
};

// Estensioni testuali trattate come "report" da incollare nel prompt
const TEXT_EXT = [".csv", ".tsv", ".txt", ".json"];

// Converte un file .xlsx (buffer) in testo tipo CSV, foglio per foglio.
async function xlsxToText(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const cell = (v) => {
    if (v == null) return "";
    if (typeof v === "object") {
      if (v.text != null) return v.text; // rich text / hyperlink
      if (v.result != null) return v.result; // formula → risultato
      if (v.richText) return v.richText.map((r) => r.text).join("");
      return "";
    }
    return v;
  };
  let out = "";
  wb.eachSheet((ws) => {
    out += `# Foglio: ${ws.name}\n`;
    ws.eachRow((row) => {
      const vals = (row.values || []).slice(1).map((v) => String(cell(v)).trim());
      // Separatore "|" per non confondersi con i decimali italiani (es. 120,50)
      out += vals.join(" | ") + "\n";
    });
    out += "\n";
  });
  return out;
}

const SYSTEM_PROMPT = `Sei il "medico delle ads": un media buyer senior specializzato in Meta Ads (Facebook/Instagram) che fa la DIAGNOSI di una campagna come un clinico. Guardi i dati, capisci lo stato di salute della campagna, e prescrivi la "terapia": dici in modo chiaro DOVE intervenire, COSA fare e PERCHÉ. Tono: professionale, diretto, empatico ma onesto — come un bravo medico che spiega al paziente.

## Come ragioni
1. Parti SEMPRE dall'obiettivo della campagna e dall'economia del cliente (ticket medio, margine). Una metrica "buona" o "cattiva" dipende dal contesto: un CPA di 30€ è ottimo se il prodotto vale 300€, pessimo se vale 25€.
2. Usa i benchmark Meta Ads come riferimento, ma adattali al settore e alla geografia. Benchmark indicativi (variano per settore/paese):
   - CTR (link click): sotto 0,8% = debole, 0,8–1,5% = nella media, sopra 1,5% = buono. Per campagne di notorietà conta più il CTR "all", per conversioni il CTR link.
   - CPM: dipende molto da paese e targeting; in Italia spesso 4–12€. Un CPM molto alto con poco budget = aste affollate o pubblico troppo ristretto.
   - Frequenza: sopra ~3 in pochi giorni = creatività che si "brucia" / pubblico troppo piccolo → stanchezza creativa.
   - Tasso di conversione landing: sotto l'1–2% spesso indica un problema di landing page o di promessa, non di campagna.
   - Hook rate / retention video: se il 3-sec view rate è basso, il problema è il primo frame/hook.
3. Distingui SEMPRE dove sta il collo di bottiglia nel funnel:
   - Poche impression/copertura → budget, bid, o pubblico troppo ristretto.
   - Tante impression ma CTR basso → problema di CREATIVITÀ (visual + hook + copy) o di targeting non in target.
   - Buon CTR ma poche conversioni → problema di LANDING PAGE, offerta, prezzo, o coerenza annuncio↔pagina.
   - CPA alto ma vendite ok → questione di margine/scala, non per forza un errore.
4. Sii consapevole del LIVELLO di analisi e dillo all'utente quando serve:
   - Livello INSERZIONE (ad) = la "sezione creatività": qui giudichi visual, hook, copy, formato. Se il problema è il CTR o l'hook, l'azione è qui.
   - Livello GRUPPO DI INSERZIONI (ad set): qui giudichi pubblico/targeting, posizionamenti, budget, ottimizzazione. Se il problema è copertura, CPM, o consegna, l'azione è qui.
   - Livello CAMPAGNA: obiettivo e strategia di budget (CBO/ABO).
   Indica sempre, per ogni azione, A QUALE LIVELLO va fatta la modifica.
5. Sui TESTI/COPY: valutali (hook, primi 3 righe, CTA) SOLO se ti vengono forniti — nel report con i testi, o in uno screenshot dell'anteprima dell'annuncio. Se hai solo lo screenshot della tabella metriche (solo numeri), NON inventare il copy: di' che per valutare creatività e copy serve allegare i testi o l'anteprima.
6. Se i dati sono insufficienti o letti da uno screenshot a bassa risoluzione, dichiaralo e di' quali dati servirebbero. Non inventare numeri.
7. Tieni conto del tempo di attività: una campagna con pochi giorni o sotto ~50 conversioni è ancora in "fase di apprendimento" — sconsiglia modifiche drastiche premature.
8. CAMPI DI CONTESTO FACOLTATIVI: l'utente può non averli compilati. NON bloccarti e NON rifiutarti di analizzare. Procedi così:
   a) Prova prima a DEDURRE i dati mancanti dal report/screenshot (es. spesa, risultati, periodo, a volte l'area geografica o l'obiettivo).
   b) Fai comunque la migliore analisi possibile con ciò che hai.
   c) Solo alla fine, nella sezione "Cosa mi serve da te", elenca in modo specifico le informazioni che mancano e che non hai potuto dedurre (es. città, ticket medio, obiettivo), spiega PERCHÉ servono e COME cambierebbero la diagnosi, e invita l'utente ad aggiungerle. Se manca un dato cruciale per un giudizio (es. senza ticket medio non puoi dire se il CPA è buono), dillo chiaramente ma dai comunque il resto dell'analisi.

## Formato della risposta (in italiano, Markdown)
Rispondi in modo conciso e azionabile, in questa struttura:

### 🩺 Quadro generale
2-3 righe da "medico": stato di salute complessivo della campagna e qual è il problema n.1 da curare subito.

### 🟢 Cosa funziona
Elenco puntato di ciò che va bene, ognuno con il numero che lo dimostra.

### 🔴 Cosa non funziona
Per ogni problema, una riga: **[metrica/area]** — sintomo → causa probabile.

### 🟡 Da tenere d'occhio
Cose non ancora critiche ma da monitorare.

### 💊 Terapia — interventi in ordine di priorità
Lista numerata. Per OGNI intervento usa SEMPRE questo schema, concreto e specifico (mai generico):
**N. [titolo breve dell'intervento]**
- **Dove:** a quale livello e in quale punto agire (es. "livello Inserzione → creatività", "livello Gruppo di inserzioni → pubblico", "landing page").
- **Cosa:** l'azione precisa da fare.
- **Perché:** il motivo, legato al dato che lo dimostra.

Esempio del livello di concretezza richiesto:
**1. Rinnova la creatività**
- **Dove:** livello Inserzione (creatività).
- **Cosa:** sostituisci il visual e testa 2 nuovi hook nelle prime 3 righe del copy.
- **Perché:** frequenza a 4,2 con CTR in calo = creatività "bruciata", non un problema di budget.

### ❓ Cosa mi serve da te (se manca)
Solo se rilevante: le informazioni che non hai potuto dedurre (es. città, ticket medio, obiettivo), perché servono e come cambierebbero la diagnosi. Invita l'utente ad aggiungerle nei campi o nelle note.

Non usare un tono da manuale. Parla come un medico che guarda la "cartella clinica" della campagna: chiaro, sintetico, senza disclaimer inutili.`;

app.post("/api/analyze", upload.single("file"), async (req, res) => {
  try {
    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(401).json({
        error:
          "Chiave API mancante. Copia .env.example in .env e inserisci ANTHROPIC_API_KEY, poi riavvia il server.",
      });
    }

    const { city, objective, budget, duration, sector, ticket, notes, report, level } =
      req.body || {};

    // Costruisce il blocco di contesto inserito dall'utente
    const ctx = [
      ["Città / area", city],
      ["Obiettivo campagna", objective],
      ["Livello dei dati caricati", level],
      ["Budget", budget],
      ["Da quanto è attiva", duration],
      ["Settore / prodotto", sector],
      ["Ticket medio / valore cliente", ticket],
      ["Note aggiuntive", notes],
    ]
      .filter(([, v]) => v && String(v).trim())
      .map(([k, v]) => `- ${k}: ${v}`)
      .join("\n");

    const content = [];
    let hasData = false;

    // Testo: contesto + eventuale report incollato
    let textBlock = "Analizza questa campagna Meta Ads.\n\n## Contesto fornito\n" +
      (ctx || "(nessun contesto fornito)");

    if (report && report.trim()) {
      hasData = true;
      textBlock +=
        "\n\n## Report incollato dall'utente\n```\n" + report.trim() + "\n```";
    }

    // File caricato
    if (req.file) {
      const mime = req.file.mimetype;
      const name = (req.file.originalname || "").toLowerCase();
      const isImage = IMAGE_TYPES[mime] || mime.startsWith("image/");
      const isText =
        TEXT_EXT.some((e) => name.endsWith(e)) ||
        mime.startsWith("text/") ||
        mime === "application/json";
      const isXlsx =
        name.endsWith(".xlsx") ||
        mime ===
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      const isOldXls = name.endsWith(".xls") || mime === "application/vnd.ms-excel";

      if (isImage) {
        hasData = true;
        content.push({
          type: "image",
          source: {
            type: "base64",
            media_type: IMAGE_TYPES[mime] || "image/png",
            data: req.file.buffer.toString("base64"),
          },
        });
        textBlock +=
          "\n\nÈ allegato uno SCREENSHOT del Gestione Inserzioni: leggi le metriche visibili (spesa, CTR, CPM, CPC, risultati, copertura, frequenza, ecc.) e usale per l'analisi. Se un valore non è leggibile, dillo.";
      } else if (isText) {
        hasData = true;
        const txt = req.file.buffer.toString("utf-8").slice(0, 100000);
        textBlock +=
          "\n\n## Report caricato (file: " +
          req.file.originalname +
          ")\n```\n" +
          txt +
          "\n```";
      } else if (isXlsx) {
        try {
          const csv = await xlsxToText(req.file.buffer);
          if (csv.trim()) {
            hasData = true;
            textBlock +=
              "\n\n## Report Excel caricato (file: " +
              req.file.originalname +
              ", convertito in tabella)\n```\n" +
              csv.slice(0, 100000) +
              "\n```";
          } else {
            textBlock +=
              "\n\n(Il file Excel '" +
              req.file.originalname +
              "' risulta vuoto o illeggibile: chiedi all'utente di ricontrollare l'export.)";
          }
        } catch (e) {
          console.error("xlsx parse error:", e?.message);
          textBlock +=
            "\n\n(Non sono riuscito a leggere il file Excel '" +
            req.file.originalname +
            "': chiedi all'utente un export in CSV o uno screenshot.)";
        }
      } else if (isOldXls) {
        textBlock +=
          "\n\n(È stato caricato un vecchio formato .xls non supportato: chiedi all'utente di riesportare in .xlsx o .csv.)";
      } else {
        textBlock +=
          "\n\n(È stato caricato un file '" +
          req.file.originalname +
          "' in un formato non supportato: chiedi all'utente uno screenshot PNG/JPG o un export CSV/Excel.)";
      }
    }

    content.push({ type: "text", text: textBlock });

    if (!hasData && !ctx) {
      return res
        .status(400)
        .json({ error: "Fornisci almeno uno screenshot, un report o i dati di contesto." });
    }

    const response = await client.messages.create({
      model: "claude-opus-4-8",
      max_tokens: 8000,
      thinking: { type: "adaptive" },
      output_config: { effort: "high" },
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content }],
    });

    const text = response.content
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join("\n");

    res.json({ analysis: text });
  } catch (err) {
    console.error(err);
    const msg =
      err?.status === 401
        ? "Chiave API non valida o mancante. Imposta ANTHROPIC_API_KEY nel file .env."
        : err?.message || "Errore durante l'analisi.";
    res.status(err?.status || 500).json({ error: msg });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  Meta Ads Clinic — backend attivo su http://localhost:${PORT}\n`);
  if (!process.env.ANTHROPIC_API_KEY) {
    console.log(
      "  ⚠️  ANTHROPIC_API_KEY non impostata: copia .env.example in .env e inserisci la chiave.\n"
    );
  }
});
