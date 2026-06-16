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

// Serve la pagina del widget (index.html nella stessa cartella del server).
app.get("/", (req, res) => res.sendFile(join(__dirname, "index.html")));

// Logo opzionale: se carichi un file logo.png accanto al server, viene mostrato.
app.get("/logo.png", (req, res) =>
  res.sendFile(join(__dirname, "logo.png"), (err) => {
    if (err && !res.headersSent) res.status(404).end();
  })
);

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

const SYSTEM_PROMPT = `Sei "Meta Ads Clinic", il medico delle campagne Meta Ads (Facebook/Instagram). Fai una DIAGNOSI come un clinico e restituisci SOLO un oggetto JSON conforme allo schema imposto (nessun testo fuori dal JSON). Tono dei testi: medico, chiaro, diretto, concreto. Tutto in italiano.

## Come ragioni
1. Parti SEMPRE dall'obiettivo e dall'economia del cliente (ticket medio, margine). Una metrica è buona/cattiva solo nel contesto: un CPA di 30€ è ottimo se il cliente vale 300€, pessimo se vale 25€.
2. DEDUCI dai dati/colonne/nome file e mettili nei campi:
   - "periodo": l'intervallo di date analizzato (es. "8–14 giu 2026"). Spesso è nel nome del file o nelle colonne data. Se non deducibile: "non rilevato".
   - "livello": "Inserzione (creatività)" se ci sono nomi di singole inserzioni/creatività; "Gruppo di inserzioni" se ci sono ad set; "Campagna" se solo campagne. Se non chiaro: "non rilevato".
3. Benchmark indicativi (adattali a settore/paese):
   - CTR link: <0,8% debole · 0,8–1,5% medio · >1,5% buono.
   - CPM: in Italia spesso 4–12€. Molto alto con poco budget = aste affollate / pubblico ristretto.
   - Frequenza >~3 in pochi giorni = creatività "bruciata".
   - Conversione landing <1–2% = problema landing/offerta, non campagna.
   - Video: 3-sec view basso = problema di hook/primo frame.
4. Trova il collo di bottiglia: poca copertura→budget/pubblico; tante impression ma CTR basso→creatività/targeting; buon CTR ma poche conversioni→landing/offerta/prezzo; CPA alto ma vendite ok→margine/scala.
5. Se tutto il budget è concentrato su 1 sola inserzione, segnalalo e di' come testare/bilanciare.
6. Copy: valutalo SOLO se i testi sono forniti. Le immagini delle creatività di solito NON sono nei file Excel di Meta: se non le hai, non giudicare il visual e, se utile, mettilo tra le cose che servono.
7. Pochi giorni o <~50 conversioni = fase di apprendimento: sconsiglia modifiche drastiche.
8. Non inventare numeri non presenti.

## Come compilare i campi
- "salute": "verde" (va bene), "giallo" (da sistemare), "rosso" (critico). "voto": 0–10.
- "quadro_generale": 2–3 frasi da medico sullo stato di salute complessivo.
- "problema_principale": UNA frase, il problema n.1 da curare subito.
- "cosa_funziona": 2–4 punti positivi, ognuno col numero che lo dimostra.
- "interventi": ogni voce è un BOX (un'area da curare). Ordina dal più grave. Campi:
   - "area": ambito (es. "Creatività", "Copy", "Targeting/Pubblico", "Budget", "Offerta", "Landing", "CPL/CPA").
   - "gravita": "rosso" | "giallo" | "verde".
   - "diagnosi": sintomo + causa probabile, col numero che lo dimostra.
   - "azione": COSA FARE, concreto e da medico. Esempi: "Metti in pausa l'inserzione X (CPL 18€, doppio della media)", "Rinnova il visual e testa 2 nuovi hook", "Sposta il 60% del budget sull'inserzione Y che converte", "Restringi il pubblico / escludi i 18–24".
   - "dove": livello + punto preciso (es. "livello Inserzione → creatività", "livello Gruppo di inserzioni → pubblico", "Landing page").
- "azioni_urgenti": 2–4 cose da fare SUBITO, brevissime e operative (es. "Refresh creativo sull'inserzione con frequenza 4,2", "Metti in pausa l'ad con CPL 18€").
- "cosa_serve": SOLO i dati davvero mancanti e non deducibili (es. città, ticket medio, obiettivo). Per ciascuno UNA riga col perché serve. Se non manca nulla, lascia vuoto.

Sii sintetico: frasi brevi, niente muri di testo. L'utente deve capire al volo dove intervenire e cosa fare.`;

// Schema della diagnosi (structured output) — il modello deve restituire questo JSON.
const DIAGNOSIS_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    periodo: { type: "string" },
    livello: { type: "string" },
    salute: { type: "string", enum: ["verde", "giallo", "rosso"] },
    voto: { type: "integer" },
    quadro_generale: { type: "string" },
    problema_principale: { type: "string" },
    cosa_funziona: { type: "array", items: { type: "string" } },
    interventi: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          area: { type: "string" },
          gravita: { type: "string", enum: ["rosso", "giallo", "verde"] },
          diagnosi: { type: "string" },
          azione: { type: "string" },
          dove: { type: "string" },
        },
        required: ["area", "gravita", "diagnosi", "azione", "dove"],
      },
    },
    azioni_urgenti: { type: "array", items: { type: "string" } },
    cosa_serve: { type: "array", items: { type: "string" } },
  },
  required: [
    "periodo",
    "livello",
    "salute",
    "voto",
    "quadro_generale",
    "problema_principale",
    "cosa_funziona",
    "interventi",
    "azioni_urgenti",
    "cosa_serve",
  ],
};

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
      output_config: {
        effort: "high",
        format: { type: "json_schema", schema: DIAGNOSIS_SCHEMA },
      },
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content }],
    });

    const text = response.content
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join("");

    let result;
    try {
      result = JSON.parse(text);
    } catch (e) {
      console.error("JSON parse error:", e?.message, "raw:", text.slice(0, 300));
      return res.status(502).json({
        error: "La diagnosi non è stata generata correttamente. Riprova tra poco.",
      });
    }

    res.json({ result });
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
