import express from "express";
import multer from "multer";
import Anthropic from "@anthropic-ai/sdk";
import ExcelJS from "exceljs";
import { runNightly } from "./nightly.js";
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
9. CONSIGLI SOLO SE FONDATI SUI DATI. Non inventare segmenti o azioni specifiche che i dati non supportano. Esempi:
   - Suggerisci di agire su ETÀ/GENERE (es. "tieni solo 45–54 donne") SOLO se nel file c'è il breakdown per età/genere. Altrimenti NON dirlo: mettilo in "cosa_serve" ("esporta con suddivisione per età e genere").
   - Suggerisci CBO/ABO, esclusione posizionamenti, o azioni su singole inserzioni SOLO se i dati lo permettono (livello e colonne presenti).
   - Se hai solo un dato aggregato (una riga), NON parlare di "questa inserzione" o "quel segmento": di' che serve l'export a livello Inserzione (o con i breakdown) per consigli specifici.
   In sintesi: ogni azione deve poter essere giustificata da un numero o una colonna realmente presente.

## Come compilare i campi
- "salute" e "voto": NON dare sempre lo stesso voto. Usa TUTTA la scala 0–10 e sii deciso, con questa rubrica:
   - 8–10 = "verde": campagna sana, efficiente sull'obiettivo → si può scalare. (Nessun problema rosso.)
   - 5–7 = "giallo": funziona ma ha problemi concreti da sistemare.
   - 2–4 = "rosso": in difficoltà seria / spreco di budget / nessun risultato utile.
   - 0–1 = "rosso": gravemente compromessa.
   Coerenza obbligatoria: se c'è ≥1 intervento "rosso", la salute è "rosso" e il voto ≤4. Se gli interventi sono solo "giallo", voto 5–7. Se è tutto buono, voto 8–10.
   METODO DI CALCOLO (applicalo, non dare voti a caso): parti da 10 e sottrai 3 per ogni intervento "rosso" e 1,5 per ogni "giallo"; arrotonda all'intero; minimo 1, massimo 10. Così il voto VARIA tra campagne diverse. È VIETATO dare sempre 6: se ti viene 6, ricontrolla quanti problemi reali ci sono.
- "quadro_generale": 2–3 frasi da medico sullo stato di salute complessivo.
- "problema_principale": UNA frase, il problema n.1 da curare subito.
- "metriche": i NUMERI CHIAVE letti dal report, da mostrare all'utente come "cartella clinica". Includi tutto ciò che è presente nei dati, ad esempio:
   - Importo speso; Risultati e Costo per risultato (CPL/CPA); CTR (link); CPM; CPC; Frequenza; Copertura; Impressioni.
   - Clic sul link e "Visualizzazioni della pagina di destinazione" (i caricamenti completati): mostra ENTRAMBI e calcola il RAPPORTO clic→visualizzazioni (es. "504 / 1.560 = 32%"). Se molti clic ma poche visualizzazioni (es. <70–80%) = landing lenta o abbandono prima del caricamento → stato "rosso"/"giallo" e citalo anche negli interventi (Landing/Tracking).
   - Vendite/ROAS/Valore conversioni se presenti.
   - "Link in uscita" (URL di destinazione) → stato "neutro".
   - Titolo/i principali dell'annuncio se presenti → stato "neutro".
   - Città / aree più performanti se c'è un breakdown geografico (es. "Milano 12 lead, Roma 8") → stato in base al rendimento.
   Per ogni metrica: label breve, valore con unità (es. "1.560 clic", "9,40€", "32%"), stato "verde"|"giallo"|"rosso"|"neutro". Non inventare: metti solo ciò che è nei dati.
- "rilevati": i dati di contesto che riesci a DEDURRE dal report/file, da pre-compilare nel form. Campi (stringa vuota "" se non deducibile):
   - "budget": budget rilevato (es. "270€/giorno" o "≈ 1.900€ in 7 giorni").
   - "durata": da quanto è attiva / periodo coperto (es. "8–14 giu 2026 (7 giorni)").
   - "obiettivo": prova a dedurre il tipo (es. "Contatti/Lead — Modulo Facebook", "Vendite/Conversioni", "Traffico"). Se incerto, "".
   - "citta": area geografica dove girano le ads. DEDUCILA da: (a) la suddivisione per regione/città se presente nei dati; (b) i NOMI di campagne/gruppi/inserzioni (es. "Milano", "Lombardia", "Italia", "Nord"); (c) ciò che scrive l'utente. Scrivi UNA di queste forme: "Campagna nazionale (tutta Italia)" · "Multi-città (es. Milano, Roma, Torino)" · oppure la singola area (es. "Solo Milano e provincia"). Se NON è deducibile da nulla, lascia "" e mettilo in "cosa_serve" ("esporta con suddivisione per Regione/Città per capire dove girano e rendono le ads").
   - "settore": settore/prodotto se deducibile, altrimenti "".
- "dispersione": stima del budget "disperso" = spesa finita su inserzioni/segmenti/posizionamenti/AREE GEOGRAFICHE con risultati nulli o molto costosi rispetto agli altri. PESA per SPESA e TEMPO (non per numero di ads). Se c'è la suddivisione per regione/città, considera anche la dispersione GEOGRAFICA (es. spendi su aree che non convertono) e citala. Campi:
   - "valore": € e % sul totale (es. "≈ 320€ (11% della spesa)"). Se NON calcolabile (solo dato aggregato, niente dettaglio per inserzione/segmento) → "non calcolabile".
   - "giudizio": "verde" (fisiologica), "giallo" (alta), "rosso" (grave spreco), "neutro" (non calcolabile). Regola pratica: in fase di test/apprendimento una dispersione fino a ~10–15% della spesa è normale; oltre è un problema. Rapporta SEMPRE alla spesa totale (disperdere 200€ su 30.000€ è normale; su 600€ è grave).
   - "commento": 1–2 frasi: è accettabile o no, e perché (considera apprendimento e spesa totale). Se "non calcolabile", spiega che serve l'export a livello Inserzione.
- "grafici": SOLO se nei dati ci sono dei BREAKDOWN. Crea un grafico a barre PER OGNI breakdown presente, tra cui:
   - "CPL per genere" (Uomini/Donne) — costo per contatto per genere.
   - "CPL per età" o "Lead per età" — per fascia d'età (18-24, 25-34, ...).
   - "Contatti per inserzione" o "CPL per inserzione" — confronto tra le singole inserzioni.
   - GEOGRAFICO: se ci sono dati per città/regione, "Lead per città" e/o "CPL per città" (così si vede DOVE rende e DOVE disperde il budget).
   Ogni grafico: "titolo", "unita" ("€"|"lead"|"%"|...), "voci" = lista di { label, valore (NUMERO puro, senza simboli), nota (breve, anche "") }. Ordina le voci in modo utile (es. dalla migliore alla peggiore). Se NON ci sono breakdown nei dati, lascia "grafici" VUOTO (non inventarli).
- "ads": se il file contiene le SINGOLE inserzioni (righe per inserzione) e una colonna di stato/recapito (es. "Recapito", "Stato": Attiva / Disattivata / In pausa / Bozza), elenca le inserzioni più rilevanti (le migliori e le peggiori). Per ognuna:
   - "nome": il nome dell'inserzione.
   - "stato": "attiva" | "non attiva" | "sconosciuto" (in base alla colonna; se non c'è la colonna stato, "sconosciuto").
   - "giudizio": "verde" (va bene) | "giallo" | "rosso".
   - "nota": la metrica chiave che la qualifica (es. "CPL 6€, la migliore" / "frequenza 5,1, satura").
   - "azione": cosa fare, COERENTE con lo stato. Per le ATTIVE: "Scala (+20% budget)", "Mantieni e monitora", "Rinnova la creatività", "Metti in pausa". Per le NON ATTIVE: "Riattiva e scala (andava bene)" oppure "Lasciala spenta (non rendeva)".
   Se il file NON ha il dettaglio per inserzione o NON ha la colonna stato, lascia "ads" VUOTO e aggiungi in "cosa_serve": "Esporta a livello Inserzione con la colonna Recapito/Stato per dividere ads attive e non attive".
- "cosa_funziona": 2–4 punti positivi, ognuno col numero che lo dimostra.
- "interventi": ogni voce è un BOX (un'area da curare). Ordina dal più grave. Campi:
   - "area": ambito (es. "Creatività", "Copy", "Targeting/Pubblico", "Budget", "Offerta", "Landing", "CPL/CPA").
   - "gravita": "rosso" | "giallo" | "verde".
   - "diagnosi": sintomo + causa probabile, col numero che lo dimostra.
   - "azione": COSA FARE, concreto e da medico. Esempi: "Metti in pausa l'inserzione X (CPL 18€, doppio della media)", "Rinnova il visual e testa 2 nuovi hook", "Sposta il 60% del budget sull'inserzione Y che converte", "Restringi il pubblico / escludi i 18–24".
   - "dove": livello + punto preciso (es. "livello Inserzione → creatività", "livello Gruppo di inserzioni → pubblico", "Landing page").
- "azioni_urgenti": 2–4 cose da fare SUBITO, brevissime e operative. Includi, quando i dati lo indicano: mettere in pausa un'inserzione che spreca; RIATTIVARE un'inserzione spenta che però rendeva meglio (in base ai numeri); TESTARE 1–2 nuove creatività/varianti se quelle attive sono sature (frequenza alta). Es.: "Riattiva 'Video offerta' (era a CPL 6€, oggi spendi su una a 14€)", "Metti in pausa l'ad con CPL 18€", "Testa 2 nuovi hook sull'inserzione con frequenza 4,2".
- "cosa_serve": SOLO i dati davvero mancanti e non deducibili (es. città, ticket medio, obiettivo). Per ciascuno UNA riga col perché serve. Se non manca nulla, lascia vuoto.

Sii sintetico: frasi brevi, niente muri di testo. LINGUAGGIO SEMPLICE, per chi NON è esperto di advertising: evita gergo inutile e, se usi una sigla, spiegala in due parole la prima volta (es. "CPL = costo per contatto", "frequenza = quante volte la stessa persona ha visto l'annuncio"). L'utente deve capire al volo lo stato della campagna, dove intervenire e cosa fare.`;

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
    rilevati: {
      type: "object",
      additionalProperties: false,
      properties: {
        budget: { type: "string" },
        durata: { type: "string" },
        obiettivo: { type: "string" },
        citta: { type: "string" },
        settore: { type: "string" },
      },
      required: ["budget", "durata", "obiettivo", "citta", "settore"],
    },
    dispersione: {
      type: "object",
      additionalProperties: false,
      properties: {
        valore: { type: "string" },
        giudizio: { type: "string", enum: ["verde", "giallo", "rosso", "neutro"] },
        commento: { type: "string" },
      },
      required: ["valore", "giudizio", "commento"],
    },
    grafici: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          titolo: { type: "string" },
          unita: { type: "string" },
          voci: {
            type: "array",
            items: {
              type: "object",
              additionalProperties: false,
              properties: {
                label: { type: "string" },
                valore: { type: "number" },
                nota: { type: "string" },
              },
              required: ["label", "valore", "nota"],
            },
          },
        },
        required: ["titolo", "unita", "voci"],
      },
    },
    metriche: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          label: { type: "string" },
          valore: { type: "string" },
          stato: { type: "string", enum: ["verde", "giallo", "rosso", "neutro"] },
        },
        required: ["label", "valore", "stato"],
      },
    },
    ads: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          nome: { type: "string" },
          stato: { type: "string", enum: ["attiva", "non attiva", "sconosciuto"] },
          giudizio: { type: "string", enum: ["verde", "giallo", "rosso"] },
          nota: { type: "string" },
          azione: { type: "string" },
        },
        required: ["nome", "stato", "giudizio", "nota", "azione"],
      },
    },
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
    "rilevati",
    "dispersione",
    "grafici",
    "metriche",
    "ads",
    "cosa_funziona",
    "interventi",
    "azioni_urgenti",
    "cosa_serve",
  ],
};

// Estrae il JSON da una risposta che potrebbe avere fence o testo attorno.
function extractJson(text) {
  let raw = String(text || "").trim();
  raw = raw.replace(/^```(?:json)?/i, "").replace(/```$/i, "").trim();
  const a = raw.indexOf("{"), b = raw.lastIndexOf("}");
  return a >= 0 && b > a ? raw.slice(a, b + 1) : raw;
}

// Esegue la diagnosi (riusata sia dall'endpoint manuale sia dal job notturno).
async function runDiagnosis(content) {
  const response = await client.messages.create({
    model: "claude-opus-4-8",
    max_tokens: 16000,
    thinking: { type: "adaptive" },
    output_config: {
      effort: "medium",
      format: { type: "json_schema", schema: DIAGNOSIS_SCHEMA },
    },
    system: SYSTEM_PROMPT,
    messages: [{ role: "user", content }],
  });
  const rawText = response.content
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("")
    .trim();
  try {
    return JSON.parse(extractJson(rawText));
  } catch (e) {
    const err = new Error("La diagnosi non è stata generata correttamente.");
    err.stopReason = response.stop_reason;
    err.rawHead = rawText.slice(0, 200);
    err.rawTail = rawText.slice(-200);
    err.rawLen = rawText.length;
    throw err;
  }
}

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

    let result;
    try {
      result = await runDiagnosis(content);
    } catch (e) {
      if (e.rawHead !== undefined) {
        console.error("JSON parse fail. stop_reason:", e.stopReason, "len:", e.rawLen, "head:", e.rawHead, "tail:", e.rawTail);
        const hint = e.stopReason === "max_tokens" ? " (la risposta era troppo lunga: riprova)" : "";
        return res.status(502).json({ error: e.message + " Riprova tra poco." + hint });
      }
      throw e;
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

// Endpoint del job notturno: lo chiama un cron esterno (cron-job.org) alle 00:30.
// Protetto da CRON_SECRET. Risponde subito e lavora in background.
app.get("/api/cron/run", (req, res) => {
  if (!process.env.CRON_SECRET || req.query.key !== process.env.CRON_SECRET) {
    return res.status(403).json({ error: "Accesso negato (chiave cron mancante o errata)." });
  }
  res.json({ status: "avviato", ts: new Date().toISOString() });
  runNightly(runDiagnosis)
    .then((s) => console.log("[nightly] completato:", JSON.stringify(s)))
    .catch((e) => console.error("[nightly] errore fatale:", e?.message));
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
