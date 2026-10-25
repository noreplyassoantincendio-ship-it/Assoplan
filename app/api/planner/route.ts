import { NextResponse } from 'next/server';
import { GoogleGenerativeAI } from '@google/generative-ai';

const apiKey = process.env.GEMINI_API_KEY || "";
const genAI = new GoogleGenerativeAI(apiKey);

async function generateWithRetry(model: any, contents: any, retries = 3, delay = 2000): Promise<any> {
  for (let i = 0; i < retries; i++) {
    try {
      return await model.generateContent(contents);
    } catch (error: any) {
      const is503 = error.message?.includes('503') || error.status === 503 || error.message?.includes('Service Unavailable');
      if (is503 && i < retries - 1) {
        await new Promise(res => setTimeout(res, delay));
        delay *= 2;
      } else {
        throw error;
      }
    }
  }
}

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { action, data, comandoUtente, statoAttuale, history } = body;

    if (!apiKey) {
      return NextResponse.json({ success: false, error: "API Key mancante nel file .env.local" }, { status: 500 });
    }

    const model = genAI.getGenerativeModel({ 
      model: "gemini-3.8-flash", 
      generationConfig: { 
        temperature: 0.0, 
        responseMimeType: "application/json" 
      }
    });

    if (action === 'pianifica_settimana') {
      
      // ==========================================
      // AGENTE 1: IL CARTOGRAFO LIGURE
      // ==========================================
      const promptCartografo = `Sei l'Agente Cartografo di Asso Antincendio.
Il tuo compito è analizzare i cantieri forniti e raggrupparli in "Cluster Geografici" omogenei per CAP e Località.

⚠️ REGOLA DI FERRO (DIVIETO ASSOLUTO DI INCROCIO):
Un tecnico non deve MAI fare la spola tra valli diverse o tra litorale ed entroterra nello stesso giorno. 
NON inserire MAI nello stesso cluster cantieri che appartengono a macro-zone diverse:

LA BUSSOLA LIGURE (Macro-Zone separate):
- ZONA 1 (PONENTE LITORALE): Arenzano, Voltri, Prà, Pegli, Sestri Ponente, Cornigliano, Mele, Cogoleto, Acquasanta.
- ZONA 2 (ENTROTERRA NORD / VAL POLCEVERA & SCRIVIA): Busalla, Ronco Scrivia, Isola del Cantone, Pontedecimo, Bolzaneto, Rivarolo, Certosa, Campomorone, Mignanego, Serra Riccò, Sant'Olcese.
- ZONA 3 (CENTRO / PORTO / SAMPIERDARENA): Genova Centro, Sampierdarena, Foce, Carignano, Castelletto, San Fruttuoso, Marassi.
- ZONA 4 (LEVANTE LITORALE): Albaro, Sturla, Quarto, Quinto, Nervi, Bogliasco, Pieve Ligure, Sori, Recco, Camogli, Rapallo.
- ZONA 5 (VAL BISAGNO): Staglieno, Molassana, Struppa, Gavette, Bargagli, Torriglia.

Restituisci ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_cartografo": "Spiega come hai diviso le zone applicando la Bussola Ligure",
  "cluster": [
    { "zona": "Nome Macro-Zona", "codici": ["codice1", "codice2"] }
  ]
}
CANTIERI: ${JSON.stringify(data.interventi)}`;

      const resCartografo = await generateWithRetry(model, [{ text: promptCartografo }]);
      const cartografoOutput = JSON.parse(resCartografo.response.text());
      const clusterGeografici = cartografoOutput.cluster;

      // ==========================================
      // AGENTE 2: IL CRONOMETRISTA
      // ==========================================
      const promptCronometrista = `Sei l'Agente Cronometrista di Asso Antincendio. Assegna i cluster geografici alle giornate lavorative disponibili.
REGOLE TASSATIVE:
1. Ogni giornata ha circa 480 minuti lavorativi (08:00 - 18:00 con pausa pranzo).
2. Assegna un solo cluster geografico (o zone strettamente contigue) a ciascuna giornata. Non unire MAI Ponente (es. Arenzano) con Entroterra Nord (es. Busalla) o Levante nello stesso giorno.
3. Se una giornata ha impegni PREGRESSI al pomeriggio (es. oraInizio "14:00"), DEVI inserire cantieri normali della stessa zona PRIMA del pregresso per riempire la mattina.
4. Se ci sono interventi con isSOS = true, cerca di inserirne 3 o 4 al giorno tenendo conto del CAP/Zona.
5. Negli array di ciascun giorno inserisci i codici dei cantieri normali nell'ordine di visita ottimale.

Restituisci ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_incastri": "Spiega come hai assegnato i cluster alle date",
  "giorni": {
    "YYYY-MM-DD": ["cod_1", "cod_2"]
  },
  "sospesi": ["cod_scartato"]
}
CLUSTER RICEVUTI: ${JSON.stringify(clusterGeografici)}
GIORNATE E PREGRESSI: ${JSON.stringify(data.giornate)}
CANTIERI (con durate e dettagli): ${JSON.stringify(data.interventi)}`;

      const resCronometrista = await generateWithRetry(model, [{ text: promptCronometrista }]);
      const bozzaPiano = JSON.parse(resCronometrista.response.text());

      // ==========================================
      // AGENTE 3: IL SUPERVISORE
      // ==========================================
      const promptSupervisore = `Sei il Supervisore Logistico Finale di Asso Antincendio. Verifica e correggi la bozza del piano.
CHECKLIST OBBLIGATORIA:
1. Nessun giorno deve mischiare località incompatibili (es. Arenzano con Busalla, o Ponente con Levante). Se trovi un intruso fuori zona in un giorno, spostalo nel giorno della sua zona o in "sospesi".
2. Controlla che ogni giornata con un PREGRESSO pomeridiano abbia cantieri assegnati per coprire la mattina.
3. Controlla che TUTTI i codici dei cantieri ricevuti in input siano presenti (o dentro "giorni" o dentro "sospesi"). Non perdere nessun codice.

Restituisci ESCLUSIVAMENTE un JSON valido con questa struttura:
{
  "ragionamento_supervisore": "Spiega le verifiche e correzioni effettuate",
  "giorni": {
    "YYYY-MM-DD": ["cod_corretto1", "cod_corretto2"]
  },
  "sospesi": ["cod_finale1"]
}
BOZZA DA REVISIONARE: ${JSON.stringify({ giorni: bozzaPiano.giorni, sospesi: bozzaPiano.sospesi })}
GIORNATE DISPONIBILI: ${JSON.stringify(data.giornate)}
DETTAGLI CANTIERI: ${JSON.stringify(data.interventi)}`;

      const resSupervisore = await generateWithRetry(model, [{ text: promptSupervisore }]);
      const pianoFinale = JSON.parse(resSupervisore.response.text());

      return NextResponse.json({ success: true, piano: { giorni: pianoFinale.giorni, sospesi: pianoFinale.sospesi } });
    }

    if (action === 'modifica_chat') {
      let historyTestuale = "";
      if (history && history.length > 0) {
        historyTestuale = "CRONOLOGIA RECENTE DELLA CHAT:\n";
        history.slice(-6).forEach((msg: any) => {
          historyTestuale += `${msg.role === 'user' ? 'UTENTE' : 'ASSISTENTE'}: ${msg.text}\n`;
        });
      }

      const systemPrompt = `Sei l'Assistente Operativo di Asso Antincendio. Il tuo UNICO compito è eseguire ALLA LETTERA l'ordine dato dall'utente nel COMANDO UTENTE, modificando la plancia attuale.

REGOLE DI ESECUZIONE TASSATIVE:
1. Esegui esattamente ciò che chiede l'utente (spostamenti di giorno, cambio d'ordine primo/ultimo, inserimento dai sospesi, messa in sospeso per cliente, numero BI o località).
2. Usa la mappa "giornateDisponibili" per sapere esattamente a quale data "YYYY-MM-DD" corrisponde il giorno nominato dall'utente (es. Lunedì, Martedì, Mercoledì, Giovedì, Venerdì, Sabato).
3. L'ordine dei codici nell'array di un giorno ("YYYY-MM-DD": ["cod1", "cod2", ...]) determina l'ordine cronologico dal mattino alla sera. Se l'utente chiede di mettere un cantiere "per primo" o "al mattino", mettilo all'inizio dell'array di quel giorno.
4. CONSERVAZIONE TOTALE DEI DATI: Tutti i cantieri che NON sono coinvolti dal comando dell'utente DEVONO rimanere esattamente nel giorno e nell'ordine in cui si trovano adesso. Non dimenticare e non cancellare nessun codice!
5. Tutti i giorni presenti in "giornateDisponibili" devono essere presenti come chiavi nell'oggetto "giorni".

Restituisci ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_operazione": "Indica quali codici hai individuato e quale modifica esatta hai applicato",
  "messaggioChat": "Risposta chiara e concisa per l'utente che conferma l'operazione eseguita",
  "piano": {
    "giorni": {
      "YYYY-MM-DD": ["cod1", "cod2"]
    },
    "sospesi": ["cod3"]
  }
}

${historyTestuale}
STATO ATTUALE COMPLETO DELLA PLANCIA:
${JSON.stringify(statoAttuale)}

COMANDO UTENTE DA ESEGUIRE SUBITO: "${comandoUtente}"`;

      const result = await generateWithRetry(model, [{ text: systemPrompt }]);
      const parsedOutput = JSON.parse(result.response.text());

      return NextResponse.json({ 
        success: true, 
        piano: parsedOutput.piano || { giorni: {}, sospesi: [] }, 
        messaggioChat: parsedOutput.messaggioChat || "Modifica applicata alla plancia."
      });
    }

    return NextResponse.json({ success: false, error: "Azione non riconosciuta" }, { status: 400 });

  } catch (error: any) {
    console.error("API Planner Error:", error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}