import { NextResponse } from 'next/server';
import { GoogleGenerativeAI } from '@google/generative-ai';

const apiKey = process.env.GEMINI_API_KEY || "";
const genAI = new GoogleGenerativeAI(apiKey);

// Funzione per l'estrazione sicura del JSON dalle risposte dell'IA
function extractJSON(testo: string) {
  const firstBrace = testo.indexOf('{');
  const lastBrace = testo.lastIndexOf('}');
  if (firstBrace === -1 || lastBrace === -1) {
      // Se non trova graffe, prova con le parentesi quadre (per il Cartografo)
      const firstSquare = testo.indexOf('[');
      const lastSquare = testo.lastIndexOf(']');
      if (firstSquare !== -1 && lastSquare !== -1) {
          return JSON.parse(testo.slice(firstSquare, lastSquare + 1));
      }
      throw new Error("Nessun JSON valido trovato nella risposta IA.");
  }
  return JSON.parse(testo.slice(firstBrace, lastBrace + 1));
}

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

    // Usiamo il modello veloce per non far aspettare 1 minuto all'utente con 3 chiamate
    const model = genAI.getGenerativeModel({ 
      model: "gemini-1.5-pro-latest",
      generationConfig: { temperature: 0.1 }
    });

    if (action === 'pianifica_settimana') {
      
      // ==========================================
      // AGENTE 1: IL CARTOGRAFO
      // ==========================================
      const promptCartografo = `Sei l'Agente Cartografo di Asso Antincendio.
Il tuo UNICO compito è prendere i cantieri in ingresso e raggrupparli in "Zone/Direttrici" geografiche logiche in base a CAP e Località (es. Val Polcevera, Centro, Levante, Ponente, Entroterra).
Ignora completamente gli orari, le date o la durata degli interventi. Fai solo clustering geografico.

RESTITUISCI ESCLUSIVAMENTE un array JSON puro, senza markdown, con questa struttura:
[
  { "zona": "Genova Ponente", "codici": ["CLI_123_Via...", "BI_456"] },
  { "zona": "Val Scrivia", "codici": ["BI_789"] }
]

CANTIERI DA CLUSTERIZZARE:
${JSON.stringify(data.interventi.map((c: any) => ({ codice: c.codice, indirizzo: c.indirizzo, localita: c.localita })), null, 2)}`;

      const resCartografo = await generateWithRetry(model, [{ text: promptCartografo }]);
      const clusterGeografici = extractJSON(resCartografo.response.text());
      console.log("✅ Agente 1 (Cartografo) completato.");


      // ==========================================
      // AGENTE 2: IL CRONOMETRISTA
      // ==========================================
      const promptCronometrista = `Sei l'Agente Cronometrista di Asso Antincendio.
Hai ricevuto i raggruppamenti geografici dal Cartografo e le giornate disponibili (con i relativi impegni PREGRESSI TASSATIVI di Google Calendar).

REGOLE DI INCASTRO (TIME-BLOCKING):
1. Riempi i giorni assegnando interi "Cluster Geografici" (non mischiare zone opposte nello stesso giorno).
2. I PREGRESSI (es. PREGRESSO_123) sono scogli fissi. 
3. LA REGOLA D'ORO DEL MATTINO: Se in una giornata c'è un PREGRESSO al pomeriggio (es. alle 14:00), è SEVERAMENTE VIETATO lasciare la giornata vuota prima di quell'orario. Devi obbligatoriamente assegnare al mattino dei cantieri della stessa zona (o di strada).
4. Cantieri lunghi (> 300 minuti) vanno messi sempre come primo intervento della giornata.
5. Inserisci i PREGRESSI nell'ordine corretto all'interno dell'array (se è al pomeriggio, mettilo dopo i cantieri del mattino).

RESTITUISCI ESCLUSIVAMENTE un oggetto JSON puro, senza markdown:
{
  "giorni": {
    "2026-10-05": ["codiceCliente1", "codiceCliente2", "PREGRESSO_abc", "codiceCliente3"]
  },
  "sospesi": ["codiceCliente4"]
}

DATI DA ELABORARE:
Cluster Geografici: ${JSON.stringify(clusterGeografici)}
Giornate e Pregressi: ${JSON.stringify(data.giornate)}
Tutti i cantieri (con durate): ${JSON.stringify(data.interventi)}`;

      const resCronometrista = await generateWithRetry(model, [{ text: promptCronometrista }]);
      const bozzaPiano = extractJSON(resCronometrista.response.text());
      console.log("✅ Agente 2 (Cronometrista) completato.");


      // ==========================================
      // AGENTE 3: IL SUPERVISORE
      // ==========================================
      const promptSupervisore = `Sei il Supervisore Logistico Finale di Asso Antincendio.
Stai analizzando la bozza del piano logistico generato dal tuo sottoposto. Il tuo compito è trovare e correggere i suoi errori prima della consegna definitiva.

CHECKLIST DI CORREZIONE:
1. Ha lasciato dei buchi enormi al mattino prima di un tassativo pomeridiano?
   - CORREZIONE: Cerca nei "sospesi" dei cantieri di strada o della stessa zona e infilali al mattino prima del PREGRESSO.
2. Ha sovraccaricato una singola giornata? (Considera un limite di circa 480 minuti effettivi di lavoro per tecnico al giorno, più i viaggi).
   - CORREZIONE: Se una giornata è palesemente troppo piena (> 6 interventi lunghi), togli l'ultimo dell'array e spostalo nei "sospesi".
3. Ha dimenticato i PREGRESSI? Controlla che tutti i codici che iniziano per PREGRESSO_ siano ancora nel loro giorno corretto. NON rimuovere o spostare di giorno i pregressi.

RESTITUISCI ESCLUSIVAMENTE l'oggetto JSON FINALE CORRETTO, senza markdown, con questa struttura:
{
  "giorni": {
    "2026-10-05": ["codiceCorretto1", "PREGRESSO_abc"]
  },
  "sospesi": ["codiceCorretto4"]
}

BOZZA DA REVISIONARE:
${JSON.stringify(bozzaPiano)}

DATI EXTRA PER LA REVISIONE:
Giornate originali: ${JSON.stringify(data.giornate)}
Tutti i cantieri: ${JSON.stringify(data.interventi)}`;

      const resSupervisore = await generateWithRetry(model, [{ text: promptSupervisore }]);
      const pianoFinale = extractJSON(resSupervisore.response.text());
      console.log("✅ Agente 3 (Supervisore) completato. Piano generato.");

      return NextResponse.json({ success: true, piano: pianoFinale });
    }

    if (action === 'modifica_chat') {
      const systemPrompt = `Sei l'Assistente Operativo di Asso Antincendio. Devi analizzare il comando dell'utente e applicare le modifiche alla plancia attuale.
Ricorda: l'ordine degli array indica il percorso (da inizio a fine giornata). I PREGRESSO_... sono ancore inamovibili. I cantieri tolti da un giorno vanno messi in "sospesi".

STATO LOGISTICO ATTUALE (DA MODIFICARE):
${JSON.stringify(statoAttuale, null, 2)}

RESTITUISCI ESCLUSIVAMENTE UN JSON PURO (Niente markdown):
{
  "messaggioChat": "Capo, ho spostato il cantiere come richiesto...",
  "piano": {
    "giorni": { "2026-10-05": ["codice1"] },
    "sospesi": ["codice3"]
  }
}
Comando dell'utente: "${comandoUtente}"`;

      const result = await generateWithRetry(model, [{ text: systemPrompt }]);
      const parsedOutput = extractJSON(result.response.text());

      if (!parsedOutput.piano) parsedOutput.piano = { giorni: {}, sospesi: [] };

      return NextResponse.json({ 
        success: true, 
        piano: parsedOutput.piano, 
        messaggioChat: parsedOutput.messaggioChat || "Ho aggiornato la plancia operativa."
      });
    }

    return NextResponse.json({ success: false, error: "Azione non riconosciuta" }, { status: 400 });

  } catch (error: any) {
    console.error("API Planner Error:", error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}