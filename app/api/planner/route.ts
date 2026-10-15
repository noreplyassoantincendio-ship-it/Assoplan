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

    // TEMP 0.1: Obbliga il modello a essere deterministico, logico, freddo e calcolatore. Niente "creatività".
    const model = genAI.getGenerativeModel({ 
      model: "gemini-3.8-flash",
      generationConfig: { temperature: 0.1 }
    });

    let historyTestuale = "";
    if (history && history.length > 0) {
      historyTestuale = "--- CRONOLOGIA DELLA CHAT FINORA (Usa queste info per applicare regole e preferenze dell'utente nel tempo) ---\n";
      history.forEach((msg: any) => {
        historyTestuale += `${msg.role === 'user' ? 'UTENTE' : 'ASSISTENTE'}: ${msg.text}\n`;
      });
      historyTestuale += "-------------------------------------------------------------\n\n";
    }

    // IL CERVELLO CONDIVISO: Le regole d'oro di Asso Antincendio
    const regoleAssoAntincendio = `
LE TAVOLE DELLA LEGGE DI ASSO ANTINCENDIO (REGOLE DA RISPETTARE SEMPRE):

0. OBBEDIENZA ASSOLUTA IN CHAT:
    - Sei uno strumento esecutivo. Devi ubbidire TASSATIVAMENTE agli ordini dell'utente.
    - Se l'utente ti chiede esplicitamente di spostare, assegnare o forzare un cantiere in un giorno specifico, ESEGUILO IMMEDIATAMENTE, ignorando limiti di orario o di logica geografica. L'utente ha sempre l'ultima parola.

1. LOGICA "A GOCCIA" E VICINANZA GEOGRAFICA:
   - È assolutamente VIETATO far fare "salti qua e là" al tecnico. Usa SEMPRE la logica della massima vicinanza tra i clienti.
   - Per zone fuori città: il giro DEVE iniziare dal cliente più lontano al mattino, inserendo i clienti più vicini man mano che si torna verso l'ufficio (Genova) nel pomeriggio.
   - NON mischiare MAI direttrici opposte nello stesso giorno (Es: Valle di Masone non si mischia con la riviera di Ponente).

2. ORARI E TOLLERANZA (L'OBIETTIVO È IL RIENTRO IN SEDE):
   - Inizio turno: ore 08:00 dall'ufficio. Primo appuntamento NON prima delle 08:30.
   - Il tecnico DEVE rientrare in ufficio entro le 17:00 / 17:15.
   - Tempi presidi (estintori/idranti/porte): 3 minuti cadauno. Minimo intervento: 20 minuti (se fino a 3 presidi). Impianto Rivelazione/Sprinkler: 30 minuti a impianto. Trasferimento minimo tra cantieri su Calendar: 20 minuti.
   - Tolleranza traffico: Ammesso un margine di 15-30 minuti (massimo entro le 18:00) SOLO per permettere il tragitto finale di rientro. 
   - I cantieri in eccesso che fanno sforare le 18:00 DEVONO ESSERE SCARTATI nell'array "sospesi".

3. GESTIONE DEI BUCHI A FINE GIORNATA:
   - DEVI riempire TUTTO il tempo disponibile al tecnico senza lasciare buchi, a parte la pausa.
   - Se ci sono eventi pregressi fissati a metà giornata, inserisci interventi prima e dopo tali eventi per ottimizzare i tempi.
   - Se a fine giro (es. ore 15:00) il tecnico torna verso Genova, cerca nei "sospesi" clienti di STRADA per riempire il buco. Se ci sono solo clienti lontani, non inserirli.

4. GESTIONE DEI CANTIERI LUNGHI (I MACIGNI):
   - Interventi > 6 ore: Piazzali SEMPRE a inizio giornata.
   - Interventi > 8 ore: NON entrano in una singola giornata. Scartali nei "sospesi" e nella risposta discorsiva (messaggioChat) avvisa l'utente.

5. [SOS] E [INS]:
   - [SOS] = Estintore da restituire. [INS] = Insoluto contabile. Trattali come cantieri normali a livello logistico.
`;

    if (action === 'pianifica_settimana') {
      const systemPrompt = `Sei il Responsabile Logistica Senior (Dispatcher) di Asso Antincendio a Genova. Il tuo compito è prendere una lista di cantieri grezzi e ottimizzarli per una settimana lavorativa.

${regoleAssoAntincendio}

Ricorda di consultare la cronologia della chat allegata: se l'utente ti ha dato in passato delle regole fisse o delle eccezioni, applicale rigorosamente.

OUTPUT RICHIESTO (RAGIONA PRIMA DI SCRIVERE IL JSON):
Fase 1: Scrivi un blocco testuale che inizia con "ANALISI LOGISTICA:". Qui devi elencare brevemente i raggruppamenti per zona e verificare matematicamente che la somma dei minuti non superi l'orario di fine turno.
Fase 2: Restituisci il piano in formato JSON racchiuso tra parentesi graffe. Struttura esatta:
{
  "giorni": {
    "2026-10-05": ["codiceCliente1", "codiceCliente2"]
  },
  "sospesi": ["codiceCliente4"]
}`;

      const promptData = `${historyTestuale}DATI DA ELABORARE:\n\nGIORNATE E ANCORE:\n${JSON.stringify(data.giornate, null, 2)}\n\nCANTIERI:\n${JSON.stringify(data.interventi, null, 2)}`;

      const result = await generateWithRetry(model, [{ text: systemPrompt }, { text: promptData }]);
      const responseText = result.response.text();
      
      const firstBrace = responseText.indexOf('{');
      const lastBrace = responseText.lastIndexOf('}');
      if (firstBrace === -1 || lastBrace === -1) throw new Error("Nessun JSON valido trovato nella risposta IA");
      
      const jsonStr = responseText.slice(firstBrace, lastBrace + 1);
      const pianoParsed = JSON.parse(jsonStr);
      
      return NextResponse.json({ success: true, piano: pianoParsed });
    }

    if (action === 'modifica_chat') {
      const systemPrompt = `Sei l'Assistente Logistico Senior di Asso Antincendio. Devi analizzare il comando dell'utente e applicare le modifiche allo STATO LOGISTICO ATTUALE.

${regoleAssoAntincendio}

${historyTestuale}

STATO LOGISTICO ATTUALE (DA MODIFICARE):
${JSON.stringify(statoAttuale, null, 2)}

REGOLE DI RISPOSTA:
1. Sii proattivo: Fai notare se l'utente chiede una modifica che vìola una regola (es. mischiare direttrici), ma ESEGUILA COMUNQUE.
2. Ordine: L'inizio dell'array è il mattino, la fine dell'array è il pomeriggio.
3. PREGRESSI: I codici che iniziano con "PREGRESSO_" sono intoccabili, usali come "spartiacque".
4. Non perdere dati: Ricopia integralmente i codici dei cantieri e i giorni che non hai toccato. I cantieri rimossi vanno spostati in "sospesi".

OUTPUT RICHIESTO (RAGIONA PRIMA DI SCRIVERE IL JSON):
Fase 1: Scrivi un blocco testuale che inizia con "ANALISI COMANDO:" in cui valuti l'ordine dell'utente e fai un calcolo logico.
Fase 2: Genera il JSON racchiuso tra parentesi graffe. Formato:
{
  "messaggioChat": "Ciao capo, ho spostato il cantiere...",
  "piano": {
    "giorni": {
      "2026-10-05": ["codice1", "PREGRESSO_...", "codice2"]
    },
    "sospesi": ["codice3"]
  }
}`;

      const instruction = `Comando o constatazione dell'utente: "${comandoUtente}"`;

      const result = await generateWithRetry(model, [{ text: systemPrompt }, { text: instruction }]);
      const responseText = result.response.text();
      
      const firstBrace = responseText.indexOf('{');
      const lastBrace = responseText.lastIndexOf('}');
      if (firstBrace === -1 || lastBrace === -1) throw new Error("Nessun JSON valido trovato nella risposta IA della chat");
      
      const jsonStr = responseText.slice(firstBrace, lastBrace + 1);
      
      let parsedOutput;
      try {
        parsedOutput = JSON.parse(jsonStr);
      } catch (e) {
        throw new Error("Il JSON restituito dall'IA non è formattato correttamente.");
      }

      if (!parsedOutput.piano) {
         parsedOutput.piano = { giorni: {}, sospesi: [] };
      }

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