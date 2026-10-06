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

    const model = genAI.getGenerativeModel({ model: "gemini-3.8-flash" });

    let historyTestuale = "";
    if (history && history.length > 0) {
      historyTestuale = "--- CRONOLOGIA DELLA CHAT FINORA (Usa queste info per applicare regole e preferenze dell'utente nel tempo) ---\n";
      history.forEach((msg: any) => {
        historyTestuale += `${msg.role === 'user' ? 'UTENTE' : 'ASSISTENTE'}:${msg.text}\n`;
      });
      historyTestuale += "-------------------------------------------------------------\n\n";
    }

    // IL CERVELLO CONDIVISO: Le regole d'oro di Asso Antincendio
    const regoleAssoAntincendio = `
LE TAVOLE DELLA LEGGE DI ASSO ANTINCENDIO (REGOLE DA RISPETTARE SEMPRE):

1. LOGICA "A GOCCIA" E VICINANZA GEOGRAFICA:
   - È assolutamente VIETATO far fare "salti qua e là" (ping-pong) al tecnico (es. Nervi -> Centro -> Voltri). Usa SEMPRE la logica della massima vicinanza tra i clienti.
   - Quando si coprono zone fuori città o fuori provincia, il giro DEVE iniziare dal cliente più lontano al mattino, per poi inserire i clienti man mano più vicini in direzione dell'ufficio (Genova) nel pomeriggio.
   - NON mischiare MAI direttrici opposte nello stesso giorno (Es: la Valle di Masone non si mischia con la riviera di Ponente; la Val Scrivia non si mischia con il Levante).

2. ORARI E TOLLERANZA (L'OBIETTIVO È IL RIENTRO IN SEDE):
   - Inizio turno: ore 08:00 dall'ufficio.
   - Fine turno: ore 17:30.
   - Devi calcolare il giro in modo che il tecnico arrivi al cliente finale e rientri verso l'ufficio intorno alle 17:00 / 17:15.
   - Tolleranza traffico: È ammesso un margine di 15-30 minuti (massimo entro le 18:00) SOLO per permettere il tragitto finale di rientro o completare l'ultimo intervento. Qualsiasi cantiere che faccia sforare le 18:00 deve essere SCARTATO nei "sospesi".

3. GESTIONE DEI BUCHI A FINE GIORNATA:
   - Se il giro finisce presto (es. alle 15:00) e il tecnico sta tornando verso Genova, verifica se nei "sospesi" ci sono clienti vicini (in zona) per riempire il buco.
   - SE nei sospesi sono rimasti SOLO clienti lontani (es. tecnico in centro, sospeso a Rapallo), NON INSERIRLI. Piuttosto lascia la giornata mezza vuota e fai rientrare il tecnico.

4. GESTIONE DEI CANTIERI LUNGHI (I MACIGNI):
   - Interventi > 6 ore: vanno piazzati SEMPRE a inizio giornata.
   - Interventi > 8 ore: NON entrano fisicamente in una giornata. Scartali e mettili nei "sospesi", poi nella chat avvisa l'utente e chiedigli come vuole dividerli in più giorni.

5. [SOS] E [INS]:
   - [SOS] significa "Estintore da restituire al cliente". 
   - [INS] significa insoluto contabile. Non hanno priorità logistica particolare, trattali come cantieri normali a meno che l'utente non ti dica diversamente in chat.
`;

    if (action === 'pianifica_settimana') {
      const systemPrompt = `Sei il Responsabile Logistica Senior (Dispatcher) di Asso Antincendio a Genova. Il tuo compito è prendere una lista di cantieri grezzi e suddividerli in una settimana lavorativa, restituendo un piano perfettamente ottimizzato per il tecnico.

${regoleAssoAntincendio}

Ricorda di consultare la cronologia della chat allegata: se l'utente ti ha dato in passato delle regole fisse o delle eccezioni, applicale rigorosamente alla pianificazione.

OUTPUT RICHIESTO:
Restituisci UNICAMENTE un oggetto JSON valido racchiuso tra parentesi graffe, senza markdown o testo aggiuntivo. Struttura:
{
  "giorni": {
    "2026-10-05": ["codiceCliente1", "codiceCliente2"]
  },
  "sospesi": ["codiceCliente4"]
}`;

      const promptData = `${historyTestuale}DATI DA ELABORARE:\n\nGIORNATE E ANCORE:\n${JSON.stringify(data.giornate, null, 2)}\n\nCANTIERI:\n${JSON.stringify(data.interventi, null, 2)}`;

      const result = await generateWithRetry(model, [{ text: systemPrompt }, { text: promptData }]);
      const responseText = result.response.text();
      const matchJson = responseText.match(/\{[\s\S]*\}/);
      if (!matchJson) throw new Error("Nessun JSON valido trovato nella risposta IA");
      const pianoParsed = JSON.parse(matchJson[0]);
      
      return NextResponse.json({ success: true, piano: pianoParsed });
    }

    if (action === 'modifica_chat') {
      const systemPrompt = `Sei l'Assistente Logistico Senior (Dispatcher) di Asso Antincendio a Genova. Devi analizzare il comando dell'utente, rispondere in modo proattivo e applicare le modifiche allo STATO LOGISTICO ATTUALE della plancia.

${regoleAssoAntincendio}

${historyTestuale}

STATO LOGISTICO ATTUALE RICEVUTO DAL SISTEMA FRONTEND:
${JSON.stringify(statoAttuale, null, 2)}

REGOLE ASSOLUTE DI RISPOSTA IN CHAT:
1. "messaggioChat": Sii proattivo e di supporto. Usa il tono di un vero capozona. Fai notare se l'utente ti chiede una modifica che vìola una delle regole aziendali (es. mischiare direttrici), ma eseguila comunque se è un comando diretto. Se ci sono cantieri > 8 ore nei sospesi, chiedi all'utente come vuole dividerli. 
2. Ordine dell'Array: Quando restituisci l'elenco dei cantieri nel campo "giorni", l'ordine è cruciale. L'inizio dell'array è il mattino, la fine dell'array è il pomeriggio.
3. PREGRESSI: I codici che iniziano con "PREGRESSO_" sono eventi fissi di Google Calendar. NON PUOI RIMUOVERLI. Devi solo posizionarli nell'array nel punto giusto per fare da "spartiacque" temporale (es. prima del pregresso, dopo il pregresso).
4. Nessun tralasciamento: Ricopia integralmente i codici dei cantieri e i giorni che non hai toccato. I cantieri che l'utente ti chiede di togliere da un giorno devono finire obbligatoriamente nell'array "sospesi".
5. Restituisci ESCLUSIVAMENTE JSON puro, senza blocchi di markdown \`\`\`json.

FORMATO JSON ESATTO E OBBLIGATORIO:
{
  "messaggioChat": "Ciao capo, ho rimosso il cantiere come richiesto. Ho evitato di infilare un SOS lontano perché...",
  "piano": {
    "giorni": {
      "YYYY-MM-DD": ["codice1", "PREGRESSO_...", "codice2"]
    },
    "sospesi": ["codice3"]
  }
}`;

      const instruction = `Comando o constatazione dell'utente: "${comandoUtente}"`;

      const result = await generateWithRetry(model, [{ text: systemPrompt }, { text: instruction }]);
      const responseText = result.response.text();
      
      const matchJson = responseText.match(/\{[\s\S]*\}/);
      if (!matchJson) throw new Error("Nessun JSON valido trovato nella risposta della chat");
      
      let parsedOutput;
      try {
        parsedOutput = JSON.parse(matchJson[0]);
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