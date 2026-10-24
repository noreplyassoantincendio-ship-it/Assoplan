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

    // TEMP 0.1: Massima rigidità matematica
    const model = genAI.getGenerativeModel({ 
      model: "gemini-1.5-flash-latest",
      generationConfig: { temperature: 0.1 }
    });

    let historyTestuale = "";
    if (history && history.length > 0) {
      historyTestuale = "--- CRONOLOGIA DELLA CHAT FINORA (Usa queste info per applicare regole e preferenze dell'utente) ---\n";
      history.forEach((msg: any) => {
        historyTestuale += `${msg.role === 'user' ? 'UTENTE' : 'ASSISTENTE'}: ${msg.text}\n`;
      });
      historyTestuale += "-------------------------------------------------------------\n\n";
    }

    // IL NUOVO CERVELLO MULTI-AGENTE
    const regoleAssoAntincendio = `
LE TAVOLE DELLA LEGGE DI ASSO ANTINCENDIO - ARCHITETTURA MULTI-AGENTE:

Sei un sistema direzionale composto da 3 Agenti IA che lavorano in sequenza. Esegui il tuo ragionamento diviso in 3 FASI ben distinte (pensiero ad alta voce) prima di generare il JSON finale.

FASE 1: L'AGENTE CARTOGRAFO (Stratega Spaziale)
- Analizza tutti i cantieri e raggruppali per vicinanza geografica (CAP, Località).
- Regola assoluta: Non mischiare mai direttrici opposte nello stesso giorno (Es. Levante e Ponente non si mischiano).
- I cantieri fuori città partono dal più lontano al mattino, rientrando verso Genova nel pomeriggio.

FASE 2: L'AGENTE CRONOMETRISTA (Gestione Tassativi e Time-Blocking)
- Inizio turno: 08:00 da Genova (primo cliente ~08:30). Fine turno: 17:15 (Tolleranza max 18:00).
- LA REGOLA DEL MURO (PREGRESSI): I codici "PREGRESSO_" sono eventi fissi e inamovibili (Tassativi). 
- Se trovi un PREGRESSO a metà giornata (es. alle 14:00), NON puoi lasciare la mattina vuota. Devi spezzare la giornata in due blocchi di tempo:
  * Blocco Mattina: dalle 08:30 fino all'orario di inizio del pregresso. RIEMPILO con cantieri vicini al pregresso.
  * Blocco Pomeriggio: dalla fine del pregresso fino alle 17:15. RIEMPILO con cantieri sulla via del ritorno.
- I cantieri lunghi (> 6 ore) vanno inseriti come primo intervento della giornata.
- I cantieri > 8 ore non ci stanno fisicamente: scartali nei sospesi.

FASE 3: L'AGENTE SUPERVISORE (Controllo Qualità Spietato)
- Applica questa Checklist di correzione interna prima di validare il piano:
  1. Hai lasciato ore vuote al mattino prima di un tassativo pomeridiano? Se sì, correggi inserendo cantieri.
  2. La somma dei minuti fa sforare le 18:00 al tecnico? Se sì, SCARTA l'ultimo cantiere e mettilo nei "sospesi".
  3. L'utente ti ha dato un ordine esplicito in chat? Se sì, questo prevale su tutte le regole precedenti.

ESEMPI DI ADDESTRAMENTO (FEW-SHOT PROMPTING - COPIA QUESTI RAGIONAMENTI):
- CASO A (Tassativo Pomeridiano): C'è un PREGRESSO alle 14:30 a Genova Centro. 
  Soluzione corretta: Il Cronometrista riempie il blocco 08:30-12:30 con 3 interventi a Genova Est/Centro. Pausa. Alle 14:30 c'è il pregresso. Dalle 15:30 alle 17:00 inserisce un altro intervento verso l'ufficio.
- CASO B (Buco a fine giornata): Il tecnico finisce alle 15:00 a Voltri. Nei sospesi c'è Nervi (lontano) e Pra (vicino).
  Soluzione corretta: Inserisci Pra. Lascia Nervi nei sospesi per non far attraversare la città al tecnico.

[SOS] (Reso Estintori) e [INS] (Insoluti) sono cantieri normali a livello di logistica stradale.
`;

    if (action === 'pianifica_settimana') {
      const systemPrompt = `Sei il Sistema Multi-Agente di Asso Antincendio. Devi prendere una lista di cantieri grezzi e ottimizzarli per una settimana lavorativa.

${regoleAssoAntincendio}

OUTPUT RICHIESTO:
Scrivi il tuo ragionamento analitico esplicitando le 3 fasi (CARTOGRAFO, CRONOMETRISTA, SUPERVISORE) in un blocco di testo.
Poi, alla fine, restituisci ESCLUSIVAMENTE il piano in formato JSON racchiuso tra parentesi graffe. Struttura esatta:
{
  "giorni": {
    "2026-10-05": ["codiceCliente1", "PREGRESSO_123", "codiceCliente2"]
  },
  "sospesi": ["codiceCliente4"]
}`;

      const promptData = `${historyTestuale}DATI DA ELABORARE:\n\nGIORNATE E ANCORE (PREGRESSI):\n${JSON.stringify(data.giornate, null, 2)}\n\nCANTIERI DA INCASTRARE:\n${JSON.stringify(data.interventi, null, 2)}`;

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
      const systemPrompt = `Sei il Sistema Multi-Agente di Asso Antincendio. Devi analizzare il comando dell'utente e applicare le modifiche alla plancia attuale.

${regoleAssoAntincendio}

${historyTestuale}

STATO LOGISTICO ATTUALE (DA MODIFICARE):
${JSON.stringify(statoAttuale, null, 2)}

REGOLE DI RISPOSTA CHAT:
1. Ordine: L'inizio dell'array è il mattino, la fine dell'array è il pomeriggio. Usa i PREGRESSO_... come spartiacque.
2. I cantieri rimossi da un giorno vanno obbligatoriamente spostati nell'array "sospesi".
3. Ricopia integralmente i codici che non hai toccato.

OUTPUT RICHIESTO:
Fase 1: Scrivi la tua analisi come SUPERVISORE. Valuta la fattibilità dell'ordine.
Fase 2: Genera il JSON racchiuso tra parentesi graffe. Formato:
{
  "messaggioChat": "Capo, ho fatto come hai chiesto. Ho spostato il cantiere prima del tassativo...",
  "piano": {
    "giorni": {
      "2026-10-05": ["codice1", "PREGRESSO_...", "codice2"]
    },
    "sospesi": ["codice3"]
  }
}`;

      const instruction = `Comando dell'utente: "${comandoUtente}"`;

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