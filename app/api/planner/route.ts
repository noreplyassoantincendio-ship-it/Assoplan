import { NextResponse } from 'next/server';
import { GoogleGenerativeAI } from '@google/generative-ai';

const apiKey = process.env.GEMINI_API_KEY || "";
const genAI = new GoogleGenerativeAI(apiKey);

// Funzione di utilità per riprovare in caso di sovraccarico temporaneo dei server Google
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

    // 1. IL GUINZAGLIO STRUTTURALE: Modello Flash bloccato in modalità JSON puro a creatività zero
    const model = genAI.getGenerativeModel({ 
      model: "gemini-3.8-flash", 
      generationConfig: { 
        temperature: 0.0, // Zero creatività, massima precisione matematica
        responseMimeType: "application/json" // IMPEDISCE ERRORI DI SINTASSI (Niente Markdown, niente testo libero)
      }
    });

    if (action === 'pianifica_settimana') {
      
      // ==========================================
      // AGENTE 1: IL CARTOGRAFO
      // ==========================================
      const promptCartografo = `Sei l'Agente Cartografo. Raggruppa i cantieri per CAP/Località. Non mischiare Levante e Ponente nello stesso gruppo.
Devi restituire ESCLUSIVAMENTE un JSON valido con questa struttura esatta (rispetta rigorosamente le chiavi):
{
  "ragionamento_cartografo": "Spiega brevemente come hai diviso le zone (es. Genova Est vs Ponente)",
  "cluster": [
    { "zona": "Nome Zona", "codici": ["codice1", "codice2"] }
  ]
}
CANTIERI: ${JSON.stringify(data.interventi.map((c: any) => ({ codice: c.codice, indirizzo: c.indirizzo, localita: c.localita })))}`;

      const resCartografo = await generateWithRetry(model, [{ text: promptCartografo }]);
      const cartografoOutput = JSON.parse(resCartografo.response.text());
      const clusterGeografici = cartografoOutput.cluster;
      console.log("✅ Agente 1 (Cartografo) completato:", cartografoOutput.ragionamento_cartografo);


      // ==========================================
      // AGENTE 2: IL CRONOMETRISTA
      // ==========================================
      const promptCronometrista = `Sei l'Agente Cronometrista. Assegna i cluster geografici ai giorni della settimana per creare il piano di lavoro.
REGOLE TASSATIVE:
1. I codici che iniziano con PREGRESSO_ (es. PREGRESSO_123) sono inamovibili (appuntamenti fissi) e vanno lasciati nel giorno e nell'ordine corretto.
2. REGOLA DEL MATTINO: Se un giorno contiene un PREGRESSO al pomeriggio (es. alle 14:00 o alle 15:00), è severamente vietato lasciare la mattina di quel giorno vuota. Inserisci altri cantieri dello stesso cluster PRIMA del PREGRESSO.
3. I cantieri molto lunghi (> 300 min) vanno messi come primo intervento della giornata.

Devi restituire ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_incastri": "Spiega come hai posizionato i cantieri, specie attorno ai pregressi",
  "giorni": { "2026-10-05": ["cod_1", "PREGRESSO_abc", "cod_2"] },
  "sospesi": ["cod_scartato"]
}
CLUSTER RICEVUTI: ${JSON.stringify(clusterGeografici)}
GIORNATE E PREGRESSI: ${JSON.stringify(data.giornate)}
CANTIERI (con durate per incastrarli): ${JSON.stringify(data.interventi)}`;

      const resCronometrista = await generateWithRetry(model, [{ text: promptCronometrista }]);
      const bozzaPiano = JSON.parse(resCronometrista.response.text());
      console.log("✅ Agente 2 (Cronometrista) completato:", bozzaPiano.ragionamento_incastri);


      // ==========================================
      // AGENTE 3: IL SUPERVISORE
      // ==========================================
      const promptSupervisore = `Sei il Supervisore Finale. Correggi la bozza del piano logistico prima della consegna.
CHECKLIST OBBLIGATORIA:
1. Controlla le mattine: l'agente precedente ha lasciato vuota la mattina di un giorno che ha un PREGRESSO pomeridiano? Se sì, sei obbligato a prendere un cantiere affine dai "sospesi" e inserirlo prima del PREGRESSO.
2. Controlla il carico: se la somma dei minuti di un giorno supera i 480/500 minuti (più i viaggi), sposta l'ultimo cantiere di quell'array nella lista "sospesi".
3. Intoccabilità: Non rimuovere né spostare di giorno i codici PREGRESSO_*.

Devi restituire ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_buchi": "Spiega se hai trovato e tappato buchi mattutini",
  "ragionamento_overtime": "Spiega se hai rimosso cantieri per limiti orari eccessivi",
  "giorni": { "2026-10-05": ["cod_corretto1", "PREGRESSO_abc"] },
  "sospesi": ["cod_finale1"]
}
BOZZA DA REVISIONARE: ${JSON.stringify({ giorni: bozzaPiano.giorni, sospesi: bozzaPiano.sospesi })}`;

      const resSupervisore = await generateWithRetry(model, [{ text: promptSupervisore }]);
      const pianoFinale = JSON.parse(resSupervisore.response.text());
      console.log("✅ Agente 3 (Supervisore) completato. Piano approvato.");

      // Restituiamo solo gli array puliti al frontend (nascondendo i ragionamenti all'interfaccia React)
      return NextResponse.json({ success: true, piano: { giorni: pianoFinale.giorni, sospesi: pianoFinale.sospesi } });
    }

    if (action === 'modifica_chat') {
      
      let historyTestuale = "";
      if (history && history.length > 0) {
        historyTestuale = "CRONOLOGIA DELLA CHAT FINORA:\n";
        history.forEach((msg: any) => {
          historyTestuale += `${msg.role === 'user' ? 'UTENTE' : 'ASSISTENTE'}: ${msg.text}\n`;
        });
      }

      const systemPrompt = `Sei l'Assistente Operativo in Chat. Applica il comando logistico richiesto dall'utente sulla plancia.
REGOLE: 
- L'ordine degli ID nell'array rappresenta la sequenza temporale. I PREGRESSO_... sono ancore fisse. 
- I cantieri tolti da un giorno e non riassegnati vanno messi nell'array "sospesi".

Devi restituire ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_operazione": "Spiega logicamente la modifica applicata",
  "messaggioChat": "Capo, ho spostato il cantiere a giovedì come richiesto...",
  "piano": {
    "giorni": { "2026-10-05": ["cod1"] },
    "sospesi": ["cod2"]
  }
}
${historyTestuale}
STATO ATTUALE: ${JSON.stringify(statoAttuale)}
COMANDO UTENTE: "${comandoUtente}"`;

      const result = await generateWithRetry(model, [{ text: systemPrompt }]);
      const parsedOutput = JSON.parse(result.response.text());

      return NextResponse.json({ 
        success: true, 
        piano: parsedOutput.piano || { giorni: {}, sospesi: [] }, 
        messaggioChat: parsedOutput.messaggioChat || "Ho aggiornato la plancia operativa."
      });
    }

    return NextResponse.json({ success: false, error: "Azione non riconosciuta" }, { status: 400 });

  } catch (error: any) {
    console.error("API Planner Error:", error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}