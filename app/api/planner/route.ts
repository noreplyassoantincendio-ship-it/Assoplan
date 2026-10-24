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
      model: "gemini-1.5-flash", 
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
Il tuo compito è analizzare i cantieri forniti e raggrupparli in "Cluster Geografici".

⚠️ REGOLA DI FERRO (DIVIETO ASSOLUTO DI INCROCIO):
Un tecnico non deve MAI fare la spola tra valli diverse o tra litorale ed entroterra. 
NON inserire MAI nello stesso cluster cantieri che appartengono a macro-zone diverse di questa mappa:

LA BUSSOLA LIGURE (Macro-Zone separate):
- ZONA 1 (PONENTE LITORALE): Arenzano, Voltri, Prà, Pegli, Sestri Ponente, Mele, Cogoleto, Acquasanta.
- ZONA 2 (ENTROTERRA NORD): Busalla, Ronco Scrivia, Pontedecimo, Bolzaneto, Rivarolo, Campomorone, Mignanego.
- ZONA 3 (CENTRO / PORTO): Genova Centro, Sampierdarena, Foce, Carignano, San Fruttuoso, Marassi.
- ZONA 4 (LEVANTE LITORALE): Nervi, Quarto, Quinto, Sturla, Bogliasco, Recco, Pieve Ligure.
- ZONA 5 (VAL BISAGNO): Molassana, Struppa, Staglieno, Gavette.

Se vedi un cantiere ad Arenzano e uno a Busalla, DEVI creare due cluster distinti.
Devi restituire ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_cartografo": "Spiega come hai applicato la Bussola Ligure per separare le zone",
  "cluster": [
    { "zona": "Nome Macro-Zona", "codici": ["codice1", "codice2"] }
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
      const promptCronometrista = `Sei l'Agente Cronometrista. Assegna i cluster geografici ai giorni della settimana.
REGOLE TASSATIVE:
1. I codici che iniziano con PREGRESSO_ (es. PREGRESSO_123) sono APPUNTAMENTI FISSI: lasciali nel giorno assegnato.
2. LIMITA LE TRASFERTE: Cerca di assegnare un solo cluster geografico (es. solo Ponente) a una singola giornata, salvo i pregressi.
3. REGOLA DEL MATTINO: Se c'è un PREGRESSO pomeridiano, non lasciare mai la mattina vuota. Riempila con cantieri dello STESSO cluster del pregresso (o di un cluster affine).
4. Cantieri molto lunghi (> 300 min) vanno messi come primo intervento.

Devi restituire ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_incastri": "Spiega le assegnazioni fatte",
  "giorni": { "2026-10-05": ["cod_1", "PREGRESSO_abc", "cod_2"] },
  "sospesi": ["cod_scartato"]
}
CLUSTER RICEVUTI: ${JSON.stringify(clusterGeografici)}
GIORNATE E PREGRESSI: ${JSON.stringify(data.giornate)}
CANTIERI (con durate): ${JSON.stringify(data.interventi)}`;

      const resCronometrista = await generateWithRetry(model, [{ text: promptCronometrista }]);
      const bozzaPiano = JSON.parse(resCronometrista.response.text());
      console.log("✅ Agente 2 (Cronometrista) completato:", bozzaPiano.ragionamento_incastri);

      // ==========================================
      // AGENTE 3: IL SUPERVISORE
      // ==========================================
      const promptSupervisore = `Sei il Supervisore Finale. Correggi la bozza del piano.
CHECKLIST:
1. Controlla le mattine: l'agente precedente ha lasciato vuota la mattina di un giorno che ha un PREGRESSO pomeridiano? Se sì, pesca dai sospesi e riempi.
2. Controlla il carico: se la somma dei minuti di un giorno supera i 480/500 minuti (più i viaggi), sposta l'ultimo cantiere dell'array in "sospesi".
3. Non rimuovere né spostare MAI i PREGRESSO_*.

Devi restituire ESCLUSIVAMENTE un JSON valido con questa struttura:
{
  "ragionamento_buchi": "Spiega se hai trovato buchi",
  "ragionamento_overtime": "Spiega se hai rimosso cantieri in eccesso",
  "giorni": { "2026-10-05": ["cod_corretto1", "PREGRESSO_abc"] },
  "sospesi": ["cod_finale1"]
}
BOZZA DA REVISIONARE: ${JSON.stringify({ giorni: bozzaPiano.giorni, sospesi: bozzaPiano.sospesi })}`;

      const resSupervisore = await generateWithRetry(model, [{ text: promptSupervisore }]);
      const pianoFinale = JSON.parse(resSupervisore.response.text());
      console.log("✅ Agente 3 (Supervisore) completato.");

      return NextResponse.json({ success: true, piano: { giorni: pianoFinale.giorni, sospesi: pianoFinale.sospesi } });
    }

    if (action === 'modifica_chat') {
      let historyTestuale = "";
      if (history && history.length > 0) {
        historyTestuale = "CRONOLOGIA DELLA CHAT FINORA:\n";
        history.forEach((msg: any) => historyTestuale += `${msg.role === 'user' ? 'UTENTE' : 'ASSISTENTE'}: ${msg.text}\n`);
      }

      const systemPrompt = `Sei l'Assistente Operativo in Chat. Applica il comando logistico richiesto dall'utente sulla plancia.
REGOLE: L'ordine degli ID nell'array rappresenta la sequenza temporale. I PREGRESSO_... sono ancore fisse. I cantieri tolti da un giorno e non riassegnati vanno in "sospesi".

Devi restituire ESCLUSIVAMENTE un JSON valido con questa struttura esatta:
{
  "ragionamento_operazione": "Spiega logicamente la modifica",
  "messaggioChat": "Capo, ho spostato il cantiere a giovedì...",
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
        messaggioChat: parsedOutput.messaggioChat || "Aggiornato."
      });
    }

    return NextResponse.json({ success: false, error: "Azione non riconosciuta" }, { status: 400 });

  } catch (error: any) {
    console.error("API Planner Error:", error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}