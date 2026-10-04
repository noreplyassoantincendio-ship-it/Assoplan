import { NextResponse } from 'next/server';

export async function POST(req: Request) {
  console.log("=== INIZIO ELABORAZIONE CHAT IA ===");
  
  try {
    const body = await req.json();
    console.log("1. Dati ricevuti dal frontend. Azione:", body.action);

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      console.error("❌ ERRORE: GEMINI_API_KEY non trovata nel file .env.local!");
      return NextResponse.json({ success: false, error: 'API key mancante' }, { status: 500 });
    }

    const prompt = `
      Sei l'assistente logistico IA di Asso Antincendio (Liguria/Genova).
      Il tuo compito è analizzare e modificare la pianificazione dei tecnici.
      
      AZIONE RICHIESTA: ${body.action}
      COMANDO UTENTE: "${body.comandoUtente}"
      
      DATI FORNITI (STATO ATTUALE):
      ${JSON.stringify(body.statoAttuale || body.datiGrezzi)}
      
      REGOLE DI RISPOSTA (CRITICO):
      Devi restituire ESCLUSIVAMENTE un oggetto JSON valido.
      NON usare formattazioni markdown (come i backtick \`\`\`json). Non inserire testo o spiegazioni.
      
      Formato richiesto:
      {
        "success": true,
        "piano": {
          "interventi": [ ... array aggiornato ... ]
        }
      }
    `;

    // ARRAY DI MODELLI DA PROVARE A CASCATA SE UNO E' OCCUPATO (Errore 503)
    const modelliDaProvare = [
      "gemini-flash-latest", 
      "gemini-2.5-flash-lite", 
      "gemini-3.5-flash-lite",
      "gemini-pro-latest"
    ];

    let response = null;
    let modelloUsato = "";

    for (const modelName of modelliDaProvare) {
      console.log(`2. Tentativo con il modello: ${modelName}...`);
      
      response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.1 }
        })
      });

      if (response.ok) {
        modelloUsato = modelName;
        break; // Usciamo dal loop se la chiamata ha successo
      } else {
        console.log(`⚠️ Modello ${modelName} occupato o non trovato, passo al successivo...`);
      }
    }

    if (!response || !response.ok) {
      console.error("❌ 3. TUTTI I MODELLI SONO OCCUPATI O HANNO RESTITUITO ERRORE.");
      return NextResponse.json({ success: false, error: 'Tutti i server IA sono occupati, riprova tra poco.' }, { status: 503 });
    }

    console.log(`✅ 3. Connessione stabilita con successo usando: ${modelloUsato}`);

    const data = await response.json();
    let aiText = data.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!aiText) throw new Error('Risposta vuota o formato non valido da Gemini');

    // Pulizia da blocchi markdown
    aiText = aiText.replace(/```json/gi, '').replace(/```/g, '').trim();
    const firstBrace = aiText.indexOf('{');
    const lastBrace = aiText.lastIndexOf('}');
    
    if (firstBrace !== -1 && lastBrace !== -1) {
      aiText = aiText.substring(firstBrace, lastBrace + 1);
    }

    console.log("4. Parsing del JSON in corso...");
    const jsonParsed = JSON.parse(aiText);
    
    console.log("✅ 5. Parsing riuscito! Dati inviati alla tabella.");
    return NextResponse.json(jsonParsed);

  } catch (error: any) {
    console.error('❌ ERRORE CRITICO:', error.message);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}