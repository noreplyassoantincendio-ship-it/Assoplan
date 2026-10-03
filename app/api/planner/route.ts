import { NextResponse } from 'next/server';

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;

const SYSTEM_PROMPT = `
Sei il direttore operativo esperto di Asso Antincendio, azienda di manutenzione presidi antincendio a Genova e in Liguria.
Il tuo compito è ricevere i dati grezzi dei cantieri (o i comandi di modifica via chat) e restituire una pianificazione settimanale perfetta, intelligente e rigorosa.

REGOLE OPERATIVE FONDAMENTALI DA RISPETTARE:
1. GEOGRAFIA E VALLI LIGURI: Non usare mai raggi geometrici in chilometri. Raggruppa i cantieri per macro-aree e valli naturali contigue (es. Asse Valle Scrivia con Busalla, Mignanego e Crocefieschi; Ponente; Levante; Genova centro).
2. SATURAZIONE DEL TURNO (ZERO BUCHI): I tecnici lavorano dalle 08:00 (o dalla fine degli impegni pregressi) fino alle 18:00. Se una zona esaurisce i cantieri, aggancia immediatamente la zona limitrofa o di rientro per sfruttare la giornata senza lasciare buchi.
3. GESTIONE SOS VS ORDINARI:
   - Per gli interventi SOS (restituzioni estintori), calcola i muletti basandoti esclusivamente sulla colonna eo_rv dei pezzi da restituire, senza gonfiare la durata sui totali d'anagrafica.
   - Per la manutenzione ordinaria, se il totale riassuntivo è anomalo, controlla le righe di dettaglio per stimare i minuti corretti.
4. FORMATO DI RISPOSTA: Restituisci sempre e solo un JSON strutturato con la lista degli interventi pianificati, comprensiva di giorno, orari di inizio/fine, minuti di lavoro, tempi di viaggio stimati e distinte di carico dei muletti.
`;

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { action, datiGrezzi, comandoUtente, statoAttuale } = body;

    // Costruiamo il prompt per Gemini basato sull'azione richiesta
    let promptUtente = "";
    if (action === "pianifica_iniziale") {
      promptUtente = `Analizza questi dati grezzi dei cantieri e genera la pianificazione ottimizzata per la settimana:\n${JSON.stringify(datiGrezzi)}`.slice(0, 15000);
    } else if (action === "modifica_chat") {
      promptUtente = `Stato attuale della pianificazione:\n${JSON.stringify(statoAttuale)}\n\nComando dell'utente da applicare:\n"${comandoUtente}"\n\nRicalcola e aggiorna la pianificazione rispettando il comando.`;
    }

    // Chiamata alle API ufficiali di Google Gemini (Gemini 2.5 Flash o Pro)
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [
          { role: "user", parts: [{ text: SYSTEM_PROMPT + "\n\n" + promptUtente }] }
        ],
        generationConfig: {
          responseMimeType: "application/json"
        }
      })
    });

    if (!response.ok) {
      const errText = await response.text();
      return NextResponse.json({ error: "Errore dal servizio Gemini", details: errText }, { status: 500 });
    }

    const data = await response.json();
    const generatedText = data.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
    const pianoJSON = JSON.parse(generatedText);

    return NextResponse.json({ success: true, piano: pianoJSON });

  } catch (error: any) {
    return NextResponse.json({ error: "Errore interno del server", message: error.message }, { status: 500 });
  }
}