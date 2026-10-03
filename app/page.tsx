'use-effect';
import React, { useState } from 'react';

export default function AssoPlanDashboard() {
  const [isChatOpen, setIsChatOpen] = useState(false);
  const [chatInput, setChatInput] = useState('');
  const [messages, setMessages] = useState([
    { sender: 'ai', text: 'Ciao! Sono il tuo assistente operativo. Carica il file Excel dei cantieri e dammi qualsiasi istruzione per pianificare o modificare la settimana.' }
  ]);
  const [loading, setLoading] = useState(false);
  
  // Stato simulato dei cantieri / pianificazione attuale
  const [pianificazione, setPianificazione] = useState([
    { id: 1, titolo: '[INS] [DC] CONDOMINIO VIA VEZZANI 9A - BI 2593', localita: 'RIVAROLO', orario: '08:22 - 08:46', lavoro: '24 min', viaggio: '+22m da Sede', muletti: 0 },
    { id: 2, titolo: '[DC] GADO MED S.R.L. - BI 2351', localita: 'CORNIGLIANO', orario: '08:52 - 10:48', lavoro: '116 min', viaggio: '+17m rientro sede', muletti: -1 }
  ]);

  const handleSendMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!chatInput.trim() || loading) return;

    const userMessage = chatInput;
    setMessages(prev => [...prev, { sender: 'user', text: userMessage }]);
    setChatInput('');
    setLoading(true);

    try {
      const res = await fetch('/api/planner', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'modifica_chat',
          comandoUtente: userMessage,
          statoAttuale: pianificazione
        })
      });

      const data = await res.json();
      if (data.success && data.piano) {
        setMessages(prev => [...prev, { sender: 'ai', text: 'Fatto! Ho aggiornato la pianificazione secondo le tue indicazioni.' }]);
        if (data.piano.interventi) {
          setPianificazione(data.piano.interventi);
        }
      } else {
        setMessages(prev => [...prev, { sender: 'ai', text: 'Ho riscontrato un problema nell elaborare la richiesta. Riprova.' }]);
      }
    } catch (err) {
      setMessages(prev => [...prev, { sender: 'ai', text: 'Errore di connessione con il server di pianificazione.' }]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans">
      {/* Header Principale */}
      <header className="bg-slate-900 text-white px-6 py-4 flex justify-between items-center shadow-md">
        <div>
          <h1 className="text-xl font-bold tracking-wide">ASSO ANTINCENDIO — Gestionale Operativo</h1>
          <p className="text-xs text-slate-400">Pianificazione Intelligente & Logistica Territoriale (Genoa / Liguria)</p>
        </div>
        <div className="flex items-center gap-3">
          <button 
            onClick={() => setIsChatOpen(!isChatOpen)}
            className="bg-amber-600 hover:bg-amber-500 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-2 transition shadow"
          >
            💬 Assistente IA
          </button>
        </div>
      </header>

      {/* Contenuto Principale & Chat Drawer */}
      <div className="flex-1 flex relative overflow-hidden">
        {/* Tabella / Area Principale */}
        <main className="flex-1 p-6 overflow-y-auto">
          <div className="max-w-6xl mx-auto">
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 mb-6 flex justify-between items-center">
              <div>
                <h2 className="text-lg font-bold text-slate-800">Itinerario Tecnico: Nikolas Digiglio</h2>
                <p className="text-sm text-slate-500">Gestione automatica valli, turni e distinte carichi</p>
              </div>
              <div className="text-sm bg-blue-50 text-blue-700 px-3 py-1.5 rounded-lg font-medium border border-blue-200">
                📅 16/10/2026 (Venerdì)
              </div>
            </div>

            {/* Tabella Interventi */}
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="bg-slate-100 px-6 py-3 border-b border-slate-200 text-xs font-bold text-slate-600 uppercase tracking-wider grid grid-cols-12 gap-4">
                <div className="col-span-1">Giro</div>
                <div className="col-span-6">Riferimento Cliente & Indirizzo</div>
                <div className="col-span-3">Data e Orario</div>
                <div className="col-span-2 text-right">Dettagli</div>
              </div>

              <div className="divide-y divide-slate-100">
                {pianificazione.map((item, index) => (
                  <div key={item.id || index} className="px-6 py-4 grid grid-cols-12 gap-4 items-center hover:bg-slate-50/80 transition">
                    <div className="col-span-1 font-bold text-blue-600">
                      <span className="w-7 h-7 bg-blue-100 rounded-full flex items-center justify-center text-xs">{index + 1}</span>
                    </div>
                    <div className="col-span-6">
                      <p className="font-semibold text-slate-800 text-sm">{item.titolo}</p>
                      <p className="text-xs text-slate-500">{item.localita}</p>
                    </div>
                    <div className="col-span-3">
                      <p className="text-xs font-bold text-slate-700">🕒 {item.orario}</p>
                      <div className="flex gap-2 mt-1">
                        <span className="text-[10px] bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded border border-emerald-200">🛠️️ {item.lavoro} lavoro</span>
                        <span className="text-[10px] bg-purple-50 text-purple-700 px-2 py-0.5 rounded border border-purple-200">🚗 {item.viaggio}</span>
                      </div>
                    </div>
                    <div className="col-span-2 text-right">
                      <span className="text-xs font-semibold text-slate-600 bg-slate-100 px-2.5 py-1 rounded">Muletti: {item.muletti}</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </main>

        {/* Pannello Laterale Chat IA */}
        {isChatOpen && (
          <aside className="w-96 bg-white border-l border-slate-200 flex flex-col shadow-xl z-20">
            <div className="p-4 bg-slate-900 text-white flex justify-between items-center">
              <h3 className="font-bold text-sm flex items-center gap-2">✨ Assistente Operativo IA</h3>
              <button onClick={() => setIsChatOpen(false)} className="text-slate-400 hover:text-white text-lg font-bold">×</button>
            </div>

            <div className="flex-1 p-4 overflow-y-auto space-y-3 bg-slate-50">
              {messages.map((msg, idx) => (
                <div key={idx} className={`flex ${msg.sender === 'user' ? 'justify-end' : 'justify-start'}`}>
                  <div className={`max-w-[85%] rounded-2xl px-4 py-2.5 text-xs leading-relaxed ${
                    msg.sender === 'user' 
                      ? 'bg-blue-600 text-white rounded-br-none' 
                      : 'bg-white text-slate-800 border border-slate-200 shadow-sm rounded-bl-none'
                  }`}>
                    {msg.text}
                  </div>
                </div>
              ))}
              {loading && (
                <div className="flex justify-start">
                  <div className="bg-white text-slate-500 border border-slate-200 px-4 py-2 rounded-xl text-xs shadow-sm animate-pulse">
                    Sto elaborando il ragionamento logico...
                  </div>
                </div>
              )}
            </div>

            <form onSubmit={handleSendMessage} className="p-3 bg-white border-t border-slate-200 flex gap-2">
              <input 
                type="text"
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                placeholder="Es. Sposta Gado Med a mercoledì..."
                className="flex-1 border border-slate-300 rounded-lg px-3 py-2 text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
              <button 
                type="submit"
                disabled={loading}
                className="bg-blue-600 hover:bg-blue-500 text-white px-4 py-2 rounded-lg text-xs font-semibold transition disabled:opacity-50"
              >
                Invia
              </button>
            </form>
          </aside>
        )}
      </div>
    </div>
  );
}