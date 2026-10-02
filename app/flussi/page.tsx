"use client";
import { useState } from "react";

// Lista degli appaltatori e la regola di assegnazione automatica del coordinatore
const APPALTATORI_CONFIG = {
  "Pighi": "Porcile Roberto",
  "Icil": "Porcile Roberto",
  "Sistemi di Sicurezza": "Andrea Taucci",
  "Camst": "Andrea Taucci",
  "Nazca": "Andrea Taucci",
  "Fervo": "Andrea Taucci",
  "Asso Antincendio": "" // Da decidere manualmente
};

export default function DashboardFlussi() {
  // Dati di esempio (poi collegati a Supabase)
  const [richieste, setRichieste] = useState([
    {
      id_richiesta: "RIL-2026-45",
      data_inserimento: "01/10/2026 10:30",
      appaltatore: "Pighi",
      cliente: "Lidl",
      codice_filiale: "GE-04",
      filiale: "Via XX Settembre",
      tipo_presidio: "Porte Tagliafuoco",
      motivazione: "Cerniera anta bloccata",
      stato: "In Lavorazione",
      tecnico_assegnato: "S. Adinolfi",
      coordinatore: "Porcile Roberto"
    },
    {
      id_richiesta: "RIL-2026-44",
      data_inserimento: "01/10/2026 09:15",
      appaltatore: "Camst",
      cliente: "Esselunga",
      codice_filiale: "GE-02",
      filiale: "San Martino",
      tipo_presidio: "Estintori",
      motivazione: "Controllo periodico / Anomalia supporto",
      stato: "Nuova",
      tecnico_assegnato: "-",
      coordinatore: "Andrea Taucci"
    }
  ]);

  const [filtroStato, setFiltroStato] = useState("Tutti");
  const [ricerca, setRicerca] = useState("");

  // Conteggi KPI
  const countNuove = richieste.filter(r => r.stato === "Nuova").length;
  const countLavorazione = richieste.filter(r => r.stato === "In Lavorazione").length;
  const countChiusi = richieste.filter(r => r.stato === "Chiuso").length;

  return (
    <div className="min-h-screen bg-slate-50 p-6 font-sans text-slate-800">
      
      {/* HEADER */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900 flex items-center gap-2">
            <span>📊</span> Dashboard Ingressi & Flussi Intervento
          </h1>
          <p className="text-sm text-slate-500">Gestione centralizzata richieste, coordinamento e rientro contabile</p>
        </div>
        <div className="flex gap-3">
          <button className="px-4 py-2 bg-white border border-slate-200 rounded-xl text-sm font-medium hover:bg-slate-100 transition shadow-sm flex items-center gap-2">
            🔄 Sincronizza
          </button>
          <button className="px-4 py-2 bg-blue-600 text-white rounded-xl text-sm font-medium hover:bg-blue-700 transition shadow-sm flex items-center gap-2">
            ➕ Nuova Richiesta
          </button>
        </div>
      </div>

      {/* KPI CARDS (Colpo d'occhio) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
        <div onClick={() => setFiltroStato("Nuova")} className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm hover:border-blue-300 transition cursor-pointer">
          <div className="text-slate-500 text-xs font-semibold uppercase tracking-wider mb-1">📥 In Arrivo / Da Smistare</div>
          <div className="text-3xl font-extrabold text-blue-600">{countNuove}</div>
          <div className="text-xs text-slate-400 mt-2">Richieste in attesa di coordinatore</div>
        </div>

        <div onClick={() => setFiltroStato("In Lavorazione")} className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm hover:border-amber-300 transition cursor-pointer">
          <div className="text-slate-500 text-xs font-semibold uppercase tracking-wider mb-1">🔄 In Lavorazione</div>
          <div className="text-3xl font-extrabold text-amber-600">{countLavorazione}</div>
          <div className="text-xs text-slate-400 mt-2">Assegnati ai tecnici sul campo</div>
        </div>

        <div onClick={() => setFiltroStato("Chiuso")} className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm hover:border-emerald-300 transition cursor-pointer">
          <div className="text-slate-500 text-xs font-semibold uppercase tracking-wider mb-1">✅ Chiusi (Da Contabilizzare)</div>
          <div className="text-3xl font-extrabold text-emerald-600">{countChiusi}</div>
          <div className="text-xs text-slate-400 mt-2">Pronti per il rientro in ufficio</div>
        </div>

        <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm">
          <div className="text-slate-500 text-xs font-semibold uppercase tracking-wider mb-1">⚠️ Segnalazioni / Bloccati</div>
          <div className="text-3xl font-extrabold text-rose-600">0</div>
          <div className="text-xs text-slate-400 mt-2">Richiedono attenzione ufficio</div>
        </div>
      </div>

      {/* FILTRI & RICERCA */}
      <div className="bg-white p-4 rounded-2xl border border-slate-200/80 shadow-sm mb-6 flex flex-col md:flex-row gap-4 justify-between items-center">
        <input 
          type="text" 
          placeholder="🔍 Cerca cliente, filiale o protocollo..." 
          value={ricerca}
          onChange={(e) => setRicerca(e.target.value)}
          className="w-full md:w-96 px-4 py-2 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500"
        />
        <div className="flex gap-2 w-full md:w-auto overflow-x-auto">
          {["Tutti", "Nuova", "In Lavorazione", "Chiuso"].map((stato) => (
            <button
              key={stato}
              onClick={() => setFiltroStato(stato)}
              className={`px-4 py-2 rounded-xl text-sm font-medium transition whitespace-nowrap ${
                filtroStato === stato 
                  ? "bg-slate-900 text-white shadow-sm" 
                  : "bg-slate-100 text-slate-600 hover:bg-slate-200"
              }`}
            >
              {stato}
            </button>
          ))}
        </div>
      </div>

      {/* TABELLA PRINCIPALE */}
      <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-500 uppercase tracking-wider">
                <th className="p-4">Protocollo</th>
                <th className="p-4">Data / Ora</th>
                <th className="p-4">Appaltatore</th>
                <th className="p-4">Cliente / Filiale</th>
                <th className="p-4">Presidio & Motivo</th>
                <th className="p-4">Coordinatore</th>
                <th className="p-4">Stato</th>
                <th className="p-4">Tecnico</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-sm">
              {richieste.map((r, index) => (
                <tr key={index} className="hover:bg-slate-50/80 transition">
                  <td className="p-4 font-semibold text-slate-900">{r.id_richiesta}</td>
                  <td className="p-4 text-slate-500 text-xs">{r.data_inserimento}</td>
                  <td className="p-4">
                    <span className="px-2.5 py-1 bg-slate-100 text-slate-700 rounded-lg text-xs font-medium border border-slate-200">
                      {r.appaltatore}
                    </span>
                  </td>
                  <td className="p-4">
                    <div className="font-medium text-slate-900">{r.cliente}</div>
                    <div className="text-xs text-slate-500">{r.filiale} ({r.codice_filiale})</div>
                  </td>
                  <td className="p-4">
                    <div className="font-medium text-slate-800">{r.tipo_presidio}</div>
                    <div className="text-xs text-slate-500 truncate max-w-xs">{r.motivazione}</div>
                  </td>
                  <td className="p-4 text-xs font-medium text-slate-700">{r.coordinatore}</td>
                  <td className="p-4">
                    <span className={`px-2.5 py-1 rounded-full text-xs font-semibold inline-flex items-center gap-1.5 ${
                      r.stato === "Nuova" ? "bg-blue-50 text-blue-700 border border-blue-200" :
                      r.stato === "In Lavorazione" ? "bg-amber-50 text-amber-700 border border-amber-200" :
                      "bg-emerald-50 text-emerald-700 border border-emerald-200"
                    }`}>
                      <span className={`w-1.5 h-1.5 rounded-full ${
                        r.stato === "Nuova" ? "bg-blue-500" :
                        r.stato === "In Lavorazione" ? "bg-amber-500" :
                        "bg-emerald-500"
                      }`}></span>
                      {r.stato}
                    </span>
                  </td>
                  <td className="p-4 text-xs font-medium text-slate-600">{r.tecnico_assegnato}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

    </div>
  );
}