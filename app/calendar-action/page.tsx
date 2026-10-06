"use client";

import { useEffect, useState, Suspense, useRef } from "react";
import { useSearchParams } from "next/navigation";
import { CheckCircle, Loader2, XCircle } from "lucide-react";

function ActionHandler() {
  const searchParams = useSearchParams();
  const [status, setStatus] = useState<"loading" | "success" | "error">("loading");
  const [message, setMessage] = useState("Cerco l'intervento su Google Calendar...");
  const [countdown, setCountdown] = useState(3);
  
  const hasFired = useRef(false);

  useEffect(() => {
    if (hasFired.current) return;
    hasFired.current = true;

    const executeAction = async () => {
      const bi = searchParams.get("bi");
      const techEmail = searchParams.get("tecnico");
      const action = searchParams.get("action");

      if (!bi || !techEmail || !action) {
        setStatus("error");
        setMessage("Parametri mancanti nel link.");
        return;
      }

      const storedGoogle = localStorage.getItem("asso_google_token");
      if (!storedGoogle) {
        setStatus("error");
        setMessage("Non sei connesso a Google Calendar. Apri la plancia principale ed effettua l'accesso.");
        return;
      }

      const tokenData = JSON.parse(storedGoogle);
      if (Date.now() > tokenData.expiry) {
         setStatus("error");
         setMessage("La sessione di Google è scaduta. Rifai l'accesso dalla plancia principale.");
         return;
      }
      const token = tokenData.token;

      try {
        setMessage("Individuazione cantiere in corso...");
        const searchUrl = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(techEmail)}/events?q=${encodeURIComponent(bi)}&timeMin=${new Date(Date.now() - 24*60*60*1000).toISOString()}&maxResults=50`;
        const resSearch = await fetch(searchUrl, { 
            headers: { Authorization: `Bearer ${token}` },
            cache: 'no-store'
        });
        const dataSearch = await resSearch.json();

        if (!dataSearch.items || dataSearch.items.length === 0) {
           setStatus("error");
           setMessage("Nessun cantiere trovato con questo numero BI su Calendar.");
           return;
        }

        const targetEvent = dataSearch.items.find((e: any) => e.summary && e.summary.includes("[DC]") && e.summary.includes(bi));

        if (!targetEvent) {
           setStatus("error");
           setMessage("Questo cantiere è già stato confermato (il tag [DC] non c'è più).");
           return;
        }

        setMessage("Aggiorno i parametri...");
        
        const newSummary = targetEvent.summary.replace(/\[DC\]\s*/g, "");
        const colorId = action === "email" ? "6" : null;

        const patchBody: any = { summary: newSummary };
        if (colorId) {
            patchBody.colorId = colorId;
        } else {
            patchBody.colorId = null; 
        }

        const resPatch = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(techEmail)}/events/${encodeURIComponent(targetEvent.id)}`, {
          method: 'PATCH',
          headers: { 
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(patchBody)
        });

        if (resPatch.ok) {
           setStatus("success");
           setMessage(`Tag rimosso! Cantiere colorato di ${action === 'email' ? 'ARANCIONE' : 'PREDEFINITO'}.`);
           
           // AVVIA LA CHIUSURA AUTOMATICA E IL CONTO ALLA ROVESCIA
           let counter = 3;
           const interval = setInterval(() => {
               counter -= 1;
               setCountdown(counter);
               if (counter <= 0) {
                   clearInterval(interval);
                   window.close(); // Chiude la scheda
               }
           }, 1000);

        } else {
           setStatus("error");
           setMessage("C'è stato un problema di connessione con Google Calendar.");
        }

      } catch (error) {
         setStatus("error");
         setMessage("Errore critico di rete.");
      }
    };

    executeAction();
  }, [searchParams]);

  return (
    <div className="flex flex-col items-center justify-center p-8 text-center">
        {status === "loading" && <Loader2 size={64} className="text-blue-500 animate-spin mb-4" />}
        {status === "success" && <CheckCircle size={64} className="text-emerald-500 mb-4" />}
        {status === "error" && <XCircle size={64} className="text-red-500 mb-4" />}
        
        <h1 className="text-2xl font-extrabold text-slate-800 mb-2">Sincronizzazione Asso</h1>
        <p className="text-slate-600 font-medium mb-4">{message}</p>

        {status === "success" && (
            <p className="text-sm font-bold text-emerald-600 bg-emerald-50 px-4 py-2 rounded-lg animate-pulse">
                Chiusura automatica in {countdown} secondi...
            </p>
        )}

        {status !== "loading" && (
            <button onClick={() => window.close()} className="mt-6 px-6 py-3 bg-slate-900 text-white rounded-xl font-bold hover:bg-slate-800 shadow-lg transition-all">
                Chiudi manualmente
            </button>
        )}
    </div>
  );
}

export default function CalendarActionPage() {
   return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center font-sans p-4">
         <div className="bg-white p-8 rounded-3xl shadow-xl max-w-md w-full border border-slate-200">
            <Suspense fallback={<div className="text-center font-bold text-slate-500">Avvio modulo...</div>}>
                <ActionHandler />
            </Suspense>
         </div>
      </div>
   );
}