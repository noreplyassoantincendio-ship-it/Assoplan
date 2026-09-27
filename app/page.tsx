"use client";

import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import dynamic from "next/dynamic";
import { UploadCloud, AlertCircle, Package, MapPin, Loader2, FileSpreadsheet, MapPinOff, Zap, UserCheck, Printer, Calendar, Clock, CheckSquare, Square, Mail, Timer, FileText, X, Sliders, Check, Trash2, ArrowRight, PauseCircle, PlusCircle, ExternalLink, ShieldAlert, AlertTriangle, CheckCircle, PlayCircle, Lock, CalendarPlus, GripVertical } from "lucide-react";

const MapContainer = dynamic(() => import("react-leaflet").then((mod) => mod.MapContainer), { ssr: false });
const TileLayer = dynamic(() => import("react-leaflet").then((mod) => mod.TileLayer), { ssr: false });
const Marker = dynamic(() => import("react-leaflet").then((mod) => mod.Marker), { ssr: false });
const Popup = dynamic(() => import("react-leaflet").then((mod) => mod.Popup), { ssr: false });

const GOOGLE_CLIENT_ID = "645365149295-lk8ei43kt09hsm2csupf643tgqkqqgmo.apps.googleusercontent.com";
const GOOGLE_SCOPES = "https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/calendar.events";

export default function Home() {
  const [interventiGrezzi, setInterventiGrezzi] = useState<any[]>([]);
  const [clientiInSospeso, setClientiInSospeso] = useState<any[]>([]);
  const [clientiGiaCalendarizzati, setClientiGiaCalendarizzati] = useState<any[]>([]);
  const [inElaborazione, setInElaborazione] = useState(false);
  const [leafletLoaded, setLeafletLoaded] = useState(false);
  
  const [isCheckingAuth, setIsCheckingAuth] = useState(true);
  const [googleToken, setGoogleToken] = useState<string | null>(null);
  const tokenClientRef = useRef<any>(null);

  const [datiGrezziCaricati, setDatiGrezziCaricati] = useState<any[]>([]);
  const [nomeFileCorrente, setNomeFileCorrente] = useState<string>("");

  const [dataInizio, setDataInizio] = useState<string>("2026-09-28");
  const [dataFine, setDataFine] = useState<string>("2026-10-02");
  const [giorniAttivi, setGiorniAttivi] = useState<{ [key: string]: boolean }>({
    "Lunedì": true, "Martedì": true, "Mercoledì": true, "Giovedì": true, "Venerdì": true
  });

  const [giornoSceltoFiltro, setGiornoSceltoFiltro] = useState<string>("Tutti");
  const [clienteSelezionatoScheda, setClienteSelezionatoScheda] = useState<any | null>(null);
  const [configSospesi, setConfigSospesi] = useState<{ [key: string]: { data: string, oraInizio: string } }>({});
  
  const [giornateStats, setGiornateStats] = useState<any[]>([]);
  const [creazioneInCorso, setCreazioneInCorso] = useState<{[key:string]: boolean}>({});

  const [bulkSyncStatus, setBulkSyncStatus] = useState({ active: false, progress: 0, current: 0, total: 0 });

  const tecniciAnagrafica = [
    { nome: "Fabrizio Vercellino", email: "tecnici08.assoantincendio@gmail.com" },
    { nome: "Nikolas Digiglio", email: "tecnici04.assoantincendio@gmail.com" },
    { nome: "Porcile Roberto", email: "porcile.assoantincendio@gmail.com" },
    { nome: "Daniel Rossi", email: "tecnici10.assoantincendio@gmail.com" },
    { nome: "Andrea Taucci", email: "tecnici05.assoantincendio@gmail.com" },
    { nome: "Pietro Adinolfi", email: "tecnici09.assoantincendio@gmail.com" },
    { nome: "Salvatore Russo", email: "tecnici03.assoantincendio@gmail.com" },
    { nome: "Mario Licata", email: "tecnici12.assoantincendio@gmail.com" },
    { nome: "Samuele Adinolfi", email: "tecnici07.assoantincendio@gmail.com" }
  ];

  const [tecnicoSelezionato, setTecnicoSelezionato] = useState(tecniciAnagrafica[0]);

  useEffect(() => {
    import("leaflet").then((L) => {
      // @ts-ignore
      delete L.Icon.Default.prototype._getIconUrl;
      L.Icon.Default.mergeOptions({
        iconRetinaUrl: "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-icon-2x.png",
        iconUrl: "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-icon.png",
        shadowUrl: "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-shadow.png",
      });
      setLeafletLoaded(true);
    });
  }, []);

  useEffect(() => {
    const stored = localStorage.getItem("asso_google_token");
    if (stored) {
      try {
        const parsed = JSON.parse(stored);
        if (Date.now() < parsed.expiry) {
          setGoogleToken(parsed.token);
        } else {
          localStorage.removeItem("asso_google_token");
        }
      } catch (e) {
        localStorage.removeItem("asso_google_token");
      }
    }
  }, []);

  useEffect(() => {
    const initGoogleClient = () => {
      // @ts-ignore
      if (window.google && window.google.accounts) {
        // @ts-ignore
        tokenClientRef.current = window.google.accounts.oauth2.initTokenClient({
          client_id: GOOGLE_CLIENT_ID,
          scope: GOOGLE_SCOPES,
          callback: (response: any) => {
            if (response.error) {
              console.error("Errore OAuth:", response);
              alert("Impossibile connettersi a Google: " + response.error);
              return;
            }
            if (response.access_token) {
              setGoogleToken(response.access_token);
              localStorage.setItem("asso_google_token", JSON.stringify({
                token: response.access_token,
                expiry: Date.now() + 3300 * 1000 
              }));
            }
          },
        });
        setIsCheckingAuth(false);
      }
    };

    // @ts-ignore
    if (window.google) {
      initGoogleClient();
    } else {
      const script = document.createElement("script");
      script.src = "https://accounts.google.com/gsi/client";
      script.async = true;
      script.defer = true;
      script.onload = initGoogleClient;
      document.head.appendChild(script);
    }
  }, []);

  const eseguiLoginGoogle = () => {
    if (tokenClientRef.current) {
      tokenClientRef.current.requestAccessToken();
    } else {
      alert("Il sistema di connessione Google si sta caricando. Attendi qualche secondo e riprova.");
    }
  };

  const logoutGoogle = () => {
    setGoogleToken(null);
    localStorage.removeItem("asso_google_token");
  };

  useEffect(() => {
    if (clienteSelezionatoScheda) document.body.style.overflow = "hidden";
    else document.body.style.overflow = "unset";
    return () => { document.body.style.overflow = "unset"; };
  }, [clienteSelezionatoScheda]);

  const GOOGLE_MAPS_API_KEY = "AIzaSyBLfAmnm_kaHbUc0sAVzhvkwXDF14EFCro";
  const SEDE_UFFICIO_LAT = 44.4056;
  const SEDE_UFFICIO_LNG = 8.9463;

  const calcolaDistanzaKm = (lat1: number, lon1: number, lat2: number, lon2: number) => {
    const R = 6371;
    const dLat = (lat2 - lat1) * (Math.PI / 180);
    const dLon = (lon2 - lon1) * (Math.PI / 180);
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(lat1 * (Math.PI / 180)) * Math.cos(lat2 * (Math.PI / 180)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
    return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
  };

  const trovaCoordinateGoogle = async (indirizzo: string, localita: string) => {
    let cleanAddr = indirizzo.replace(/Nessun indirizzo valido letto/gi, '').trim();
    const citta = localita && localita.trim().length > 0 ? localita.trim() : "Genova";
    const query = `${cleanAddr}, ${citta}, Italy`;

    const cacheChiave = query.toLowerCase();
    const cacheSalvata = localStorage.getItem("asso_mappe_cache");
    const memoriaFissa = cacheSalvata ? JSON.parse(cacheSalvata) : {};

    if (memoriaFissa[cacheChiave]) return { ...memoriaFissa[cacheChiave], origine: "memoria" };
    if (!GOOGLE_MAPS_API_KEY) return null;

    try {
      const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(query)}&key=${GOOGLE_MAPS_API_KEY}`);
      const data = await response.json();
      if (data.status === "OK" && data.results.length > 0) {
        const location = data.results[0].geometry.location;
        const risultato = { lat: location.lat, lng: location.lng, precisione: "strada", origine: "google" };
        memoriaFissa[cacheChiave] = risultato;
        localStorage.setItem("asso_mappe_cache", JSON.stringify(memoriaFissa));
        return risultato;
      }
    } catch (error) {}
    return null;
  };

  const getNomeGiorno = (dataStr: string) => {
    const [y, m, d] = dataStr.split("-").map(Number);
    const date = new Date(y, m - 1, d);
    const mapGiorni = ["Domenica", "Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato"];
    return mapGiorni[date.getDay()];
  };

  const getGiornateLavorative = (inizio: string, fine: string, attivi: { [key: string]: boolean }) => {
    const dates = [];
    const [yI, mI, dI] = inizio.split("-").map(Number);
    const [yF, mF, dF] = fine.split("-").map(Number);
    
    let curr = new Date(yI, mI - 1, dI);
    const end = new Date(yF, mF - 1, dF);
    const mapGiorni = ["Domenica", "Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato"];
    
    while (curr <= end) {
      const nomeGiorno = mapGiorni[curr.getDay()];
      if (attivi[nomeGiorno]) {
        const y = curr.getFullYear();
        const m = String(curr.getMonth() + 1).padStart(2, '0');
        const d = String(curr.getDate()).padStart(2, '0');
        dates.push({ dataStr: `${y}-${m}-${d}`, nomeGiorno });
      }
      curr.setDate(curr.getDate() + 1);
    }
    return dates;
  };

  const fetchEventiRealiCalendar = async (emailTecnico: string, token: string) => {
    try {
      const [yI, mI, dI] = dataInizio.split("-").map(Number);
      const dataMin = new Date(yI, mI - 1, dI - 10);
      const [yF, mF, dF] = dataFine.split("-").map(Number);
      const dataMax = new Date(yF, mF - 1, dF + 20);

      const timeMin = dataMin.toISOString();
      const timeMax = dataMax.toISOString();
      const calendarId = encodeURIComponent(emailTecnico);
      const url = `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events?timeMin=${timeMin}&timeMax=${timeMax}&singleEvents=true&orderBy=startTime&maxResults=2500`;

      const response = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: "application/json" } });

      if (!response.ok) {
        if (response.status === 401) { logoutGoogle(); alert("Token di Google scaduto. Clicca nuovamente su 'Connetti Google Calendar'."); }
        return [];
      }

      const data = await response.json();
      const eventi = [];

      for (const evt of (data.items || [])) {
        const testoCompleto = `${evt.summary || ""} ${evt.description || ""}`.toUpperCase();
        const matches = testoCompleto.match(/(?:BI[:\s]*)([A-Z0-9\/\-]+)/g);
        const biGiaLetti = matches ? matches.map(m => m.replace(/BI[:\s]*/g, "").trim()) : [];
        
        let latEvt = null; let lngEvt = null;
        if (evt.location && evt.location.trim().length > 2) {
          const coord = await trovaCoordinateGoogle(evt.location, "Genova");
          if (coord) { latEvt = coord.lat; lngEvt = coord.lng; }
        }

        if (evt.start?.date && evt.end?.date) {
            const startD = new Date(evt.start.date + "T00:00:00");
            const endD = new Date(evt.end.date + "T00:00:00"); 
            let currentD = new Date(startD);
            while (currentD < endD) {
                const y = currentD.getFullYear();
                const m = String(currentD.getMonth() + 1).padStart(2, '0');
                const d = String(currentD.getDate()).padStart(2, '0');
                const dateStr = `${y}-${m}-${d}`;

                eventi.push({
                  id: evt.id + "_" + dateStr,
                  dataStr: dateStr,
                  lat: latEvt,
                  lng: lngEvt,
                  durataMinuti: 480, 
                  numeriBiRilevati: biGiaLetti,
                });
                currentD.setDate(currentD.getDate() + 1);
            }
        } else if (evt.start?.dateTime && evt.end?.dateTime) {
            const t1 = new Date(evt.start.dateTime).getTime();
            const t2 = new Date(evt.end.dateTime).getTime();
            const startD = new Date(evt.start.dateTime);
            const y = startD.getFullYear();
            const m = String(startD.getMonth() + 1).padStart(2, '0');
            const d = String(startD.getDate()).padStart(2, '0');
            const dateStr = `${y}-${m}-${d}`;

            eventi.push({
              id: evt.id,
              dataStr: dateStr,
              lat: latEvt,
              lng: lngEvt,
              durataMinuti: Math.round((t2 - t1) / 60000),
              numeriBiRilevati: biGiaLetti,
            });
        }
      }
      return eventi;
    } catch (err) { return []; }
  };

  const generaTitoloEvento = (cliente: any) => {
    let tags = [];
    if (cliente.haInsoluto) tags.push("[INS]");
    if (cliente.resoEstintori) tags.push("[SOS]");
    if (!cliente.nome.toUpperCase().includes("ASL")) tags.push("[DC]");
    
    const tagString = tags.length > 0 ? tags.join(" ") + " " : "";
    return `${tagString}${cliente.nome} - BI ${cliente.numeroBi}`;
  };

  const generaDescrizioneEvento = (cliente: any, tecnicoNome: string) => {
    let attrezzatureTesto = cliente.dettaglioAttrezzature && cliente.dettaglioAttrezzature.length > 0 
      ? cliente.dettaglioAttrezzature.map((a: any) => `🔧 ${a.descrizione}: ${a.quantita}`).join("\n")
      : "Nessuna attrezzatura specificata";

    const telInfo = cliente.telefono ? `📞 ${cliente.telefono}` : "";
    const emailInfo = cliente.email ? `📧 Email: ${cliente.email}` : "";
    const contatti = (telInfo || emailInfo) ? `\n${telInfo}\n${emailInfo}\n` : "\n";

    return `Località: ${cliente.localita || ''}
Indirizzo: ${cliente.indirizzo || ''}
${contatti}
🛠 ATTREZZATURE DA VERIFICARE:
${attrezzatureTesto}

Durata: ${cliente.minutiLavoro} min
📍 Naviga:
https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${cliente.indirizzo}, ${cliente.localita || 'Genova'}, Italy`)}

Tecnico: ${tecnicoNome}`;
  };

  const processaRighe = async (righe: any[], nomeFileCaricato: string = "") => {
    if (!googleToken) { alert("Connetti l'account Google per elaborare i dati."); return; }
    if (righe.length === 0) { alert("Nessun dato caricato. Seleziona prima un file."); return; }

    setInElaborazione(true);
    const clientiMappa = new Map();
    const nomeFileLower = (nomeFileCaricato || nomeFileCorrente).toLowerCase();
    
    let tecnicoTrovato = tecniciAnagrafica[0];
    for (const tech of tecniciAnagrafica) {
      if (tech.nome.toLowerCase().split(" ").some(p => p.length > 2 && nomeFileLower.includes(p))) {
        tecnicoTrovato = tech; break;
      }
    }
    setTecnicoSelezionato(tecnicoTrovato);

    const impegniEsistenti = await fetchEventiRealiCalendar(tecnicoTrovato.email, googleToken);
    let tuttiIBiGiaInCalendario: string[] = [];
    impegniEsistenti.forEach(evt => { tuttiIBiGiaInCalendario = [...tuttiIBiGiaInCalendario, ...evt.numeriBiRilevati]; });

    for (const riga of righe) {
      const nomeCliente = (riga.desclifor || "").toUpperCase();
      if (!nomeCliente) continue;

      const numeroBi = String(riga.numdoc || riga.nrob || riga.id_bi || riga.documento || riga.codicebi || "BI-GENERICA").trim();
      const idCliente = riga.codclifor || riga.desclifor || Math.random().toString();

      if (!clientiMappa.has(idCliente)) {
        clientiMappa.set(idCliente, {
          codice: idCliente, numeroBi, nome: riga.desclifor || "Sconosciuto",
          indirizzo: riga.indirizzo || riga.INDIRIZZO || "", localita: riga.localita || riga.LOCALITA || "",
          telefono: riga.telefono || riga.TELEFONO || riga.tel || riga.TEL || "",
          email: riga.email || riga.EMAIL || riga['e-mail'] || "",
          haInsoluto: false, resoEstintori: false, totaleArticoli: 0, minutiLavoro: 20, dettaglioAttrezzature: [], selezionatoPerGiro: true, giorniDallUltimoSos: 8, syncedToGoogle: false 
        });
      }

      const cliente = clientiMappa.get(idCliente);
      const testoRiga = JSON.stringify(riga).toUpperCase();
      const numFattura = riga.numft ? String(riga.numft).trim() : "";
      if (numFattura !== "" && numFattura !== "-" && (parseFloat(riga.imptotft) || 0) > 0 && !(riga.codpagam || "").toUpperCase().trim().startsWith("RB")) cliente.haInsoluto = true;
      if ((riga.codclose ? String(riga.codclose).toUpperCase().trim() : "") === "SOS" || testoRiga.includes("SOS")) cliente.resoEstintori = true;

      let minutiTotaliRiga = 0; let pezziTotaliRiga = 0; const dettaglio: any[] = [];
      if (cliente.resoEstintori) {
        const eoRv = Number(riga.eo_rv) || 0;
        if (eoRv > 0) { minutiTotaliRiga += eoRv*2; pezziTotaliRiga += eoRv; dettaglio.push({ descrizione: "Revisione Estintori (SOS)", quantita: eoRv, categoria: "Rev (2m)", minutiImpegno: eoRv*2 }); }
      } else {
        const items = [
          { v: Number(riga.eo_mt)||0, m: 2, n: "Estintori (Manutenzione)" }, { v: Number(riga.id_mt)||0, m: 2, n: "Idranti" },
          { v: Number(riga.po_mt)||0, m: 3, n: "Porte Tagliafuoco / REI" }, { v: Number(riga.ss_mt)||0, m: 2, n: "Uscite di Emergenza" },
          { v: Number(riga.la_mt)||0, m: 1, n: "Lampade di Emergenza" }, { v: Number(riga.ir_mt)||0, m: 30, n: "Impianto Rilevazione (IRAI)" },
          { v: Number(riga.ri_mt)||0, m: 15, n: "Rete Idrica" }, { v: Number(riga.gp_mt)||0, m: 20, n: "Gruppo Pressurizzazione" }
        ];
        items.forEach(it => { if(it.v > 0){ minutiTotaliRiga += it.v*it.m; pezziTotaliRiga += it.v; dettaglio.push({ descrizione: it.n, quantita: it.v, categoria: `(${it.m}m/pz)`, minutiImpegno: it.v*it.m }); } });
      }
      if (minutiTotaliRiga > 0 || cliente.minutiLavoro === 20) { cliente.totaleArticoli = pezziTotaliRiga; cliente.minutiLavoro = 20 + minutiTotaliRiga; cliente.dettaglioAttrezzature = dettaglio; }
    }

    const datiConCoordinate = [];
    for (const cliente of Array.from(clientiMappa.values())) {
      let coordinate = null;
      if (cliente.indirizzo?.trim().length > 2) coordinate = await trovaCoordinateGoogle(cliente.indirizzo, cliente.localita);
      datiConCoordinate.push({ ...cliente, lat: coordinate?.lat || 44.4056, lng: coordinate?.lng || 8.9463 });
    }

    const clientiDaPianificare = [];
    const clientiSospesiTmp = [];
    const clientiScartatiTmp = [];

    for (let i = 0; i < datiConCoordinate.length; i++) {
      const c = datiConCoordinate[i];
      const biCliente = c.numeroBi.toUpperCase();
      
      if (biCliente !== "" && biCliente !== "BI-GENERICA" && tuttiIBiGiaInCalendario.includes(biCliente)) {
        clientiScartatiTmp.push(c); continue;
      }
      if (c.resoEstintori && c.giorniDallUltimoSos < 5) { clientiSospesiTmp.push({ ...c, motivoSospeso: `SOS Bloccato (<5 gg)` }); continue; }
      
      const txt = `${c.indirizzo} ${c.localita} ${c.nome}`.toUpperCase();
      if (txt.match(/RONCO SCRIVIA|BUSALLA|SAVIGNONE|ISOLA DEL CANTONE|MELE|MASONE|CAMPO LIGURE/)) { clientiSospesiTmp.push({ ...c, motivoSospeso: "Entroterra / Fuori Zona" }); continue; }

      let minDistanzaAltra = Infinity;
      for (let j = 0; j < datiConCoordinate.length; j++) {
        if (i === j) continue;
        const dist = calcolaDistanzaKm(c.lat, c.lng, datiConCoordinate[j].lat, datiConCoordinate[j].lng);
        if (dist < minDistanzaAltra) minDistanzaAltra = dist;
      }
      if (minDistanzaAltra > 15 && datiConCoordinate.length > 3) { clientiSospesiTmp.push({ ...c, motivoSospeso: `Isolato (${minDistanzaAltra.toFixed(1)} km)` }); } 
      else { clientiDaPianificare.push(c); }
    }

    const dateRange = getGiornateLavorative(dataInizio, dataFine, giorniAttivi);
    const giornateStrutturate = dateRange.map(g => {
      const impegniDelGiorno = impegniEsistenti.filter(e => e.dataStr === g.dataStr);
      let minutiOccupati = 0; let anchorLat = null; let anchorLng = null;
      impegniDelGiorno.forEach(imp => {
        minutiOccupati += imp.durataMinuti;
        if (imp.lat && imp.lng && !anchorLat) { anchorLat = imp.lat; anchorLng = imp.lng; }
      });
      let minutiDisponibili = Math.max(0, 510 - minutiOccupati);
      let statoGiornata = "APERTO";
      if (minutiOccupati >= 240 || minutiDisponibili < 150) { minutiDisponibili = 0; statoGiornata = "PIENO (Sospeso)"; }

      return { ...g, minutiOccupati, minutiDisponibili, statoGiornata, anchorLat, anchorLng };
    });

    setGiornateStats(giornateStrutturate);
    const clientiPianificati = [];

    for (const giornata of giornateStrutturate) {
      if (giornata.minutiDisponibili <= 0) continue; 
      let ultimaCoordGiorno = giornata.anchorLat ? { lat: giornata.anchorLat, lng: giornata.anchorLng } : null;
      let aggiuntoQualcuno = true;
      let orarioCorrenteMinuti = 8 * 60 + 30; // Partenza alle 08:30

      while (aggiuntoQualcuno && clientiDaPianificare.length > 0) {
        aggiuntoQualcuno = false;
        let bestIdx = -1; let minDistanza = Infinity;
        const puntoPartenza = ultimaCoordGiorno || { lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG };

        for (let i = 0; i < clientiDaPianificare.length; i++) {
          const c = clientiDaPianificare[i];
          const dist = calcolaDistanzaKm(puntoPartenza.lat, puntoPartenza.lng, c.lat, c.lng);
          if (dist < minDistanza) { minDistanza = dist; bestIdx = i; }
        }

        if (bestIdx !== -1) {
          const candidato = clientiDaPianificare[bestIdx];
          const minViaggio = Math.round(minDistanza * 2);
          const minRientro = Math.round(calcolaDistanzaKm(candidato.lat, candidato.lng, SEDE_UFFICIO_LAT, SEDE_UFFICIO_LNG) * 2);
          const impegnoReale = minViaggio + candidato.minutiLavoro;

          if ((giornata.minutiDisponibili - impegnoReale) >= minRientro) {
            
            const inizioLavoroMin = orarioCorrenteMinuti + minViaggio;
            const fineLavoroMin = inizioLavoroMin + candidato.minutiLavoro;
            
            const hhStart = String(Math.floor(inizioLavoroMin / 60)).padStart(2, '0');
            const mmStart = String(inizioLavoroMin % 60).padStart(2, '0');
            const hhEnd = String(Math.floor(fineLavoroMin / 60)).padStart(2, '0');
            const mmEnd = String(fineLavoroMin % 60).padStart(2, '0');

            clientiPianificati.push({ 
                ...candidato, 
                settimana: `${dataInizio} al ${dataFine}`, 
                giorno: giornata.nomeGiorno, 
                dataAssegnata: giornata.dataStr,
                oraInizio: `${hhStart}:${mmStart}`,
                oraFine: `${hhEnd}:${mmEnd}`
            });

            giornata.minutiDisponibili -= impegnoReale;
            ultimaCoordGiorno = { lat: candidato.lat, lng: candidato.lng };
            clientiDaPianificare.splice(bestIdx, 1);
            orarioCorrenteMinuti = fineLavoroMin;
            aggiuntoQualcuno = true;
          } else { break; }
        }
      }
    }

    clientiDaPianificare.forEach(c => { clientiSospesiTmp.push({ ...c, motivoSospeso: "Tempo esaurito: i giorni selezionati sono pieni" }); });
    
    setInterventiGrezzi(clientiPianificati);
    setClientiInSospeso(clientiSospesiTmp);
    setClientiGiaCalendarizzati(clientiScartatiTmp);
    setInElaborazione(false);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!googleToken) { alert("Devi connettere Google Calendar prima di caricare il file."); e.target.value = ''; return; }
    const file = e.target.files?.[0];
    if (!file) return;

    setNomeFileCorrente(file.name);

    if (file.name.toLowerCase().endsWith('.txt') || file.name.toLowerCase().endsWith('.csv')) {
      Papa.parse(file, { header: true, skipEmptyLines: true, dynamicTyping: true, complete: (res) => { setDatiGrezziCaricati(res.data); }});
    } else {
      const reader = new FileReader();
      reader.onload = async (evento) => {
        const data = evento.target?.result as ArrayBuffer;
        try {
          const arr = new Uint8Array(data);
          const wb = XLSX.read(arr, { type: "array" });
          const righe = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
          if (righe.length > 0) { setDatiGrezziCaricati(righe); return; }
        } catch (errExcel) {}
        try {
          const testo = new TextDecoder("windows-1252").decode(data);
          if (testo.toLowerCase().includes("<table")) {
            const parser = new DOMParser(); const doc = parser.parseFromString(testo, "text/html"); const table = doc.querySelector("table");
            if (table) { const workbookHTML = XLSX.utils.table_to_book(table); setDatiGrezziCaricati(XLSX.utils.sheet_to_json(workbookHTML.Sheets[workbookHTML.SheetNames[0]])); return; }
          }
          alert("Formato non supportato.");
        } catch (errFinal) { alert("Impossibile leggere il file."); }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const toggleGiornoAttivo = (giorno: string) => { setGiorniAttivi(prev => ({ ...prev, [giorno]: !prev[giorno] })); };
  const toggleSelezioneCliente = (indexGlobale: number) => { const nuovi = [...interventiGrezzi]; nuovi[indexGlobale].selezionatoPerGiro = !nuovi[indexGlobale].selezionatoPerGiro; setInterventiGrezzi(nuovi); };
  const calcolaDataDalGiorno = (giornoSelezionato: string) => { return calcolaDataDalGiornoSicura(giornoSelezionato, dataInizio); };

  const aggiornaConfigSospeso = (codice: string, campo: string, valore: string) => {
    setConfigSospesi(prev => {
      const current = prev[codice] || { data: dataInizio, oraInizio: "08:30" };
      return { ...prev, [codice]: { ...current, [campo]: valore } };
    });
  };

  const creaSuCalendarESposta = async (clienteSospeso: any, conf: any) => {
    if (clienteSospeso.resoEstintori && clienteSospeso.giorniDallUltimoSos < 5) { alert(`Impossibile ripianificare l'intervento SOS: sono passati solo ${clienteSospeso.giorniDallUltimoSos} giorni lavorativi.`); return; }
    if (!googleToken) { alert("Devi connettere Google Calendar!"); return; }

    setCreazioneInCorso(prev => ({...prev, [clienteSospeso.codice]: true}));

    try {
      const startDateTime = new Date(`${conf.data}T${conf.oraInizio}:00`);
      const endDateTime = new Date(startDateTime.getTime() + clienteSospeso.minutiLavoro * 60000);
      const hhEnd = String(endDateTime.getHours()).padStart(2, '0');
      const mmEnd = String(endDateTime.getMinutes()).padStart(2, '0');

      const event = {
        summary: generaTitoloEvento(clienteSospeso),
        location: `${clienteSospeso.indirizzo}, ${clienteSospeso.localita || 'Genova'}, Italia`,
        description: generaDescrizioneEvento(clienteSospeso, tecnicoSelezionato.nome),
        start: { dateTime: startDateTime.toISOString(), timeZone: 'Europe/Rome' },
        end: { dateTime: endDateTime.toISOString(), timeZone: 'Europe/Rome' },
      };

      const res = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(tecnicoSelezionato.email)}/events`, {
        method: 'POST', headers: { 'Authorization': `Bearer ${googleToken}`, 'Content-Type': 'application/json' }, body: JSON.stringify(event)
      });

      if (res.ok) {
        const giornoScelto = getNomeGiorno(conf.data);
        setInterventiGrezzi(prev => [...prev, { ...clienteSospeso, settimana: `${dataInizio} al ${dataFine}`, giorno: giornoScelto, dataAssegnata: conf.data, oraInizio: conf.oraInizio, oraFine: `${hhEnd}:${mmEnd}`, selezionatoPerGiro: true, syncedToGoogle: true }]);
        setClientiInSospeso(prev => prev.filter(c => c.codice !== clienteSospeso.codice));
      } else { alert("Errore durante la creazione API su Google."); }
    } catch (e) { alert("Errore di rete."); } 
    finally { setCreazioneInCorso(prev => ({...prev, [clienteSospeso.codice]: false})); }
  };

  const inviaPianificazioneAGoogle = async () => {
    if (!googleToken) { alert("Devi connettere Google Calendar!"); return; }
    const itemsToSync = interventiGrezzi.filter(i => i.selezionatoPerGiro && !i.syncedToGoogle);
    if (itemsToSync.length === 0) { alert("Tutti i cantieri selezionati sono già stati inviati o non c'è nulla da inviare."); return; }

    setBulkSyncStatus({ active: true, progress: 0, current: 0, total: itemsToSync.length });

    let syncedCount = 0;
    const nuoviInterventiGrezzi = [...interventiGrezzi];

    for (const cliente of itemsToSync) {
        const startDateTime = new Date(`${cliente.dataAssegnata}T${cliente.oraInizio}:00`);
        const endDateTime = new Date(`${cliente.dataAssegnata}T${cliente.oraFine}:00`);

        const event = {
            summary: generaTitoloEvento(cliente),
            location: `${cliente.indirizzo}, ${cliente.localita || 'Genova'}, Italia`,
            description: generaDescrizioneEvento(cliente, tecnicoSelezionato.nome),
            start: { dateTime: startDateTime.toISOString(), timeZone: 'Europe/Rome' },
            end: { dateTime: endDateTime.toISOString(), timeZone: 'Europe/Rome' },
        };

        try {
            const res = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(tecnicoSelezionato.email)}/events`, {
                method: 'POST', headers: { 'Authorization': `Bearer ${googleToken}`, 'Content-Type': 'application/json' }, body: JSON.stringify(event)
            });
            if (res.ok) {
                const idx = nuoviInterventiGrezzi.findIndex(i => i.codice === cliente.codice);
                if (idx !== -1) nuoviInterventiGrezzi[idx].syncedToGoogle = true;
            }
        } catch (e) { console.error("Errore sync", e); }

        syncedCount++;
        setBulkSyncStatus({ active: true, progress: Math.round((syncedCount / itemsToSync.length) * 100), current: syncedCount, total: itemsToSync.length });
        setInterventiGrezzi([...nuoviInterventiGrezzi]); 
    }

    setTimeout(() => setBulkSyncStatus({ active: false, progress: 0, current: 0, total: 0 }), 2000); 
  };

  // ---- GESTIONE DRAG & DROP PER RIORDINARE LA LISTA PIANIFICATI ----
  
  const handleDragStartReorder = (e: React.DragEvent, codice: string) => {
    e.dataTransfer.setData("application/reorder-lista", codice);
  };

  const handleDragOverReorder = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleDropReorder = (e: React.DragEvent, targetCodice: string) => {
    e.preventDefault();
    const draggedCodice = e.dataTransfer.getData("application/reorder-lista");
    if (!draggedCodice || draggedCodice === targetCodice) return;

    setInterventiGrezzi(prev => {
      const result = Array.from(prev);
      const draggedIndex = result.findIndex(i => i.codice === draggedCodice);
      const targetIndex = result.findIndex(i => i.codice === targetCodice);

      if (draggedIndex === -1 || targetIndex === -1) return prev;

      // Sposta l'elemento nell'array
      const [removed] = result.splice(draggedIndex, 1);
      result.splice(targetIndex, 0, removed);

      // --- Ricalcolo Matematico degli Orari per le giornate interessate ---
      const giorniPresenti = [...new Set(result.map(i => i.dataAssegnata))];

      giorniPresenti.forEach(dataStr => {
         const itemsGiorno = result.filter(i => i.dataAssegnata === dataStr);
         let orarioCorrenteMinuti = 8 * 60 + 30; // Riparte dalle 08:30 per ri-scandire la giornata
         
         let ultimaCoordGiorno = null;
         const stat = giornateStats.find(g => g.dataStr === dataStr);
         if (stat && stat.anchorLat) {
             ultimaCoordGiorno = { lat: stat.anchorLat, lng: stat.anchorLng };
         }

         itemsGiorno.forEach(candidato => {
             const puntoPartenza = ultimaCoordGiorno || { lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG };
             const dist = calcolaDistanzaKm(puntoPartenza.lat, puntoPartenza.lng, candidato.lat, candidato.lng);
             const minViaggio = Math.round(dist * 2);

             const inizioLavoroMin = orarioCorrenteMinuti + minViaggio;
             const fineLavoroMin = inizioLavoroMin + candidato.minutiLavoro;

             const hhStart = String(Math.floor(inizioLavoroMin / 60)).padStart(2, '0');
             const mmStart = String(inizioLavoroMin % 60).padStart(2, '0');
             const hhEnd = String(Math.floor(fineLavoroMin / 60)).padStart(2, '0');
             const mmEnd = String(fineLavoroMin % 60).padStart(2, '0');

             candidato.oraInizio = `${hhStart}:${mmStart}`;
             candidato.oraFine = `${hhEnd}:${mmEnd}`;

             orarioCorrenteMinuti = fineLavoroMin;
             ultimaCoordGiorno = { lat: candidato.lat, lng: candidato.lng };
         });
      });

      return result;
    });
  };

  const interventiVisibili = interventiGrezzi.filter(i => giornoSceltoFiltro === "Tutti" || i.giorno === giornoSceltoFiltro);
  const clientiAttivi = interventiVisibili.filter(i => i.selezionatoPerGiro);
  const clientiAttiviCount = clientiAttivi.length;
  const minutiLavoroTotali = clientiAttivi.reduce((acc, curr) => acc + curr.minutiLavoro, 0);
  
  let minutiTrasferimentoTotali = 0; let minutiRientroUfficio = 0;
  for (let i = 0; i < clientiAttivi.length - 1; i++) {
    if (clientiAttivi[i].lat && clientiAttivi[i+1].lat) { minutiTrasferimentoTotali += Math.round(calcolaDistanzaKm(clientiAttivi[i].lat, clientiAttivi[i].lng, clientiAttivi[i+1].lat, clientiAttivi[i+1].lng) * 2); }
  }
  if (clientiAttivi.length > 0 && clientiAttivi[clientiAttivi.length - 1].lat) {
    minutiRientroUfficio = Math.round(calcolaDistanzaKm(clientiAttivi[clientiAttivi.length - 1].lat, clientiAttivi[clientiAttivi.length - 1].lng, SEDE_UFFICIO_LAT, SEDE_UFFICIO_LNG) * 2);
  }

  const minutiTotaliGiornata = minutiLavoroTotali + minutiTrasferimentoTotali + minutiRientroUfficio;
  const oreTotaliGiornata = (minutiTotaliGiornata / 60).toFixed(1);

  const calendarEmbedUrl = `https://calendar.google.com/calendar/embed?src=${encodeURIComponent(tecnicoSelezionato.email)}&ctz=Europe%2FRome&mode=WEEK&showTitle=0&showPrint=0`;

  return (
    <main className="min-h-screen bg-slate-100 text-slate-900 font-sans p-6 md:p-10 overflow-x-hidden">
      <div className="max-w-7xl mx-auto">
        
        {/* HEADER */}
        <header className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-8 gap-4">
          <div>
            <h1 className="text-2xl md:text-3xl font-extrabold tracking-tight text-slate-900 flex items-center gap-3">
              <span className="p-2.5 bg-blue-600 text-white rounded-xl shadow-md"><Zap size={24} /></span>
              Asso Antincendio <span className="text-blue-600 font-normal text-xl">| Smart Logistics</span>
            </h1>
            <p className="text-xs text-slate-500 mt-1">Sincronizzazione API, orari calcolati e gestione cantieri drag & drop.</p>
          </div>

          <div className="flex flex-col md:flex-row items-center gap-4 w-full md:w-auto">
            {isCheckingAuth ? (
              <div className="bg-slate-50 text-slate-500 px-5 py-2.5 rounded-xl font-bold flex items-center justify-center gap-2 border border-slate-200 text-sm"><Loader2 size={18} className="animate-spin text-blue-600" /> Verifica connessione...</div>
            ) : !googleToken ? (
              <button onClick={eseguiLoginGoogle} className="bg-slate-900 hover:bg-slate-800 text-white px-5 py-2.5 rounded-xl font-bold flex items-center gap-2 shadow-sm text-sm"><Calendar size={18} /> Avvia Connessione Google</button>
            ) : (
              <div className="flex items-center gap-2">
                <span className="bg-emerald-50 text-emerald-800 px-4 py-2.5 rounded-xl font-bold flex items-center gap-2 border border-emerald-200 text-sm"><Check size={18} /> Calendar Connesso</span>
                <button onClick={logoutGoogle} className="bg-slate-100 hover:bg-red-50 text-slate-600 hover:text-red-600 px-3 py-2.5 rounded-xl border border-slate-200"><X size={18} /></button>
              </div>
            )}
          </div>
        </header>

        {/* PANNELLO CONFIG E UPLOAD */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
          
          <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col justify-between">
            <div>
              <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2 mb-4"><Calendar className="text-blue-600" size={18} /> 1. Configura Periodo</h2>
              <div className="grid grid-cols-2 gap-3 mb-4">
                <div><label className="text-[11px] font-bold text-slate-500 block mb-1">Dal</label><input type="date" value={dataInizio} onChange={(e) => setDataInizio(e.target.value)} className="w-full bg-slate-50 border border-slate-300 rounded-xl p-2.5 text-xs font-semibold focus:outline-none" /></div>
                <div><label className="text-[11px] font-bold text-slate-500 block mb-1">Al</label><input type="date" value={dataFine} onChange={(e) => setDataFine(e.target.value)} className="w-full bg-slate-50 border border-slate-300 rounded-xl p-2.5 text-xs font-semibold focus:outline-none" /></div>
              </div>
              <div className="flex flex-wrap gap-1.5 mb-4">
                {Object.keys(giorniAttivi).map((g) => <button key={g} onClick={() => toggleGiornoAttivo(g)} className={`px-2.5 py-1 rounded-lg text-xs font-bold ${giorniAttivi[g] ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-400 line-through'}`}>{g.slice(0, 3)}</button>)}
              </div>
            </div>
          </div>

          <div className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm flex flex-col justify-between">
             <div className="border-2 border-dashed border-slate-300 rounded-xl p-4 text-center bg-slate-50 flex flex-col items-center justify-center transition-all hover:border-blue-400 mb-4 h-full">
               <div className="p-2 bg-blue-100 text-blue-600 rounded-xl mb-2"><FileSpreadsheet size={24} /></div>
               <h2 className="text-sm font-bold text-slate-800 mb-1">{datiGrezziCaricati.length > 0 ? "File in Memoria" : "Carica file"}</h2>
               <label className={`bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl font-bold text-xs shadow-md transition-all inline-block mt-1 ${!googleToken ? 'opacity-50 cursor-not-allowed' : 'cursor-pointer'}`}>
                 {datiGrezziCaricati.length > 0 ? "Sostituisci" : "Sfoglia file"}
                 <input type="file" accept=".xlsx, .xls, .csv, .txt" className="hidden" onChange={handleFileUpload} disabled={!googleToken} />
               </label>
             </div>
             
             <div className="border-t border-slate-100 pt-3 flex flex-col gap-1.5">
               <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1"><UserCheck size={14} className="text-blue-600"/> 2. Tecnico Operativo</label>
               <select value={tecnicoSelezionato.nome} onChange={(e) => { const t = tecniciAnagrafica.find(t => t.nome === e.target.value); if(t) setTecnicoSelezionato(t); }} className="w-full bg-slate-50 border border-slate-300 rounded-xl p-2 text-xs font-bold text-slate-800 focus:outline-none cursor-pointer">
                  {tecniciAnagrafica.map((t, idx) => <option key={idx} value={t.nome}>{t.nome}</option>)}
               </select>
             </div>
          </div>

          <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col justify-between items-center text-center">
             <div className="mb-4">
               <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center justify-center gap-2 mb-2"><PlayCircle className="text-emerald-600" size={18} /> 3. Avvio Analisi Automatica</h2>
               <p className="text-xs text-slate-500">Incrocia percorsi, calcola orari esatti e taglia i doppioni dal Calendar.</p>
             </div>
             <button
               onClick={() => processaRighe(datiGrezziCaricati, nomeFileCorrente)}
               disabled={datiGrezziCaricati.length === 0 || inElaborazione}
               className={`w-full py-3.5 px-4 rounded-xl text-sm font-extrabold transition-all flex items-center justify-center gap-2 ${(datiGrezziCaricati.length > 0 && !inElaborazione) ? 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-md' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}
             >
               {inElaborazione ? <><Loader2 size={16} className="animate-spin"/> Elaborazione...</> : <><Sliders size={16} /> Analizza & Crea Orari</>}
             </button>
          </div>

        </div>

        {/* SEZIONE CENTRALE: SOSPESI E CALENDARIO GOOGLE */}
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6 mb-8">
          
          <div className="bg-amber-50 border border-amber-300 rounded-2xl p-5 shadow-sm flex flex-col h-[700px]">
            <h3 className="text-sm font-extrabold text-amber-900 flex items-center gap-2 mb-1"><PauseCircle size={18} className="text-amber-600" /> Sospesi ({clientiInSospeso.length})</h3>
            <p className="text-[11px] text-amber-800 mb-4 leading-tight">Gestisci l'orario e invia a Calendar.</p>
            
            <div className="space-y-4 overflow-y-auto pr-2 flex-1">
              {clientiInSospeso.map((s, idx) => {
                  const conf = configSospesi[s.codice] || { data: dataInizio, oraInizio: "08:30" };
                  const isSosBloccato = s.resoEstintori && s.giorniDallUltimoSos < 5;
                  const isCreating = creazioneInCorso[s.codice] || false;

                  return (
                    <div key={idx} className={`bg-white border p-4 rounded-2xl shadow-sm flex flex-col gap-3 transition-all ${isSosBloccato ? 'border-red-300 bg-red-50/40 opacity-70' : 'border-amber-200 hover:border-blue-400'}`}>
                      <div>
                        <div className="flex justify-between items-start mb-1"><span className="bg-amber-100 text-amber-900 font-mono text-[10px] px-2 py-0.5 rounded font-bold">BI: {s.numeroBi}</span><span className="text-[11px] text-slate-500 font-bold flex items-center gap-1"><Timer size={12} className="text-amber-600"/> {s.minutiLavoro}m</span></div>
                        <h4 className="font-bold text-slate-900 text-sm leading-tight mb-0.5">{s.nome}</h4>
                        <p className="text-[11px] text-slate-500 truncate">{s.indirizzo}, {s.localita}</p>
                        {isSosBloccato && <span className="text-[10px] text-red-600 font-bold flex items-start gap-1 leading-tight mt-2"><AlertTriangle size={12} className="flex-shrink-0" /> {s.motivoSospeso}</span>}
                      </div>
                      
                      <div className="border-t border-slate-100 pt-3 space-y-3">
                          <div className="grid grid-cols-2 gap-3">
                            <div>
                              <label className="text-[10px] font-bold text-slate-500 block mb-1">Data</label>
                              <input type="date" value={conf.data} onChange={(e) => aggiornaConfigSospeso(s.codice, 'data', e.target.value)} className="w-full bg-slate-50 border border-slate-300 rounded-lg p-2 text-[11px] font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-400" />
                            </div>
                            <div>
                              <label className="text-[10px] font-bold text-slate-500 block mb-1">Ora Inizio</label>
                              <input type="time" value={conf.oraInizio} onChange={(e) => aggiornaConfigSospeso(s.codice, 'oraInizio', e.target.value)} className="w-full bg-slate-50 border border-slate-300 rounded-lg p-2 text-[11px] font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-400" />
                            </div>
                          </div>

                          <button onClick={() => creaSuCalendarESposta(s, conf)} disabled={isSosBloccato || isCreating} className={`w-full text-xs py-2.5 rounded-xl font-bold flex items-center justify-center gap-1.5 shadow-sm transition-all ${isSosBloccato ? 'bg-slate-200 text-slate-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 text-white'}`}>
                              {isCreating ? <><Loader2 size={14} className="animate-spin" /> Creazione...</> : <><CalendarPlus size={14} /> Invia a Google</>}
                          </button>
                      </div>
                    </div>
                  );
              })}
              {clientiInSospeso.length === 0 && <div className="text-center text-xs text-amber-700/60 mt-10 font-medium">Nessun sospeso da assegnare.</div>}
            </div>
          </div>

          <div className="lg:col-span-3 bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col h-[700px] overflow-hidden">
            <div className="mb-4 flex justify-between items-center shrink-0">
              <div><h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-2 mb-1"><Calendar className="text-blue-600" size={18} /> Google Calendar ({tecnicoSelezionato.nome})</h3></div>
            </div>
            <div className="w-full h-[600px] rounded-xl overflow-hidden border border-slate-200 bg-slate-50">
              <iframe src={calendarEmbedUrl} style={{ border: 0 }} width="100%" height="100%" frameBorder="0" scrolling="no"></iframe>
            </div>
          </div>
        </div>

        {/* BOX RIEPILOGO, TRASPARENZA E INVIO MASSIVO */}
        {interventiGrezzi.length > 0 && !inElaborazione && (
          <div className="mb-8 bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6">
              <h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-2"><CheckCircle size={18} className="text-emerald-600" /> Riepilogo Pianificazione e Tempi</h3>
              
              {bulkSyncStatus.active ? (
                <div className="bg-blue-50 border border-blue-200 p-3 rounded-xl flex flex-col gap-2 w-full md:w-72">
                   <div className="flex justify-between text-[11px] font-bold text-blue-800">
                      <span className="flex items-center gap-1.5"><Loader2 size={12} className="animate-spin"/> Invio a Google in corso...</span>
                      <span>{bulkSyncStatus.current} / {bulkSyncStatus.total}</span>
                   </div>
                   <div className="w-full bg-blue-200 rounded-full h-2">
                      <div className="bg-blue-600 h-2 rounded-full transition-all duration-300" style={{width: `${bulkSyncStatus.progress}%`}}></div>
                   </div>
                </div>
              ) : (
                <button 
                  onClick={inviaPianificazioneAGoogle} 
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2.5 rounded-xl font-bold flex items-center gap-2 shadow-md text-xs transition-all w-full md:w-auto"
                >
                  <UploadCloud size={16} /> Invia Pianificazione a Google ({interventiGrezzi.filter(i => i.selezionatoPerGiro && !i.syncedToGoogle).length})
                </button>
              )}
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
               <div className="p-4 bg-emerald-50 rounded-xl border border-emerald-100 flex flex-col justify-center"><span className="block text-xs text-emerald-700 font-bold mb-1 uppercase tracking-wider">Pianificati Generati</span><span className="text-3xl font-extrabold text-emerald-900">{interventiGrezzi.length}</span></div>
               <div className="p-4 bg-amber-50 rounded-xl border border-amber-100 flex flex-col justify-center"><span className="block text-xs text-amber-700 font-bold mb-1 uppercase tracking-wider">Finiti nei Sospesi</span><span className="text-3xl font-extrabold text-amber-900">{clientiInSospeso.length}</span></div>
               <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 flex flex-col justify-center"><span className="block text-xs text-slate-500 font-bold mb-1 uppercase tracking-wider">Scartati (Già a Calendario)</span><span className="text-3xl font-extrabold text-slate-700">{clientiGiaCalendarizzati.length}</span></div>
            </div>
            
            <div className="mt-5 pt-5 border-t border-slate-100">
               <p className="text-[11px] font-bold text-slate-500 mb-3 uppercase tracking-wider">Scannerizzazione Spazio Google Calendar Effettuata:</p>
               <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                 {giornateStats.map((g, idx) => (
                   <div key={idx} className={`p-3 border rounded-xl flex flex-col gap-1 ${g.statoGiornata === "APERTO" ? 'bg-emerald-50/50 border-emerald-200' : 'bg-red-50/50 border-red-200'}`}>
                     <span className="text-xs font-bold text-slate-800">{g.dataStr}</span>
                     <span className="text-[10px] text-slate-500">{g.nomeGiorno}</span>
                     <div className="flex items-center gap-1.5 mt-1 font-mono text-[10px]">
                        {g.statoGiornata === "APERTO" ? <CheckCircle size={12} className="text-emerald-500" /> : <Lock size={12} className="text-red-500" />}
                        <span className={g.statoGiornata === "APERTO" ? "text-emerald-700 font-bold" : "text-red-700 font-bold"}>
                          {(g.minutiOccupati / 60).toFixed(1)}h occupate
                        </span>
                     </div>
                   </div>
                 ))}
               </div>
            </div>
          </div>
        )}

        {/* TABELLA E MAPPA DEI PIANIFICATI */}
        {interventiGrezzi.length > 0 && leafletLoaded && !inElaborazione && (
          <>
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden mb-12">
              <div className="p-6 border-b border-slate-200 flex flex-col md:flex-row justify-between items-start md:items-center gap-2">
                <div>
                  <h3 className="text-lg font-extrabold text-slate-900">Lista Pianificati | {tecnicoSelezionato.nome}</h3>
                  <p className="text-xs text-slate-500 mt-1">Puoi trascinare (drag & drop) le righe per riordinarle. I tempi verranno ricalcolati in automatico!</p>
                </div>
                <button onClick={() => window.print()} className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-xl font-bold flex items-center gap-2 shadow-sm text-xs transition-all"><Printer size={15} /> Stampa Foglio di Marcia</button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-50 border-b border-slate-200 text-xs uppercase text-slate-500 tracking-wider">
                    <tr>
                      <th className="p-4 w-12 text-center"></th>
                      <th className="p-4 w-16 text-center">Giro</th>
                      <th className="p-4">Riferimento BI & Cliente</th>
                      <th className="p-4">Data Assegnata</th>
                      <th className="p-4">Orario (Stimato)</th>
                      <th className="p-4">Stato Google</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {interventiVisibili.map((intervento, index) => {
                      const indiceReale = interventiGrezzi.findIndex(i => i.codice === intervento.codice);
                      return (
                        <tr 
                          key={intervento.codice} 
                          draggable
                          onDragStart={(e) => handleDragStartReorder(e, intervento.codice)}
                          onDragOver={handleDragOverReorder}
                          onDrop={(e) => handleDropReorder(e, intervento.codice)}
                          className={`hover:bg-slate-50 transition-colors cursor-grab active:cursor-grabbing ${!intervento.selezionatoPerGiro ? 'opacity-40 bg-slate-50/80' : ''}`}
                          title="Trascina su o giù per riordinare e ricalcolare i tempi"
                        >
                          <td className="p-4 text-center text-slate-300 hover:text-slate-500">
                            <GripVertical size={20} className="mx-auto" />
                          </td>
                          <td className="p-4 text-center">
                            <button onClick={() => toggleSelezioneCliente(indiceReale)} className="text-blue-600">{intervento.selezionatoPerGiro ? <CheckSquare size={20} /> : <Square size={20} className="text-slate-400" />}</button>
                          </td>
                          <td className="p-4">
                            <div className="flex items-center gap-2 mb-1"><span className="bg-blue-50 text-blue-700 font-mono text-[11px] px-2 py-0.5 rounded-md font-bold">BI: {intervento.numeroBi}</span></div>
                            <span className="font-bold text-slate-900 text-base block">{intervento.nome}</span>
                            <p className="text-slate-500 text-xs">{intervento.indirizzo} {intervento.localita ? `- ${intervento.localita}` : ""}</p>
                          </td>
                          <td className="p-4"><span className="bg-slate-100 text-slate-700 px-3 py-1.5 rounded-xl text-xs font-bold border border-slate-200 inline-block">{intervento.dataAssegnata} ({intervento.giorno})</span></td>
                          <td className="p-4">
                            <span className="block text-[11px] text-slate-500 font-mono font-bold flex items-center gap-1.5 mb-1"><Clock size={12}/> {intervento.oraInizio} - {intervento.oraFine}</span>
                            <span className="bg-emerald-50 text-emerald-700 px-2 py-1 rounded-lg text-[10px] font-bold inline-flex items-center gap-1 w-max border border-emerald-100"><Timer size={10} /> {intervento.minutiLavoro} min lavoro</span>
                          </td>
                          <td className="p-4">
                            {intervento.syncedToGoogle ? (
                              <span className="bg-emerald-50 text-emerald-700 px-3 py-1.5 rounded-xl text-xs font-bold flex items-center gap-1 w-max border border-emerald-200"><CheckCircle size={14} /> Sincronizzato</span>
                            ) : (
                              <span className="bg-slate-50 text-slate-500 px-3 py-1.5 rounded-xl text-xs font-bold flex items-center gap-1 w-max border border-slate-200"><Clock size={14} /> In Attesa</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="mb-10 bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
              <h3 className="text-base font-bold mb-4 flex items-center gap-2 text-slate-900"><MapPin className="text-red-500" /> Mappa Operativa GPS ({giornoSceltoFiltro})</h3>
              <div className="h-[480px] w-full rounded-xl overflow-hidden z-0 border border-slate-100">
                <MapContainer center={[SEDE_UFFICIO_LAT, SEDE_UFFICIO_LNG]} zoom={12} style={{ height: "100%", width: "100%" }}>
                  <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
                  {interventiVisibili.filter(i => i.lat && i.lng).map((c, i) => (
                    <Marker key={i} position={[c.lat, c.lng]}>
                      <Popup>
                        <div className="p-2">
                          <p className="font-bold text-sm text-slate-900">{c.nome}</p>
                          <p className="text-[11px] text-blue-600 font-mono font-bold mb-1">BI: {c.numeroBi}</p>
                          <p className="text-xs text-slate-600 mb-2">{c.indirizzo}, {c.localita}</p>
                        </div>
                      </Popup>
                    </Marker>
                  ))}
                </MapContainer>
              </div>
            </div>
          </>
        )}
      </div>
    </main>
  );
}