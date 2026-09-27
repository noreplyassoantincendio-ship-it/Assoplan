"use client";

import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import dynamic from "next/dynamic";
import { UploadCloud, AlertCircle, Package, MapPin, Loader2, FileSpreadsheet, MapPinOff, Zap, UserCheck, Printer, Calendar, Clock, CheckSquare, Square, Mail, Timer, FileText, X, Sliders, Check, Trash2, ArrowRight, PauseCircle, PlusCircle, ExternalLink, ShieldAlert, AlertTriangle, CheckCircle, PlayCircle, Lock, CalendarPlus, GripVertical, Phone, Book, User, Banknote, Search } from "lucide-react";

const MapContainer = dynamic(() => import("react-leaflet").then((mod) => mod.MapContainer), { ssr: false });
const TileLayer = dynamic(() => import("react-leaflet").then((mod) => mod.TileLayer), { ssr: false });
const Marker = dynamic(() => import("react-leaflet").then((mod) => mod.Marker), { ssr: false });
const Popup = dynamic(() => import("react-leaflet").then((mod) => mod.Popup), { ssr: false });

const GOOGLE_CLIENT_ID = "645365149295-lk8ei43kt09hsm2csupf643tgqkqqgmo.apps.googleusercontent.com";
const GOOGLE_SCOPES = "https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/calendar.events";

const FINE_GIORNATA_ASSOLUTA = 18 * 60; // 18:00 limite max assoluto
const INIZIO_PAUSA = 12 * 60 + 30; // 12:30
const FINE_PAUSA = 14 * 60; // 14:00
const DURATA_PAUSA = 90;

const formattaDataVisuale = (dataStr: string) => {
  if (!dataStr) return "";
  const parts = dataStr.split("-");
  if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
  return dataStr;
};

const getNextWeekDays = () => {
  const today = new Date();
  const dayOfWeek = today.getDay(); 
  const daysUntilNextMonday = dayOfWeek === 0 ? 1 : 8 - dayOfWeek;
  
  const nextMonday = new Date(today);
  nextMonday.setDate(today.getDate() + daysUntilNextMonday);
  
  const nextFriday = new Date(nextMonday);
  nextFriday.setDate(nextMonday.getDate() + 4);

  const fmt = (d: Date) => {
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  };
  
  return { start: fmt(nextMonday), end: fmt(nextFriday) };
};

const elaboraFileExcelGenerico = (file: File, callback: (righe: any[]) => void) => {
  if (file.name.toLowerCase().endsWith('.txt') || file.name.toLowerCase().endsWith('.csv')) {
    Papa.parse(file, { header: true, skipEmptyLines: true, dynamicTyping: true, complete: (res) => { callback(res.data); }});
  } else {
    const reader = new FileReader();
    reader.onload = async (evento) => {
      const data = evento.target?.result as ArrayBuffer;

      try {
        const testo = new TextDecoder("windows-1252").decode(data);
        if (testo.toLowerCase().includes("<table")) {
          const parser = new DOMParser(); 
          const doc = parser.parseFromString(testo, "text/html"); 
          const table = doc.querySelector("table");
          if (table) { 
            const workbookHTML = XLSX.utils.table_to_book(table); 
            callback(XLSX.utils.sheet_to_json(workbookHTML.Sheets[workbookHTML.SheetNames[0]])); 
            return; 
          }
        }
      } catch (err) {}

      const originalConsoleError = console.error;
      console.error = () => {}; 
      try {
        const arr = new Uint8Array(data);
        const wb = XLSX.read(arr, { type: "array" });
        callback(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));
      } catch (err) {
        alert("Formato file non supportato o file corrotto.");
      } finally {
        console.error = originalConsoleError;
      }
    };
    reader.readAsArrayBuffer(file);
  }
};

const calcolaTempistiche = (orarioPartenza: number, minViaggio: number, minutiLavoro: number) => {
    let inizioLavoro = orarioPartenza + minViaggio;
    
    if (inizioLavoro >= INIZIO_PAUSA && inizioLavoro < FINE_PAUSA) {
        inizioLavoro = FINE_PAUSA;
    }

    let fineLavoro = inizioLavoro + minutiLavoro;

    if (inizioLavoro < INIZIO_PAUSA && fineLavoro > INIZIO_PAUSA) {
        fineLavoro += DURATA_PAUSA;
    }

    return { inizioLavoro, fineLavoro };
};

export default function Home() {
  const defaultDates = getNextWeekDays();
  const [dataInizio, setDataInizio] = useState<string>(defaultDates.start);
  const [dataFine, setDataFine] = useState<string>(defaultDates.end);

  const [interventiGrezzi, setInterventiGrezzi] = useState<any[]>([]);
  const [clientiInSospeso, setClientiInSospeso] = useState<any[]>([]);
  const [clientiGiaCalendarizzati, setClientiGiaCalendarizzati] = useState<any[]>([]);
  const [inElaborazione, setInElaborazione] = useState(false);
  
  const [leafletLoaded, setLeafletLoaded] = useState(false);
  const leafletRef = useRef<any>(null);
  
  const [isCheckingAuth, setIsCheckingAuth] = useState(true);
  const [googleToken, setGoogleToken] = useState<string | null>(null);
  const tokenClientRef = useRef<any>(null);

  const [datiGrezziCaricati, setDatiGrezziCaricati] = useState<any[]>([]);
  const [nomeFileCorrente, setNomeFileCorrente] = useState<string>("");
  
  const [rubricaClienti, setRubricaClienti] = useState<{ [codice: string]: any }>({});
  const [listaInsoluti, setListaInsoluti] = useState<any[] | null>(null);

  const [giorniAttivi, setGiorniAttivi] = useState<{ [key: string]: boolean }>({
    "Lunedì": true, "Martedì": true, "Mercoledì": true, "Giovedì": true, "Venerdì": true, "Sabato": false
  });

  const [giornoSceltoFiltro, setGiornoSceltoFiltro] = useState<string>("Tutti");
  const [clienteSelezionatoScheda, setClienteSelezionatoScheda] = useState<any | null>(null);
  
  const [sospesoInModifica, setSospesoInModifica] = useState<any | null>(null);
  const [configSospesoSingolo, setConfigSospesoSingolo] = useState<{data: string, oraInizio: string}>({data: dataInizio, oraInizio: "08:30"});

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
      leafletRef.current = L;
      // @ts-ignore
      delete L.Icon.Default.prototype._getIconUrl;
      L.Icon.Default.mergeOptions({
        iconRetinaUrl: "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-icon-2x.png",
        iconUrl: "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-icon.png",
        shadowUrl: "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-shadow.png",
      });
      setLeafletLoaded(true);
    });

    const storedGoogle = localStorage.getItem("asso_google_token");
    if (storedGoogle) {
      try {
        const parsed = JSON.parse(storedGoogle);
        if (Date.now() < parsed.expiry) setGoogleToken(parsed.token);
        else localStorage.removeItem("asso_google_token");
      } catch (e) { localStorage.removeItem("asso_google_token"); }
    }

    const storedRubrica = localStorage.getItem("asso_rubrica_clienti");
    if (storedRubrica) { try { setRubricaClienti(JSON.parse(storedRubrica)); } catch (e) {} }
  }, []);

  useEffect(() => {
    const initGoogleClient = () => {
      // @ts-ignore
      if (window.google && window.google.accounts) {
        // @ts-ignore
        tokenClientRef.current = window.google.accounts.oauth2.initTokenClient({
          client_id: GOOGLE_CLIENT_ID, scope: GOOGLE_SCOPES,
          callback: (response: any) => {
            if (response.error) { alert("Impossibile connettersi a Google: " + response.error); return; }
            if (response.access_token) {
              setGoogleToken(response.access_token);
              localStorage.setItem("asso_google_token", JSON.stringify({ token: response.access_token, expiry: Date.now() + 3300 * 1000 }));
            }
          },
        });
        setIsCheckingAuth(false);
      }
    };
    // @ts-ignore
    if (window.google) initGoogleClient();
    else {
      const script = document.createElement("script");
      script.src = "https://accounts.google.com/gsi/client";
      script.async = true; script.defer = true; script.onload = initGoogleClient;
      document.head.appendChild(script);
    }
  }, []);

  const eseguiLoginGoogle = () => {
    if (tokenClientRef.current) tokenClientRef.current.requestAccessToken();
    else alert("Caricamento in corso. Riprova tra un istante.");
  };

  const logoutGoogle = () => { setGoogleToken(null); localStorage.removeItem("asso_google_token"); };

  useEffect(() => {
    if (clienteSelezionatoScheda || sospesoInModifica) document.body.style.overflow = "hidden";
    else document.body.style.overflow = "unset";
    return () => { document.body.style.overflow = "unset"; };
  }, [clienteSelezionatoScheda, sospesoInModifica]);

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
        const y = curr.getFullYear(); const m = String(curr.getMonth() + 1).padStart(2, '0'); const d = String(curr.getDate()).padStart(2, '0');
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

      const url = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(emailTecnico)}/events?timeMin=${dataMin.toISOString()}&timeMax=${dataMax.toISOString()}&singleEvents=true&orderBy=startTime&maxResults=2500`;
      const response = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: "application/json" } });

      if (!response.ok) {
        if (response.status === 401) { logoutGoogle(); alert("Token di Google scaduto."); }
        return [];
      }

      const data = await response.json();
      const eventi = [];

      for (const evt of (data.items || [])) {
        const testoCompleto = `${evt.summary || ""} ${evt.description || ""}`.toUpperCase();
        const matches = testoCompleto.match(/(?:BI[:\s]*)([A-Z0-9\/\-]+)/g);
        const biGiaLetti = matches ? matches.map(m => m.replace(/BI[:\s]*/g, "").trim()) : [];
        
        let latEvt: number | null = null; let lngEvt: number | null = null;
        if (evt.location && evt.location.trim().length > 2) {
          const coord = await trovaCoordinateGoogle(evt.location, "Genova");
          if (coord) { latEvt = coord.lat; lngEvt = coord.lng; }
        }

        if (evt.start?.date && evt.end?.date) {
            const y = parseInt(evt.start.date.split("-")[0]);
            const m = parseInt(evt.start.date.split("-")[1]);
            const d = parseInt(evt.start.date.split("-")[2]);
            const startD = new Date(y, m - 1, d, 8, 0);
            const endD = new Date(y, m - 1, d, 17, 0); 
            
            eventi.push({ 
                id: evt.id, dataStr: evt.start.date, 
                lat: latEvt, lng: lngEvt, 
                durataMinuti: 480, fineMinuti: 17 * 60, numeriBiRilevati: biGiaLetti,
                summary: evt.summary || "", description: evt.description || "", location: evt.location || "",
                startD, endD
            });
        } else if (evt.start?.dateTime && evt.end?.dateTime) {
            const startD = new Date(evt.start.dateTime);
            const endD = new Date(evt.end.dateTime);
            const y = startD.getFullYear(); const m = String(startD.getMonth() + 1).padStart(2, '0'); const d = String(startD.getDate()).padStart(2, '0');
            const fineMinutiLocali = endD.getHours() * 60 + endD.getMinutes();

            eventi.push({ 
                id: evt.id, dataStr: `${y}-${m}-${d}`, 
                lat: latEvt, lng: lngEvt, 
                durataMinuti: Math.round((endD.getTime() - startD.getTime()) / 60000), fineMinuti: fineMinutiLocali, numeriBiRilevati: biGiaLetti,
                summary: evt.summary || "", description: evt.description || "", location: evt.location || "",
                startD, endD
            });
        }
      }
      return eventi;
    } catch (err) { return []; }
  };

  const generaTitoloEvento = (cliente: any) => {
    if (cliente.isDistinta) return cliente.nome;
    if (cliente.isPregresso) return cliente.nome;
    let tags = [];
    if (cliente.haInsoluto) tags.push("[INS]");
    if (cliente.resoEstintori) tags.push("[SOS]");
    if (!cliente.nome.toUpperCase().includes("ASL")) tags.push("[DC]");
    
    const tagString = tags.length > 0 ? tags.join(" ") + " " : "";
    return `${tagString}${cliente.nome} - BI ${cliente.numeroBi}`;
  };

  const generaDescrizioneEvento = (cliente: any, tecnicoNome: string) => {
    if (cliente.isDistinta) return cliente.descrizioneDistinta;
    if (cliente.isPregresso) return cliente.descrizionePregressa;

    let attrezzatureTesto = cliente.dettaglioAttrezzature && cliente.dettaglioAttrezzature.length > 0 
      ? cliente.dettaglioAttrezzature.map((a: any) => `🔧 ${a.descrizione}: ${a.quantita}`).join("\n")
      : "Nessuna attrezzatura specificata";

    const contattoInfo = cliente.contatto ? `👤 Contatto: ${cliente.contatto}` : "";
    const telInfo = cliente.telefono ? `📞 ${cliente.telefono}` : "";
    const emailInfo = cliente.email ? `📧 Email: ${cliente.email}` : "";
    
    let stringaContatti = [contattoInfo, telInfo, emailInfo].filter(Boolean).join("\n");
    if (stringaContatti.length > 0) stringaContatti = `\n${stringaContatti}\n`;

    return `Località: ${cliente.localita || ''}
Indirizzo: ${cliente.indirizzo || ''}
${stringaContatti}
🛠 ATTREZZATURE DA VERIFICARE:
${attrezzatureTesto}

Durata: ${cliente.minutiLavoro} min
📍 Naviga:
https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${cliente.indirizzo}, ${cliente.localita || 'Genova'}, Italy`)}

Tecnico: ${tecnicoNome}`;
  };

  const ricalcolaDistinteEOrari = (listaDaOrdinare: any[], statsLocali: any[] = giornateStats) => {
      const listaPulita = listaDaOrdinare.filter(i => !i.isDistinta);
      const mapGiorni = new Map();
      const vecchieDistinte = listaDaOrdinare.filter(i => i.isDistinta);
      const nuoveStats = statsLocali.map(g => ({ ...g }));

      listaPulita.forEach(i => {
        if (!mapGiorni.has(i.dataAssegnata)) mapGiorni.set(i.dataAssegnata, []);
        mapGiorni.get(i.dataAssegnata).push(i);
      });

      const nuovaLista: any[] = [];

      mapGiorni.forEach((interventiGiorno, dataStr) => {
        const stat = statsLocali.find(g => g.dataStr === dataStr);
        let orarioCorrenteMinuti = stat && stat.orarioPartenzaMinuti ? stat.orarioPartenzaMinuti : 8 * 60;
        
        let ultimaCoordGiorno: { lat: number; lng: number } | null = null;
        let furgone = 0; let piccoNegativo = 0; let totaleMovimenti = 0;

        if (stat && stat.anchorLat) {
            ultimaCoordGiorno = { lat: stat.anchorLat, lng: stat.anchorLng };
        }

        interventiGiorno.forEach((c: any) => {
            if (c.isPregresso) {
                const [hEnd, mEnd] = c.oraFine.split(":").map(Number);
                orarioCorrenteMinuti = Math.max(orarioCorrenteMinuti, (hEnd * 60 + mEnd) + 30);
                if (c.lat && c.lng) ultimaCoordGiorno = { lat: c.lat, lng: c.lng };
            } else if (!c.selezionatoPerGiro) {
                c.oraInizio = "-";
                c.oraFine = "-";
            } else {
                const puntoPartenza = ultimaCoordGiorno || { lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG };
                const dist = calcolaDistanzaKm(puntoPartenza.lat, puntoPartenza.lng, c.lat, c.lng);
                const minViaggio = Math.round(dist * 2);

                const { inizioLavoro, fineLavoro } = calcolaTempistiche(orarioCorrenteMinuti, minViaggio, c.minutiLavoro);

                const hhStart = String(Math.floor(inizioLavoro / 60)).padStart(2, '0');
                const mmStart = String(inizioLavoro % 60).padStart(2, '0');
                const hhEnd = String(Math.floor(fineLavoro / 60)).padStart(2, '0');
                const mmEnd = String(fineLavoro % 60).padStart(2, '0');

                c.oraInizio = `${hhStart}:${mmStart}`;
                c.oraFine = `${hhEnd}:${mmEnd}`;

                orarioCorrenteMinuti = fineLavoro;
                ultimaCoordGiorno = { lat: c.lat, lng: c.lng };

                let delta = c.resoEstintori ? c.muletti_eorv : -c.muletti_eorv;
                furgone += delta;
                totaleMovimenti += c.muletti_eorv;
                if (furgone < piccoNegativo) piccoNegativo = furgone;
            }
        });

        const targetStat = nuoveStats.find(s => s.dataStr === dataStr);
        if (targetStat) {
            targetStat.orarioPartenzaMinuti = orarioCorrenteMinuti;
            targetStat.statoGiornata = (FINE_GIORNATA_ASSOLUTA - orarioCorrenteMinuti) < 60 ? "PIENO (Sospeso)" : "APERTO";
        }

        const fabbisognoNetto = Math.abs(piccoNegativo);
        const scortaJolly = Math.max(2, Math.ceil(fabbisognoNetto * 0.20));
        const totaleDaCaricare = fabbisognoNetto + scortaJolly;

        let descStr = `🛠 DISTINTA MULETTI (Calcolo a Cascata):\n\n`;
        descStr += `• Fabbisogno Netto (Scoperto Max Tappe): ${fabbisognoNetto} pz\n`;
        descStr += `• Scorta Jolly Sicurezza (20% Fabbisogno, min 2): ${scortaJolly} pz\n`;
        descStr += `-----------------------------------\n`;
        descStr += `TOTALE MULETTI DA CARICARE: ${totaleDaCaricare} pz\n`;
        descStr += `-----------------------------------\n`;
        descStr += `(Movimentazione totale prevista: ${totaleMovimenti} pz)`;

        const primaRiga = interventiGiorno[0];
        const oldDistinta = vecchieDistinte.find(d => d.dataAssegnata === dataStr);

        const distintaObj = {
            codice: `DISTINTA_${dataStr}`, isDistinta: true, nome: `📦 DISTINTA DI CARICO (${formattaDataVisuale(dataStr)})`, numeroBi: "MAGAZZINO", indirizzo: "Sede Asso Antincendio", localita: "Genova", minutiLavoro: 60, giorno: primaRiga.giorno, dataAssegnata: dataStr, oraInizio: "07:00", oraFine: "08:00", descrizioneDistinta: descStr, selezionatoPerGiro: true, syncedToGoogle: oldDistinta ? oldDistinta.syncedToGoogle : false, lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG
        };

        nuovaLista.push(distintaObj, ...interventiGiorno);
      });

      return { listaOrdinata: nuovaLista, statsAggiornate: nuoveStats };
  };

  const processaRighe = async (righe: any[]) => {
    if (!googleToken) { alert("Connetti l'account Google per elaborare i dati."); return; }
    if (righe.length === 0) { alert("Nessun dato caricato. Seleziona prima un file."); return; }

    setInElaborazione(true);
    const clientiMappa = new Map();
    
    const impegniEsistenti = await fetchEventiRealiCalendar(tecnicoSelezionato.email, googleToken);
    let tuttiIBiGiaInCalendario: string[] = [];
    impegniEsistenti.forEach(evt => { tuttiIBiGiaInCalendario = [...tuttiIBiGiaInCalendario, ...evt.numeriBiRilevati]; });

    for (const riga of righe) {
      const nomeCliente = (riga.desclifor || "").toUpperCase();
      if (!nomeCliente) continue;

      const numeroBi = String(riga.numdoc || riga.nrob || riga.id_bi || riga.documento || riga.codicebi || "BI-GENERICA").trim();
      const codPerRubrica = riga.codclifor ? String(riga.codclifor).trim() : "";
      const idCliente = riga.codclifor || riga.desclifor || Math.random().toString();

      const datiRubrica = rubricaClienti[codPerRubrica] || {};
      const telDaUsare = datiRubrica.cellulare || datiRubrica.telefono || riga.telefono || riga.TELEFONO || riga.tel || riga.TEL || "";
      const emailDaUsare = datiRubrica.email || riga.email || riga.EMAIL || riga['e-mail'] || "";
      const contattoDaUsare = datiRubrica.contatto || "";

      if (!clientiMappa.has(idCliente)) {
        clientiMappa.set(idCliente, {
          codice: idCliente, numeroBi, nome: riga.desclifor || "Sconosciuto",
          indirizzo: riga.indirizzo || riga.INDIRIZZO || "", localita: riga.localita || riga.LOCALITA || "",
          telefono: telDaUsare, email: emailDaUsare, contatto: contattoDaUsare,
          haInsoluto: false, resoEstintori: false, totaleArticoli: 0, minutiLavoro: 20, dettaglioAttrezzature: [], selezionatoPerGiro: true, giorniDallUltimoSos: 8, syncedToGoogle: false, muletti_eorv: 0 
        });
      }

      const cliente = clientiMappa.get(idCliente);
      const testoRiga = JSON.stringify(riga).toUpperCase();
      const numFattura = riga.numft ? String(riga.numft).trim() : "";
      if (numFattura !== "" && numFattura !== "-" && (parseFloat(riga.imptotft) || 0) > 0 && !(riga.codpagam || "").toUpperCase().trim().startsWith("RB")) cliente.haInsoluto = true;
      if ((riga.codclose ? String(riga.codclose).toUpperCase().trim() : "") === "SOS" || testoRiga.includes("SOS")) cliente.resoEstintori = true;

      const eoRv = Number(riga.eo_rv || riga.EO_RV) || 0;
      cliente.muletti_eorv += eoRv;

      let minutiTotaliRiga = 0; let pezziTotaliRiga = 0; const dettaglio: any[] = [];
      if (cliente.resoEstintori) {
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

    const dateRange = getGiornateLavorative(dataInizio, dataFine, giorniAttivi);

    const eventiPregressiUI = impegniEsistenti.filter(evt => dateRange.some(d => d.dataStr === evt.dataStr)).map(evt => {
        const hhStart = String(evt.startD.getHours()).padStart(2, '0');
        const mmStart = String(evt.startD.getMinutes()).padStart(2, '0');
        const hhEnd = String(evt.endD.getHours()).padStart(2, '0');
        const mmEnd = String(evt.endD.getMinutes()).padStart(2, '0');

        let nomeCliente = evt.summary || "Impegno in Calendario";
        let numeroBi = evt.numeriBiRilevati.length > 0 ? evt.numeriBiRilevati[0] : "-";
        nomeCliente = nomeCliente.replace(/\[.*?\]/g, "").trim();
        if (nomeCliente.includes("- BI")) nomeCliente = nomeCliente.split("- BI")[0].trim();

        let telefono = ""; let email = ""; let contatto = "";
        if (evt.description) {
            const telMatch = evt.description.match(/📞\s*([0-9\+\s]+)/);
            if (telMatch) telefono = telMatch[1].trim();
            const emailMatch = evt.description.match(/📧\s*Email:\s*([^\s\n]+)/);
            if (emailMatch) email = emailMatch[1].trim();
            const contattoMatch = evt.description.match(/👤\s*Contatto:\s*([^\n]+)/);
            if (contattoMatch) contatto = contattoMatch[1].trim();
        }

        return {
            codice: `PREGRESSO_${evt.id}`,
            isPregresso: true, isDistinta: false,
            nome: nomeCliente, numeroBi: numeroBi,
            indirizzo: evt.location || "Vedi Google Calendar", localita: "",
            oraInizio: `${hhStart}:${mmStart}`, oraFine: `${hhEnd}:${mmEnd}`,
            dataAssegnata: evt.dataStr, giorno: getNomeGiorno(evt.dataStr),
            selezionatoPerGiro: true, syncedToGoogle: true,
            minutiLavoro: evt.durataMinuti, lat: evt.lat, lng: evt.lng,
            contatto, telefono, email, muletti_eorv: 0,
            descrizionePregressa: evt.description || "Nessun dettaglio da Google"
        };
    });

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

    const giornateStrutturate = dateRange.map(g => {
      const pregressiDelGiorno = eventiPregressiUI.filter(e => e.dataAssegnata === g.dataStr);
      let maxFineMinuti = 8 * 60;
      let anchorLat: number | null = null; let anchorLng: number | null = null;

      pregressiDelGiorno.forEach(imp => {
        const [h, m] = imp.oraFine.split(":").map(Number);
        const fine = h * 60 + m;
        if (fine > maxFineMinuti) maxFineMinuti = fine;
        if (imp.lat && imp.lng && !anchorLat) { anchorLat = imp.lat; anchorLng = imp.lng; }
      });

      if (maxFineMinuti > 8 * 60) maxFineMinuti += 30; // +30 min respiro

      if (maxFineMinuti >= INIZIO_PAUSA && maxFineMinuti < FINE_PAUSA) {
         maxFineMinuti = FINE_PAUSA;
      }

      let minutiDisponibili = Math.max(0, FINE_GIORNATA_ASSOLUTA - maxFineMinuti);
      let statoGiornata = "APERTO";
      if (minutiDisponibili < 60) { minutiDisponibili = 0; statoGiornata = "PIENO (Sospeso)"; }

      return { ...g, minutiDisponibili, statoGiornata, anchorLat, anchorLng, orarioPartenzaMinuti: maxFineMinuti };
    });

    const clientiPianificati = [];

    for (const giornata of giornateStrutturate) {
      if (giornata.minutiDisponibili <= 0) continue; 
      
      let ultimaCoordGiorno: { lat: number; lng: number } | null = giornata.anchorLat !== null && giornata.anchorLng !== null ? { lat: giornata.anchorLat, lng: giornata.anchorLng } : null;
      let orarioAttualeSimulato = giornata.orarioPartenzaMinuti; 
      let aggiuntoQualcuno = true;

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
          
          const { inizioLavoro, fineLavoro } = calcolaTempistiche(orarioAttualeSimulato, minViaggio, candidato.minutiLavoro);

          if ((fineLavoro + minRientro) <= FINE_GIORNATA_ASSOLUTA) {
            clientiPianificati.push({ 
                ...candidato, 
                settimana: `${formattaDataVisuale(dataInizio)} al ${formattaDataVisuale(dataFine)}`, 
                giorno: giornata.nomeGiorno, 
                dataAssegnata: giornata.dataStr 
            });
            orarioAttualeSimulato = fineLavoro;
            ultimaCoordGiorno = { lat: candidato.lat, lng: candidato.lng };
            clientiDaPianificare.splice(bestIdx, 1);
            aggiuntoQualcuno = true;
          } else { break; }
        }
      }
    }

    clientiDaPianificare.forEach(c => { clientiSospesiTmp.push({ ...c, motivoSospeso: "Tempo esaurito (Giornate Piene)" }); });
    
    const datiCombinati = [...eventiPregressiUI, ...clientiPianificati].sort((a, b) => {
        if (a.dataAssegnata !== b.dataAssegnata) return a.dataAssegnata.localeCompare(b.dataAssegnata);
        if (a.isPregresso && !b.isPregresso) return -1;
        if (!a.isPregresso && b.isPregresso) return 1;
        if (a.isPregresso && b.isPregresso) return a.oraInizio.localeCompare(b.oraInizio);
        return 0;
    });

    const res = ricalcolaDistinteEOrari(datiCombinati, giornateStrutturate);
    setInterventiGrezzi(res.listaOrdinata);
    setGiornateStats(res.statsAggiornate);
    setClientiInSospeso(clientiSospesiTmp);
    setClientiGiaCalendarizzati(clientiScartatiTmp);
    setInElaborazione(false);
  };

  const toggleSelezioneCliente = (codice: string) => { 
    const nuovi = [...interventiGrezzi]; 
    const idx = nuovi.findIndex(i => i.codice === codice);
    if(idx !== -1) {
      nuovi[idx].selezionatoPerGiro = !nuovi[idx].selezionatoPerGiro; 
      const resRic = ricalcolaDistinteEOrari(nuovi, giornateStats);
      setInterventiGrezzi(resRic.listaOrdinata);
      setGiornateStats(resRic.statsAggiornate);
    }
  };

  const rimuoviDaGiornata = (codice: string) => {
    const itemToRemove = interventiGrezzi.find(i => i.codice === codice);
    if (!itemToRemove) return;
    const nuoviGrezzi = interventiGrezzi.filter(i => i.codice !== codice);
    setClientiInSospeso(prev => [...prev, { ...itemToRemove, motivoSospeso: "Rimosso manualmente", selezionatoPerGiro: true }]);
    const resRic = ricalcolaDistinteEOrari(nuoviGrezzi, giornateStats);
    setInterventiGrezzi(resRic.listaOrdinata);
    setGiornateStats(resRic.statsAggiornate);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!googleToken) { alert("Devi connettere Google Calendar prima di caricare il file."); e.target.value = ''; return; }
    const file = e.target.files?.[0];
    if (!file) return;

    const nomeFileLower = file.name.toLowerCase();
    setNomeFileCorrente(file.name);
    
    for (const tech of tecniciAnagrafica) {
      if (tech.nome.toLowerCase().split(" ").some(p => p.length > 2 && nomeFileLower.includes(p))) {
        setTecnicoSelezionato(tech);
        break;
      }
    }

    elaboraFileExcelGenerico(file, (righe) => {
       if (righe.length > 0) setDatiGrezziCaricati(righe);
    });
    e.target.value = '';
  };

  const handleRubricaUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    elaboraFileExcelGenerico(file, (righe) => {
      const nuovaRubrica: any = {};
      righe.forEach(r => {
        const findKey = (searchTerms: string[]) => {
            const key = Object.keys(r).find(k => searchTerms.some(term => k.toLowerCase().includes(term.toLowerCase())));
            return key ? r[key] : "";
        };
        const cod = findKey(["codice", "codclifor", "cliente"]);
        if (cod) {
          nuovaRubrica[String(cod).trim()] = {
            cellulare: findKey(["cell"]),
            telefono: findKey(["telefono", "tel"]),
            contatto: findKey(["contatto", "referente"]),
            email: findKey(["email", "e-mail", "mail"])
          };
        }
      });
      setRubricaClienti(nuovaRubrica);
      localStorage.setItem("asso_rubrica_clienti", JSON.stringify(nuovaRubrica));
      alert(`Rubrica caricata con successo in memoria! Registrati ${Object.keys(nuovaRubrica).length} contatti.`);
    });
    e.target.value = '';
  };

  const handleInsolutiUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    elaboraFileExcelGenerico(file, (righe) => {
       const insolutiMappa = new Map();
       
       righe.forEach(r => {
          const numeroBi = String(r.numdoc || r.nrob || r.id_bi || r.documento || r.codicebi || "").trim();
          const numFattura = r.numft ? String(r.numft).trim() : "";
          const importo = parseFloat(r.imptotft) || 0;
          const codpagam = String(r.codpagam || "").toUpperCase().trim();
          const despagam = String(r.despagam || r.des_pagam || r.pagamento || "").trim();
          const modalitaPagamento = despagam ? `${codpagam} - ${despagam}` : codpagam;
          
          if (numFattura && numFattura !== "-" && importo > 0 && !codpagam.startsWith("RB")) {
             const chiaveUnivoca = numeroBi || `${numFattura}_${r.codclifor}`;
             
             if (!insolutiMappa.has(chiaveUnivoca)) {
                 const codclifor = String(r.codclifor || "").trim();
                 const rubrica = rubricaClienti[codclifor] || {};
                 const telefono = rubrica.cellulare || rubrica.telefono || r.telefono || r.TELEFONO || r.tel || r.TEL || "";
                 const email = rubrica.email || r.email || r.EMAIL || r['e-mail'] || "";
                 const contatto = rubrica.contatto || "";
                 const note = r.note_contabilita || r.note || r.NOTE || "";
                 
                 insolutiMappa.set(chiaveUnivoca, {
                    codiceCliente: codclifor,
                    anagrafica: r.desclifor || "Sconosciuto",
                    numFattura,
                    importo,
                    telefono, email, contatto,
                    note,
                    modalitaPagamento
                 });
             }
          }
       });
       setListaInsoluti(Array.from(insolutiMappa.values()));
    });
    e.target.value = '';
  };

  const toggleGiornoAttivo = (giorno: string) => { setGiorniAttivi(prev => ({ ...prev, [giorno]: !prev[giorno] })); };
  
  const apriModaleSospeso = (s: any) => {
    setSospesoInModifica(s);
    setConfigSospesoSingolo({ data: dataInizio, oraInizio: "08:30" });
  };

  const creaSospesoDaModale = async (action: 'lista' | 'calendar') => {
    if (!sospesoInModifica) return;
    const clienteSospeso = sospesoInModifica;
    const conf = configSospesoSingolo;

    if (clienteSospeso.resoEstintori && clienteSospeso.giorniDallUltimoSos < 5) { alert(`Impossibile ripianificare l'intervento SOS.`); return; }
    if (action === 'calendar' && !googleToken) { alert("Devi connettere Google Calendar!"); return; }

    if (action === 'calendar') {
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
          const newClient = { ...clienteSospeso, settimana: `${formattaDataVisuale(dataInizio)} al ${formattaDataVisuale(dataFine)}`, giorno: giornoScelto, dataAssegnata: conf.data, oraInizio: conf.oraInizio, oraFine: `${hhEnd}:${mmEnd}`, selezionatoPerGiro: true, syncedToGoogle: true };
          
          const resRic = ricalcolaDistinteEOrari([...interventiGrezzi, newClient], giornateStats);
          setInterventiGrezzi(resRic.listaOrdinata);
          setGiornateStats(resRic.statsAggiornate);
          
          setClientiInSospeso(prev => prev.filter(c => c.codice !== clienteSospeso.codice));
          setSospesoInModifica(null);
        } else { alert("Errore durante la creazione API su Google."); }
      } catch (e) { alert("Errore di rete."); } 
    } else {
      const startDateTime = new Date(`${conf.data}T${conf.oraInizio}:00`);
      const endDateTime = new Date(startDateTime.getTime() + clienteSospeso.minutiLavoro * 60000);
      const hhEnd = String(endDateTime.getHours()).padStart(2, '0');
      const mmEnd = String(endDateTime.getMinutes()).padStart(2, '0');

      const newClient = { ...clienteSospeso, settimana: `${formattaDataVisuale(dataInizio)} al ${formattaDataVisuale(dataFine)}`, giorno: getNomeGiorno(conf.data), dataAssegnata: conf.data, oraInizio: conf.oraInizio, oraFine: `${hhEnd}:${mmEnd}`, selezionatoPerGiro: true, syncedToGoogle: false };
      
      const resRic = ricalcolaDistinteEOrari([...interventiGrezzi, newClient], giornateStats);
      setInterventiGrezzi(resRic.listaOrdinata);
      setGiornateStats(resRic.statsAggiornate);
      
      setClientiInSospeso(prev => prev.filter(c => c.codice !== clienteSospeso.codice));
      setSospesoInModifica(null);
    }
  };

  const inviaPianificazioneAGoogle = async () => {
    if (!googleToken) { alert("Devi connettere Google Calendar!"); return; }
    const itemsToSync = interventiGrezzi.filter(i => i.selezionatoPerGiro && !i.syncedToGoogle && !i.isPregresso);
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
        } catch (e) {}

        syncedCount++;
        setBulkSyncStatus({ active: true, progress: Math.round((syncedCount / itemsToSync.length) * 100), current: syncedCount, total: itemsToSync.length });
        setInterventiGrezzi([...nuoviInterventiGrezzi]); 
    }

    setTimeout(() => setBulkSyncStatus({ active: false, progress: 0, current: 0, total: 0 }), 2000); 
  };

  const handleDragStartReorder = (e: React.DragEvent, codice: string) => { e.dataTransfer.setData("application/reorder-lista", codice); };
  const handleDragOverReorder = (e: React.DragEvent) => { e.preventDefault(); };

  const handleDropReorder = (e: React.DragEvent, targetCodice: string) => {
    e.preventDefault();
    const draggedCodice = e.dataTransfer.getData("application/reorder-lista");
    if (!draggedCodice || draggedCodice === targetCodice) return;

    setInterventiGrezzi(prev => {
      const result = Array.from(prev);
      const draggedIndex = result.findIndex(i => i.codice === draggedCodice);
      const targetIndex = result.findIndex(i => i.codice === targetCodice);

      if (draggedIndex === -1 || targetIndex === -1) return prev;

      const [removed] = result.splice(draggedIndex, 1);
      result.splice(targetIndex, 0, removed);

      const resRic = ricalcolaDistinteEOrari(result, giornateStats);
      setGiornateStats(resRic.statsAggiornate);
      return resRic.listaOrdinata;
    });
  };

  const createNumberedIcon = (numero: number) => {
    if (leafletRef.current) {
      return leafletRef.current.divIcon({
        className: "custom-marker-icon",
        html: `<div style="background-color: #2563eb; color: white; border-radius: 50%; width: 26px; height: 26px; display: flex; align-items: center; justify-content: center; font-weight: bold; border: 2px solid white; box-shadow: 0 2px 4px rgba(0,0,0,0.4); font-size: 11px;">${numero}</div>`,
        iconSize: [26, 26],
        iconAnchor: [13, 26],
        popupAnchor: [0, -26]
      });
    }
    return undefined;
  };

  let currProgressivo = 0;
  const interventiVisibiliArr = interventiGrezzi
    .filter(i => giornoSceltoFiltro === "Tutti" || i.dataAssegnata === giornoSceltoFiltro)
    .map(i => {
       if (i.isDistinta || !i.selezionatoPerGiro || i.isPregresso) return { ...i, numProgressivo: null };
       currProgressivo++;
       return { ...i, numProgressivo: currProgressivo };
    });

  const clientiAttiviCount = interventiVisibiliArr.filter(i => i.selezionatoPerGiro && !i.isDistinta && !i.isPregresso).length;
  
  const calDataMin = dataInizio.replace(/-/g, '');
  const calDataMax = dataFine.replace(/-/g, '');
  const calendarEmbedUrl = `https://calendar.google.com/calendar/embed?src=${encodeURIComponent(tecnicoSelezionato.email)}&ctz=Europe%2FRome&mode=WEEK&showTitle=0&showPrint=0&dates=${calDataMin}%2F${calDataMax}`;
  
  const giorniVisibiliPerStampa = [...new Set(interventiVisibiliArr.map(i => i.dataAssegnata))].sort();

  if (listaInsoluti) {
     return (
        <div className="min-h-screen bg-white text-black p-8 font-sans print:p-0">
           
           <style dangerouslySetInnerHTML={{__html: `
             @media print {
               @page { size: landscape; margin: 10mm; }
               body { background-color: white !important; margin: 0; padding: 0; }
               * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
             }
           `}} />

           <div className="flex justify-between items-center mb-8 border-b-2 border-slate-900 pb-4 print:hidden">
              <div>
                 <h1 className="text-3xl font-extrabold text-slate-900 flex items-center gap-2"><Banknote className="text-emerald-600" size={32}/> Report Contabilità: Insoluti [INS]</h1>
                 <p className="text-slate-500 font-bold mt-1">Verifica di correttezza dati estratti dai fogli di viaggio e incrociati con la rubrica contatti.</p>
              </div>
              <div className="flex gap-4">
                 <button onClick={() => window.print()} className="bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-2.5 rounded-xl font-bold flex items-center gap-2 shadow-sm"><Printer size={18} /> Stampa Report PDF</button>
                 <button onClick={() => setListaInsoluti(null)} className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-6 py-2.5 rounded-xl font-bold flex items-center gap-2 border border-slate-300"><X size={18} /> Chiudi</button>
              </div>
           </div>

           <div className="hidden print:flex border-b-2 border-black pb-3 mb-6 justify-between items-end">
               <div>
                 <h1 className="text-2xl font-extrabold uppercase tracking-tight">REPORT CONTABILITA': INSOLUTI [INS]</h1>
                 <p className="text-gray-600 font-bold mt-1 uppercase">Stampa generata in data: {new Date().toLocaleDateString('it-IT')}</p>
               </div>
               <div className="text-right">
                 <p className="text-xs uppercase text-gray-500 font-bold mb-1">Totale insoluti rilevati</p>
                 <p className="text-lg font-extrabold text-red-600">{listaInsoluti.length} posizioni</p>
               </div>
            </div>

           <table className="w-full text-left text-sm border-collapse border border-slate-900 print:border-black">
              <thead>
                 <tr className="bg-slate-200 border-y-2 border-slate-900 print:border-black print:bg-slate-200">
                    <th className="p-3 uppercase text-xs border border-slate-400 print:border-black w-[35%]">Anagrafica Cliente</th>
                    <th className="p-3 uppercase text-xs border border-slate-400 print:border-black w-[20%]">Contatti</th>
                    <th className="p-3 uppercase text-xs border border-slate-400 print:border-black w-[10%]">Fattura / Pagamento</th>
                    <th className="p-3 uppercase text-xs text-right border border-slate-400 print:border-black w-[10%]">Importo</th>
                    <th className="p-3 uppercase text-xs border border-slate-400 print:border-black w-[25%]">Note Contabilità</th>
                 </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 print:divide-black">
                 {listaInsoluti.map((ins, i) => (
                    <tr key={i} className="hover:bg-slate-50 print:hover:bg-transparent">
                       <td className="p-3 border border-slate-300 print:border-black align-top">
                         <span className="font-extrabold text-slate-900 block">{ins.anagrafica}</span>
                         <span className="text-[10px] text-slate-500 font-mono mt-1 block">Cod: {ins.codiceCliente}</span>
                       </td>
                       <td className="p-3 border border-slate-300 print:border-black align-top text-[11px] text-slate-700">
                           {ins.contatto && <div className="flex items-center gap-1 mb-1 font-mono"><User size={10} /> {ins.contatto}</div>}
                           {ins.telefono && <div className="flex items-center gap-1 mb-1 font-mono"><Phone size={10} /> {ins.telefono}</div>}
                           {ins.email && <div className="flex items-center gap-1 font-mono break-all"><Mail size={10} className="shrink-0" /> {ins.email}</div>}
                           {!ins.telefono && !ins.email && !ins.contatto && <span className="text-slate-400 italic">Nessun contatto</span>}
                       </td>
                       <td className="p-3 border border-slate-300 print:border-black align-top">
                          <span className="font-mono font-bold block">{ins.numFattura}</span>
                          <span className="text-[9px] text-slate-500 font-bold uppercase block mt-1">{ins.modalitaPagamento}</span>
                       </td>
                       <td className="p-3 border border-slate-300 print:border-black align-top font-mono font-extrabold text-red-600 text-right">€ {ins.importo.toFixed(2)}</td>
                       <td className="p-3 border border-slate-300 print:border-black align-top text-xs italic text-slate-500 whitespace-pre-wrap">{ins.note || "-"}</td>
                    </tr>
                 ))}
                 {listaInsoluti.length === 0 && <tr><td colSpan={5} className="p-6 text-center border border-slate-300 print:border-black text-slate-500 font-bold">Nessun insoluto rilevato nel file caricato.</td></tr>}
              </tbody>
           </table>
        </div>
     );
  }

  return (
    <>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" integrity="sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY=" crossOrigin="" />
    
    <style dangerouslySetInnerHTML={{__html: `
      @media print {
        body { background-color: white !important; margin: 0; padding: 0; }
        @page { size: auto; margin: 10mm; }
        .break-after-page { break-after: page; page-break-after: always; }
        * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
      }
    `}} />

    <main className="min-h-screen bg-slate-100 text-slate-900 font-sans p-6 md:p-10 overflow-x-hidden print:hidden">
      <div className="max-w-7xl mx-auto">
        
        {/* HEADER */}
        <header className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-8 gap-4">
          <div>
            <h1 className="text-2xl md:text-3xl font-extrabold tracking-tight text-slate-900 flex items-center gap-3">
              <span className="p-2.5 bg-blue-600 text-white rounded-xl shadow-md"><Zap size={24} /></span>
              Asso Antincendio <span className="text-blue-600 font-normal text-xl">| Smart Logistics</span>
            </h1>
            <p className="text-xs text-slate-500 mt-1">Rubrica in memoria, Foglio Viaggio e Pausa Pranzo Dinamica.</p>
          </div>

          <div className="flex flex-col gap-2 w-full md:w-auto">
            <div className="flex items-center gap-4">
              {isCheckingAuth ? (
                <div className="bg-slate-50 text-slate-500 px-5 py-2.5 rounded-xl font-bold flex items-center justify-center gap-2 border border-slate-200 text-sm w-full"><Loader2 size={18} className="animate-spin text-blue-600" /> Verifica connessione...</div>
              ) : !googleToken ? (
                <button onClick={eseguiLoginGoogle} className="bg-slate-900 hover:bg-slate-800 text-white px-5 py-2.5 rounded-xl font-bold flex items-center justify-center gap-2 shadow-sm text-sm w-full"><Calendar size={18} /> Avvia Connessione Google</button>
              ) : (
                <div className="flex items-center gap-2 w-full">
                  <span className="bg-emerald-50 text-emerald-800 px-4 py-2.5 rounded-xl font-bold flex items-center gap-2 border border-emerald-200 text-sm flex-1"><Check size={18} /> Calendar Connesso</span>
                  <button onClick={logoutGoogle} className="bg-slate-100 hover:bg-red-50 text-slate-600 hover:text-red-600 px-3 py-2.5 rounded-xl border border-slate-200 shrink-0"><X size={18} /></button>
                </div>
              )}
            </div>
            
            <label className={`w-full py-2.5 rounded-xl font-bold text-xs shadow-sm flex items-center justify-center gap-2 transition-all cursor-pointer ${Object.keys(rubricaClienti).length > 0 ? 'bg-amber-100 text-amber-800 hover:bg-amber-200 border border-amber-200' : 'bg-slate-100 text-slate-600 hover:bg-slate-200 border border-slate-300'}`}>
              <Book size={16} />
              {Object.keys(rubricaClienti).length > 0 ? `Aggiorna Rubrica Clienti (${Object.keys(rubricaClienti).length} salvati)` : "Inserisci Rubrica Clienti"}
              <input type="file" accept=".xlsx, .xls, .csv, .txt" className="hidden" onChange={handleRubricaUpload} />
            </label>
          </div>
        </header>

        {/* PANNELLO CONFIG */}
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
             <div className="border-2 border-dashed border-slate-300 rounded-xl p-4 text-center bg-slate-50 flex flex-col items-center justify-center transition-all hover:border-blue-400 mb-4 h-full relative">
               <div className="p-2 bg-blue-100 text-blue-600 rounded-xl mb-1"><FileSpreadsheet size={20} /></div>
               <h2 className="text-sm font-bold text-slate-800 mb-1 truncate px-2 w-full max-w-[200px]" title={nomeFileCorrente}>{datiGrezziCaricati.length > 0 ? nomeFileCorrente : "Carica file cantieri"}</h2>
               <label className={`bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-xl font-bold text-xs shadow-md transition-all inline-block mt-1 ${!googleToken ? 'opacity-50 cursor-not-allowed' : 'cursor-pointer'}`}>
                 {datiGrezziCaricati.length > 0 ? "Sostituisci" : "Sfoglia"}
                 <input type="file" accept=".xlsx, .xls, .csv, .txt" className="hidden" onChange={handleFileUpload} disabled={!googleToken} />
               </label>
             </div>

             <div className="border-t border-slate-100 pt-3 mt-3 flex flex-col gap-1.5">
               <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1"><UserCheck size={14} className="text-blue-600"/> 2. Tecnico Operativo</label>
               <select value={tecnicoSelezionato.nome} onChange={(e) => { const t = tecniciAnagrafica.find(t => t.nome === e.target.value); if(t) setTecnicoSelezionato(t); }} className="w-full bg-slate-50 border border-slate-300 rounded-xl p-2 text-xs font-bold text-slate-800 focus:outline-none cursor-pointer">
                  {tecniciAnagrafica.map((t, idx) => <option key={idx} value={t.nome}>{t.nome}</option>)}
               </select>
             </div>
          </div>

          <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col justify-between items-center text-center">
             <div className="w-full">
               <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center justify-center gap-2 mb-2"><PlayCircle className="text-emerald-600" size={18} /> 3. Avvio Analisi Automatica</h2>
               <p className="text-xs text-slate-500 mb-4">Incrocia percorsi, pausa pranzo 12:30-14:00, orari e interventi pregressi.</p>
             </div>
             
             <div className="w-full flex flex-col gap-3 mt-auto">
                 <button
                   onClick={() => processaRighe(datiGrezziCaricati)}
                   disabled={datiGrezziCaricati.length === 0 || inElaborazione}
                   className={`w-full py-3 px-4 rounded-xl text-sm font-extrabold transition-all flex items-center justify-center gap-2 ${(datiGrezziCaricati.length > 0 && !inElaborazione) ? 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-md' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}
                 >
                   {inElaborazione ? <><Loader2 size={16} className="animate-spin"/> Elaborazione...</> : <><Sliders size={16} /> Analizza & Crea Orari</>}
                 </button>
                 
                 <label className="w-full py-2.5 px-4 rounded-xl text-[12px] font-bold transition-all flex items-center justify-center gap-2 border border-indigo-700 cursor-pointer bg-indigo-600 text-white hover:bg-indigo-700 shadow-md">
                   <Search size={16} /> Verifica Insoluti (File Globale)
                   <input type="file" accept=".xlsx, .xls, .csv, .txt" className="hidden" onChange={handleInsolutiUpload} />
                 </label>
             </div>
          </div>
        </div>

        {/* GOOGLE CALENDAR E SOSPESI */}
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6 mb-8 items-start">
          
          <div className="bg-amber-50 border border-amber-300 rounded-2xl p-5 shadow-sm flex flex-col h-[600px]">
            <h3 className="text-sm font-extrabold text-amber-900 flex items-center gap-2 mb-1"><PauseCircle size={18} className="text-amber-600" /> Sospesi ({clientiInSospeso.length})</h3>
            <p className="text-[11px] text-amber-800 mb-4 leading-tight">Clicca per forzare pianificazione.</p>
            
            <div className="space-y-3 overflow-y-auto pr-2 flex-1">
              {clientiInSospeso.map((s, idx) => {
                  const isSosBloccato = s.resoEstintori && s.giorniDallUltimoSos < 5;

                  return (
                    <div 
                        key={idx} 
                        onClick={() => !isSosBloccato && apriModaleSospeso(s)}
                        className={`bg-white border p-3 rounded-xl shadow-sm flex flex-col gap-1 transition-all ${isSosBloccato ? 'border-red-300 bg-red-50/40 opacity-70 cursor-not-allowed' : 'border-amber-200 hover:border-blue-400 hover:shadow-md cursor-pointer'}`}
                    >
                        <div className="flex justify-between items-start">
                            <span className="bg-amber-100 text-amber-900 font-mono text-[9px] px-2 py-0.5 rounded font-bold">BI: {s.numeroBi}</span>
                            <span className="text-[10px] text-slate-500 font-bold flex items-center gap-1"><Timer size={10} className="text-amber-600"/> {s.minutiLavoro}m</span>
                        </div>
                        <h4 className="font-bold text-slate-900 text-[11px] leading-tight mt-1">{s.nome}</h4>
                        <p className="text-[10px] text-slate-500 truncate mt-0.5">{s.indirizzo}</p>
                        {isSosBloccato && <span className="text-[9px] text-red-600 font-bold flex items-start gap-1 leading-tight mt-1"><AlertTriangle size={10} className="flex-shrink-0" /> {s.motivoSospeso}</span>}
                        {!isSosBloccato && s.motivoSospeso === "Tempo esaurito (Giornate Piene)" && <span className="text-[9px] text-amber-600 font-bold flex items-start gap-1 leading-tight mt-1"><AlertTriangle size={10} className="flex-shrink-0" /> Turni esauriti</span>}
                    </div>
                  );
              })}
              {clientiInSospeso.length === 0 && <div className="text-center text-xs text-amber-700/60 mt-10 font-medium">Nessun sospeso.</div>}
            </div>
          </div>

          <div className="lg:col-span-3 flex flex-col gap-6">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col h-[600px] overflow-hidden">
              <div className="mb-4 flex justify-between items-center shrink-0">
                <div>
                   <h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-2 mb-1"><Calendar className="text-blue-600" size={18} /> Google Calendar ({tecnicoSelezionato.nome})</h3>
                   <p className="text-[11px] text-slate-500 font-bold">Settimana visualizzata: {formattaDataVisuale(dataInizio)} - {formattaDataVisuale(dataFine)}</p>
                </div>
              </div>
              <div className="w-full flex-1 rounded-xl overflow-hidden border border-slate-200 bg-slate-50">
                <iframe src={calendarEmbedUrl} style={{ border: 0 }} width="100%" height="100%" frameBorder="0" scrolling="no"></iframe>
              </div>
            </div>
          </div>
        </div>

        {/* BOX RIEPILOGO PIANIFICAZIONE */}
        {interventiGrezzi.length > 0 && leafletLoaded && !inElaborazione && (
          <div className="w-full flex flex-col gap-6 mb-8">
             <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col lg:flex-row gap-8 items-start">
                <div className="flex-1 w-full">
                   <h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-2 mb-5"><CheckCircle size={18} className="text-emerald-600" /> Riepilogo Pianificazione</h3>
                   <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-5">
                     <div className="p-3 bg-emerald-50 rounded-xl border border-emerald-100 flex flex-col justify-center"><span className="block text-[10px] text-emerald-700 font-bold uppercase tracking-wider mb-0.5">Pianificati Nuovi</span><span className="text-2xl font-extrabold text-emerald-900">{clientiAttiviCount}</span></div>
                     <div className="p-3 bg-amber-50 rounded-xl border border-amber-100 flex flex-col justify-center"><span className="block text-[10px] text-amber-700 font-bold uppercase tracking-wider mb-0.5">Finiti nei Sospesi</span><span className="text-2xl font-extrabold text-amber-900">{clientiInSospeso.length}</span></div>
                     <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex flex-col justify-center"><span className="block text-[10px] text-slate-500 font-bold uppercase tracking-wider mb-0.5">Scartati / Esistenti</span><span className="text-2xl font-extrabold text-slate-700">{clientiGiaCalendarizzati.length}</span></div>
                   </div>

                   {bulkSyncStatus.active ? (
                      <div className="bg-blue-50 border border-blue-200 p-3 rounded-xl flex flex-col gap-2 w-full">
                         <div className="flex justify-between text-[11px] font-bold text-blue-800">
                            <span className="flex items-center gap-1.5"><Loader2 size={12} className="animate-spin"/> Invio in corso...</span>
                            <span>{bulkSyncStatus.current} / {bulkSyncStatus.total}</span>
                         </div>
                         <div className="w-full bg-blue-200 rounded-full h-2">
                            <div className="bg-blue-600 h-2 rounded-full transition-all duration-300" style={{width: `${bulkSyncStatus.progress}%`}}></div>
                         </div>
                      </div>
                    ) : (
                      <button onClick={inviaPianificazioneAGoogle} className="w-full md:w-auto bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-3 rounded-xl font-bold flex items-center justify-center gap-2 shadow-md text-xs transition-all">
                        <UploadCloud size={16} /> Invia a Google ({interventiGrezzi.filter(i => i.selezionatoPerGiro && !i.syncedToGoogle && !i.isPregresso).length} pronti)
                      </button>
                    )}
                </div>

                <div className="flex-1 w-full lg:border-l border-slate-100 lg:pl-8 pt-6 lg:pt-0 border-t lg:border-t-0">
                   <div className="flex justify-between items-center mb-4">
                     <p className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Filtra per giornata:</p>
                     <button onClick={() => setGiornoSceltoFiltro("Tutti")} className={`text-xs px-4 py-2 rounded-lg font-bold transition-all ${giornoSceltoFiltro === "Tutti" ? "bg-blue-600 text-white shadow-md" : "bg-slate-100 text-slate-600 hover:bg-slate-200 border border-slate-200"}`}>Mostra Tutti</button>
                   </div>
                   
                   <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
                     {giornateStats.map((g, idx) => {
                       const hh = String(Math.floor(g.orarioPartenzaMinuti / 60)).padStart(2, '0');
                       const mm = String(g.orarioPartenzaMinuti % 60).padStart(2, '0');
                       
                       return (
                         <div key={idx} onClick={() => setGiornoSceltoFiltro(g.dataStr)} className={`p-3 border rounded-xl flex flex-col gap-0.5 cursor-pointer transition-all ${g.statoGiornata === "APERTO" ? 'bg-emerald-50/50 border-emerald-200 hover:bg-emerald-100' : 'bg-red-50/50 border-red-200 hover:bg-red-100'} ${giornoSceltoFiltro === g.dataStr ? 'ring-2 ring-blue-500 shadow-md scale-[1.02]' : 'opacity-80 hover:opacity-100'}`}>
                           <span className="text-[12px] font-bold text-slate-800">{formattaDataVisuale(g.dataStr)}</span>
                           <span className="text-[10px] text-slate-500">{g.nomeGiorno}</span>
                           <div className="flex items-center gap-1.5 mt-1.5 font-mono text-[10px]">
                              {g.statoGiornata === "APERTO" ? <CheckCircle size={12} className="text-emerald-500" /> : <Lock size={12} className="text-red-500" />}
                              <span className={g.statoGiornata === "APERTO" ? "text-emerald-700 font-bold" : "text-red-700 font-bold"}>
                                 Fine st.: {hh}:{mm}
                              </span>
                           </div>
                         </div>
                       )
                     })}
                   </div>
                </div>
             </div>
          </div>
        )}

        {/* TABELLA E MAPPA */}
        {interventiGrezzi.length > 0 && leafletLoaded && !inElaborazione && (
          <div className="w-full flex flex-col gap-6">
            
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="p-6 border-b border-slate-200 flex flex-col md:flex-row justify-between items-start md:items-center gap-2">
                <div>
                  <h3 className="text-lg font-extrabold text-slate-900 flex items-center gap-2"><Package className="text-blue-600"/> Lista Pianificati e Distinte | {tecnicoSelezionato.nome}</h3>
                  <p className="text-xs text-slate-500 mt-1">Trascina le righe per riordinarle. Puoi deselezionarle (chiudendo l'orario) o buttarle nei Sospesi col cestino!</p>
                </div>
                <button onClick={() => window.print()} className="bg-slate-900 hover:bg-slate-800 text-white px-5 py-2.5 rounded-xl font-bold flex items-center gap-2 shadow-md text-sm transition-all"><Printer size={16} /> Stampa Foglio di Viaggio</button>
              </div>
              
              <div className="overflow-x-auto overflow-y-auto max-h-[600px] relative border-b border-slate-200 bg-white">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-50 border-b border-slate-200 text-xs uppercase text-slate-500 tracking-wider sticky top-0 z-10 shadow-sm">
                    <tr>
                      <th className="p-4 w-12 text-center"></th>
                      <th className="p-4 w-28 text-center">Giro</th>
                      <th className="p-4">Riferimento BI & Cliente</th>
                      <th className="p-4">Data e Orario</th>
                      <th className="p-4 text-center">Scheda</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {interventiVisibiliArr.map((intervento) => {
                      
                      if (intervento.isDistinta) {
                        return (
                          <tr key={intervento.codice} className="bg-orange-50/50 border-b border-orange-100">
                             <td className="p-4"></td>
                             <td className="p-4 text-center"><Package size={20} className="text-orange-500 mx-auto" /></td>
                             <td className="p-4">
                               <span className="font-extrabold text-orange-900 text-sm block">{intervento.nome}</span>
                               <span className="text-[10px] text-orange-700 font-bold block mt-1 uppercase tracking-wider">Basato sull'analisi a cascata dei Muletti.</span>
                             </td>
                             <td className="p-4">
                               <span className="block text-[11px] text-slate-500 font-mono font-bold flex items-center gap-1.5 mb-1"><Clock size={12}/> {intervento.oraInizio} - {intervento.oraFine}</span>
                               <span className="bg-orange-100 text-orange-800 px-3 py-1.5 rounded-xl text-xs font-bold border border-orange-200 inline-block">{formattaDataVisuale(intervento.dataAssegnata)} ({intervento.giorno})</span>
                             </td>
                             <td className="p-4 text-center">
                               {intervento.syncedToGoogle && <span className="bg-emerald-50 text-emerald-700 px-2 py-1.5 rounded-lg text-[10px] font-bold inline-flex items-center gap-1 border border-emerald-200 mb-1"><CheckCircle size={10} /> Inviato a Calendar</span>}
                               <button onClick={() => setClienteSelezionatoScheda(intervento)} className="bg-orange-200 hover:bg-orange-300 text-orange-900 px-3.5 py-2 rounded-xl text-xs font-bold inline-flex items-center gap-1.5 shadow-sm block mx-auto mt-1"><FileText size={14} /> Dettagli Carico</button>
                             </td>
                          </tr>
                        )
                      }

                      return (
                        <tr 
                          key={intervento.codice} 
                          draggable
                          onDragStart={(e) => handleDragStartReorder(e, intervento.codice)}
                          onDragOver={handleDragOverReorder}
                          onDrop={(e) => handleDropReorder(e, intervento.codice)}
                          className={`hover:bg-slate-50 transition-colors cursor-grab active:cursor-grabbing ${!intervento.selezionatoPerGiro ? 'opacity-40 bg-slate-50/80' : ''} ${intervento.isPregresso ? 'bg-indigo-50/30' : ''}`}
                          title="Trascina per riordinare e ricalcolare i tempi"
                        >
                          <td className="p-4 text-center text-slate-300 hover:text-slate-500 align-middle">
                            <GripVertical size={20} className="mx-auto" />
                          </td>
                          <td className="p-4 text-center align-middle">
                            <div className="flex flex-col items-center gap-1.5">
                               {intervento.numProgressivo !== null && !intervento.isPregresso && (
                                  <span className="bg-blue-100 text-blue-800 font-extrabold w-6 h-6 rounded-full flex items-center justify-center text-[11px] shadow-sm">
                                      {intervento.numProgressivo}
                                  </span>
                               )}
                               {!intervento.isPregresso && (
                                  <div className="flex items-center gap-2 mt-1">
                                      <button onClick={() => toggleSelezioneCliente(intervento.codice)} className="text-blue-600 hover:scale-110 transition-transform" title="Deseleziona / Escludi orario">
                                          {intervento.selezionatoPerGiro ? <CheckSquare size={18} /> : <Square size={18} className="text-slate-400" />}
                                      </button>
                                      <button onClick={() => rimuoviDaGiornata(intervento.codice)} className="text-red-500 hover:text-red-700 hover:scale-110 transition-transform bg-red-50 p-1.5 rounded-lg border border-red-100 shadow-sm" title="Rimuovi dalla giornata (Sposta in Sospesi)">
                                          <Trash2 size={15} />
                                      </button>
                                  </div>
                               )}
                               {intervento.isPregresso && <CalendarPlus size={18} className="text-indigo-400 mt-1" title="Evento Pregresso" />}
                            </div>
                          </td>
                          <td className="p-4">
                            <span className="font-bold text-slate-900 text-sm block leading-tight">{generaTitoloEvento(intervento)}</span>
                            {intervento.isPregresso && <span className="bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-xl text-[9px] font-extrabold block w-max mt-1.5 border border-indigo-200">📌 GIÀ A CALENDARIO</span>}
                            <p className="text-slate-500 text-[11px] mt-1">{intervento.indirizzo} {intervento.localita ? `- ${intervento.localita}` : ""}</p>
                          </td>
                          <td className="p-4">
                            <span className="block text-[11px] text-slate-500 font-mono font-bold flex items-center gap-1.5 mb-1"><Clock size={12} className={intervento.isPregresso ? "text-indigo-500" : ""}/> {intervento.oraInizio} - {intervento.oraFine}</span>
                            <span className="bg-emerald-50 text-emerald-700 px-2 py-1 rounded-lg text-[10px] font-bold inline-flex items-center gap-1 w-max border border-emerald-100 mr-2"><Timer size={10} /> {intervento.minutiLavoro} min {intervento.isPregresso ? 'stimati' : 'lavoro'}</span>
                            {!intervento.isPregresso && intervento.resoEstintori && (
                               <span className="bg-blue-50 text-blue-700 px-2 py-1 rounded-lg text-[10px] font-bold inline-flex items-center gap-1 w-max border border-blue-200">Muletti: +{intervento.muletti_eorv}</span>
                            )}
                            {!intervento.isPregresso && !intervento.resoEstintori && (
                               <span className="bg-amber-50 text-amber-700 px-2 py-1 rounded-lg text-[10px] font-bold inline-flex items-center gap-1 w-max border border-amber-200">Muletti: -{intervento.muletti_eorv}</span>
                            )}
                          </td>
                          <td className="p-4 text-center">
                             {!intervento.isPregresso && intervento.syncedToGoogle && (
                               <span className="bg-emerald-50 text-emerald-700 px-2 py-1 rounded-lg text-[10px] font-bold inline-flex items-center gap-1 border border-emerald-200 mb-1"><CheckCircle size={10} /> Inviato</span>
                             )}
                             {!intervento.isPregresso && !intervento.syncedToGoogle && (
                               <span className="bg-slate-50 text-slate-500 px-2 py-1 rounded-lg text-[10px] font-bold inline-flex items-center gap-1 border border-slate-200 mb-1"><Clock size={10} /> In Attesa</span>
                             )}
                             <button onClick={() => setClienteSelezionatoScheda(intervento)} className="bg-slate-100 hover:bg-blue-50 text-slate-700 px-3 py-1.5 rounded-lg text-[10px] font-bold inline-flex items-center gap-1.5 block mx-auto mt-1"><FileText size={12} /> Info</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
              <h3 className="text-base font-bold mb-4 flex items-center gap-2 text-slate-900"><MapPin className="text-red-500" /> Mappa Percorso Ordinato ({giornoSceltoFiltro === "Tutti" ? "Tutti i Giorni" : formattaDataVisuale(giornoSceltoFiltro)})</h3>
              <div className="h-[480px] w-full rounded-xl overflow-hidden z-0 border border-slate-100">
                <MapContainer key={giornoSceltoFiltro} center={[SEDE_UFFICIO_LAT, SEDE_UFFICIO_LNG]} zoom={12} style={{ height: "100%", width: "100%" }}>
                  <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
                  {interventiVisibiliArr.filter(i => i.lat && i.lng && !i.isDistinta && i.selezionatoPerGiro && !i.isPregresso).map((c) => (
                    <Marker key={c.codice} position={[c.lat, c.lng]} icon={createNumberedIcon(c.numProgressivo)}>
                      <Popup>
                        <div className="p-2">
                          <p className="font-bold text-sm text-slate-900">#{c.numProgressivo} - {c.nome}</p>
                          <p className="text-[11px] text-blue-600 font-mono font-bold mb-1">BI: {c.numeroBi}</p>
                          <p className="text-xs text-slate-600 mb-2">{c.indirizzo}, {c.localita}</p>
                        </div>
                      </Popup>
                    </Marker>
                  ))}
                  {/* Marker Pregressi Mappa */}
                  {interventiVisibiliArr.filter(i => i.lat && i.lng && !i.isDistinta && i.isPregresso).map((c) => (
                    <Marker key={c.codice} position={[c.lat, c.lng]}>
                      <Popup>
                        <div className="p-2">
                          <span className="bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-xl text-[9px] font-extrabold inline-block mb-1 border border-indigo-200">📌 GIÀ A CALENDARIO</span>
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

          </div>
        )}

        {/* MODALE ASSEGNAZIONE SOSPESI */}
        {sospesoInModifica && (
          <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-md z-50 flex items-center justify-center p-4 animate-in fade-in duration-200">
            <div className="bg-white rounded-3xl shadow-2xl max-w-sm w-full p-8 relative border border-slate-200">
              <button onClick={() => setSospesoInModifica(null)} className="absolute top-6 right-6 text-slate-400 bg-slate-100 p-2 rounded-full hover:bg-slate-200"><X size={18} /></button>
              
              <div className="mb-6">
                 <div className="p-3 bg-amber-100 text-amber-600 rounded-xl inline-block mb-4"><PauseCircle size={24} /></div>
                 <h2 className="text-xl font-extrabold text-slate-900 leading-tight">Pianifica Sospeso</h2>
                 <p className="text-xs text-slate-500 mt-1">{sospesoInModifica.nome}</p>
              </div>

              <div className="flex flex-col gap-4 mb-6">
                <div>
                  <label className="text-[11px] font-bold text-slate-500 block mb-1">Data assegnazione</label>
                  <input type="date" value={configSospesoSingolo.data} onChange={(e) => setConfigSospesoSingolo({...configSospesoSingolo, data: e.target.value})} className="w-full bg-slate-50 border border-slate-300 rounded-xl p-3 text-sm font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-400" />
                </div>
                <div>
                  <label className="text-[11px] font-bold text-slate-500 block mb-1">Ora Inizio stimata</label>
                  <input type="time" value={configSospesoSingolo.oraInizio} onChange={(e) => setConfigSospesoSingolo({...configSospesoSingolo, oraInizio: e.target.value})} className="w-full bg-slate-50 border border-slate-300 rounded-xl p-3 text-sm font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-400" />
                </div>
              </div>

              <div className="flex flex-col gap-2.5">
                 <button onClick={() => creaSospesoDaModale('calendar')} disabled={creazioneInCorso[sospesoInModifica.codice]} className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3.5 rounded-xl text-sm font-bold flex items-center justify-center gap-2 shadow-md transition-all">
                    {creazioneInCorso[sospesoInModifica.codice] ? <><Loader2 size={16} className="animate-spin"/> Invio in corso...</> : <><CalendarPlus size={16}/> Crea su Calendar</>}
                 </button>
                 <button onClick={() => creaSospesoDaModale('lista')} disabled={creazioneInCorso[sospesoInModifica.codice]} className="w-full bg-slate-100 hover:bg-slate-200 text-slate-700 py-3.5 rounded-xl text-sm font-bold flex items-center justify-center gap-2 transition-all">
                    <CheckSquare size={16}/> Aggiungi solo in Lista
                 </button>
              </div>
            </div>
          </div>
        )}

        {/* MODALE SCHEDA TECNICA CLIENTE E DISTINTA */}
        {clienteSelezionatoScheda && (
          <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-md z-50 flex items-center justify-center p-4 animate-in fade-in duration-200 overflow-y-auto">
            <div className="bg-white rounded-3xl shadow-2xl max-w-2xl w-full p-8 relative border border-slate-200 my-auto">
              <button onClick={() => setClienteSelezionatoScheda(null)} className="absolute top-6 right-6 text-slate-400 bg-slate-100 p-2.5 rounded-full hover:bg-slate-200"><X size={20} /></button>
              
              {clienteSelezionatoScheda.isDistinta ? (
                <>
                  <div className="flex items-center gap-4 mb-6">
                    <div className="p-4 bg-orange-100 text-orange-600 rounded-2xl"><Package size={32} /></div>
                    <div>
                      <span className="bg-orange-100 text-orange-800 font-mono text-xs px-3 py-1 rounded-xl font-bold mb-1 inline-block">REPORT MAGAZZINO</span>
                      <h2 className="text-xl font-extrabold text-slate-900">Distinta di Carico</h2>
                      <p className="text-xs text-slate-500">Giornata lavorativa: {formattaDataVisuale(clienteSelezionatoScheda.dataAssegnata)}</p>
                    </div>
                  </div>
                  <div className="bg-slate-50 p-6 rounded-2xl border border-slate-200 text-sm whitespace-pre-wrap font-mono leading-relaxed mb-6">
                    {clienteSelezionatoScheda.descrizioneDistinta}
                  </div>
                </>
              ) : (
                <>
                  <div className="flex items-center gap-4 mb-6">
                    <div className="p-4 bg-blue-50 text-blue-600 rounded-2xl"><FileText size={32} /></div>
                    <div>
                      <span className="bg-blue-100 text-blue-800 font-mono text-xs px-3 py-1 rounded-xl font-bold mb-1 inline-block">BI: {clienteSelezionatoScheda.numeroBi}</span>
                      <h2 className="text-2xl font-extrabold text-slate-900">{clienteSelezionatoScheda.nome}</h2>
                      <p className="text-xs text-slate-500">{clienteSelezionatoScheda.indirizzo}, {clienteSelezionatoScheda.localita}</p>
                      {clienteSelezionatoScheda.isPregresso && <span className="bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-xl text-[9px] font-extrabold inline-block mt-2 border border-indigo-200">📌 PRE-ESISTENTE IN CALENDAR</span>}
                    </div>
                  </div>
                  
                  <div className="grid grid-cols-2 gap-4 mb-6 bg-slate-50 p-5 rounded-2xl border border-slate-200 text-xs">
                    <div><span className="text-slate-400 uppercase font-bold text-[10px] block mb-0.5">Turno Assegnato</span><b className="text-slate-900 text-sm">{formattaDataVisuale(clienteSelezionatoScheda.dataAssegnata)} ({clienteSelezionatoScheda.giorno})</b></div>
                    <div><span className="text-slate-400 uppercase font-bold text-[10px] block mb-0.5">Tempo Stimato</span><b className="text-emerald-600 text-sm">{clienteSelezionatoScheda.minutiLavoro} min</b></div>
                  </div>
                  
                  {(clienteSelezionatoScheda.contatto || clienteSelezionatoScheda.telefono || clienteSelezionatoScheda.email) && (
                  <div className="mb-6 bg-slate-50 p-5 rounded-2xl border border-slate-200 text-xs">
                    <span className="text-slate-400 uppercase font-bold text-[10px] block mb-2">Contatti e Riferimenti</span>
                    {clienteSelezionatoScheda.contatto && <div className="flex items-center gap-2 mb-1.5"><User size={14} className="text-slate-400"/> <b className="text-slate-700">{clienteSelezionatoScheda.contatto}</b></div>}
                    {clienteSelezionatoScheda.telefono && <div className="flex items-center gap-2 mb-1.5"><Phone size={14} className="text-slate-400"/> <b className="text-slate-700">{clienteSelezionatoScheda.telefono}</b></div>}
                    {clienteSelezionatoScheda.email && <div className="flex items-center gap-2"><Mail size={14} className="text-slate-400"/> <b className="text-slate-700">{clienteSelezionatoScheda.email}</b></div>}
                  </div>
                  )}

                  {clienteSelezionatoScheda.isPregresso && (
                  <div className="mb-6 bg-indigo-50 p-5 rounded-2xl border border-indigo-200 text-xs text-indigo-900 whitespace-pre-wrap">
                    <span className="text-indigo-400 uppercase font-bold text-[10px] block mb-2">Note / Descrizione Calendario</span>
                    {clienteSelezionatoScheda.descrizionePregressa}
                  </div>
                  )}
                </>
              )}

              <div className="flex gap-4">
                <button onClick={() => setClienteSelezionatoScheda(null)} className="flex-1 bg-slate-900 text-white px-6 py-2.5 rounded-xl text-xs font-bold shadow-md hover:bg-slate-800 transition-all">Chiudi Modale</button>
              </div>
            </div>
          </div>
        )}

      </div>
    </main>

    {/* ========================================= */}
    {/* VISTA DEDICATA SOLO PER LA STAMPA CARTACEA */}
    {/* ========================================= */}
    <div className="hidden print:block w-full text-black bg-white font-sans text-sm">
       {giorniVisibiliPerStampa.map((dataStr) => {
          const interventiGiorno = interventiVisibiliArr.filter(i => i.dataAssegnata === dataStr && (i.selezionatoPerGiro || i.isPregresso));
          if (interventiGiorno.length === 0) return null;

          const distinta = interventiGiorno.find(i => i.isDistinta);
          const lavori = interventiGiorno.filter(i => !i.isDistinta).sort((a,b) => a.oraInizio.localeCompare(b.oraInizio));
          const nomeGiorno = getNomeGiorno(dataStr);

          return (
             <div key={dataStr} className="break-after-page mb-10">
                
                <div className="border-b-2 border-black pb-3 mb-6 flex justify-between items-end">
                   <div>
                     <h1 className="text-2xl font-extrabold uppercase tracking-tight">FOGLIO DI VIAGGIO</h1>
                     <p className="text-gray-600 font-bold mt-1 uppercase">{nomeGiorno} {formattaDataVisuale(dataStr)}</p>
                   </div>
                   <div className="text-right">
                     <p className="text-xs uppercase text-gray-500 font-bold mb-1">Tecnico Assegnato</p>
                     <p className="text-lg font-extrabold">{tecnicoSelezionato.nome}</p>
                   </div>
                </div>

                {distinta && (
                  <div className="mb-6 p-4 border border-black bg-gray-50 rounded-lg">
                     <h2 className="font-extrabold text-base mb-2 uppercase flex items-center gap-2"><Package size={18} /> Distinta di Carico Magazzino</h2>
                     <pre className="font-mono text-sm whitespace-pre-wrap">{distinta.descrizioneDistinta}</pre>
                  </div>
                )}

                <table className="w-full text-left border-collapse mt-4">
                   <thead>
                      <tr className="border-b-2 border-black bg-slate-200 print:bg-slate-200">
                         <th className="py-2 w-12 text-center uppercase text-xs">#</th>
                         <th className="py-2 w-24 uppercase text-xs">Orario</th>
                         <th className="py-2 w-28 uppercase text-xs">Num. BI</th>
                         <th className="py-2 uppercase text-xs">Cliente / Indirizzo</th>
                         <th className="py-2 w-56 uppercase text-xs">Contatti</th>
                      </tr>
                   </thead>
                   <tbody>
                      {lavori.map(lavoro => (
                         <tr key={lavoro.codice} className="border-b border-gray-300">
                            <td className="py-4 align-top text-center font-extrabold text-lg">
                              {lavoro.isPregresso ? "*" : lavoro.numProgressivo}
                            </td>
                            <td className="py-4 align-top">
                              <span className="font-extrabold text-sm">{lavoro.oraInizio}</span>
                              <br/><span className="text-xs text-gray-500">{lavoro.oraFine}</span>
                            </td>
                            <td className="py-4 align-top font-mono font-bold text-sm">
                              {lavoro.numeroBi}
                            </td>
                            <td className="py-4 align-top pr-4">
                               <div className="font-extrabold text-base leading-tight mb-1">{generaTitoloEvento(lavoro)}</div>
                               {lavoro.isPregresso && <span className="text-[10px] font-bold uppercase block mb-1">📌 Intervento pre-esistente</span>}
                               <div className="text-xs text-gray-600 font-medium">{lavoro.indirizzo}, {lavoro.localita}</div>
                            </td>
                            <td className="py-4 align-top text-xs text-gray-800">
                               {lavoro.contatto && <div className="flex items-center gap-1 mb-1 font-mono"><User size={12} /> {lavoro.contatto}</div>}
                               {lavoro.telefono && <div className="flex items-center gap-1 mb-1 font-mono"><Phone size={12} /> {lavoro.telefono}</div>}
                               {lavoro.email && <div className="flex items-center gap-1 font-mono"><Mail size={12} /> {lavoro.email}</div>}
                               {!lavoro.telefono && !lavoro.email && !lavoro.contatto && <span className="text-gray-400 italic">Nessun contatto</span>}
                            </td>
                         </tr>
                      ))}
                   </tbody>
                </table>

             </div>
          )
       })}
    </div>
    </>
  );
}