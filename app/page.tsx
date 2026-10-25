"use client";

import Link from "next/link";
import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import dynamic from "next/dynamic";
import {
  UploadCloud, AlertCircle, Package, MapPin, Loader2, FileSpreadsheet,
  Zap, UserCheck, Printer, Calendar, Clock, CheckSquare, Mail, Timer,
  FileText, X, Sliders, Check, PauseCircle, ExternalLink, AlertTriangle,
  CheckCircle, PlayCircle, Lock, CalendarPlus, Phone, Book, User,
  Banknote, Car, Home as HomeIcon, AlertOctagon, MessageSquare,
  ArrowRightLeft, List, FileUp, Wrench
} from "lucide-react";

const MapContainer = dynamic(() => import("react-leaflet").then((mod) => mod.MapContainer), { ssr: false });
const TileLayer = dynamic(() => import("react-leaflet").then((mod) => mod.TileLayer), { ssr: false });
const Marker = dynamic(() => import("react-leaflet").then((mod) => mod.Marker), { ssr: false });
const Popup = dynamic(() => import("react-leaflet").then((mod) => mod.Popup), { ssr: false });

const GOOGLE_CLIENT_ID = "645365149295-lk8ei43kt09hsm2csupf643tgqkqqgmo.apps.googleusercontent.com";
const GOOGLE_SCOPES = "https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/calendar.events";

const FINE_GIORNATA_ASSOLUTA = 18 * 60; // 18:00
const INIZIO_PAUSA = 12 * 60 + 30; // 12:30
const FINE_PAUSA = 14 * 60; // 14:00
const DURATA_PAUSA = 90;
const GOOGLE_MAPS_API_KEY = "AIzaSyBLfAmnm_kaHbUc0sAVzhvkwXDF14EFCro";
const SEDE_UFFICIO_LAT = 44.4056;
const SEDE_UFFICIO_LNG = 8.9463;

const formattaDataVisuale = (dataStr: string) => {
  if (!dataStr) return "";
  const parts = dataStr.split("-");
  if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
  return dataStr;
};

const OrarioInMinuti = (oraStr: string, fallback = 480) => {
  if (!oraStr || typeof oraStr !== "string" || !oraStr.includes(":")) return fallback;
  const [h, m] = oraStr.split(":").map(Number);
  if (isNaN(h) || isNaN(m)) return fallback;
  return h * 60 + m;
};

const MinutiInOrario = (min: number) => {
  const hh = String(Math.floor(min / 60)).padStart(2, "0");
  const mm = String(Math.floor(min % 60)).padStart(2, "0");
  return `${hh}:${mm}`;
};

const deveMostrareDC = (nomeCliente: string) => {
  const n = String(nomeCliente || "").toUpperCase();
  if (n.includes("ASL") || n.includes("ADE") || n.includes("ADER")) return false;
  return true;
};

const getNextWeekDays = () => {
  const today = new Date();
  const dayOfWeek = today.getDay(); 
  const daysToMonday = dayOfWeek === 0 ? 1 : (dayOfWeek === 6 ? 2 : 1 - dayOfWeek);
  
  const startMonday = new Date(today);
  startMonday.setDate(today.getDate() + daysToMonday);
  
  const endFriday = new Date(startMonday);
  endFriday.setDate(startMonday.getDate() + 4);

  const fmt = (d: Date) => {
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  };
  
  return { start: fmt(startMonday), end: fmt(endFriday) };
};

const getNomeGiorno = (dataStr: string) => {
  if (!dataStr || typeof dataStr !== 'string') return "Lunedì";
  const parts = dataStr.split("-");
  if (parts.length < 3) return "Lunedì";
  const [y, m, d] = parts.map(Number);
  const date = new Date(y, m - 1, d);
  const mapGiorni = ["Domenica", "Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato"];
  return mapGiorni[date.getDay()] || "Lunedì";
};

const getGiornateLavorative = (inizio: string, fine: string, attivi: { [key: string]: boolean }) => {
  const dates = [];
  if (!inizio || !fine) return dates;
  
  const partsI = inizio.split("-");
  const partsF = fine.split("-");
  if(partsI.length < 3 || partsF.length < 3) return dates;

  const [yI, mI, dI] = partsI.map(Number);
  const [yF, mF, dF] = partsF.map(Number);
  let curr = new Date(yI, mI - 1, dI);
  const end = new Date(yF, mF - 1, dF);
  const mapGiorni = ["Domenica", "Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato"];
  
  let limit = 0;
  while (curr <= end && limit < 365) {
    const nomeGiorno = mapGiorni[curr.getDay()];
    if (attivi[nomeGiorno]) {
      const y = curr.getFullYear(); 
      const m = String(curr.getMonth() + 1).padStart(2, '0'); 
      const d = String(curr.getDate()).padStart(2, '0');
      dates.push({ dataStr: `${y}-${m}-${d}`, nomeGiorno });
    }
    curr.setDate(curr.getDate() + 1);
    limit++;
  }
  return dates;
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

const calcolaTempistiche = (orarioPartenza: number | string, minViaggio: number | string, minutiLavoro: number | string) => {
    let inizioLavoro = Math.round(Number(orarioPartenza) + Number(minViaggio));
    
    if (inizioLavoro >= INIZIO_PAUSA && inizioLavoro < FINE_PAUSA) {
        inizioLavoro = FINE_PAUSA;
    }

    let fineLavoro = inizioLavoro + Number(minutiLavoro);

    if (inizioLavoro < INIZIO_PAUSA && fineLavoro > INIZIO_PAUSA) {
        fineLavoro += DURATA_PAUSA;
    }

    return { inizioLavoro, fineLavoro };
};

const calcolaDistanzaKm = (lat1: number, lon1: number, lat2: number, lon2: number) => {
    const R = 6371;
    const dLat = (lat2 - lat1) * (Math.PI / 180);
    const dLon = (lon2 - lon1) * (Math.PI / 180);
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(lat1 * (Math.PI / 180)) * Math.cos(lat2 * (Math.PI / 180)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
    return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
};

let googleMapsPromise: Promise<void> | null = null;
const loadGoogleMapsAPI = () => {
    if (googleMapsPromise) return googleMapsPromise;
    if (typeof window !== "undefined" && window.google && window.google.maps) {
        googleMapsPromise = Promise.resolve();
        return googleMapsPromise;
    }
    googleMapsPromise = new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = `https://maps.googleapis.com/maps/api/js?key=${GOOGLE_MAPS_API_KEY}`;
        script.async = true;
        script.defer = true;
        script.onload = () => resolve();
        script.onerror = () => reject(new Error("Google Maps SDK failed to load"));
        document.head.appendChild(script);
    });
    return googleMapsPromise;
};

const calcolaTempoDistanzaGoogle = async (lat1: number, lon1: number, lat2: number, lon2: number) => {
    if (Math.abs(lat1 - lat2) < 0.0001 && Math.abs(lon1 - lon2) < 0.0001) {
        return { minuti: 0, km: 0 };
    }

    const cacheKey = `${lat1.toFixed(4)},${lon1.toFixed(4)}-${lat2.toFixed(4)},${lon2.toFixed(4)}`;
    const cacheSalvata = localStorage.getItem("asso_routes_cache");
    const memoria = cacheSalvata ? JSON.parse(cacheSalvata) : {};

    if (memoria[cacheKey]) return memoria[cacheKey];

    try {
        await loadGoogleMapsAPI();
        const service = new window.google.maps.DistanceMatrixService();
        const origin = new window.google.maps.LatLng(lat1, lon1);
        const destination = new window.google.maps.LatLng(lat2, lon2);
        
        const response: any = await new Promise((resolve, reject) => {
            const timer = setTimeout(() => reject(new Error("Timeout DistanceMatrix")), 3000);
            service.getDistanceMatrix({
                origins: [origin],
                destinations: [destination],
                travelMode: window.google.maps.TravelMode.DRIVING,
            }, (res: any, status: any) => {
                clearTimeout(timer);
                if (status === "OK") resolve(res);
                else reject(status);
            });
        });

        if (response && response.rows && response.rows[0] && response.rows[0].elements[0].status === "OK") {
           const element = response.rows[0].elements[0];
           const secondi = element.duration.value;
           const metri = element.distance.value;
           const risultato = { minuti: Math.round(secondi / 60), km: Number((metri / 1000).toFixed(1)) };
           memoria[cacheKey] = risultato;
           localStorage.setItem("asso_routes_cache", JSON.stringify(memoria));
           return risultato;
        }
    } catch (e) {
        console.error("Fallback distanza lineare:", e);
    }
    
    const dist = calcolaDistanzaKm(lat1, lon1, lat2, lon2);
    return { minuti: Math.round(dist * 2), km: Number(dist.toFixed(1)) };
};

export default function Home() {
  const defaultDates = getNextWeekDays();
  const [dataInizio, setDataInizio] = useState<string>(defaultDates.start);
  const [dataFine, setDataFine] = useState<string>(defaultDates.end);

  const [interventiGrezzi, setInterventiGrezzi] = useState<any[]>([]);
  const [clientiInSospeso, setClientiInSospeso] = useState<any[]>([]);
  const [clientiGiaCalendarizzati, setClientiGiaCalendarizzati] = useState<any[]>([]);
  
  const [inElaborazione, setInElaborazione] = useState(false);
  const [progressIA, setProgressIA] = useState({ active: false, percent: 0, text: "" });
  const [menuAperto, setMenuAperto] = useState<string | null>(null);
  
  const [leafletLoaded, setLeafletLoaded] = useState(false);
  const leafletRef = useRef<any>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null); 
  
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

  const [giornoSceltoFiltro, setGiornoSceltoFiltro] = useState<string>("sospesi"); 
  const [clienteSelezionatoScheda, setClienteSelezionatoScheda] = useState<any | null>(null);
  
  const [sospesoInModifica, setSospesoInModifica] = useState<any | null>(null);
  const [configSospesoSingolo, setConfigSospesoSingolo] = useState<{data: string, oraInizio: string}>({data: dataInizio, oraInizio: "08:30"});

  const [giornateStats, setGiornateStats] = useState<any[]>([]);
  const [creazioneInCorso, setCreazioneInCorso] = useState<{[key:string]: boolean}>({});
  const [bulkSyncStatus, setBulkSyncStatus] = useState({ active: false, progress: 0, current: 0, total: 0 });

  const [isChatOpen, setIsChatOpen] = useState(false);
  const [chatInput, setChatInput] = useState('');
  const [messages, setMessages] = useState([
    { sender: 'ai', text: 'Ciao! Sono il tuo assistente operativo. Dimmi pure quali cantieri vuoi spostare, riordinare o inserire dai sospesi.' }
  ]);
  const [chatLoading, setChatLoading] = useState(false);
  
  const [dragOverCard, setDragOverCard] = useState<string | null>(null);

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
    const range = getGiornateLavorative(dataInizio, dataFine, giorniAttivi);
    const statsBase = range.map(g => ({
      ...g,
      minutiDisponibili: 480,
      statoGiornata: "APERTO",
      anchorLat: null,
      anchorLng: null,
      orarioPartenzaMinuti: 8 * 60,
      orarioFineMinuti: 8 * 60
    }));
    setGiornateStats(statsBase);
  }, [dataInizio, dataFine, giorniAttivi]);

  useEffect(() => {
    import("leaflet").then((L) => {
      leafletRef.current = L;
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
    if (isChatOpen) {
        messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
    }
  }, [messages, chatLoading, isChatOpen]);

  useEffect(() => {
    const initGoogleClient = () => {
      if (window.google && window.google.accounts) {
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

  const trovaCoordinateGoogle = async (indirizzo: string, localita: string) => {
    let cleanAddr = String(indirizzo).replace(/Nessun indirizzo valido letto/gi, '').trim();
    
    if (cleanAddr.toLowerCase().includes("acquasanta") || (localita && String(localita).toLowerCase().includes("acquasanta"))) {
      return { lat: 44.4568, lng: 8.7678, precisione: "strada", origine: "override_manuale" };
    }

    let citta = localita && String(localita).trim().length > 0 ? String(localita).trim() : "";
    if (cleanAddr.includes("-") && !citta) {
       const parts = cleanAddr.split("-");
       cleanAddr = parts[0].trim();
       citta = parts[1].trim();
    }

    const comuniProvincia = ["ronco scrivia", "busalla", "mignanego", "crocefieschi", "arenzano", "mele", "voltri", "sanremo", "imperia", "vado ligure", "savona", "pontedecimo", "bolzaneto", "campomorone", "ceranesi", "sant'olcese", "serra riccò"];
    const addrLower = cleanAddr.toLowerCase();
    
    if (!citta || citta.toLowerCase() === "genova") {
       for (const comune of comuniProvincia) {
          if (addrLower.includes(comune)) {
              citta = comune;
              break;
          }
       }
    }

    if (!citta) citta = "Genova";
    
    let query = "";
    if (cleanAddr.toLowerCase().includes("italia") || cleanAddr.toLowerCase().includes("italy")) {
        query = cleanAddr; 
    } else {
        query = `${cleanAddr}, ${citta}, Italy`;
    }

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
        
        let localitaEsattaDaNote = "";
        if (evt.description) {
            const matchLoc = evt.description.match(/Localit[aà]:\s*([^\n]+)/i);
            if (matchLoc && matchLoc[1]) localitaEsattaDaNote = matchLoc[1].split("-")[0].trim();
        }

        let latEvt: number | null = null; let lngEvt: number | null = null;
        if (evt.location && evt.location.trim().length > 2) {
          let addressToSearch = evt.location;
          if (localitaEsattaDaNote) addressToSearch += `, ${localitaEsattaDaNote}`;
          
          const coord = await trovaCoordinateGoogle(addressToSearch, localitaEsattaDaNote); 
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
            const fineMinutiLocali = Math.min(endD.getHours() * 60 + endD.getMinutes(), FINE_GIORNATA_ASSOLUTA);

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
    if (deveMostrareDC(cliente.nome)) tags.push("[DC]");
    
    const tagString = tags.length > 0 ? tags.join(" ") + " " : "";
    return `${tagString}${cliente.nome} - BI ${cliente.numeroBi}`;
  };

  const generaDescrizioneEvento = (cliente: any, tecnicoNome: string, tecnicoEmail: string, interventiGiorno: any[] = []) => {
    if (cliente.isDistinta) {
        const tappe = interventiGiorno.filter(i => !i.isDistinta && i.selezionatoPerGiro);
        let navUrl = "";
        if (tappe.length > 0) {
            const origin = encodeURIComponent(`${SEDE_UFFICIO_LAT},${SEDE_UFFICIO_LNG}`);
            const destination = encodeURIComponent(`${SEDE_UFFICIO_LAT},${SEDE_UFFICIO_LNG}`);
            const waypoints = tappe.map(t => encodeURIComponent(`${t.indirizzo}, ${t.localita || 'Genova'}, Italy`)).join("|");
            navUrl = `\n📍 MASTER LINK NAVIGAZIONE GIORNATA (Tutte le tappe):\nhttps://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${destination}&waypoints=${waypoints}\n`;
        }
        return `${cliente.descrizioneDistinta}\n${navUrl}`;
    }
    if (cliente.isPregresso) return cliente.descrizionePregressa;

    let attrezzatureTesto = cliente.dettaglioAttrezzature && cliente.dettaglioAttrezzature.length > 0 
      ? cliente.dettaglioAttrezzature.map((a: any) => `🔧 ${a.descrizione}: ${a.quantita}`).join("\n")
      : "Nessuna attrezzatura specificata";

    const contattoInfo = cliente.contatto ? `👤 Contatto: ${cliente.contatto}` : "";
    const telInfo = cliente.telefono ? `📞 ${cliente.telefono}` : "";
    const emailInfo = cliente.email ? `📧 Email: ${cliente.email}` : "";
    
    let stringaContatti = [contattoInfo, telInfo, emailInfo].filter(Boolean).join("\n");
    if (stringaContatti.length > 0) stringaContatti = `\n${stringaContatti}\n`;

    return `Località: ${cliente.cap ? cliente.cap + ' ' : ''}${cliente.localita || ''}
Indirizzo: ${cliente.indirizzo || ''}
${stringaContatti}
🛠 ATTREZZATURE DA VERIFICARE:
${attrezzatureTesto}

Durata: ${cliente.minutiLavoro} min
📍 Naviga:
https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${cliente.indirizzo}, ${cliente.cap ? cliente.cap + ' ' : ''}${cliente.localita || 'Genova'}, Italy`)}

Tecnico: ${tecnicoNome}`;
  };

  // MOTORE ORARIO CON INCASTRO MATTINA/POMERIGGIO ATTORNO AI PREGRESSI
  const ricalcolaDistinteEOrari = async (
      listaDaOrdinare: any[],
      statsLocali: any[] = giornateStats,
      usaOrdinamentoGeografico = false,
      scartaOvertimeAutomatico = false
  ) => {
      const listaPulita = listaDaOrdinare.filter(i => !i.isDistinta && i.dataAssegnata && typeof i.dataAssegnata === 'string');
      const mapGiorni = new globalThis.Map();
      const vecchieDistinte = listaDaOrdinare.filter(i => i.isDistinta);
      const nuoveStats = statsLocali.map(g => ({ ...g, orarioFineMinuti: 8 * 60, statoGiornata: "APERTO" }));
      
      const scartatiTemporanei: any[] = []; 

      listaPulita.forEach(i => {
        if (!mapGiorni.has(i.dataAssegnata)) mapGiorni.set(i.dataAssegnata, []);
        mapGiorni.get(i.dataAssegnata).push(i);
      });

      const nuovaLista: any[] = [];

      for (const [dataStr, interventiGiorno] of Array.from(mapGiorni.entries())) {
        let orarioCorrenteMinuti = 8 * 60; // Partenza alle 08:00
        let ultimaCoordGiorno: { lat: number; lng: number } | null = null;

        // Separiamo i pregressi (ordinati per oraInizio) dai cantieri normali
        const pregressiFila = interventiGiorno
            .filter((c: any) => c.isPregresso)
            .sort((a: any, b: any) => OrarioInMinuti(a.oraInizio, 480) - OrarioInMinuti(b.oraInizio, 480));

        let normaliFila = interventiGiorno.filter((c: any) => !c.isPregresso && c.selezionatoPerGiro);

        if (usaOrdinamentoGeografico && normaliFila.length > 1) {
            let daOrdinare = [...normaliFila];
            let ordinatiGeo = [];
            let currPos = { lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG };
            while (daOrdinare.length > 0) {
                daOrdinare.sort((a, b) => calcolaDistanzaKm(currPos.lat, currPos.lng, a.lat, a.lng) - calcolaDistanzaKm(currPos.lat, currPos.lng, b.lat, b.lng));
                const next = daOrdinare.shift();
                ordinatiGeo.push(next);
                if (next.lat && next.lng) currPos = { lat: next.lat, lng: next.lng };
            }
            normaliFila = ordinatiGeo;
        }

        let furgone = 0; let piccoNegativo = 0; let totaleMovimenti = 0;
        let prevCliente: any = null;
        let clientiAmmessi: any[] = [];

        let idxPreg = 0;
        let idxNorm = 0;

        while (idxPreg < pregressiFila.length || idxNorm < normaliFila.length) {
            const proxPreg = idxPreg < pregressiFila.length ? pregressiFila[idxPreg] : null;
            const proxNorm = idxNorm < normaliFila.length ? normaliFila[idxNorm] : null;

            let prendiPregresso = false;

            if (proxPreg && !proxNorm) {
                prendiPregresso = true;
            } else if (proxPreg && proxNorm) {
                const startPregMin = OrarioInMinuti(proxPreg.oraInizio, 480);
                const puntoPartenza = ultimaCoordGiorno || { lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG };
                const { minuti: minViaggioTest } = await calcolaTempoDistanzaGoogle(puntoPartenza.lat, puntoPartenza.lng, proxNorm.lat, proxNorm.lng);
                const { fineLavoro: fineNormTest } = calcolaTempistiche(orarioCorrenteMinuti, minViaggioTest, proxNorm.minutiLavoro);

                // Se il cantiere normale finisce prima dell'inizio del pregresso (con 10 min di margine), lo incastriamo PRIMA del pregresso!
                if (fineNormTest + 10 <= startPregMin) {
                    prendiPregresso = false;
                } else {
                    prendiPregresso = true;
                }
            }

            if (prendiPregresso && proxPreg) {
                idxPreg++;
                const c = proxPreg;
                const finePregMin = OrarioInMinuti(c.oraFine, 1080);

                if (prevCliente && ultimaCoordGiorno && c.lat && c.lng) {
                    const { minuti } = await calcolaTempoDistanzaGoogle(ultimaCoordGiorno.lat, ultimaCoordGiorno.lng, c.lat, c.lng);
                    prevCliente.minutiVersoProssimo = Number(minuti);
                    prevCliente.isRientroSede = false;
                    c.minutiDaPrecedente = Number(minuti);
                } else {
                    c.minutiDaPrecedente = 0;
                }

                c.isPartenzaDaSede = !prevCliente;
                orarioCorrenteMinuti = Math.max(orarioCorrenteMinuti, finePregMin + 15);
                if (orarioCorrenteMinuti >= INIZIO_PAUSA && orarioCorrenteMinuti < FINE_PAUSA) {
                    orarioCorrenteMinuti = FINE_PAUSA;
                }
                if (c.lat && c.lng) ultimaCoordGiorno = { lat: c.lat, lng: c.lng };
                prevCliente = c;
                clientiAmmessi.push(c);
            } else if (proxNorm) {
                idxNorm++;
                const c = proxNorm;
                const puntoPartenza = ultimaCoordGiorno || { lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG };
                const { minuti: minViaggio } = await calcolaTempoDistanzaGoogle(puntoPartenza.lat, puntoPartenza.lng, c.lat, c.lng);
                const { inizioLavoro, fineLavoro } = calcolaTempistiche(orarioCorrenteMinuti, minViaggio, c.minutiLavoro);

                if (scartaOvertimeAutomatico && Number(fineLavoro) > 1080) {
                    c.selezionatoPerGiro = false;
                    c.dataAssegnata = null;
                    c.motivoSospeso = "Scartato per Limite Orario (Overtime > 18:00)";
                    scartatiTemporanei.push(c);
                    continue;
                }

                c.minutiDaPrecedente = Number(minViaggio);
                c.isPartenzaDaSede = !prevCliente;

                if (prevCliente) {
                    prevCliente.minutiVersoProssimo = Number(minViaggio);
                    prevCliente.isRientroSede = false;
                }

                c.oraInizio = MinutiInOrario(inizioLavoro);
                c.oraFine = MinutiInOrario(fineLavoro);

                orarioCorrenteMinuti = fineLavoro;
                if (c.lat && c.lng) ultimaCoordGiorno = { lat: c.lat, lng: c.lng };
                prevCliente = c;

                let delta = c.resoEstintori ? c.muletti_eorv : -c.muletti_eorv;
                furgone += delta;
                totaleMovimenti += c.muletti_eorv;
                if (furgone < piccoNegativo) piccoNegativo = furgone;

                clientiAmmessi.push(c);
            }
        }
        
        if (prevCliente && ultimaCoordGiorno) {
            const { minuti } = await calcolaTempoDistanzaGoogle(ultimaCoordGiorno.lat, ultimaCoordGiorno.lng, SEDE_UFFICIO_LAT, SEDE_UFFICIO_LNG);
            prevCliente.minutiVersoProssimo = Number(minuti);
            prevCliente.isRientroSede = true;
        }

        const targetStat = nuoveStats.find(s => s.dataStr === dataStr);
        if (targetStat) {
            targetStat.orarioFineMinuti = orarioCorrenteMinuti;
            targetStat.statoGiornata = (FINE_GIORNATA_ASSOLUTA - orarioCorrenteMinuti) < 60 ? "PIENO" : "APERTO";
        }

        const fabbisognoNetto = Math.abs(piccoNegativo);
        const scortaJolly = Math.max(2, Math.ceil(fabbisognoNetto * 0.20));
        const totaleDaCaricare = fabbisognoNetto + scortaJolly;

        let descStr = `🛠 DISTINTA MULETTI:\n• Fabbisogno Netto: ${fabbisognoNetto} pz\n• Scorta Jolly (20%): ${scortaJolly} pz\n-----------------------------------\nTOTALE MULETTI DA CARICARE: ${totaleDaCaricare} pz\n-----------------------------------\n(Movimentazione totale: ${totaleMovimenti} pz)`;

        const primaRiga = clientiAmmessi[0] || { giorno: getNomeGiorno(dataStr) };
        const oldDistinta = vecchieDistinte.find(d => d.dataAssegnata === dataStr);

        const distintaObj = {
            codice: `DISTINTA_${dataStr}`, isDistinta: true, nome: `📦 DISTINTA DI CARICO (${formattaDataVisuale(dataStr)})`, numeroBi: "MAGAZZINO", indirizzo: "Sede Asso Antincendio", localita: "Genova", minutiLavoro: 60, giorno: primaRiga.giorno, dataAssegnata: dataStr, oraInizio: "07:00", oraFine: "08:00", descrizioneDistinta: descStr, selezionatoPerGiro: true, syncedToGoogle: oldDistinta ? oldDistinta.syncedToGoogle : false, lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG,
            fabbisognoNetto, scortaJolly, totaleDaCaricare, totaleMovimenti
        };

        distintaObj.descrizioneDistinta = generaDescrizioneEvento(distintaObj, tecnicoSelezionato.nome, tecnicoSelezionato.email, clientiAmmessi);

        nuovaLista.push(distintaObj, ...clientiAmmessi);
      }

      if (scartatiTemporanei.length > 0) {
          setClientiInSospeso(prev => {
              const prevCodes = new Set(prev.map(p => p.codice));
              const nuoviScarti = scartatiTemporanei.filter(s => !prevCodes.has(s.codice));
              return [...prev, ...nuoviScarti];
          });
      }

      return { listaOrdinata: nuovaLista, statsAggiornate: nuoveStats };
  };

  const processaRighe = async (righe: any[]) => {
    if (!googleToken) { alert("Connetti l'account Google per elaborare i dati."); return; }
    if (righe.length === 0) { alert("Nessun dato caricato. Seleziona prima un file."); return; }

    setInElaborazione(true);
    setProgressIA({ active: true, percent: 5, text: "Avvio elaborazione e lettura dati Excel..." });

    try {
        const clientiMappa = new globalThis.Map();
        
        setProgressIA({ active: true, percent: 15, text: "Sincronizzazione impegni Calendar in corso..." });
        const impegniEsistenti = await fetchEventiRealiCalendar(tecnicoSelezionato.email, googleToken);
        
        let tuttiIBiGiaInCalendario: string[] = [];
        impegniEsistenti.forEach(evt => { tuttiIBiGiaInCalendario = [...tuttiIBiGiaInCalendario, ...evt.numeriBiRilevati]; });

        setProgressIA({ active: true, percent: 30, text: "Analisi foglio Excel e anagrafiche Rubrica..." });
        
        let righeScartate = 0;

        for (const riga of righe) {
          const nomeCliente = String(riga.desclifor || "").toUpperCase();
          if (!nomeCliente) {
             righeScartate++;
             continue; 
          }

          const numeroBi = String(riga.numdoc || riga.nrob || riga.id_bi || riga.documento || riga.codicebi || "").trim();
          const codPerRubrica = riga.codclifor ? String(riga.codclifor).trim() : "";
          
          let idUnivocoCantiere = "";
          if (numeroBi && numeroBi !== "0" && numeroBi !== "-") {
              idUnivocoCantiere = `BI_${numeroBi}`;
          } else {
              const ind = String(riga.indirizzo || riga.INDIRIZZO || "NoInd");
              idUnivocoCantiere = `CLI_${codPerRubrica}_${ind}`;
          }

          const datiRubrica = rubricaClienti[codPerRubrica] || {};
          const telDaUsare = String(datiRubrica.cellulare || datiRubrica.telefono || riga.telefono || riga.TELEFONO || riga.tel || riga.TEL || "");
          const emailDaUsare = String(datiRubrica.email || riga.email || riga.EMAIL || riga['e-mail'] || "");
          const contattoDaUsare = String(datiRubrica.contatto || riga.referente || riga.REFERENTE || riga.contatto || "");

          if (!clientiMappa.has(idUnivocoCantiere)) {
            clientiMappa.set(idUnivocoCantiere, {
              codice: idUnivocoCantiere,
              codiceCliente: codPerRubrica,
              numeroBi: numeroBi || "BI-GENERICA", 
              nome: String(riga.desclifor || "Sconosciuto"),
              indirizzo: String(riga.indirizzo || riga.INDIRIZZO || ""), 
              localita: String(riga.localita || riga.LOCALITA || ""),
              cap: String(riga.cap || riga.CAP || ""),
              telefono: telDaUsare, email: emailDaUsare, contatto: contattoDaUsare,
              haInsoluto: false, resoEstintori: false, totaleArticoli: 0, minutiLavoro: 10, dettaglioAttrezzature: [], selezionatoPerGiro: false, giorniDallUltimoSos: 8, syncedToGoogle: false, muletti_eorv: 0,
              _rawItems: [], _maxMt: { eo: 0, po: 0, id: 0, la: 0, ir: 0, ss: 0, ri: 0, gp: 0, rv: 0 }
            });
          }

          const cliente = clientiMappa.get(idUnivocoCantiere);
          const testoRiga = JSON.stringify(riga).toUpperCase();
          const numFattura = riga.numft ? String(riga.numft).trim() : "";
          const valNomod = String(riga.nomod || riga.NOMOD || "").toUpperCase().trim();
          if (valNomod === "VERO" || valNomod === "TRUE" || (numFattura !== "" && numFattura !== "-" && (parseFloat(riga.imptotft) || 0) > 0 && !String(riga.codpagam || "").toUpperCase().trim().startsWith("RB"))) {
              cliente.haInsoluto = true;
          }
          if (String(riga.codclose || riga.SOS || riga.sos || "").toUpperCase().trim() === "SOS" || String(riga.SOS || riga.sos || "").toUpperCase().trim() === "VERO" || testoRiga.includes('"SOS"')) {
              cliente.resoEstintori = true;
          }

          const qta = Number(riga.quantita) || 0;
          const desc = String(riga.desdocrig || "");
          if (qta > 0 && desc) {
              cliente._rawItems.push({ desc, qta });
          }
          
          cliente._maxMt.eo = Math.max(cliente._maxMt.eo, Number(riga.eo_mt) || 0);
          cliente._maxMt.po = Math.max(cliente._maxMt.po, Number(riga.po_mt) || 0);
          cliente._maxMt.id = Math.max(cliente._maxMt.id, Number(riga.id_mt) || 0);
          cliente._maxMt.la = Math.max(cliente._maxMt.la, Number(riga.la_mt) || 0);
          cliente._maxMt.ir = Math.max(cliente._maxMt.ir, Number(riga.ir_mt) || 0);
          cliente._maxMt.ss = Math.max(cliente._maxMt.ss, Number(riga.ss_mt) || 0);
          cliente._maxMt.ri = Math.max(cliente._maxMt.ri, Number(riga.ri_mt) || 0);
          cliente._maxMt.gp = Math.max(cliente._maxMt.gp, Number(riga.gp_mt) || 0);
          cliente._maxMt.rv = Math.max(cliente._maxMt.rv, Number(riga.eo_rv || riga.EO_RV) || 0);
        }

        setProgressIA({ active: true, percent: 45, text: "Calcolo Geografico e Rapportini..." });
        for (const cliente of Array.from(clientiMappa.values())) {
            if (cliente.resoEstintori) {
                let veriMuletti = cliente._maxMt.rv || 0;
                cliente.muletti_eorv = veriMuletti;
                if (veriMuletti > 0) {
                    cliente.minutiLavoro = 10 + (veriMuletti * 2);
                    cliente.totaleArticoli = veriMuletti;
                    cliente.dettaglioAttrezzature = [{ descrizione: "Restituzione Estintori (SOS)", quantita: veriMuletti, categoria: "Rev (2m)", minutiImpegno: veriMuletti * 2 }];
                } else {
                    cliente.minutiLavoro = 10;
                    cliente.totaleArticoli = 0;
                    cliente.dettaglioAttrezzature = [{ descrizione: "Ritiro/Restituzione SOS", quantita: 1, categoria: "SOS", minutiImpegno: 10 }];
                }
            } else {
                let sumQtaEstintori = 0;
                let sumQtaPorte = 0;
                let sumQtaIdranti = 0;
                let sumQtaLampade = 0;
                let sumQtaIrai = 0;
                let sumQtaReti = 0;
                let sumQtaGruppi = 0;
                let sumQtaUscite = 0;

                cliente._rawItems.forEach((item: any) => {
                    const d = item.desc.toLowerCase();
                    const q = item.qta;
                    if (d.includes("estintor") || d.includes("schiuma") || d.includes("co2") || d.includes("polvere") || d.includes("idrico")) sumQtaEstintori += q;
                    else if (d.includes("porta") || d.includes("rei") || d.includes("tagliafuoco")) sumQtaPorte += q;
                    else if (d.includes("naspo") || d.includes("idrant") || d.includes("manichett") || d.includes("manichet")) sumQtaIdranti += q;
                    else if (d.includes("lampad")) sumQtaLampade += q;
                    else if (d.includes("emergenza") || d.includes("antipanico")) sumQtaUscite += q;
                    else if (d.includes("rivelazion") || d.includes("rilevazion") || d.includes("fumi") || d.includes("irai")) sumQtaIrai += q;
                    else if (d.includes("rete idrica") || d.includes("mandata")) sumQtaReti += q;
                    else if (d.includes("gruppo") || d.includes("pressurizzazion") || d.includes("pompe")) sumQtaGruppi += q;
                });

                const usaRighe = (
                    (sumQtaEstintori > 0 && cliente._maxMt.eo <= 1 && sumQtaEstintori > 1) || 
                    (sumQtaPorte > 0 && cliente._maxMt.po === 0) ||
                    (sumQtaIdranti > 0 && cliente._maxMt.id === 0) ||
                    (sumQtaLampade > 0 && cliente._maxMt.la === 0)
                );

                let minutiTotali = 0;
                let pezziTotali = 0;
                const dettaglio: any[] = [];

                if (usaRighe) {
                    const itemsRighe = [
                        { v: sumQtaEstintori, m: 2, n: "Estintori" },
                        { v: sumQtaIdranti, m: 2, n: "Idranti" },
                        { v: sumQtaPorte, m: 3, n: "Porte REI" },
                        { v: sumQtaUscite, m: 2, n: "Uscite Emergenza" },
                        { v: sumQtaLampade, m: 1, n: "Lampade Emergenza" },
                        { v: sumQtaIrai, m: 30, n: "Impianto IRAI" },
                        { v: sumQtaReti, m: 15, n: "Rete Idrica" },
                        { v: sumQtaGruppi, m: 20, n: "Gruppo Pompaggio" }
                    ];
                    itemsRighe.forEach(it => { 
                        if(it.v > 0) { 
                            minutiTotali += it.v * it.m; pezziTotali += it.v; 
                            dettaglio.push({ descrizione: it.n, quantita: it.v, categoria: `(${it.m}m/pz)`, minutiImpegno: it.v * it.m }); 
                        } 
                    });
                } else {
                    const itemsMt = [
                        { v: cliente._maxMt.eo, m: 2, n: "Estintori" },
                        { v: cliente._maxMt.id, m: 2, n: "Idranti" },
                        { v: cliente._maxMt.po, m: 3, n: "Porte REI" },
                        { v: cliente._maxMt.ss, m: 2, n: "Uscite Emergenza" },
                        { v: cliente._maxMt.la, m: 1, n: "Lampade Emergenza" },
                        { v: cliente._maxMt.ir, m: 30, n: "Impianto IRAI" },
                        { v: cliente._maxMt.ri, m: 15, n: "Rete Idrica" },
                        { v: cliente._maxMt.gp, m: 20, n: "Gruppo Pompaggio" }
                    ];
                    itemsMt.forEach(it => { 
                        if(it.v > 0) { 
                            minutiTotali += it.v * it.m; pezziTotali += it.v; 
                            dettaglio.push({ descrizione: it.n, quantita: it.v, categoria: `(${it.m}m/pz)`, minutiImpegno: it.v * it.m }); 
                        } 
                    });
                }
                
                cliente.muletti_eorv = cliente._maxMt.rv || 0;
                
                if (minutiTotali > 0) {
                    cliente.totaleArticoli = pezziTotali;
                    cliente.minutiLavoro = 10 + minutiTotali;
                    cliente.dettaglioAttrezzature = dettaglio;
                }
            }
        }

        const datiConCoordinate = [];
        for (const cliente of Array.from(clientiMappa.values())) {
          let coordinate = null;
          let origineCoord = "default";
          if (cliente.indirizzo?.trim().length > 2) {
            coordinate = await trovaCoordinateGoogle(cliente.indirizzo, cliente.localita);
            if (coordinate) origineCoord = coordinate.origine || "google";
          }
          datiConCoordinate.push({ ...cliente, lat: coordinate?.lat || 44.4056, lng: coordinate?.lng || 8.9463, origineCoord });
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
                contatto, telefono, email, muletti_eorv: 0, origineCoord: "calendar",
                descrizionePregressa: evt.description || "Impegno pregresso da Google Calendar"
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
          
          clientiDaPianificare.push(c);
        }

        const giornateStrutturate = dateRange.map(g => {
          const pregressiDelGiorno = eventiPregressiUI.filter(e => e.dataAssegnata === g.dataStr);
          const minutiOccupatiDaPregressi = pregressiDelGiorno.reduce((acc, imp) => acc + (Number(imp.minutiLavoro) || 60), 0);
          let minutiDisponibili = Math.max(0, 480 - minutiOccupatiDaPregressi);
          let statoGiornata = minutiDisponibili < 60 ? "PIENO" : "APERTO";

          return { ...g, minutiDisponibili, statoGiornata, anchorLat: null, anchorLng: null, orarioPartenzaMinuti: 8 * 60, orarioFineMinuti: 8 * 60 };
        });

        let unassigned = [...clientiDaPianificare];
        const clientiPianificati: any[] = [];

        setProgressIA({ active: true, percent: 55, text: "Connessione ai 3 Agenti IA..." });

        let simInterval = setInterval(() => {
            setProgressIA(prev => {
                if (prev.percent < 90) {
                    return { ...prev, percent: prev.percent + 1, text: "Ragionamento IA in corso (Cartografo -> Cronometrista -> Supervisore)..." };
                }
                return prev;
            });
        }, 1000);

        try {
            const payload = {
                interventi: unassigned.map(c => ({
                    codice: c.codice,
                    numeroBi: c.numeroBi,
                    nome: c.nome,
                    indirizzo: c.indirizzo,
                    localita: c.localita,
                    cap: c.cap,
                    minutiLavoro: c.minutiLavoro,
                    isSOS: c.resoEstintori
                })),
                giornate: giornateStrutturate.map(g => ({
                    dataStr: g.dataStr,
                    nomeGiorno: g.nomeGiorno,
                    minutiDisponibili: g.minutiDisponibili,
                    pregressi: eventiPregressiUI.filter(p => p.dataAssegnata === g.dataStr).map(p => ({
                        codice: p.codice,
                        nome: p.nome,
                        indirizzo: p.indirizzo,
                        oraInizio: p.oraInizio,
                        oraFine: p.oraFine,
                        minutiLavoro: p.minutiLavoro
                    }))
                }))
            };

            const res = await fetch('/api/planner', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    action: 'pianifica_settimana', 
                    data: payload,
                    history: messages.map(m => ({
                        role: m.sender === 'user' ? 'user' : 'model',
                        text: m.text
                    }))
                })
            });
            
            clearInterval(simInterval);
            setProgressIA({ active: true, percent: 92, text: "Applicazione rotte e calcolo tempistiche..." });

            const iaResult = await res.json();
            
            if (iaResult.success && iaResult.piano && iaResult.piano.giorni) {
                unassigned.forEach(c => {
                    let assigned = false;
                    for (const [dataStr, codici] of Object.entries(iaResult.piano.giorni)) {
                        const arrCodici = codici as string[];
                        const idx = arrCodici.indexOf(c.codice);
                        if (idx !== -1) {
                            c.dataAssegnata = dataStr;
                            c.giorno = getNomeGiorno(dataStr);
                            c.selezionatoPerGiro = true;
                            c._aiOrder = idx;
                            clientiPianificati.push(c);
                            assigned = true;
                            break;
                        }
                    }
                    if (!assigned) {
                        clientiSospesiTmp.push({ ...c, motivoSospeso: "In Sospeso (Logistica / Tempo)", selezionatoPerGiro: false, dataAssegnata: null });
                    }
                });
            } else {
                throw new Error(iaResult.error || "L'Agente IA non ha restituito un piano valido");
            }
        } catch (e: any) {
            clearInterval(simInterval);
            console.error("Dettaglio errore IA:", e);
            if (typeof window !== 'undefined') {
                window.alert(`Errore IA: ${e.message || "Errore sconosciuto"}. I cantieri sono stati inseriti nei Sospesi.`);
            }
            unassigned.forEach(c => {
                clientiSospesiTmp.push({ ...c, motivoSospeso: `Errore: ${e.message || "Sconosciuto"}`, selezionatoPerGiro: false, dataAssegnata: null });
            });
        }

        setProgressIA({ active: true, percent: 98, text: "Finalizzazione plancia logistica..." });

        // Rispettiamo l'ordine restituito dall'IA
        clientiPianificati.sort((a, b) => (a._aiOrder ?? 999) - (b._aiOrder ?? 999));
        const datiCombinati = [...eventiPregressiUI, ...clientiPianificati];

        const resRic = await ricalcolaDistinteEOrari(datiCombinati, giornateStrutturate, true, true);
        setInterventiGrezzi(resRic.listaOrdinata);
        setGiornateStats(resRic.statsAggiornate);
        setClientiInSospeso(prev => {
            const codiciSosp = new Set(clientiSospesiTmp.map(s => s.codice));
            const extra = prev.filter(p => !codiciSosp.has(p.codice) && !clientiPianificati.some(cp => cp.codice === p.codice));
            return [...clientiSospesiTmp, ...extra];
        });
        setClientiGiaCalendarizzati(clientiScartatiTmp);
        
        if (giornateStrutturate.length > 0) {
            setGiornoSceltoFiltro(giornateStrutturate[0].dataStr);
        } else {
            setGiornoSceltoFiltro("sospesi");
        }

        setProgressIA({ active: false, percent: 100, text: "Completato!" });
        setInElaborazione(false);
        
    } catch (error: any) {
        console.error("Errore critico durante processaRighe:", error);
        alert(`Generazione interrotta: ${error.message}`);
        setProgressIA({ active: false, percent: 0, text: "" });
        setInElaborazione(false);
    }
  };

  const handleDragStartKanban = (e: React.DragEvent, codice: string) => {
      e.dataTransfer.setData("application/kanban-card", codice);
  };

  const handleDragOverKanban = (e: React.DragEvent) => { 
      e.preventDefault(); 
  };

  const handleDropToDay = async (e: React.DragEvent, targetDataAssegnata: string | null, targetCodice?: string) => {
      e.preventDefault();
      e.stopPropagation();
      setDragOverCard(null);

      const draggedCodice = e.dataTransfer.getData("application/kanban-card");
      if (!draggedCodice || draggedCodice === targetCodice) return;

      await spostaCantiereRapido(draggedCodice, targetDataAssegnata, targetCodice);
  };

  const spostaCantiereRapido = async (codiceCantiere: string, dataDestinazione: string | null, targetCodice?: string) => {
      setMenuAperto(null);
      setInElaborazione(true);
      
      try {
          let result = Array.from(interventiGrezzi);
          let item = result.find(i => i.codice === codiceCantiere);
          let isFromSospesi = false;

          if (!item) {
              const sospesoItem = clientiInSospeso.find(i => i.codice === codiceCantiere);
              if (sospesoItem) {
                  item = { ...sospesoItem, selezionatoPerGiro: true, dataAssegnata: dataDestinazione };
                  isFromSospesi = true;
              }
          }

          if (!item) return;

          if (isFromSospesi) {
              setClientiInSospeso(prev => prev.filter(c => c.codice !== codiceCantiere));
          } else {
              result = result.filter(i => i.codice !== codiceCantiere);
          }

          if (dataDestinazione) {
              item.dataAssegnata = dataDestinazione;
              item.giorno = getNomeGiorno(dataDestinazione);
              item.selezionatoPerGiro = true;
              
              if (targetCodice) {
                  const targetIndex = result.findIndex(i => i.codice === targetCodice && i.dataAssegnata === dataDestinazione);
                  if (targetIndex !== -1) {
                      result.splice(targetIndex, 0, item);
                  } else {
                      result.push(item);
                  }
              } else {
                  result.push(item);
              }
              
              // scartaOvertimeAutomatico = false per non annullare mai lo spostamento manuale dell'utente
              const resRic = await ricalcolaDistinteEOrari(result, giornateStats, false, false);
              setGiornateStats(resRic.statsAggiornate);
              setInterventiGrezzi(resRic.listaOrdinata);
          } else {
              item.dataAssegnata = null;
              item.selezionatoPerGiro = false;
              item.motivoSospeso = "Spostato manualmente in Sospesi";
              setClientiInSospeso(prev => [...prev, item]);
              const resRic = await ricalcolaDistinteEOrari(result, giornateStats, false, false);
              setGiornateStats(resRic.statsAggiornate);
              setInterventiGrezzi(resRic.listaOrdinata);
          }
      } finally {
          setInElaborazione(false);
      }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!googleToken) { alert("Devi connettere Google Calendar prima di caricare il file."); e.target.value = ''; return; }
    const file = e.target.files?.[0];
    if (!file) return;

    const nomeFileLower = file.name.toLowerCase();
    setNomeFileCorrente(file.name);
    
    let matchTrovato = false;
    for (const tech of tecniciAnagrafica) {
      const paroleNome = tech.nome.toLowerCase().split(" ");
      for (const parola of paroleNome) {
        if (parola.length > 3 && nomeFileLower.includes(parola)) {
          setTecnicoSelezionato(tech);
          matchTrovato = true;
          break;
        }
      }
      if (matchTrovato) break;
    }

    elaboraFileExcelGenerico(file, (righe) => {
       if (righe.length > 0) setDatiGrezziCaricati(righe);
    });
    e.target.value = '';
  };

  // CHAT IA ARRICCHITA CON CONTESTO COMPLETO (GIORNI, LOCALITÀ, BI, ORARI)
  const handleSendMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!chatInput.trim() || chatLoading) return;

    const userMessage = chatInput;
    const updatedMessages = [...messages, { sender: 'user', text: userMessage }];
    setMessages(updatedMessages);
    setChatInput('');
    setChatLoading(true);
    setProgressIA({ active: true, percent: 15, text: "Esecuzione comando IA in corso..." });

    let simInterval = setInterval(() => {
        setProgressIA(prev => {
            if (prev.percent < 90) return { ...prev, percent: prev.percent + 2, text: "Aggiornamento plancia secondo le tue istruzioni..." };
            return prev;
        });
    }, 600);

    try {
      const res = await fetch('/api/planner', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'modifica_chat',
          comandoUtente: userMessage,
          statoAttuale: {
            giornateDisponibili: giornateStats.map(g => ({
              dataStr: g.dataStr,
              nomeGiorno: g.nomeGiorno
            })),
            interventiInRotta: interventiGrezzi.filter(i => !i.isDistinta).map(i => ({
              codice: i.codice,
              numeroBi: i.numeroBi,
              nome: i.nome,
              indirizzo: i.indirizzo,
              localita: i.localita,
              cap: i.cap,
              dataAssegnata: i.dataAssegnata,
              nomeGiorno: getNomeGiorno(i.dataAssegnata),
              oraInizio: i.oraInizio,
              oraFine: i.oraFine,
              minutiLavoro: i.minutiLavoro,
              isPregresso: !!i.isPregresso,
              isSOS: !!i.resoEstintori,
              isInsoluto: !!i.haInsoluto
            })),
            sospesi: clientiInSospeso.map(s => ({
              codice: s.codice,
              numeroBi: s.numeroBi,
              nome: s.nome,
              indirizzo: s.indirizzo,
              localita: s.localita,
              cap: s.cap,
              minutiLavoro: s.minutiLavoro,
              isSOS: !!s.resoEstintori,
              isInsoluto: !!s.haInsoluto
            }))
          },
          history: updatedMessages.map(m => ({
            role: m.sender === 'user' ? 'user' : 'model',
            text: m.text
          }))
        })
      });

      clearInterval(simInterval);
      setProgressIA({ active: true, percent: 95, text: "Applicazione modifiche e ricalcolo orari..." });

      const data = await res.json();
      if (data.success) {
        const rispostaDiscorsiva = data.messaggioChat || "Ho aggiornato la pianificazione come richiesto.";
        setMessages(prev => [...prev, { sender: 'ai', text: rispostaDiscorsiva }]);
        
        if (data.piano && data.piano.giorni) {
          try {
            let tuttaLaLista = [...interventiGrezzi.filter(i => !i.isDistinta), ...clientiInSospeso];
            let nuoviAssegnati: any[] = [];
            let nuoviSospesi: any[] = [];

            tuttaLaLista.forEach(item => {
              // I pregressi restano sempre ancorati al loro giorno originale
              if (item.isPregresso) {
                nuoviAssegnati.push(item);
                return;
              }

              item._aiIndex = 9999; 
              let assegnato = false;
              
              for (const [dataStr, codici] of Object.entries(data.piano.giorni)) {
                const codiciArray = codici as string[];
                const indiceIA = codiciArray.indexOf(item.codice);
                
                if (indiceIA !== -1) {
                  item.dataAssegnata = dataStr;
                  item.giorno = getNomeGiorno(dataStr);
                  item.selezionatoPerGiro = true;
                  item._aiIndex = indiceIA; 
                  nuoviAssegnati.push(item);
                  assegnato = true;
                  break;
                }
              }
              
              if (!assegnato) {
                if (data.piano.sospesi && (data.piano.sospesi as string[]).includes(item.codice)) {
                  item.selezionatoPerGiro = false;
                  item.dataAssegnata = null;
                  item.motivoSospeso = "Spostato in Sospesi via Chat IA";
                  nuoviSospesi.push(item);
                } else if (item.dataAssegnata) {
                  nuoviAssegnati.push(item);
                } else {
                  item.selezionatoPerGiro = false;
                  nuoviSospesi.push(item);
                }
              }
            });

            nuoviAssegnati.sort((a, b) => {
               if (a.dataAssegnata !== b.dataAssegnata) return a.dataAssegnata.localeCompare(b.dataAssegnata);
               return (a._aiIndex || 0) - (b._aiIndex || 0);
            });

            // Non scartiamo l'overtime in automatico così l'ordine dato in chat viene SEMPRE eseguito
            const resRic = await ricalcolaDistinteEOrari(nuoviAssegnati, giornateStats, false, false);
            setInterventiGrezzi(resRic.listaOrdinata);
            setGiornateStats(resRic.statsAggiornate);
            setClientiInSospeso(nuoviSospesi);
          } catch (recalcErr) {
            console.error("Errore nel ricalcolo locale:", recalcErr);
          }
        }
      } else {
        setMessages(prev => [...prev, { sender: 'ai', text: `Errore: ${data.error || 'Non sono riuscito ad applicare la modifica.'}` }]);
      }
    } catch (err) {
      clearInterval(simInterval);
      setMessages(prev => [...prev, { sender: 'ai', text: 'Errore di connessione con il server.' }]);
    } finally {
      setProgressIA({ active: false, percent: 100, text: "Fatto!" });
      setChatLoading(false);
    }
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
            contatto: findKey(["contatto", "referente", "persona"]),
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
       const insolutiMappa = new globalThis.Map();
       
       righe.forEach(r => {
          const numeroBi = String(r.numdoc || r.nrob || r.id_bi || r.documento || r.codicebi || "").trim();
          const numFattura = r.numft ? String(r.numft).trim() : "";
          const importo = parseFloat(r.imptotft) || 0;
          const codpagam = String(r.codpagam || "").toUpperCase().trim();
          const despagam = String(r.despagam || r.des_pagam || r.pagamento || "").trim();
          const modalitaPagamento = despagam ? `${codpagam} - ${despagam}` : codpagam;
          
          if (numFattura && numFattura !== "-" && importo > 0 && !String(codpagam).startsWith("RB")) {
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
          location: `${clienteSospeso.indirizzo}, ${clienteSospeso.cap ? clienteSospeso.cap + ' ' : ''}${clienteSospeso.localita || 'Genova'}, Italia`,
          description: generaDescrizioneEvento(clienteSospeso, tecnicoSelezionato.nome, tecnicoSelezionato.email),
          start: { dateTime: startDateTime.toISOString(), timeZone: 'Europe/Rome' },
          end: { dateTime: endDateTime.toISOString(), timeZone: 'Europe/Rome' },
        };

        const res = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(tecnicoSelezionato.email)}/events`, {
          method: 'POST', headers: { 'Authorization': `Bearer ${googleToken}`, 'Content-Type': 'application/json' }, body: JSON.stringify(event)
        });

        if (res.ok) {
          const giornoScelto = getNomeGiorno(conf.data);
          const newClient = { ...clienteSospeso, settimana: `${formattaDataVisuale(dataInizio)} al ${formattaDataVisuale(dataFine)}`, giorno: giornoScelto, dataAssegnata: conf.data, oraInizio: conf.oraInizio, oraFine: `${hhEnd}:${mmEnd}`, selezionatoPerGiro: true, syncedToGoogle: true };
          
          setInElaborazione(true);
          const resRic = await ricalcolaDistinteEOrari([...interventiGrezzi, newClient], giornateStats, false, false);
          setInterventiGrezzi(resRic.listaOrdinata);
          setGiornateStats(resRic.statsAggiornate);
          setInElaborazione(false);
          
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
      
      setInElaborazione(true);
      const resRic = await ricalcolaDistinteEOrari([...interventiGrezzi, newClient], giornateStats, false, false);
      setInterventiGrezzi(resRic.listaOrdinata);
      setGiornateStats(resRic.statsAggiornate);
      setInElaborazione(false);

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
            location: `${cliente.indirizzo}, ${cliente.cap ? cliente.cap + ' ' : ''}${cliente.localita || 'Genova'}, Italia`,
            description: generaDescrizioneEvento(cliente, tecnicoSelezionato.nome, tecnicoSelezionato.email),
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

  const createNumberedIcon = (numero: number) => {
    if (leafletRef.current) {
      return leafletRef.current.divIcon({
        className: "custom-marker-icon",
        html: `<div style="background-color: #0f172a; color: white; border-radius: 50%; width: 26px; height: 26px; display: flex; align-items: center; justify-content: center; font-weight: bold; border: 2px solid white; box-shadow: 0 2px 5px rgba(0,0,0,0.4); font-size: 11px;">${numero}</div>`,
        iconSize: [26, 26],
        iconAnchor: [13, 26],
        popupAnchor: [0, -26]
      });
    }
    return undefined;
  };

  const clientiAttiviCount = interventiGrezzi.filter(i => i.selezionatoPerGiro && !i.isDistinta && !i.isPregresso).length;
  const globaleAssegnati = interventiGrezzi.filter(i => i.selezionatoPerGiro && !i.isDistinta && !i.isPregresso);
  const globalPregressiCount = interventiGrezzi.filter(i => i.isPregresso).length;
  const totaleInCarico = globaleAssegnati.length + clientiInSospeso.length + globalPregressiCount;
  
  const giorniVisibiliPerStampa = [...new Set(interventiGrezzi.map(i => i.dataAssegnata).filter(Boolean))].sort();

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
              </tbody>
           </table>
        </div>
     );
  }

  let listaAttiva: any[] = [];
  let statGiornoAttivo: any = null;

  if (giornoSceltoFiltro === "sospesi") {
      listaAttiva = clientiInSospeso;
  } else {
      listaAttiva = interventiGrezzi.filter(i => i.selezionatoPerGiro && i.dataAssegnata === giornoSceltoFiltro);
      statGiornoAttivo = giornateStats.find(g => g.dataStr === giornoSceltoFiltro);
  }

  return (
    <div className="flex flex-col h-screen w-full bg-slate-100 text-slate-900 font-sans overflow-hidden relative print:h-auto print:overflow-visible print:bg-white" onClick={() => setMenuAperto(null)}>
      
      {/* OVERLAY BARRA DI PROGRESSIONE IA */}
      {progressIA.active && (
        <div className="fixed top-6 left-1/2 -translate-x-1/2 z-[9999] bg-white rounded-2xl shadow-2xl border border-blue-200 p-4 w-[90%] max-w-md flex flex-col gap-3">
           <div className="flex justify-between items-center">
              <span className="text-xs font-bold text-slate-700 flex items-center gap-2"><Loader2 size={14} className="animate-spin text-blue-600"/> {progressIA.text}</span>
              <span className="text-xs font-extrabold text-blue-700">{progressIA.percent}%</span>
           </div>
           <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden">
              <div className="bg-blue-600 h-full rounded-full transition-all duration-500 ease-out" style={{ width: `${progressIA.percent}%` }}></div>
           </div>
        </div>
      )}

      {/* HEADER PRINCIPALE */}
      <header className="w-full bg-white border-b border-slate-200 px-6 py-3 flex flex-wrap justify-between items-center gap-4 shrink-0 z-30 print:hidden">
         <div className="flex items-center gap-3">
            <span className="p-2 bg-blue-600 text-white rounded-xl shadow-sm"><Zap size={20} /></span>
            <div>
              <h1 className="text-lg font-extrabold tracking-tight text-slate-900 flex items-center gap-2">
                Asso Antincendio <span className="text-blue-600 font-normal text-sm">| Smart Logistics</span>
              </h1>
              <p className="text-[11px] text-slate-500">Rubrica Clienti, Dati Rapportino & Motore Multi-Agente IA</p>
            </div>
         </div>

         <div className="flex flex-wrap items-center gap-2">
            <Link href="/flussi" className="px-3 py-2 bg-blue-50 text-blue-700 border border-blue-200 rounded-xl text-xs font-bold hover:bg-blue-100 transition flex items-center gap-1.5">
              📊 Gestione Flussi
            </Link>

            <a href="https://calendar.google.com" target="_blank" rel="noopener noreferrer" className="px-3 py-2 bg-slate-50 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold hover:bg-slate-100 transition flex items-center gap-1.5">
              <Phone size={14} /> Chiamate <ExternalLink size={12} className="opacity-60" />
            </a>

            <label className={`px-3 py-2 rounded-xl font-bold text-xs flex items-center gap-1.5 transition-all cursor-pointer border ${Object.keys(rubricaClienti).length > 0 ? 'bg-amber-50 text-amber-800 border-amber-200 hover:bg-amber-100' : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'}`}>
              <Book size={14} />
              {Object.keys(rubricaClienti).length > 0 ? `Rubrica (${Object.keys(rubricaClienti).length})` : "Carica Rubrica"}
              <input type="file" accept=".xlsx, .xls, .csv, .txt" className="hidden" onChange={handleRubricaUpload} />
            </label>

            <label className="px-3 py-2 rounded-xl font-bold text-xs flex items-center gap-1.5 transition-all cursor-pointer bg-red-50 text-red-700 hover:bg-red-100 border border-red-200">
              <Banknote size={14} /> Lista Insoluti
              <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={handleInsolutiUpload} />
            </label>

            {isCheckingAuth ? (
              <span className="text-slate-500 font-bold flex items-center gap-1.5 text-xs px-3 py-2"><Loader2 size={14} className="animate-spin text-blue-600" /> Verifica...</span>
            ) : !googleToken ? (
              <button onClick={eseguiLoginGoogle} className="bg-slate-900 hover:bg-slate-800 text-white px-4 py-2 rounded-xl font-bold flex items-center gap-1.5 shadow-sm text-xs"><Calendar size={14} /> Connetti Google</button>
            ) : (
              <div className="flex items-center gap-1">
                <span className="bg-emerald-50 text-emerald-800 px-3 py-2 rounded-xl font-bold flex items-center gap-1.5 border border-emerald-200 text-xs"><Check size={14} /> Calendar OK</span>
                <button onClick={logoutGoogle} title="Disconnetti Google" className="bg-white hover:bg-red-50 text-slate-500 hover:text-red-600 p-2 rounded-xl border border-slate-200"><X size={14} /></button>
              </div>
            )}
         </div>
      </header>

      {/* CRUSCOTTO DI CONTROLLO OPERATIVO */}
      <div className="bg-white border-b border-slate-200 px-6 py-3 shrink-0 z-20 print:hidden">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 items-center">
          
          <div className="lg:col-span-4 flex flex-wrap items-center gap-2 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
            <div className="flex items-center gap-1.5">
              <span className="text-[10px] font-extrabold uppercase text-slate-500">Dal</span>
              <input type="date" value={dataInizio} onChange={(e) => setDataInizio(e.target.value)} className="bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs font-bold focus:outline-none" />
            </div>
            <div className="flex items-center gap-1.5">
              <span className="text-[10px] font-extrabold uppercase text-slate-500">Al</span>
              <input type="date" value={dataFine} onChange={(e) => setDataFine(e.target.value)} className="bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs font-bold focus:outline-none" />
            </div>
            <div className="flex gap-1 ml-auto">
              {Object.keys(giorniAttivi).map((g) => (
                <button key={g} onClick={() => toggleGiornoAttivo(g)} className={`px-2 py-1 rounded-md text-[10px] font-extrabold transition ${giorniAttivi[g] ? 'bg-blue-600 text-white' : 'bg-slate-200 text-slate-400 line-through'}`}>
                  {g.slice(0, 3)}
                </button>
              ))}
            </div>
          </div>

          <div className="lg:col-span-4 flex flex-wrap items-center gap-2 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
            <select value={tecnicoSelezionato.nome} onChange={(e) => { const t = tecniciAnagrafica.find(t => t.nome === e.target.value); if(t) setTecnicoSelezionato(t); }} className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-bold text-slate-800 focus:outline-none cursor-pointer flex-1 min-w-[140px]">
              {tecniciAnagrafica.map((t, idx) => <option key={idx} value={t.nome}>{t.nome}</option>)}
            </select>

            <label className={`px-3 py-1.5 rounded-lg font-bold text-xs shadow-sm transition-all flex items-center gap-1.5 border ${!googleToken ? 'opacity-50 cursor-not-allowed bg-slate-200 text-slate-500' : datiGrezziCaricati.length > 0 ? 'bg-blue-50 text-blue-800 border-blue-200 cursor-pointer hover:bg-blue-100' : 'bg-white text-slate-700 border-slate-300 cursor-pointer hover:bg-slate-100'}`} title={nomeFileCorrente}>
              <FileSpreadsheet size={14} className="text-blue-600" />
              <span className="truncate max-w-[110px]">{datiGrezziCaricati.length > 0 ? nomeFileCorrente : "Sfoglia Excel"}</span>
              <input type="file" accept=".xlsx, .xls, .csv, .txt" className="hidden" onChange={handleFileUpload} disabled={!googleToken} />
            </label>

            <button
              onClick={() => processaRighe(datiGrezziCaricati)}
              disabled={datiGrezziCaricati.length === 0 || inElaborazione}
              className={`px-4 py-1.5 rounded-lg text-xs font-extrabold transition-all flex items-center gap-1.5 ${(datiGrezziCaricati.length > 0 && !inElaborazione) ? 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-sm cursor-pointer' : 'bg-slate-200 text-slate-400 cursor-not-allowed'}`}
            >
              {inElaborazione ? <><Loader2 size={14} className="animate-spin"/> Elaboro...</> : <><Sliders size={14} /> Genera Plancia</>}
            </button>
          </div>

          <div className="lg:col-span-4 grid grid-cols-4 gap-2">
            <div className="px-3 py-1.5 bg-blue-50 rounded-xl border border-blue-200 flex flex-col justify-center">
              <span className="text-[9px] text-blue-700 font-bold uppercase truncate">In Carico</span>
              <span className="text-lg font-extrabold text-blue-900 leading-tight">{totaleInCarico}</span>
            </div>
            <div className="px-3 py-1.5 bg-emerald-50 rounded-xl border border-emerald-200 flex flex-col justify-center">
              <span className="text-[9px] text-emerald-700 font-bold uppercase truncate">In Rotta</span>
              <span className="text-lg font-extrabold text-emerald-900 leading-tight">{clientiAttiviCount}</span>
            </div>
            <div className="px-3 py-1.5 bg-amber-50 rounded-xl border border-amber-200 flex flex-col justify-center">
              <span className="text-[9px] text-amber-700 font-bold uppercase truncate">Sospesi</span>
              <span className="text-lg font-extrabold text-amber-900 leading-tight">{clientiInSospeso.length}</span>
            </div>
            <div className="px-3 py-1.5 bg-slate-100 rounded-xl border border-slate-200 flex flex-col justify-center">
              <span className="text-[9px] text-slate-500 font-bold uppercase truncate">Già Fatti</span>
              <span className="text-lg font-extrabold text-slate-700 leading-tight">{clientiGiaCalendarizzati.length}</span>
            </div>
          </div>

        </div>
      </div>

      {/* AREA PRINCIPALE: TAB GIORNI + SPLIT SCREEN */}
      <div className="flex-1 flex flex-col overflow-hidden print:hidden">
          
          {interventiGrezzi.length > 0 && (
              <div className="bg-white border-b border-slate-200 px-6 py-2 flex items-center justify-between gap-2 overflow-x-auto shrink-0 custom-scrollbar z-20">
                  <div className="flex items-center gap-2">
                      <button 
                          onClick={() => setGiornoSceltoFiltro("sospesi")}
                          className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl font-bold text-xs transition-colors whitespace-nowrap cursor-pointer ${giornoSceltoFiltro === "sospesi" ? 'bg-amber-500 text-white shadow-sm' : 'bg-amber-50 text-amber-900 border border-amber-200 hover:bg-amber-100'}`}
                      >
                          <PauseCircle size={14} /> SOSPESI <span className={`px-1.5 py-0.5 rounded-full text-[10px] font-extrabold ${giornoSceltoFiltro === "sospesi" ? 'bg-amber-700 text-white' : 'bg-amber-200 text-amber-900'}`}>{clientiInSospeso.length}</span>
                      </button>
                      
                      <div className="w-px h-6 bg-slate-200 mx-1"></div>
                      
                      {giornateStats.map(g => {
                          const isOvertime = g.orarioFineMinuti > 1080;
                          const countGiorno = interventiGrezzi.filter(i => i.selezionatoPerGiro && !i.isDistinta && i.dataAssegnata === g.dataStr).length;
                          const hh = String(Math.floor(g.orarioFineMinuti / 60)).padStart(2, '0');
                          const mm = String(Math.floor(g.orarioFineMinuti % 60)).padStart(2, '0');
                          return (
                              <button 
                                  key={g.dataStr}
                                  onClick={() => setGiornoSceltoFiltro(g.dataStr)}
                                  className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl font-bold text-xs transition-all whitespace-nowrap border cursor-pointer ${giornoSceltoFiltro === g.dataStr ? 'bg-slate-900 border-slate-900 text-white shadow-sm' : 'bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100'}`}
                              >
                                  <span>{g.nomeGiorno} {formattaDataVisuale(g.dataStr).slice(0,5)}</span>
                                  <span className={`px-1.5 py-0.5 rounded text-[10px] font-mono ${giornoSceltoFiltro === g.dataStr ? 'bg-slate-700 text-slate-200' : 'bg-white text-slate-600 border border-slate-200'}`}>
                                      {countGiorno} ({hh}:{mm})
                                  </span>
                                  {isOvertime && <AlertOctagon size={13} className="text-red-400" />}
                              </button>
                          );
                      })}
                  </div>

                  <div className="flex items-center gap-2 shrink-0">
                      <button onClick={() => window.print()} className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-3.5 py-1.5 rounded-xl font-bold flex items-center gap-1.5 text-xs border border-slate-200 cursor-pointer">
                          <Printer size={14} /> Stampa Fogli Viaggio
                      </button>
                      {bulkSyncStatus.active ? (
                          <div className="bg-blue-50 border border-blue-200 px-4 py-1.5 rounded-xl flex items-center gap-2 text-xs font-bold text-blue-800">
                              <Loader2 size={14} className="animate-spin"/> Sync... {bulkSyncStatus.current}/{bulkSyncStatus.total}
                          </div>
                      ) : (
                          <button onClick={inviaPianificazioneAGoogle} className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-1.5 rounded-xl font-bold flex items-center gap-1.5 shadow-sm text-xs transition-all cursor-pointer">
                              <UploadCloud size={14} /> Salva su Calendar ({globaleAssegnati.filter(i=>!i.syncedToGoogle).length})
                          </button>
                      )}
                  </div>
              </div>
          )}

          {/* SPLIT SCREEN: SCHEDE OPERATIVE COMPLETE A SINISTRA + MAPPA A DESTRA */}
          {interventiGrezzi.length > 0 ? (
              <div className="flex-1 flex overflow-hidden">
                  
                  {/* PANNELLO SINISTRO: SCHEDE COMPLETE CON REFERENTI, TELEFONI E RAPPORTINO */}
                  <div className="w-full md:w-[54%] lg:w-[50%] h-full bg-white border-r border-slate-200 flex flex-col z-10">
                      
                      <div className="px-4 py-3 border-b border-slate-200 bg-slate-50 flex justify-between items-center shrink-0">
                          <div>
                              <h2 className="font-extrabold text-slate-900 text-sm uppercase tracking-tight">
                                  {giornoSceltoFiltro === "sospesi" ? `⏸ Cantieri in Sospeso (${clientiInSospeso.length})` : `📅 ${statGiornoAttivo?.nomeGiorno} ${formattaDataVisuale(giornoSceltoFiltro)}`}
                              </h2>
                              <p className="text-[11px] text-slate-500 font-medium">
                                  Trascina le schede per riordinare la giornata oppure usa ⇄ per cambiare giorno.
                              </p>
                          </div>
                          {statGiornoAttivo && (
                              <div className="text-right shrink-0">
                                  <span className={`text-xs font-extrabold px-2.5 py-1 rounded-lg flex items-center gap-1 ${statGiornoAttivo.orarioFineMinuti > 1080 ? 'bg-red-100 text-red-800 border border-red-200' : 'bg-emerald-100 text-emerald-800 border border-emerald-200'}`}>
                                      {statGiornoAttivo.orarioFineMinuti > 1080 ? <AlertOctagon size={13}/> : <CheckCircle size={13}/>}
                                      Fine: {MinutiInOrario(statGiornoAttivo.orarioFineMinuti)}
                                  </span>
                              </div>
                          )}
                      </div>

                      <div className="flex-1 overflow-y-auto p-3 space-y-3 custom-scrollbar bg-slate-50/60" onDragOver={handleDragOverKanban} onDrop={(e) => handleDropToDay(e, giornoSceltoFiltro === "sospesi" ? null : giornoSceltoFiltro)}>
                          {listaAttiva.length === 0 ? (
                              <div className="flex flex-col items-center justify-center h-48 text-slate-400 border-2 border-dashed border-slate-200 rounded-2xl m-2 bg-white">
                                  <List size={28} className="mb-2 opacity-50" />
                                  <p className="text-xs font-bold">Nessun cantiere in questa vista</p>
                              </div>
                          ) : (
                              listaAttiva.map((item, idx) => {
                                  const numTappa = idx + (listaAttiva[0]?.isDistinta ? 0 : 1);

                                  // CARD SPECIALIZZATA PER DISTINTA DI CARICO MAGAZZINO
                                  if (item.isDistinta) {
                                      return (
                                          <div key={item.codice} className="p-3.5 rounded-xl border border-orange-200 bg-orange-50/70 shadow-xs">
                                              <div className="flex items-center justify-between gap-2 mb-2">
                                                  <div className="flex items-center gap-2">
                                                      <span className="px-2 py-0.5 rounded bg-orange-200 text-orange-900 font-extrabold text-[10px]">07:00 - 08:00 • MAGAZZINO</span>
                                                      <span className="font-extrabold text-xs text-slate-900">{item.nome}</span>
                                                  </div>
                                                  <button onClick={() => setClienteSelezionatoScheda(item)} className="text-orange-800 hover:bg-orange-100 px-2.5 py-1 rounded-lg text-[11px] font-bold border border-orange-200 bg-white cursor-pointer">
                                                      Dettaglio & Link Maps
                                                  </button>
                                              </div>
                                              <div className="text-[11px] font-mono text-slate-700 bg-white/90 p-2.5 rounded-lg border border-orange-200/80 whitespace-pre-wrap">
                                                  {item.descrizioneDistinta}
                                              </div>
                                          </div>
                                      );
                                  }

                                  // CARD COMPLETA PER CANTIERE O IMPEGNO PREGRESSO
                                  return (
                                      <div 
                                          key={item.codice}
                                          draggable={!(item.resoEstintori && item.giorniDallUltimoSos < 5) && !item.isPregresso}
                                          onDragStart={(e) => handleDragStartKanban(e, item.codice)}
                                          onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDragOverCard(item.codice); }}
                                          onDragLeave={() => setDragOverCard(null)}
                                          onDrop={(e) => handleDropToDay(e, giornoSceltoFiltro === "sospesi" ? null : giornoSceltoFiltro, item.codice)}
                                          className={`group relative flex flex-col p-3.5 rounded-xl border transition-all
                                              ${dragOverCard === item.codice ? 'border-t-4 border-t-blue-600 mt-1' : ''}
                                              ${item.isPregresso ? 'bg-indigo-50/70 border-indigo-200' : 'bg-white border-slate-200 hover:border-blue-400 cursor-grab active:cursor-grabbing shadow-xs hover:shadow-md'}
                                          `}
                                      >
                                          {/* RIGA 1: TESTATA ORARIO, VIAGGIO, BI, TAG E AZIONI */}
                                          <div className="flex flex-wrap items-center justify-between gap-2 pb-2 mb-2 border-b border-slate-100">
                                              <div className="flex flex-wrap items-center gap-1.5">
                                                  {giornoSceltoFiltro !== "sospesi" && (
                                                      <span className="w-6 h-6 rounded-full bg-slate-900 text-white flex items-center justify-center text-[11px] font-extrabold shrink-0">
                                                          {numTappa}
                                                      </span>
                                                  )}

                                                  {item.isPregresso ? (
                                                      <span className="text-[11px] font-mono font-extrabold text-indigo-900 bg-indigo-100 px-2 py-0.5 rounded border border-indigo-200 flex items-center gap-1">
                                                          <Lock size={10}/> TASSATIVO {item.oraInizio} - {item.oraFine}
                                                      </span>
                                                  ) : giornoSceltoFiltro !== "sospesi" ? (
                                                      <span className="text-[11px] font-mono font-extrabold text-blue-800 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                                                          ⏰ {item.oraInizio} - {item.oraFine}
                                                      </span>
                                                  ) : null}

                                                  {giornoSceltoFiltro !== "sospesi" && item.minutiDaPrecedente !== undefined && (
                                                      <span className="text-[10px] font-bold text-slate-600 bg-slate-100 px-2 py-0.5 rounded border border-slate-200 flex items-center gap-1">
                                                          <Car size={11}/> +{item.minutiDaPrecedente}m {item.isPartenzaDaSede ? "da Sede" : "viaggio"}
                                                      </span>
                                                  )}

                                                  <span className="text-[10px] font-mono font-extrabold text-slate-700 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                                                      BI: {item.numeroBi}
                                                  </span>

                                                  <span className="text-[10px] font-extrabold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                                                      ⏱ {item.minutiLavoro} min
                                                  </span>

                                                  {!item.isPregresso && deveMostrareDC(item.nome) && (
                                                      <span className="bg-slate-800 text-white text-[9px] font-extrabold px-1.5 py-0.5 rounded">DC</span>
                                                  )}
                                                  {item.resoEstintori && (
                                                      <span className="bg-amber-100 text-amber-900 border border-amber-300 text-[10px] font-extrabold px-2 py-0.5 rounded">SOS • RESO ESTINTORI</span>
                                                  )}
                                                  {item.haInsoluto && (
                                                      <span className="bg-red-100 text-red-800 border border-red-300 text-[10px] font-extrabold px-2 py-0.5 rounded">INS • INSOLUTO</span>
                                                  )}
                                                  {item.syncedToGoogle && !item.isPregresso && (
                                                      <span className="bg-emerald-100 text-emerald-800 text-[9px] font-extrabold px-1.5 py-0.5 rounded">✓ Sync</span>
                                                  )}
                                              </div>

                                              {/* PULSANTI AZIONE */}
                                              <div className="flex items-center gap-1.5">
                                                  {giornoSceltoFiltro === "sospesi" && !(item.resoEstintori && item.giorniDallUltimoSos < 5) && (
                                                      <button onClick={(e) => { e.stopPropagation(); apriModaleSospeso(item); }} className="bg-blue-600 hover:bg-blue-700 text-white px-2.5 py-1 rounded-lg text-[10px] font-extrabold cursor-pointer">
                                                          📅 Pianifica
                                                      </button>
                                                  )}
                                                  {!item.isPregresso && (
                                                      <div className="relative">
                                                          <button onClick={(e) => { e.stopPropagation(); setMenuAperto(menuAperto === item.codice ? null : item.codice); }} title="Sposta in un altro giorno" className="text-slate-700 hover:text-blue-700 px-2 py-1 rounded-lg hover:bg-blue-50 transition-colors cursor-pointer bg-slate-50 border border-slate-200 flex items-center gap-1 text-[10px] font-bold">
                                                              <ArrowRightLeft size={12}/> Sposta
                                                          </button>
                                                          {menuAperto === item.codice && (
                                                              <div className="absolute top-full right-0 mt-1 bg-white border border-slate-200 shadow-2xl rounded-xl p-1.5 z-[99] w-40">
                                                                  <div className="text-[9px] font-extrabold text-slate-400 px-2 py-1 uppercase">Assegna a...</div>
                                                                  {giornateStats.map(gs => (
                                                                      <button key={gs.dataStr} onClick={(e) => { e.stopPropagation(); spostaCantiereRapido(item.codice, gs.dataStr); }} className="w-full text-left px-2 py-1.5 text-[11px] font-bold text-slate-700 hover:bg-blue-50 hover:text-blue-700 rounded-lg transition-colors cursor-pointer">
                                                                          {gs.nomeGiorno} {formattaDataVisuale(gs.dataStr).slice(0,5)}
                                                                      </button>
                                                                  ))}
                                                                  {giornoSceltoFiltro !== "sospesi" && (
                                                                      <>
                                                                          <div className="border-t border-slate-100 my-1"></div>
                                                                          <button onClick={(e) => { e.stopPropagation(); spostaCantiereRapido(item.codice, null); }} className="w-full text-left px-2 py-1.5 text-[11px] font-bold text-amber-700 hover:bg-amber-50 rounded-lg transition-colors cursor-pointer">
                                                                              ⏸ Manda in Sospesi
                                                                          </button>
                                                                      </>
                                                                  )}
                                                              </div>
                                                          )}
                                                      </div>
                                                  )}
                                                  <button onClick={(e) => { e.stopPropagation(); setClienteSelezionatoScheda(item); }} title="Apri scheda dettagli" className="text-slate-600 hover:text-blue-600 p-1.5 rounded-lg hover:bg-blue-50 transition-colors cursor-pointer bg-slate-50 border border-slate-200">
                                                      <FileText size={13}/>
                                                  </button>
                                              </div>
                                          </div>

                                          {/* RIGA 2: NOME CLIENTE E INDIRIZZO COMPLETO */}
                                          <div className="mb-2">
                                              <div className="text-sm font-extrabold text-slate-900 leading-snug">
                                                  {item.nome}
                                              </div>
                                              <div className="text-xs font-semibold text-slate-600 flex items-center gap-1 mt-0.5">
                                                  <MapPin size={12} className="text-blue-600 shrink-0"/>
                                                  <span>{item.indirizzo || "Indirizzo non specificato"}{item.cap ? ` - ${item.cap}` : ""}{item.localita ? ` ${item.localita}` : ""}</span>
                                              </div>
                                          </div>

                                          {/* RIGA 3: REFERENTE, TELEFONO ED EMAIL (SEMPRE VISIBILI) */}
                                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 bg-slate-50 px-3 py-2 rounded-lg border border-slate-200/80 mb-2 text-[11px]">
                                              <div className="flex items-center gap-1.5 truncate">
                                                  <User size={12} className="text-slate-500 shrink-0"/>
                                                  <span className="text-slate-500 font-semibold">Ref:</span>
                                                  <span className="font-extrabold text-slate-800 truncate">{item.contatto || "Non indicato"}</span>
                                              </div>
                                              <div className="flex items-center gap-1.5 truncate">
                                                  <Phone size={12} className="text-emerald-600 shrink-0"/>
                                                  <span className="text-slate-500 font-semibold">Tel:</span>
                                                  {item.telefono ? (
                                                      <a href={`tel:${item.telefono}`} onClick={(e) => e.stopPropagation()} className="font-mono font-extrabold text-blue-700 hover:underline truncate">
                                                          {item.telefono}
                                                      </a>
                                                  ) : (
                                                      <span className="text-slate-400 italic">Nessun tel.</span>
                                                  )}
                                              </div>
                                              <div className="flex items-center gap-1.5 truncate">
                                                  <Mail size={12} className="text-slate-500 shrink-0"/>
                                                  <span className="font-mono text-slate-700 truncate">{item.email || "-"}</span>
                                              </div>
                                          </div>

                                          {/* RIGA 4: DATI DEL RAPPORTINO / ATTREZZATURE E MULETTI (SEMPRE VISIBILI) */}
                                          {item.isPregresso ? (
                                              <div className="bg-indigo-100/50 border border-indigo-200/80 rounded-lg px-3 py-2 text-[11px] text-indigo-950 whitespace-pre-wrap font-medium">
                                                  {item.descrizionePregressa}
                                              </div>
                                          ) : (
                                              <div className="bg-blue-50/50 border border-blue-200/70 rounded-lg px-3 py-2">
                                                  <div className="flex flex-wrap items-center justify-between gap-2 mb-1">
                                                      <span className="text-[10px] font-extrabold uppercase text-blue-900 flex items-center gap-1">
                                                          <Wrench size={11}/> Dati Rapportino & Attrezzature ({item.totaleArticoli || 0} pz totali)
                                                      </span>
                                                      {item.muletti_eorv > 0 && (
                                                          <span className="text-[10px] font-extrabold bg-orange-100 text-orange-900 border border-orange-300 px-2 py-0.5 rounded">
                                                              📦 Muletti EO_RV: {item.muletti_eorv} pz
                                                          </span>
                                                      )}
                                                  </div>
                                                  {item.dettaglioAttrezzature && item.dettaglioAttrezzature.length > 0 ? (
                                                      <div className="flex flex-wrap gap-1.5 mt-1">
                                                          {item.dettaglioAttrezzature.map((att: any, iAtt: number) => (
                                                              <span key={iAtt} className="bg-white border border-blue-200 text-slate-800 px-2 py-0.5 rounded text-[11px] font-bold">
                                                                  {att.descrizione}: <strong className="text-blue-700 font-mono">{att.quantita}</strong> <span className="text-[10px] text-slate-400">({att.minutiImpegno}m)</span>
                                                              </span>
                                                          ))}
                                                      </div>
                                                  ) : (
                                                      <span className="text-[11px] text-slate-500 italic">Intervento standard (10 min base)</span>
                                                  )}
                                              </div>
                                          )}

                                          {/* MOTIVO SOSPESO O RIENTRO IN SEDE */}
                                          {giornoSceltoFiltro === "sospesi" && item.motivoSospeso && (
                                              <div className="mt-2 text-[10px] font-bold text-amber-800 bg-amber-50 border border-amber-200 px-2.5 py-1 rounded-lg">
                                                  ⚠️ Stato: {item.motivoSospeso}
                                              </div>
                                          )}

                                          {giornoSceltoFiltro !== "sospesi" && item.isRientroSede && item.minutiVersoProssimo !== undefined && (
                                              <div className="mt-2 text-[10px] font-bold text-slate-700 bg-slate-100 border border-slate-200 px-2.5 py-1 rounded-lg flex items-center gap-1.5">
                                                  <HomeIcon size={12} className="text-slate-600"/> Rientro in Sede stimato: +{item.minutiVersoProssimo} min di viaggio
                                              </div>
                                          )}
                                      </div>
                                  );
                              })
                          )}
                      </div>
                  </div>

                  {/* PANNELLO DESTRO: MAPPA GIGANTE AFFIANCATA */}
                  <div className="hidden md:block flex-1 bg-slate-100 relative z-0">
                      {leafletLoaded ? (
                          <MapContainer key={giornoSceltoFiltro} center={[SEDE_UFFICIO_LAT, SEDE_UFFICIO_LNG]} zoom={11} style={{ height: "100%", width: "100%" }}>
                              <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
                              {listaAttiva.filter(i => i.lat && i.lng && !i.isDistinta).map((c, idx) => (
                                  <Marker key={c.codice} position={[c.lat, c.lng]} icon={createNumberedIcon(idx + 1)}>
                                      <Popup>
                                          <div className="p-1">
                                              <p className="font-bold text-xs text-slate-900">#{idx + 1} - {c.nome}</p>
                                              <p className="text-[10px] text-slate-600">{c.indirizzo}, {c.localita}</p>
                                              {c.telefono && <p className="text-[10px] font-mono text-slate-700 mt-0.5">📞 {c.telefono}</p>}
                                              <p className="text-[10px] font-mono font-bold text-blue-600 mt-1">{c.oraInizio || "Sospeso"} ({c.minutiLavoro}m)</p>
                                          </div>
                                      </Popup>
                                  </Marker>
                              ))}
                          </MapContainer>
                      ) : (
                          <div className="w-full h-full flex items-center justify-center text-slate-400">
                              <Loader2 size={32} className="animate-spin" />
                          </div>
                      )}
                  </div>
              </div>
          ) : (
              /* SCHERMATA INIZIALE RICCA QUANDO NESSUN FILE È ANCORA ELABORATO */
              <div className="flex-1 overflow-y-auto p-6 flex flex-col items-center justify-center">
                  <div className="max-w-4xl w-full">
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-5 mb-6">
                          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-xs relative overflow-hidden">
                              <div className="p-3 bg-blue-50 text-blue-600 rounded-xl inline-block mb-3"><Calendar size={24}/></div>
                              <span className="bg-blue-100 text-blue-800 font-extrabold text-[10px] px-2.5 py-1 rounded-md ml-2 uppercase">1. Periodo & Tecnico</span>
                              <h3 className="font-extrabold text-slate-900 text-base mb-1">Imposta la Settimana</h3>
                              <p className="text-xs text-slate-500 leading-relaxed">Configura nella barra qui sopra i giorni lavorativi attivi e seleziona il tecnico di riferimento.</p>
                          </div>
                          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-xs relative overflow-hidden">
                              <div className="p-3 bg-amber-50 text-amber-600 rounded-xl inline-block mb-3"><FileUp size={24}/></div>
                              <span className="bg-amber-100 text-amber-800 font-extrabold text-[10px] px-2.5 py-1 rounded-md ml-2 uppercase">2. Caricamento Dati</span>
                              <h3 className="font-extrabold text-slate-900 text-base mb-1">Importa il File Cantieri</h3>
                              <p className="text-xs text-slate-500 leading-relaxed">Clicca su <strong>Sfoglia Excel</strong>. Il sistema estrarrà referenti, telefoni e attrezzature del rapportino per ogni BI.</p>
                          </div>
                          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-xs relative overflow-hidden">
                              <div className="p-3 bg-emerald-50 text-emerald-600 rounded-xl inline-block mb-3"><Sliders size={24}/></div>
                              <span className="bg-emerald-100 text-emerald-800 font-extrabold text-[10px] px-2.5 py-1 rounded-md ml-2 uppercase">3. Multi-Agente IA</span>
                              <h3 className="font-extrabold text-slate-900 text-base mb-1">Genera Plancia & Mappa</h3>
                              <p className="text-xs text-slate-500 leading-relaxed">Premi <strong>Genera Plancia</strong>: Cartografo, Cronometrista e Supervisore creeranno i giri ottimali.</p>
                          </div>
                      </div>
                  </div>
              </div>
          )}
      </div>

      {/* MODALE PIANIFICA SOSPESO SINGOLO */}
      {sospesoInModifica && (
        <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-md z-[9999] flex items-center justify-center p-4 print:hidden">
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
                <input type="date" value={configSospesoSingolo.data} onChange={(e) => setConfigSospesoSingolo({...configSospesoSingolo, data: e.target.value})} className="w-full bg-slate-50 border border-slate-300 rounded-xl p-3 text-sm font-semibold text-slate-800 focus:outline-none" />
              </div>
              <div>
                <label className="text-[11px] font-bold text-slate-500 block mb-1">Ora Inizio stimata</label>
                <input type="time" value={configSospesoSingolo.oraInizio} onChange={(e) => setConfigSospesoSingolo({...configSospesoSingolo, oraInizio: e.target.value})} className="w-full bg-slate-50 border border-slate-300 rounded-xl p-3 text-sm font-semibold text-slate-800 focus:outline-none" />
              </div>
            </div>
            <div className="flex flex-col gap-2.5">
               <button onClick={() => creaSospesoDaModale('calendar')} disabled={creazioneInCorso[sospesoInModifica.codice]} className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3.5 rounded-xl text-sm font-bold flex items-center justify-center gap-2 shadow-md transition-all">
                  {creazioneInCorso[sospesoInModifica.codice] ? (
                    <span className="flex items-center gap-2"><Loader2 size={16} className="animate-spin"/> Invio in corso...</span>
                  ) : (
                    <span className="flex items-center gap-2"><CalendarPlus size={16}/> Crea su Calendar</span>
                  )}
               </button>
               <button onClick={() => creaSospesoDaModale('lista')} className="w-full bg-slate-100 hover:bg-slate-200 text-slate-700 py-3.5 rounded-xl text-sm font-bold flex items-center justify-center gap-2 transition-all"><CheckSquare size={16}/> Aggiungi solo in Lista</button>
            </div>
          </div>
        </div>
      )}

      {/* MODALE DETTAGLIO SCHEDA CLIENTE */}
      {clienteSelezionatoScheda && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9999] flex items-center justify-center p-4 overflow-y-auto print:hidden">
          <div className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full p-8 relative border border-slate-200 my-auto max-h-[90vh] flex flex-col">
            <button onClick={() => setClienteSelezionatoScheda(null)} className="absolute top-6 right-6 text-slate-400 bg-slate-50 border border-slate-200 p-2 rounded-full hover:bg-slate-100 hover:text-slate-900 transition-colors"><X size={18} /></button>
            
            {clienteSelezionatoScheda.isDistinta ? (
              <>
                <div className="flex items-center gap-4 mb-6">
                  <div className="p-4 bg-slate-100 text-slate-900 rounded-xl border border-slate-200"><Package size={24} /></div>
                  <div>
                    <span className="bg-slate-100 text-slate-700 font-mono text-[10px] px-2 py-0.5 rounded border border-slate-200 font-bold mb-1 inline-block">MAGAZZINO</span>
                    <h2 className="text-xl font-extrabold text-slate-900">Distinta di Carico</h2>
                  </div>
                </div>
                <div className="bg-slate-50 p-6 rounded-xl border border-slate-200 text-xs whitespace-pre-wrap font-mono leading-relaxed mb-6 overflow-y-auto flex-1">
                  {clienteSelezionatoScheda.descrizioneDistinta}
                </div>
              </>
            ) : (
              <>
                <div className="flex items-center gap-4 mb-6 pb-4 border-b border-slate-100">
                  <div className="p-4 bg-slate-100 text-slate-900 rounded-xl border border-slate-200"><FileText size={24} /></div>
                  <div>
                    <div className="flex gap-2 mb-1">
                       <span className="bg-slate-100 text-slate-700 font-mono text-[10px] px-2 py-0.5 rounded border border-slate-200 font-bold">BI: {clienteSelezionatoScheda.numeroBi}</span>
                       {clienteSelezionatoScheda.haInsoluto && <span className="bg-red-50 text-red-700 border border-red-200 font-mono text-[10px] px-2 py-0.5 rounded font-bold">INSOLUTO</span>}
                       {clienteSelezionatoScheda.resoEstintori && <span className="bg-orange-50 text-orange-700 border border-orange-200 font-mono text-[10px] px-2 py-0.5 rounded font-bold">RESO ESTINTORI</span>}
                    </div>
                    <h2 className="text-xl font-extrabold text-slate-900">{clienteSelezionatoScheda.nome}</h2>
                    <p className="text-xs text-slate-500 font-medium mt-1">{clienteSelezionatoScheda.indirizzo}, {clienteSelezionatoScheda.localita}</p>
                  </div>
                </div>

                <div className="space-y-6 overflow-y-auto pr-2 flex-1 mb-6 custom-scrollbar">
                   <div className="grid grid-cols-1 md:grid-cols-2 gap-3 bg-slate-50 p-4 rounded-xl border border-slate-200 text-xs">
                      <div><strong className="text-slate-500 block mb-0.5 uppercase text-[9px] tracking-wider">Referente</strong> <span className="font-bold text-slate-900">{clienteSelezionatoScheda.contatto || "-"}</span></div>
                      <div><strong className="text-slate-500 block mb-0.5 uppercase text-[9px] tracking-wider">Telefono</strong> <span className="font-bold text-slate-900">{clienteSelezionatoScheda.telefono || "-"}</span></div>
                      <div className="col-span-2"><strong className="text-slate-500 block mb-0.5 uppercase text-[9px] tracking-wider">Email</strong> <span className="font-bold text-slate-900">{clienteSelezionatoScheda.email || "-"}</span></div>
                   </div>

                   <div>
                      <h4 className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-2 border-b border-slate-100 pb-1">Attrezzature del Rapportino ({clienteSelezionatoScheda.totaleArticoli || 0} pz)</h4>
                      {clienteSelezionatoScheda.dettaglioAttrezzature && clienteSelezionatoScheda.dettaglioAttrezzature.length > 0 ? (
                         <div className="border border-slate-200 rounded-xl overflow-hidden divide-y divide-slate-100">
                            {clienteSelezionatoScheda.dettaglioAttrezzature.map((att: any, idx: number) => (
                               <div key={idx} className="p-3 flex justify-between items-center bg-white text-xs">
                                  <span className="font-bold text-slate-800">{att.descrizione}</span>
                                  <span className="font-mono bg-slate-50 border border-slate-200 px-2 py-0.5 rounded text-slate-700">Q.tà: {att.quantita} ({att.minutiImpegno}m)</span>
                               </div>
                            ))}
                         </div>
                      ) : (
                         <p className="text-xs text-slate-400 italic bg-slate-50 p-3 rounded-xl border border-slate-200">Standard</p>
                      )}
                   </div>
                </div>
              </>
            )}

            <div className="pt-4 border-t border-slate-100 flex justify-end">
              <button onClick={() => setClienteSelezionatoScheda(null)} className="bg-slate-900 text-white px-6 py-2.5 rounded-xl text-xs font-bold shadow-sm hover:bg-slate-800 transition-all">Chiudi</button>
            </div>
          </div>
        </div>
      )}

      {/* VISTA STAMPA FOGLIO DI VIAGGIO COMPLETA CON REFERENTI E TELEFONI */}
      <div className="hidden print:block w-full text-black bg-white font-sans text-sm p-6">
         {giorniVisibiliPerStampa.map((dataStr: any) => {
            const interventiGiorno = interventiGrezzi.filter(i => i.dataAssegnata === dataStr && (i.selezionatoPerGiro || i.isPregresso));
            if (interventiGiorno.length === 0) return null;

            const distinta = interventiGiorno.find(i => i.isDistinta);
            const lavori = interventiGiorno.filter(i => !i.isDistinta).sort((a,b) => (a.oraInizio || "").localeCompare(b.oraInizio || ""));
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
                       <pre className="font-mono text-xs whitespace-pre-wrap">{distinta.descrizioneDistinta}</pre>
                    </div>
                  )}

                  <div className="mt-4">
                     <h2 className="font-extrabold text-base mb-3 uppercase border-b border-black pb-1">Elenco Interventi Giornalieri</h2>
                     <table className="w-full text-left text-xs border-collapse border border-black">
                        <thead>
                           <tr className="bg-gray-100 border-b border-black">
                              <th className="p-2 border border-black w-24">Orario</th>
                              <th className="p-2 border border-black w-48">Cliente & BI</th>
                              <th className="p-2 border border-black w-48">Indirizzo & Contatti</th>
                              <th className="p-2 border border-black">Dati Rapportino / Attrezzature</th>
                           </tr>
                        </thead>
                        <tbody>
                           {lavori.map((lav, idx) => (
                              <tr key={idx} className="border-b border-black">
                                 <td className="p-2 border border-black font-mono font-bold align-top">{lav.oraInizio} - {lav.oraFine}</td>
                                 <td className="p-2 border border-black font-bold align-top">
                                    {lav.nome}
                                    <span className="block font-mono font-normal text-[10px] text-gray-600 mt-0.5">BI: {lav.numeroBi}</span>
                                 </td>
                                 <td className="p-2 border border-black align-top">
                                    <div>{lav.indirizzo} {lav.localita ? `- ${lav.localita}` : ""}</div>
                                    {(lav.contatto || lav.telefono) && (
                                       <div className="font-mono text-[10px] mt-1">
                                          {lav.contatto ? `Ref: ${lav.contatto} ` : ""}{lav.telefono ? `Tel: ${lav.telefono}` : ""}
                                       </div>
                                    )}
                                 </td>
                                 <td className="p-2 border border-black align-top text-[11px]">
                                    {lav.dettaglioAttrezzature && lav.dettaglioAttrezzature.length > 0 
                                       ? lav.dettaglioAttrezzature.map((d: any) => `${d.descrizione}: ${d.quantita}`).join(", ") 
                                       : "Verifica standard"}
                                    {lav.muletti_eorv > 0 ? ` (Muletti: ${lav.muletti_eorv})` : ""}
                                 </td>
                              </tr>
                           ))}
                        </tbody>
                     </table>
                  </div>
               </div>
            );
         })}
      </div>

      {/* CHAT IA FLUTTUANTE */}
      <button
        onClick={() => setIsChatOpen(!isChatOpen)}
        className={`fixed bottom-6 right-6 z-[100] w-14 h-14 rounded-full shadow-lg flex items-center justify-center transition-all duration-300 print:hidden ${
          isChatOpen ? 'bg-slate-900 text-white hover:bg-slate-800' : 'bg-blue-600 text-white hover:bg-blue-700'
        }`}
      >
        {isChatOpen ? <X size={24} /> : <MessageSquare size={24} className={chatLoading ? 'animate-pulse' : ''} />}
      </button>

      <div 
        className={`fixed bottom-24 right-6 z-[90] w-[360px] h-[520px] max-h-[75vh] bg-white rounded-2xl shadow-2xl border border-slate-200 flex flex-col overflow-hidden transition-all duration-300 origin-bottom-right print:hidden ${
          isChatOpen ? 'scale-100 opacity-100 translate-y-0' : 'scale-95 opacity-0 translate-y-8 pointer-events-none'
        }`}
      >
        <div className="p-4 bg-slate-900 text-white flex justify-between items-center shrink-0">
          <div className="flex items-center gap-2">
            <Zap size={16} className="text-blue-400" />
            <h3 className="font-bold text-xs tracking-wide">Assistente IA Operativo</h3>
          </div>
          <button onClick={() => setIsChatOpen(false)} className="text-slate-400 hover:text-white transition-colors"><X size={16} /></button>
        </div>

        <div className="flex-1 p-4 overflow-y-auto space-y-4 bg-slate-50 custom-scrollbar">
          {messages.map((msg, idx) => (
            <div key={idx} className={`flex ${msg.sender === 'user' ? 'justify-end' : 'justify-start'}`}>
              <div className={`max-w-[85%] px-3 py-2 text-xs leading-relaxed shadow-xs ${
                msg.sender === 'user' 
                  ? 'bg-blue-600 text-white rounded-xl rounded-br-sm' 
                  : 'bg-white text-slate-700 border border-slate-200 rounded-xl rounded-bl-sm'
              }`}>
                {msg.text}
              </div>
            </div>
          ))}
          {chatLoading && (
            <div className="flex justify-start">
              <div className="bg-white text-slate-500 border border-slate-200 px-3 py-2 rounded-xl rounded-bl-sm text-xs shadow-xs flex items-center gap-2">
                <Loader2 size={12} className="animate-spin text-blue-600" /> Elaborazione...
              </div>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        <form onSubmit={handleSendMessage} className="p-3 bg-white border-t border-slate-100 flex gap-2 items-center shrink-0">
          <input 
            type="text"
            value={chatInput}
            onChange={(e) => setChatInput(e.target.value)}
            placeholder="Es. Sposta Rossi a giovedì mattina..."
            className="flex-1 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-blue-500 transition-colors"
          />
          <button type="submit" disabled={chatLoading} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-bold transition-colors disabled:opacity-50 shrink-0">
            Invia
          </button>
        </form>
      </div>

    </div>
  );
}