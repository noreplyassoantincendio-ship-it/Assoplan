"use client";

import Link from "next/link";
import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import dynamic from "next/dynamic";
import { UploadCloud, AlertCircle, Package, MapPin, Loader2, FileSpreadsheet, MapPinOff, Zap, UserCheck, Printer, Calendar, Clock, CheckSquare, Square, Mail, Timer, FileText, X, Sliders, Check, Trash2, ArrowRight, PauseCircle, PlusCircle, ExternalLink, ShieldAlert, AlertTriangle, CheckCircle, PlayCircle, Lock, CalendarPlus, GripVertical, Phone, Book, User, Banknote, Search, Database, Car, Home as HomeIcon, AlertOctagon, MessageSquare } from "lucide-react";

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
    // @ts-ignore
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
        // @ts-ignore
        const service = new window.google.maps.DistanceMatrixService();
        // @ts-ignore
        const origin = new window.google.maps.LatLng(lat1, lon1);
        // @ts-ignore
        const destination = new window.google.maps.LatLng(lat2, lon2);
        
        const response: any = await new Promise((resolve, reject) => {
            service.getDistanceMatrix({
                origins: [origin],
                destinations: [destination],
                // @ts-ignore
                travelMode: window.google.maps.TravelMode.DRIVING,
            }, (res: any, status: any) => {
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
        console.error("Errore Distance Matrix API JS SDK:", e);
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

  const [giornoSceltoFiltro, setGiornoSceltoFiltro] = useState<string>("Tutti");
  const [clienteSelezionatoScheda, setClienteSelezionatoScheda] = useState<any | null>(null);
  
  const [sospesoInModifica, setSospesoInModifica] = useState<any | null>(null);
  const [configSospesoSingolo, setConfigSospesoSingolo] = useState<{data: string, oraInizio: string}>({data: dataInizio, oraInizio: "08:30"});

  const [giornateStats, setGiornateStats] = useState<any[]>([]);
  const [creazioneInCorso, setCreazioneInCorso] = useState<{[key:string]: boolean}>({});
  const [bulkSyncStatus, setBulkSyncStatus] = useState({ active: false, progress: 0, current: 0, total: 0 });

  // CHAT FLUTTUANTE
  const [isChatOpen, setIsChatOpen] = useState(false);
  const [chatInput, setChatInput] = useState('');
  const [messages, setMessages] = useState([
    { sender: 'ai', text: 'Ciao! Sono il tuo assistente operativo. Carica il file Excel e chiedimi di pianificare o spostare cantieri.' }
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
    if (isChatOpen) {
        messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
    }
  }, [messages, chatLoading, isChatOpen]);

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

  const trovaCoordinateGoogle = async (indirizzo: string, localita: string) => {
    let cleanAddr = indirizzo.replace(/Nessun indirizzo valido letto/gi, '').trim();
    
    if (cleanAddr.toLowerCase().includes("acquasanta") || (localita && localita.toLowerCase().includes("acquasanta"))) {
      return { lat: 44.4568, lng: 8.7678, precisione: "strada", origine: "override_manuale" };
    }

    let citta = localita && localita.trim().length > 0 ? localita.trim() : "";
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
    if (!cliente.nome.toUpperCase().includes("ASL")) tags.push("[DC]");
    
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

  const ricalcolaDistinteEOrari = async (listaDaOrdinare: any[], statsLocali: any[] = giornateStats, usaOrdinamentoGeografico = true) => {
      const listaPulita = listaDaOrdinare.filter(i => !i.isDistinta && i.dataAssegnata && typeof i.dataAssegnata === 'string');
      const mapGiorni = new Map();
      const vecchieDistinte = listaDaOrdinare.filter(i => i.isDistinta);
      const nuoveStats = statsLocali.map(g => ({ ...g }));
      
      const scartatiTemporanei: any[] = []; 

      listaPulita.forEach(i => {
        if (!mapGiorni.has(i.dataAssegnata)) mapGiorni.set(i.dataAssegnata, []);
        mapGiorni.get(i.dataAssegnata).push(i);
      });

      const nuovaLista: any[] = [];

      for (const [dataStr, interventiGiorno] of Array.from(mapGiorni.entries())) {
        const stat = statsLocali.find(g => g.dataStr === dataStr);
        let orarioCorrenteMinuti = 8 * 60; 
        let ultimaCoordGiorno: { lat: number; lng: number } | null = null;
        let listaFinaleGiorno: any[] = [];
        
        if (usaOrdinamentoGeografico) {
            let daOrdinare = interventiGiorno.filter(i => !i.isPregresso && i.selezionatoPerGiro);
            const pregressiGiorno = interventiGiorno.filter(c => c.isPregresso);
            
            if (pregressiGiorno.length > 0) {
                let maxPregressoEnd = 8 * 60;
                pregressiGiorno.forEach(p => {
                    const oraFineStr = p.oraFine || "18:00";
                    const [h, m] = String(oraFineStr).split(":").map(Number);
                    const fine = (Number(h) * 60) + Number(m);
                    if (fine > maxPregressoEnd) maxPregressoEnd = fine;
                    if (p.lat && p.lng && !ultimaCoordGiorno) {
                        ultimaCoordGiorno = { lat: p.lat, lng: p.lng };
                    }
                });
                orarioCorrenteMinuti = Math.max(Number(orarioCorrenteMinuti), Number(maxPregressoEnd) + 30);
                if (orarioCorrenteMinuti >= INIZIO_PAUSA && orarioCorrenteMinuti < FINE_PAUSA) {
                    orarioCorrenteMinuti = FINE_PAUSA;
                }
            } else if (stat && stat.anchorLat) {
                ultimaCoordGiorno = { lat: stat.anchorLat, lng: stat.anchorLng };
            }

            let ordinatiGoogle = [];
            let currPos = ultimaCoordGiorno || { lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG };
            
            while (daOrdinare.length > 0) {
                daOrdinare.sort((a,b) => calcolaDistanzaKm(currPos.lat, currPos.lng, a.lat, a.lng) - calcolaDistanzaKm(currPos.lat, currPos.lng, b.lat, b.lng));
                const next = daOrdinare.shift();
                ordinatiGoogle.push(next);
                if (next.lat && next.lng) currPos = { lat: next.lat, lng: next.lng };
            }

            listaFinaleGiorno = [...pregressiGiorno, ...ordinatiGoogle];
        } else {
            listaFinaleGiorno = [...interventiGiorno];
            if (stat && stat.anchorLat && (!listaFinaleGiorno[0] || !listaFinaleGiorno[0].isPregresso)) {
                ultimaCoordGiorno = { lat: stat.anchorLat, lng: stat.anchorLng };
            }
        }

        let furgone = 0; let piccoNegativo = 0; let totaleMovimenti = 0;
        let prevCliente: any = null;
        ultimaCoordGiorno = null;
        
        let clientiAmmessi = [];

        for (const c of listaFinaleGiorno) {
            if (c.isPregresso) {
                const oraFineStr = c.oraFine || "18:00";
                const [hEnd, mEnd] = String(oraFineStr).split(":").map(Number);
                
                const finePregressoMinuti = (Number(hEnd) * 60 + Number(mEnd));
                orarioCorrenteMinuti = Math.max(Number(orarioCorrenteMinuti), finePregressoMinuti + 30);
                
                if (prevCliente && ultimaCoordGiorno && c.lat && c.lng) {
                    const { minuti } = await calcolaTempoDistanzaGoogle(ultimaCoordGiorno.lat, ultimaCoordGiorno.lng, c.lat, c.lng);
                    prevCliente.minutiVersoProssimo = Number(minuti);
                    prevCliente.isRientroSede = false;
                }
                
                c.minutiDaPrecedente = 0;
                c.isPartenzaDaSede = !prevCliente;
                if (c.lat && c.lng) ultimaCoordGiorno = { lat: c.lat, lng: c.lng };
                prevCliente = c;
                clientiAmmessi.push(c);
            } else {
                const puntoPartenza = ultimaCoordGiorno || { lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG };
                const { minuti: minViaggio } = await calcolaTempoDistanzaGoogle(puntoPartenza.lat, puntoPartenza.lng, c.lat, c.lng);

                const { inizioLavoro, fineLavoro } = calcolaTempistiche(orarioCorrenteMinuti, minViaggio, c.minutiLavoro);
                
                if (Number(fineLavoro) > 1080) {
                   c.selezionatoPerGiro = false;
                   c.dataAssegnata = null;
                   c.motivoSospeso = "Scartato per Limite Orario (Overtime)";
                   scartatiTemporanei.push(c);
                   continue; 
                }

                c.minutiDaPrecedente = Number(minViaggio);
                c.isPartenzaDaSede = !prevCliente;

                if (prevCliente) {
                    prevCliente.minutiVersoProssimo = Number(minViaggio);
                    prevCliente.isRientroSede = false;
                }

                const hhStart = String(Math.floor(inizioLavoro / 60)).padStart(2, '0');
                const mmStart = String(inizioLavoro % 60).padStart(2, '0');
                const hhEnd = String(Math.floor(fineLavoro / 60)).padStart(2, '0');
                const mmEnd = String(fineLavoro % 60).padStart(2, '0');

                c.oraInizio = `${hhStart}:${mmStart}`;
                c.oraFine = `${hhEnd}:${mmEnd}`;

                orarioCorrenteMinuti = fineLavoro;
                ultimaCoordGiorno = { lat: c.lat, lng: c.lng };
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

        let descStr = `🛠 DISTINTA MULETTI:\n\n• Fabbisogno Netto: ${fabbisognoNetto} pz\n• Scorta Jolly (20%): ${scortaJolly} pz\n-----------------------------------\nTOTALE MULETTI DA CARICARE: ${totaleDaCaricare} pz\n-----------------------------------\n(Movimentazione totale: ${totaleMovimenti} pz)`;

        const primaRiga = clientiAmmessi[0] || { giorno: getNomeGiorno(dataStr) };
        const oldDistinta = vecchieDistinte.find(d => d.dataAssegnata === dataStr);

        const distintaObj = {
            codice: `DISTINTA_${dataStr}`, isDistinta: true, nome: `📦 DISTINTA DI CARICO (${formattaDataVisuale(dataStr)})`, numeroBi: "MAGAZZINO", indirizzo: "Sede Asso Antincendio", localita: "Genova", minutiLavoro: 60, giorno: primaRiga.giorno, dataAssegnata: dataStr, oraInizio: "07:00", oraFine: "08:00", descrizioneDistinta: descStr, selezionatoPerGiro: true, syncedToGoogle: oldDistinta ? oldDistinta.syncedToGoogle : false, lat: SEDE_UFFICIO_LAT, lng: SEDE_UFFICIO_LNG
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
          haInsoluto: false, resoEstintori: false, totaleArticoli: 0, minutiLavoro: 10, dettaglioAttrezzature: [], selezionatoPerGiro: false, giorniDallUltimoSos: 8, syncedToGoogle: false, muletti_eorv: 0,
          _rawItems: [], _maxMt: { eo: 0, po: 0, id: 0, la: 0, ir: 0, ss: 0, ri: 0, gp: 0, rv: 0 }
        });
      }

      const cliente = clientiMappa.get(idCliente);
      const testoRiga = JSON.stringify(riga).toUpperCase();
      const numFattura = riga.numft ? String(riga.numft).trim() : "";
      if (numFattura !== "" && numFattura !== "-" && (parseFloat(riga.imptotft) || 0) > 0 && !(riga.codpagam || "").toUpperCase().trim().startsWith("RB")) cliente.haInsoluto = true;
      if ((riga.codclose ? String(riga.codclose).toUpperCase().trim() : "") === "SOS" || testoRiga.includes("SOS")) cliente.resoEstintori = true;

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
                cliente.dettaglioAttrezzature = [];
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
                    { v: sumQtaEstintori, m: 2, n: "Estintori (Da Righe)" },
                    { v: sumQtaIdranti, m: 2, n: "Idranti (Da Righe)" },
                    { v: sumQtaPorte, m: 3, n: "Porte Tagliafuoco (Da Righe)" },
                    { v: sumQtaUscite, m: 2, n: "Uscite Emergenza (Da Righe)" },
                    { v: sumQtaLampade, m: 1, n: "Lampade Emergenza (Da Righe)" },
                    { v: sumQtaIrai, m: 30, n: "Impianto Rilevazione (Da Righe)" },
                    { v: sumQtaReti, m: 15, n: "Rete Idrica (Da Righe)" },
                    { v: sumQtaGruppi, m: 20, n: "Gruppo Pressurizzazione (Da Righe)" }
                ];
                itemsRighe.forEach(it => { 
                    if(it.v > 0) { 
                        minutiTotali += it.v * it.m; pezziTotali += it.v; 
                        dettaglio.push({ descrizione: it.n, quantita: it.v, categoria: `(${it.m}m/pz)`, minutiImpegno: it.v * it.m }); 
                    } 
                });
            } else {
                const itemsMt = [
                    { v: cliente._maxMt.eo, m: 2, n: "Estintori (Manutenzione)" },
                    { v: cliente._maxMt.id, m: 2, n: "Idranti" },
                    { v: cliente._maxMt.po, m: 3, n: "Porte Tagliafuoco / REI" },
                    { v: cliente._maxMt.ss, m: 2, n: "Uscite di Emergenza" },
                    { v: cliente._maxMt.la, m: 1, n: "Lampade di Emergenza" },
                    { v: cliente._maxMt.ir, m: 30, n: "Impianto Rilevazione (IRAI)" },
                    { v: cliente._maxMt.ri, m: 15, n: "Rete Idrica" },
                    { v: cliente._maxMt.gp, m: 20, n: "Gruppo Pressurizzazione" }
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
      
      clientiDaPianificare.push(c);
    }

    const giornateStrutturate = dateRange.map(g => {
      const pregressiDelGiorno = eventiPregressiUI.filter(e => e.dataAssegnata === g.dataStr);
      let maxFineMinuti = 8 * 60;
      let anchorLat: number | null = null; let anchorLng: number | null = null;

      pregressiDelGiorno.forEach(imp => {
        const oraFineStr = imp.oraFine || "18:00";
        const [h, m] = String(oraFineStr).split(":").map(Number);
        const fine = (Number(h) * 60) + Number(m);
        if (fine > maxFineMinuti) maxFineMinuti = fine;
        if (imp.lat && imp.lng && !anchorLat) { anchorLat = imp.lat; anchorLng = imp.lng; }
      });

      if (maxFineMinuti > 8 * 60) maxFineMinuti += 30;

      if (maxFineMinuti >= INIZIO_PAUSA && maxFineMinuti < FINE_PAUSA) {
         maxFineMinuti = FINE_PAUSA;
      }

      maxFineMinuti = Math.min(maxFineMinuti, FINE_GIORNATA_ASSOLUTA);
      let minutiDisponibili = Math.max(0, FINE_GIORNATA_ASSOLUTA - maxFineMinuti);
      let statoGiornata = "APERTO";
      if (minutiDisponibili < 60) { minutiDisponibili = 0; statoGiornata = "PIENO (Sospeso)"; }

      return { ...g, minutiDisponibili, statoGiornata, anchorLat, anchorLng, orarioPartenzaMinuti: maxFineMinuti };
    });

    let unassigned = [...clientiDaPianificare];
    const clientiPianificati = [];

    try {
        const payload = {
            interventi: unassigned.map(c => ({
                codice: c.codice,
                nome: c.nome,
                indirizzo: c.indirizzo,
                localita: c.localita,
                minutiLavoro: c.minutiLavoro
            })),
            giornate: giornateStrutturate.map(g => ({
                dataStr: g.dataStr,
                nomeGiorno: g.nomeGiorno,
                pregressi: eventiPregressiUI.filter(p => p.dataAssegnata === g.dataStr).map(p => ({
                    nome: p.nome,
                    indirizzo: p.indirizzo,
                    descrizione: p.descrizionePregressa
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
        
        const iaResult = await res.json();
        
        if (iaResult.success && iaResult.piano) {
            unassigned.forEach(c => {
                let assigned = false;
                for (const [dataStr, codici] of Object.entries(iaResult.piano.giorni)) {
                    if ((codici as string[]).includes(c.codice)) {
                        c.dataAssegnata = dataStr;
                        c.giorno = getNomeGiorno(dataStr);
                        c.selezionatoPerGiro = true;
                        clientiPianificati.push(c);
                        assigned = true;
                        break;
                    }
                }
                if (!assigned) {
                    clientiSospesiTmp.push({ ...c, motivoSospeso: "Scartato dall'IA (Logistica / Tempo)", selezionatoPerGiro: false, dataAssegnata: null });
                }
            });
        } else {
            throw new Error(iaResult.error || "L'Agente IA non ha restituito un piano valido");
        }
    } catch (e: any) {
        console.error("Dettaglio errore IA:", e);
        if (typeof window !== 'undefined') {
            window.alert(`Errore IA: ${e.message || "Errore sconosciuto"}. I cantieri sono stati inseriti nei Sospesi.`);
        }
        unassigned.forEach(c => {
            clientiSospesiTmp.push({ ...c, motivoSospeso: `Errore: ${e.message || "Sconosciuto"}`, selezionatoPerGiro: false, dataAssegnata: null });
        });
    }

    const datiCombinati = [...eventiPregressiUI, ...clientiPianificati].sort((a, b) => {
        if (a.dataAssegnata !== b.dataAssegnata) return a.dataAssegnata.localeCompare(b.dataAssegnata);
        if (a.isPregresso && !b.isPregresso) return -1;
        if (!a.isPregresso && b.isPregresso) return 1;
        if (a.isPregresso && b.isPregresso) return a.oraInizio.localeCompare(b.oraInizio);
        return 0;
    });

    const resRic = await ricalcolaDistinteEOrari(datiCombinati, giornateStrutturate, true);
    setInterventiGrezzi(resRic.listaOrdinata);
    setGiornateStats(resRic.statsAggiornate);
    setClientiInSospeso(clientiSospesiTmp); 
    setClientiGiaCalendarizzati(clientiScartatiTmp);
    setInElaborazione(false);
  };

  const handleDragStartKanban = (e: React.DragEvent, codice: string) => {
      e.dataTransfer.setData("application/kanban-card", codice);
  };

  const handleDragOverKanban = (e: React.DragEvent) => { 
      e.preventDefault(); 
  };

  const handleDropToDay = async (e: React.DragEvent, targetDataAssegnata: string, targetCodice?: string) => {
      e.preventDefault();
      e.stopPropagation();
      setDragOverCard(null);

      const draggedCodice = e.dataTransfer.getData("application/kanban-card");
      if (!draggedCodice || draggedCodice === targetCodice) return;

      setInElaborazione(true);
      try {
          let result = Array.from(interventiGrezzi);
          let item = result.find(i => i.codice === draggedCodice);
          let isFromSospesi = false;

          if (!item) {
              const sospesoItem = clientiInSospeso.find(i => i.codice === draggedCodice);
              if (sospesoItem) {
                  item = { ...sospesoItem, selezionatoPerGiro: true, dataAssegnata: targetDataAssegnata };
                  isFromSospesi = true;
              }
          }

          if (!item) return;

          if (isFromSospesi) {
              setClientiInSospeso(prev => prev.filter(c => c.codice !== draggedCodice));
          } else {
              result = result.filter(i => i.codice !== draggedCodice);
          }

          item.dataAssegnata = targetDataAssegnata;

          if (targetCodice) {
              const targetIndex = result.findIndex(i => i.codice === targetCodice && i.dataAssegnata === targetDataAssegnata);
              if (targetIndex !== -1) {
                  result.splice(targetIndex, 0, item);
              } else {
                  result.push(item);
              }
          } else {
              result.push(item);
          }

          const resRic = await ricalcolaDistinteEOrari(result, giornateStats, false);
          setGiornateStats(resRic.statsAggiornate);
          setInterventiGrezzi(resRic.listaOrdinata);
      } finally {
          setInElaborazione(false);
      }
  };

  const handleDropToSospesi = async (e: React.DragEvent) => {
      e.preventDefault();
      setDragOverCard(null);
      const draggedCodice = e.dataTransfer.getData("application/kanban-card");
      if (!draggedCodice) return;

      setInElaborazione(true);
      try {
          const itemToRemove = interventiGrezzi.find(i => i.codice === draggedCodice);
          if (itemToRemove && !itemToRemove.isDistinta && !itemToRemove.isPregresso) {
              const nuoviGrezzi = interventiGrezzi.filter(i => i.codice !== draggedCodice);
              setClientiInSospeso(prev => [...prev, { ...itemToRemove, motivoSospeso: "Spostato in Sospesi", selezionatoPerGiro: false, dataAssegnata: null }]);
              const resRic = await ricalcolaDistinteEOrari(nuoviGrezzi, giornateStats, false);
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

  const handleSendMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!chatInput.trim() || chatLoading) return;

    const userMessage = chatInput;
    const updatedMessages = [...messages, { sender: 'user', text: userMessage }];
    setMessages(updatedMessages);
    setChatInput('');
    setChatLoading(true);

    try {
      const res = await fetch('/api/planner', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'modifica_chat',
          comandoUtente: userMessage,
          statoAttuale: {
            interventi: interventiGrezzi.filter(i => !i.isDistinta).map(i => ({ codice: i.codice, nome: i.nome, dataAssegnata: i.dataAssegnata })),
            sospesi: clientiInSospeso.map(s => ({ codice: s.codice, nome: s.nome }))
          },
          history: updatedMessages.map(m => ({
            role: m.sender === 'user' ? 'user' : 'model',
            text: m.text
          }))
        })
      });

      const data = await res.json();
      if (data.success) {
        const rispostaDiscorsiva = data.messaggioChat || "Ho aggiornato la pianificazione secondo le tue indicazioni.";
        setMessages(prev => [...prev, { sender: 'ai', text: rispostaDiscorsiva }]);
        
        if (data.piano && data.piano.giorni) {
          try {
            let tuttaLaLista = [...interventiGrezzi.filter(i => !i.isDistinta), ...clientiInSospeso];
            let nuoviAssegnati: any[] = [];
            let nuoviSospesi: any[] = [];

            tuttaLaLista.forEach(item => {
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
                  nuoviSospesi.push(item);
                } else {
                  if (item.dataAssegnata) {
                    nuoviAssegnati.push(item);
                  } else {
                    item.selezionatoPerGiro = false;
                    nuoviSospesi.push(item);
                  }
                }
              }
            });

            nuoviAssegnati.sort((a, b) => {
               if (a.dataAssegnata !== b.dataAssegnata) return 0;
               return (a._aiIndex || 0) - (b._aiIndex || 0);
            });

            const resRic = await ricalcolaDistinteEOrari(nuoviAssegnati, giornateStats, false);
            setInterventiGrezzi(resRic.listaOrdinata);
            setGiornateStats(resRic.statsAggiornate);
            setClientiInSospeso(nuoviSospesi);
          } catch (recalcErr) {
            console.error("Errore nel ricalcolo locale:", recalcErr);
          }
        }
      } else {
        setMessages(prev => [...prev, { sender: 'ai', text: 'Non sono riuscito ad applicare la modifica. Riprova.' }]);
      }
    } catch (err) {
      setMessages(prev => [...prev, { sender: 'ai', text: 'Errore di connessione con il server.' }]);
    } finally {
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
          const resRic = await ricalcolaDistinteEOrari([...interventiGrezzi, newClient], giornateStats, true);
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
      const resRic = await ricalcolaDistinteEOrari([...interventiGrezzi, newClient], giornateStats, true);
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
            location: `${cliente.indirizzo}, ${cliente.localita || 'Genova'}, Italia`,
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
        html: `<div style="background-color: #2563eb; color: white; border-radius: 50%; width: 26px; height: 26px; display: flex; align-items: center; justify-content: center; font-weight: bold; border: 2px solid white; box-shadow: 0 2px 4px rgba(0,0,0,0.4); font-size: 11px;">${numero}</div>`,
        iconSize: [26, 26],
        iconAnchor: [13, 26],
        popupAnchor: [0, -26]
      });
    }
    return undefined;
  };

  const interventiVisibiliArr = interventiGrezzi.filter(i => giornoSceltoFiltro === "Tutti" || i.dataAssegnata === giornoSceltoFiltro);
  const clientiAttiviCount = interventiVisibiliArr.filter(i => i.selezionatoPerGiro && !i.isDistinta && !i.isPregresso).length;
  const globaleAssegnati = interventiGrezzi.filter(i => i.selezionatoPerGiro && !i.isDistinta && !i.isPregresso);
  const globalPregressiCount = interventiGrezzi.filter(i => i.isPregresso).length;
  const totaleInCarico = globaleAssegnati.length + clientiInSospeso.length + globalPregressiCount;
  
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

  return (
    <div className="flex h-screen w-full bg-slate-100 text-slate-900 font-sans overflow-hidden relative print:h-auto print:overflow-visible print:bg-white">
      <main className="w-full h-full overflow-y-auto custom-scrollbar p-6 md:p-10 print:p-0 print:overflow-visible relative">
        <div className="w-full mx-auto">
          
          <header className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-8 gap-4 print:hidden">
            <div>
              <h1 className="text-2xl md:text-3xl font-extrabold tracking-tight text-slate-900 flex items-center gap-3">
                <span className="p-2.5 bg-blue-600 text-white rounded-xl shadow-md"><Zap size={24} /></span>
                Asso Antincendio <span className="text-blue-600 font-normal text-xl">| Smart Logistics</span>
              </h1>
              <p className="text-xs text-slate-500 mt-1">Rubrica in memoria, Foglio Viaggio, Pausa Pranzo Dinamica & Assistente IA.</p>
            </div>

            <div className="flex flex-col gap-2 w-full md:w-auto">
              <div className="flex items-center gap-2">
                <Link 
                  href="/flussi" 
                  className="px-4 py-2.5 bg-blue-600 text-white rounded-xl text-xs font-bold hover:bg-blue-700 transition flex items-center justify-center gap-2 shadow-sm w-full"
                >
                  📊 Gestione Flussi
                </Link>
              </div>

              <div className="flex flex-col gap-2 p-2.5 bg-slate-50 border border-slate-200 rounded-2xl">
                  <div className="flex items-center gap-2">
                    {isCheckingAuth ? (
                      <div className="text-slate-500 font-bold flex items-center justify-center gap-2 text-sm w-full py-2.5"><Loader2 size={18} className="animate-spin text-blue-600" /> Verifica...</div>
                    ) : !googleToken ? (
                      <button onClick={eseguiLoginGoogle} className="bg-slate-900 hover:bg-slate-800 text-white px-5 py-2.5 rounded-xl font-bold flex items-center justify-center gap-2 shadow-sm text-sm w-full"><Calendar size={18} /> Avvia Connessione Google</button>
                    ) : (
                      <div className="flex items-center gap-2 w-full">
                        <span className="bg-emerald-50 text-emerald-800 px-4 py-2.5 rounded-xl font-bold flex items-center justify-center gap-2 border border-emerald-200 text-sm flex-1"><Check size={18} /> Calendar Connesso</span>
                        <button onClick={logoutGoogle} className="bg-white hover:bg-red-50 text-slate-600 hover:text-red-600 px-3 py-2.5 rounded-xl border border-slate-200 shadow-sm shrink-0"><X size={18} /></button>
                      </div>
                    )}
                  </div>
                  
                  <a 
                    href="https://calendar.google.com" 
                    target="_blank" 
                    rel="noopener noreferrer" 
                    className="bg-blue-100 text-blue-800 px-4 py-2.5 rounded-xl font-bold flex items-center justify-center gap-2 border border-blue-200 shadow-sm text-xs transition-all hover:bg-blue-200"
                  >
                    <Phone size={16} /> Gestione Chiamate <ExternalLink size={14} className="opacity-60" />
                  </a>
              </div>
              
              {/* BOTTONI AFFIANCATI PER RUBRICA E INSOLUTI */}
              <div className="flex gap-2 w-full">
                  <label className={`flex-1 py-2.5 rounded-xl font-bold text-[11px] shadow-sm flex items-center justify-center gap-1.5 transition-all cursor-pointer ${Object.keys(rubricaClienti).length > 0 ? 'bg-amber-100 text-amber-800 hover:bg-amber-200 border border-amber-200' : 'bg-slate-100 text-slate-600 hover:bg-slate-200 border border-slate-300'}`}>
                    <Book size={14} />
                    {Object.keys(rubricaClienti).length > 0 ? `Rubrica (${Object.keys(rubricaClienti).length})` : "Carica Rubrica"}
                    <input type="file" accept=".xlsx, .xls, .csv, .txt" className="hidden" onChange={handleRubricaUpload} />
                  </label>
                  
                  <label className="flex-1 py-2.5 rounded-xl font-bold text-[11px] shadow-sm flex items-center justify-center gap-1.5 transition-all cursor-pointer bg-red-50 text-red-700 hover:bg-red-100 border border-red-200">
                    <Banknote size={14} />
                    Lista Insoluti
                    <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={handleInsolutiUpload} />
                  </label>
              </div>
            </div>
          </header>

          <div className="w-full mb-8 print:hidden">
            <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
               <h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-2 mb-5"><CheckCircle size={18} className="text-emerald-600" /> Riepilogo Pianificazione</h3>
               <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                 <div className="p-4 bg-blue-50 rounded-xl border border-blue-200 flex flex-col justify-center">
                    <span className="block text-[11px] text-blue-700 font-bold uppercase tracking-wider mb-1">Totale In Carico</span>
                    <span className="text-3xl font-extrabold text-blue-900">{totaleInCarico}</span>
                 </div>
                 <div className="p-4 bg-emerald-50 rounded-xl border border-emerald-100 flex flex-col justify-center">
                    <span className="block text-[11px] text-emerald-700 font-bold uppercase tracking-wider mb-1">Nuovi in Rotta</span>
                    <span className="text-3xl font-extrabold text-emerald-900">{clientiAttiviCount}</span>
                 </div>
                 <div className="p-4 bg-amber-50 rounded-xl border border-amber-100 flex flex-col justify-center">
                    <span className="block text-[11px] text-amber-700 font-bold uppercase tracking-wider mb-1">Finiti in Sospeso</span>
                    <span className="text-3xl font-extrabold text-amber-900">{clientiInSospeso.length}</span>
                 </div>
                 <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 flex flex-col justify-center">
                    <span className="block text-[11px] text-slate-500 font-bold uppercase tracking-wider mb-1">Già Fatti / Scartati</span>
                    <span className="text-3xl font-extrabold text-slate-700">{clientiGiaCalendarizzati.length}</span>
                 </div>
               </div>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8 print:hidden">
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
               <div className="w-full flex flex-col gap-3 h-full justify-center">
                   <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center justify-center gap-2 mb-2"><PlayCircle className="text-emerald-600" size={18} /> 3. Genera Plancia</h2>
                   <p className="text-[11px] text-slate-500 mb-2">Lascia che l'Intelligenza Artificiale pianifichi i giri incrociando cantieri, ancore e limiti orari.</p>
                   <button
                     onClick={() => processaRighe(datiGrezziCaricati)}
                     disabled={datiGrezziCaricati.length === 0 || inElaborazione}
                     className={`w-full py-4 rounded-xl text-sm font-extrabold transition-all flex items-center justify-center gap-2 ${(datiGrezziCaricati.length > 0 && !inElaborazione) ? 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-md' : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}
                   >
                     {inElaborazione ? <><Loader2 size={16} className="animate-spin"/> IA in elaborazione...</> : <><Sliders size={16} /> Analizza & Popola la Board</>}
                   </button>
               </div>
            </div>
          </div>

          <div className={`mb-10 relative print:hidden ${inElaborazione ? 'opacity-50 pointer-events-none' : ''}`}>
               {interventiGrezzi.length > 0 && (
                   <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-4 gap-4 bg-white p-4 rounded-2xl border border-slate-200 shadow-sm">
                       <div>
                          <h2 className="text-xl font-extrabold text-slate-900 flex items-center gap-2"><GripVertical className="text-blue-600" /> Plancia Operativa Settimanale</h2>
                          <p className="text-xs text-slate-500">Trascina le schede tra le colonne per affinare la programmazione. O rilascia una scheda su un'altra per posizionarla esattamente lì.</p>
                       </div>
                       <div className="flex gap-4">
                           {bulkSyncStatus.active ? (
                                <div className="bg-blue-50 border border-blue-200 px-4 py-2.5 rounded-xl flex items-center gap-3 text-xs font-bold text-blue-800">
                                    <Loader2 size={14} className="animate-spin"/> Sync a Calendar... {bulkSyncStatus.current} / {bulkSyncStatus.total}
                                </div>
                            ) : (
                                <button onClick={inviaPianificazioneAGoogle} className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2.5 rounded-xl font-bold flex items-center justify-center gap-2 shadow-md text-xs transition-all">
                                    <UploadCloud size={16} /> Salva su Google Calendar ({globaleAssegnati.filter(i=>!i.syncedToGoogle).length})
                                </button>
                           )}
                       </div>
                   </div>
               )}

               {interventiGrezzi.length > 0 && (
                   <div className="flex items-start gap-4 pb-4 w-full">
                       <div 
                           className="min-w-[320px] max-w-[320px] h-[70vh] min-h-[500px] max-h-[800px] bg-slate-200 border-2 border-slate-300 rounded-3xl p-4 flex flex-col shrink-0 z-20 shadow-xl"
                           onDragOver={handleDragOverKanban}
                           onDrop={handleDropToSospesi}
                       >
                           <h3 className="font-extrabold text-slate-800 mb-4 flex justify-between items-center px-1">
                               <span className="flex items-center gap-1.5"><PauseCircle size={16}/> SOSPESI</span>
                               <span className="bg-slate-800 text-white px-2.5 py-0.5 rounded-full text-[10px] shadow-sm">{clientiInSospeso.length}</span>
                           </h3>
                           <div className="flex-1 overflow-y-auto space-y-3 pr-2 custom-scrollbar">
                               {clientiInSospeso.map(item => (
                                  <div 
                                      key={item.codice} 
                                      draggable={!(item.resoEstintori && item.giorniDallUltimoSos < 5)}
                                      onDragStart={(e) => handleDragStartKanban(e, item.codice)}
                                      onClick={() => !(item.resoEstintori && item.giorniDallUltimoSos < 5) && apriModaleSospeso(item)}
                                      className={`bg-white border p-3 rounded-xl shadow-sm flex flex-col gap-1 transition-all ${(item.resoEstintori && item.giorniDallUltimoSos < 5) ? 'border-red-300 bg-red-50/40 opacity-70 cursor-not-allowed' : 'border-amber-200 hover:border-blue-400 hover:shadow-md cursor-grab active:cursor-grabbing relative'}`}
                                  >
                                      <div className="flex justify-between items-start">
                                          <span className="bg-amber-100 text-amber-900 font-mono text-[9px] px-2 py-0.5 rounded font-bold">BI: {item.numeroBi}</span>
                                          <span className="text-[10px] text-slate-500 font-bold flex items-center gap-1"><Timer size={10} className="text-amber-600"/> {item.minutiLavoro}m</span>
                                      </div>
                                      <h4 className="font-bold text-slate-900 text-[11px] leading-tight mt-1 pr-6">{item.nome}</h4>
                                      <p className="text-[10px] text-slate-500 truncate mt-0.5">{item.indirizzo}</p>
                                      <button onClick={(e) => { e.stopPropagation(); setClienteSelezionatoScheda(item); }} className="absolute top-2 right-2 text-slate-400 hover:text-blue-600 bg-slate-50 p-1.5 rounded-lg transition-colors"><FileText size={12}/></button>
                                  </div>
                               ))}
                               {clientiInSospeso.length === 0 && <div className="text-center text-xs text-slate-400 mt-10 font-bold border-2 border-dashed border-slate-300 p-4 rounded-xl">Nessun sospeso. Trascina qui le schede da scartare.</div>}
                           </div>
                       </div>

                       <div className="flex overflow-x-auto items-start gap-4 pb-4 w-full snap-x custom-scrollbar h-[70vh] min-h-[500px] max-h-[800px]">
                           {giornateStats.map(g => {
                               const taskGiorno = interventiGrezzi.filter(i => i.selezionatoPerGiro && i.dataAssegnata === g.dataStr);
                               const isOvertime = g.orarioFineMinuti > 1080; // > 18:00
                               const hh = String(Math.floor(g.orarioFineMinuti / 60)).padStart(2, '0');
                               const mm = String(Math.floor(g.orarioFineMinuti % 60)).padStart(2, '0');

                               return (
                                 <div 
                                     key={g.dataStr}
                                     className={`min-w-[340px] max-w-[340px] h-full bg-slate-50 border-2 rounded-3xl p-4 flex flex-col snap-start shrink-0 transition-all ${isOvertime ? 'border-red-400 shadow-[0_0_20px_rgba(239,68,68,0.15)] bg-red-50/50' : 'border-slate-200 shadow-sm'}`}
                                     onDragOver={handleDragOverKanban}
                                     onDrop={(e) => handleDropToDay(e, g.dataStr)}
                                 >
                                     <div className="mb-4 pb-4 border-b border-slate-200 px-1">
                                         <div className="flex justify-between items-start mb-2">
                                             <h3 className="font-extrabold text-slate-900 text-lg uppercase tracking-tight">{g.nomeGiorno}</h3>
                                             <span className="bg-blue-100 text-blue-800 text-[10px] font-bold px-2 py-1 rounded-lg">{formattaDataVisuale(g.dataStr)}</span>
                                         </div>
                                         
                                         {isOvertime ? (
                                             <div className="bg-red-100 text-red-800 text-[11px] font-bold px-2.5 py-1.5 rounded-lg flex items-center justify-between shadow-sm">
                                                 <span className="flex items-center gap-1.5"><AlertOctagon size={14}/> OVERTIME TURNO:</span>
                                                 <span className="text-sm">{hh}:{mm}</span>
                                             </div>
                                         ) : (
                                             <div className="bg-emerald-100 text-emerald-800 text-[11px] font-bold px-2.5 py-1.5 rounded-lg flex items-center justify-between shadow-sm">
                                                 <span className="flex items-center gap-1.5"><CheckCircle size={14}/> Fine turno stimata:</span>
                                                 <span className="text-sm">{hh}:{mm}</span>
                                             </div>
                                         )}
                                     </div>
                                     
                                     <div className="flex-1 overflow-y-auto space-y-3 pr-2 custom-scrollbar">
                                         {taskGiorno.length === 0 && <div className="text-center text-xs text-slate-400 mt-10 font-bold border-2 border-dashed border-slate-200 p-4 rounded-xl">Giorno vuoto. Trascina qui le schede.</div>}
                                         {taskGiorno.map(item => (
                                            <div 
                                                key={item.codice}
                                                draggable={!item.isPregresso && !item.isDistinta}
                                                onDragStart={(e) => handleDragStartKanban(e, item.codice)}
                                                onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDragOverCard(item.codice); }}
                                                onDragLeave={() => setDragOverCard(null)}
                                                onDrop={(e) => handleDropToDay(e, g.dataStr, item.codice)}
                                                className={`p-3 rounded-xl border bg-white shadow-sm text-left relative transition-all
                                                    ${dragOverCard === item.codice ? 'border-t-4 border-t-blue-500 mt-2' : ''}
                                                    ${item.isPregresso ? 'border-indigo-200 bg-indigo-50/50 cursor-not-allowed' : item.isDistinta ? 'border-orange-200 bg-orange-50/50 cursor-default' : 'border-slate-200 hover:border-blue-400 cursor-grab active:cursor-grabbing'}
                                                `}
                                            >
                                               <div className="flex justify-between items-start mb-1.5">
                                                  {item.isDistinta ? (
                                                    <span className="bg-orange-100 text-orange-800 text-[9px] font-extrabold px-2 py-0.5 rounded-md flex items-center gap-1"><Package size={10}/> MAGAZZINO</span>
                                                  ) : item.isPregresso ? (
                                                    <span className="bg-indigo-100 text-indigo-800 text-[9px] font-extrabold px-2 py-0.5 rounded-md flex items-center gap-1"><CalendarPlus size={10}/> PREGRESSO</span>
                                                  ) : (
                                                    <span className="bg-slate-100 text-slate-700 text-[9px] font-mono font-bold px-2 py-0.5 rounded-md border border-slate-200">BI: {item.numeroBi}</span>
                                                  )}
                                                  {!item.isDistinta && (
                                                      <span className="text-[10px] font-mono font-bold text-slate-600 flex items-center gap-1"><Clock size={10}/> {item.oraInizio}-{item.oraFine}</span>
                                                  )}
                                               </div>
                                               <p className={`font-bold text-[11px] leading-tight mt-1 pr-6 ${item.isDistinta ? 'text-orange-900' : 'text-slate-900'}`}>{item.nome}</p>
                                               <div className="flex justify-between items-end mt-2">
                                                   <p className="text-[9px] text-slate-500 font-medium truncate w-[70%]"><MapPin size={10} className="inline mr-0.5"/>{item.localita || item.indirizzo}</p>
                                                   <div className="flex items-center gap-1.5">
                                                      {item.resoEstintori && <span className="text-amber-600" title="SOS"><AlertTriangle size={12}/></span>}
                                                      {!item.isPregresso && !item.isDistinta && <span className="bg-emerald-50 text-emerald-700 px-1.5 py-0.5 rounded text-[9px] font-bold border border-emerald-100">{item.minutiLavoro}m</span>}
                                                   </div>
                                               </div>
                                               <button onClick={() => setClienteSelezionatoScheda(item)} className="absolute top-2 right-2 text-slate-400 hover:text-blue-600 bg-slate-50 hover:bg-blue-50 p-1.5 rounded-lg transition-colors"><FileText size={12}/></button>
                                            </div>
                                         ))}
                                     </div>
                                 </div>
                               )
                           })}
                       </div>
                   </div>
               )}
          </div>

          {/* MAPPA IN BASSO CON FILTRI ACCORPATI */}
          {interventiGrezzi.length > 0 && leafletLoaded && (
            <div className="w-full flex flex-col gap-6 mb-8 print:hidden">
              <div className={`bg-white p-6 rounded-2xl shadow-sm border border-slate-200 ${inElaborazione ? 'opacity-50 pointer-events-none' : ''}`}>
                
                <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 gap-4 border-b border-slate-100 pb-4">
                  <h3 className="text-lg font-extrabold flex items-center gap-2 text-slate-900"><MapPin className="text-red-500" /> Mappa Globale Cantieri</h3>
                  <div className="flex items-center gap-3">
                    <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider hidden md:block">Filtra per giornata:</span>
                    <button onClick={() => setGiornoSceltoFiltro("Tutti")} className={`text-xs px-4 py-2 rounded-lg font-bold transition-all ${giornoSceltoFiltro === "Tutti" ? "bg-blue-600 text-white shadow-md" : "bg-slate-100 text-slate-600 hover:bg-slate-200 border border-slate-200"}`}>Mostra Tutti</button>
                  </div>
                </div>

                <div className="grid grid-cols-2 md:grid-cols-6 gap-3 mb-6">
                  {giornateStats.map((g, idx) => {
                    const minVal = g.orarioFineMinuti || g.orarioPartenzaMinuti || 480;
                    const isOvertime = minVal > 1080;
                    const hh = String(Math.floor(minVal / 60)).padStart(2, '0');
                    const mm = String(Math.floor(minVal % 60)).padStart(2, '0');
                    
                    return (
                      <div key={idx} onClick={() => setGiornoSceltoFiltro(g.dataStr)} className={`p-3 border rounded-xl flex flex-col gap-0.5 cursor-pointer transition-all ${isOvertime ? 'bg-red-50/50 border-red-200 hover:bg-red-100' : 'bg-emerald-50/50 border-emerald-200 hover:bg-emerald-100'} ${giornoSceltoFiltro === g.dataStr ? 'ring-2 ring-blue-500 shadow-md scale-[1.02]' : 'opacity-80 hover:opacity-100'}`}>
                        <span className="text-[12px] font-bold text-slate-800">{formattaDataVisuale(g.dataStr)}</span>
                        <span className="text-[10px] text-slate-500">{g.nomeGiorno}</span>
                        <div className="flex items-center gap-1.5 mt-1.5 font-mono text-[10px]">
                           {isOvertime ? <AlertOctagon size={12} className="text-red-500" /> : <CheckCircle size={12} className="text-emerald-500" />}
                           <span className={isOvertime ? "text-red-700 font-bold" : "text-emerald-700 font-bold"}>
                              Fine st.: {hh}:{mm}
                           </span>
                        </div>
                      </div>
                    )
                  })}
                </div>

                <div className="h-[480px] w-full rounded-xl overflow-hidden z-0 border border-slate-100 relative">
                  <MapContainer key={giornoSceltoFiltro} center={[SEDE_UFFICIO_LAT, SEDE_UFFICIO_LNG]} zoom={11} style={{ height: "100%", width: "100%" }}>
                    <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
                    {interventiGrezzi.filter(i => (giornoSceltoFiltro === "Tutti" || i.dataAssegnata === giornoSceltoFiltro) && i.lat && i.lng && !i.isDistinta && i.selezionatoPerGiro).map((c, idx) => (
                      <Marker key={c.codice} position={[c.lat, c.lng]} icon={createNumberedIcon(idx + 1)}>
                        <Popup>
                          <div className="p-2">
                            <p className="font-bold text-sm text-slate-900">{c.nome}</p>
                          </div>
                        </Popup>
                      </Marker>
                    ))}
                  </MapContainer>
                </div>
              </div>
            </div>
          )}

          {/* MODALI IN Z-INDEX MASSIMO Z-[9999] */}
          {sospesoInModifica && (
            <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-md z-[9999] flex items-center justify-center p-4">
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

          {clienteSelezionatoScheda && (
            <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-md z-[9999] flex items-center justify-center p-4 overflow-y-auto">
              <div className="bg-white rounded-3xl shadow-2xl max-w-2xl w-full p-8 relative border border-slate-200 my-auto max-h-[90vh] flex flex-col">
                <button onClick={() => setClienteSelezionatoScheda(null)} className="absolute top-6 right-6 text-slate-400 bg-slate-100 p-2.5 rounded-full hover:bg-slate-200"><X size={20} /></button>
                
                {clienteSelezionatoScheda.isDistinta ? (
                  <>
                    <div className="flex items-center gap-4 mb-6">
                      <div className="p-4 bg-orange-100 text-orange-600 rounded-2xl"><Package size={32} /></div>
                      <div>
                        <span className="bg-orange-100 text-orange-800 font-mono text-xs px-3 py-1 rounded-xl font-bold mb-1 inline-block">REPORT MAGAZZINO</span>
                        <h2 className="text-xl font-extrabold text-slate-900">Distinta di Carico</h2>
                      </div>
                    </div>
                    <div className="bg-slate-50 p-6 rounded-2xl border border-slate-200 text-sm whitespace-pre-wrap font-mono leading-relaxed mb-6 overflow-y-auto flex-1">
                      {clienteSelezionatoScheda.descrizioneDistinta}
                    </div>
                  </>
                ) : (
                  <>
                    <div className="flex items-center gap-4 mb-6 pb-4 border-b border-slate-100">
                      <div className="p-4 bg-blue-50 text-blue-600 rounded-2xl"><FileText size={32} /></div>
                      <div>
                        <div className="flex gap-2 mb-1">
                           <span className="bg-blue-100 text-blue-800 font-mono text-xs px-3 py-1 rounded-xl font-bold">BI: {clienteSelezionatoScheda.numeroBi}</span>
                           {clienteSelezionatoScheda.haInsoluto && <span className="bg-red-100 text-red-800 font-mono text-xs px-3 py-1 rounded-xl font-bold">INSOLUTO [INS]</span>}
                           {clienteSelezionatoScheda.resoEstintori && <span className="bg-amber-100 text-amber-800 font-mono text-xs px-3 py-1 rounded-xl font-bold">SOS [ESTINTORI]</span>}
                        </div>
                        <h2 className="text-2xl font-extrabold text-slate-900">{clienteSelezionatoScheda.nome}</h2>
                        <p className="text-xs text-slate-500 font-medium mt-0.5">{clienteSelezionatoScheda.indirizzo}, {clienteSelezionatoScheda.localita}</p>
                      </div>
                    </div>

                    <div className="space-y-4 overflow-y-auto pr-2 flex-1 mb-6">
                       <div className="grid grid-cols-1 md:grid-cols-2 gap-3 bg-slate-50 p-4 rounded-2xl border border-slate-200 text-xs">
                          <div><strong className="text-slate-500 block mb-0.5">Referente / Contatto:</strong> <span className="font-bold text-slate-900">{clienteSelezionatoScheda.contatto || "Nessun referente specificato"}</span></div>
                          <div><strong className="text-slate-500 block mb-0.5">Telefono:</strong> <span className="font-bold text-slate-900">{clienteSelezionatoScheda.telefono || "Nessun telefono"}</span></div>
                          <div className="col-span-2"><strong className="text-slate-500 block mb-0.5">Email:</strong> <span className="font-bold text-slate-900">{clienteSelezionatoScheda.email || "Nessuna email"}</span></div>
                       </div>

                       <div>
                          <h4 className="text-xs font-extrabold text-slate-900 uppercase tracking-wider mb-2">Attrezzature da Verificare ({clienteSelezionatoScheda.totaleArticoli || 0} pz):</h4>
                          {clienteSelezionatoScheda.dettaglioAttrezzature && clienteSelezionatoScheda.dettaglioAttrezzature.length > 0 ? (
                             <div className="border border-slate-200 rounded-xl overflow-hidden divide-y divide-slate-100">
                                {clienteSelezionatoScheda.dettaglioAttrezzature.map((att: any, idx: number) => (
                                   <div key={idx} className="p-3 flex justify-between items-center bg-white text-xs">
                                      <span className="font-bold text-slate-800">{att.descrizione}</span>
                                      <span className="font-mono bg-slate-100 px-2.5 py-1 rounded-lg font-bold text-slate-700">Q.tà: {att.quantita}</span>
                                   </div>
                                ))}
                             </div>
                          ) : (
                             <p className="text-xs text-slate-400 italic bg-slate-50 p-3 rounded-xl border border-slate-200">Nessuna attrezzatura dettagliata trovata per questo cantiere.</p>
                          )}
                       </div>

                       <div className="p-4 bg-emerald-50 rounded-2xl border border-emerald-100 flex justify-between items-center text-xs">
                          <div>
                             <strong className="text-emerald-800 block">Tempo stimato di intervento:</strong>
                             <span className="text-emerald-900 font-extrabold text-sm">{clienteSelezionatoScheda.minutiLavoro} minuti</span>
                          </div>
                          <div className="text-right">
                             <strong className="text-emerald-800 block">Orario Assegnato:</strong>
                             <span className="text-emerald-900 font-mono font-extrabold text-sm">{clienteSelezionatoScheda.oraInizio} - {clienteSelezionatoScheda.oraFine}</span>
                          </div>
                       </div>
                    </div>
                  </>
                )}

                <div className="flex gap-4 pt-4 border-t border-slate-100">
                  <button onClick={() => setClienteSelezionatoScheda(null)} className="flex-1 bg-slate-900 text-white px-6 py-3 rounded-xl text-xs font-bold shadow-md hover:bg-slate-800 transition-all">Chiudi Scheda</button>
                </div>
              </div>
            </div>
          )}

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
                           <pre className="font-mono text-xs whitespace-pre-wrap">{distinta.descrizioneDistinta}</pre>
                        </div>
                      )}

                      <div className="mt-4">
                         <h2 className="font-extrabold text-base mb-3 uppercase border-b border-black pb-1">Elenco Interventi Giornalieri</h2>
                         <table className="w-full text-left text-xs border-collapse border border-black">
                            <thead>
                               <tr className="bg-gray-100 border-b border-black">
                                  <th className="p-2 border border-black w-24">Orario</th>
                                  <th className="p-2 border border-black w-48">Riferimento BI & Cliente</th>
                                  <th className="p-2 border border-black">Indirizzo & Località</th>
                                  <th className="p-2 border border-black">Attrezzature / Note</th>
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
                                     <td className="p-2 border border-black align-top">{lav.indirizzo} {lav.localita ? `- ${lav.localita}` : ""}</td>
                                     <td className="p-2 border border-black align-top text-[11px]">
                                        {lav.dettaglioAttrezzature && lav.dettaglioAttrezzature.length > 0 
                                           ? lav.dettaglioAttrezzature.map((d: any) => `${d.descrizione}: ${d.quantita}`).join(", ") 
                                           : "Verifica standard"}
                                     </td>
                                  </tr>
                               ))}
                            </tbody>
                         </table>
                      </div>
                   </div>
                )
             })}
          </div>

        </div>
      </main>

      {/* PULSANTE FLUTTUANTE PER APRIRE LA CHAT IA */}
      <button
        onClick={() => setIsChatOpen(!isChatOpen)}
        className={`fixed bottom-8 right-8 z-[100] w-16 h-16 rounded-full shadow-[0_10px_40px_rgba(37,99,235,0.4)] flex items-center justify-center transition-all duration-300 hover:scale-105 print:hidden ${
          isChatOpen ? 'bg-slate-800 text-white hover:bg-slate-700 shadow-slate-500/30' : 'bg-gradient-to-br from-blue-600 to-blue-700 text-white'
        }`}
      >
        {isChatOpen ? <X size={28} /> : (
           <div className="relative flex items-center justify-center">
              <MessageSquare size={26} className={chatLoading ? 'animate-pulse' : ''} />
              {!isChatOpen && messages.length > 1 && (
                 <span className="absolute -top-1 -right-2 bg-red-500 text-white text-[10px] font-bold w-5 h-5 flex items-center justify-center rounded-full border-2 border-white shadow-sm">
                    {messages.length - 1}
                 </span>
              )}
           </div>
        )}
      </button>

      {/* FINESTRA CHAT FLUTTUANTE */}
      <div 
        className={`fixed bottom-28 right-8 z-[90] w-[380px] h-[600px] max-h-[75vh] bg-white rounded-3xl shadow-[0_20px_60px_-15px_rgba(0,0,0,0.2)] border border-slate-200 flex flex-col overflow-hidden transition-all duration-300 origin-bottom-right print:hidden ${
          isChatOpen ? 'scale-100 opacity-100 translate-y-0' : 'scale-95 opacity-0 translate-y-8 pointer-events-none'
        }`}
      >
        <div className="p-4 bg-gradient-to-r from-slate-900 to-slate-800 text-white flex justify-between items-center shrink-0">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-blue-500/20 rounded-xl backdrop-blur-sm border border-blue-400/30">
               <Zap size={18} className="text-blue-300" />
            </div>
            <div>
              <h3 className="font-extrabold text-sm tracking-wide">Assistente IA</h3>
              <p className="text-[10px] text-slate-300 font-medium">Operativo e in ascolto</p>
            </div>
          </div>
          <button onClick={() => setIsChatOpen(false)} className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-700/50 rounded-full transition-colors">
             <X size={18} />
          </button>
        </div>

        <div className="flex-1 p-4 overflow-y-auto space-y-4 bg-slate-50/70 custom-scrollbar">
          {messages.map((msg, idx) => (
            <div key={idx} className={`flex ${msg.sender === 'user' ? 'justify-end' : 'justify-start'}`}>
              <div className={`max-w-[85%] px-4 py-3 text-[13px] leading-relaxed shadow-sm ${
                msg.sender === 'user' 
                  ? 'bg-blue-600 text-white rounded-2xl rounded-br-sm font-medium' 
                  : 'bg-white text-slate-700 border border-slate-200/80 rounded-2xl rounded-bl-sm font-medium'
              }`}>
                {msg.text}
              </div>
            </div>
          ))}
          {chatLoading && (
            <div className="flex justify-start">
              <div className="bg-white text-slate-500 border border-slate-200 px-4 py-3 rounded-2xl rounded-bl-sm text-[13px] shadow-sm flex items-center gap-2.5">
                <Loader2 size={14} className="animate-spin text-blue-600" /> Sto elaborando le modifiche...
              </div>
            </div>
          )}
          <div ref={messagesEndRef} />
        </div>

        <form onSubmit={handleSendMessage} className="p-4 bg-white border-t border-slate-100 flex gap-2 items-center shrink-0">
          <input 
            type="text"
            value={chatInput}
            onChange={(e) => setChatInput(e.target.value)}
            placeholder="Es. Sposta Kalmar a mercoledì..."
            className="flex-1 bg-slate-50 border border-slate-200 rounded-2xl px-4 py-3 text-[13px] font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:bg-white transition-all shadow-inner"
          />
          <button 
            type="submit"
            disabled={chatLoading}
            className="bg-blue-600 hover:bg-blue-500 text-white px-4 py-3 rounded-2xl text-[13px] font-bold transition-all shadow-md disabled:opacity-50 shrink-0"
          >
            Invia
          </button>
        </form>
      </div>
    </div>
  );
}