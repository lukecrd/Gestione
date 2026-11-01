import { mensileWeekStart, mondayDates } from './utils/mensileWeek';
import { normalizeMensilePrograms, legacyMensileProgram, mergeMensilePrograms, serializeMensilePrograms, updateMensilePeriod, clearMensilePrograms } from './utils/mensilePrograms';
import { domenicaPeriodKey, normalizeDomenicaPrograms, selectDomenicaProgram, readProgramSelection, legacyDomenicaProgram, mergeDomenicaPrograms, serializeDomenicaPrograms, updateDomenicaPeriod, clearDomenicaPrograms } from './utils/domenicaPrograms';
import { createSchedulingLedger, weeklyDates, weeklyRoleConflicts } from './utils/automaticScheduling';
import { personNameKey, duplicatePersonNames, syncRegistryLinks } from './utils/peopleRegistry';
import { exportProgramExcel } from './utils/programExcel';
import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Users,
  Calendar,
  CalendarDays,
  Settings,
  RefreshCw,
  Download,
  Upload,
  Printer,
  FileSpreadsheet,
  AlertTriangle,
  Trash2,
  Edit3,
  Search,
  Check,
  Sparkles,
  UserPlus,
  Clock,
  Archive,
  Save,
  FolderOpen,
  Eye,
  BookmarkCheck,
  BarChart2,
  RotateCcw,
  ShieldCheck,
  LogOut,
  Lock,
  UserCheck,
  Key,
  MapPin,
  Compass,
  Terminal,
  Cpu,
  Activity,
  Radio,
  BookOpen,
} from 'lucide-react';
import { Person, StateData, MensileRow, DomenicaRow, DomenicaMonthProgram, MensileMonthProgram, ArchivedProgram, AuthUser, ChecklistProgramKey, TabellaAllegato } from './types';
import { StatsView } from './components/StatsView';
import { UtentiView } from './components/UtentiView';
import { ascoltaUtenti, ProfiloUtente } from './lib/utenti';
import { contaInAttesa, puoVedere, puoModificare, chiaviStatoModificabili, PROGRAMMI } from './utils/utenti';
import { OperaPubblicaView } from './components/OperaPubblicaView';
import { ServizioCampoView } from './components/ServizioCampoView';
import { VitaEMinisteroView } from './components/VitaEMinisteroView';
import { TerritoriView } from './components/TerritoriView';
import { DiscorsiEsterniView } from './components/DiscorsiEsterniView';
import { normalizeDiscorsiEsterni } from './utils/discorsiEsterni';
import { normalizeTerritori, alertTerritori } from './utils/territori';
import { ModernSectionHub, SectionKey } from './components/ModernSectionHub';
import { TabellaInformazioniView, VOCI_TABELLA, ZOOM } from './components/TabellaInformazioniView';
import { MensileProgramSheet, MensileFieldKey, fmtShort } from './components/MensileProgramSheet';
import { WeekendProgramSheet } from './components/WeekendProgramSheet';
import { apriPdfProgramma, creaPdf } from './utils/programPdf';
import { ACCOUNT_DRIVE, accediDrive, caricaSuDrive } from './lib/drive';
import { eliminaAllegato, leggiAllegato } from './lib/allegati';
import { ModernSquareNavbar } from './components/ModernSquareNavbar';
import { ConflittiSettimana } from './components/ConflittiSettimana';
import { SECTION_META, accentStyle } from './components/sectionMeta';
import { ExcelImportModal } from './components/ExcelImportModal';
import { DEFAULT_VITA_MINISTERO_DATA } from './data/defaultVitaEMinistero';
import { SCHEMI_DISCORSI_PUBBLICI } from './data/titoliDiscorsiPubblici';
import { auth } from './firebase';
import {
  subscribeToCongregation,
  pushStateToFirestore,
  pushActiveProgramsToFirestore,
  ActiveProgramsData,
} from './lib/firestoreSync';

const STORAGE_KEY = 'dashboard_congregazione_v2';
const MENSILE_PROGRAMS_KEY = 'dashboard_congregazione_mensile_programs_v1';
const DOMENICA_PROGRAMS_KEY = 'dashboard_congregazione_domenica_programs_v1';
const GIORNI = ['Domenica', 'Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato'];
const MESI = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
const OCTOBRE_2026_WEEKEND: Record<string, { presidente?: string; oratore?: string; congregazione?: string; titoloDiscorso: string; lettore?: string; assemblea?: boolean }> = {
  '2026-10-04': { presidente: 'Giuseppe Facchi', oratore: 'Antonio Gramegna', congregazione: 'Gavorrano', titoloDiscorso: 'Impariamo dal diluvio dei giorni di Noè', lettore: 'Michele Verzeletti' },
  '2026-10-11': { presidente: 'Jacopo Pastorelli', oratore: 'Matteo Cannistraci', congregazione: 'Marciana Marina', titoloDiscorso: 'Imitiamo il Padre delle tenere misericordie', lettore: 'Luca Cardelli' },
  '2026-10-18': { presidente: 'Tonino Beniamini', oratore: 'Peter Semprevivo', congregazione: 'Orbetello', titoloDiscorso: 'Chi è qualificato per fare discepoli?', lettore: 'Marco D’Angiolella' },
  '2026-10-25': { titoloDiscorso: 'Assemblea di Circoscrizione con il Rappresentante della filiale', assemblea: true },
};

const _mem: Record<string, string> = {};
const storage = {
  get(k: string) {
    try { return window.localStorage.getItem(k); } catch (e) { return (k in _mem) ? _mem[k] : null; }
  },
  set(k: string, v: string) {
    try { window.localStorage.setItem(k, v); } catch (e) { _mem[k] = v; }
  },
};

function normalize(s: any): StateData {
  s = s || {};
  const people: Person[] = Array.isArray(s.people) ? s.people : [];
  const unavail: Record<string, string[]> = s.unavail || {};
  const special: Record<string, string> = s.special || {};
  const groups = s.groups || { riassetto: '1', pulizie: 'Massa' };
  const mensileArchives: ArchivedProgram[] = Array.isArray(s.mensileArchives) ? s.mensileArchives : [];

  people.forEach((p: any) => {
    p.roles = p.roles || {};
    ['uscieri','console','microfoni','presidente','preghiera','lettore'].forEach(r => {
      p.roles[r] = !!p.roles[r];
    });
  });

  const operaPubblica = s.operaPubblica || { participants: [], schedule: [] };
  const servizioCampo = s.servizioCampo || { conductors: [], schedule: [], locations: [] };
  const vitaEMinistero = s.vitaEMinistero || DEFAULT_VITA_MINISTERO_DATA;
  const programResponsibles = s.programResponsibles && typeof s.programResponsibles === 'object' ? s.programResponsibles : {};
  const territori = normalizeTerritori(s.territori);
  const discorsiEsterni = normalizeDiscorsiEsterni(s.discorsiEsterni);
  const tabellaAllegati = s.tabellaAllegati && typeof s.tabellaAllegati === 'object' ? s.tabellaAllegati : {};

  return syncRegistryLinks({ people, unavail, special, groups, mensileArchives, operaPubblica, servizioCampo, vitaEMinistero, programResponsibles, territori, discorsiEsterni, tabellaAllegati });
}

function loadInitialState(): StateData {
  try {
    const raw = storage.get(STORAGE_KEY);
    if (raw) return normalize(JSON.parse(raw));
  } catch (e) {}
  try {
    const old = storage.get('dashboard_congregazione_v1');
    if (old) return normalize(JSON.parse(old));
  } catch (e) {}
  return normalize({});
}

function uid() {
  return 'p' + Math.random().toString(36).slice(2, 9);
}

function timeNow() {
  const d = new Date();
  return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
}

function iso(dt: Date) {
  return dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0') + '-' + String(dt.getDate()).padStart(2, '0');
}

function fmtDate(dt: Date) {
  return String(dt.getDate()).padStart(2, '0') + '/' + String(dt.getMonth() + 1).padStart(2, '0') + '/' + dt.getFullYear();
}

function addDays(dt: Date, n: number) {
  const d = new Date(dt);
  d.setDate(d.getDate() + n);
  return d;
}

// Righe del programma domenicale come le mostrano la sezione e il PDF: completa Ottobre 2026 dai dati importati
// e, se il mese ha meno di 5 domeniche, aggiunge una domenica vuota in fondo.
function weekendDisplayRows(domRows: DomenicaRow[] | null, domTitle: string) {
  const importedOctoberRows = domRows?.map(row => {
    const sample = domTitle === 'Ottobre 2026' ? OCTOBRE_2026_WEEKEND[iso(row.date)] : undefined;
    if (!sample || row.titoloDiscorso !== undefined) return row;
    return {
      ...row,
      presidente: sample.presidente || '',
      oratore: sample.oratore || '',
      congregazione: sample.congregazione || '',
      titoloDiscorso: sample.titoloDiscorso,
      lettore: sample.lettore || '',
      special: sample.assemblea ? sample.titoloDiscorso : row.special,
    };
  });
  return importedOctoberRows && importedOctoberRows.length > 0 && importedOctoberRows.length < 5
    ? [...importedOctoberRows, { date: addDays(importedOctoberRows[importedOctoberRows.length - 1].date, 7), placeholder: true, oratore: '', congregazione: '', titoloDiscorso: '', presidente: '', lettore: '' }]
    : importedOctoberRows;
}

function sundayDates(year: number, month: number) {
  const dates: Date[] = [];
  const days = new Date(year, month + 1, 0).getDate();
  for (let d = 1; d <= days; d++) {
    const dt = new Date(year, month, d);
    if (dt.getDay() === 0) dates.push(dt);
  }
  return dates;
}

function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function downloadCsv(filename: string, header: string[], rows: (string | undefined)[][]) {
  const esc = (v: any) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  const lines = [header.map(esc).join(','), ...rows.map(r => r.map(esc).join(','))];
  const blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  triggerDownload(blob, filename);
}

// L'utente arriva già approvato da AuthGate: il ruolo viene dal profilo in Firestore, protetto dalle regole.
export default function App({ utente, onEsci }: { utente: AuthUser; onEsci: () => void }) {
  const [state, setState] = useState<StateData>(loadInitialState);
  const [activeTab, setActiveTab] = useState<SectionKey | 'hub'>('hub');

  const currentUser: AuthUser = utente;
  const [dbVuoto, setDbVuoto] = useState(false);
  const handleLogout = onEsci;

  // Elenco utenti solo per gli amministratori (le regole lo negano agli altri).
  const [utenti, setUtenti] = useState<ProfiloUtente[]>([]);
  useEffect(() => {
    if (currentUser.role !== 'admin') return;
    return ascoltaUtenti(setUtenti, err => console.warn('Elenco utenti non disponibile:', err));
  }, [currentUser.role]);

  // Permessi per programma: la sezione aperta decide se si può modificare (le regole di Firestore ricontrollano).
  const canEdit = puoModificare(currentUser, activeTab);
  const checkAdminPermission = (): boolean => {
    if (!canEdit) {
      showToast('Azione non consentita: per questa sezione hai solo la consultazione');
      return false;
    }
    return true;
  };

  // Sync Status
  const [syncMsg, setSyncMsg] = useState<string>('Connessione…');
  const [syncClass, setSyncClass] = useState<string>('text-gray-400');

  // Toast
  const [toastMsg, setToastMsg] = useState<string | null>(null);
  const toastTimer = useRef<any>(null);

  // Custom Confirm Modal State
  interface ConfirmModalConfig {
    title: string;
    message: string;
    confirmText?: string;
    confirmVariant?: 'danger' | 'primary';
    onConfirm: () => void;
  }
  const [confirmModalState, setConfirmModalState] = useState<ConfirmModalConfig | null>(null);

  const showToast = (msg: string) => {
    setToastMsg(msg);
    if (toastTimer.current) clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastMsg(null), 2600);
  };

  // Auto-save & Remote Firestore sync handler
  const saveState = (newState: StateData) => {
    newState = normalize(newState);
    setState(newState);
    storage.set(STORAGE_KEY, JSON.stringify(newState));
    setSyncMsg('Salvataggio Cloud…');
    setSyncClass('text-blue-500 font-medium');
    pushStateToFirestore(newState, currentUser.role === 'admin' ? undefined : chiaviStatoModificabili(currentUser)
      .filter(k => JSON.stringify((newState as any)[k]) !== JSON.stringify((state as any)[k])))
      .then(() => {
        setSyncMsg('✓ Sincronizzato Cloud ' + timeNow());
        setSyncClass('text-emerald-600 font-medium');
      })
      .catch((err) => {
        console.error('Firestore save error:', err);
        setSyncMsg('Offline — Salvato locale');
        setSyncClass('text-amber-600');
      });
  };

  // Helper to sync active programs to Firestore
  const syncActivePrograms = (overrides: Partial<ActiveProgramsData>) => {
    const year = overrides.domYear ?? domYear;
    const month = overrides.domMonth ?? domMonth;
    const key = domenicaPeriodKey(year, month);
    // Include locally migrated periods that do not yet have a cloud entry.
    let updates: Record<string, DomenicaMonthProgram> = {
      ...(Object.fromEntries(Object.entries(domProgramsRef.current).filter(([period]) => !domCloudKeys.current.has(period))) as Record<string, DomenicaMonthProgram>),
      ...(overrides.domPrograms || {}),
    };
    if ('domRows' in overrides || 'domTitle' in overrides || 'domWarn' in overrides) {
      const patch: Partial<DomenicaMonthProgram> = {};
      if ('domRows' in overrides) patch.rows = overrides.domRows;
      if ('domTitle' in overrides) patch.title = overrides.domTitle;
      if ('domWarn' in overrides) patch.warn = overrides.domWarn;
      const updated = updateDomenicaPeriod({ ...domProgramsRef.current, ...updates }, year, month, patch);
      if (key && updated[key]) updates = { ...updates, [key]: updated[key] };
    }
    if (Object.keys(updates).length) persistDomenicaPrograms({ ...domProgramsRef.current, ...updates });
    const menY = overrides.menYear ?? menYear;
    const menM = overrides.menMonth ?? menMonth;
    const menKey = domenicaPeriodKey(menY, menM);
    let menUpdates: Record<string, MensileMonthProgram> = {
      ...(Object.fromEntries(Object.entries(menProgramsRef.current).filter(([period]) => !menCloudKeys.current.has(period))) as Record<string, MensileMonthProgram>),
      ...(overrides.menPrograms || {}),
    };
    if ('menRows' in overrides || 'menTitle' in overrides || 'menWarn' in overrides) {
      const patch: Partial<MensileMonthProgram> = {};
      if ('menRows' in overrides) patch.rows = overrides.menRows;
      if ('menTitle' in overrides) patch.title = overrides.menTitle;
      if ('menWarn' in overrides) patch.warn = overrides.menWarn;
      const updated = updateMensilePeriod({ ...menProgramsRef.current, ...menUpdates }, menY, menM, patch);
      if (menKey && updated[menKey]) menUpdates = { ...menUpdates, [menKey]: updated[menKey] };
    }
    if (Object.keys(menUpdates).length) persistMensilePrograms({ ...menProgramsRef.current, ...menUpdates });
    const data: ActiveProgramsData = {
      menRows,
      menTitle,
      menWarn,
      menMonth,
      menYear,
      domRows,
      domTitle,
      domWarn,
      domMonth,
      domYear,
      ...overrides,
      domPrograms: updates,
      menPrograms: menUpdates,
    };
    pushActiveProgramsToFirestore(data, { men: puoModificare(currentUser, 'mensile'), dom: puoModificare(currentUser, 'domenica') }).catch((err) => {
      console.error('Active program sync error:', err);
      setSyncMsg('Offline — Programmi salvati in locale');
      setSyncClass('text-amber-600');
    });
  };

  useEffect(() => {
    setSyncMsg('Connessione Cloud…');
    setSyncClass('text-blue-500 font-medium');

    const unsubscribeData = subscribeToCongregation(
      ({ state: remoteState, activePrograms: remotePrograms }) => {
        setDbVuoto(false);
        if (remoteState && Array.isArray(remoteState.people)) {
          const norm = normalize(remoteState);
          setState(norm);
          storage.set(STORAGE_KEY, JSON.stringify(norm));
        }

        if (remotePrograms || remoteState) {
          Object.keys(remotePrograms?.menPrograms || {}).forEach(period => menCloudKeys.current.add(period));
          persistMensilePrograms(mergeMensilePrograms(menProgramsRef.current, remotePrograms, remoteState?.mensileArchives || []));
          if (remotePrograms && !hasLoadedMonthlySelection.current) {
            const legacy = legacyMensileProgram(remotePrograms);
            setMenMonth(legacy?.month ?? remotePrograms.menMonth);
            setMenYear(legacy?.year ?? remotePrograms.menYear);
            hasLoadedMonthlySelection.current = true;
          }
        }

        if (remotePrograms) {
          Object.keys(remotePrograms.domPrograms || {}).forEach(period => domCloudKeys.current.add(period));
          persistDomenicaPrograms(mergeDomenicaPrograms(domProgramsRef.current, remotePrograms));
          // Remote writes update saved data, without changing the month being viewed.
          if (!hasLoadedSundaySelection.current) {
            const legacy = legacyDomenicaProgram(remotePrograms);
            setDomMonth(legacy?.month ?? remotePrograms.domMonth);
            setDomYear(legacy?.year ?? remotePrograms.domYear);
            hasLoadedSundaySelection.current = true;
          }
        }

        setSyncMsg('✓ Sincronizzato Cloud ' + timeNow());
        setSyncClass('text-emerald-600 font-medium');
      },
      (error) => {
        console.warn('Firestore subscription error:', error);
        setSyncMsg('Offline — Dati salvati in locale');
        setSyncClass('text-amber-600');
      },
      () => {
        setDbVuoto(true);
        setSyncMsg('Database vuoto');
        setSyncClass('text-amber-600');
      }
    );

    return () => {
      unsubscribeData();
    };
  }, []);

  // Helpers for logic
  const personById = (id: string) => state.people.find(p => p.id === id);
  const areSpouses = (a?: Person | null, b?: Person | null) => {
    if (!a || !b) return false;
    return a.spouseId === b.id || b.spouseId === a.id;
  };

  const isUnavailable = (personId: string, sun: Date, wed: Date) => {
    const list = state.unavail[personId] || [];
    return list.includes(iso(sun)) || list.includes(iso(wed));
  };

  const fairPick = (
    candidates: Person[],
    counts: Record<string, number>,
    lastUsed: Record<string, number>,
    meetingIdx: number,
    excludeIds: string[]
  ) => {
    const pool = candidates.filter(p => !excludeIds.includes(p.id));
    if (pool.length === 0) return null;
    pool.sort((a, b) => {
      const ca = counts[a.id] || 0, cb = counts[b.id] || 0;
      if (ca !== cb) return ca - cb;
      const la = lastUsed[a.id] ?? -999, lb = lastUsed[b.id] ?? -999;
      if (la !== lb) return la - lb;
      return a.name.localeCompare(b.name);
    });
    return pool[0];
  };

  const markUsed = (p: Person, counts: Record<string, number>, lastUsed: Record<string, number>, idx: number) => {
    counts[p.id] = (counts[p.id] || 0) + 1;
    lastUsed[p.id] = idx;
  };

  const consolePairValid = (a: Person, b: Person) => {
    return a.gender === b.gender || areSpouses(a, b);
  };

  const renderPersonName = (name: string | undefined, duplicates?: string[]) => {
    if (!name || name === '—') return '—';
    if (duplicates?.includes(name)) {
      return (
        <span className="inline-flex items-center gap-1 font-semibold text-amber-700 dark:text-amber-400" title="⚠ Attenzione: persona assegnata a più ruoli nella stessa settimana">
          <span className="text-amber-600 dark:text-amber-400 font-bold" aria-label="Avviso">⚠</span>
          {name}
        </span>
      );
    }
    return name;
  };

  // Cella dell'incarico nel Programma mensile: menu a tendina per Admin, testo semplice per Viewer
  const renderMensileCell = (
    rowIndex: number,
    field: MensileFieldKey,
    currentValue: string | undefined,
    pool: Person[],
    duplicates?: string[]
  ) => {
    if (!canEdit) {
      return renderPersonName(currentValue, duplicates);
    }
    const value = currentValue && currentValue !== '—' ? currentValue : '';
    const isDup = !!value && !!duplicates?.includes(value);
    const names = new Set(pool.map(p => p.name));
    if (value) names.add(value);
    const options = Array.from(names).sort((a, b) => a.localeCompare(b));

    return (
      <select
        value={value}
        onChange={e => updateMensileField(rowIndex, field, e.target.value)}
        className={`mensile-select${isDup ? ' mensile-select-warn' : ''}`}
        title={isDup ? '⚠ Attenzione: persona assegnata a più ruoli nella stessa settimana' : 'Cambia nominativo'}
      >
        <option value="">— (nessuno)</option>
        {options.map(name => (
          <option key={name} value={name}>{name}</option>
        ))}
      </select>
    );
  };

  // Opzioni disponibili per Riassetto e Pulizie del fine settimana
  const GROUP_OPTIONS: { value: string; label: string }[] = [
    { value: '1', label: 'Gruppo 1' },
    { value: '2', label: 'Gruppo 2' },
    { value: '3', label: 'Gruppo 3' },
    { value: '4', label: 'Gruppo 4' },
    { value: 'Massa', label: 'Massa Marittima' },
  ];

  // Cambia manualmente il gruppo di Riassetto/Pulizie assegnato a una riga del programma mensile
  const updateMensileGroup = (rowIndex: number, field: 'riassetto' | 'pulizie', value: string) => {
    if (!checkAdminPermission()) return;
    if (!menRows) return;
    const updatedRows = menRows.map((r, i) => {
      if (i !== rowIndex || r.special) return r;
      return { ...r, [field]: value };
    });
    syncActivePrograms({ menRows: updatedRows });
    showToast('Gruppo aggiornato');
  };

  // Cella di Riassetto/Pulizie: select con etichetta estesa a schermo per Admin,
  // testo compatto (es. "2", "Massa") per Viewer e in stampa/PDF
  const renderGroupCell = (rowIndex: number, field: 'riassetto' | 'pulizie', currentValue: string | undefined) => {
    const compact = currentValue || '—';
    if (!canEdit) {
      return compact;
    }
    return (
      <>
        <select
          value={currentValue || ''}
          onChange={e => updateMensileGroup(rowIndex, field, e.target.value)}
          className="mensile-select no-print"
          title="Cambia gruppo"
        >
          {GROUP_OPTIONS.map(opt => (
            <option key={opt.value} value={opt.value}>{opt.label}</option>
          ))}
        </select>
        <span className="print-only">{compact}</span>
      </>
    );
  };

  // --- ANAGRAFICA FORM STATE ---
  const [personId, setPersonId] = useState('');
  const [personName, setPersonName] = useState('');
  const [personGender, setPersonGender] = useState<'M' | 'F'>('M');
  const [personSpouse, setPersonSpouse] = useState('');
  const [roles, setRoles] = useState({
    uscieri: false,
    console: false,
    microfoni: false,
    presidente: false,
    presidentePubblica: false,
    presidenteInfrasettimanale: false,
    preghiera: false,
    lettore: false,
  });
  const [peopleSearch, setPeopleSearch] = useState('');
  const [anagraficaFilter, setAnagraficaFilter] = useState<'all' | 'M' | 'F' | 'uscieri' | 'console' | 'microfoni' | 'presidente' | 'presidenteInfrasettimanale' | 'lettore' | 'preghiera'>('all');
  const [showPersonEditor, setShowPersonEditor] = useState(false);
  const [showExcelImportModal, setShowExcelImportModal] = useState(false);

  const resetPersonForm = () => {
    setPersonId('');
    setPersonName('');
    setPersonGender('M');
    setPersonSpouse('');
    setRoles({
      uscieri: false,
      console: false,
      microfoni: false,
      presidente: false,
    presidentePubblica: false,
    presidenteInfrasettimanale: false,
      preghiera: false,
      lettore: false,
    });
  };

  const handlePersonSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!checkAdminPermission()) return;
    const name = personName.trim().replace(/\s+/g, ' ');
    if (!name) return;
    if (state.people.some(person => person.id !== personId && personNameKey(person.name) === personNameKey(name))) {
      showToast('Questo nominativo è già in anagrafica. Modifica la persona esistente.');
      return;
    }
    const id = personId || uid();
    const newPerson: Person = {
      id,
      name,
      gender: personGender,
      spouseId: personSpouse || null,
      roles: { ...roles },
    };

    const newPeople = [...state.people];
    const idx = newPeople.findIndex(p => p.id === id);
    if (idx >= 0) newPeople[idx] = newPerson;
    else newPeople.push(newPerson);

    // Sync spouses
    newPeople.forEach(p => {
      if (p.id !== newPerson.id && p.spouseId === newPerson.id && newPerson.spouseId !== p.id) {
        p.spouseId = null;
      }
    });
    if (newPerson.spouseId) {
      const sp = newPeople.find(p => p.id === newPerson.spouseId);
      if (sp) sp.spouseId = newPerson.id;
    }

    saveState({ ...state, people: newPeople });
    resetPersonForm();
    setShowPersonEditor(false);
    showToast(idx >= 0 ? 'Persona aggiornata' : 'Persona aggiunta');
  };

  const editPerson = (id: string) => {
    const p = personById(id);
    if (!p) return;
    setPersonId(p.id);
    setPersonName(p.name);
    setPersonGender(p.gender);
    setPersonSpouse(p.spouseId || '');
    setRoles({ ...p.roles, presidentePubblica: !!p.roles.presidentePubblica, presidenteInfrasettimanale: !!p.roles.presidenteInfrasettimanale });
    setShowPersonEditor(true);
  };

  const deletePerson = (id: string) => {
    if (!checkAdminPermission()) return;
    const p = personById(id);
    setConfirmModalState({
      title: 'Elimina Persona',
      message: `Sei sicuro di voler eliminare ${p?.name || 'questa persona'} dall'anagrafica?`,
      confirmText: 'Elimina',
      confirmVariant: 'danger',
      onConfirm: () => {
        const newPeople = state.people.filter(p => p.id !== id);
        newPeople.forEach(p => { if (p.spouseId === id) p.spouseId = null; });
        const newUnavail = { ...state.unavail };
        delete newUnavail[id];
        const newSpecial = { ...state.special };
        delete newSpecial[id];
        saveState({ ...state, people: newPeople, unavail: newUnavail, special: newSpecial });
        showToast('Persona eliminata');
        setConfirmModalState(null);
      }
    });
  };

  // --- IMPORTAZIONE ANAGRAFICA DA EXCEL ---
  const handleImportPeopleFromExcel = (peopleToAdd: Person[], updateExisting: boolean) => {
    if (!checkAdminPermission()) return;
    if (!peopleToAdd.length) return;

    const currentPeople = [...state.people];
    let addedCount = 0;
    let updatedCount = 0;

    peopleToAdd.forEach(importedPerson => {
      const existingIdx = currentPeople.findIndex(
        p => personNameKey(p.name) === personNameKey(importedPerson.name)
      );

      if (existingIdx >= 0) {
        if (updateExisting) {
          currentPeople[existingIdx] = {
            ...currentPeople[existingIdx],
            gender: importedPerson.gender,
          };
          updatedCount++;
        }
      } else {
        currentPeople.push(importedPerson);
        addedCount++;
      }
    });

    saveState({ ...state, people: currentPeople });
    const msg = [
      addedCount > 0 ? `${addedCount} nuovi ${addedCount === 1 ? 'proclamatore aggiunto' : 'proclamatori aggiunti'}` : '',
      updatedCount > 0 ? `${updatedCount} aggiornati` : '',
    ].filter(Boolean).join(', ');
    showToast(msg ? `Importazione completata: ${msg}` : 'Nessuna modifica effettuata');
  };

  // --- CHECKLIST MENSILE: RESPONSABILI PROGRAMMI ---
  const updateProgramResponsible = (key: ChecklistProgramKey, responsible: { name: string; email: string } | null) => {
    if (!checkAdminPermission()) return;
    const current = { ...(state.programResponsibles || {}) };
    const name = responsible?.name.trim() || '';
    const email = responsible?.email.trim() || '';
    if (!responsible || (!name && !email)) {
      delete current[key];
    } else {
      current[key] = { name, email };
    }
    saveState({ ...state, programResponsibles: current });
    showToast(name || email ? 'Responsabile aggiornato' : 'Responsabile rimosso');
  };

  const salvaAllegatoTabella = (chiave: string, allegato: TabellaAllegato | null) => {
    if (currentUser.role !== 'admin') { showToast('Solo un amministratore può collegare link e PDF'); return; }
    const { [chiave]: precedente, ...altri } = state.tabellaAllegati || {};
    saveState({ ...state, tabellaAllegati: allegato ? { ...altri, [chiave]: allegato } : altri });
    // Il PDF sostituito o rimosso non serve più: lo si cancella dopo aver salvato il nuovo riferimento.
    if (precedente?.fileId && precedente.fileId !== allegato?.fileId) {
      eliminaAllegato(precedente.fileId).catch(() => showToast('Collegamento aggiornato, ma il vecchio PDF non è stato eliminato'));
    }
    showToast(allegato ? 'Collegamento salvato' : 'Collegamento rimosso');
  };

  // --- PROGRAMMA MENSILE STATE ---
  const now = new Date();
  const initialMenSelection = readProgramSelection(storage.get(MENSILE_PROGRAMS_KEY + '_selection'), now);
  const [menMonth, setMenMonth] = useState(initialMenSelection.month);
  const [menYear, setMenYear] = useState(initialMenSelection.year);
  const [menPrograms, setMenPrograms] = useState<Record<string, MensileMonthProgram>>(() => {
    try { return mergeMensilePrograms(JSON.parse(storage.get(MENSILE_PROGRAMS_KEY) || '{}'), {}, state.mensileArchives || []); }
    catch { return mergeMensilePrograms({}, {}, state.mensileArchives || []); }
  });
  const menProgramsRef = useRef(menPrograms);
  const menCloudKeys = useRef(new Set<string>());
  const hasLoadedMonthlySelection = useRef(!!storage.get(MENSILE_PROGRAMS_KEY + '_selection'));
  const persistMensilePrograms = (programs: Record<string, MensileMonthProgram>) => {
    const normalized = normalizeMensilePrograms(programs);
    menProgramsRef.current = normalized;
    setMenPrograms(normalized);
    storage.set(MENSILE_PROGRAMS_KEY, JSON.stringify(serializeMensilePrograms(normalized)));
  };
  const changeMensilePeriod = (year: number, month: number) => {
    const key = domenicaPeriodKey(year, month);
    if (!key) return;
    hasLoadedMonthlySelection.current = true;
    setMenYear(year);
    setMenMonth(month);
    setActiveArchiveId(null);
    storage.set(MENSILE_PROGRAMS_KEY + '_selection', key);
  };
  const menProgram = menPrograms[domenicaPeriodKey(menYear, menMonth)];
  const menRows = menProgram?.rows ?? null;
  const menWarn = menProgram?.warn ?? null;
  const menTitle = menProgram?.title || `Mese di ${MESI[menMonth]} ${menYear}`;
  const allMensileRows = (Object.values(menPrograms) as MensileMonthProgram[]).flatMap(program => program.rows || []);

  const generateMensile = (y: number, m: number) => {
    if (!checkAdminPermission()) return;
    const mondays = mondayDates(y, m);
    const ledger = createSchedulingLedger(state, { mensileRows: allMensileRows, domenicaRows: allDomenicaRows }, {
      program: 'mensile', dates: mondays.flatMap(monday => weeklyDates(monday)),
    });
    const uscieriPool = state.people.filter(p => p.roles.uscieri);
    const micPool = state.people.filter(p => p.roles.microfoni);
    const consolePool = state.people.filter(p => p.roles.console);
    const uc: Record<string, number> = {}, ul: Record<string, number> = {};
    const mc: Record<string, number> = {}, ml: Record<string, number> = {};
    const cc: Record<string, number> = {}, cl: Record<string, number> = {};

    const riassGroups = ['1', '2', '3', '4'];
    const pulGroups = ['Massa', '1', '2', '3', '4'];
    let ri = Math.max(0, riassGroups.indexOf(state.groups.riassetto));
    let pi = Math.max(0, pulGroups.indexOf(state.groups.pulizie));

    const warnings: string[] = [];
    if (uscieriPool.length < 3) warnings.push('Servono almeno 3 uscieri.');
    if (micPool.length < 2) warnings.push('Servono almeno 2 microfonisti.');
    if (consolePool.length < 2) warnings.push('Servono almeno 2 persone per Audio/Video (console).');

    const rows: MensileRow[] = [];
    mondays.forEach((monday, idx) => {
      const special = state.special[iso(monday)] || state.special[iso(addDays(monday, 6))];
      if (special) {
        rows.push({ date: monday, special });
        return;
      }

      const dates = weeklyDates(monday);
      const unavailIds = state.people.filter(p => ledger.isUnavailable(p, dates) || ledger.isBusy(p, dates)).map(p => p.id);

      // Presidente of Vita e Ministero on this midweek evening cannot also be usciere.
      const presidenti = new Set(ledger.entries.filter(e => e.program === 'vitaEMinistero' && e.role === 'Presidente' && weeklyDates(e.date)[0] === dates[0]).map(e => e.key));
      const chosenUsc: Person[] = [];
      for (let s = 0; s < 3; s++) {
        const p = fairPick(uscieriPool.filter(p => !ledger.isBusy(p, dates) && !presidenti.has(ledger.identity(p))), uc, ul, idx, [...unavailIds, ...chosenUsc.map(c => c.id)]);
        if (p) { chosenUsc.push(p); markUsed(p, uc, ul, idx); ledger.reserve(p, dates); }
      }

      // Microfonisti: prefer people not already assigned to Uscieri in the same week
      const usedInWeek = [...unavailIds, ...chosenUsc.map(c => c.id)];
      const chosenMic: Person[] = [];
      for (let s = 0; s < 2; s++) {
        const p = fairPick(micPool.filter(p => !ledger.isBusy(p, dates)), mc, ml, idx, [...usedInWeek, ...chosenMic.map(c => c.id)]);
        if (p) { chosenMic.push(p); markUsed(p, mc, ml, idx); ledger.reserve(p, dates); }
      }

      // Audio/Video: prefer people not already assigned to Uscieri or Microfonisti
      const assignedInWeekIds = [...usedInWeek, ...chosenMic.map(c => c.id)];
      let c1: Person | null = null, c2: Person | null = null;

      const orderedPrimary = consolePool.filter(p => !assignedInWeekIds.includes(p.id) && !ledger.isBusy(p, dates)).sort((a, b) => {
        const ca = cc[a.id] || 0, cb = cc[b.id] || 0;
        if (ca !== cb) return ca - cb;
        return (cl[a.id] ?? -999) - (cl[b.id] ?? -999);
      });

      outerPrimary:
      for (let x = 0; x < orderedPrimary.length; x++) {
        for (let y = 0; y < orderedPrimary.length; y++) {
          if (x === y || ledger.identity(orderedPrimary[x]) === ledger.identity(orderedPrimary[y])) continue;
          if (consolePairValid(orderedPrimary[x], orderedPrimary[y])) {
            c1 = orderedPrimary[x];
            c2 = orderedPrimary[y];
            break outerPrimary;
          }
        }
      }

      // A shortage leaves a role empty; do not reuse someone already assigned.
      if (!c1 && orderedPrimary.length > 0) c1 = orderedPrimary[0];
      if (c1) { markUsed(c1, cc, cl, idx); ledger.reserve(c1, dates); }
      if (c2) { markUsed(c2, cc, cl, idx); ledger.reserve(c2, dates); }

      const missing = [];
      if (chosenUsc.length < 3) missing.push(`${3 - chosenUsc.length} uscieri`);
      if (chosenMic.length < 2) missing.push(`${2 - chosenMic.length} microfonisti`);
      if (!c1 || !c2) missing.push(`${2 - Number(!!c1) - Number(!!c2)} audio/video`);
      if (missing.length) warnings.push(`Settimana del ${fmtShort(monday)}: da assegnare ${missing.join(', ')}. Nessun nominativo idoneo libero.`);

      const riassetto = riassGroups[ri % riassGroups.length]; ri++;
      const pulizie = pulGroups[pi % pulGroups.length]; pi++;

      const row: MensileRow = {
        date: monday,
        ingresso: [chosenUsc[0]?.name || '—', chosenUsc[1]?.name || '—'],
        auditorium: chosenUsc[2]?.name || '—',
        microfoni: [chosenMic[0]?.name || '—', chosenMic[1]?.name || '—'],
        audioVideo: [c1?.name || '—', c2?.name || '—'],
        riassetto,
        pulizie,
      };

      // Validation check: duplicate roles in the same week
      const roleAssignments: { name: string; role: string }[] = [];
      if (row.ingresso?.[0] && row.ingresso[0] !== '—') roleAssignments.push({ name: row.ingresso[0], role: 'Uscieri ingresso' });
      if (row.ingresso?.[1] && row.ingresso[1] !== '—') roleAssignments.push({ name: row.ingresso[1], role: 'Uscieri ingresso' });
      if (row.auditorium && row.auditorium !== '—') roleAssignments.push({ name: row.auditorium, role: 'Usciere Auditorium' });
      if (row.microfoni?.[0] && row.microfoni[0] !== '—') roleAssignments.push({ name: row.microfoni[0], role: 'Microfonista' });
      if (row.microfoni?.[1] && row.microfoni[1] !== '—') roleAssignments.push({ name: row.microfoni[1], role: 'Microfonista' });
      if (row.audioVideo?.[0] && row.audioVideo[0] !== '—') roleAssignments.push({ name: row.audioVideo[0], role: 'Audio/Video' });
      if (row.audioVideo?.[1] && row.audioVideo[1] !== '—') roleAssignments.push({ name: row.audioVideo[1], role: 'Audio/Video' });

      const nameMap: Record<string, string[]> = {};
      roleAssignments.forEach(({ name, role }) => {
        if (!nameMap[name]) nameMap[name] = [];
        nameMap[name].push(role);
      });

      const duplicates: string[] = [];
      const rowWarns: string[] = [];
      Object.entries(nameMap).forEach(([name, rolesList]) => {
        if (rolesList.length > 1) {
          duplicates.push(name);
          rowWarns.push(`${name} è assegnato/a a più ruoli (${rolesList.join(', ')}) nella settimana del ${fmtShort(monday)}.`);
        }
      });

      if (duplicates.length > 0) {
        row.duplicates = duplicates;
        row.warnings = rowWarns;
        warnings.push(...rowWarns);
      }

      rows.push(row);
    });

    const title = `Mese di ${MESI[m]} ${y}`;
    const warn = warnings.join(' ') || null;
    setActiveArchiveId(null);
    syncActivePrograms({
      menRows: rows,
      menTitle: title,
      menWarn: warn,
      menMonth: m,
      menYear: y,
    });
  };

  // Ricalcola sovrapposizioni (stessa persona su più incarichi nella stessa settimana)
  const computeRowDuplicates = (row: MensileRow): { duplicates: string[]; warnings: string[] } => {
    if (row.special) return { duplicates: [], warnings: [] };

    const roleAssignments: { name: string; role: string }[] = [];
    if (row.ingresso?.[0] && row.ingresso[0] !== '—') roleAssignments.push({ name: row.ingresso[0], role: 'Uscieri ingresso' });
    if (row.ingresso?.[1] && row.ingresso[1] !== '—') roleAssignments.push({ name: row.ingresso[1], role: 'Uscieri ingresso' });
    if (row.auditorium && row.auditorium !== '—') roleAssignments.push({ name: row.auditorium, role: 'Usciere Auditorium' });
    if (row.microfoni?.[0] && row.microfoni[0] !== '—') roleAssignments.push({ name: row.microfoni[0], role: 'Microfonista' });
    if (row.microfoni?.[1] && row.microfoni[1] !== '—') roleAssignments.push({ name: row.microfoni[1], role: 'Microfonista' });
    if (row.audioVideo?.[0] && row.audioVideo[0] !== '—') roleAssignments.push({ name: row.audioVideo[0], role: 'Audio/Video' });
    if (row.audioVideo?.[1] && row.audioVideo[1] !== '—') roleAssignments.push({ name: row.audioVideo[1], role: 'Audio/Video' });

    const nameMap: Record<string, string[]> = {};
    roleAssignments.forEach(({ name, role }) => {
      if (!nameMap[name]) nameMap[name] = [];
      nameMap[name].push(role);
    });

    const duplicates: string[] = [];
    const warnings: string[] = [];
    Object.entries(nameMap).forEach(([name, rolesList]) => {
      if (rolesList.length > 1) {
        duplicates.push(name);
        warnings.push(`${name} è assegnato/a a più ruoli (${rolesList.join(', ')}) nella settimana del ${fmtShort(mensileWeekStart(row.date))}.`);
      }
    });

    return { duplicates, warnings };
  };


  // Cambia manualmente il nominativo assegnato a un incarico di una riga del programma mensile
  const updateMensileField = (rowIndex: number, field: MensileFieldKey, newName: string) => {
    if (!checkAdminPermission()) return;
    if (!menRows) return;
    const value = newName || '—';

    const updatedRows = menRows.map((r, i) => {
      if (i !== rowIndex || r.special) return r;
      const next: MensileRow = {
        ...r,
        ingresso: r.ingresso ? [...r.ingresso] as [string, string] : undefined,
        microfoni: r.microfoni ? [...r.microfoni] as [string, string] : undefined,
        audioVideo: r.audioVideo ? [...r.audioVideo] as [string, string] : undefined,
      };
      switch (field) {
        case 'ingresso0':
          next.ingresso = [value, next.ingresso?.[1] || '—'];
          break;
        case 'ingresso1':
          next.ingresso = [next.ingresso?.[0] || '—', value];
          break;
        case 'auditorium':
          next.auditorium = value;
          break;
        case 'microfoni0':
          next.microfoni = [value, next.microfoni?.[1] || '—'];
          break;
        case 'microfoni1':
          next.microfoni = [next.microfoni?.[0] || '—', value];
          break;
        case 'audioVideo0':
          next.audioVideo = [value, next.audioVideo?.[1] || '—'];
          break;
        case 'audioVideo1':
          next.audioVideo = [next.audioVideo?.[0] || '—', value];
          break;
      }
      const { duplicates, warnings } = computeRowDuplicates(next);
      next.duplicates = duplicates.length ? duplicates : undefined;
      next.warnings = warnings.length ? warnings : undefined;
      return next;
    });

    const combinedWarn = updatedRows.flatMap(r => r.warnings || []).join(' ') || null;
    syncActivePrograms({ menRows: updatedRows, menWarn: combinedWarn });
    showToast('Incarico aggiornato');
  };

  const [activeArchiveId, setActiveArchiveId] = useState<string | null>(null);

  const saveMensileToArchive = () => {
    if (!menRows) return;
    const nowStr = new Date().toLocaleString('it-IT', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit'
    });
    const archiveId = `men-${menYear}-${menMonth + 1}-${Date.now()}`;
    const newArchive: ArchivedProgram = {
      id: archiveId,
      month: menMonth,
      year: menYear,
      title: menTitle || `Mese di ${MESI[menMonth]} ${menYear}`,
      savedAt: nowStr,
      rows: menRows.map(r => ({
        dateStr: r.date.toISOString(),
        special: r.special,
        ingresso: r.ingresso,
        auditorium: r.auditorium,
        microfoni: r.microfoni,
        audioVideo: r.audioVideo,
        riassetto: r.riassetto,
        pulizie: r.pulizie,
        duplicates: r.duplicates,
        warnings: r.warnings
      })),
      warn: menWarn,
    };

    const existing = state.mensileArchives || [];
    const filtered = existing.filter(a => !(a.month === menMonth && a.year === menYear));
    const updatedArchives = [newArchive, ...filtered];

    saveState({ ...state, mensileArchives: updatedArchives });
    setActiveArchiveId(archiveId);
    showToast(`Programma di ${MESI[menMonth]} ${menYear} archiviato`);
  };

  const exportMensileToExcel = async () => {
try {
    if (!menRows) return;
    const rows: (string | number)[][] = [
      ['Programma Incontro Vita e Ministero / Servizi Infrasettimanali', '', '', '', '', '', ''],
      [`${menTitle}`, '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['Settimana del', 'Uscieri Ingresso', 'Usciere Auditorium', 'Microfonisti', 'Audio / Video', 'Riassetto', 'Pulizie del Fine Settimana'],
    ];

    menRows.forEach(r => {
      if (r.special) {
        rows.push([fmtShort(mensileWeekStart(r.date)), r.special, '', '', '', '', '']);
      } else {
        const ing = (r.ingresso || []).filter(Boolean).join(' — ');
        const mic = (r.microfoni || []).filter(Boolean).join(' — ');
        const av = (r.audioVideo || []).filter(Boolean).join(' — ');
        rows.push([
          fmtShort(mensileWeekStart(r.date)),
          ing || '---',
          r.auditorium || '---',
          mic || '---',
          av || '---',
          r.riassetto || '---',
          r.pulizie || '---',
        ]);
      }
    });

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    worksheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 6 } },
    ];
    worksheet['!cols'] = [
      { wch: 18 },
      { wch: 28 },
      { wch: 22 },
      { wch: 28 },
      { wch: 28 },
      { wch: 22 },
      { wch: 25 },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Programma Infrasettimanale');
    await exportProgramExcel(workbook, `Programma_Infrasettimanale_${menTitle.replace(/\s+/g, '_')}.xlsx`, '.print-sheet');
    showToast('Programma Infrasettimanale esportato in Excel (.xlsx)!');

} catch (error) { showToast(error instanceof Error ? error.message : 'Errore durante l’esportazione Excel.'); }
};

  const exportDomenicaToExcel = async () => {
try {
    if (!domenicaProgramRows) return;
    const rows: (string | number)[][] = [
      ['Programma Adunanza del Fine Settimana', '', '', '', '', ''],
      [`${domTitle}`, '', '', '', '', ''],
      ['', '', '', '', '', ''],
      ['Domenica', 'Presidente', 'Oratore', 'Congregazione', 'Tema Discorso', 'Lettore TdG'],
    ];

    domenicaProgramRows.forEach(r => {
      if (r.special) {
        rows.push([r.placeholder ? 'Domenica' : `${GIORNI[r.date.getDay()]} ${fmtDate(r.date)}`, r.presidente || '', r.oratore || '', r.congregazione || '', r.titoloDiscorso ?? r.special ?? '', r.lettore || '']);
      } else {
        rows.push([
          r.placeholder ? 'Domenica' : `${GIORNI[r.date.getDay()]} ${fmtDate(r.date)}`,
          r.presidente || '---',
          r.oratore || '---',
          r.congregazione || '---',
          r.titoloDiscorso || '---',
          r.lettore || '---',
        ]);
      }
    });

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    worksheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
    ];
    worksheet['!cols'] = [
      { wch: 22 },
      { wch: 24 },
      { wch: 26 },
      { wch: 26 },
      { wch: 55 },
      { wch: 24 },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Adunanza Domenica');
    await exportProgramExcel(workbook, `Programma_Domenica_${domTitle.replace(/\s+/g, '_')}.xlsx`, '.weekend-program');
    showToast('Programma Domenica esportato in Excel (.xlsx)!');

} catch (error) { showToast(error instanceof Error ? error.message : 'Errore durante l’esportazione Excel.'); }
};

  const loadMensileArchive = (archive: ArchivedProgram) => {
    const rows: MensileRow[] = archive.rows.map(r => ({
      date: new Date(r.dateStr),
      special: r.special,
      ingresso: r.ingresso,
      auditorium: r.auditorium,
      microfoni: r.microfoni,
      audioVideo: r.audioVideo,
      riassetto: r.riassetto,
      pulizie: r.pulizie,
      duplicates: r.duplicates,
      warnings: r.warnings
    }));
    changeMensilePeriod(archive.year, archive.month);
    setActiveArchiveId(archive.id);
    showToast(`Caricato dall'archivio: ${archive.title}`);
    syncActivePrograms({
      menRows: rows,
      menTitle: archive.title,
      menWarn: archive.warn || null,
      menMonth: archive.month,
      menYear: archive.year,
    });
  };

  const deleteMensileArchive = (archiveId: string, title: string) => {
    if (!checkAdminPermission()) return;
    setConfirmModalState({
      title: 'Elimina dall\'Archivio',
      message: `Sei sicuro di voler eliminare dall'archivio il programma "${title}"?`,
      confirmText: 'Elimina',
      confirmVariant: 'danger',
      onConfirm: () => {
        const updated = (state.mensileArchives || []).filter(a => a.id !== archiveId);
        saveState({ ...state, mensileArchives: updated });
        if (activeArchiveId === archiveId) {
          setActiveArchiveId(null);
        }
        showToast('Programma rimosso dall\'archivio');
        setConfirmModalState(null);
      }
    });
  };

  const clearAllArchives = () => {
    if (currentUser.role !== 'admin') { showToast('Azione riservata agli amministratori'); return; }
    if (!state.mensileArchives || state.mensileArchives.length === 0) return;
    setConfirmModalState({
      title: 'Svuota Archivio Programmi',
      message: 'Sei sicuro di voler cancellare TUTTI i programmi mensili salvati nell\'archivio?',
      confirmText: 'Svuota Archivio',
      confirmVariant: 'danger',
      onConfirm: () => {
        saveState({ ...state, mensileArchives: [] });
        setActiveArchiveId(null);
        showToast('Archivio programmi svuotato');
        setConfirmModalState(null);
      }
    });
  };

  const clearActiveArchive = () => {
    setActiveArchiveId(null);
  };

  const resetMensileProgram = () => {
    if (!checkAdminPermission()) return;
    if (!menRows) {
      showToast('Nessun programma mensile da azzerare');
      return;
    }
    setConfirmModalState({
      title: 'Azzera Programma Mensile',
      message: 'Sei sicuro di voler azzerare il programma mensile attualmente visualizzato?',
      confirmText: 'Azzera Programma',
      confirmVariant: 'danger',
      onConfirm: () => {
        setActiveArchiveId(null);
        showToast('Programma mensile azzerato');
        setConfirmModalState(null);
        syncActivePrograms({ menRows: null, menWarn: null });
      }
    });
  };

  const resetDomenicaProgram = () => {
    if (!checkAdminPermission()) return;
    if (!domRows) {
      showToast('Nessun turno adunanza domenica da azzerare');
      return;
    }
    setConfirmModalState({
      title: 'Azzera Turni Domenica',
      message: 'Sei sicuro di voler azzerare i turni dell\'Adunanza Domenica?',
      confirmText: 'Azzera Turni',
      confirmVariant: 'danger',
      onConfirm: () => {
        showToast('Turni adunanza domenica azzerati');
        setConfirmModalState(null);
        syncActivePrograms({ domRows: null, domWarn: null });
      }
    });
  };

  const resetAllPrograms = () => {
    if (currentUser.role !== 'admin') { showToast('Azione riservata agli amministratori'); return; }
    if (!menRows && !domRows) {
      showToast('Nessun programma attivo da azzerare');
      return;
    }
    setConfirmModalState({
      title: 'Azzera Tutti i Programmi',
      message: 'Vuoi azzerare sia il Programma Mensile che l\'Adunanza Domenica attualmente visualizzati?',
      confirmText: 'Azzera Tutti i Programmi',
      confirmVariant: 'danger',
      onConfirm: () => {
        setActiveArchiveId(null);
        showToast('Tutti i programmi attivi sono stati azzerati');
        setConfirmModalState(null);
        syncActivePrograms({ menRows: null, menWarn: null, domRows: null, domWarn: null });
      }
    });
  };

  const resetToInitialDefaults = () => {
    if (currentUser.role !== 'admin') { showToast('Azione riservata agli amministratori'); return; }
    setConfirmModalState({
      title: 'Ripristina Dati Iniziali',
      message: 'Vuoi ripristinare tutti i dati della congregazione allo stato iniziale? Verranno ripristinate le persone ed eliminati tutti i programmi salvati.',
      confirmText: 'Ripristina Tutto',
      confirmVariant: 'danger',
      onConfirm: () => {
        localStorage.removeItem(STORAGE_KEY);
        const initial = loadInitialState();
        saveState(initial);
        setActiveArchiveId(null);
        showToast('Dati ripristinati allo stato iniziale');
        setConfirmModalState(null);
        syncActivePrograms({
          menRows: null, menWarn: null, domRows: null, domWarn: null, domTitle: '',
          domPrograms: clearDomenicaPrograms(domProgramsRef.current),
          menPrograms: clearMensilePrograms(menProgramsRef.current),
        });
      }
    });
  };

  // --- ADUNANZA DOMENICA STATE ---
  const initialDomSelection = readProgramSelection(storage.get(DOMENICA_PROGRAMS_KEY + '_selection'), now);
  const [domMonth, setDomMonth] = useState(initialDomSelection.month);
  const [domYear, setDomYear] = useState(initialDomSelection.year);
  const [domPrograms, setDomPrograms] = useState<Record<string, DomenicaMonthProgram>>(() => {
    try { return normalizeDomenicaPrograms(JSON.parse(storage.get(DOMENICA_PROGRAMS_KEY) || '{}')); }
    catch { return {}; }
  });
  const domProgramsRef = useRef(domPrograms);
  const domCloudKeys = useRef(new Set<string>());
  const hasLoadedSundaySelection = useRef(!!storage.get(DOMENICA_PROGRAMS_KEY + '_selection'));
  const persistDomenicaPrograms = (programs: Record<string, DomenicaMonthProgram>) => {
    const normalized = normalizeDomenicaPrograms(programs);
    domProgramsRef.current = normalized;
    setDomPrograms(normalized);
    storage.set(DOMENICA_PROGRAMS_KEY, JSON.stringify(serializeDomenicaPrograms(normalized)));
  };
  const changeDomenicaPeriod = (year: number, month: number) => {
    const key = domenicaPeriodKey(year, month);
    if (!key) return;
    hasLoadedSundaySelection.current = true;
    setDomYear(year);
    setDomMonth(month);
    storage.set(DOMENICA_PROGRAMS_KEY + '_selection', key);
  };
  const domProgram = selectDomenicaProgram(domPrograms, domYear, domMonth);
  const domRows = domProgram?.rows ?? null;
  const domWarn = domProgram?.warn ?? null;
  const domTitle = domProgram?.title || `${MESI[domMonth]} ${domYear}`;
  const domenicaProgramRows = weekendDisplayRows(domRows, domTitle);

  const allDomenicaRows = (Object.values(domPrograms) as DomenicaMonthProgram[]).flatMap(program => program.rows || []);
  // Stesso nominativo in più ruoli nella stessa settimana, su tutti i programmi.
  const conflittiRuoli = activeTab in PROGRAMMI ? weeklyRoleConflicts(state, { mensileRows: allMensileRows, domenicaRows: allDomenicaRows }) : [];

  // --- PDF DEI PROGRAMMI (Tabella informazioni): programma del mese in corso, generato al clic ---
  const foglioPdf = (sezione: SectionKey): React.ReactElement | null => {
    const y = now.getFullYear();
    const m = now.getMonth();
    const delMese = (dateStr?: string) => !!dateStr && dateStr.startsWith(`${y}-${String(m + 1).padStart(2, '0')}`);
    const testo = (value?: string) => (value && value !== '—' ? value : '');
    // Solo lettura: la vista disegnata per il PDF non deve mai salvare.
    const vista = { activePrograms: { mensileRows: allMensileRows, domenicaRows: allDomenicaRows }, onSaveState: () => {}, isAdmin: false, onShowToast: () => {}, checkAdminPermission: () => false, stampa: true };
    switch (sezione) {
      case 'mensile': {
        const programma = menPrograms[domenicaPeriodKey(y, m)];
        if (!programma?.rows?.length) return null;
        return <MensileProgramSheet rows={programma.rows} title={programma.title || `${MESI[m]} ${y}`} warn={null}
          cell={(_i, _f, value) => testo(value) || '—'} group={(_i, _f, value) => value || '—'} />;
      }
      case 'domenica': {
        const programma = selectDomenicaProgram(domPrograms, y, m);
        const title = programma?.title || `${MESI[m]} ${y}`;
        const rows = weekendDisplayRows(programma?.rows ?? null, title);
        if (!rows?.length) return null;
        return <WeekendProgramSheet rows={rows} monthLabel={title.split(' ')[0] || MESI[m]} field={(_i, _k, value) => testo(value)} />;
      }
      case 'vitaEMinistero': {
        const meetings = (state.vitaEMinistero?.meetings || []).filter(mt => delMese(mt.dateStr));
        if (!meetings.length) return null;
        return <VitaEMinisteroView {...vista} state={{ ...state, vitaEMinistero: { ...state.vitaEMinistero!, meetings } }} />;
      }
      case 'servizioCampo':
        return (state.servizioCampo?.schedule || []).some(s => delMese(s.dateStr)) ? <ServizioCampoView {...vista} state={state} /> : null;
      case 'operaPubblica':
        return (state.operaPubblica?.schedule || []).some(s => delMese(s.dateStr)) ? <OperaPubblicaView {...vista} state={state} /> : null;
      default:
        return null;
    }
  };

  // Tabella informazioni in PDF (un foglio A4, a due colonne come il manifesto) salvata sull'account Drive configurato.
  // Ogni riquadro è un link: i PDF dei programmi del mese e quelli caricati vengono messi su Drive come pubblici (con link).
  const salvaTabellaSuDrive = async () => {
    try {
      const token = await accediDrive(); // prima di tutto: il popup di Google deve partire dal clic
      const mese = `${MESI[now.getMonth()]} ${now.getFullYear()}`;
      const nome = `Tabella Informazioni - ${mese}.pdf`;
      const link: Record<string, string> = { zoom: ZOOM.url };
      const perSezione: Partial<Record<SectionKey, string>> = {}; // incarichi settimanali e pulizie: stesso PDF mensile
      for (const voce of VOCI_TABELLA) {
        const allegato = state.tabellaAllegati?.[voce.img];
        if (voce.sezione) {
          if (!(voce.sezione in perSezione)) {
            const foglio = foglioPdf(voce.sezione);
            if (!foglio) continue; // programma del mese non ancora pronto: riquadro senza link
            showToast(`Drive: carico "${voce.titolo}"…`);
            const titolo = `${voce.titolo} - ${mese}`;
            perSezione[voce.sezione] = (await caricaSuDrive(token, await creaPdf(foglio, titolo), `${titolo}.pdf`, { pubblico: true })).webViewLink;
          }
          link[voce.img] = perSezione[voce.sezione]!;
        } else if (allegato?.fileId) {
          showToast(`Drive: carico "${voce.titolo}"…`);
          const file = await leggiAllegato(allegato.fileId);
          link[voce.img] = (await caricaSuDrive(token, file, allegato.nome || `${voce.titolo}.pdf`, { pubblico: true })).webViewLink;
        } else if (allegato?.url) {
          link[voce.img] = allegato.url;
        }
      }
      const tabella = (
        <TabellaInformazioniView onSelectSection={() => {}} visibile={sez => puoVedere(currentUser, sez)} allegati={state.tabellaAllegati || {}}
          isAdmin={false} onSalvaAllegato={() => {}} onErrore={() => {}} onApriPdf={() => Promise.resolve()} />
      );
      const pdf = await creaPdf(tabella, nome, { stampa: false, larghezzaPx: 1000, finestraPx: 1280, unaPagina: true, link });
      await caricaSuDrive(token, pdf, nome);
      showToast(`Tabella salvata su Drive (${ACCOUNT_DRIVE}): ${nome}`);
    } catch (err: any) {
      console.error(err);
      showToast(err?.message || 'Salvataggio su Drive non riuscito');
    }
  };

  const apriPdfSezione = (sezione: SectionKey, titolo: string) => {
    const foglio = foglioPdf(sezione);
    if (!foglio) {
      showToast(`Il programma "${titolo}" di ${MESI[now.getMonth()]} ${now.getFullYear()} non è ancora pronto`);
      return Promise.resolve();
    }
    return apriPdfProgramma(foglio, `${titolo} - ${MESI[now.getMonth()]} ${now.getFullYear()}`)
      .catch(err => { console.error(err); showToast('Non è stato possibile creare il PDF'); });
  };

  const generateDomenica = (y: number, m: number) => {
    if (!checkAdminPermission()) return;
    const sundays = sundayDates(y, m);
    const ledger = createSchedulingLedger(state, { mensileRows: allMensileRows, domenicaRows: allDomenicaRows }, {
      program: 'domenica', dates: sundays.map(iso),
    });
    const presPool = state.people.filter(p => p.roles.presidentePubblica);
    const lettPool = state.people.filter(p => p.roles.lettore);
    const pc: Record<string, number> = {}, pl: Record<string, number> = {};
    const lc: Record<string, number> = {}, ll: Record<string, number> = {};

    const rows: DomenicaRow[] = [];
    const warnings: string[] = [];
    if (presPool.length < 1) warnings.push('Nessun Presidente abilitato.');
    if (lettPool.length < 1) warnings.push('Nessun Lettore Torre di Guardia abilitato.');

    sundays.forEach((sun, idx) => {
      const key = iso(sun);
      const sample = y === 2026 && m === 9 ? OCTOBRE_2026_WEEKEND[key] : undefined;
      if (state.special[key]) {
        rows.push({ date: sun, special: state.special[key], oratore: '', congregazione: '', titoloDiscorso: state.special[key], presidente: '', lettore: '' });
        return;
      }
      if (sample?.assemblea) {
        rows.push({ date: sun, special: sample.titoloDiscorso, oratore: '', congregazione: '', titoloDiscorso: sample.titoloDiscorso, presidente: '', lettore: '' });
        return;
      }
      const existing = domenicaProgramRows?.find(row => !row.placeholder && iso(row.date) === key);
      const oratore = existing?.oratore ?? sample?.oratore ?? '';
      if (oratore) ledger.reserve({ name: oratore }, key);
      const unavailIds = state.people.filter(p => ledger.isUnavailable(p, key) || ledger.isBusy(p, key)).map(p => p.id);
      const used = [...unavailIds];

      const pres = fairPick(presPool, pc, pl, idx, used);
      if (pres) { markUsed(pres, pc, pl, idx); used.push(pres.id); ledger.reserve(pres, key); }

      const lett = fairPick(lettPool.filter(p => !ledger.isBusy(p, key)), lc, ll, idx, used);
      if (lett) { markUsed(lett, lc, ll, idx); used.push(lett.id); ledger.reserve(lett, key); }
      if (!pres) warnings.push(`${fmtDate(sun)}: presidente da assegnare, nessun nominativo idoneo libero.`);
      if (!lett) warnings.push(`${fmtDate(sun)}: lettore da assegnare, nessun nominativo idoneo libero.`);

      const row: DomenicaRow = {
        date: sun,
        oratore,
        congregazione: existing?.congregazione ?? sample?.congregazione ?? '',
        titoloDiscorso: existing?.titoloDiscorso ?? sample?.titoloDiscorso ?? '',
        presidente: pres?.name || '—',
        lettore: lett?.name || '—',
      };

      // Validation check for duplicates
      const roleAssignments: { name: string; role: string }[] = [];
      if (row.presidente && row.presidente !== '—') roleAssignments.push({ name: row.presidente, role: 'Presidente' });
      if (row.lettore && row.lettore !== '—') roleAssignments.push({ name: row.lettore, role: 'Lettore' });

      const nameMap: Record<string, string[]> = {};
      roleAssignments.forEach(({ name, role }) => {
        if (!nameMap[name]) nameMap[name] = [];
        nameMap[name].push(role);
      });

      const duplicates: string[] = [];
      const rowWarns: string[] = [];
      Object.entries(nameMap).forEach(([name, rolesList]) => {
        if (rolesList.length > 1) {
          duplicates.push(name);
          rowWarns.push(`${name} è assegnato/a a più ruoli (${rolesList.join(', ')}) nell'adunanza del ${fmtDate(sun)}.`);
        }
      });

      if (duplicates.length > 0) {
        row.duplicates = duplicates;
        row.warnings = rowWarns;
        warnings.push(...rowWarns);
      }

      rows.push(row);
    });

    if (sundays.length > 0 && sundays.length < 5) {
      rows.push({
        date: addDays(sundays[sundays.length - 1], 7),
        placeholder: true,
        oratore: '',
        congregazione: '',
        titoloDiscorso: '',
        presidente: '',
        lettore: '',
      });
    }

    const title = `${MESI[m]} ${y}`;
    const warn = warnings.join(' ') || null;
    syncActivePrograms({
      domRows: rows,
      domTitle: title,
      domWarn: warn,
      domMonth: m,
      domYear: y,
    });
  };

  const updateDomenicaField = (rowIndex: number, field: 'oratore' | 'congregazione' | 'titoloDiscorso' | 'presidente' | 'lettore', value: string) => {
    if (!checkAdminPermission() || !domenicaProgramRows) return;
    const updatedRows = domenicaProgramRows.map((row, index) => index === rowIndex ? { ...row, [field]: value } : row);
    syncActivePrograms({ domRows: updatedRows });
  };

  const renderDomenicaField = (rowIndex: number, field: 'oratore' | 'congregazione' | 'titoloDiscorso' | 'presidente' | 'lettore', value?: string) => {
    const current = value && value !== '—' ? value : '';
    if (!canEdit) return current;

    if (field === 'oratore' || field === 'congregazione' || field === 'titoloDiscorso') {
      const placeholder = field === 'oratore' ? 'Nome oratore' : field === 'congregazione' ? 'Congregazione' : 'Seleziona o scrivi il tema';
      return <>
        <input
          className={`weekend-program__edit no-print${field === 'titoloDiscorso' ? ' weekend-program__edit--topic' : ''}`}
          value={current}
          placeholder={placeholder}
          list={field === 'titoloDiscorso' ? 'titoli-discorso-pubblico' : undefined}
          onChange={e => updateDomenicaField(rowIndex, field, e.target.value)}
        />
        <span className="weekend-program__value print-only">{current}</span>
      </>;
    }

    const role = field === 'presidente' ? 'presidentePubblica' : 'lettore';
    const eligible = state.people.filter(person => person.roles[role]);
    return <>
      <select className="weekend-program__edit no-print" value={current} onChange={e => updateDomenicaField(rowIndex, field, e.target.value)}>
        <option value="">—</option>
        {current && !eligible.some(person => person.name === current) && <option value={current}>{current} (non abilitato)</option>}
        {eligible.map(person => <option key={person.id} value={person.name}>{person.name}</option>)}
      </select>
      <span className="weekend-program__value print-only">{current}</span>
    </>;
  };

  // --- INDISPONIBILITÀ & DATE SPECIALI ---
  const [unPerson, setUnPerson] = useState('');
  const [unMode, setUnMode] = useState<'single' | 'range'>('single');
  const [unDate, setUnDate] = useState('');
  const [unStartDate, setUnStartDate] = useState('');
  const [unEndDate, setUnEndDate] = useState('');
  const [spDate, setSpDate] = useState('');
  const [spLabel, setSpLabel] = useState('');

  const handleAddUnavail = () => {
    if (!checkAdminPermission()) return;
    if (!unPerson) {
      showToast('Seleziona una persona');
      return;
    }

    if (unMode === 'single') {
      if (!unDate) {
        showToast('Seleziona la data');
        return;
      }
      const list = state.unavail[unPerson] || [];
      if (!list.includes(unDate)) {
        const newUnavail = { ...state.unavail, [unPerson]: [...list, unDate] };
        saveState({ ...state, unavail: newUnavail });
        showToast('Indisponibilità aggiunta');
      } else {
        showToast('Data già presente tra le indisponibilità');
      }
    } else {
      if (!unStartDate || !unEndDate) {
        showToast('Seleziona sia la data di inizio che la data di fine');
        return;
      }
      if (unStartDate > unEndDate) {
        showToast('La data di inizio deve essere precedente o uguale alla data di fine');
        return;
      }

      const datesToAdd: string[] = [];
      const cur = new Date(unStartDate + 'T00:00:00');
      const end = new Date(unEndDate + 'T00:00:00');

      while (cur <= end) {
        const yyyy = cur.getFullYear();
        const mm = String(cur.getMonth() + 1).padStart(2, '0');
        const dd = String(cur.getDate()).padStart(2, '0');
        datesToAdd.push(`${yyyy}-${mm}-${dd}`);
        cur.setDate(cur.getDate() + 1);
      }

      const list = state.unavail[unPerson] || [];
      const newList = Array.from(new Set([...list, ...datesToAdd]));
      const addedCount = newList.length - list.length;

      if (addedCount > 0) {
        const newUnavail = { ...state.unavail, [unPerson]: newList };
        saveState({ ...state, unavail: newUnavail });
        showToast(`Aggiunte ${addedCount} giornate di indisponibilità!`);
      } else {
        showToast('Tutte le date nell\'intervallo erano già state aggiunte');
      }
    }
  };

  const handleRemoveUnavail = (pid: string, d: string) => {
    if (!checkAdminPermission()) return;
    const list = (state.unavail[pid] || []).filter(x => x !== d);
    const newUnavail = { ...state.unavail };
    if (list.length === 0) delete newUnavail[pid];
    else newUnavail[pid] = list;
    saveState({ ...state, unavail: newUnavail });
    showToast('Rimossa');
  };

  const handleAddSpecial = () => {
    if (!checkAdminPermission()) return;
    if (!spDate || !spLabel.trim()) {
      showToast('Inserisci data ed etichetta');
      return;
    }
    const newSpecial = { ...state.special, [spDate]: spLabel.trim() };
    saveState({ ...state, special: newSpecial });
    setSpLabel('');
    showToast('Data speciale aggiunta');
  };

  const handleRemoveSpecial = (d: string) => {
    if (!checkAdminPermission()) return;
    const newSpecial = { ...state.special };
    delete newSpecial[d];
    saveState({ ...state, special: newSpecial });
    showToast('Rimossa');
  };

  // Import / Export JSON
  // Backup completo: lo stato da solo non contiene i programmi mensili e domenicali.
  const handleExportData = () => {
    const backup = {
      formato: 'gestione-congregazione-backup',
      versione: 2,
      esportatoIl: new Date().toISOString(),
      state,
      menPrograms: serializeMensilePrograms(menProgramsRef.current),
      domPrograms: serializeDomenicaPrograms(domProgramsRef.current),
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
    triggerDownload(blob, `backup_congregazione_${new Date().toISOString().slice(0, 10)}.json`);
  };

  // Primo caricamento nel database (vuoto) dei dati salvati in questo browser.
  const caricaDatiLocali = () => {
    if (currentUser.role !== 'admin' || !checkAdminPermission()) return;
    saveState(state);
    syncActivePrograms({ menPrograms: menProgramsRef.current, domPrograms: domProgramsRef.current });
    showToast(`Caricati nel database: ${state.people.length} persone, ${Object.keys(menProgramsRef.current).length} programmi mensili, ${Object.keys(domProgramsRef.current).length} domenicali`);
  };

  const handleImportData = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (currentUser.role !== 'admin') { showToast('Importazione riservata agli amministratori'); return; }
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onerror = () => showToast('Impossibile leggere il file');
    reader.onload = () => {
      let text = String(reader.result || '');
      if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
      text = text.trim();
      let data: any;
      try {
        data = JSON.parse(text);
      } catch (err) {
        showToast('File non è un JSON valido');
        return;
      }
      // Formato nuovo (backup completo) o vecchio (solo lo stato).
      const completo = data?.formato === 'gestione-congregazione-backup';
      const statoImportato = completo ? data.state : data;
      if (!statoImportato || !Array.isArray(statoImportato.people)) {
        showToast('File non valido: manca l\'anagrafica');
        return;
      }
      try {
        const norm = normalize(statoImportato);
        saveState(norm);
        if (completo) {
          const men = normalizeMensilePrograms(data.menPrograms || {});
          const dom = normalizeDomenicaPrograms(data.domPrograms || {});
          persistMensilePrograms(men);
          persistDomenicaPrograms(dom);
          syncActivePrograms({ menPrograms: men, domPrograms: dom });
        }
        showToast('Dati importati: ' + norm.people.length + ' persone' + (completo ? ' e programmi' : ''));
      } catch (err: any) {
        showToast('Errore nell\'applicare i dati: ' + (err?.message || err));
      }
    };
    reader.readAsText(file, 'UTF-8');
    e.target.value = '';
  };

  const filteredPeople = state.people
    .filter(p => {
      const matchSearch = !peopleSearch.trim() || p.name.toLowerCase().includes(peopleSearch.trim().toLowerCase());
      if (!matchSearch) return false;
      if (anagraficaFilter === 'all') return true;
      if (anagraficaFilter === 'M') return p.gender === 'M';
      if (anagraficaFilter === 'F') return p.gender === 'F';
      if (anagraficaFilter === 'uscieri') return p.roles.uscieri;
      if (anagraficaFilter === 'console') return p.roles.console;
      if (anagraficaFilter === 'microfoni') return p.roles.microfoni;
      if (anagraficaFilter === 'presidenteInfrasettimanale') return p.roles.presidenteInfrasettimanale;
      if (anagraficaFilter === 'presidente') return p.roles.presidentePubblica;
      if (anagraficaFilter === 'lettore') return p.roles.lettore;
      if (anagraficaFilter === 'preghiera') return p.roles.preghiera;
      return true;
    })
    .sort((a, b) => a.name.localeCompare(b.name));

  const countMen = state.people.filter(p => p.gender === 'M').length;
  const countWomen = state.people.filter(p => p.gender === 'F').length;
  const countUscieri = state.people.filter(p => p.roles.uscieri).length;
  const countConsole = state.people.filter(p => p.roles.console).length;
  const countMic = state.people.filter(p => p.roles.microfoni).length;
  const registryDuplicates = duplicatePersonNames(state.people);
  const countPres = state.people.filter(p => p.roles.presidentePubblica).length;
  const countLett = state.people.filter(p => p.roles.lettore).length;
  const countPreg = state.people.filter(p => p.roles.preghiera).length;

  const unavailEntries: { pid: string; d: string }[] = [];
  Object.entries(state.unavail).forEach(([pid, dates]) => {
    (dates as string[]).forEach(d => unavailEntries.push({ pid, d }));
  });
  unavailEntries.sort((a, b) => a.d.localeCompare(b.d));

  const specialEntries = Object.entries(state.special).sort((a, b) => a[0].localeCompare(b[0]));

  const activeSectionTitle: Record<SectionKey | 'hub', string> = {
    hub: 'Dashboard',
    tabella: 'Tabella informazioni',
    anagrafica: 'Persone',
    mensile: 'Programma mensile',
    domenica: 'Adunanza domenica',
    vitaEMinistero: 'Vita e ministero',
    servizioCampo: 'Servizio di campo',
    operaPubblica: 'Opera pubblica',
    territori: 'Territori',
    discorsiEsterni: 'Discorsi esterni',
    utenti: 'Utenti',
    impostazioni: 'Assenze e calendario',
    statistiche: 'Statistiche',
  };

  return (
    <div id="app" className="min-h-screen bg-slate-50 font-sans">
      <div className="app-shell" style={accentStyle(activeTab)}>
        <ModernSquareNavbar
          activeTab={activeTab}
          onSelectTab={(tab) => setActiveTab(tab)}
          visibile={sez => puoVedere(currentUser, sez)}
          peopleCount={state.people.length}
          territoriAlert={alertTerritori(normalizeTerritori(state.territori).assegnazioni).length}
          utentiInAttesa={currentUser.role === 'admin' ? contaInAttesa(utenti) : null}
        />

        <div className="app-main">
          <header className="app-topbar no-print">
            <div className="min-w-0 flex items-center gap-3">
              {(() => { const Icon = SECTION_META[activeTab].icon; return <span className="app-topbar-icon" aria-hidden="true"><Icon className="w-5 h-5" /></span>; })()}
              <div className="min-w-0">
                <div className="text-sm text-slate-500">Gestione Congregazione</div>
                <div className="text-lg font-semibold text-slate-950 truncate">{activeSectionTitle[activeTab]}</div>
              </div>
            </div>

            <div className="flex items-center gap-2 sm:gap-3">
              <button
                onClick={() => showToast(syncMsg)}
                className="sync-chip hidden sm:inline-flex"
                title="Stato sincronizzazione"
              >
                <span className={`sync-dot ${syncMsg.toLowerCase().includes('offline') ? 'is-warning' : ''}`} />
                <span className="max-w-[190px] truncate">{syncMsg}</span>
              </button>

              <button onClick={handleExportData} className="icon-button hidden md:inline-flex" title="Esporta dati">
                <Download className="w-4 h-4" />
              </button>
              <label className="icon-button hidden md:inline-flex cursor-pointer" title="Importa dati">
                <Upload className="w-4 h-4" />
                <input type="file" accept="application/json" onChange={handleImportData} className="hidden" />
              </label>

              <div className="user-chip">
                {currentUser.role === 'admin' ? <ShieldCheck className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                <span className="hidden sm:inline">{currentUser.displayName || (currentUser.role === 'admin' ? 'Admin' : 'Consultazione')}</span>
              </div>
              <button onClick={handleLogout} className="icon-button" title="Esci">
                <LogOut className="w-4 h-4" />
              </button>
            </div>
          </header>

          {!canEdit && activeTab !== 'hub' && (
            <div className="viewer-banner no-print">
              <Eye className="w-4 h-4 shrink-0" />
              <span>Sola lettura per questa sezione. Le modifiche richiedono il permesso di un amministratore.</span>
              <button onClick={handleLogout}>Cambia accesso</button>
            </div>
          )}

          <main className="app-content">
      {!puoVedere(currentUser, activeTab) ? (
        <div className="card text-sm text-slate-600" role="status">Non hai accesso a questa sezione. Chiedi a un amministratore.</div>
      ) : <>
      {activeTab in PROGRAMMI && <ConflittiSettimana conflitti={conflittiRuoli} sezione={activeTab} visibile={sez => puoVedere(currentUser, sez)} />}
      {/* ============ HUB PRINCIPALE A QUADRATI MODERNI ============ */}
      {activeTab === 'hub' && (
        <ModernSectionHub
          state={state}
          // La checklist valuta il mese corrente, non il mese aperto negli editor
          menRows={menPrograms[domenicaPeriodKey(now.getFullYear(), now.getMonth())]?.rows ?? null}
          menMonth={now.getMonth()}
          menYear={now.getFullYear()}
          domRows={selectDomenicaProgram(domPrograms, now.getFullYear(), now.getMonth())?.rows ?? null}
          domMonth={now.getMonth()}
          domYear={now.getFullYear()}
          onSelectSection={(sec) => setActiveTab(sec)}
          visibile={sez => puoVedere(currentUser, sez)}
          currentUserRole={currentUser?.role}
          onUpdateResponsible={updateProgramResponsible}
        />
      )}

      {activeTab === 'tabella' && (
        <TabellaInformazioniView
          onSelectSection={(sec) => setActiveTab(sec)}
          visibile={sez => puoVedere(currentUser, sez)}
          allegati={state.tabellaAllegati || {}}
          isAdmin={currentUser.role === 'admin'}
          onSalvaAllegato={salvaAllegatoTabella}
          onErrore={showToast}
          onApriPdf={apriPdfSezione}
          onSalvaDrive={currentUser.role === 'admin' ? salvaTabellaSuDrive : undefined}
        />
      )}

      {/* ============ ANAGRAFICA ============ */}
      {activeTab === 'anagrafica' && (
        <section className="tab-panel no-print space-y-5">
          <div className="page-heading">
            <div>
              <h1 className="page-title">Persone</h1>
              <p className="page-description">Consulta le persone e gestisci le abilitazioni agli incarichi.</p>
            </div>
            {currentUser.role === 'admin' && (
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  className="btn-ghost flex items-center gap-1.5 border border-slate-200 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-200 font-semibold"
                  onClick={() => setShowExcelImportModal(true)}
                  title="Importa anagrafica proclamatori da file Excel (.xlsx, .xls, .csv)"
                >
                  <FileSpreadsheet className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                  <span>Importa da Excel</span>
                </button>
                <button
                  type="button"
                  className="btn-primary"
                  onClick={() => { resetPersonForm(); setShowPersonEditor(true); }}
                >
                  <UserPlus className="w-4 h-4" /> Aggiungi persona
                </button>
              </div>
            )}
          </div>

          {registryDuplicates.length > 0 && <div role="alert" className="status-banner text-amber-700">
            Nominativi ripetuti nell’anagrafica generale: {registryDuplicates.join(', ')}. Verifica i record; nessuna persona viene eliminata automaticamente.
          </div>}
          <p className="text-sm text-slate-500">Anagrafica generale unica. Le liste dei programmi aggiungono abilitazioni e disponibilità ai nominativi collegati.</p>
          {/* Quick statistics */}
          <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2.5">
            <div className="metric-card compact">
              <div className="metric-value">{state.people.length}</div>
              <div className="metric-label">Totale</div>
            </div>
            <div className="metric-card compact">
              <div className="metric-value">{countMen}</div>
              <div className="metric-label">Uomini</div>
            </div>
            <div className="metric-card compact">
              <div className="metric-value">{countWomen}</div>
              <div className="metric-label">Donne</div>
            </div>
            <div className="metric-card compact">
              <div className="metric-value">{countUscieri}</div>
              <div className="metric-label">Uscieri</div>
            </div>
            <div className="metric-card compact">
              <div className="metric-value">{countConsole}</div>
              <div className="metric-label">Audio/Video</div>
            </div>
            <div className="metric-card compact">
              <div className="metric-value">{countMic}</div>
              <div className="metric-label">Microfoni</div>
            </div>
            <div className="metric-card compact">
              <div className="metric-value">{countPres}</div>
              <div className="metric-label">Presidenti</div>
            </div>
            <div className="metric-card compact">
              <div className="metric-value">{countLett}</div>
              <div className="metric-label">Lettori</div>
            </div>
          </div>

          {/* Person editor drawer */}
          {showPersonEditor && (
            <div className="person-editor-layer">
              <button
                type="button"
                className="person-editor-backdrop"
                aria-label="Chiudi modifica persona"
                onClick={() => { resetPersonForm(); setShowPersonEditor(false); }}
              />
              <div className="person-editor-panel">
            <div className="flex items-center justify-between gap-2 mb-4 pb-3 border-b border-slate-200">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-sky-100 dark:bg-sky-950 text-sky-600 dark:text-sky-400 rounded-lg">
                  <UserPlus className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="card-title text-sky-950 dark:text-sky-100">
                    {personId ? 'Modifica persona' : 'Aggiungi persona'}
                  </h2>
                  <p className="text-xs text-sky-700/80 dark:text-sky-400/80">
                    Gestione anagrafica proclamatori e abilitazione incarichi
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                {personId && (
                  <span className="text-xs px-2.5 py-1 rounded-full bg-amber-100 text-amber-800 font-semibold border border-amber-200">
                    Modifica
                  </span>
                )}
                <button
                  type="button"
                  className="icon-button"
                  onClick={() => { resetPersonForm(); setShowPersonEditor(false); }}
                  aria-label="Chiudi"
                >
                  ×
                </button>
              </div>
            </div>
            <form onSubmit={handlePersonSubmit} className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-3">
              <input type="hidden" value={personId} />
              <div className="flex flex-col gap-1">
                <label className="lbl text-sky-900 dark:text-sky-200">Nome e cognome</label>
                <input
                  value={personName}
                  onChange={e => setPersonName(e.target.value)}
                  className="inp focus:border-sky-500 focus:ring-sky-500/20"
                  placeholder="Es. Mario Rossi"
                  required
                />
              </div>
              <div className="flex flex-col gap-1">
                <label className="lbl text-sky-900 dark:text-sky-200">Sesso</label>
                <select
                  value={personGender}
                  onChange={e => setPersonGender(e.target.value as 'M' | 'F')}
                  className="inp focus:border-sky-500"
                >
                  <option value="M">Uomo</option>
                  <option value="F">Donna</option>
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="lbl text-sky-900 dark:text-sky-200">Coniuge (per regola console)</label>
                <select
                  value={personSpouse}
                  onChange={e => setPersonSpouse(e.target.value)}
                  className="inp focus:border-sky-500"
                >
                  <option value="">— Nessuno —</option>
                  {state.people
                    .filter(p => p.id !== personId)
                    .sort((a, b) => a.name.localeCompare(b.name))
                    .map(p => (
                      <option key={p.id} value={p.id}>
                        {p.name}
                      </option>
                    ))}
                </select>
              </div>
              <div className="sm:col-span-2 bg-slate-50 p-3.5 rounded-xl border border-slate-200 mt-1">
                <span className="lbl text-sky-900 dark:text-sky-200 block mb-2 font-bold">Ruoli abilitati:</span>
                <div className="flex flex-wrap items-center gap-x-5 gap-y-2.5">
                  <label className="chk">
                    <input
                      type="checkbox"
                      checked={roles.uscieri}
                      onChange={e => setRoles({ ...roles, uscieri: e.target.checked })}
                      className="accent-sky-600"
                    /> Usciere
                  </label>
                  <label className="chk">
                    <input
                      type="checkbox"
                      checked={roles.console}
                      onChange={e => setRoles({ ...roles, console: e.target.checked })}
                      className="accent-sky-600"
                    /> Audio/Video (console)
                  </label>
                  <label className="chk">
                    <input
                      type="checkbox"
                      checked={roles.microfoni}
                      onChange={e => setRoles({ ...roles, microfoni: e.target.checked })}
                      className="accent-sky-600"
                    /> Microfonista
                  </label>
                  <label className="chk">
                    <input type="checkbox" checked={roles.presidenteInfrasettimanale}
                      onChange={e => setRoles({ ...roles, presidenteInfrasettimanale: e.target.checked })} className="accent-sky-600" />
                    Presidente adunanza infrasettimanale
                  </label>
                  <label className="chk">
                    <input
                      type="checkbox"
                      checked={roles.presidentePubblica}
                      onChange={e => setRoles({ ...roles, presidentePubblica: e.target.checked })}
                      className="accent-sky-600"
                    /> Presidente adunanza pubblica
                  </label>
                  <label className="chk">
                    <input
                      type="checkbox"
                      checked={roles.preghiera}
                      onChange={e => setRoles({ ...roles, preghiera: e.target.checked })}
                      className="accent-sky-600"
                    /> Preghiera finale
                  </label>
                  <label className="chk">
                    <input
                      type="checkbox"
                      checked={roles.lettore}
                      onChange={e => setRoles({ ...roles, lettore: e.target.checked })}
                      className="accent-sky-600"
                    /> Lettore Torre di Guardia
                  </label>
                </div>
              </div>
              <div className="sm:col-span-2 flex gap-2.5 mt-1">
                <button type="submit" className="btn-sky">
                  <Check className="w-4 h-4" />
                  <span>{personId ? 'Aggiorna persona' : 'Salva persona'}</span>
                </button>
                <button
                  type="button"
                  onClick={() => { resetPersonForm(); setShowPersonEditor(false); }}
                  className="btn-ghost"
                >
                  Annulla
                </button>
              </div>
            </form>
              </div>
            </div>
          )}

          {/* Directory List Card */}
          <div className="card border-sky-200 dark:border-sky-900/60 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-3 pb-2 border-b border-sky-100 dark:border-sky-900/60">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-sky-100 dark:bg-sky-950 text-sky-600 dark:text-sky-400 rounded-lg">
                  <Users className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="card-title text-sky-950 dark:text-sky-100">
                    Elenco Anagrafica <span className="text-sky-600 dark:text-sky-400 font-normal ml-1">({filteredPeople.length}{filteredPeople.length !== state.people.length ? ` di ${state.people.length}` : ''})</span>
                  </h2>
                  <p className="text-xs text-slate-500 dark:text-slate-400">
                    Consulta, cerca e filtra i fratelli e le sorelle della congregazione
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2 w-full sm:w-auto">
                <div className="relative flex-1 sm:flex-none">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    value={peopleSearch}
                    onChange={e => setPeopleSearch(e.target.value)}
                    className="inp pl-9 pr-3 py-2 sm:w-[260px]"
                    placeholder="Cerca per nome..."
                  />
                </div>
                {currentUser.role === 'admin' && (
                  <div className="flex items-center gap-1.5 sm:hidden">
                    <button
                      type="button"
                      className="icon-button border border-slate-200 dark:border-slate-700 text-emerald-600 dark:text-emerald-400"
                      onClick={() => setShowExcelImportModal(true)}
                      aria-label="Importa da Excel"
                      title="Importa da Excel"
                    >
                      <FileSpreadsheet className="w-4 h-4" />
                    </button>
                    <button
                      type="button"
                      className="btn-primary"
                      onClick={() => { resetPersonForm(); setShowPersonEditor(true); }}
                      aria-label="Aggiungi persona"
                    >
                      <UserPlus className="w-4 h-4" />
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* Interactive Filter Pills */}
            <div className="flex items-center gap-1.5 flex-wrap mb-3.5 pb-2 border-b border-slate-100 dark:border-slate-800">
              <span className="text-xs font-semibold text-slate-500 dark:text-slate-400 mr-1">Filtra per:</span>
              <button
                onClick={() => setAnagraficaFilter('all')}
                className={`text-xs px-2.5 py-1 rounded-full font-semibold transition-colors ${
                  anagraficaFilter === 'all'
                    ? 'bg-sky-600 text-white shadow-xs'
                    : 'bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300 hover:bg-sky-100 hover:text-sky-800'
                }`}
              >
                Tutti ({state.people.length})
              </button>
              <button
                onClick={() => setAnagraficaFilter('M')}
                className={`text-xs px-2.5 py-1 rounded-full font-semibold transition-colors ${
                  anagraficaFilter === 'M'
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'bg-blue-50 text-blue-800 dark:bg-blue-950/60 dark:text-blue-300 hover:bg-blue-100'
                }`}
              >
                Uomini ({countMen})
              </button>
              <button
                onClick={() => setAnagraficaFilter('F')}
                className={`text-xs px-2.5 py-1 rounded-full font-semibold transition-colors ${
                  anagraficaFilter === 'F'
                    ? 'bg-rose-600 text-white shadow-xs'
                    : 'bg-rose-50 text-rose-800 dark:bg-rose-950/60 dark:text-rose-300 hover:bg-rose-100'
                }`}
              >
                Donne ({countWomen})
              </button>
              <button
                onClick={() => setAnagraficaFilter('uscieri')}
                className={`text-xs px-2.5 py-1 rounded-full font-semibold transition-colors ${
                  anagraficaFilter === 'uscieri'
                    ? 'bg-amber-600 text-white shadow-xs'
                    : 'bg-amber-50 text-amber-800 dark:bg-amber-950/60 dark:text-amber-300 hover:bg-amber-100'
                }`}
              >
                Uscieri ({countUscieri})
              </button>
              <button
                onClick={() => setAnagraficaFilter('console')}
                className={`text-xs px-2.5 py-1 rounded-full font-semibold transition-colors ${
                  anagraficaFilter === 'console'
                    ? 'bg-indigo-600 text-white shadow-xs'
                    : 'bg-indigo-50 text-indigo-800 dark:bg-indigo-950/60 dark:text-indigo-300 hover:bg-indigo-100'
                }`}
              >
                Audio/Video ({countConsole})
              </button>
              <button
                onClick={() => setAnagraficaFilter('microfoni')}
                className={`text-xs px-2.5 py-1 rounded-full font-semibold transition-colors ${
                  anagraficaFilter === 'microfoni'
                    ? 'bg-teal-600 text-white shadow-xs'
                    : 'bg-teal-50 text-teal-800 dark:bg-teal-950/60 dark:text-teal-300 hover:bg-teal-100'
                }`}
              >
                Microfoni ({countMic})
              </button>
              <button onClick={() => setAnagraficaFilter('presidenteInfrasettimanale')} className="text-xs px-2.5 py-1 rounded-full font-semibold bg-indigo-50 text-indigo-800">
                Presidenti infrasettimanale ({state.people.filter(p => p.roles.presidenteInfrasettimanale).length})
              </button>
              <button
                onClick={() => setAnagraficaFilter('presidente')}
                className={`text-xs px-2.5 py-1 rounded-full font-semibold transition-colors ${
                  anagraficaFilter === 'presidente'
                    ? 'bg-purple-600 text-white shadow-xs'
                    : 'bg-purple-50 text-purple-800 dark:bg-purple-950/60 dark:text-purple-300 hover:bg-purple-100'
                }`}
              >
                Presidenti pubblica ({countPres})
              </button>
              <button
                onClick={() => setAnagraficaFilter('lettore')}
                className={`text-xs px-2.5 py-1 rounded-full font-semibold transition-colors ${
                  anagraficaFilter === 'lettore'
                    ? 'bg-emerald-600 text-white shadow-xs'
                    : 'bg-emerald-50 text-emerald-800 dark:bg-emerald-950/60 dark:text-emerald-300 hover:bg-emerald-100'
                }`}
              >
                Lettori ({countLett})
              </button>
            </div>

            <div className="table-wrapper border border-sky-200 dark:border-sky-900/60 rounded-xl">
              <table className="tbl">
                <thead>
                  <tr className="bg-sky-100/70 dark:bg-sky-950/60 text-sky-950 dark:text-sky-200">
                    <th className="font-bold">Nome</th>
                    <th className="font-bold">Sesso</th>
                    <th className="font-bold">Coniuge</th>
                    <th className="text-center font-bold">Usc.</th>
                    <th className="text-center font-bold">A/V</th>
                    <th className="text-center font-bold">Micr.</th>
                    <th className="text-center font-bold" title="Presidente adunanza infrasettimanale">Pres. infr.</th>
                    <th className="text-center font-bold" title="Presidente adunanza pubblica">Pres. pub.</th>
                    <th className="text-center font-bold">Pregh.</th>
                    <th className="text-center font-bold">Lett.</th>
                    <th className="text-right font-bold">Azioni</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredPeople.map(p => {
                    const spouse = p.spouseId ? (personById(p.spouseId)?.name || '—') : '—';
                    const chk = (v: boolean) => (
                      v ? <Check className="w-4 h-4 text-sky-600 dark:text-sky-400 mx-auto stroke-[2.5]" /> : <span className="text-slate-300 dark:text-slate-700">—</span>
                    );
                    return (
                      <tr key={p.id} className="hover:bg-sky-50/40 dark:hover:bg-sky-950/30 transition-colors">
                        <td className="font-semibold text-slate-900 dark:text-slate-100">{p.name}</td>
                        <td>
                          <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold ${
                            p.gender === 'M'
                              ? 'bg-blue-100 text-blue-800 border border-blue-200 dark:bg-blue-950/60 dark:text-blue-300 dark:border-blue-800'
                              : 'bg-rose-100 text-rose-800 border border-rose-200 dark:bg-rose-950/60 dark:text-rose-300 dark:border-rose-800'
                          }`}>
                            {p.gender === 'M' ? 'Uomo' : 'Donna'}
                          </span>
                        </td>
                        <td className="text-slate-500 dark:text-slate-400">{spouse}</td>
                        <td className="text-center">{chk(p.roles.uscieri)}</td>
                        <td className="text-center">{chk(p.roles.console)}</td>
                        <td className="text-center">{chk(p.roles.microfoni)}</td>
                        <td className="text-center">{chk(!!p.roles.presidenteInfrasettimanale)}</td>
                        <td className="text-center">{chk(!!p.roles.presidentePubblica)}</td>
                        <td className="text-center">{chk(p.roles.preghiera)}</td>
                        <td className="text-center">{chk(p.roles.lettore)}</td>
                        <td className="text-right whitespace-nowrap">
                          <button
                            onClick={() => editPerson(p.id)}
                            className="bg-sky-50 hover:bg-sky-100 text-sky-700 border border-sky-200 dark:bg-sky-950/50 dark:hover:bg-sky-900/60 dark:text-sky-300 dark:border-sky-800/60 rounded-lg px-2.5 py-1 text-xs font-semibold inline-flex items-center gap-1 transition-colors mr-1"
                            title="Modifica"
                          >
                            <Edit3 className="w-3.5 h-3.5" />
                            <span>Modifica</span>
                          </button>
                          <button
                            onClick={() => deletePerson(p.id)}
                            className="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 dark:bg-rose-950/50 dark:hover:bg-rose-900/60 dark:text-rose-300 dark:border-rose-800/60 rounded-lg px-2.5 py-1 text-xs font-semibold inline-flex items-center gap-1 transition-colors"
                            title="Elimina"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            <span>Elimina</span>
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {filteredPeople.length === 0 && (
              <div className="text-center py-10 space-y-3">
                <p className="text-slate-400">
                  {peopleSearch || anagraficaFilter !== 'all'
                    ? 'Nessuna persona corrisponde ai criteri di ricerca/filtro.'
                    : 'Nessuna persona in anagrafica.'}
                </p>
                {currentUser.role === 'admin' && !peopleSearch && anagraficaFilter === 'all' && (
                  <div className="flex items-center justify-center gap-2.5 pt-2">
                    <button
                      type="button"
                      className="btn-ghost text-xs border border-slate-200 dark:border-slate-700 inline-flex items-center gap-1.5"
                      onClick={() => setShowExcelImportModal(true)}
                    >
                      <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                      Importa da file Excel
                    </button>
                    <button
                      type="button"
                      className="btn-primary text-xs inline-flex items-center gap-1.5"
                      onClick={() => { resetPersonForm(); setShowPersonEditor(true); }}
                    >
                      <UserPlus className="w-4 h-4" /> Aggiungi persona
                    </button>
                  </div>
                )}
              </div>
            )}
          </div>
        </section>
      )}

      {/* ============ PROGRAMMA MENSILE ============ */}
      {activeTab === 'mensile' && (
        <section className="tab-panel space-y-5">
          <div className="page-heading no-print">
            <div>
              <h1 className="page-title">Programma mensile</h1>
              <p className="page-description">Genera e gestisci i turni di sala per il mese selezionato.</p>
            </div>
          </div>
          <div className="card no-print">
            <div className="flex items-center gap-2 mb-1 pb-2 border-b border-slate-100 dark:border-slate-800">
              <Sparkles className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
              <h2 className="card-title">Genera Programma Mensile</h2>
            </div>
            <p className="text-slate-500 dark:text-slate-400 text-xs mt-1">
              Un unico programma per settimana: Uscieri (2 ingresso + 1 auditorium), Microfonisti (2), Audio/Video console (2, mai uomo+donna salvo coniugi), Riassetto (gruppi 1‑4) e Pulizie (gruppi 1‑4 + Massa).
            </p>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5 mt-4">
              <div className="flex flex-col gap-1">
                <label className="lbl" htmlFor="men-month">Mese</label>
                <select
                  id="men-month"
                  value={menMonth}
                  onChange={e => changeMensilePeriod(menYear, parseInt(e.target.value, 10))}
                  className="inp"
                >
                  {MESI.map((m, i) => (
                    <option key={i} value={i}>{m}</option>
                  ))}
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="lbl" htmlFor="men-year">Anno</label>
                <input
                  id="men-year"
                  type="number"
                  value={menYear}
                  onChange={e => changeMensilePeriod(parseInt(e.target.value, 10), menMonth)}
                  className="inp"
                />
              </div>
              <div className="flex items-end">
                <button
                  onClick={() => generateMensile(menYear, menMonth)}
                  className="btn-primary w-full"
                >
                  <Sparkles className="w-4 h-4" />
                  <span>Genera programma</span>
                </button>
              </div>
            </div>
            {menRows && (
              <div className="flex items-center gap-2 mt-4 pt-3 border-t border-slate-100 dark:border-slate-800 flex-wrap">
                <button
                  onClick={saveMensileToArchive}
                  className="btn-primary"
                  title="Salva o aggiorna questo programma nell'archivio permanente"
                >
                  <Save className="w-4 h-4" />
                  <span>{activeArchiveId ? 'Aggiorna in Archivio' : 'Salva in Archivio'}</span>
                </button>
                <button onClick={() => window.print()} className="btn-ghost" title="Stampa o salva in PDF">
                  <Printer className="w-4 h-4 text-indigo-600 dark:text-indigo-400" />
                  <span>Stampa / PDF</span>
                </button>
                <button
                  onClick={exportMensileToExcel}
                  className="btn-ghost text-emerald-700 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-950/30 hover:bg-emerald-100"
                  title="Esporta in formato Excel (.xlsx)"
                >
                  <FileSpreadsheet className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                  <span>Esporta Excel (.xlsx)</span>
                </button>
                <button
                  onClick={() => {
                    if (!menRows) return;
                    const header = ['Settimana del', 'Uscieri ingresso 1', 'Uscieri ingresso 2', 'Usciere Auditorium', 'Microfonista 1', 'Microfonista 2', 'Audio/Video 1', 'Audio/Video 2', 'Riassetto', 'Pulizie'];
                    const data = menRows.map(r => r.special
                      ? [fmtShort(mensileWeekStart(r.date)), r.special, '', '', '', '', '', '', '', '']
                      : [fmtShort(mensileWeekStart(r.date)), r.ingresso?.[0], r.ingresso?.[1], r.auditorium, r.microfoni?.[0], r.microfoni?.[1], r.audioVideo?.[0], r.audioVideo?.[1], r.riassetto, r.pulizie]);
                    downloadCsv('programma_mensile_' + menTitle.replace(/\s+/g, '_') + '.csv', header, data);
                  }}
                  className="btn-ghost text-slate-500 text-xs"
                  title="Esporta CSV grezzo"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>CSV</span>
                </button>
                <button
                  onClick={resetMensileProgram}
                  className="btn-ghost text-rose-600 dark:text-rose-400 hover:bg-rose-50 dark:hover:bg-rose-950/40"
                  title="Azzera e svuota il programma mensile attualmente visualizzato"
                >
                  <RotateCcw className="w-4 h-4" />
                  <span>Azzera Programma</span>
                </button>
              </div>
            )}
          </div>

          {!menRows && (
            <div className="card no-print text-sm text-slate-500" role="status">
              Nessun programma salvato per {MESI[menMonth]} {menYear}.
            </div>
          )}
          {/* ARCHIVIO PROGRAMMI MENSILE */}
          <div className="card mb-5 no-print">
            <div className="flex items-center justify-between gap-2 mb-2 pb-2 border-b border-slate-100 dark:border-slate-800">
              <div className="flex items-center gap-2">
                <Archive className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                <h2 className="card-title">Archivio Programmi Mensili</h2>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-semibold px-2.5 py-0.5 rounded-full bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400">
                  {(state.mensileArchives || []).length} { (state.mensileArchives || []).length === 1 ? 'salvato' : 'salvati' }
                </span>
                {(state.mensileArchives || []).length > 0 && (
                  <button
                    onClick={clearAllArchives}
                    className="btn-ghost text-xs text-rose-600 dark:text-rose-400 py-1 px-2 hover:bg-rose-50 dark:hover:bg-rose-950/40"
                    title="Svuota l'intero archivio dei programmi salvati"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Svuota Archivio</span>
                  </button>
                )}
              </div>
            </div>

            {(!state.mensileArchives || state.mensileArchives.length === 0) ? (
              <p className="text-xs text-slate-400 dark:text-slate-500 py-3 text-center">
                I programmi vengono salvati automaticamente per mese e anno. Puoi conservare anche una copia nell'archivio.
              </p>
            ) : (
              <div className="table-wrapper border border-slate-200 dark:border-slate-800 rounded-xl mt-3">
                <table className="tbl">
                  <thead>
                    <tr>
                      <th>Periodo</th>
                      <th>Data salvataggio</th>
                      <th className="text-center">Settimane</th>
                      <th className="text-right">Azioni</th>
                    </tr>
                  </thead>
                  <tbody>
                    {state.mensileArchives.map(a => {
                      const isViewing = activeArchiveId === a.id;
                      return (
                        <tr key={a.id} className={isViewing ? "bg-indigo-50/70 dark:bg-indigo-950/30 font-medium" : undefined}>
                          <td className="font-semibold text-slate-900 dark:text-slate-100">
                            <div className="flex items-center gap-2">
                              <span>{a.title}</span>
                              {isViewing && (
                                <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-indigo-100 dark:bg-indigo-900 text-indigo-700 dark:text-indigo-300">
                                  In visione
                                </span>
                              )}
                            </div>
                          </td>
                          <td className="text-xs text-slate-500 dark:text-slate-400">{a.savedAt}</td>
                          <td className="text-center text-xs text-slate-600 dark:text-slate-400">{a.rows.length}</td>
                          <td className="text-right whitespace-nowrap">
                            <button
                              onClick={() => loadMensileArchive(a)}
                              className={`icon-btn mr-1.5 ${isViewing ? 'bg-indigo-600 text-white hover:bg-indigo-700' : 'edit'}`}
                              title="Consulta programma"
                            >
                              <Eye className="w-3.5 h-3.5" />
                              <span>{isViewing ? 'In visione' : 'Consulta'}</span>
                            </button>
                            <button
                              onClick={() => deleteMensileArchive(a.id, a.title)}
                              className="icon-btn"
                              title="Elimina dall'archivio"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                              <span>Elimina</span>
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          {activeArchiveId && (
            <div className="bg-indigo-50 dark:bg-indigo-950/40 border border-indigo-200 dark:border-indigo-800/80 rounded-xl p-3.5 mb-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 no-print shadow-sm">
              <div className="flex items-center gap-2 text-indigo-900 dark:text-indigo-200 font-medium text-xs sm:text-sm">
                <Archive className="w-4 h-4 text-indigo-600 dark:text-indigo-400 shrink-0" />
                <span>
                  Stai consultando la versione archiviata di: <strong>{menTitle}</strong> (salvata il {state.mensileArchives?.find(a => a.id === activeArchiveId)?.savedAt}).
                </span>
              </div>
              <button
                onClick={clearActiveArchive}
                className="btn-ghost text-xs whitespace-nowrap bg-white dark:bg-slate-900 border border-indigo-200 dark:border-indigo-800 text-indigo-700 dark:text-indigo-300 hover:bg-indigo-100"
              >
                Torna a generazione libera
              </button>
            </div>
          )}

          {menRows && (() => {
            const uscieriOptions = state.people.filter(p => p.roles.uscieri).sort((a, b) => a.name.localeCompare(b.name));
            const microfoniOptions = state.people.filter(p => p.roles.microfoni).sort((a, b) => a.name.localeCompare(b.name));
            const audioVideoOptions = state.people.filter(p => p.roles.console).sort((a, b) => a.name.localeCompare(b.name));
            const poolFor = (field: MensileFieldKey) =>
              field.startsWith('microfoni') ? microfoniOptions : field.startsWith('audioVideo') ? audioVideoOptions : uscieriOptions;
            return (
              <MensileProgramSheet
                rows={menRows}
                title={menTitle}
                warn={menWarn}
                cell={(i, field, value, duplicates) => renderMensileCell(i, field, value, poolFor(field), duplicates)}
                group={renderGroupCell}
                hint={canEdit && (
                  <p className="text-xs text-slate-500 dark:text-slate-400 mb-2 no-print">
                    Clicca su un nominativo per assegnare un'altra persona all'incarico.
                  </p>
                )}
              />
            );
          })()}
        </section>
      )}

      {/* ============ ADUNANZA DOMENICA ============ */}
      {activeTab === 'domenica' && (
        <section className="tab-panel space-y-5">
          <div className="page-heading no-print">
            <div>
              <h1 className="page-title">Adunanza domenica</h1>
              <p className="page-description">Pianifica oratore, titolo del discorso pubblico, presidente e lettore per le domeniche del mese.</p>
            </div>
          </div>
          <div className="card no-print">
            <div className="flex items-center gap-2 mb-1 pb-2 border-b border-slate-100 dark:border-slate-800">
              <CalendarDays className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
              <h2 className="card-title">Genera programma Adunanza Domenica</h2>
            </div>
            <p className="text-slate-500 dark:text-slate-400 text-xs mt-1">
              Per ogni domenica: oratore e discorso pubblico, Presidente e Lettore Torre di Guardia. Rotazione equa; una persona non ricopre due ruoli nella stessa adunanza. Rispetta indisponibilità e date speciali.
            </p>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5 mt-4">
              <div className="flex flex-col gap-1">
                <label className="lbl" htmlFor="dom-month">Mese</label>
                <select
                  id="dom-month"
                  value={domMonth}
                  onChange={e => changeDomenicaPeriod(domYear, parseInt(e.target.value, 10))}
                  className="inp"
                >
                  {MESI.map((m, i) => (
                    <option key={i} value={i}>{m}</option>
                  ))}
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="lbl" htmlFor="dom-year">Anno</label>
                <input
                  id="dom-year"
                  type="number"
                  value={domYear}
                  onChange={e => changeDomenicaPeriod(parseInt(e.target.value, 10), domMonth)}
                  className="inp"
                />
              </div>
              <div className="flex items-end">
                <button
                  onClick={() => generateDomenica(domYear, domMonth)}
                  className="btn-primary w-full"
                >
                  <Sparkles className="w-4 h-4" />
                  <span>Genera turni</span>
                </button>
              </div>
            </div>
          </div>

          {!domenicaProgramRows && (
            <div className="card no-print text-sm text-slate-500" role="status">
              Nessun programma salvato per {MESI[domMonth]} {domYear}.
            </div>
          )}
          {domenicaProgramRows && (
            <WeekendProgramSheet
              rows={domenicaProgramRows}
              monthLabel={domTitle.split(' ')[0] || MESI[domMonth]}
              field={renderDomenicaField}
              toolbar={
                <div className="no-print flex flex-wrap justify-end gap-2 p-3">
                  <button onClick={() => window.print()} className="btn-ghost" title="Apri la stampa del programma">
                    <Printer className="w-4 h-4 text-indigo-600 dark:text-indigo-400" />
                    <span>Stampa</span>
                  </button>
                  <button onClick={() => window.print()} className="btn-ghost" title="Scegli Salva come PDF nella finestra di stampa">
                    <Download className="w-4 h-4 text-indigo-600 dark:text-indigo-400" />
                    <span>Esporta PDF</span>
                  </button>
                  <button
                    onClick={exportDomenicaToExcel}
                    className="btn-ghost text-emerald-700 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-950/30 hover:bg-emerald-100"
                    title="Esporta in formato Excel (.xlsx)"
                  >
                    <FileSpreadsheet className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                    <span>Esporta Excel (.xlsx)</span>
                  </button>
                </div>
              }
              warn={domWarn && <div className="warn no-print"><AlertTriangle className="w-4 h-4 shrink-0 text-amber-600 dark:text-amber-400" /> <span>{domWarn}</span></div>}
            >
              <datalist id="titoli-discorso-pubblico">
                {SCHEMI_DISCORSI_PUBBLICI.map(s => <option key={s.numero} value={s.titolo} label={`n. ${s.numero}`} />)}
              </datalist>
            </WeekendProgramSheet>
          )}
        </section>
      )}

      {/* ============ INDISPONIBILITÀ & DATE SPECIALI ============ */}
      {activeTab === 'impostazioni' && (
        <section className="tab-panel no-print space-y-5">
          <div className="page-heading">
            <div>
              <h1 className="page-title">Assenze e calendario</h1>
              <p className="page-description">Registra indisponibilità e date speciali che influenzano la generazione dei programmi.</p>
            </div>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            <div className="card">
              <div className="flex items-center gap-2 mb-1 pb-2 border-b border-slate-100 dark:border-slate-800">
                <Clock className="w-5 h-5 text-amber-600 dark:text-amber-400" />
                <h2 className="card-title">Indisponibilità</h2>
              </div>
              <p className="text-slate-500 dark:text-slate-400 text-xs mt-1">
                Segna le date o un intervallo di più settimane in cui una persona non è disponibile.
              </p>

              {/* Mode Switcher */}
              <div className="flex items-center gap-3 mt-3 bg-slate-100 dark:bg-slate-800/60 p-1.5 rounded-xl text-xs">
                <button
                  type="button"
                  onClick={() => setUnMode('single')}
                  className={`flex-1 py-1.5 px-3 rounded-lg font-medium transition-all ${
                    unMode === 'single'
                      ? 'bg-white dark:bg-slate-900 text-slate-900 dark:text-white shadow-sm'
                      : 'text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
                  }`}
                >
                  📅 Data Singola
                </button>
                <button
                  type="button"
                  onClick={() => setUnMode('range')}
                  className={`flex-1 py-1.5 px-3 rounded-lg font-medium transition-all ${
                    unMode === 'range'
                      ? 'bg-white dark:bg-slate-900 text-slate-900 dark:text-white shadow-sm'
                      : 'text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
                  }`}
                >
                  🗓️ Intervallo (Più settimane)
                </button>
              </div>

              <div className="flex flex-col gap-3 mt-4">
                <div className="flex flex-col gap-1">
                  <label className="lbl">Persona</label>
                  <select
                    value={unPerson}
                    onChange={e => setUnPerson(e.target.value)}
                    className="inp text-xs"
                  >
                    <option value="">— Seleziona persona —</option>
                    {state.people
                      .slice()
                      .sort((a, b) => a.name.localeCompare(b.name))
                      .map(p => (
                        <option key={p.id} value={p.id}>{p.name}</option>
                      ))}
                  </select>
                </div>

                {unMode === 'single' ? (
                  <div className="flex flex-col gap-1">
                    <label className="lbl">Data</label>
                    <input
                      type="date"
                      value={unDate}
                      onChange={e => setUnDate(e.target.value)}
                      className="inp text-xs"
                    />
                  </div>
                ) : (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div className="flex flex-col gap-1">
                      <label className="lbl">Data Inizio Range</label>
                      <input
                        type="date"
                        value={unStartDate}
                        onChange={e => setUnStartDate(e.target.value)}
                        className="inp text-xs"
                      />
                    </div>
                    <div className="flex flex-col gap-1">
                      <label className="lbl">Data Fine Range</label>
                      <input
                        type="date"
                        value={unEndDate}
                        onChange={e => setUnEndDate(e.target.value)}
                        className="inp text-xs"
                      />
                    </div>
                  </div>
                )}
              </div>

              <button onClick={handleAddUnavail} className="btn-primary mt-4 w-full justify-center">
                <Clock className="w-4 h-4" />
                <span>
                  {unMode === 'single' ? 'Aggiungi Indisponibilità Singola' : 'Aggiungi Intervallo Indisponibilità'}
                </span>
              </button>
              <div className="table-wrapper border border-slate-200 dark:border-slate-800 rounded-xl mt-4">
                <table className="tbl">
                  <thead>
                    <tr>
                      <th>Persona</th>
                      <th>Data</th>
                      <th className="text-right">Azioni</th>
                    </tr>
                  </thead>
                  <tbody>
                    {unavailEntries.map(({ pid, d }) => {
                      const name = personById(pid)?.name || '(eliminato)';
                      const [y, mo, da] = d.split('-');
                      return (
                        <tr key={`${pid}-${d}`}>
                          <td className="font-medium text-slate-900 dark:text-slate-100">{name}</td>
                          <td className="text-slate-600 dark:text-slate-300">{`${da}/${mo}/${y}`}</td>
                          <td className="text-right">
                            <button
                              onClick={() => handleRemoveUnavail(pid, d)}
                              className="icon-btn"
                              title="Rimuovi"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                              <span>Rimuovi</span>
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {unavailEntries.length === 0 && (
                <p className="text-slate-400 text-center py-6 text-xs">Nessuna indisponibilità registrata.</p>
              )}
            </div>

            <div className="card">
              <div className="flex items-center gap-2 mb-1 pb-2 border-b border-slate-100 dark:border-slate-800">
                <Calendar className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                <h2 className="card-title">Date speciali</h2>
              </div>
              <p className="text-slate-500 dark:text-slate-400 text-xs mt-1">
                Settimane senza adunanza normale (assemblea, commemorazione, ecc.). La riga mostrerà l'etichetta al posto degli incarichi.
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5 mt-4">
                <div className="flex flex-col gap-1">
                  <label className="lbl">Data (domenica della settimana)</label>
                  <input
                    type="date"
                    value={spDate}
                    onChange={e => setSpDate(e.target.value)}
                    className="inp"
                  />
                </div>
                <div className="flex flex-col gap-1">
                  <label className="lbl">Etichetta</label>
                  <input
                    value={spLabel}
                    onChange={e => setSpLabel(e.target.value)}
                    className="inp"
                    placeholder="Es. Assemblea di circoscrizione"
                  />
                </div>
              </div>
              <button onClick={handleAddSpecial} className="btn-primary mt-4">
                <Calendar className="w-4 h-4" />
                <span>Aggiungi data speciale</span>
              </button>
              <div className="table-wrapper border border-slate-200 dark:border-slate-800 rounded-xl mt-4">
                <table className="tbl">
                  <thead>
                    <tr>
                      <th>Data</th>
                      <th>Etichetta</th>
                      <th className="text-right">Azioni</th>
                    </tr>
                  </thead>
                  <tbody>
                    {specialEntries.map(([d, label]) => {
                      const [y, mo, da] = d.split('-');
                      return (
                        <tr key={d}>
                          <td className="font-medium text-slate-900 dark:text-slate-100">{`${da}/${mo}/${y}`}</td>
                          <td className="text-slate-600 dark:text-slate-300">{label}</td>
                          <td className="text-right">
                            <button
                              onClick={() => handleRemoveSpecial(d)}
                              className="icon-btn"
                              title="Rimuovi"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                              <span>Rimuovi</span>
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {specialEntries.length === 0 && (
                <p className="text-slate-400 text-center py-6 text-xs">Nessuna data speciale registrata.</p>
              )}
            </div>

            <div className="card md:col-span-2">
              <div className="flex items-center gap-2 mb-1 pb-2 border-b border-slate-100 dark:border-slate-800">
                <Settings className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                <h2 className="card-title">Rotazione gruppi</h2>
              </div>
              <p className="text-slate-500 dark:text-slate-400 text-xs mt-1">
                Imposta da quale gruppo partire ogni mese; la rotazione avanza di una posizione ad ogni settimana.
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4">
                <div className="flex flex-col gap-1">
                  <label className="lbl">Riassetto — gruppo di partenza</label>
                  <select
                    value={state.groups.riassetto}
                    onChange={e => checkAdminPermission() && saveState({ ...state, groups: { ...state.groups, riassetto: e.target.value } })}
                    disabled={!canEdit}
                    className="inp"
                  >
                    <option value="1">Gruppo 1</option>
                    <option value="2">Gruppo 2</option>
                    <option value="3">Gruppo 3</option>
                    <option value="4">Gruppo 4</option>
                  </select>
                </div>
                <div className="flex flex-col gap-1">
                  <label className="lbl">Pulizie — gruppo di partenza</label>
                  <select
                    value={state.groups.pulizie}
                    onChange={e => checkAdminPermission() && saveState({ ...state, groups: { ...state.groups, pulizie: e.target.value } })}
                    disabled={!canEdit}
                    className="inp"
                  >
                    <option value="Massa">Massa</option>
                    <option value="1">Gruppo 1</option>
                    <option value="2">Gruppo 2</option>
                    <option value="3">Gruppo 3</option>
                    <option value="4">Gruppo 4</option>
                  </select>
                </div>
              </div>
            </div>

            <div className="card md:col-span-2">
              <div className="flex items-center gap-2 mb-1 pb-2 border-b border-slate-100 dark:border-slate-800 text-rose-600 dark:text-rose-400">
                <RotateCcw className="w-5 h-5" />
                <h2 className="card-title text-slate-900 dark:text-slate-100">Gestione & Azzeramento Dati</h2>
              </div>
              <p className="text-slate-500 dark:text-slate-400 text-xs mt-1">
                Azzera o ripristina i programmi generati, l'archivio salvato o l'intera anagrafica della congregazione.
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5 mt-4">
                <button
                  onClick={resetAllPrograms}
                  className="btn-ghost justify-start text-rose-600 dark:text-rose-400 border border-slate-200 dark:border-slate-800 p-3 hover:bg-rose-50 dark:hover:bg-rose-950/30"
                >
                  <RotateCcw className="w-4 h-4 shrink-0" />
                  <div className="text-left">
                    <div className="font-semibold text-xs">Azzera Programmi Attivi</div>
                    <div className="text-[11px] text-slate-500 dark:text-slate-400 font-normal">Svuota la tabella mensile e domenica in visione</div>
                  </div>
                </button>

                <button
                  onClick={clearAllArchives}
                  className="btn-ghost justify-start text-rose-600 dark:text-rose-400 border border-slate-200 dark:border-slate-800 p-3 hover:bg-rose-50 dark:hover:bg-rose-950/30"
                >
                  <Trash2 className="w-4 h-4 shrink-0" />
                  <div className="text-left">
                    <div className="font-semibold text-xs">Svuota Archivio Programmi</div>
                    <div className="text-[11px] text-slate-500 dark:text-slate-400 font-normal">Cancella tutti i programmi mensili salvati</div>
                  </div>
                </button>

                <button
                  onClick={resetToInitialDefaults}
                  className="btn-ghost justify-start text-rose-700 dark:text-rose-400 border border-rose-200 dark:border-rose-900/60 p-3 hover:bg-rose-100 dark:hover:bg-rose-950/50"
                >
                  <RotateCcw className="w-4 h-4 shrink-0" />
                  <div className="text-left">
                    <div className="font-semibold text-xs">Ripristina Dati Iniziali</div>
                    <div className="text-[11px] text-slate-500 dark:text-slate-400 font-normal">Ripristina la congregazione di esempio predefinita</div>
                  </div>
                </button>
              </div>
            </div>
          </div>
        </section>
      )}

      {/* ============ VITA E MINISTERO ============ */}
      {activeTab === 'vitaEMinistero' && (
        <VitaEMinisteroView
          state={state}
          activePrograms={{ mensileRows: allMensileRows, domenicaRows: allDomenicaRows }}
          onSaveState={saveState}
          isAdmin={canEdit}
          onShowToast={showToast}
          checkAdminPermission={checkAdminPermission}
        />
      )}

      {/* ============ SERVIZIO DI CAMPO ============ */}
      {activeTab === 'servizioCampo' && (
        <ServizioCampoView
          state={state}
          activePrograms={{ mensileRows: allMensileRows, domenicaRows: allDomenicaRows }}
          onSaveState={saveState}
          isAdmin={canEdit}
          onShowToast={showToast}
          checkAdminPermission={checkAdminPermission}
        />
      )}

      {/* ============ OPERA PUBBLICA ============ */}
      {activeTab === 'operaPubblica' && (
        <OperaPubblicaView
          state={state}
          activePrograms={{ mensileRows: allMensileRows, domenicaRows: allDomenicaRows }}
          onSaveState={saveState}
          isAdmin={canEdit}
          onShowToast={showToast}
          checkAdminPermission={checkAdminPermission}
        />
      )}

      {dbVuoto && currentUser.role === 'admin' && (
        <div className="status-banner mb-6 no-print" role="alert">
          <AlertTriangle className="w-5 h-5 shrink-0" />
          <div className="flex-1">
            <div className="font-semibold text-slate-900">Il database è vuoto</div>
            <div className="text-sm text-slate-600 mt-0.5">
              In questo browser ci sono dati salvati: {state.people.length} persone, {Object.keys(menPrograms).length} programmi mensili e {Object.keys(domPrograms).length} domenicali.
              Caricali nel database perché tutti li vedano. Fallo dal browser con i dati più aggiornati.
            </div>
          </div>
          <button type="button" className="btn-primary shrink-0" onClick={caricaDatiLocali} disabled={!state.people.length}>
            Carica nel database
          </button>
        </div>
      )}

      {/* ============ UTENTI (solo amministratori) ============ */}
      {activeTab === 'utenti' && currentUser.role === 'admin' && (
        <UtentiView utenti={utenti} mioUid={currentUser.uid} onShowToast={showToast} />
      )}

      {/* ============ DISCORSI ESTERNI ============ */}
      {activeTab === 'discorsiEsterni' && (
        <DiscorsiEsterniView
          state={state}
          onSaveState={saveState}
          isAdmin={canEdit}
          onShowToast={showToast}
          checkAdminPermission={checkAdminPermission}
        />
      )}

      {/* ============ TERRITORI ============ */}
      {activeTab === 'territori' && (
        <TerritoriView
          state={state}
          onSaveState={saveState}
          isAdmin={canEdit}
          onShowToast={showToast}
          checkAdminPermission={checkAdminPermission}
        />
      )}

      {/* ============ STATISTICHE ============ */}
      {activeTab === 'statistiche' && (
        <StatsView
          people={state.people}
          menRows={menRows}
          menTitle={menTitle}
          domRows={domRows}
          domTitle={domTitle}
          archives={state.mensileArchives || []}
        />
      )}

      </>}
          </main>
        </div>
      </div>

      {/* Excel Import Modal */}
      {showExcelImportModal && (
        <ExcelImportModal
          isOpen={showExcelImportModal}
          onClose={() => setShowExcelImportModal(false)}
          existingPeople={state.people}
          onImportPeople={handleImportPeopleFromExcel}
          onShowToast={showToast}
        />
      )}

      {/* Custom Confirm Modal */}
      {confirmModalState && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-fadeIn">
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl shadow-2xl max-w-md w-full p-5 space-y-4">
            <div className="flex items-center gap-3 text-rose-600 dark:text-rose-400">
              <div className="p-2.5 bg-rose-50 dark:bg-rose-950/60 rounded-xl">
                <RotateCcw className="w-5 h-5" />
              </div>
              <h3 className="text-base font-bold text-slate-900 dark:text-slate-100">
                {confirmModalState.title}
              </h3>
            </div>
            <p className="text-sm text-slate-600 dark:text-slate-300 leading-relaxed">
              {confirmModalState.message}
            </p>
            <div className="flex items-center justify-end gap-2.5 pt-3 border-t border-slate-100 dark:border-slate-800">
              <button
                onClick={() => setConfirmModalState(null)}
                className="btn-ghost text-xs px-3.5 py-2"
              >
                Annulla
              </button>
              <button
                onClick={confirmModalState.onConfirm}
                className={`${confirmModalState.confirmVariant === 'danger' ? 'btn-danger' : 'btn-primary'} text-xs px-4 py-2`}
              >
                {confirmModalState.confirmText || 'Conferma'}
              </button>
            </div>
          </div>
        </div>
      )}

      {toastMsg && (
        <div className="toast">
          <Check className="w-4 h-4 text-emerald-400" />
          <span>{toastMsg}</span>
        </div>
      )}
    </div>
  );
}
