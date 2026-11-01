import { ServizioCampoPrintTemplate, DEFAULT_FIELD_TOPIC } from './ServizioCampoPrintTemplate';
import { createSchedulingLedger } from '../utils/automaticScheduling';
import { exportProgramExcel } from '../utils/programExcel';
import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Compass,
  Calendar,
  Users,
  MapPin,
  Clock,
  Printer,
  Sparkles,
  Plus,
  Trash2,
  Edit2,
  Share2,
  FileSpreadsheet,
  Check,
  Search,
  Filter,
  BarChart2,
  ChevronLeft,
  ChevronRight,
  UserCheck,
  UserPlus,
  AlertTriangle,
  Info,
  Layers,
  Settings2,
  CheckSquare,
  Square,
  ArrowUpDown,
  Sun,
  Sunset,
  Download,
  Eye,
  RotateCcw,
  Star,
  Flag,
  PartyPopper,
  Bookmark,
  CalendarPlus,
  X
} from 'lucide-react';
import {
  SchedulingPrograms,
  StateData,
  ServizioCampoData,
  ServizioCampoConductor,
  ServizioCampoMeetingAssignment,
  ServizioCampoDayKey,
  ServizioCampoMeetingType,
  ServizioCampoDefaultSlotConfig,
  Person
} from '../types';

interface ServizioCampoViewProps {
  state: StateData;
  activePrograms?: SchedulingPrograms;
  onSaveState: (newState: StateData) => void;
  isAdmin: boolean;
  onShowToast: (msg: string) => void;
  checkAdminPermission: () => boolean;
  stampa?: boolean; // apre direttamente la sottoscheda di stampa (PDF della Tabella informazioni)
}

const MESI = [
  'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
  'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
];

const MESI_ABBR = ['gen', 'feb', 'mar', 'apr', 'mag', 'giu', 'lug', 'ago', 'set', 'ott', 'nov', 'dic'];
const GIORNI_SETT = ['Domenica', 'Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato'];

const DEFAULT_LOCATIONS = [
  'Sala del Regno',
  'Piazza Principale (Ribolla)',
  'Piazza del Popolo (Roccastrada)',
  'Parco Comunale',
  'Zona Residenziale',
  'Centro Commerciale',
  'Zoom (Online)'
];

const DEFAULT_SLOT_SETTINGS: Record<ServizioCampoDayKey, ServizioCampoDefaultSlotConfig> = {
  martediMattina: { time: '09:30', defaultLocation: 'Sala del Regno', active: true },
  giovediMattina: { time: '09:30', defaultLocation: 'Sala del Regno', active: true },
  sabatoPomeriggio: { time: '15:00', defaultLocation: 'Sala del Regno', active: true },
  domenicaPomeriggio: { time: '15:00', defaultLocation: 'Zoom (Online)', active: true },
  speciale: { time: '09:30', defaultLocation: 'Sala del Regno', active: true }
};

export const isZoomLocation = (loc: string | undefined | null): boolean => {
  if (!loc) return false;
  return loc.toLowerCase().includes('zoom');
};

export const isSundayDate = (dateStr: string): boolean => {
  if (!dateStr) return false;
  const parts = dateStr.split('-');
  if (parts.length < 3) return false;
  const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  return d.getDay() === 0;
};

export const SLOT_INFO: Record<ServizioCampoDayKey, { label: string; dayName: string; timeSlot: 'mattina' | 'pomeriggio' | 'sera'; colorBadge: string; icon: any }> = {
  martediMattina: {
    label: 'Martedì Mattina',
    dayName: 'martedi',
    timeSlot: 'mattina',
    colorBadge: 'bg-sky-100 text-sky-800 dark:bg-sky-950/70 dark:text-sky-300 border-sky-200 dark:border-sky-800',
    icon: Sun
  },
  giovediMattina: {
    label: 'Giovedì Mattina',
    dayName: 'giovedi',
    timeSlot: 'mattina',
    colorBadge: 'bg-amber-100 text-amber-800 dark:bg-amber-950/70 dark:text-amber-300 border-amber-200 dark:border-amber-800',
    icon: Sun
  },
  sabatoPomeriggio: {
    label: 'Sabato Pomeriggio',
    dayName: 'sabato',
    timeSlot: 'pomeriggio',
    colorBadge: 'bg-indigo-100 text-indigo-800 dark:bg-indigo-950/70 dark:text-indigo-300 border-indigo-200 dark:border-indigo-800',
    icon: Sunset
  },
  domenicaPomeriggio: {
    label: 'Domenica Pomeriggio',
    dayName: 'domenica',
    timeSlot: 'pomeriggio',
    colorBadge: 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950/70 dark:text-emerald-300 border-emerald-200 dark:border-emerald-800',
    icon: Sunset
  },
  speciale: {
    label: 'Data Particolare',
    dayName: 'speciale',
    timeSlot: 'mattina',
    colorBadge: 'bg-purple-100 text-purple-800 dark:bg-purple-950/70 dark:text-purple-300 border-purple-200 dark:border-purple-800',
    icon: Sparkles
  }
};

export const MEETING_TYPE_INFO: Record<ServizioCampoMeetingType, { label: string; badge: string; icon: any; prefix: string }> = {
  standard: {
    label: 'Standard',
    badge: 'bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300 border-slate-300 dark:border-slate-700',
    icon: Calendar,
    prefix: ''
  },
  festivo: {
    label: 'Giorno Festivo',
    badge: 'bg-rose-100 text-rose-800 dark:bg-rose-950/80 dark:text-rose-300 border-rose-300 dark:border-rose-800',
    icon: PartyPopper,
    prefix: '🎉 FESTIVO'
  },
  unificata: {
    label: 'Adunanza Unificata',
    badge: 'bg-purple-100 text-purple-800 dark:bg-purple-950/80 dark:text-purple-300 border-purple-300 dark:border-purple-800',
    icon: Users,
    prefix: '👥 UNIFICATA'
  },
  visita: {
    label: 'Visita Sorvegliante',
    badge: 'bg-amber-100 text-amber-800 dark:bg-amber-950/80 dark:text-amber-300 border-amber-300 dark:border-amber-800',
    icon: Star,
    prefix: '👔 VISITA SORVEGLIANTE'
  },
  campagna: {
    label: 'Campagna Speciale',
    badge: 'bg-cyan-100 text-cyan-800 dark:bg-cyan-950/80 dark:text-cyan-300 border-cyan-300 dark:border-cyan-800',
    icon: Flag,
    prefix: '🚀 CAMPAGNA SPECIALE'
  },
  speciale: {
    label: 'Altro Straordinario',
    badge: 'bg-indigo-100 text-indigo-800 dark:bg-indigo-950/80 dark:text-indigo-300 border-indigo-300 dark:border-indigo-800',
    icon: Sparkles,
    prefix: '⭐ SPECIALE'
  }
};

// Known Italian Holidays helper
const ITALIAN_HOLIDAYS_FIXED: Record<string, string> = {
  '01-01': 'Capodanno',
  '01-06': 'Epifania',
  '04-25': 'Festa della Liberazione',
  '05-01': 'Festa del Lavoro',
  '06-02': 'Festa della Repubblica',
  '08-15': 'Ferragosto',
  '11-01': 'Tutti i Santi',
  '12-08': 'Immacolata Concezione',
  '12-25': 'Natale',
  '12-26': 'Santo Stefano'
};

function formatShortDate(dateStr: string) {
  const parts = dateStr.split('-');
  if (parts.length < 3) return dateStr;
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[2], 10);
  const dt = new Date(y, m, d);
  const dayName = GIORNI_SETT[dt.getDay()];
  return `${dayName.slice(0, 3)} ${d} ${MESI_ABBR[m]}`;
}

function getDayOfWeekFromDate(dateStr: string): 'lunedi' | 'martedi' | 'mercoledi' | 'giovedi' | 'venerdi' | 'sabato' | 'domenica' {
  const parts = dateStr.split('-');
  if (parts.length < 3) return 'martedi';
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[2], 10);
  const dt = new Date(y, m, d);
  const day = dt.getDay();
  switch (day) {
    case 0: return 'domenica';
    case 1: return 'lunedi';
    case 2: return 'martedi';
    case 3: return 'mercoledi';
    case 4: return 'giovedi';
    case 5: return 'venerdi';
    case 6: return 'sabato';
    default: return 'martedi';
  }
}

export function ServizioCampoView({
  state,
  activePrograms,
  onSaveState,
  isAdmin,
  onShowToast,
  checkAdminPermission,
  stampa
}: ServizioCampoViewProps) {
  const now = new Date();
  const [selectedMonth, setSelectedMonth] = useState<number>(now.getMonth());
  const [selectedYear, setSelectedYear] = useState<number>(now.getFullYear());
  const [generationWarnings, setGenerationWarnings] = useState<string[]>([]);
  const [activeSubTab, setActiveSubTab] = useState<'programma' | 'conduttori' | 'luoghi' | 'statistiche' | 'stampa'>(stampa ? 'stampa' : 'programma');

  // Modal / Form for adding a Special Date (Festivo / Unificata / etc.)
  const [isSpecialModalOpen, setIsSpecialModalOpen] = useState(false);
  const [specialDateValue, setSpecialDateValue] = useState<string>(() => {
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    return `${now.getFullYear()}-${mm}-01`;
  });
  const [specialMeetingType, setSpecialMeetingType] = useState<ServizioCampoMeetingType>('festivo');
  const [specialTime, setSpecialTime] = useState('09:30');
  const [specialTimeSlot, setSpecialTimeSlot] = useState<'mattina' | 'pomeriggio' | 'sera'>('mattina');
  const [specialLocation, setSpecialLocation] = useState('Sala del Regno');
  const [specialConductorId, setSpecialConductorId] = useState<string>('');
  const [specialNoteInput, setSpecialNoteInput] = useState('');

  // Conductor form state
  const [editingConductorId, setEditingConductorId] = useState<string | null>(null);
  const [conductorName, setConductorName] = useState('');
  const [selectedPersonId, setSelectedPersonId] = useState<string>('');
  const [conductorAvailability, setConductorAvailability] = useState({
    martediMattina: true,
    giovediMattina: true,
    sabatoPomeriggio: true,
    domenicaPomeriggio: true
  });
  const [conductorNotes, setConductorNotes] = useState('');
  const [conductorSearch, setConductorSearch] = useState('');

  // Location form state
  const [newLocationName, setNewLocationName] = useState('');
  const [editingLocationIndex, setEditingLocationIndex] = useState<number | null>(null);
  const [editLocationValue, setEditLocationValue] = useState('');

  // Custom confirmation modals state (avoids sandboxed iframe window.confirm blocks)
  const [conductorToDelete, setConductorToDelete] = useState<{ id: string; name: string } | null>(null);
  const [meetingToDelete, setMeetingToDelete] = useState<ServizioCampoMeetingAssignment | null>(null);
  const [locationToDelete, setLocationToDelete] = useState<string | null>(null);

  // Schedule filtering / search
  const [scheduleFilterType, setScheduleFilterType] = useState<string>('all');
  const [scheduleSearch, setScheduleSearch] = useState('');

  // Extract or initialize data
  const data: ServizioCampoData = useMemo(() => {
    const raw: ServizioCampoData = state.servizioCampo || {
      conductors: [],
      schedule: [],
      locations: DEFAULT_LOCATIONS,
      defaultSettings: DEFAULT_SLOT_SETTINGS
    };
    return {
      printTopic: raw.printTopic ?? DEFAULT_FIELD_TOPIC,
      conductors: Array.isArray(raw.conductors) ? raw.conductors : [],
      schedule: Array.isArray(raw.schedule) ? raw.schedule : [],
      locations: Array.isArray(raw.locations) && raw.locations.length > 0 ? raw.locations : DEFAULT_LOCATIONS,
      defaultSettings: raw.defaultSettings || DEFAULT_SLOT_SETTINGS
    };
  }, [state.servizioCampo]);

  const conductors = data.conductors;
  const schedule = data.schedule;
  const locations = data.locations;
  const defaultSettings = data.defaultSettings || DEFAULT_SLOT_SETTINGS;

  // Locations used for rotation on non-Sunday meetings (strictly physical locations, outdoor points if available, never Zoom)
  const rotationLocations = useMemo(() => {
    const physicalOnly = locations.filter(l => !isZoomLocation(l));
    const nonSalaPhysical = physicalOnly.filter(l => l !== 'Sala del Regno');
    if (nonSalaPhysical.length > 0) return nonSalaPhysical;
    if (physicalOnly.length > 0) return physicalOnly;
    return locations;
  }, [locations]);

  // A location missing from the list shows as the first <select> option on screen but would print as-is: treat it as unset.
  const isKnownLocation = (loc: string | undefined | null): loc is string => !!loc && locations.includes(loc);
  const sundayFallbackLocation = locations.find(isZoomLocation) || locations[0];

  const saveServizioData = (updated: ServizioCampoData) => {
    onSaveState({
      ...state,
      servizioCampo: updated
    });
  };

  // Helper to get conductor by ID
  const getConductorById = (id: string | null | undefined): ServizioCampoConductor | undefined => {
    if (!id) return undefined;
    return conductors.find(c => c.id === id);
  };

  // Generate standard 4-slots meetings list for the month
  const standardMonthMeetingsList = useMemo(() => {
    const dates: { dateStr: string; dayOfWeek: 'martedi' | 'giovedi' | 'sabato' | 'domenica'; slotKey: ServizioCampoDayKey }[] = [];
    const daysInMonth = new Date(selectedYear, selectedMonth + 1, 0).getDate();

    for (let day = 1; day <= daysInMonth; day++) {
      const dt = new Date(selectedYear, selectedMonth, day);
      const dayNum = dt.getDay(); // 0 = Dom, 2 = Mar, 4 = Gio, 6 = Sab
      const mm = String(selectedMonth + 1).padStart(2, '0');
      const dd = String(day).padStart(2, '0');
      const dateStr = `${selectedYear}-${mm}-${dd}`;

      if (dayNum === 2) {
        dates.push({ dateStr, dayOfWeek: 'martedi', slotKey: 'martediMattina' });
      } else if (dayNum === 4) {
        dates.push({ dateStr, dayOfWeek: 'giovedi', slotKey: 'giovediMattina' });
      } else if (dayNum === 6) {
        dates.push({ dateStr, dayOfWeek: 'sabato', slotKey: 'sabatoPomeriggio' });
      } else if (dayNum === 0) {
        dates.push({ dateStr, dayOfWeek: 'domenica', slotKey: 'domenicaPomeriggio' });
      }
    }
    return dates;
  }, [selectedYear, selectedMonth]);

  // Combined month schedule items: standard recurring slots (with rotation for non-Sunday) + custom special/festive meetings
  const currentMonthSchedule = useMemo(() => {
    let nonSundayRotationIndex = 0;

    // 1. Map standard days
    const standardItems: ServizioCampoMeetingAssignment[] = standardMonthMeetingsList.map(item => {
      const existing = schedule.find(s => s.dateStr === item.dateStr && s.slotKey === item.slotKey);
      const slotDef = defaultSettings[item.slotKey] || DEFAULT_SLOT_SETTINGS[item.slotKey];
      const specialNoteFromGlobal = state.special[item.dateStr] || '';

      // Check if Italian holiday
      const monthDayKey = item.dateStr.slice(5);
      const holidayName = ITALIAN_HOLIDAYS_FIXED[monthDayKey] || '';
      const inferredType: ServizioCampoMeetingType = holidayName ? 'festivo' : 'standard';

      // Determine default location if not explicitly assigned
      let assignedLocation: string;
      if (item.slotKey === 'domenicaPomeriggio') {
        assignedLocation = isKnownLocation(slotDef.defaultLocation) ? slotDef.defaultLocation : sundayFallbackLocation;
      } else if (isKnownLocation(slotDef.defaultLocation)) {
        assignedLocation = slotDef.defaultLocation;
      } else {
        assignedLocation = rotationLocations[nonSundayRotationIndex % rotationLocations.length];
        nonSundayRotationIndex++;
      }

      if (existing) {
        return {
          ...existing,
          location: isKnownLocation(existing.location) ? existing.location : assignedLocation,
          meetingType: existing.meetingType || (holidayName ? 'festivo' : 'standard'),
          // Il nome della festività non va nelle note (anche se salvato in passato).
          specialNote: (existing.specialNote && existing.specialNote !== holidayName ? existing.specialNote : '') || specialNoteFromGlobal
        };
      }

      // Default meeting object if not yet saved in schedule
      const newMeeting: ServizioCampoMeetingAssignment = {
        id: `${item.dateStr}_${item.slotKey}`,
        dateStr: item.dateStr,
        dayOfWeek: item.dayOfWeek,
        timeSlot: SLOT_INFO[item.slotKey].timeSlot,
        slotKey: item.slotKey,
        meetingType: inferredType,
        time: slotDef.time,
        location: assignedLocation,
        conductorId: null,
        isActive: slotDef.active !== false,
        specialNote: specialNoteFromGlobal,
        notes: ''
      };
      return newMeeting;
    });

    // 2. Add any custom/extra meetings for this month that are not standard recurring slots
    const customItems: ServizioCampoMeetingAssignment[] = schedule.filter(s => {
      const parts = s.dateStr.split('-');
      if (parts.length < 2) return false;
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10) - 1;
      if (y !== selectedYear || m !== selectedMonth) return false;

      // Keep if marked as custom or slotKey === 'speciale' or not in standard slots list
      const isStandardRecurring = standardMonthMeetingsList.some(
        std => std.dateStr === s.dateStr && std.slotKey === s.slotKey
      );
      return s.isCustom || s.slotKey === 'speciale' || !isStandardRecurring;
    }).map(s => (isKnownLocation(s.location) ? s : { ...s, location: locations[0] }));

    // 3. Merge & sort chronologically by date and time
    const merged = [...standardItems, ...customItems];
    merged.sort((a, b) => {
      if (a.dateStr !== b.dateStr) {
        return a.dateStr.localeCompare(b.dateStr);
      }
      return (a.time || '').localeCompare(b.time || '');
    });

    return merged;
  }, [standardMonthMeetingsList, schedule, defaultSettings, locations, rotationLocations, state.special, selectedYear, selectedMonth]);

  // --- AUTOMATIC SCHEDULE GENERATION (with location rotation on non-Sundays, Zoom on Sunday) ---
  const handleAutoGenerate = () => {
    if (!checkAdminPermission()) return;
    if (conductors.length === 0) {
      onShowToast('Nessun conduttore inserito! Aggiungi prima i conduttori con le relative disponibilità.');
      setActiveSubTab('conduttori');
      return;
    }

    const ledger = createSchedulingLedger(state, activePrograms, {
      program: 'servizioCampo', dates: currentMonthSchedule.map(meeting => meeting.dateStr),
    });
    const warnings: string[] = [];
    const usageCounts: Record<string, number> = {};
    conductors.forEach(c => (usageCounts[c.id] = 0));

    // Count all-time usage to balance long-term
    schedule.filter(s => s.isActive !== false && !currentMonthSchedule.some(m => m.dateStr === s.dateStr)).forEach(s => {
      if (s.conductorId && usageCounts[s.conductorId] !== undefined) {
        usageCounts[s.conductorId]++;
      }
    });

    const monthUsageCounts: Record<string, number> = {};
    conductors.forEach(c => (monthUsageCounts[c.id] = 0));

    let lastConductorId: string | null = null;
    let autoRotationIndex = 0;

    const newMonthSchedule: ServizioCampoMeetingAssignment[] = currentMonthSchedule.map(meeting => {
      if (meeting.isActive === false) return meeting;
      const slotKey = meeting.slotKey;
      const slotDef = defaultSettings[slotKey] || DEFAULT_SLOT_SETTINGS[slotKey] || DEFAULT_SLOT_SETTINGS.speciale;
      const specialNote = meeting.specialNote || state.special[meeting.dateStr] || '';

      // Determine meeting location: preserve existing if already specified, else use slot default or rotation
      let meetingLocation = meeting.location;
      if (!meetingLocation) {
        if (slotKey === 'domenicaPomeriggio') {
          meetingLocation = slotDef.defaultLocation || 'Zoom (Online)';
        } else if (slotDef.defaultLocation && slotDef.defaultLocation !== 'rotazione') {
          meetingLocation = slotDef.defaultLocation;
        } else {
          meetingLocation = rotationLocations[autoRotationIndex % rotationLocations.length] || 'Sala del Regno';
          autoRotationIndex++;
        }
      }

      // Find available candidates
      const candidates = conductors.filter(c => {
        if (c.isActive === false) return false;
        if (ledger.isBusy(c, meeting.dateStr) || ledger.isUnavailable(c, meeting.dateStr)) return false;

        // For standard slotKeys, check specific availability; for 'speciale' slots, any active conductor can be considered
        if (slotKey !== 'speciale' && c.availability && !c.availability[slotKey]) {
          return false;
        }

        // Check unavail in global congregational calendar if linked to anagrafica person
        if (c.personId && state.unavail[c.personId]) {
          if (state.unavail[c.personId].includes(meeting.dateStr)) return false;
        }
        return true;
      });

      if (candidates.length === 0) {
        warnings.push(`${meeting.dateStr} ${meeting.time}: conduttore da assegnare.`);
        return {
          ...meeting,
          conductorId: null,
          time: meeting.time || slotDef.time,
          location: meetingLocation,
          specialNote
        };
      }

      // Sort by fewest month assignments, then all-time assignments, avoiding consecutive days
      candidates.sort((a, b) => {
        const ma = monthUsageCounts[a.id] || 0;
        const mb = monthUsageCounts[b.id] || 0;
        if (ma !== mb) return ma - mb;

        if (a.id === lastConductorId) return 1;
        if (b.id === lastConductorId) return -1;

        const ta = usageCounts[a.id] || 0;
        const tb = usageCounts[b.id] || 0;
        if (ta !== tb) return ta - tb;

        return a.name.localeCompare(b.name);
      });

      const picked = candidates[0];
      ledger.reserve(picked, meeting.dateStr);
      monthUsageCounts[picked.id] = (monthUsageCounts[picked.id] || 0) + 1;
      usageCounts[picked.id] = (usageCounts[picked.id] || 0) + 1;
      lastConductorId = picked.id;

      return {
        ...meeting,
        conductorId: picked.id,
        time: meeting.time || slotDef.time,
        location: meetingLocation,
        specialNote,
        isActive: meeting.isActive !== false
      };
    });

    // Merge with other months
    const otherSchedules = schedule.filter(s => {
      const parts = s.dateStr.split('-');
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10) - 1;
      return y !== selectedYear || m !== selectedMonth;
    });

    saveServizioData({
      ...data,
      schedule: [...otherSchedules, ...newMonthSchedule]
    });

    setGenerationWarnings(warnings);
    onShowToast(warnings.length ? `Programma generato: ${warnings.length} conduttori da assegnare, senza duplicare incarichi.` : 'Programma generato senza sovrapposizioni nella stessa data.');
  };

  const handleClearMonth = () => {
    if (!checkAdminPermission()) return;
    const otherSchedules = schedule.filter(s => {
      const parts = s.dateStr.split('-');
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10) - 1;
      return y !== selectedYear || m !== selectedMonth;
    });

    saveServizioData({
      ...data,
      schedule: otherSchedules
    });
    onShowToast(`Programma di ${MESI[selectedMonth]} azzerato.`);
  };

  const handleUpdateMeeting = (
    meetingId: string | undefined,
    dateStr: string,
    slotKey: ServizioCampoDayKey,
    updates: Partial<ServizioCampoMeetingAssignment>
  ) => {
    if (!checkAdminPermission()) return;

    let newSchedule = [...schedule];
    const existingIndex = newSchedule.findIndex(
      s => (meetingId && s.id === meetingId) || (s.dateStr === dateStr && s.slotKey === slotKey)
    );

    if (existingIndex >= 0) {
      newSchedule[existingIndex] = {
        ...newSchedule[existingIndex],
        ...updates
      };
    } else {
      const base = currentMonthSchedule.find(
        s => (meetingId && s.id === meetingId) || (s.dateStr === dateStr && s.slotKey === slotKey)
      );
      if (base) {
        newSchedule.push({
          ...base,
          ...updates
        });
      }
    }

    saveServizioData({
      ...data,
      schedule: newSchedule
    });
  };

  const handleDeleteMeeting = (meeting: ServizioCampoMeetingAssignment) => {
    if (!checkAdminPermission()) return;
    setMeetingToDelete(meeting);
  };

  const executeDeleteMeeting = () => {
    if (!meetingToDelete) return;
    const filtered = schedule.filter(s => {
      if (meetingToDelete.id && s.id === meetingToDelete.id) return false;
      if (s.dateStr === meetingToDelete.dateStr && s.slotKey === meetingToDelete.slotKey) return false;
      return true;
    });
    saveServizioData({
      ...data,
      schedule: filtered
    });
    setMeetingToDelete(null);
    onShowToast(`Adunanza del ${formatShortDate(meetingToDelete.dateStr)} rimossa.`);
  };

  // --- ADD SPECIAL DATE / MEETING ---
  const handleOpenAddSpecialModal = () => {
    const mm = String(selectedMonth + 1).padStart(2, '0');
    setSpecialDateValue(`${selectedYear}-${mm}-01`);
    setSpecialMeetingType('festivo');
    setSpecialTime('09:30');
    setSpecialTimeSlot('mattina');
    setSpecialLocation(rotationLocations[0] || 'Sala del Regno');
    setSpecialConductorId('');
    setSpecialNoteInput('');
    setIsSpecialModalOpen(true);
  };

  const handleSaveSpecialDateMeeting = (e: React.FormEvent) => {
    e.preventDefault();
    if (!checkAdminPermission()) return;

    if (!specialDateValue) {
      onShowToast('Seleziona una data valida');
      return;
    }

    const dayOfWeek = getDayOfWeekFromDate(specialDateValue);
    const newId = `custom_${specialDateValue}_${Math.random().toString(36).slice(2, 8)}`;

    const newMeeting: ServizioCampoMeetingAssignment = {
      id: newId,
      dateStr: specialDateValue,
      dayOfWeek,
      timeSlot: specialTimeSlot,
      slotKey: 'speciale',
      meetingType: specialMeetingType,
      time: specialTime || '09:30',
      location: specialLocation || locations[0] || 'Sala del Regno',
      conductorId: specialConductorId || null,
      isActive: true,
      specialNote: specialNoteInput.trim() || (specialMeetingType !== 'standard' ? MEETING_TYPE_INFO[specialMeetingType].label : ''),
      isCustom: true
    };

    saveServizioData({
      ...data,
      schedule: [...schedule, newMeeting]
    });

    setIsSpecialModalOpen(false);
    onShowToast(`Adunanza speciale aggiunta per ${formatShortDate(specialDateValue)}!`);
  };

  // Quick helper to add common Italian holiday if in current month
  const currentMonthHolidays = useMemo(() => {
    const mm = String(selectedMonth + 1).padStart(2, '0');
    const result: { dateStr: string; name: string }[] = [];
    Object.entries(ITALIAN_HOLIDAYS_FIXED).forEach(([key, name]) => {
      if (key.startsWith(`${mm}-`)) {
        const dd = key.split('-')[1];
        result.push({
          dateStr: `${selectedYear}-${mm}-${dd}`,
          name
        });
      }
    });
    return result;
  }, [selectedMonth, selectedYear]);

  // --- CONDUCTOR MANAGEMENT ---
  const handleSaveConductor = (e: React.FormEvent) => {
    e.preventDefault();
    if (!checkAdminPermission()) return;
    const name = conductorName.trim();
    if (!name) {
      onShowToast('Inserisci il nome del conduttore');
      return;
    }

    let gender: 'M' | 'F' = 'M';
    if (selectedPersonId) {
      const p = state.people.find(x => x.id === selectedPersonId);
      if (p) gender = p.gender;
    }

    if (editingConductorId) {
      const updatedList = conductors.map(c => {
        if (c.id === editingConductorId) {
          return {
            ...c,
            name,
            personId: selectedPersonId || null,
            gender,
            availability: { ...conductorAvailability },
            notes: conductorNotes.trim()
          };
        }
        return c;
      });
      saveServizioData({ ...data, conductors: updatedList });
      onShowToast(`Conduttore "${name}" aggiornato!`);
    } else {
      const newId = 'sc_' + Math.random().toString(36).slice(2, 9);
      const newConductor: ServizioCampoConductor = {
        id: newId,
        name,
        personId: selectedPersonId || null,
        gender,
        availability: { ...conductorAvailability },
        notes: conductorNotes.trim(),
        isActive: true
      };
      saveServizioData({ ...data, conductors: [...conductors, newConductor] });
      onShowToast(`Conduttore "${name}" aggiunto!`);
    }

    // Reset form
    setEditingConductorId(null);
    setConductorName('');
    setSelectedPersonId('');
    setConductorAvailability({
      martediMattina: true,
      giovediMattina: true,
      sabatoPomeriggio: true,
      domenicaPomeriggio: true
    });
    setConductorNotes('');
  };

  const handleEditConductor = (c: ServizioCampoConductor) => {
    if (!checkAdminPermission()) return;
    setEditingConductorId(c.id);
    setConductorName(c.name);
    setSelectedPersonId(c.personId || '');
    setConductorAvailability({
      martediMattina: !!c.availability?.martediMattina,
      giovediMattina: !!c.availability?.giovediMattina,
      sabatoPomeriggio: !!c.availability?.sabatoPomeriggio,
      domenicaPomeriggio: !!c.availability?.domenicaPomeriggio
    });
    setConductorNotes(c.notes || '');
  };

  const handleDeleteConductor = (id: string, name: string) => {
    if (!checkAdminPermission()) return;
    setConductorToDelete({ id, name });
  };

  const executeDeleteConductor = () => {
    if (!conductorToDelete) return;
    const { id, name } = conductorToDelete;
    const filtered = conductors.filter(c => c.id !== id);
    // Unassign this conductor from all scheduled meetings
    const updatedSchedule = schedule.map(s => {
      if (s.conductorId === id) {
        return { ...s, conductorId: null };
      }
      return s;
    });

    saveServizioData({ ...data, conductors: filtered, schedule: updatedSchedule });
    setConductorToDelete(null);
    onShowToast(`Conduttore "${name}" eliminato.`);
  };

  const handleImportBrothersFromAnagrafica = () => {
    if (!checkAdminPermission()) return;
    const brothers = state.people.filter(p => p.gender === 'M');
    if (brothers.length === 0) {
      onShowToast('Nessun fratello presente in anagrafica');
      return;
    }

    let addedCount = 0;
    const currentPersonIds = new Set(conductors.map(c => c.personId).filter(Boolean));
    const newConductors = [...conductors];

    brothers.forEach(b => {
      if (!currentPersonIds.has(b.id)) {
        newConductors.push({
          id: 'sc_' + Math.random().toString(36).slice(2, 9),
          name: b.name,
          personId: b.id,
          gender: 'M',
          availability: {
            martediMattina: true,
            giovediMattina: true,
            sabatoPomeriggio: true,
            domenicaPomeriggio: true
          },
          isActive: true
        });
        addedCount++;
      }
    });

    if (addedCount > 0) {
      saveServizioData({ ...data, conductors: newConductors });
      onShowToast(`Importati ${addedCount} fratelli dall'anagrafica!`);
    } else {
      onShowToast('Tutti i fratelli sono già presenti tra i conduttori.');
    }
  };

  // --- LOCATION MANAGEMENT ---
  const handleAddLocation = (e: React.FormEvent) => {
    e.preventDefault();
    if (!checkAdminPermission()) return;
    const loc = newLocationName.trim();
    if (!loc) return;
    if (locations.includes(loc)) {
      onShowToast('Questo luogo è già presente nell\'elenco');
      return;
    }
    const updated = [...locations, loc];
    saveServizioData({ ...data, locations: updated });
    setNewLocationName('');
    onShowToast(`Luogo "${loc}" aggiunto all'elenco!`);
  };

  const handleSaveEditLocation = (index: number) => {
    if (!checkAdminPermission()) return;
    const val = editLocationValue.trim();
    if (!val) return;
    const oldVal = locations[index];
    const updated = [...locations];
    updated[index] = val;

    // Update existing assignments referencing oldVal
    const updatedSchedule = schedule.map(s => {
      if (s.location === oldVal) {
        return { ...s, location: val };
      }
      return s;
    });

    saveServizioData({ ...data, locations: updated, schedule: updatedSchedule });
    setEditingLocationIndex(null);
    setEditLocationValue('');
    onShowToast(`Luogo aggiornato in "${val}"!`);
  };

  const handleDeleteLocation = (loc: string) => {
    if (!checkAdminPermission()) return;
    if (locations.length <= 1) {
      onShowToast('Deve rimanere almeno un luogo nell\'elenco.');
      return;
    }
    setLocationToDelete(loc);
  };

  const executeDeleteLocation = () => {
    if (!locationToDelete) return;
    if (locations.length <= 1) {
      onShowToast('Deve rimanere almeno un luogo nell\'elenco.');
      setLocationToDelete(null);
      return;
    }
    const updated = locations.filter(l => l !== locationToDelete);
    saveServizioData({ ...data, locations: updated });
    setLocationToDelete(null);
    onShowToast(`Luogo "${locationToDelete}" rimosso.`);
  };

  const handleSaveSlotSetting = (
    key: ServizioCampoDayKey,
    updates: Partial<ServizioCampoDefaultSlotConfig>
  ) => {
    if (!checkAdminPermission()) return;
    const current = defaultSettings[key] || DEFAULT_SLOT_SETTINGS[key];
    const updatedSettings = {
      ...defaultSettings,
      [key]: {
        ...current,
        ...updates
      }
    };
    saveServizioData({ ...data, defaultSettings: updatedSettings });
    onShowToast('Impostazioni orario e luogo predefinito salvate.');
  };

  // --- STATS / USAGE SUMMARY ---
  const conductorStats = useMemo(() => {
    return conductors.map(c => {
      let monthTotal = 0;
      let monthMartedi = 0;
      let monthGiovedi = 0;
      let monthSabato = 0;
      let monthDomenica = 0;
      let monthSpeciali = 0;
      const monthAssignments: { dateStr: string; dayLabel: string; slotKey: ServizioCampoDayKey; meetingType?: ServizioCampoMeetingType; time: string; location: string }[] = [];

      let allTimeTotal = 0;
      let lastDate: string | null = null;

      schedule.forEach(s => {
        if (s.conductorId === c.id) {
          allTimeTotal++;
          if (!lastDate || s.dateStr > lastDate) {
            lastDate = s.dateStr;
          }

          const parts = s.dateStr.split('-');
          const y = parseInt(parts[0], 10);
          const m = parseInt(parts[1], 10) - 1;
          if (y === selectedYear && m === selectedMonth) {
            monthTotal++;
            if (s.slotKey === 'martediMattina') monthMartedi++;
            else if (s.slotKey === 'giovediMattina') monthGiovedi++;
            else if (s.slotKey === 'sabatoPomeriggio') monthSabato++;
            else if (s.slotKey === 'domenicaPomeriggio') monthDomenica++;
            else monthSpeciali++;

            monthAssignments.push({
              dateStr: s.dateStr,
              dayLabel: formatShortDate(s.dateStr),
              slotKey: s.slotKey,
              meetingType: s.meetingType,
              time: s.time,
              location: s.location
            });
          }
        }
      });

      const totalAvailSlots = (c.availability?.martediMattina ? 1 : 0) +
        (c.availability?.giovediMattina ? 1 : 0) +
        (c.availability?.sabatoPomeriggio ? 1 : 0) +
        (c.availability?.domenicaPomeriggio ? 1 : 0);

      return {
        id: c.id,
        name: c.name,
        gender: c.gender || 'M',
        personId: c.personId,
        isActive: c.isActive !== false,
        availability: c.availability,
        totalAvailSlots,
        monthTotal,
        monthMartedi,
        monthGiovedi,
        monthSabato,
        monthDomenica,
        monthSpeciali,
        monthAssignments,
        allTimeTotal,
        lastDate
      };
    });
  }, [conductors, schedule, selectedYear, selectedMonth]);

  const activeConductorsInMonth = conductorStats.filter(c => c.monthTotal > 0).length;
  const totalMonthMeetings = currentMonthSchedule.filter(s => s.isActive !== false).length;
  const assignedMonthMeetings = currentMonthSchedule.filter(s => s.isActive !== false && s.conductorId).length;

  // --- EXPORT WHATSAPP ---
  const handleCopyWhatsApp = () => {
    let text = `📖 *PROGRAMMA ADUNANZE SERVIZIO DI CAMPO*\n`;
    text += `📅 *${MESI[selectedMonth].toUpperCase()} ${selectedYear}*\n`;
    text += `━━━━━━━━━━━━━━━━━━━━━\n\n`;

    currentMonthSchedule.forEach(m => {
      const dtFormatted = formatShortDate(m.dateStr);
      const slot = SLOT_INFO[m.slotKey] || SLOT_INFO.speciale;
      const typeInfo = MEETING_TYPE_INFO[m.meetingType || 'standard'];
      const conductor = getConductorById(m.conductorId);

      if (m.isActive === false) {
        text += `❌ *${dtFormatted}* (${slot.label}) - _Sospesa_\n`;
        if (m.specialNote) text += `   ℹ️ ${m.specialNote}\n`;
        text += `\n`;
        return;
      }

      let typeTag = '';
      if (m.meetingType && m.meetingType !== 'standard') {
        typeTag = ` [${typeInfo.prefix || typeInfo.label.toUpperCase()}]`;
      }

      text += `📍 *${dtFormatted}*${typeTag} - Ore *${m.time}*\n`;
      text += `   🏛️ *Luogo:* ${m.location}\n`;
      text += `   👤 *Conduttore:* ${conductor ? conductor.name : '— _Da assegnare_'}\n`;
      if (m.specialNote) text += `   ℹ️ _${m.specialNote}_\n`;
      text += `\n`;
    });

    text += `━━━━━━━━━━━━━━━━━━━━━\n`;
    text += `_Si prega di arrivare con qualche minuto di anticipo._`;

    navigator.clipboard.writeText(text);
    onShowToast('Programma formattato copiato negli appunti per WhatsApp!');
  };

  // --- EXPORT EXCEL ---
  const handleExportExcel = async () => {
try {
setActiveSubTab('stampa');
await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));

    const headers = [
      'Data',
      'Giorno',
      'Appuntamento',
      'Orario',
      'Luogo di Ritrovo',
      'Conduttore Assegnato',
      'Note / Specifiche'
    ];

    const rows = currentMonthSchedule.map(m => {
      const conductor = getConductorById(m.conductorId);
      const slot = SLOT_INFO[m.slotKey] || SLOT_INFO.speciale;

      return [
        m.dateStr,
        formatShortDate(m.dateStr),
        slot.label,
        m.time,
        m.location,
        conductor ? conductor.name : '—',
        m.specialNote || ''
      ];
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 14 },
      { wch: 20 },
      { wch: 22 },
      { wch: 10 },
      { wch: 30 },
      { wch: 28 },
      { wch: 35 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Servizio di Campo');
    await exportProgramExcel(wb, `Adunanze_Servizio_Campo_${MESI[selectedMonth]}_${selectedYear}.xlsx`, '.excel-service');
    onShowToast('Programma esportato in Excel (.xlsx)!');

} catch (error) { onShowToast(error instanceof Error ? error.message : 'Errore durante l’esportazione Excel.'); }
};

  // Filter schedule rows
  const filteredSchedule = useMemo(() => {
    return currentMonthSchedule.filter(m => {
      // Type filter
      if (scheduleFilterType === 'speciali') {
        if (m.meetingType === 'standard' && m.slotKey !== 'speciale' && !m.isCustom) return false;
      } else if (scheduleFilterType !== 'all' && m.slotKey !== scheduleFilterType && m.meetingType !== scheduleFilterType) {
        return false;
      }

      // Search filter
      if (scheduleSearch.trim()) {
        const query = scheduleSearch.toLowerCase();
        const conductor = getConductorById(m.conductorId);
        const matchName = conductor ? conductor.name.toLowerCase().includes(query) : false;
        const matchLoc = (m.location || '').toLowerCase().includes(query);
        const matchNote = (m.specialNote || '').toLowerCase().includes(query);
        const matchDate = m.dateStr.includes(query);
        return matchName || matchLoc || matchNote || matchDate;
      }

      return true;
    });
  }, [currentMonthSchedule, scheduleFilterType, scheduleSearch, conductors]);

  return (
    <div className="space-y-6">
      {/* Header Banner */}
      <div className="no-print bg-gradient-to-r from-sky-900 via-indigo-900 to-slate-900 text-white p-5 rounded-2xl shadow-md">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="p-3 bg-sky-500/20 rounded-xl border border-sky-400/30 text-sky-300">
              <Compass className="w-7 h-7" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl sm:text-2xl font-black tracking-tight">
                  Adunanze per il Servizio di Campo
                </h1>
                <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-sky-400/20 text-sky-200 border border-sky-300/30">
                  Gestione Conduttori &amp; Date Particolari
                </span>
              </div>
              <p className="text-xs sm:text-sm text-slate-300 mt-0.5">
                Pianificazione adunanze settimanali, giorni festivi, unificate e uscite straordinarie
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <div className="bg-slate-800/80 border border-slate-700/80 px-3.5 py-2 rounded-xl text-center">
              <span className="block text-xs text-slate-400 font-medium">Conduttori</span>
              <span className="text-lg font-bold text-sky-300">{conductors.length}</span>
            </div>
            <div className="bg-slate-800/80 border border-slate-700/80 px-3.5 py-2 rounded-xl text-center">
              <span className="block text-xs text-slate-400 font-medium">Adunanze Mese</span>
              <span className="text-lg font-bold text-emerald-300">{currentMonthSchedule.length}</span>
            </div>
            <div className="bg-slate-800/80 border border-slate-700/80 px-3.5 py-2 rounded-xl text-center">
              <span className="block text-xs text-slate-400 font-medium">Copertura</span>
              <span className="text-lg font-bold text-amber-300">
                {assignedMonthMeetings}/{totalMonthMeetings}
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Sub-Navigation Tabs */}
      <div className="flex items-center gap-1.5 border-b border-slate-200 dark:border-slate-800 overflow-x-auto pb-0.5 no-print">
        <button
          type="button"
          onClick={() => setActiveSubTab('programma')}
          className={`flex items-center gap-2 py-2.5 px-4 text-xs font-semibold rounded-t-xl border-b-2 transition-all ${
            activeSubTab === 'programma'
              ? 'border-sky-600 text-sky-600 dark:text-sky-400 bg-sky-50/50 dark:bg-sky-950/20'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
          }`}
        >
          <Calendar className="w-4 h-4 text-sky-600 dark:text-sky-400" />
          <span>Programma Mese</span>
          <span className="px-1.5 py-0.2 rounded-full bg-sky-100 dark:bg-sky-950 text-[10px] font-bold text-sky-700 dark:text-sky-300">
            {MESI_ABBR[selectedMonth]} {selectedYear}
          </span>
        </button>

        <button
          type="button"
          onClick={() => setActiveSubTab('conduttori')}
          className={`flex items-center gap-2 py-2.5 px-4 text-xs font-semibold rounded-t-xl border-b-2 transition-all ${
            activeSubTab === 'conduttori'
              ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400 bg-indigo-50/50 dark:bg-indigo-950/20'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
          }`}
        >
          <Users className="w-4 h-4 text-indigo-600 dark:text-indigo-400" />
          <span>Conduttori &amp; Disponibilità</span>
          <span className="px-1.5 py-0.2 rounded-full bg-indigo-100 dark:bg-indigo-950 text-[10px] font-bold text-indigo-700 dark:text-indigo-300">
            {conductors.length}
          </span>
        </button>

        <button
          type="button"
          onClick={() => setActiveSubTab('luoghi')}
          className={`flex items-center gap-2 py-2.5 px-4 text-xs font-semibold rounded-t-xl border-b-2 transition-all ${
            activeSubTab === 'luoghi'
              ? 'border-amber-600 text-amber-600 dark:text-amber-400 bg-amber-50/50 dark:bg-amber-950/20'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
          }`}
        >
          <MapPin className="w-4 h-4 text-amber-600 dark:text-amber-400" />
          <span>Luoghi &amp; Orari Predefiniti</span>
          <span className="px-1.5 py-0.2 rounded-full bg-amber-100 dark:bg-amber-950 text-[10px] font-bold text-amber-700 dark:text-amber-300">
            {locations.length}
          </span>
        </button>

        <button
          type="button"
          onClick={() => setActiveSubTab('statistiche')}
          className={`flex items-center gap-2 py-2.5 px-4 text-xs font-semibold rounded-t-xl border-b-2 transition-all ${
            activeSubTab === 'statistiche'
              ? 'border-purple-600 text-purple-600 dark:text-purple-400 bg-purple-50/50 dark:bg-purple-950/20'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
          }`}
        >
          <BarChart2 className="w-4 h-4 text-purple-600 dark:text-purple-400" />
          <span>Riepilogo Utilizzo</span>
          <span className="px-1.5 py-0.2 rounded-full bg-purple-100 dark:bg-purple-950 text-[10px] font-bold text-purple-700 dark:text-purple-300">
            {activeConductorsInMonth}/{conductors.length}
          </span>
        </button>

        <button
          type="button"
          onClick={() => setActiveSubTab('stampa')}
          className={`flex items-center gap-2 py-2.5 px-4 text-xs font-semibold rounded-t-xl border-b-2 transition-all ${
            activeSubTab === 'stampa'
              ? 'border-emerald-600 text-emerald-600 dark:text-emerald-400 bg-emerald-50/50 dark:bg-emerald-950/20'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
          }`}
        >
          <Printer className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
          <span>Stampa &amp; Condividi</span>
        </button>
      </div>

      {/* ======================================================== */}
      {/* SUBTAB 1: PROGRAMMA MESE                                */}
      {/* ======================================================== */}
      {activeSubTab === 'programma' && generationWarnings.length > 0 && (
        <div className="card no-print border-amber-300 bg-amber-50 dark:bg-amber-950/30 text-xs" role="status">
          <p className="font-bold mb-2">Incarichi da completare</p>
          <ul className="list-disc pl-4 space-y-1">{generationWarnings.map((warning, i) => <li key={i}>{warning}</li>)}</ul>
        </div>
      )}
      {activeSubTab === 'programma' && (
        <div className="space-y-6">
          {/* Controls Bar */}
          <div className="card flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-center gap-3 flex-wrap">
              <div className="flex items-center gap-1.5 bg-slate-100 dark:bg-slate-800/80 p-1 rounded-xl border border-slate-200 dark:border-slate-700">
                <button
                  type="button"
                  onClick={() => {
                    if (selectedMonth === 0) {
                      setSelectedMonth(11);
                      setSelectedYear(selectedYear - 1);
                    } else {
                      setSelectedMonth(selectedMonth - 1);
                    }
                  }}
                  className="btn-ghost p-1.5"
                  title="Mese precedente"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <div className="text-sm font-bold text-slate-800 dark:text-slate-100 min-w-[130px] text-center">
                  {MESI[selectedMonth]} {selectedYear}
                </div>
                <button
                  type="button"
                  onClick={() => {
                    if (selectedMonth === 11) {
                      setSelectedMonth(0);
                      setSelectedYear(selectedYear + 1);
                    } else {
                      setSelectedMonth(selectedMonth + 1);
                    }
                  }}
                  className="btn-ghost p-1.5"
                  title="Mese successivo"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>

              <select
                value={selectedMonth}
                onChange={e => setSelectedMonth(Number(e.target.value))}
                className="inp py-1.5 text-xs font-semibold"
              >
                {MESI.map((m, idx) => (
                  <option key={m} value={idx}>
                    {m}
                  </option>
                ))}
              </select>

              <select
                value={selectedYear}
                onChange={e => setSelectedYear(Number(e.target.value))}
                className="inp py-1.5 text-xs font-semibold"
              >
                {[selectedYear - 1, selectedYear, selectedYear + 1].map(y => (
                  <option key={y} value={y}>
                    {y}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={handleOpenAddSpecialModal}
                disabled={!isAdmin}
                className="btn-primary bg-indigo-600 hover:bg-indigo-700 text-white flex items-center gap-1.5"
                title="Aggiungi una data particolare (giorno festivo, adunanza unificata, visita o orario speciale)"
              >
                <CalendarPlus className="w-4 h-4" />
                <span>+ Data Particolare / Festivo</span>
              </button>

              <button
                type="button"
                onClick={handleAutoGenerate}
                disabled={!isAdmin}
                className="btn-primary flex items-center gap-1.5"
                title="Assegna automaticamente i conduttori in base alle disponibilità"
              >
                <Sparkles className="w-4 h-4" />
                <span>Genera Programma Mese</span>
              </button>

              <button
                type="button"
                onClick={handleCopyWhatsApp}
                className="btn-ghost text-xs flex items-center gap-1.5 text-emerald-600 dark:text-emerald-400 hover:bg-emerald-50 dark:hover:bg-emerald-950/30 border border-emerald-200 dark:border-emerald-800"
              >
                <Share2 className="w-3.5 h-3.5" />
                <span>WhatsApp</span>
              </button>

              <button
                type="button"
                onClick={handleExportExcel}
                className="btn-ghost text-xs flex items-center gap-1.5 text-sky-600 dark:text-sky-400 hover:bg-sky-50 dark:hover:bg-sky-950/30 border border-sky-200 dark:border-sky-800"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span>Excel</span>
              </button>

              <button
                type="button"
                onClick={handleClearMonth}
                disabled={!isAdmin}
                className="btn-ghost text-xs text-rose-500 hover:text-rose-700"
                title="Azzera il programma di questo mese"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Azzera</span>
              </button>
            </div>
          </div>

          {/* Quick Festive Helper Bar (if current month has holidays) */}
          {currentMonthHolidays.length > 0 && (
            <div className="bg-amber-50/70 dark:bg-amber-950/30 border border-amber-200/80 dark:border-amber-900/60 p-3 rounded-2xl flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs">
              <div className="flex items-center gap-2 text-amber-900 dark:text-amber-200">
                <PartyPopper className="w-4 h-4 text-amber-600 dark:text-amber-400 shrink-0" />
                <span>
                  <strong>Festività del mese:</strong>{' '}
                  {currentMonthHolidays.map(h => `${formatShortDate(h.dateStr)} (${h.name})`).join(', ')}
                </span>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-[11px] text-amber-800 dark:text-amber-300">
                  Puoi aggiungere o personalizzare uscite festive o adunanze unificate per queste date.
                </span>
              </div>
            </div>
          )}

          {/* Schedule Table */}
          <div className="card space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100 dark:border-slate-800">
              <div className="flex items-center gap-2">
                <Calendar className="w-5 h-5 text-sky-600 dark:text-sky-400" />
                <h2 className="card-title">
                  Appuntamenti di {MESI[selectedMonth]} {selectedYear} ({currentMonthSchedule.length} uscite)
                </h2>
              </div>

              <div className="flex items-center gap-2 flex-wrap">
                <div className="relative">
                  <input
                    type="text"
                    placeholder="Cerca conduttore, luogo o data..."
                    value={scheduleSearch}
                    onChange={e => setScheduleSearch(e.target.value)}
                    className="inp pl-8 text-xs py-1.5 w-52"
                  />
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                </div>

                <select
                  value={scheduleFilterType}
                  onChange={e => setScheduleFilterType(e.target.value)}
                  className="inp text-xs py-1.5"
                >
                  <option value="all">Tutti gli appuntamenti</option>
                  <option value="speciali">🌟 Solo Date Particolari / Festive / Unificate</option>
                  <option value="martediMattina">Solo Martedì Mattina</option>
                  <option value="giovediMattina">Solo Giovedì Mattina</option>
                  <option value="sabatoPomeriggio">Solo Sabato Pomeriggio</option>
                  <option value="domenicaPomeriggio">Solo Domenica Pomeriggio</option>
                </select>
              </div>
            </div>

            <div className="overflow-x-auto rounded-2xl border border-slate-200 dark:border-slate-800">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100/90 dark:bg-slate-800/90 text-slate-700 dark:text-slate-300 font-semibold border-b border-slate-200 dark:border-slate-700">
                    <th className="p-3 w-36">Data e Giorno</th>
                    <th className="p-3 w-44">Tipologia / Evento</th>
                    <th className="p-3 w-28">Orario</th>
                    <th className="p-3 min-w-[180px]">Luogo di Ritrovo</th>
                    <th className="p-3 min-w-[200px]">Conduttore Assegnato</th>
                    <th className="p-3 min-w-[180px]">Note / Specifiche</th>
                    <th className="p-3 w-24 text-center">Stato</th>
                    <th className="p-3 w-14 text-center">Azioni</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                  {filteredSchedule.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="p-8 text-center text-slate-400">
                        Nessun appuntamento trovato con i filtri selezionati.
                      </td>
                    </tr>
                  ) : (
                    filteredSchedule.map(meeting => {
                      const slot = SLOT_INFO[meeting.slotKey] || SLOT_INFO.speciale;
                      const isSuspended = meeting.isActive === false;
                      const meetingType = meeting.meetingType || 'standard';
                      const typeInfo = MEETING_TYPE_INFO[meetingType] || MEETING_TYPE_INFO.standard;

                      // Filter conductors available for this specific slot if standard, or all if speciale
                      const availableConductors = conductors.filter(
                        c => c.isActive !== false && (meeting.slotKey === 'speciale' || c.availability[meeting.slotKey])
                      );
                      const otherConductors = conductors.filter(
                        c => c.isActive !== false && (meeting.slotKey !== 'speciale' && !c.availability[meeting.slotKey])
                      );

                      return (
                        <tr
                          key={meeting.id || `${meeting.dateStr}_${meeting.slotKey}`}
                          className={`transition-colors ${
                            isSuspended
                              ? 'bg-rose-50/40 dark:bg-rose-950/20 opacity-70'
                              : meetingType === 'festivo'
                              ? 'bg-rose-50/20 dark:bg-rose-950/10 hover:bg-rose-50/40'
                              : meetingType === 'unificata'
                              ? 'bg-purple-50/20 dark:bg-purple-950/10 hover:bg-purple-50/40'
                              : meetingType === 'visita'
                              ? 'bg-amber-50/20 dark:bg-amber-950/10 hover:bg-amber-50/40'
                              : 'hover:bg-slate-50/80 dark:hover:bg-slate-800/40'
                          }`}
                        >
                          {/* Date and Day */}
                          <td className="p-3 font-semibold text-slate-900 dark:text-slate-100">
                            <div className="flex items-center gap-1.5">
                              <span>{formatShortDate(meeting.dateStr)}</span>
                              {meeting.isCustom && (
                                <span className="px-1 py-0.2 rounded bg-indigo-100 dark:bg-indigo-950 text-indigo-700 dark:text-indigo-300 text-[9px] font-bold" title="Adunanza straordinaria aggiunta manualmente">
                                  EXTRA
                                </span>
                              )}
                            </div>
                            <div className="text-[10px] text-slate-400 font-normal">{meeting.dateStr}</div>
                          </td>

                          {/* Meeting Type Selector / Badge */}
                          <td className="p-3">
                            <div className="space-y-1">
                              <select
                                value={meetingType}
                                disabled={!isAdmin || isSuspended}
                                onChange={e =>
                                  handleUpdateMeeting(meeting.id, meeting.dateStr, meeting.slotKey, {
                                    meetingType: e.target.value as ServizioCampoMeetingType
                                  })
                                }
                                className={`inp py-1 px-2 text-xs font-bold w-full rounded-lg border ${typeInfo.badge}`}
                              >
                                <option value="standard">📅 Standard ({slot.label})</option>
                                <option value="festivo">🎉 Giorno Festivo</option>
                                <option value="unificata">👥 Adunanza Unificata</option>
                                <option value="visita">👔 Visita Sorvegliante</option>
                                <option value="campagna">🚀 Campagna Speciale</option>
                                <option value="speciale">⭐ Altro Straordinario</option>
                              </select>
                            </div>
                          </td>

                          {/* Time picker */}
                          <td className="p-3">
                            <input
                              type="time"
                              value={meeting.time}
                              disabled={!isAdmin || isSuspended}
                              onChange={e =>
                                handleUpdateMeeting(meeting.id, meeting.dateStr, meeting.slotKey, {
                                  time: e.target.value
                                })
                              }
                              className="inp py-1 px-2 text-xs font-mono font-bold w-24"
                            />
                          </td>

                          {/* Location selector */}
                          <td className="p-3">
                            <select
                              value={meeting.location}
                              disabled={!isAdmin || isSuspended}
                              onChange={e =>
                                handleUpdateMeeting(meeting.id, meeting.dateStr, meeting.slotKey, {
                                  location: e.target.value
                                })
                              }
                              className={`inp py-1 px-2 text-xs font-semibold w-full ${
                                isZoomLocation(meeting.location)
                                  ? 'bg-emerald-50/60 dark:bg-emerald-950/30 text-emerald-900 dark:text-emerald-200 border-emerald-300 dark:border-emerald-800'
                                  : 'text-slate-900 dark:text-slate-100'
                              }`}
                            >
                              {locations.map(loc => (
                                <option key={loc} value={loc}>
                                  {isZoomLocation(loc) ? `💻 ${loc}` : loc}
                                </option>
                              ))}
                            </select>
                          </td>

                          {/* Conductor selector */}
                          <td className="p-3">
                            <select
                              value={meeting.conductorId || ''}
                              disabled={!isAdmin || isSuspended}
                              onChange={e =>
                                handleUpdateMeeting(meeting.id, meeting.dateStr, meeting.slotKey, {
                                  conductorId: e.target.value || null
                                })
                              }
                              className={`inp py-1 px-2 text-xs font-semibold w-full ${
                                !meeting.conductorId && !isSuspended
                                  ? 'border-amber-400 bg-amber-50/50 dark:bg-amber-950/20 text-amber-900 dark:text-amber-200'
                                  : 'text-slate-900 dark:text-slate-100'
                              }`}
                            >
                              <option value="">— Seleziona Conduttore —</option>
                              <optgroup label="✓ Disponibili per questo turno">
                                {availableConductors.map(c => (
                                  <option key={c.id} value={c.id}>
                                    {c.name}
                                  </option>
                                ))}
                              </optgroup>
                              {otherConductors.length > 0 && (
                                <optgroup label="⚠ Altri conduttori">
                                  {otherConductors.map(c => (
                                    <option key={c.id} value={c.id}>
                                      {c.name}
                                    </option>
                                  ))}
                                </optgroup>
                              )}
                            </select>
                          </td>

                          {/* Special Note / Event detail */}
                          <td className="p-3">
                            <input
                              type="text"
                              placeholder="es. 1° Maggio, Visita Sorvegliante..."
                              value={meeting.specialNote || ''}
                              disabled={!isAdmin}
                              onChange={e =>
                                handleUpdateMeeting(meeting.id, meeting.dateStr, meeting.slotKey, {
                                  specialNote: e.target.value
                                })
                              }
                              className="inp py-1 px-2 text-xs w-full"
                            />
                          </td>

                          {/* Active / Suspended toggle */}
                          <td className="p-3 text-center">
                            <button
                              type="button"
                              disabled={!isAdmin}
                              onClick={() =>
                                handleUpdateMeeting(meeting.id, meeting.dateStr, meeting.slotKey, {
                                  isActive: !meeting.isActive
                                })
                              }
                              className={`p-1.5 rounded-lg border text-xs font-bold transition-all ${
                                meeting.isActive !== false
                                  ? 'bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-950/60 dark:text-emerald-300 dark:border-emerald-800'
                                  : 'bg-rose-100 text-rose-700 border-rose-300 dark:bg-rose-950 dark:text-rose-300 dark:border-rose-800'
                              }`}
                              title={meeting.isActive !== false ? 'Adunanza attiva (clicca per sospendere)' : 'Adunanza sospesa (clicca per riattivare)'}
                            >
                              {meeting.isActive !== false ? 'Attiva' : 'Sospesa'}
                            </button>
                          </td>

                          {/* Action (e.g. Delete if custom extra date) */}
                          <td className="p-3 text-center">
                            {meeting.isCustom ? (
                              <button
                                type="button"
                                disabled={!isAdmin}
                                onClick={() => handleDeleteMeeting(meeting)}
                                className="p-1 text-slate-400 hover:text-rose-600 rounded"
                                title="Elimina questa data particolare"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            ) : (
                              <span className="text-slate-300 dark:text-slate-700 text-xs">—</span>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* MODAL: AGGIUNGI DATA PARTICOLARE / FESTIVO / UNIFICATA   */}
      {/* ======================================================== */}
      {isSpecialModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-fadeIn">
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl shadow-2xl max-w-lg w-full p-6 space-y-5">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100 dark:border-slate-800">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-indigo-100 dark:bg-indigo-950/80 text-indigo-600 dark:text-indigo-300 rounded-xl">
                  <CalendarPlus className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-900 dark:text-slate-100">
                    Aggiungi Adunanza Particolare / Festiva
                  </h3>
                  <p className="text-xs text-slate-500">
                    Specifica uscite festive, adunanze unificate o appuntamenti straordinari
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsSpecialModalOpen(false)}
                className="p-1 text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 rounded-lg"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveSpecialDateMeeting} className="space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="label">Data dell'Appuntamento *</label>
                  <input
                    type="date"
                    required
                    value={specialDateValue}
                    onChange={e => setSpecialDateValue(e.target.value)}
                    className="inp text-xs font-semibold"
                  />
                  {specialDateValue && (
                    <span className="text-[11px] text-slate-500 mt-1 block">
                      Giorno: <strong>{formatShortDate(specialDateValue)}</strong>
                    </span>
                  )}
                </div>

                <div>
                  <label className="label">Tipologia Evento *</label>
                  <select
                    value={specialMeetingType}
                    onChange={e => setSpecialMeetingType(e.target.value as ServizioCampoMeetingType)}
                    className="inp text-xs font-semibold"
                  >
                    <option value="festivo">🎉 Giorno Festivo</option>
                    <option value="unificata">👥 Adunanza Unificata</option>
                    <option value="visita">👔 Visita Sorvegliante</option>
                    <option value="campagna">🚀 Campagna Speciale</option>
                    <option value="speciale">⭐ Straordinaria / Altro</option>
                    <option value="standard">📅 Standard</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="label">Orario Incontro *</label>
                  <input
                    type="time"
                    required
                    value={specialTime}
                    onChange={e => setSpecialTime(e.target.value)}
                    className="inp text-xs font-mono font-bold"
                  />
                </div>

                <div>
                  <label className="label">Fascia Oraria</label>
                  <select
                    value={specialTimeSlot}
                    onChange={e => setSpecialTimeSlot(e.target.value as any)}
                    className="inp text-xs"
                  >
                    <option value="mattina">🌅 Mattina</option>
                    <option value="pomeriggio">🌇 Pomeriggio</option>
                    <option value="sera">🌙 Sera</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="label">Luogo di Ritrovo *</label>
                <div className="flex gap-2">
                  <select
                    value={specialLocation}
                    onChange={e => setSpecialLocation(e.target.value)}
                    className="inp text-xs flex-1"
                  >
                    {locations.map(loc => (
                      <option key={loc} value={loc}>
                        {isZoomLocation(loc) ? `💻 ${loc} (Online)` : loc}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="label">Conduttore Assegnato (Opzionale)</label>
                <select
                  value={specialConductorId}
                  onChange={e => setSpecialConductorId(e.target.value)}
                  className="inp text-xs"
                >
                  <option value="">— Lascia da assegnare / Assegna con auto-generazione —</option>
                  {conductors.map(c => (
                    <option key={c.id} value={c.id}>
                      {c.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="label">Note / Dettaglio Evento</label>
                <input
                  type="text"
                  placeholder="es. Festa del Lavoro, Adunanza Unificata tutti i gruppi..."
                  value={specialNoteInput}
                  onChange={e => setSpecialNoteInput(e.target.value)}
                  className="inp text-xs"
                />
              </div>

              <div className="flex items-center justify-end gap-2.5 pt-3 border-t border-slate-100 dark:border-slate-800">
                <button
                  type="button"
                  onClick={() => setIsSpecialModalOpen(false)}
                  className="btn-ghost text-xs px-4 py-2"
                >
                  Annulla
                </button>
                <button
                  type="submit"
                  className="btn-primary bg-indigo-600 hover:bg-indigo-700 text-white text-xs px-4 py-2 flex items-center gap-1.5"
                >
                  <Check className="w-4 h-4" />
                  <span>Aggiungi Adunanza</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* SUBTAB 2: CONDUTTORI & DISPONIBILITÀ                    */}
      {/* ======================================================== */}
      {activeSubTab === 'conduttori' && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Form to Add / Edit Conductor */}
            <div className="card space-y-4 lg:col-span-1">
              <div className="flex items-center justify-between pb-3 border-b border-slate-100 dark:border-slate-800">
                <div className="flex items-center gap-2">
                  <UserPlus className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                  <h2 className="card-title">
                    {editingConductorId ? 'Modifica Conduttore' : 'Nuovo Conduttore'}
                  </h2>
                </div>
                {editingConductorId && (
                  <button
                    type="button"
                    onClick={() => {
                      setEditingConductorId(null);
                      setConductorName('');
                      setSelectedPersonId('');
                    }}
                    className="text-xs text-slate-500 hover:text-slate-800 dark:hover:text-slate-200"
                  >
                    Annulla
                  </button>
                )}
              </div>

              <form onSubmit={handleSaveConductor} className="space-y-4">
                {/* Pick from Congregation Anagrafica */}
                <div>
                  <label className="label">Collega da Anagrafica (Opzionale)</label>
                  <select
                    value={selectedPersonId}
                    disabled={!isAdmin}
                    onChange={e => {
                      const pid = e.target.value;
                      setSelectedPersonId(pid);
                      if (pid) {
                        const person = state.people.find(p => p.id === pid);
                        if (person) setConductorName(person.name);
                      }
                    }}
                    className="inp text-xs"
                  >
                    <option value="">-- Seleziona da Anagrafica Congregazione --</option>
                    {state.people.map(p => (
                      <option key={p.id} value={p.id}>
                        {p.name} ({p.gender === 'M' ? 'Fratello' : 'Sorella'})
                      </option>
                    ))}
                  </select>
                </div>

                {/* Name */}
                <div>
                  <label className="label">Nome e Cognome *</label>
                  <input
                    type="text"
                    required
                    disabled={!isAdmin}
                    placeholder="es. Mario Rossi"
                    value={conductorName}
                    onChange={e => setConductorName(e.target.value)}
                    className="inp text-xs font-semibold"
                  />
                </div>

                {/* Availability Checkboxes for the 4 standard slots */}
                <div className="space-y-2 pt-2 border-t border-slate-100 dark:border-slate-800">
                  <label className="label">Disponibilità Dichiarata per i Giorni</label>
                  <p className="text-[11px] text-slate-500 dark:text-slate-400">
                    Seleziona gli appuntamenti in cui questo conduttore può prestare servizio:
                  </p>

                  <div className="space-y-2 mt-2">
                    <label className="flex items-center gap-2 p-2.5 rounded-xl border border-sky-200 dark:border-sky-900 bg-sky-50/50 dark:bg-sky-950/20 cursor-pointer hover:bg-sky-50 transition-colors">
                      <input
                        type="checkbox"
                        disabled={!isAdmin}
                        checked={conductorAvailability.martediMattina}
                        onChange={e =>
                          setConductorAvailability({
                            ...conductorAvailability,
                            martediMattina: e.target.checked
                          })
                        }
                        className="rounded text-sky-600 focus:ring-sky-500"
                      />
                      <div className="text-xs">
                        <span className="font-bold text-sky-900 dark:text-sky-200">🌅 Martedì Mattina</span>
                        <span className="block text-[10px] text-slate-500">Ore {defaultSettings.martediMattina.time}</span>
                      </div>
                    </label>

                    <label className="flex items-center gap-2 p-2.5 rounded-xl border border-amber-200 dark:border-amber-900 bg-amber-50/50 dark:bg-amber-950/20 cursor-pointer hover:bg-amber-50 transition-colors">
                      <input
                        type="checkbox"
                        disabled={!isAdmin}
                        checked={conductorAvailability.giovediMattina}
                        onChange={e =>
                          setConductorAvailability({
                            ...conductorAvailability,
                            giovediMattina: e.target.checked
                          })
                        }
                        className="rounded text-amber-600 focus:ring-amber-500"
                      />
                      <div className="text-xs">
                        <span className="font-bold text-amber-900 dark:text-amber-200">☀️ Giovedì Mattina</span>
                        <span className="block text-[10px] text-slate-500">Ore {defaultSettings.giovediMattina.time}</span>
                      </div>
                    </label>

                    <label className="flex items-center gap-2 p-2.5 rounded-xl border border-indigo-200 dark:border-indigo-900 bg-indigo-50/50 dark:bg-indigo-950/20 cursor-pointer hover:bg-indigo-50 transition-colors">
                      <input
                        type="checkbox"
                        disabled={!isAdmin}
                        checked={conductorAvailability.sabatoPomeriggio}
                        onChange={e =>
                          setConductorAvailability({
                            ...conductorAvailability,
                            sabatoPomeriggio: e.target.checked
                          })
                        }
                        className="rounded text-indigo-600 focus:ring-indigo-500"
                      />
                      <div className="text-xs">
                        <span className="font-bold text-indigo-900 dark:text-indigo-200">🌇 Sabato Pomeriggio</span>
                        <span className="block text-[10px] text-slate-500">Ore {defaultSettings.sabatoPomeriggio.time}</span>
                      </div>
                    </label>

                    <label className="flex items-center gap-2 p-2.5 rounded-xl border border-emerald-200 dark:border-emerald-900 bg-emerald-50/50 dark:bg-emerald-950/20 cursor-pointer hover:bg-emerald-50 transition-colors">
                      <input
                        type="checkbox"
                        disabled={!isAdmin}
                        checked={conductorAvailability.domenicaPomeriggio}
                        onChange={e =>
                          setConductorAvailability({
                            ...conductorAvailability,
                            domenicaPomeriggio: e.target.checked
                          })
                        }
                        className="rounded text-emerald-600 focus:ring-emerald-500"
                      />
                      <div className="text-xs">
                        <span className="font-bold text-emerald-900 dark:text-emerald-200">🌆 Domenica Pomeriggio</span>
                        <span className="block text-[10px] text-slate-500">Ore {defaultSettings.domenicaPomeriggio.time}</span>
                      </div>
                    </label>
                  </div>
                </div>

                {/* Notes */}
                <div>
                  <label className="label">Note Particolari</label>
                  <input
                    type="text"
                    disabled={!isAdmin}
                    placeholder="es. Preferisce sabato, disponibile con auto"
                    value={conductorNotes}
                    onChange={e => setConductorNotes(e.target.value)}
                    className="inp text-xs"
                  />
                </div>

                <div className="pt-2">
                  <button type="submit" disabled={!isAdmin} className="btn-primary w-full flex items-center justify-center gap-2">
                    <Check className="w-4 h-4" />
                    <span>{editingConductorId ? 'Aggiorna Conduttore' : 'Aggiungi Conduttore'}</span>
                  </button>
                </div>
              </form>

              {/* Fast Import Button */}
              <div className="pt-3 border-t border-slate-100 dark:border-slate-800">
                <button
                  type="button"
                  disabled={!isAdmin}
                  onClick={handleImportBrothersFromAnagrafica}
                  className="btn-ghost w-full text-xs text-indigo-600 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-800 flex items-center justify-center gap-1.5"
                >
                  <Users className="w-3.5 h-3.5" />
                  <span>Importa tutti i Fratelli dall'Anagrafica</span>
                </button>
              </div>
            </div>

            {/* List of Conductors */}
            <div className="card space-y-4 lg:col-span-2">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100 dark:border-slate-800">
                <div className="flex items-center gap-2">
                  <Users className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                  <h2 className="card-title">
                    Elenco Conduttori Registrati ({conductors.length})
                  </h2>
                </div>

                <div className="relative w-full sm:w-64">
                  <input
                    type="text"
                    placeholder="Cerca conduttore..."
                    value={conductorSearch}
                    onChange={e => setConductorSearch(e.target.value)}
                    className="inp pl-8 text-xs py-1.5"
                  />
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                </div>
              </div>

              {conductors.length === 0 ? (
                <div className="text-center py-12 text-slate-400 text-xs bg-slate-50/50 dark:bg-slate-900/30 rounded-2xl border border-dashed border-slate-200 dark:border-slate-800">
                  Nessun conduttore registrato. Aggiungi il primo conduttore dal modulo o importali dall'anagrafica.
                </div>
              ) : (
                <div className="overflow-x-auto rounded-2xl border border-slate-200 dark:border-slate-800">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100/90 dark:bg-slate-800/90 text-slate-700 dark:text-slate-300 font-semibold border-b border-slate-200 dark:border-slate-700">
                        <th className="p-3">Conduttore</th>
                        <th className="p-3 text-center">🌅 Mar Mattina</th>
                        <th className="p-3 text-center">☀️ Gio Mattina</th>
                        <th className="p-3 text-center">🌇 Sab Pom</th>
                        <th className="p-3 text-center">🌆 Dom Pom</th>
                        <th className="p-3">Note</th>
                        <th className="p-3 text-right">Azioni</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                      {conductors
                        .filter(c => !conductorSearch.trim() || c.name.toLowerCase().includes(conductorSearch.toLowerCase()))
                        .map(c => (
                          <tr key={c.id} className="hover:bg-slate-50/80 dark:hover:bg-slate-800/40">
                            <td className="p-3 font-semibold text-slate-900 dark:text-slate-100">
                              <div className="flex items-center gap-2">
                                <span>{c.name}</span>
                                {c.personId && (
                                  <span className="px-1.5 py-0.2 rounded bg-indigo-100 dark:bg-indigo-950 text-indigo-700 dark:text-indigo-300 text-[10px]" title="Collegato all'anagrafica">
                                    Anagrafica
                                  </span>
                                )}
                              </div>
                            </td>

                            <td className="p-3 text-center">
                              {c.availability.martediMattina ? (
                                <span className="inline-block p-1 bg-sky-100 text-sky-700 dark:bg-sky-950 dark:text-sky-300 rounded-md font-bold">✓</span>
                              ) : (
                                <span className="text-slate-300 dark:text-slate-700">—</span>
                              )}
                            </td>

                            <td className="p-3 text-center">
                              {c.availability.giovediMattina ? (
                                <span className="inline-block p-1 bg-amber-100 text-amber-700 dark:bg-amber-950 dark:text-amber-300 rounded-md font-bold">✓</span>
                              ) : (
                                <span className="text-slate-300 dark:text-slate-700">—</span>
                              )}
                            </td>

                            <td className="p-3 text-center">
                              {c.availability.sabatoPomeriggio ? (
                                <span className="inline-block p-1 bg-indigo-100 text-indigo-700 dark:bg-indigo-950 dark:text-indigo-300 rounded-md font-bold">✓</span>
                              ) : (
                                <span className="text-slate-300 dark:text-slate-700">—</span>
                              )}
                            </td>

                            <td className="p-3 text-center">
                              {c.availability.domenicaPomeriggio ? (
                                <span className="inline-block p-1 bg-emerald-100 text-emerald-700 dark:bg-emerald-950 dark:text-emerald-300 rounded-md font-bold">✓</span>
                              ) : (
                                <span className="text-slate-300 dark:text-slate-700">—</span>
                              )}
                            </td>

                            <td className="p-3 text-slate-500 text-[11px]">
                              {c.notes || '—'}
                            </td>

                            <td className="p-3 text-right">
                              <div className="flex items-center justify-end gap-1">
                                <button
                                  type="button"
                                  disabled={!isAdmin}
                                  onClick={() => handleEditConductor(c)}
                                  className="p-1 text-slate-500 hover:text-indigo-600 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800"
                                  title="Modifica"
                                >
                                  <Edit2 className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  disabled={!isAdmin}
                                  onClick={() => handleDeleteConductor(c.id, c.name)}
                                  className="p-1 text-slate-500 hover:text-rose-600 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800"
                                  title="Elimina"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* SUBTAB 3: LUOGHI & ORARI PREDEFINITI                    */}
      {/* ======================================================== */}
      {activeSubTab === 'luoghi' && (
        <div className="space-y-6">
          {/* Default Slot Settings (Times & Default Place) */}
          <div className="card space-y-4">
            <div className="flex items-center gap-2 pb-3 border-b border-slate-100 dark:border-slate-800">
              <Settings2 className="w-5 h-5 text-amber-600 dark:text-amber-400" />
              <div>
                <h2 className="card-title">Orari e Gestione Luoghi dei 4 Appuntamenti</h2>
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                  Configura gli orari standard e i luoghi per ciascun appuntamento. È possibile scegliere la rotazione automatica dei luoghi fisici o impostare un luogo fisso (inclusi <strong>Zoom</strong> o la Sala).
                </p>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              {(['martediMattina', 'giovediMattina', 'sabatoPomeriggio', 'domenicaPomeriggio'] as ServizioCampoDayKey[]).map(slotKey => {
                const info = SLOT_INFO[slotKey];
                const conf = defaultSettings[slotKey] || DEFAULT_SLOT_SETTINGS[slotKey];
                const isSunday = slotKey === 'domenicaPomeriggio';
                const isConfiguredZoom = isZoomLocation(conf.defaultLocation) || (isSunday && (!conf.defaultLocation || isZoomLocation(conf.defaultLocation)));

                return (
                  <div key={slotKey} className="p-4 rounded-2xl border border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/40 space-y-3">
                    <div className="flex items-center justify-between font-bold text-xs">
                      <div className="flex items-center gap-2">
                        <info.icon className="w-4 h-4 text-amber-600" />
                        <span>{info.label}</span>
                      </div>
                      {isConfiguredZoom ? (
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300">
                          💻 Zoom
                        </span>
                      ) : conf.defaultLocation && conf.defaultLocation !== 'rotazione' ? (
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800 dark:bg-amber-950 dark:text-amber-300">
                          📍 Fisso
                        </span>
                      ) : (
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-sky-100 text-sky-800 dark:bg-sky-950 dark:text-sky-300">
                          🔄 A Rotazione
                        </span>
                      )}
                    </div>

                    <div>
                      <label className="text-[11px] font-semibold text-slate-500">Orario Standard</label>
                      <input
                        type="time"
                        disabled={!isAdmin}
                        value={conf.time}
                        onChange={e => handleSaveSlotSetting(slotKey, { time: e.target.value })}
                        className="inp py-1 px-2 text-xs font-mono font-bold mt-1"
                      />
                    </div>

                    <div>
                      <label className="text-[11px] font-semibold text-slate-500">
                        {isSunday ? 'Luogo Predefinito (Domenica)' : 'Luogo Predefinito / Rotazione'}
                      </label>
                      <select
                        disabled={!isAdmin}
                        value={isKnownLocation(conf.defaultLocation) ? conf.defaultLocation : (isSunday ? sundayFallbackLocation : 'rotazione')}
                        onChange={e => handleSaveSlotSetting(slotKey, { defaultLocation: e.target.value })}
                        className={`inp py-1 px-2 text-xs mt-1 font-semibold ${
                          isConfiguredZoom
                            ? 'bg-emerald-50/50 dark:bg-emerald-950/30 text-emerald-900 dark:text-emerald-200 border-emerald-300 dark:border-emerald-800'
                            : ''
                        }`}
                      >
                        {!isSunday && (
                          <option value="rotazione">🔄 A rotazione tra i luoghi fisici</option>
                        )}
                        {locations.map(loc => (
                          <option key={loc} value={loc}>
                            {isZoomLocation(loc) ? `💻 ${loc}` : loc}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Location Management List */}
          <div className="card space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100 dark:border-slate-800">
              <div className="flex items-center gap-2">
                <MapPin className="w-5 h-5 text-amber-600 dark:text-amber-400" />
                <div>
                  <h2 className="card-title">Elenco Luoghi di Ritrovo ({locations.length})</h2>
                  <p className="text-xs text-slate-500">I luoghi registrati possono essere selezionati per qualsiasi appuntamento o utilizzati nella rotazione automatica.</p>
                </div>
              </div>

              {/* Add New Location Form */}
              <form onSubmit={handleAddLocation} className="flex items-center gap-2">
                <input
                  type="text"
                  disabled={!isAdmin}
                  placeholder="Nome nuovo luogo..."
                  value={newLocationName}
                  onChange={e => setNewLocationName(e.target.value)}
                  className="inp text-xs py-1.5 w-64"
                />
                <button type="submit" disabled={!isAdmin || !newLocationName.trim()} className="btn-primary text-xs py-1.5 flex items-center gap-1">
                  <Plus className="w-3.5 h-3.5" />
                  <span>Aggiungi</span>
                </button>
              </form>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
              {locations.map((loc, idx) => {
                const isZoom = isZoomLocation(loc);
                return (
                  <div
                    key={idx}
                    className={`p-3 rounded-xl border flex items-center justify-between gap-2 shadow-2xs ${
                      isZoom
                        ? 'border-emerald-200 dark:border-emerald-900 bg-emerald-50/40 dark:bg-emerald-950/20'
                        : 'border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900'
                    }`}
                  >
                    {editingLocationIndex === idx ? (
                      <div className="flex items-center gap-1.5 w-full">
                        <input
                          type="text"
                          autoFocus
                          value={editLocationValue}
                          onChange={e => setEditLocationValue(e.target.value)}
                          className="inp text-xs py-1 px-2 flex-1"
                        />
                        <button
                          type="button"
                          onClick={() => handleSaveEditLocation(idx)}
                          className="p-1 text-emerald-600 hover:bg-emerald-50 rounded"
                          title="Salva"
                        >
                          <Check className="w-3.5 h-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => setEditingLocationIndex(null)}
                          className="p-1 text-slate-400 hover:bg-slate-100 rounded"
                          title="Annulla"
                        >
                          ✕
                        </button>
                      </div>
                    ) : (
                      <>
                        <div className="flex flex-col min-w-0 flex-1">
                          <div className="flex items-center gap-2 truncate">
                            {isZoom ? (
                              <span className="text-sm shrink-0">💻</span>
                            ) : (
                              <MapPin className="w-4 h-4 text-amber-500 shrink-0" />
                            )}
                            <span className="text-xs font-semibold text-slate-800 dark:text-slate-200 truncate" title={loc}>
                              {loc}
                            </span>
                          </div>
                          <span className="text-[10px] text-slate-500 dark:text-slate-400 mt-0.5 ml-6 truncate">
                            {isZoom ? 'Online (Zoom)' : 'In Rotazione fisica'}
                          </span>
                        </div>
                        <div className="flex items-center gap-1 shrink-0">
                          <button
                            type="button"
                            disabled={!isAdmin}
                            onClick={() => {
                              setEditingLocationIndex(idx);
                              setEditLocationValue(loc);
                            }}
                            className="p-1 text-slate-400 hover:text-indigo-600 rounded"
                            title="Modifica luogo"
                          >
                            <Edit2 className="w-3 h-3" />
                          </button>
                          <button
                            type="button"
                            disabled={!isAdmin || locations.length <= 1}
                            onClick={() => handleDeleteLocation(loc)}
                            className="p-1 text-slate-400 hover:text-rose-600 rounded"
                            title="Elimina luogo"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* SUBTAB 4: RIEPILOGO UTILIZZO & STATISTICHE              */}
      {/* ======================================================== */}
      {activeSubTab === 'statistiche' && (
        <div className="space-y-6">
          {/* Quick Metrics */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            <div className="card p-4 border-l-4 border-l-sky-500">
              <span className="text-[11px] font-semibold text-slate-400">Conduttori Impiegati nel Mese</span>
              <div className="text-2xl font-black text-slate-900 dark:text-slate-100 mt-1">
                {activeConductorsInMonth} <span className="text-xs font-normal text-slate-500">/ {conductors.length}</span>
              </div>
              <div className="text-xs text-slate-500 mt-1">
                Copertura: {conductors.length > 0 ? Math.round((activeConductorsInMonth / conductors.length) * 100) : 0}%
              </div>
            </div>

            <div className="card p-4 border-l-4 border-l-indigo-500">
              <span className="text-[11px] font-semibold text-slate-400">Adunanze Totali Assegnate</span>
              <div className="text-2xl font-black text-indigo-600 dark:text-indigo-400 mt-1">
                {assignedMonthMeetings} <span className="text-xs font-normal text-slate-500">/ {totalMonthMeetings}</span>
              </div>
              <div className="text-xs text-slate-500 mt-1">
                Nel mese di {MESI[selectedMonth]} {selectedYear}
              </div>
            </div>

            <div className="card p-4 border-l-4 border-l-emerald-500">
              <span className="text-[11px] font-semibold text-slate-400">Media Incarichi</span>
              <div className="text-2xl font-black text-emerald-600 dark:text-emerald-400 mt-1">
                {activeConductorsInMonth > 0 ? (assignedMonthMeetings / activeConductorsInMonth).toFixed(1) : '0'}
              </div>
              <div className="text-xs text-slate-500 mt-1">
                Adunanze per conduttore attivo
              </div>
            </div>
          </div>

          {/* Usage Table */}
          <div className="card space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100 dark:border-slate-800">
              <div className="flex items-center gap-2">
                <BarChart2 className="w-5 h-5 text-purple-600 dark:text-purple-400" />
                <h2 className="card-title">
                  Statistiche di Utilizzo Conduttori ({MESI[selectedMonth]} {selectedYear})
                </h2>
              </div>
            </div>

            <div className="overflow-x-auto rounded-2xl border border-slate-200 dark:border-slate-800">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100/90 dark:bg-slate-800/90 text-slate-700 dark:text-slate-300 font-semibold border-b border-slate-200 dark:border-slate-700">
                    <th className="p-3">Conduttore</th>
                    <th className="p-3 text-center">Turni nel Mese</th>
                    <th className="p-3 text-center">🌅 Mar</th>
                    <th className="p-3 text-center">☀️ Gio</th>
                    <th className="p-3 text-center">🌇 Sab</th>
                    <th className="p-3 text-center">🌆 Dom</th>
                    <th className="p-3 text-center">⭐ Festivi/Spec.</th>
                    <th className="p-3">Date nel Mese</th>
                    <th className="p-3 text-center">Totale Storico</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                  {conductorStats
                    .slice()
                    .sort((a, b) => b.monthTotal - a.monthTotal || a.name.localeCompare(b.name))
                    .map(c => (
                      <tr key={c.id} className="hover:bg-slate-50/80 dark:hover:bg-slate-800/40">
                        <td className="p-3 font-semibold text-slate-900 dark:text-slate-100">
                          {c.name}
                        </td>
                        <td className="p-3 text-center">
                          <span
                            className={`inline-block px-2.5 py-0.5 rounded-full text-xs font-bold ${
                              c.monthTotal === 0
                                ? 'bg-slate-100 dark:bg-slate-800 text-slate-500'
                                : c.monthTotal <= 2
                                ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300'
                                : 'bg-indigo-100 text-indigo-800 dark:bg-indigo-950 dark:text-indigo-300'
                            }`}
                          >
                            {c.monthTotal} {c.monthTotal === 1 ? 'turno' : 'turni'}
                          </span>
                        </td>
                        <td className="p-3 text-center">{c.monthMartedi || '—'}</td>
                        <td className="p-3 text-center">{c.monthGiovedi || '—'}</td>
                        <td className="p-3 text-center">{c.monthSabato || '—'}</td>
                        <td className="p-3 text-center">{c.monthDomenica || '—'}</td>
                        <td className="p-3 text-center">{c.monthSpeciali || '—'}</td>
                        <td className="p-3">
                          {c.monthAssignments.length === 0 ? (
                            <span className="text-slate-400 italic text-[11px]">Nessun turno</span>
                          ) : (
                            <div className="flex flex-wrap gap-1">
                              {c.monthAssignments.map((a, idx) => (
                                <span
                                  key={idx}
                                  className="px-1.5 py-0.5 rounded bg-slate-100 dark:bg-slate-800 text-[10px] font-medium"
                                  title={`${a.dayLabel} - ${a.location}`}
                                >
                                  {a.dayLabel}
                                </span>
                              ))}
                            </div>
                          )}
                        </td>
                        <td className="p-3 text-center font-semibold text-slate-800 dark:text-slate-200">
                          {c.allTimeTotal}
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* SUBTAB 5: STAMPA & CONDIVISIONE                         */}
      {/* ======================================================== */}
      {activeSubTab === 'stampa' && (
        <div className="space-y-6">
          <div className="card flex items-center justify-between no-print">
            <div className="flex items-center gap-2">
              <Printer className="w-5 h-5 text-emerald-600 dark:text-emerald-400" />
              <div>
                <h2 className="card-title">Foglio di Stampa Programma Servizio di Campo</h2>
                <p className="text-xs text-slate-500">Ottimizzato per la bacheca e la stampa cartacea (o esportazione PDF)</p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => window.print()}
                className="btn-primary flex items-center gap-1.5"
              >
                <Printer className="w-4 h-4" />
                <span>Stampa Foglio</span>
              </button>
              <button
                type="button"
                onClick={handleCopyWhatsApp}
                className="btn-ghost text-xs flex items-center gap-1.5 text-emerald-600 border border-emerald-300"
              >
                <Share2 className="w-3.5 h-3.5" />
                <span>WhatsApp</span>
              </button>
            </div>
          </div>

          {/* Printable Layout Sheet */}
          <div className="no-print mb-4"><label className="text-sm font-semibold">Cosa trattare?<input className="input w-full mt-1" value={data.printTopic || ''} onChange={e => saveServizioData({ ...data, printTopic: e.target.value })} /></label></div>
          <ServizioCampoPrintTemplate meetings={currentMonthSchedule.map(m => ({ ...m, conductorName: getConductorById(m.conductorId)?.name || '—' }))} year={selectedYear} month={selectedMonth} topic={data.printTopic || DEFAULT_FIELD_TOPIC} defaults={defaultSettings} />
        </div>
      )}

      {/* ======================================================== */}
      {/* IN-APP CONFIRMATION MODALS (FOR IFRAME RELIABILITY)      */}
      {/* ======================================================== */}
      {/* 1. Delete Conductor Modal */}
      {conductorToDelete && (
        <div className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in duration-150">
          <div className="card max-w-md w-full p-6 space-y-4 border-rose-200 dark:border-rose-900/50 shadow-2xl">
            <div className="flex items-center gap-3 text-rose-600 dark:text-rose-400">
              <div className="p-2 bg-rose-100 dark:bg-rose-950 rounded-xl">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="font-bold text-base text-slate-900 dark:text-slate-100">Elimina Conduttore</h3>
                <p className="text-xs text-slate-500">Questa operazione è irreversibile.</p>
              </div>
            </div>

            <p className="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
              Sei sicuro di voler eliminare il conduttore <strong>"{conductorToDelete.name}"</strong> dall'elenco del servizio di campo? Gli eventuali turni già assegnati a questo conduttore torneranno da assegnare.
            </p>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100 dark:border-slate-800">
              <button
                type="button"
                onClick={() => setConductorToDelete(null)}
                className="btn-ghost text-xs"
              >
                Annulla
              </button>
              <button
                type="button"
                onClick={executeDeleteConductor}
                className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs shadow-md transition-colors"
              >
                Elimina Conduttore
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 2. Delete Meeting Modal */}
      {meetingToDelete && (
        <div className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in duration-150">
          <div className="card max-w-md w-full p-6 space-y-4 border-rose-200 dark:border-rose-900/50 shadow-2xl">
            <div className="flex items-center gap-3 text-rose-600 dark:text-rose-400">
              <div className="p-2 bg-rose-100 dark:bg-rose-950 rounded-xl">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="font-bold text-base text-slate-900 dark:text-slate-100">Elimina Adunanza</h3>
                <p className="text-xs text-slate-500">Rimuove l'adunanza dal calendario</p>
              </div>
            </div>

            <p className="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
              Sei sicuro di voler rimuovere l'adunanza del <strong>{formatShortDate(meetingToDelete.dateStr)}</strong> ({meetingToDelete.time} - {meetingToDelete.location})?
            </p>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100 dark:border-slate-800">
              <button
                type="button"
                onClick={() => setMeetingToDelete(null)}
                className="btn-ghost text-xs"
              >
                Annulla
              </button>
              <button
                type="button"
                onClick={executeDeleteMeeting}
                className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs shadow-md transition-colors"
              >
                Elimina Adunanza
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 3. Delete Location Modal */}
      {locationToDelete && (
        <div className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in duration-150">
          <div className="card max-w-md w-full p-6 space-y-4 border-rose-200 dark:border-rose-900/50 shadow-2xl">
            <div className="flex items-center gap-3 text-rose-600 dark:text-rose-400">
              <div className="p-2 bg-rose-100 dark:bg-rose-950 rounded-xl">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="font-bold text-base text-slate-900 dark:text-slate-100">Rimuovi Luogo di Ritrovo</h3>
                <p className="text-xs text-slate-500">Elimina il punto di ritrovo dalla rotazione</p>
              </div>
            </div>

            <p className="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
              Sei sicuro di voler rimuovere il luogo <strong>"{locationToDelete}"</strong> dall'elenco dei luoghi di ritrovo per il servizio di campo?
            </p>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100 dark:border-slate-800">
              <button
                type="button"
                onClick={() => setLocationToDelete(null)}
                className="btn-ghost text-xs"
              >
                Annulla
              </button>
              <button
                type="button"
                onClick={executeDeleteLocation}
                className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs shadow-md transition-colors"
              >
                Rimuovi Luogo
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
