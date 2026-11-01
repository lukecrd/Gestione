import { generateVitaAssignments, vitaMonthlyRepeats } from '../utils/automaticScheduling';
import { fetchWolJson } from '../utils/wolApi';
import React, { useState } from 'react';
import {
  SchedulingPrograms,
  StateData,
  VitaEMinisteroData,
  VitaEMinisteroMeeting,
  VitaEMinisteroParticipant,
  MinisteroPart,
  VitaCristianaPart,
} from '../types';
import {
  DEFAULT_VITA_MINISTERO_DATA,
  DEFAULT_VITA_MINISTERO_PARTICIPANTS,
  DEFAULT_VITA_MINISTERO_MEETINGS,
} from '../data/defaultVitaEMinistero';
import { VitaEMinisteroPrintView } from './VitaEMinisteroPrintView';
import { VitaEMinisteroParticipantsModal } from './VitaEMinisteroParticipantsModal';
import { VitaEMinisteroWolModal } from './VitaEMinisteroWolModal';
import { VitaEMinisteroStatsView } from './VitaEMinisteroStatsView';
import { S89Modal } from './S89Modal';
import {
  S89Item,
  downloadS89Pdf,
  downloadCombinedS89Pdf,
  getS89FileName,
} from '../utils/s89Pdf';
import {
  computeVitaStats,
  compareParticipantsByLastAssignmentAsc,
} from '../utils/vitaEMinisteroStats';
import {
  getEffectivePartTypes,
  inferPartTypeIds,
  isParticipantEligibleForPartTypes,
  slugifyPartTypeLabel,
} from '../utils/ministeroPartTypes';
import {
  Calendar,
  Users,
  Printer,
  Plus,
  Trash2,
  ChevronLeft,
  ChevronRight,
  BookOpen,
  Sparkles,
  Building2,
  Music,
  CheckCircle2,
  AlertTriangle,
  Globe,
  RefreshCw,
  Loader2,
  Download,
  BarChart3,
  FileText,
  Clock,
} from 'lucide-react';

interface VitaEMinisteroViewProps {
  state: StateData;
  activePrograms?: SchedulingPrograms;
  onSaveState: (newState: StateData) => void;
  isAdmin: boolean;
  onShowToast: (msg: string) => void;
  checkAdminPermission: () => boolean;
  stampa?: boolean; // apre direttamente la sottoscheda di stampa (PDF della Tabella informazioni)
}

const MESI_FULL = [
  'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
  'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
];

export function VitaEMinisteroView({
  state,
  activePrograms,
  onSaveState,
  isAdmin,
  onShowToast,
  checkAdminPermission,
  stampa,
}: VitaEMinisteroViewProps) {
  // Current active sub-tab
  const [activeSubTab, setActiveSubTab] = useState<'programma' | 'nominativi' | 'statistiche' | 'stampa'>(stampa ? 'stampa' : 'programma');

  const [generationWarnings, setGenerationWarnings] = useState<string[]>([]);

  // WOL sync modal state
  const [isWolModalOpen, setIsWolModalOpen] = useState<boolean>(false);
  const [isQuickSyncing, setIsQuickSyncing] = useState<boolean>(false);

  // Active S-89 Assignment Slip Modal
  const [activeS89Modal, setActiveS89Modal] = useState<{
    item: S89Item;
    meetingId: string;
    targetType: 'lettura' | 'ministeroPart';
    partIdx?: number;
  } | null>(null);

  // Month & Year state: current month
  const [selectedMonth, setSelectedMonth] = useState<number>(() => new Date().getMonth());
  const [selectedYear, setSelectedYear] = useState<number>(() => new Date().getFullYear());

  // Get or initialize Vita e Ministero data
  const vmData: VitaEMinisteroData = React.useMemo(() => {
    if (state.vitaEMinistero) {
      return {
        congregationName: state.vitaEMinistero.congregationName || DEFAULT_VITA_MINISTERO_DATA.congregationName,
        participants: state.vitaEMinistero.participants ?? DEFAULT_VITA_MINISTERO_PARTICIPANTS,
        meetings: state.vitaEMinistero.meetings ?? DEFAULT_VITA_MINISTERO_MEETINGS,
        ministeroPartTypes: state.vitaEMinistero.ministeroPartTypes ?? [],
      };
    }
    return DEFAULT_VITA_MINISTERO_DATA;
  }, [state.vitaEMinistero]);

  const participants = vmData.participants || DEFAULT_VITA_MINISTERO_PARTICIPANTS;
  const meetings = vmData.meetings || [];

  // Tipi di parte "Efficaci nel ministero": lista fissa + eventuali aggiunte custom
  const ministeroPartTypes = React.useMemo(
    () => getEffectivePartTypes(vmData.ministeroPartTypes),
    [vmData.ministeroPartTypes]
  );

  // Compute stats Map for rotation priority and last assignment tracking
  const statsMap = React.useMemo(() => {
    return computeVitaStats(participants, meetings);
  }, [participants, meetings]);

  const updateVmData = (newData: VitaEMinisteroData) => {
    onSaveState({
      ...state,
      vitaEMinistero: newData,
    });
  };

  const handleUpdateParticipants = (newParticipants: VitaEMinisteroParticipant[]) => {
    updateVmData({
      ...vmData,
      participants: newParticipants,
    });
  };

  // Aggiunge un nuovo tipo di parte personalizzato alla lista fissa
  const handleAddCustomPartType = (label: string) => {
    const trimmed = label.trim();
    if (!trimmed) return;
    const alreadyExists = ministeroPartTypes.some(
      t => t.label.trim().toLowerCase() === trimmed.toLowerCase()
    );
    if (alreadyExists) {
      onShowToast('Esiste già un tipo di parte con questo nome.');
      return;
    }
    const newType = { id: slugifyPartTypeLabel(trimmed), label: trimmed, isDefault: false };
    updateVmData({
      ...vmData,
      ministeroPartTypes: [...(vmData.ministeroPartTypes || []), newType],
    });
    onShowToast(`Aggiunto tipo di parte "${trimmed}"`);
  };

  // Rimuove un tipo di parte personalizzato (non è possibile rimuovere quelli predefiniti)
  // e ripulisce i riferimenti da partecipanti e parti già configurate.
  const handleRemoveCustomPartType = (id: string) => {
    const cleanedParticipants = participants.map(p =>
      p.roles.ministeroTipiAbilitati?.includes(id)
        ? { ...p, roles: { ...p.roles, ministeroTipiAbilitati: p.roles.ministeroTipiAbilitati!.filter(t => t !== id) } }
        : p
    );
    const cleanedMeetings = meetings.map(m =>
      m.ministeroParts.some(mp => mp.partTypeIds?.includes(id))
        ? {
            ...m,
            ministeroParts: m.ministeroParts.map(mp =>
              mp.partTypeIds?.includes(id)
                ? { ...mp, partTypeIds: mp.partTypeIds.filter(t => t !== id) }
                : mp
            ),
          }
        : m
    );
    updateVmData({
      ...vmData,
      participants: cleanedParticipants,
      meetings: cleanedMeetings,
      ministeroPartTypes: (vmData.ministeroPartTypes || []).filter(t => t.id !== id),
    });
    onShowToast('Tipo di parte rimosso');
  };

  const handleUpdateMeetings = (newMeetings: VitaEMinisteroMeeting[]) => {
    updateVmData({
      ...vmData,
      meetings: newMeetings,
    });
  };

  // Filter meetings for selected month & year
  const meetingsInMonth = meetings.filter(m => {
    if (!m.dateStr) return true;
    const parts = m.dateStr.split('-');
    if (parts.length >= 2) {
      const y = parseInt(parts[0], 10);
      const mIdx = parseInt(parts[1], 10) - 1;
      return y === selectedYear && mIdx === selectedMonth;
    }
    return true;
  });
  const monthlyRepeats = vitaMonthlyRepeats(meetingsInMonth, participants);

  const handleAutoGenerate = () => {
    if (!checkAdminPermission()) return;
    const result = generateVitaAssignments(state, vmData, selectedYear, selectedMonth, activePrograms);
    setGenerationWarnings(result.warnings);
    if (!result.generatedCount) {
      onShowToast('Nessuna adunanza ordinaria da generare nel mese selezionato. Scarica prima il programma da WOL.');
      return;
    }
    handleUpdateMeetings(result.meetings);
    onShowToast(result.warnings.length
      ? `Generazione completata: ${result.warnings.length} segnalazioni da verificare.`
      : 'Nominativi assegnati in ordine di ultima assegnazione, senza sovrapposizioni nella stessa data.');
  };

  const handlePrevMonth = () => {
    if (selectedMonth === 0) {
      setSelectedMonth(11);
      setSelectedYear(y => y - 1);
    } else {
      setSelectedMonth(m => m - 1);
    }
  };

  const handleNextMonth = () => {
    if (selectedMonth === 11) {
      setSelectedMonth(0);
      setSelectedYear(y => y + 1);
    } else {
      setSelectedMonth(m => m + 1);
    }
  };

  // Quick sync currently selected month from wol.jw.org
  const handleQuickSyncCurrentMonth = async () => {
    if (!checkAdminPermission()) return;
    setIsQuickSyncing(true);
    try {
      const monthNum = selectedMonth + 1;
      const data = await fetchWolJson(`/api/wol/month?year=${selectedYear}&month=${monthNum}`);
      if (!data.success) {
        throw new Error(data.error || 'Errore durante la connessione a wol.jw.org');
      }
      const fetched: VitaEMinisteroMeeting[] = data.meetings || [];
      if (fetched.length === 0) {
        onShowToast(`Nessuna adunanza trovata su wol.jw.org per ${MESI_FULL[selectedMonth]} ${selectedYear}`);
        return;
      }

      // Merge preserving existing participants
      const meetingMap = new Map<string, VitaEMinisteroMeeting>();
      meetings.forEach(m => {
        if (m.dateStr) meetingMap.set(m.dateStr, m);
      });

      fetched.forEach(wolM => {
        const existing = wolM.dateStr ? meetingMap.get(wolM.dateStr) : undefined;
        if (!existing) {
          meetingMap.set(wolM.dateStr, wolM);
        } else {
          const mergedMinisteroParts = wolM.ministeroParts.map((wp, idx) => {
            const ep = existing.ministeroParts.find(p => p.number === wp.number) || existing.ministeroParts[idx];
            return {
              ...wp,
              hasAssistant: ep?.hasAssistant ?? wp.hasAssistant,
              studentId: ep?.studentId || '',
              assistantId: (ep?.hasAssistant ?? wp.hasAssistant) ? ep?.assistantId || '' : '',
              isSent: ep?.isSent || false,
              room: ep?.room || 'main',
              // Mantiene i tipi di parte già impostati manualmente in precedenza,
              // altrimenti usa quelli dedotti automaticamente dallo scraper (wp.partTypeIds).
              partTypeIds: ep?.partTypeIds && ep.partTypeIds.length > 0
                ? ep.partTypeIds
                : (wp.partTypeIds || inferPartTypeIds(wp.minutes, wp.hasAssistant)),
            };
          });

          const mergedVitaParts = wolM.vitaCristianaParts.map((wvp, idx) => {
            const evp = existing.vitaCristianaParts.find(p => p.number === wvp.number) || existing.vitaCristianaParts[idx];
            return {
              ...wvp,
              speakerId: evp?.speakerId || '',
            };
          });

          meetingMap.set(wolM.dateStr, {
            ...wolM,
            id: existing.id || wolM.id,
            presidenteId: existing.presidenteId || '',
            preghieraInizialeId: existing.preghieraInizialeId || '',
            tesori1SpeakerId: existing.tesori1SpeakerId || '',
            tesoriGemmeSpeakerId: existing.tesoriGemmeSpeakerId || '',
            tesoriLetturaReaderId: existing.tesoriLetturaReaderId || '',
            tesoriLetturaSent: existing.tesoriLetturaSent || false,
            tesoriLetturaRoom: existing.tesoriLetturaRoom || 'main',
            ministeroParts: mergedMinisteroParts,
            vitaCristianaParts: mergedVitaParts,
            studioBiblicoType: existing.studioBiblicoType,
            discorsoSorveglianteTitle: existing.discorsoSorveglianteTitle,
            studioBiblicoConductorId: existing.studioBiblicoConductorId || '',
            studioBiblicoReaderId: existing.studioBiblicoReaderId || '',
            preghieraFinaleId: existing.preghieraFinaleId || '',
            isSpecialEvent: existing.isSpecialEvent,
            specialEventTitle: existing.specialEventTitle,
            notes: existing.notes,
          });
        }
      });

      const updatedMeetings = Array.from(meetingMap.values());
      updatedMeetings.sort((a, b) => (a.dateStr || '').localeCompare(b.dateStr || ''));
      handleUpdateMeetings(updatedMeetings);
      onShowToast(`Aggiornato ${MESI_FULL[selectedMonth]} ${selectedYear} da wol.jw.org (${fetched.length} adunanze)!`);
    } catch (err: any) {
      console.error('Quick sync error:', err);
      onShowToast(`Errore: ${err.message}`);
    } finally {
      setIsQuickSyncing(false);
    }
  };

  // Add a new meeting week
  const handleAddMeeting = () => {
    if (!checkAdminPermission()) return;

    // Calculate a default date in selected month
    const d = new Date(selectedYear, selectedMonth, 1);
    // Find first Monday
    while (d.getDay() !== 1 && d.getMonth() === selectedMonth) {
      d.setDate(d.getDate() + 1);
    }
    const isoDate = `${selectedYear}-${String(selectedMonth + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const dateLabel = `${d.getDate()} ${MESI_FULL[selectedMonth].toLowerCase()} ${selectedYear}`;

    const newMeeting: VitaEMinisteroMeeting = {
      id: 'vm_m_' + Math.random().toString(36).substring(2, 9),
      dateStr: isoDate,
      dateLabel,
      bibleReading: 'LETTURA BIBLICA',
      presidenteId: '',
      canticoIniziale: 'Cantico 1: ',
      preghieraInizialeId: '',
      tesori1Title: 'Discorso',
      tesori1Minutes: 10,
      tesori1SpeakerId: '',
      tesoriGemmeTitle: 'Gemme spirituali',
      tesoriGemmeMinutes: 10,
      tesoriGemmeSpeakerId: '',
      tesoriLetturaTitle: 'Lettura biblica',
      tesoriLetturaMinutes: 4,
      tesoriLetturaReaderId: '',
      ministeroParts: [
        {
          id: 'mp_' + Math.random().toString(36).substring(2, 7),
          number: 4,
          title: 'Iniziare una conversazione',
          minutes: 3,
          studentId: '',
          assistantId: '',
          hasAssistant: true,
          partTypeIds: inferPartTypeIds(3, true),
        },
        {
          id: 'mp_' + Math.random().toString(36).substring(2, 7),
          number: 5,
          title: 'Coltivare l’interesse',
          minutes: 4,
          studentId: '',
          assistantId: '',
          hasAssistant: true,
          partTypeIds: inferPartTypeIds(4, true),
        },
      ],
      canticoIntermedio: 'Cantico 2: ',
      vitaCristianaParts: [
        {
          id: 'vcp_' + Math.random().toString(36).substring(2, 7),
          number: 6,
          title: 'Bisogni locali',
          minutes: 15,
          speakerId: '',
        },
      ],
      studioBiblicoTitle: 'Studio biblico di congregazione',
      studioBiblicoMinutes: 30,
      studioBiblicoConductorId: '',
      studioBiblicoReaderId: '',
      canticoFinale: 'Cantico 3: ',
      preghieraFinaleId: '',
    };

    handleUpdateMeetings([...meetings, newMeeting]);
    onShowToast(`Aggiunta nuova settimana: ${dateLabel}`);
  };

  const handleDeleteMeeting = (id: string, dateLabel: string) => {
    if (!checkAdminPermission()) return;
    if (confirm(`Eliminare l'adunanza di "${dateLabel}"?`)) {
      handleUpdateMeetings(meetings.filter(m => m.id !== id));
      onShowToast(`Settimana ${dateLabel} eliminata.`);
    }
  };

  const handleUpdateSingleMeeting = (id: string, updatedFields: Partial<VitaEMinisteroMeeting>) => {
    if (!isAdmin) return;
    const updated = meetings.map(m => (m.id === id ? { ...m, ...updatedFields } : m));
    handleUpdateMeetings(updated);
  };

  // Helper: trova nome partecipante dato l'ID
  const getParticipantName = (id?: string) => {
    if (!id) return '';
    const p = participants.find(part => part.id === id);
    return p ? p.name : '';
  };

  // S-89: Apertura modale anteprima per Lettura Biblica (Parte 3)
  const handleOpenLetturaS89 = (meeting: VitaEMinisteroMeeting) => {
    const studentName = getParticipantName(meeting.tesoriLetturaReaderId);
    if (!studentName) {
      onShowToast('Seleziona prima il lettore per compilare il foglietto S-89');
      return;
    }
    const item: S89Item = {
      id: `${meeting.id}_lettura`,
      studentName,
      assistantName: '',
      dateLabel: meeting.dateLabel || meeting.dateStr,
      partNumber: 3,
      partTitle: meeting.tesoriLetturaTitle || 'Lettura biblica',
      room: meeting.tesoriLetturaRoom || 'main',
      isSent: !!meeting.tesoriLetturaSent,
    };
    setActiveS89Modal({
      item,
      meetingId: meeting.id,
      targetType: 'lettura',
    });
  };

  // S-89: Download diretto PDF per Lettura Biblica (Parte 3) con nome dello studente nel file
  const handleDownloadLetturaS89 = (meeting: VitaEMinisteroMeeting) => {
    const studentName = getParticipantName(meeting.tesoriLetturaReaderId);
    if (!studentName) {
      onShowToast('Seleziona prima il lettore per scaricare il foglietto S-89');
      return;
    }
    const item: S89Item = {
      studentName,
      assistantName: '',
      dateLabel: meeting.dateLabel || meeting.dateStr,
      partNumber: 3,
      partTitle: meeting.tesoriLetturaTitle || 'Lettura biblica',
      room: meeting.tesoriLetturaRoom || 'main',
      isSent: !!meeting.tesoriLetturaSent,
    };
    downloadS89Pdf(item);
    onShowToast(`Scaricato foglietto: ${getS89FileName(studentName, 3)}`);
  };

  // S-89: Toggle flag "Inviato o non inviato" per Lettura Biblica
  const handleToggleLetturaSent = (meeting: VitaEMinisteroMeeting) => {
    if (!checkAdminPermission()) return;
    const newSent = !meeting.tesoriLetturaSent;
    handleUpdateSingleMeeting(meeting.id, { tesoriLetturaSent: newSent });
    const studentName = getParticipantName(meeting.tesoriLetturaReaderId) || 'Lettura biblica';
    onShowToast(`Foglietto S-89 per ${studentName}: ${newSent ? 'INVIATO ✓' : 'NON INVIATO'}`);
  };

  // S-89: Apertura modale anteprima per Esercitazione Efficaci nel Ministero
  const handleOpenPartS89 = (meeting: VitaEMinisteroMeeting, part: MinisteroPart, pIdx: number) => {
    const studentName = getParticipantName(part.studentId);
    if (!studentName) {
      onShowToast('Seleziona prima lo studente per compilare il foglietto S-89');
      return;
    }
    const assistantName = part.hasAssistant ? getParticipantName(part.assistantId) : '';
    const item: S89Item = {
      id: `${meeting.id}_part_${part.id}`,
      studentName,
      assistantName,
      dateLabel: meeting.dateLabel || meeting.dateStr,
      partNumber: part.number,
      partTitle: part.title,
      room: part.room || 'main',
      isSent: !!part.isSent,
    };
    setActiveS89Modal({
      item,
      meetingId: meeting.id,
      targetType: 'ministeroPart',
      partIdx: pIdx,
    });
  };

  // S-89: Download diretto PDF per parte di Efficaci nel Ministero con nome dello studente nel file
  const handleDownloadPartS89 = (meeting: VitaEMinisteroMeeting, part: MinisteroPart) => {
    const studentName = getParticipantName(part.studentId);
    if (!studentName) {
      onShowToast('Seleziona prima lo studente per scaricare il foglietto S-89');
      return;
    }
    const assistantName = part.hasAssistant ? getParticipantName(part.assistantId) : '';
    const item: S89Item = {
      studentName,
      assistantName,
      dateLabel: meeting.dateLabel || meeting.dateStr,
      partNumber: part.number,
      partTitle: part.title,
      room: part.room || 'main',
      isSent: !!part.isSent,
    };
    downloadS89Pdf(item);
    onShowToast(`Scaricato foglietto: ${getS89FileName(studentName, part.number)}`);
  };

  // S-89: Toggle flag "Inviato o non inviato" per parte di Efficaci nel Ministero
  const handleTogglePartSent = (meeting: VitaEMinisteroMeeting, pIdx: number) => {
    if (!checkAdminPermission()) return;
    const updatedParts = [...meeting.ministeroParts];
    const part = updatedParts[pIdx];
    const newSent = !part.isSent;
    updatedParts[pIdx] = { ...part, isSent: newSent };
    handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
    const studentName = getParticipantName(part.studentId) || `Parte ${part.number}`;
    onShowToast(`Foglietto S-89 per ${studentName}: ${newSent ? 'INVIATO ✓' : 'NON INVIATO'}`);
  };

  // S-89: Scarica tutti i foglietti S-89 dell'adunanza in un unico file PDF multipagina
  const handleDownloadAllMeetingS89 = (meeting: VitaEMinisteroMeeting) => {
    const items: S89Item[] = [];

    // 1. Lettura biblica (parte 3)
    if (meeting.tesoriLetturaReaderId) {
      const sName = getParticipantName(meeting.tesoriLetturaReaderId);
      if (sName) {
        items.push({
          studentName: sName,
          dateLabel: meeting.dateLabel || meeting.dateStr,
          partNumber: 3,
          partTitle: meeting.tesoriLetturaTitle || 'Lettura biblica',
          room: meeting.tesoriLetturaRoom || 'main',
          isSent: !!meeting.tesoriLetturaSent,
        });
      }
    }

    // 2. Efficaci nel ministero
    (meeting.ministeroParts || []).forEach(part => {
      if (part.studentId) {
        const sName = getParticipantName(part.studentId);
        if (sName) {
          items.push({
            studentName: sName,
            assistantName: part.hasAssistant ? getParticipantName(part.assistantId) : '',
            dateLabel: meeting.dateLabel || meeting.dateStr,
            partNumber: part.number,
            partTitle: part.title,
            room: part.room || 'main',
            isSent: !!part.isSent,
          });
        }
      }
    });

    if (items.length === 0) {
      onShowToast('Nessuna assegnazione studente presente in questa settimana');
      return;
    }

    const weekFile = `Foglietti_S-89_${(meeting.dateLabel || meeting.dateStr || 'Settimana').replace(/[/\\?%*:|"<>]/g, '').replace(/\s+/g, '_')}.pdf`;
    downloadCombinedS89Pdf(items, weekFile);
    onShowToast(`Scaricati ${items.length} foglietti S-89 per la settimana`);
  };

  // S-89: Callback salvataggio modifiche da dentro la finestra modale
  const handleUpdateS89ModalStatus = (isSent: boolean, room: 'main' | 'aux1' | 'aux2') => {
    if (!activeS89Modal) return;
    const { meetingId, targetType, partIdx } = activeS89Modal;

    if (targetType === 'lettura') {
      handleUpdateSingleMeeting(meetingId, {
        tesoriLetturaSent: isSent,
        tesoriLetturaRoom: room,
      });
    } else if (targetType === 'ministeroPart' && partIdx !== undefined) {
      const targetMeeting = meetings.find(m => m.id === meetingId);
      if (targetMeeting) {
        const updatedParts = [...targetMeeting.ministeroParts];
        if (updatedParts[partIdx]) {
          updatedParts[partIdx] = {
            ...updatedParts[partIdx],
            isSent,
            room,
          };
          handleUpdateSingleMeeting(meetingId, { ministeroParts: updatedParts });
        }
      }
    }

    setActiveS89Modal(prev => (prev ? {
      ...prev,
      item: {
        ...prev.item,
        isSent,
        room,
      }
    } : null));
  };

  // Helper for participant select options:
  // Components not yet assigned or assigned longest ago appear at the top;
  // Components assigned most recently sink to the bottom ("vengono selezionati come ultimi")
  const renderParticipantOptions = (
    roleKey?: keyof VitaEMinisteroParticipant['roles'],
    currentSelectedId?: string,
    filterGender?: 'M' | 'F',
    requiredPartTypeIds?: string[],
    onlyEligible = false
  ) => {
    const eligible = participants
      .filter(p => {
        const matchGender = !filterGender || p.gender === filterGender;
        const matchRole = !roleKey || !!p.roles[roleKey];
        const matchPartType = !requiredPartTypeIds || isParticipantEligibleForPartTypes(p.roles, requiredPartTypeIds);
        return matchGender && matchRole && matchPartType;
      })
      .sort((a, b) => compareParticipantsByLastAssignmentAsc(a, b, statsMap));

    const others = participants
      .filter(p => !eligible.some(el => el.id === p.id))
      .sort((a, b) => compareParticipantsByLastAssignmentAsc(a, b, statsMap));

    const formatOptionLabel = (p: VitaEMinisteroParticipant) => {
      const s = statsMap.get(p.id);
      const genderStr = p.gender === 'M' ? 'Fratello' : 'Sorella';
      if (!s || !s.lastAssignmentDate) {
        return `${p.name} (${genderStr}) — ★ Pronto (Mai assegnato)`;
      }
      return `${p.name} (${genderStr}) — Ultima: ${s.lastAssignmentLabel} [tot. ${s.totalAssignments}]`;
    };

    return (
      <>
        <option value="">-- Seleziona --</option>
        {eligible.length > 0 && (
          <optgroup label="✓ Abilitati (Ordinati per rotazione: ultimi assegnati in fondo)">
            {eligible.map(p => (
              <option key={p.id} value={p.id}>
                {formatOptionLabel(p)}
              </option>
            ))}
          </optgroup>
        )}
        {!onlyEligible && others.length > 0 && (
          <optgroup label="Altri nominativi (non specificamente abilitati)">
            {others.map(p => (
              <option key={p.id} value={p.id}>
                {formatOptionLabel(p)}
              </option>
            ))}
          </optgroup>
        )}
      </>
    );
  };

  return (
    <div className="space-y-6">
      {/* Top Header & Sub-Navigation Bar */}
      <div className="card flex flex-col md:flex-row md:items-center md:justify-between gap-4">
        <div>
          <div className="flex items-center gap-2">
            <div className="p-2 bg-indigo-50 dark:bg-indigo-950/60 text-indigo-700 dark:text-indigo-300 rounded-xl">
              <BookOpen className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-xl font-black text-slate-900 dark:text-slate-100">
                Programmazione Vita e Ministero
              </h1>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Gestione parti e proclamatori per l'adunanza infrasettimanale (wol.jw.org)
              </p>
            </div>
          </div>
        </div>

        {/* Sub-Tabs Selector */}
        <div className="flex items-center bg-slate-100 dark:bg-slate-900 rounded-xl p-1 border border-slate-200 dark:border-slate-800 text-xs font-bold">
          <button
            onClick={() => setActiveSubTab('programma')}
            className={`px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all ${
              activeSubTab === 'programma'
                ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-xs'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
            }`}
          >
            <Calendar className="w-3.5 h-3.5" />
            Programma
          </button>

          <button
            onClick={() => setActiveSubTab('nominativi')}
            className={`px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all ${
              activeSubTab === 'nominativi'
                ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-xs'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
            }`}
          >
            <Users className="w-3.5 h-3.5" />
            Nominativi &amp; Parti ({participants.length})
          </button>

          <button
            onClick={() => setActiveSubTab('statistiche')}
            className={`px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all ${
              activeSubTab === 'statistiche'
                ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-xs'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
            }`}
          >
            <BarChart3 className="w-3.5 h-3.5" />
            Statistiche &amp; Rotazione
          </button>

          <button
            onClick={() => setActiveSubTab('stampa')}
            className={`px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all ${
              activeSubTab === 'stampa'
                ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-xs'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
            }`}
          >
            <Printer className="w-3.5 h-3.5" />
            Stampa / PDF Ufficiale
          </button>
        </div>
      </div>

      {/* =========================================================================
          SUB-TAB 1: PROGRAMMA (MEETING EDITOR)
          ========================================================================= */}
      {activeSubTab === 'programma' && (
        <div className="space-y-6">
          {/* Month Navigator & Controls */}
          <div className="card flex flex-wrap items-center justify-between gap-3 bg-white dark:bg-slate-800/80">
            <div className="flex items-center gap-2">
              <button
                onClick={handlePrevMonth}
                className="p-1.5 rounded-lg border border-slate-200 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <div className="text-base font-extrabold text-slate-900 dark:text-slate-100 min-w-[170px] text-center">
                {MESI_FULL[selectedMonth]} {selectedYear}
              </div>
              <button
                onClick={handleNextMonth}
                className="p-1.5 rounded-lg border border-slate-200 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300"
              >
                <ChevronRight className="w-4 h-4" />
              </button>

              {isAdmin && (
                <button
                  onClick={handleQuickSyncCurrentMonth}
                  disabled={isQuickSyncing}
                  title={`Sincronizza ${MESI_FULL[selectedMonth]} ${selectedYear} da wol.jw.org`}
                  className="px-2 py-1 text-xs font-semibold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-950/60 hover:bg-indigo-100 dark:hover:bg-indigo-900/60 border border-indigo-200 dark:border-indigo-800 rounded-lg flex items-center gap-1 transition-colors"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isQuickSyncing ? 'animate-spin text-indigo-600' : ''}`} />
                  <span className="hidden sm:inline">Aggiorna mese da WOL</span>
                </button>
              )}
            </div>

            <div className="flex items-center gap-2.5">
              {/* Congregation Name input */}
              <div className="flex items-center gap-1.5">
                <Building2 className="w-4 h-4 text-slate-400" />
                <span className="text-xs text-slate-500 font-medium hidden sm:inline">Congregazione:</span>
                <input
                  type="text"
                  value={vmData.congregationName || 'Roccastrada'}
                  disabled={!isAdmin}
                  onChange={e => updateVmData({ ...vmData, congregationName: e.target.value })}
                  placeholder="Nome Congregazione"
                  className="inp text-xs py-1 px-2 w-32 sm:w-36 font-semibold"
                />
              </div>

              {isAdmin && (
                <>
                  <button
                    onClick={handleAutoGenerate}
                    disabled={isQuickSyncing || meetingsInMonth.length === 0}
                    title="Rigenera i nominativi del mese: priorità a chi non ha incarichi da più tempo, rispettando abilitazioni, indisponibilità e altri programmi."
                    className="px-3 py-1.5 text-xs font-bold bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white rounded-lg flex items-center gap-1.5"
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    Genera automaticamente
                  </button>
                  <button
                    onClick={() => setIsWolModalOpen(true)}
                    className="px-3 py-1.5 text-xs font-bold bg-indigo-50 hover:bg-indigo-100 dark:bg-indigo-950/60 dark:hover:bg-indigo-900/60 text-indigo-700 dark:text-indigo-300 border border-indigo-300 dark:border-indigo-800 rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
                  >
                    <Globe className="w-3.5 h-3.5 text-indigo-600" />
                    <span>Scarica da wol.jw.org</span>
                  </button>

                  <button
                    onClick={handleAddMeeting}
                    className="px-3 py-1.5 text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span className="hidden sm:inline">Aggiungi Settimana</span>
                  </button>
                </>
              )}
            </div>
          </div>

          {monthlyRepeats.length > 0 && (
            <div className="card no-print border-amber-300 bg-amber-50 dark:bg-amber-950/30 text-xs" role="alert">
              <p className="font-bold mb-2">Stessa persona ripetuta nel mese</p>
              <ul className="list-disc pl-4 space-y-1">{monthlyRepeats.map((warning, i) => <li key={i}>{warning}</li>)}</ul>
            </div>
          )}

          {generationWarnings.length > 0 && (
            <div className="card no-print border-amber-300 bg-amber-50 dark:bg-amber-950/30 text-xs" role="status">
              <p className="font-bold mb-2">Incarichi da verificare</p>
              <ul className="list-disc pl-4 space-y-1">{generationWarnings.map((warning, i) => <li key={i}>{warning}</li>)}</ul>
            </div>
          )}

          {/* List of Meetings in this month */}
          {meetingsInMonth.length === 0 ? (
            <div className="card text-center py-12 space-y-4">
              <Calendar className="w-10 h-10 text-slate-300 dark:text-slate-600 mx-auto" />
              <div>
                <div className="text-sm font-bold text-slate-800 dark:text-slate-200">
                  Nessuna adunanza presente per {MESI_FULL[selectedMonth]} {selectedYear}.
                </div>
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-md mx-auto">
                  Puoi scaricare le settimane ufficiali di questo mese o di tutti i mesi direttamente dalla Biblioteca Online Watchtower (wol.jw.org).
                </p>
              </div>

              {isAdmin && (
                <div className="flex flex-wrap items-center justify-center gap-3 pt-2">
                  <button
                    onClick={handleQuickSyncCurrentMonth}
                    disabled={isQuickSyncing}
                    className="px-4 py-2 text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl inline-flex items-center gap-2 shadow-xs transition-colors"
                  >
                    {isQuickSyncing ? (
                      <Loader2 className="w-4 h-4 animate-spin" />
                    ) : (
                      <Globe className="w-4 h-4" />
                    )}
                    Scarica {MESI_FULL[selectedMonth]} {selectedYear} da wol.jw.org
                  </button>

                  <button
                    onClick={() => setIsWolModalOpen(true)}
                    className="px-4 py-2 text-xs font-bold bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-xl inline-flex items-center gap-2 transition-colors"
                  >
                    <Download className="w-4 h-4 text-indigo-600" />
                    Opzioni scaricamento tutti i mesi...
                  </button>
                </div>
              )}
            </div>
          ) : (
            meetingsInMonth.map((meeting, mIdx) => (
              <div
                key={meeting.id}
                className="card space-y-5 border border-slate-300 dark:border-slate-700 shadow-sm"
              >
                {/* Meeting Header Bar */}
                <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3 pb-3 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="font-extrabold text-base text-slate-900 dark:text-slate-100 flex items-center gap-2">
                      <span className="w-2.5 h-2.5 rounded-full bg-indigo-600"></span>
                      <input
                        type="text"
                        disabled={!isAdmin}
                        value={meeting.dateLabel}
                        onChange={e => handleUpdateSingleMeeting(meeting.id, { dateLabel: e.target.value })}
                        className="font-bold text-sm bg-transparent border-b border-transparent hover:border-slate-300 focus:border-indigo-500 outline-hidden px-1"
                        placeholder="Es. 5 ottobre 2026"
                      />
                      <span className="text-slate-400">|</span>
                      <input
                        type="text"
                        disabled={!isAdmin}
                        value={meeting.bibleReading}
                        onChange={e => handleUpdateSingleMeeting(meeting.id, { bibleReading: e.target.value })}
                        className="font-bold text-sm text-indigo-700 dark:text-indigo-400 bg-transparent border-b border-transparent hover:border-slate-300 focus:border-indigo-500 outline-hidden px-1 uppercase"
                        placeholder="Es. GEREMIA 40-41"
                      />
                    </div>

                    <label className="chk text-xs bg-slate-100 dark:bg-slate-800 px-2 py-1 rounded-md">
                      <input
                        type="checkbox"
                        disabled={!isAdmin}
                        checked={!!meeting.isSpecialEvent}
                        onChange={e => handleUpdateSingleMeeting(meeting.id, { isSpecialEvent: e.target.checked })}
                      />
                      <span>Evento Speciale / Assemblea</span>
                    </label>
                  </div>

                  <div className="flex flex-wrap items-center gap-2 self-end md:self-auto">
                    {/* S-89 Weekly Summary & Bulk Download */}
                    {(() => {
                      let count = 0;
                      let sent = 0;
                      if (meeting.tesoriLetturaReaderId) {
                        count++;
                        if (meeting.tesoriLetturaSent) sent++;
                      }
                      (meeting.ministeroParts || []).forEach(p => {
                        if (p.studentId) {
                          count++;
                          if (p.isSent) sent++;
                        }
                      });

                      if (count === 0) return null;

                      return (
                        <div className="flex items-center gap-1.5">
                          <span className={`text-[11px] font-bold px-2 py-0.5 rounded-md border flex items-center gap-1 ${
                            sent === count
                              ? 'bg-emerald-50 text-emerald-700 border-emerald-300 dark:bg-emerald-950/50 dark:text-emerald-300 dark:border-emerald-800'
                              : 'bg-amber-50 text-amber-800 border-amber-300 dark:bg-amber-950/50 dark:text-amber-200 dark:border-amber-800'
                          }`}>
                            <FileText className="w-3 h-3 text-amber-500" />
                            <span>S-89: {sent}/{count} inviati</span>
                          </span>

                          <button
                            type="button"
                            onClick={() => handleDownloadAllMeetingS89(meeting)}
                            title="Scarica tutti i foglietti S-89 di questa settimana in un unico file PDF"
                            className="text-[11px] font-bold px-2 py-1 rounded-md bg-indigo-50 dark:bg-indigo-950/50 text-indigo-700 dark:text-indigo-300 hover:bg-indigo-100 border border-indigo-200 dark:border-indigo-800 flex items-center gap-1 transition-colors"
                          >
                            <Download className="w-3 h-3" />
                            <span className="hidden sm:inline">Scarica</span> S-89 ({count})
                          </button>
                        </div>
                      );
                    })()}

                    {isAdmin && (
                      <button
                        onClick={() => handleDeleteMeeting(meeting.id, meeting.dateLabel)}
                        className="text-xs text-rose-600 hover:text-rose-800 dark:text-rose-400 flex items-center gap-1 ml-2"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                        Elimina Settimana
                      </button>
                    )}
                  </div>
                </div>

                {/* If Special Event (Assembly / Circuit Overseer visit) */}
                {meeting.isSpecialEvent ? (
                  <div className="py-6 px-4 bg-amber-50/50 dark:bg-amber-950/20 rounded-xl border border-amber-200 dark:border-amber-900/40 text-center space-y-3">
                    <Sparkles className="w-6 h-6 text-amber-600 mx-auto" />
                    <label className="lbl block">Titolo Evento Speciale / Assemblea:</label>
                    <input
                      type="text"
                      disabled={!isAdmin}
                      value={meeting.specialEventTitle || 'Assemblea di circoscrizione'}
                      onChange={e => handleUpdateSingleMeeting(meeting.id, { specialEventTitle: e.target.value })}
                      className="inp text-base font-bold text-center max-w-md mx-auto"
                    />
                    <p className="text-xs text-amber-700 dark:text-amber-300">
                      Durante questa settimana l'adunanza infrasettimanale è sospesa o dedicata all'evento indicato.
                    </p>
                  </div>
                ) : (
                  <div className="space-y-5">
                    {/* INTESTAZIONE: Presidente, Cantico iniziale, Preghiera iniziale */}
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 p-3 bg-slate-50 dark:bg-slate-900/50 rounded-xl border border-slate-200 dark:border-slate-800 text-xs">
                      <div>
                        <label className="lbl block mb-1">Presidente dell'adunanza (M)</label>
                        <select
                          disabled={!isAdmin}
                          value={meeting.presidenteId || ''}
                          onChange={e => handleUpdateSingleMeeting(meeting.id, { presidenteId: e.target.value })}
                          className="inp text-xs py-1"
                        >
                          {renderParticipantOptions('presidente', meeting.presidenteId, 'M')}
                        </select>
                      </div>

                      <div>
                        <label className="lbl block mb-1">Cantico Iniziale</label>
                        <input
                          type="text"
                          disabled={!isAdmin}
                          value={meeting.canticoIniziale || ''}
                          onChange={e => handleUpdateSingleMeeting(meeting.id, { canticoIniziale: e.target.value })}
                          placeholder="Es. Cantico 33: Getta su Geova..."
                          className="inp text-xs py-1"
                        />
                      </div>

                      <div>
                        <label className="lbl block mb-1">Preghiera Iniziale (M)</label>
                        <select
                          disabled={!isAdmin}
                          value={meeting.preghieraInizialeId || ''}
                          onChange={e => handleUpdateSingleMeeting(meeting.id, { preghieraInizialeId: e.target.value })}
                          className="inp text-xs py-1"
                        >
                          {renderParticipantOptions('preghiera', meeting.preghieraInizialeId, 'M')}
                        </select>
                      </div>
                    </div>

                    {/* ================= SECTION 1: TESORI DELLA PAROLA DI DIO ================= */}
                    <div className="space-y-2">
                      <div
                        style={{ backgroundColor: '#2E5B6E' }}
                        className="text-white text-xs font-bold tracking-wider px-3 py-1.5 rounded-lg uppercase flex items-center justify-between"
                      >
                        <span>TESORI DELLA PAROLA DI DIO</span>
                        <span className="text-[10px] font-normal opacity-80">Oratori fratelli idonei</span>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 p-3 bg-slate-50/70 dark:bg-slate-900/30 rounded-xl border border-slate-200 dark:border-slate-800 text-xs">
                        {/* 1. Discorso 10 min */}
                        <div className="space-y-1">
                          <span className="font-bold text-slate-800 dark:text-slate-200 block">
                            1. Discorso (10 min)
                          </span>
                          <input
                            type="text"
                            disabled={!isAdmin}
                            value={meeting.tesori1Title || ''}
                            onChange={e => handleUpdateSingleMeeting(meeting.id, { tesori1Title: e.target.value })}
                            placeholder="Titolo del discorso"
                            className="inp text-xs py-1 mb-1"
                          />
                          <select
                            disabled={!isAdmin}
                            value={meeting.tesori1SpeakerId || ''}
                            onChange={e => handleUpdateSingleMeeting(meeting.id, { tesori1SpeakerId: e.target.value })}
                            className="inp text-xs py-1"
                          >
                            {renderParticipantOptions('tesoriDiscorso', meeting.tesori1SpeakerId, 'M')}
                          </select>
                        </div>

                        {/* 2. Gemme spirituali 10 min */}
                        <div className="space-y-1">
                          <span className="font-bold text-slate-800 dark:text-slate-200 block">
                            2. Gemme spirituali (10 min)
                          </span>
                          <input
                            type="text"
                            disabled={!isAdmin}
                            value={meeting.tesoriGemmeTitle || 'Gemme spirituali'}
                            onChange={e => handleUpdateSingleMeeting(meeting.id, { tesoriGemmeTitle: e.target.value })}
                            placeholder="Titolo gemme"
                            className="inp text-xs py-1 mb-1"
                          />
                          <select
                            disabled={!isAdmin}
                            value={meeting.tesoriGemmeSpeakerId || ''}
                            onChange={e => handleUpdateSingleMeeting(meeting.id, { tesoriGemmeSpeakerId: e.target.value })}
                            className="inp text-xs py-1"
                          >
                            {renderParticipantOptions('tesoriGemme', meeting.tesoriGemmeSpeakerId, 'M')}
                          </select>
                        </div>

                        {/* 3. Lettura biblica 4 min */}
                        <div className="space-y-1">
                          <span className="font-bold text-slate-800 dark:text-slate-200 block">
                            3. Lettura biblica (4 min)
                          </span>
                          <input
                            type="text"
                            disabled={!isAdmin}
                            value={meeting.tesoriLetturaTitle || 'Lettura biblica'}
                            onChange={e => handleUpdateSingleMeeting(meeting.id, { tesoriLetturaTitle: e.target.value })}
                            placeholder="Titolo lettura"
                            className="inp text-xs py-1 mb-1"
                          />
                          <select
                            disabled={!isAdmin}
                            value={meeting.tesoriLetturaReaderId || ''}
                            onChange={e => handleUpdateSingleMeeting(meeting.id, { tesoriLetturaReaderId: e.target.value })}
                            className="inp text-xs py-1"
                          >
                            {renderParticipantOptions('tesoriLettura', meeting.tesoriLetturaReaderId, 'M')}
                          </select>

                          {/* S-89 Assignment Controls for Lettura Biblica */}
                          {meeting.tesoriLetturaReaderId && (
                            <div className="mt-2 p-2 rounded-lg bg-amber-50/70 dark:bg-amber-950/30 border border-amber-200 dark:border-amber-800/60 flex flex-wrap items-center justify-between gap-1.5">
                              <div className="flex items-center gap-1.5">
                                <button
                                  type="button"
                                  disabled={!isAdmin}
                                  onClick={() => handleToggleLetturaSent(meeting)}
                                  title="Clicca per invertire lo stato Inviato / Non inviato"
                                  className={`px-2 py-0.5 rounded text-[10px] font-bold flex items-center gap-1 border transition-colors cursor-pointer select-none ${
                                    meeting.tesoriLetturaSent
                                      ? 'bg-emerald-500 text-white border-emerald-600 shadow-2xs hover:bg-emerald-600'
                                      : 'bg-amber-100 dark:bg-amber-900/60 text-amber-800 dark:text-amber-200 border-amber-300 dark:border-amber-700 hover:bg-amber-200'
                                  }`}
                                >
                                  {meeting.tesoriLetturaSent ? (
                                    <>
                                      <CheckCircle2 className="w-2.5 h-2.5" />
                                      <span>Inviato ✓</span>
                                    </>
                                  ) : (
                                    <>
                                      <Clock className="w-2.5 h-2.5" />
                                      <span>Non inviato</span>
                                    </>
                                  )}
                                </button>

                                <select
                                  disabled={!isAdmin}
                                  value={meeting.tesoriLetturaRoom || 'main'}
                                  onChange={e => handleUpdateSingleMeeting(meeting.id, { tesoriLetturaRoom: e.target.value as any })}
                                  className="text-[10px] font-semibold py-0.5 px-1 rounded border border-amber-300/80 dark:border-amber-700 bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-300"
                                >
                                  <option value="main">Sala principale</option>
                                  <option value="aux1">Sala secondaria 1</option>
                                  <option value="aux2">Sala secondaria 2</option>
                                </select>
                              </div>

                              <div className="flex items-center gap-1">
                                <button
                                  type="button"
                                  onClick={() => handleDownloadLetturaS89(meeting)}
                                  title={`Scarica file PDF: ${getS89FileName(getParticipantName(meeting.tesoriLetturaReaderId), 3)}`}
                                  className="p-1 px-1.5 rounded bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 border border-slate-200 dark:border-slate-700 hover:bg-indigo-50 dark:hover:bg-slate-700 transition-colors flex items-center gap-1 text-[10px] font-bold"
                                >
                                  <Download className="w-3 h-3" />
                                  <span>PDF</span>
                                </button>

                                <button
                                  type="button"
                                  onClick={() => handleOpenLetturaS89(meeting)}
                                  title="Vedi e compila foglietto S-89"
                                  className="p-1 px-1.5 rounded bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-300 border border-slate-200 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors flex items-center gap-1 text-[10px] font-semibold"
                                >
                                  <FileText className="w-3 h-3 text-amber-500" />
                                  <span>S-89</span>
                                </button>
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* ================= SECTION 2: EFFICACI NEL MINISTERO ================= */}
                    <div className="space-y-2">
                      <div
                        style={{ backgroundColor: '#9C732B' }}
                        className="text-white text-xs font-bold tracking-wider px-3 py-1.5 rounded-lg uppercase flex items-center justify-between"
                      >
                        <span>EFFICACI NEL MINISTERO (ESERCITAZIONI STUDENTI)</span>
                        {isAdmin && (
                          <button
                            onClick={() => {
                              const newPartNum = (meeting.ministeroParts.length > 0 ? meeting.ministeroParts[meeting.ministeroParts.length - 1].number + 1 : 4);
                              const newPart: MinisteroPart = {
                                id: 'mp_' + Math.random().toString(36).substring(2, 7),
                                number: newPartNum,
                                title: 'Iniziare una conversazione',
                                minutes: 3,
                                studentId: '',
                                assistantId: '',
                                hasAssistant: true,
                                partTypeIds: inferPartTypeIds(3, true),
                              };
                              handleUpdateSingleMeeting(meeting.id, {
                                ministeroParts: [...meeting.ministeroParts, newPart],
                              });
                            }}
                            className="px-2 py-0.5 text-[10px] font-bold bg-white/20 hover:bg-white/30 text-white rounded flex items-center gap-1 transition-colors"
                          >
                            <Plus className="w-3 h-3" /> Aggiungi Parte
                          </button>
                        )}
                      </div>

                      <div className="space-y-2 p-3 bg-slate-50/70 dark:bg-slate-900/30 rounded-xl border border-slate-200 dark:border-slate-800 text-xs">
                        {meeting.ministeroParts.map((part, pIdx) => (
                          <div
                            key={part.id}
                            className="p-3 bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-2xs space-y-2.5"
                          >
                            {/* Row 1: Number, Title, Student, Assistant, Delete */}
                            <div className="grid grid-cols-1 sm:grid-cols-12 gap-2.5 items-center">
                              {/* Part Number & Title */}
                              <div className="sm:col-span-4 flex items-center gap-1.5">
                                <span className="font-bold text-slate-500 shrink-0 w-5">
                                  {part.number}.
                                </span>
                                <input
                                  type="text"
                                  disabled={!isAdmin}
                                  value={part.title}
                                  onChange={e => {
                                    const updatedParts = [...meeting.ministeroParts];
                                    updatedParts[pIdx].title = e.target.value;
                                    handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                  }}
                                  placeholder="Titolo parte (es. Iniziare una conversazione)"
                                  className="inp text-xs py-1 flex-1 font-medium"
                                />
                                <div className="flex items-center gap-0.5 shrink-0">
                                  <input
                                    type="number"
                                    disabled={!isAdmin}
                                    value={part.minutes}
                                    onChange={e => {
                                      const updatedParts = [...meeting.ministeroParts];
                                      updatedParts[pIdx].minutes = parseInt(e.target.value, 10) || 1;
                                      handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                    }}
                                    className="inp text-xs py-1 w-12 text-center"
                                  />
                                  <span className="text-[10px] text-slate-400">min</span>
                                </div>
                              </div>

                              {/* Student */}
                              <div className="sm:col-span-3">
                                <label className="text-[10px] text-slate-500 font-semibold block mb-0.5">
                                  Studente / Titolare:
                                </label>
                                <select
                                  disabled={!isAdmin}
                                  value={part.studentId || ''}
                                  onChange={e => {
                                    const updatedParts = [...meeting.ministeroParts];
                                    updatedParts[pIdx].studentId = e.target.value;
                                    handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                  }}
                                  className="inp text-xs py-1"
                                >
                                  {renderParticipantOptions('ministeroStudente', part.studentId, undefined, part.partTypeIds, !part.hasAssistant)}
                                </select>
                              </div>

                              {/* Assistant Checkbox & Selector */}
                              <div className="sm:col-span-4">
                                <div className="flex items-center justify-between mb-0.5">
                                  <label className="chk text-[10px]">
                                    <input
                                      type="checkbox"
                                      disabled={!isAdmin}
                                      checked={part.hasAssistant}
                                      onChange={e => {
                                        const hasAssistant = e.target.checked;
                                        const partTypeIds = [
                                          ...(part.partTypeIds || []).filter(id => id !== 'dimostrazione' && id !== 'discorso'),
                                          hasAssistant ? 'dimostrazione' : 'discorso',
                                        ];
                                        const student = participants.find(p => p.id === part.studentId);
                                        const studentId = !hasAssistant && student &&
                                          (!student.roles.ministeroStudente || !isParticipantEligibleForPartTypes(student.roles, partTypeIds))
                                          ? '' : part.studentId;
                                        const updatedParts = meeting.ministeroParts.map((p, index) => index === pIdx ? {
                                          ...p,
                                          hasAssistant,
                                          assistantId: hasAssistant ? p.assistantId : '',
                                          studentId,
                                          partTypeIds,
                                          isSent: false,
                                        } : p);
                                        handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                      }}
                                    />
                                    <span>Con Assistente</span>
                                  </label>
                                </div>

                                {part.hasAssistant ? (
                                  <select
                                    disabled={!isAdmin}
                                    value={part.assistantId || ''}
                                    onChange={e => {
                                      const updatedParts = [...meeting.ministeroParts];
                                      updatedParts[pIdx].assistantId = e.target.value;
                                      handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                    }}
                                    className="inp text-xs py-1"
                                  >
                                    {renderParticipantOptions('ministeroAssistente', part.assistantId)}
                                  </select>
                                ) : (
                                  <div className="text-[11px] text-slate-400 italic py-1">
                                    Discorso singolo (nessun assistente)
                                  </div>
                                )}
                              </div>

                              {/* Delete Part Button */}
                              {isAdmin && (
                                <div className="sm:col-span-1 text-right">
                                  <button
                                    onClick={() => {
                                      const updatedParts = meeting.ministeroParts.filter((_, idx) => idx !== pIdx);
                                      handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                    }}
                                    className="p-1 text-slate-400 hover:text-rose-600 rounded"
                                    title="Rimuovi parte"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              )}
                            </div>

                            {/* Row 1.5: Tipo di parte (criteri per filtrare i proclamatori idonei) */}
                            <div className="flex flex-wrap items-center gap-1.5 pt-1">
                              <span className="text-[10px] text-slate-500 font-semibold shrink-0">
                                Tipo di parte:
                              </span>
                              {ministeroPartTypes.map(pt => {
                                const active = !!part.partTypeIds?.includes(pt.id);
                                return (
                                  <button
                                    key={pt.id}
                                    type="button"
                                    disabled={!isAdmin}
                                    onClick={() => {
                                      const updatedParts = [...meeting.ministeroParts];
                                      const current = updatedParts[pIdx].partTypeIds || [];
                                      updatedParts[pIdx].partTypeIds = active
                                        ? current.filter(id => id !== pt.id)
                                        : [...current, pt.id];
                                      handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                    }}
                                    title={pt.label}
                                    className={`px-2 py-0.5 rounded-full text-[10px] font-semibold border transition-colors ${
                                      active
                                        ? 'bg-violet-100 border-violet-300 text-violet-800 dark:bg-violet-950/70 dark:border-violet-700 dark:text-violet-300'
                                        : 'bg-white border-slate-200 text-slate-500 dark:bg-slate-900 dark:border-slate-700 dark:text-slate-400'
                                    } ${!isAdmin ? 'opacity-70' : 'hover:border-violet-300'}`}
                                  >
                                    {pt.label}
                                  </button>
                                );
                              })}
                              {isAdmin && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    const updatedParts = [...meeting.ministeroParts];
                                    updatedParts[pIdx].partTypeIds = inferPartTypeIds(part.minutes, part.hasAssistant);
                                    handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                  }}
                                  title="Reimposta i tipi in base a durata e formato della parte"
                                  className="px-2 py-0.5 rounded-full text-[10px] font-semibold border border-dashed border-slate-300 text-slate-500 hover:border-slate-400 dark:border-slate-600 dark:text-slate-400"
                                >
                                  Auto
                                </button>
                              )}
                            </div>

                            {/* Row 2: S-89 Assignment Slip Details */}
                            {part.studentId && (
                              <div className="pt-2 border-t border-slate-100 dark:border-slate-700/80 flex flex-wrap items-center justify-between gap-2 text-xs">
                                <div className="flex flex-wrap items-center gap-2">
                                  <span className="text-[11px] font-bold text-slate-600 dark:text-slate-400 flex items-center gap-1">
                                    <FileText className="w-3.5 h-3.5 text-amber-500" />
                                    Foglietto S-89:
                                  </span>

                                  {/* Flag Inviato / Non Inviato */}
                                  <button
                                    type="button"
                                    disabled={!isAdmin}
                                    onClick={() => handleTogglePartSent(meeting, pIdx)}
                                    title="Clicca per invertire lo stato Inviato / Non inviato"
                                    className={`px-2.5 py-0.5 rounded-md text-[10px] font-bold flex items-center gap-1 border transition-all cursor-pointer select-none active:scale-95 ${
                                      part.isSent
                                        ? 'bg-emerald-500 text-white border-emerald-600 shadow-2xs hover:bg-emerald-600'
                                        : 'bg-amber-50 dark:bg-amber-950/40 text-amber-700 dark:text-amber-300 border-amber-300 dark:border-amber-700 hover:bg-amber-100'
                                    }`}
                                  >
                                    {part.isSent ? (
                                      <>
                                        <CheckCircle2 className="w-3 h-3" />
                                        <span>Inviato ✓</span>
                                      </>
                                    ) : (
                                      <>
                                        <Clock className="w-3 h-3" />
                                        <span>Non inviato</span>
                                      </>
                                    )}
                                  </button>

                                  {/* Sala */}
                                  <div className="flex items-center gap-1">
                                    <span className="text-[10px] text-slate-400">Sala:</span>
                                    <select
                                      disabled={!isAdmin}
                                      value={part.room || 'main'}
                                      onChange={e => {
                                        const updatedParts = [...meeting.ministeroParts];
                                        updatedParts[pIdx].room = e.target.value as any;
                                        handleUpdateSingleMeeting(meeting.id, { ministeroParts: updatedParts });
                                      }}
                                      className="text-[10px] font-semibold py-0.5 px-1.5 rounded border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-300"
                                    >
                                      <option value="main">Sala principale</option>
                                      <option value="aux1">Sala secondaria 1</option>
                                      <option value="aux2">Sala secondaria 2</option>
                                    </select>
                                  </div>
                                </div>

                                {/* Actions: Scarica PDF e Visualizza Foglietto */}
                                <div className="flex items-center gap-1.5">
                                  <button
                                    type="button"
                                    onClick={() => handleDownloadPartS89(meeting, part)}
                                    title={`Scarica file PDF: ${getS89FileName(getParticipantName(part.studentId), part.number)}`}
                                    className="px-2.5 py-1 rounded-md text-[11px] font-bold text-white bg-gradient-to-b from-indigo-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 border border-indigo-500 border-b-[2px] border-b-indigo-800 flex items-center gap-1 shadow-2xs active:translate-y-[1px] transition-all"
                                  >
                                    <Download className="w-3 h-3" />
                                    <span>Scarica PDF ({getParticipantName(part.studentId)})</span>
                                  </button>

                                  <button
                                    type="button"
                                    onClick={() => handleOpenPartS89(meeting, part, pIdx)}
                                    title="Visualizza e compila il modulo S-89"
                                    className="px-2.5 py-1 rounded-md text-[11px] font-semibold text-slate-700 dark:text-slate-200 bg-slate-100 dark:bg-slate-700 hover:bg-slate-200 dark:hover:bg-slate-600 border border-slate-300 dark:border-slate-600 flex items-center gap-1 transition-colors"
                                  >
                                    <FileText className="w-3 h-3 text-amber-500" />
                                    <span>Foglietto S-89</span>
                                  </button>
                                </div>
                              </div>
                            )}
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* ================= SECTION 3: VITA CRISTIANA ================= */}
                    <div className="space-y-2">
                      <div
                        style={{ backgroundColor: '#83262E' }}
                        className="text-white text-xs font-bold tracking-wider px-3 py-1.5 rounded-lg uppercase flex items-center justify-between"
                      >
                        <span>VITA CRISTIANA</span>
                        {isAdmin && (
                          <button
                            onClick={() => {
                              const newPartNum = (meeting.ministeroParts.length + meeting.vitaCristianaParts.length + 4);
                              const newVcp: VitaCristianaPart = {
                                id: 'vcp_' + Math.random().toString(36).substring(2, 7),
                                number: newPartNum,
                                title: 'Bisogni locali',
                                minutes: 15,
                                speakerId: '',
                              };
                              handleUpdateSingleMeeting(meeting.id, {
                                vitaCristianaParts: [...meeting.vitaCristianaParts, newVcp],
                              });
                            }}
                            className="px-2 py-0.5 text-[10px] font-bold bg-white/20 hover:bg-white/30 text-white rounded flex items-center gap-1 transition-colors"
                          >
                            <Plus className="w-3 h-3" /> Aggiungi Parte
                          </button>
                        )}
                      </div>

                      <div className="space-y-3 p-3 bg-slate-50/70 dark:bg-slate-900/30 rounded-xl border border-slate-200 dark:border-slate-800 text-xs">
                        {/* Cantico Intermedio */}
                        <div className="flex items-center gap-2">
                          <Music className="w-4 h-4 text-rose-600 shrink-0" />
                          <label className="font-semibold text-slate-700 dark:text-slate-300 shrink-0">
                            Cantico Intermedio:
                          </label>
                          <input
                            type="text"
                            disabled={!isAdmin}
                            value={meeting.canticoIntermedio || ''}
                            onChange={e => handleUpdateSingleMeeting(meeting.id, { canticoIntermedio: e.target.value })}
                            placeholder='Es. Cantico 17: "Lo voglio"'
                            className="inp text-xs py-1 max-w-sm"
                          />
                        </div>

                        {/* Vita Cristiana Parts (e.g. 8. Geova è il difensore delle vedove / Bisogni locali) */}
                        {meeting.vitaCristianaParts.map((vp, vIdx) => (
                          <div
                            key={vp.id}
                            className="p-2.5 bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 grid grid-cols-1 sm:grid-cols-12 gap-2.5 items-center"
                          >
                            <div className="sm:col-span-5 flex items-center gap-1.5">
                              <span className="font-bold text-slate-500 shrink-0 w-5">
                                {vp.number}.
                              </span>
                              <input
                                type="text"
                                disabled={!isAdmin}
                                value={vp.title}
                                onChange={e => {
                                  const updated = [...meeting.vitaCristianaParts];
                                  updated[vIdx].title = e.target.value;
                                  handleUpdateSingleMeeting(meeting.id, { vitaCristianaParts: updated });
                                }}
                                placeholder="Titolo parte"
                                className="inp text-xs py-1 flex-1 font-medium"
                              />
                              <div className="flex items-center gap-0.5 shrink-0">
                                <input
                                  type="number"
                                  disabled={!isAdmin}
                                  value={vp.minutes}
                                  onChange={e => {
                                    const updated = [...meeting.vitaCristianaParts];
                                    updated[vIdx].minutes = parseInt(e.target.value, 10) || 1;
                                    handleUpdateSingleMeeting(meeting.id, { vitaCristianaParts: updated });
                                  }}
                                  className="inp text-xs py-1 w-12 text-center"
                                />
                                <span className="text-[10px] text-slate-400">min</span>
                              </div>
                            </div>

                            <div className="sm:col-span-6">
                              <label className="text-[10px] text-slate-500 font-semibold block mb-0.5">
                                Oratore (Fratello idoneo):
                              </label>
                              <select
                                disabled={!isAdmin}
                                value={vp.speakerId || ''}
                                onChange={e => {
                                  const updated = [...meeting.vitaCristianaParts];
                                  updated[vIdx].speakerId = e.target.value;
                                  handleUpdateSingleMeeting(meeting.id, { vitaCristianaParts: updated });
                                }}
                                className="inp text-xs py-1"
                              >
                                {renderParticipantOptions('vitaCristianaParti', vp.speakerId, 'M')}
                              </select>
                            </div>

                            {isAdmin && (
                              <div className="sm:col-span-1 text-right">
                                <button
                                  onClick={() => {
                                    const updated = meeting.vitaCristianaParts.filter((_, idx) => idx !== vIdx);
                                    handleUpdateSingleMeeting(meeting.id, { vitaCristianaParts: updated });
                                  }}
                                  className="p-1 text-slate-400 hover:text-rose-600 rounded"
                                  title="Rimuovi parte"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            )}
                          </div>
                        ))}

                        {/* Studio biblico o discorso del sorvegliante */}
                        <div className="p-2.5 bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 grid grid-cols-1 sm:grid-cols-2 gap-3">
                          <div className="sm:col-span-2">
                            <label className="lbl block mb-1" htmlFor={`studio-type-${meeting.id}`}>Parte conclusiva di Vita cristiana</label>
                            <select
                              id={`studio-type-${meeting.id}`}
                              disabled={!isAdmin}
                              value={meeting.studioBiblicoType || 'studio'}
                              onChange={e => handleUpdateSingleMeeting(meeting.id, { studioBiblicoType: e.target.value as 'studio' | 'discorsoSorvegliante' })}
                              className="inp text-xs py-1"
                            >
                              <option value="studio">Studio biblico di congregazione</option>
                              <option value="discorsoSorvegliante">Discorso del sorvegliante</option>
                            </select>
                          </div>
                          {meeting.studioBiblicoType === 'discorsoSorvegliante' ? (
                            <div className="sm:col-span-2">
                              <label className="lbl block mb-1" htmlFor={`discorso-title-${meeting.id}`}>Titolo del discorso del sorvegliante</label>
                              <input
                                id={`discorso-title-${meeting.id}`}
                                type="text"
                                disabled={!isAdmin}
                                value={meeting.discorsoSorveglianteTitle || ''}
                                onChange={e => handleUpdateSingleMeeting(meeting.id, { discorsoSorveglianteTitle: e.target.value })}
                                placeholder="Inserisci il titolo del discorso"
                                className="inp text-xs py-1"
                              />
                            </div>
                          ) : (<>
                          <div>
                            <label className="lbl block mb-1">
                              Studio Biblico di Congregazione (30 min) - Conduttore (M)
                            </label>
                            <select
                              disabled={!isAdmin}
                              value={meeting.studioBiblicoConductorId || ''}
                              onChange={e => handleUpdateSingleMeeting(meeting.id, { studioBiblicoConductorId: e.target.value })}
                              className="inp text-xs py-1"
                            >
                              {renderParticipantOptions('studioBiblicoConduttore', meeting.studioBiblicoConductorId, 'M')}
                            </select>
                          </div>

                          <div>
                            <label className="lbl block mb-1">
                              Studio Biblico di Congregazione - Lettore (M)
                            </label>
                            <select
                              disabled={!isAdmin}
                              value={meeting.studioBiblicoReaderId || ''}
                              onChange={e => handleUpdateSingleMeeting(meeting.id, { studioBiblicoReaderId: e.target.value })}
                              className="inp text-xs py-1"
                            >
                              {renderParticipantOptions('studioBiblicoLettore', meeting.studioBiblicoReaderId, 'M')}
                            </select>
                          </div>
                          </>)}
                        </div>

                        {/* Cantico Finale e Preghiera Finale */}
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                          <div>
                            <label className="lbl block mb-1">Cantico Finale</label>
                            <input
                              type="text"
                              disabled={!isAdmin}
                              value={meeting.canticoFinale || ''}
                              onChange={e => handleUpdateSingleMeeting(meeting.id, { canticoFinale: e.target.value })}
                              placeholder="Es. Cantico 38: Dio ti renderà forte"
                              className="inp text-xs py-1"
                            />
                          </div>

                          <div>
                            <label className="lbl block mb-1">Preghiera Finale (M)</label>
                            <select
                              disabled={!isAdmin}
                              value={meeting.preghieraFinaleId || ''}
                              onChange={e => handleUpdateSingleMeeting(meeting.id, { preghieraFinaleId: e.target.value })}
                              className="inp text-xs py-1"
                            >
                              {renderParticipantOptions('preghiera', meeting.preghieraFinaleId, 'M')}
                            </select>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            ))
          )}
        </div>
      )}

      {/* =========================================================================
          SUB-TAB 2: NOMINATIVI & PARTI
          ========================================================================= */}
      {activeSubTab === 'nominativi' && (
        <VitaEMinisteroParticipantsModal
          participants={participants}
          congregationPeople={state.people}
          isAdmin={isAdmin}
          onUpdateParticipants={handleUpdateParticipants}
          onShowToast={onShowToast}
          ministeroPartTypes={ministeroPartTypes}
          onAddCustomPartType={handleAddCustomPartType}
          onRemoveCustomPartType={handleRemoveCustomPartType}
        />
      )}

      {/* =========================================================================
          SUB-TAB 3: STATISTICHE & ROTAZIONE
          ========================================================================= */}
      {activeSubTab === 'statistiche' && (
        <VitaEMinisteroStatsView
          participants={participants}
          meetings={meetings}
          onShowToast={onShowToast}
        />
      )}

      {/* =========================================================================
          SUB-TAB 4: STAMPA / ANTEPRIMA PDF
          ========================================================================= */}
      {activeSubTab === 'stampa' && (
        <VitaEMinisteroPrintView
          data={vmData}
          participants={participants}
        />
      )}

      {/* Modal: Sincronizzazione da wol.jw.org */}
      {isWolModalOpen && (
        <VitaEMinisteroWolModal
          isOpen={isWolModalOpen}
          onClose={() => setIsWolModalOpen(false)}
          currentYear={selectedYear}
          currentMonth={selectedMonth}
          existingMeetings={meetings}
          onApplyMeetings={(newMeetings, msg) => {
            handleUpdateMeetings(newMeetings);
          }}
          onShowToast={onShowToast}
        />
      )}

      {/* Modal: Foglietto Assegnazione S-89 */}
      {activeS89Modal && (
        <S89Modal
          isOpen={!!activeS89Modal}
          onClose={() => setActiveS89Modal(null)}
          item={activeS89Modal.item}
          isAdmin={isAdmin}
          onUpdateStatus={handleUpdateS89ModalStatus}
        />
      )}
    </div>
  );
}
