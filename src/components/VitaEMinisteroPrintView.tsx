import { VitaEMinisteroPrintTemplate } from './VitaEMinisteroPrintTemplate';
import { exportProgramExcel } from '../utils/programExcel';
import React, { useState, useMemo } from 'react';
import { VitaEMinisteroData, VitaEMinisteroMeeting, VitaEMinisteroParticipant } from '../types';
import { Printer, Download, Sparkles, Filter, FileText } from 'lucide-react';
import * as XLSX from 'xlsx';
import { S89Item, downloadCombinedS89Pdf } from '../utils/s89Pdf';

const MESI_FULL = [
  'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
  'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
];

interface VitaEMinisteroPrintViewProps {
  data: VitaEMinisteroData;
  participants: VitaEMinisteroParticipant[];
  onPrint?: () => void;
}

export function VitaEMinisteroPrintView({
  data,
  participants,
}: VitaEMinisteroPrintViewProps) {
  // Default: current month, if it has meetings.
  const [filterPeriod, setFilterPeriod] = useState<string>(() => {
    const now = new Date();
    const current = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    return (data.meetings || []).some(m => m.dateStr?.startsWith(current)) ? current : 'all';
  });

  // Extract distinct available periods
  const availablePeriods = useMemo(() => {
    const map = new Map<string, string>();
    (data.meetings || []).forEach(m => {
      if (m.dateStr && m.dateStr.length >= 7) {
        const key = m.dateStr.substring(0, 7); // "YYYY-MM"
        const [yStr, mStr] = key.split('-');
        const y = parseInt(yStr, 10);
        const mIdx = parseInt(mStr, 10) - 1;
        if (!map.has(key) && mIdx >= 0 && mIdx < 12) {
          map.set(key, `${MESI_FULL[mIdx]} ${y}`);
        }
      }
    });
    return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  }, [data.meetings]);

  const displayedMeetings = useMemo(() => {
    if (filterPeriod === 'all') return data.meetings || [];
    return (data.meetings || []).filter(m => m.dateStr && m.dateStr.startsWith(filterPeriod));
  }, [data.meetings, filterPeriod]);

  const getParticipantName = (id?: string) => {
    if (!id) return '';
    const p = participants.find(part => part.id === id);
    return p ? p.name : '';
  };

  const handlePrint = () => {
    window.print();
  };

  const handleDownloadAllS89 = () => {
    const items: S89Item[] = [];

    displayedMeetings.forEach(m => {
      // 1. Lettura biblica
      if (m.tesoriLetturaReaderId) {
        const name = getParticipantName(m.tesoriLetturaReaderId);
        if (name) {
          items.push({
            studentName: name,
            dateLabel: m.dateLabel || m.dateStr,
            partNumber: 3,
            partTitle: m.tesoriLetturaTitle || 'Lettura biblica',
            room: m.tesoriLetturaRoom || 'main',
            isSent: !!m.tesoriLetturaSent,
          });
        }
      }

      // 2. Efficaci nel ministero
      (m.ministeroParts || []).forEach(part => {
        if (part.studentId) {
          const sName = getParticipantName(part.studentId);
          if (sName) {
            items.push({
              studentName: sName,
              assistantName: part.hasAssistant ? getParticipantName(part.assistantId) : '',
              dateLabel: m.dateLabel || m.dateStr,
              partNumber: part.number,
              partTitle: part.title,
              room: part.room || 'main',
              isSent: !!part.isSent,
            });
          }
        }
      });
    });

    if (items.length === 0) {
      alert('Nessuna assegnazione studente trovata per il periodo selezionato.');
      return;
    }

    const filename = filterPeriod === 'all'
      ? 'Tutti_Foglietti_S-89.pdf'
      : `Foglietti_S-89_${filterPeriod}.pdf`;

    downloadCombinedS89Pdf(items, filename);
  };

  const handleExportExcel = async () => {
try {
    const rows: any[] = [];
    const cong = data.congregationName || 'Congregazione';

    rows.push(['CONGREGAZIONE:', cong, '', 'PROGRAMMA ADUNANZA INFRASETTIMANALE']);
    rows.push([]);

    displayedMeetings.forEach(m => {
      rows.push(['SETTIMANA / DATA:', m.dateLabel, 'LETTURA BIBLICA:', m.bibleReading]);

      if (m.isSpecialEvent) {
        rows.push(['EVENTO SPECIALE:', m.specialEventTitle || 'Assemblea']);
        rows.push([]);
        return;
      }

      rows.push(['Presidente:', getParticipantName(m.presidenteId), 'Cantico iniziale:', m.canticoIniziale || '']);
      rows.push(['Preghiera iniziale:', getParticipantName(m.preghieraInizialeId), '', '']);
      rows.push([]);

      // TESORI
      rows.push(['--- TESORI DELLA PAROLA DI DIO ---', '', '', '']);
      rows.push([
        `1. ${m.tesori1Title || 'Discorso'} (${m.tesori1Minutes || 10} min)`,
        getParticipantName(m.tesori1SpeakerId),
      ]);
      rows.push([
        `2. ${m.tesoriGemmeTitle || 'Gemme spirituali'} (${m.tesoriGemmeMinutes || 10} min)`,
        getParticipantName(m.tesoriGemmeSpeakerId),
      ]);
      rows.push([
        `3. ${m.tesoriLetturaTitle || 'Lettura biblica'} (${m.tesoriLetturaMinutes || 4} min)`,
        getParticipantName(m.tesoriLetturaReaderId),
      ]);
      rows.push([]);

      // EFFICACI NEL MINISTERO
      rows.push(['--- EFFICACI NEL MINISTERO ---', '', '', '']);
      m.ministeroParts.forEach(mp => {
        const student = getParticipantName(mp.studentId);
        const assistant = mp.hasAssistant && mp.assistantId ? ` / ${getParticipantName(mp.assistantId)}` : '';
        rows.push([`${mp.number}. ${mp.title} (${mp.minutes} min)`, `${student}${assistant}`]);
      });
      rows.push([]);

      // VITA CRISTIANA
      rows.push(['--- VITA CRISTIANA ---', '', '', '']);
      rows.push(['Cantico intermedio:', m.canticoIntermedio || '']);
      m.vitaCristianaParts.forEach(vp => {
        rows.push([`${vp.number}. ${vp.title} (${vp.minutes} min)`, getParticipantName(vp.speakerId)]);
      });
      const isOverseerTalk = m.studioBiblicoType === 'discorsoSorvegliante';
      const finalPartTitle = isOverseerTalk
        ? `Discorso del sorvegliante${m.discorsoSorveglianteTitle?.trim() ? `: ${m.discorsoSorveglianteTitle.trim()}` : ''}`
        : m.studioBiblicoTitle || 'Studio biblico di congregazione';
      const cbsConductor = getParticipantName(m.studioBiblicoConductorId);
      const cbsReader = getParticipantName(m.studioBiblicoReaderId);
      rows.push([
        `${finalPartTitle} (${m.studioBiblicoMinutes || 30} min)`,
        isOverseerTalk ? 'Sorvegliante di circoscrizione' : `${cbsConductor}${cbsReader ? ` / ${cbsReader}` : ''}`,
      ]);
      rows.push(['Cantico finale:', m.canticoFinale || '', 'Preghiera finale:', getParticipantName(m.preghieraFinaleId)]);
      rows.push([]);
      rows.push(['-------------------------------------------------------------------------------------']);
      rows.push([]);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Vita e Ministero');
    await exportProgramExcel(workbook, `Programma_Vita_e_Ministero_${cong}.xlsx`, '.excel-ministry');

} catch (error) { alert(error instanceof Error ? error.message : 'Errore durante l’esportazione Excel.'); }
};


  return (
    <div className="space-y-6">
      {/* Top action bar (hidden during print) */}
      <div className="no-print flex flex-wrap items-center justify-between gap-3 p-4 bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-xs">
        <div>
          <h3 className="text-base font-bold text-slate-800 dark:text-slate-100 flex items-center gap-2">
            <Sparkles className="w-4 h-4 text-amber-500" />
            Anteprima e Stampa Documento Ufficiale
          </h3>
          <p className="text-xs text-slate-500 dark:text-slate-400">
            Formato grafico fedele al modello delle adunanze infrasettimanali, pronto per la stampa A4 o salvataggio PDF.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          {/* Period Filter Dropdown */}
          {availablePeriods.length > 1 && (
            <div className="flex items-center gap-1.5 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg px-2.5 py-1 text-xs">
              <Filter className="w-3.5 h-3.5 text-slate-400" />
              <span className="text-slate-500 font-medium">Mese:</span>
              <select
                value={filterPeriod}
                onChange={e => setFilterPeriod(e.target.value)}
                className="bg-transparent font-bold text-slate-800 dark:text-slate-200 outline-none cursor-pointer text-xs"
              >
                <option value="all">Tutti i mesi ({data.meetings.length} adunanze)</option>
                {availablePeriods.map(([key, label]) => (
                  <option key={key} value={key}>
                    {label}
                  </option>
                ))}
              </select>
            </div>
          )}

          <button
            onClick={handleExportExcel}
            className="px-3 py-1.5 text-xs font-semibold bg-emerald-50 text-emerald-700 hover:bg-emerald-100 dark:bg-emerald-950/60 dark:text-emerald-300 border border-emerald-300 dark:border-emerald-800 rounded-lg flex items-center gap-1.5 transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            Esporta Excel (.xlsx)
          </button>
          <button
            onClick={handleDownloadAllS89}
            className="px-3.5 py-1.5 text-xs font-bold bg-amber-500 hover:bg-amber-600 text-white rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
            title="Scarica tutti i foglietti S-89 compilati per le adunanze selezionate"
          >
            <FileText className="w-4 h-4" />
            Scarica Foglietti S-89 (PDF)
          </button>
          <button
            onClick={handlePrint}
            className="px-4 py-1.5 text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
          >
            <Printer className="w-4 h-4" />
            Stampa / Salva in PDF
          </button>
        </div>
      </div>

      {/* The Printable Page Sheet */}
      <VitaEMinisteroPrintTemplate meetings={displayedMeetings} participants={participants} congregation={data.congregationName || 'Roccastrada'} />

    </div>
  );
}
