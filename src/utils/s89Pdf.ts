import { jsPDF } from 'jspdf';

export interface S89Item {
  id?: string;
  studentName: string;
  assistantName?: string;
  dateLabel: string; // e.g. "5 ottobre 2026"
  partNumber: string | number; // e.g. "3" or "4"
  partTitle?: string; // e.g. "Lettura biblica" or "Iniziare una conversazione"
  room?: 'main' | 'aux1' | 'aux2'; // Sala principale, Sala secondaria 1, Sala secondaria 2
  isSent?: boolean;
}

/**
 * Pulisce il nome per renderlo sicuro in un nome di file e restituisce il nome file formattato
 * con il nome dello studente: es. "Mario Rossi - S-89.pdf"
 */
export function getS89FileName(studentName: string, partNumber?: string | number): string {
  const cleanName = (studentName || 'Studente')
    .trim()
    .replace(/[/\\?%*:|"<>]/g, '')
    .trim();
  
  if (partNumber) {
    return `${cleanName} - S-89 (Parte ${partNumber}).pdf`;
  }
  return `${cleanName} - S-89.pdf`;
}

/** Formato di stampa del foglietto S-89: 7,5 × 9,5 cm. */
export const S89_FORMAT: [number, number] = [75, 95];

/**
 * Disegna una singola pagina del foglietto S-89 (7,5 × 9,5 cm) su un documento jsPDF
 */
export function renderS89Page(doc: jsPDF, item: S89Item): void {
  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 5;
  const rightEdge = pageWidth - margin;
  const dotted = (from: number, yy: number) => {
    if (from >= rightEdge) return;
    doc.setDrawColor(150, 150, 150);
    doc.setLineWidth(0.2);
    doc.setLineDashPattern([0.6, 0.9], 0);
    doc.line(from, yy + 0.4, rightEdge, yy + 0.4);
    doc.setLineDashPattern([], 0);
  };

  // Intestazione
  doc.setTextColor(20, 20, 20);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(9.5);
  doc.text('PARTE PER L’ADUNANZA', pageWidth / 2, 9, { align: 'center' });
  doc.text('VITA CRISTIANA E MINISTERO', pageWidth / 2, 13.5, { align: 'center' });

  let y = 21;
  const lineSpacing = 7.5;

  const drawField = (label: string, val?: string) => {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    const fullLabel = label + ': ';
    doc.text(fullLabel, margin, y);
    const valueX = margin + doc.getTextWidth(fullLabel) + 0.5;
    const cleanVal = val?.trim() || '';
    if (cleanVal) {
      // Valori lunghi: carattere più piccolo, poi troncati, per restare su una riga.
      doc.setFont('helvetica', 'normal');
      let size = 8.5;
      while (size > 6 && doc.setFontSize(size).getTextWidth(cleanVal) > rightEdge - valueX) size -= 0.5;
      const text = doc.splitTextToSize(cleanVal, rightEdge - valueX)[0];
      doc.text(text, valueX, y);
      dotted(valueX + doc.getTextWidth(text) + 1, y);
    } else dotted(valueX, y);
    y += lineSpacing;
  };

  drawField('Nome e cognome', item.studentName);
  drawField('Assistente', item.assistantName);
  drawField('Data', item.dateLabel);
  const partDisplay = item.partTitle
    ? `${item.partNumber} - ${item.partTitle}`
    : item.partNumber
    ? String(item.partNumber)
    : '';
  drawField('Parte n.', partDisplay);

  // Da svolgere nella:
  y += 2;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.text('Da svolgere nella:', margin, y);
  y += 4.5;

  const activeRoom = item.room || 'main';
  const rooms: Array<{ key: 'main' | 'aux1' | 'aux2'; label: string }> = [
    { key: 'main', label: 'Sala principale' },
    { key: 'aux1', label: 'Sala secondaria 1' },
    { key: 'aux2', label: 'Sala secondaria 2' },
  ];

  rooms.forEach(r => {
    doc.setDrawColor(30, 30, 30);
    doc.setLineWidth(0.3);
    doc.rect(margin + 0.8, y - 2.8, 3.2, 3.2);
    if (activeRoom === r.key) {
      doc.setLineWidth(0.5);
      doc.line(margin + 1.4, y - 1.1, margin + 2.1, y - 0.3);
      doc.line(margin + 2.1, y - 0.3, margin + 3.4, y - 2.2);
    }
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text(r.label, margin + 6, y);
    y += 4.8;
  });

  // Nota per lo studente
  y += 2;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(6.5);
  const noteParagraph =
    'La fonte e la lezione che riguardano la tua parte sono indicate nella Guida per l’adunanza Vita e ministero. Ripassa le Istruzioni per l’adunanza Vita cristiana e ministero (S-38) relative alla tua parte.';
  doc.text(doc.splitTextToSize('Nota per lo studente: ' + noteParagraph, rightEdge - margin), margin, y, { lineHeightFactor: 1.3 });

  // Codice modulo in basso a sinistra
  doc.setFontSize(6);
  doc.setTextColor(60, 60, 60);
  doc.text('S-89-I   11/23', margin, doc.internal.pageSize.getHeight() - 4);
}

/**
 * Genera il file PDF del singolo foglietto S-89 e ne avvia il download.
 * Il nome del file conterrà il nome dello studente (es. "Mario Rossi - S-89.pdf")
 */
export function downloadS89Pdf(item: S89Item): void {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: S89_FORMAT });
  renderS89Page(doc, item);
  doc.save(getS89FileName(item.studentName, item.partNumber));
}

/**
 * Genera un unico file PDF multipagina con tutti i foglietti S-89 indicati
 */
export function downloadCombinedS89Pdf(items: S89Item[], customFileName?: string): void {
  if (!items || items.length === 0) return;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: S89_FORMAT });
  items.forEach((item, index) => {
    if (index > 0) doc.addPage(S89_FORMAT, 'portrait');
    renderS89Page(doc, item);
  });
  doc.save(customFileName || `Foglietti_S-89_Vita_e_Ministero.pdf`);
}
