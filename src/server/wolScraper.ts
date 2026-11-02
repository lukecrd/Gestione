import * as cheerio from 'cheerio';
import JSZip from 'jszip';
import type { VitaEMinisteroMeeting, MinisteroPart, VitaCristianaPart } from '../types.js';
import { inferPartTypeIds } from '../utils/ministeroPartTypes.js';

// wol.jw.org blocks server requests with an AWS WAF challenge, so we use the
// official jw-cdn publication API and parse the EPUB of each bimonthly issue.
const PUB_MEDIA_API = 'https://b.jw-cdn.org/apis/pub-media/GETPUBMEDIALINKS';
const LANG = 'I';
const FIRST_YEAR = 2020;

const IT_MONTHS = [
  'gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno',
  'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'
];

// Each issue covers two months and starts on an odd month (01, 03, ..., 11)
const ISSUE_START_MONTHS = [1, 3, 5, 7, 9, 11];

// In-memory cache of successfully parsed issues: key = "YYYYMM"
const issueCache = new Map<string, VitaEMinisteroMeeting[]>();

function issueCode(year: number, month: number): string {
  const start = month % 2 === 1 ? month : month - 1;
  return `${year}${String(start).padStart(2, '0')}`;
}

/**
 * Returns the EPUB download URL of an issue, or null if it is not published.
 */
async function getIssueEpubUrl(issue: string): Promise<string | null> {
  const url = `${PUB_MEDIA_API}?pub=mwb&issue=${issue}&langwritten=${LANG}&fileformat=EPUB&output=json`;
  const res = await fetch(url);
  if (res.status === 404) return null;
  if (!res.ok) {
    throw new Error(`Errore caricamento da jw.org (${res.status}): fascicolo ${issue}`);
  }
  const data = await res.json();
  return data?.files?.[LANG]?.EPUB?.[0]?.file?.url ?? null;
}

/**
 * Parses date header like "5-11 OTTOBRE" or "28 settembre – 4 ottobre" or "26 ottobre – 1º novembre"
 */
function parseWeekHeader(rawTxt: string, year: number): { day: number; month: number; dateLabel: string; iso: string } | null {
  const clean = rawTxt.toLowerCase().replace(/[º°]/g, '').replace(/\s+/g, ' ').trim();

  // Format 1: "28 settembre – 4 ottobre"
  const m1 = clean.match(/^(\d+)\s+([a-z]+)/);
  if (m1) {
    const day = parseInt(m1[1], 10);
    const mName = m1[2];
    const mIdx = IT_MONTHS.indexOf(mName);
    if (mIdx !== -1) {
      return {
        day,
        month: mIdx + 1,
        dateLabel: `${day} ${IT_MONTHS[mIdx]} ${year}`,
        iso: `${year}-${String(mIdx + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
      };
    }
  }

  // Format 2: "5-11 ottobre"
  const m2 = clean.match(/^(\d+)\s*-\s*(\d+)\s+([a-z]+)/);
  if (m2) {
    const day = parseInt(m2[1], 10);
    const mName = m2[3];
    const mIdx = IT_MONTHS.indexOf(mName);
    if (mIdx !== -1) {
      return {
        day,
        month: mIdx + 1,
        dateLabel: `${day} ${IT_MONTHS[mIdx]} ${year}`,
        iso: `${year}-${String(mIdx + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
      };
    }
  }

  return null;
}

/**
 * Parses the XHTML of a single meeting week (one chapter of the EPUB)
 */
export function parseMeetingHtml(html: string, year: number): VitaEMinisteroMeeting | null {
  const $ = cheerio.load(html);
  // Drop the Bible text footnotes appended at the end of each chapter
  $('.groupExt, aside').remove();

  const rawHeaderDate = $('h1').first().text().replace(/\s+/g, ' ').trim();
  const dateInfo = parseWeekHeader(rawHeaderDate, year);
  if (!dateInfo) {
    return null;
  }

  const reading = $('header h2').first().text().replace(/\s+/g, ' ').trim();

  // TESORI
  let tesori1Title = '';
  let tesori1Minutes = 10;
  let tesoriGemmeTitle = 'Gemme spirituali';
  let tesoriGemmeMinutes = 10;
  let tesoriLetturaTitle = 'Lettura biblica';
  let tesoriLetturaMinutes = 4;

  $('h3').each((_, el) => {
    const txt = $(el).text().replace(/\s+/g, ' ').trim();
    const nextTxt = $(el).next('div').text() || $(el).next().text() || '';
    if (txt.startsWith('1.')) {
      tesori1Title = txt.replace(/^1\.\s*/, '');
      const m = nextTxt.match(/\((\d+)\s*min\)/);
      if (m) tesori1Minutes = parseInt(m[1], 10);
    } else if (txt.startsWith('2.')) {
      tesoriGemmeTitle = txt.replace(/^2\.\s*/, '');
      const m = nextTxt.match(/\((\d+)\s*min\)/);
      if (m) tesoriGemmeMinutes = parseInt(m[1], 10);
    } else if (txt.startsWith('3.')) {
      tesoriLetturaTitle = txt.replace(/^3\.\s*/, '');
      const m = nextTxt.match(/\((\d+)\s*min\)/);
      if (m) tesoriLetturaMinutes = parseInt(m[1], 10);
      // Try to extract verses, e.g. "Ger 40:1-10"
      const verseMatch = nextTxt.match(/\((\d+\s*min\)\s*)([^(]+)/i);
      if (verseMatch && verseMatch[2].trim()) {
        const verses = verseMatch[2].trim();
        tesoriLetturaTitle = `Lettura biblica (${verses})`;
      }
    }
  });

  // SONGS, EFFICACI NEL MINISTERO & VITA CRISTIANA
  const ministeroParts: MinisteroPart[] = [];
  const vitaCristianaParts: VitaCristianaPart[] = [];
  let canticoIniziale = '';
  let canticoIntermedio = '';
  let studioBiblicoTitle = 'Studio biblico di congregazione';
  let studioBiblicoMinutes = 30;
  let canticoFinale = '';

  let currentSection = ''; // '' (intro) | 'tesori' | 'ministero' | 'vita' | 'conclusione'

  $('h2, h3, p').each((_, el) => {
    const tag = (el as any).tagName;
    const text = $(el).text().replace(/\s+/g, ' ').trim();
    if (tag === 'h2' && text.includes('TESORI DELLA PAROLA')) {
      currentSection = 'tesori';
      return;
    }
    if (tag === 'h2' && text.includes('EFFICACI NEL MINISTERO')) {
      currentSection = 'ministero';
      return;
    }
    if (tag === 'h2' && text.includes('VITA CRISTIANA')) {
      currentSection = 'vita';
      return;
    }

    const songMatch = text.match(/^Cantico\s*(\d+)/i);
    if (songMatch) {
      const song = `Cantico ${songMatch[1]}`;
      if (currentSection === '' && !canticoIniziale) canticoIniziale = song;
      else if (currentSection === 'vita' && !canticoIntermedio) canticoIntermedio = song;
      else if (currentSection === 'conclusione' && !canticoFinale) canticoFinale = song;
      return;
    }

    if (tag !== 'h3') return;

    if (currentSection === 'ministero') {
      const matchNum = text.match(/^(\d+)\.\s*(.+)/);
      if (matchNum) {
        const num = parseInt(matchNum[1], 10);
        const title = matchNum[2];
        const nextDivText = $(el).next('div').text();
        const minMatch = nextDivText.match(/\((\d+)\s*min\)/);
        const minutes = minMatch ? parseInt(minMatch[1], 10) : 3;
        const isTalkOnly =
          nextDivText.toLowerCase().includes('discorso') &&
          !nextDivText.toLowerCase().includes('dimostrazione');

        ministeroParts.push({
          id: 'mp_wol_' + Math.random().toString(36).substring(2, 8),
          number: num,
          title,
          minutes,
          studentId: '',
          assistantId: '',
          hasAssistant: !isTalkOnly,
          partTypeIds: inferPartTypeIds(minutes, !isTalkOnly),
        });
      }
    }

    if (currentSection === 'vita') {
      if (text.includes('Studio biblico di congregazione')) {
        studioBiblicoTitle = 'Studio biblico di congregazione';
        const m = $(el).next('div').text().match(/\((\d+)\s*min\)/);
        if (m) studioBiblicoMinutes = parseInt(m[1], 10);
        return;
      }
      if (text.includes('Commenti conclusivi')) {
        // Older layouts put the final song in the heading itself
        const finalSongMatch = text.match(/Cantico\s*(\d+)/i);
        if (finalSongMatch) {
          canticoFinale = `Cantico ${finalSongMatch[1]}`;
        }
        currentSection = 'conclusione';
        return;
      }
      const matchNum = text.match(/^(\d+)\.\s*(.+)/);
      if (matchNum) {
        const num = parseInt(matchNum[1], 10);
        const title = matchNum[2];
        const m = $(el).next('div').text().match(/\((\d+)\s*min\)/);
        const minutes = m ? parseInt(m[1], 10) : 15;
        vitaCristianaParts.push({
          id: 'vcp_wol_' + Math.random().toString(36).substring(2, 8),
          number: num,
          title,
          minutes,
          speakerId: '',
        });
      }
    }
  });

  return {
    id: `vm_wol_${dateInfo.iso}`,
    dateStr: dateInfo.iso,
    dateLabel: dateInfo.dateLabel,
    bibleReading: reading || 'LETTURA BIBLICA',
    canticoIniziale,
    tesori1Title,
    tesori1Minutes,
    tesoriGemmeTitle,
    tesoriGemmeMinutes,
    tesoriLetturaTitle,
    tesoriLetturaMinutes,
    ministeroParts,
    canticoIntermedio,
    vitaCristianaParts,
    studioBiblicoTitle,
    studioBiblicoMinutes,
    canticoFinale,
  };
}

/**
 * Downloads and parses all meetings of a bimonthly issue.
 * Returns null if the issue is not published yet.
 */
async function fetchIssueMeetings(year: number, startMonth: number): Promise<VitaEMinisteroMeeting[] | null> {
  const issue = issueCode(year, startMonth);
  if (issueCache.has(issue)) {
    return issueCache.get(issue)!;
  }

  const epubUrl = await getIssueEpubUrl(issue);
  if (!epubUrl) return null;

  const res = await fetch(epubUrl);
  if (!res.ok) {
    throw new Error(`Errore download EPUB da jw.org (${res.status}): fascicolo ${issue}`);
  }
  const zip = await JSZip.loadAsync(await res.arrayBuffer());

  const chapterNames = Object.keys(zip.files)
    .filter(name => name.endsWith('.xhtml') && !name.endsWith('-extracted.xhtml'))
    .sort();

  const meetings: VitaEMinisteroMeeting[] = [];
  for (const name of chapterNames) {
    const html = await zip.files[name].async('string');
    if (!html.includes('docClass-106')) continue; // only weekly meeting chapters
    try {
      const meeting = parseMeetingHtml(html, year);
      if (meeting && !meetings.some(m => m.dateStr === meeting.dateStr)) {
        meetings.push(meeting);
      }
    } catch (e) {
      console.error(`Failed to parse ${name} of issue ${issue}:`, e);
    }
  }

  meetings.sort((a, b) => (a.dateStr || '').localeCompare(b.dateStr || ''));
  if (meetings.length) issueCache.set(issue, meetings);
  return meetings;
}

/**
 * Returns available years of mwb from jw.org
 */
export async function getAvailableYears(): Promise<{ year: number; title: string }[]> {
  let lastYear = new Date().getFullYear();
  try {
    // Next year's first issue is usually published a couple of months in advance
    if (await getIssueEpubUrl(`${lastYear + 1}01`)) lastYear++;
  } catch (err) {
    console.error('getAvailableYears error:', err);
  }

  const years: { year: number; title: string }[] = [];
  for (let y = lastYear; y >= FIRST_YEAR; y--) {
    years.push({ year: y, title: `Guida per l’adunanza Vita e ministero ${y}` });
  }
  return years;
}

/**
 * Fetch all meetings for a given month and year
 * @param month 1-12 (1 = Gennaio, 10 = Ottobre)
 */
export async function fetchMonthMeetings(year: number, month: number): Promise<VitaEMinisteroMeeting[]> {
  const meetings = await fetchIssueMeetings(year, month);
  if (!meetings) return [];
  return meetings.filter(m => parseInt((m.dateStr || '').split('-')[1], 10) === month);
}

/**
 * Fetch all meetings for all available months in a year
 */
export async function fetchAllYearMeetings(year: number): Promise<VitaEMinisteroMeeting[]> {
  const allMeetings: VitaEMinisteroMeeting[] = [];

  for (const startMonth of ISSUE_START_MONTHS) {
    try {
      const meetings = await fetchIssueMeetings(year, startMonth);
      for (const meeting of meetings || []) {
        if (!allMeetings.some(m => m.dateStr === meeting.dateStr)) {
          allMeetings.push(meeting);
        }
      }
    } catch (err) {
      console.warn(`Issue ${issueCode(year, startMonth)} not available:`, err);
    }
  }

  allMeetings.sort((a, b) => (a.dateStr || '').localeCompare(b.dateStr || ''));
  return allMeetings;
}
