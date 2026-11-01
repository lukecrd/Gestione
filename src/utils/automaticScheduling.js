import { mensileWeekStart } from './mensileWeek';
import { personNameKey } from './peopleRegistry';
import { inferPartTypeIds, isParticipantEligibleForPartTypes } from './ministeroPartTypes';

const validName = name => !!name?.trim() && !['—', '---', '-'].includes(name.trim());

export function dateKey(value) {
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
}

export function weeklyDates(value) {
  const monday = mensileWeekStart(value);
  if (Number.isNaN(monday.getTime())) return [];
  const wednesday = new Date(monday.getTime());
  const sunday = new Date(monday.getTime());
  wednesday.setDate(wednesday.getDate() + 2);
  sunday.setDate(sunday.getDate() + 6);
  return [dateKey(wednesday), dateKey(sunday)];
}

/**
 * One identity across the registry and programme-specific IDs.
 * Names also cover older imported rows and unlinked participants.
 * @param {import('../types').StateData} state
 */
/** @param {import('../types').StateData} state
 * @param {import('../types').SchedulingPrograms} programs
 * @param {{ program?: string, dates?: string[] }} excluded */
export function createSchedulingLedger(state, programs = {}, excluded = {}) {
  const people = state.people || [];
  const byId = new Map(people.map(person => [person.id, person]));
  const byName = new Map();
  for (const person of people) {
    const key = personNameKey(person.name);
    const matches = byName.get(key) || [];
    matches.push(person);
    byName.set(key, matches);
  }
  const identity = record => {
    if (!record) return '';
    const person = byId.get(record.personId || record.id);
    if (person) return 'name:' + personNameKey(person.name);
    return validName(record.name) ? 'name:' + personNameKey(record.name) : '';
  };
  const occupied = new Map();
  const reserve = (record, dates) => {
    const key = identity(record);
    if (!key) return;
    for (const date of typeof dates === 'string' ? [dates] : dates) {
      if (!date) continue;
      if (!occupied.has(date)) occupied.set(date, new Set());
      occupied.get(date).add(key);
    }
  };
  const isBusy = (record, dates) => {
    const key = identity(record);
    return (typeof dates === 'string' ? [dates] : dates).some(date => occupied.get(date)?.has(key));
  };
  const isUnavailable = (record, dates) => {
    const person = byId.get(record?.personId || record?.id);
    const matches = person ? [person] : byName.get(personNameKey(record?.name || '')) || [];
    const keys = typeof dates === 'string' ? [dates] : dates;
    return matches.some(p => keys.some(date => state.unavail?.[p.id]?.includes(date)));
  };
  const excludedDates = new Set(excluded.dates || []);
  // Every assignment, for the weekly multi-role check (see weeklyRoleConflicts).
  const entries = [];
  const add = (program, record, dates, role, busy = true) => {
    const list = (typeof dates === 'string' ? [dates] : dates).filter(Boolean);
    const key = identity(record);
    const name = byId.get(record?.personId || record?.id)?.name || byName.get(personNameKey(record?.name || ''))?.[0].name || record?.name;
    if (key) list.forEach(date => entries.push({ key, name, program, role, date }));
    if (busy) reserve(record, list.filter(date => !(excluded.program === program && excludedDates.has(date))));
  };
  const named = name => ({ name });
  const monthly = row => {
    if (row.special || row.placeholder) return;
    const dates = weeklyDates(row.date || row.dateStr);
    const slots = [...(row.ingresso || []).map(n => [n, 'Usciere ingresso']), [row.auditorium, 'Usciere auditorium'],
      ...(row.microfoni || []).map(n => [n, 'Microfoni']), ...(row.audioVideo || []).map(n => [n, 'Audio/Video'])];
    slots.filter(([name]) => validName(name)).forEach(([name, role]) => add('mensile', named(name), dates, role));
  };
  // Most recently saved archive wins; the active programme overrides its dates.
  const monthlyRows = new Map();
  [...(state.mensileArchives || [])].sort((a, b) => (a.savedAt || '').localeCompare(b.savedAt || '')).forEach(archive =>
    (archive.rows || []).forEach(row => monthlyRows.set(dateKey(mensileWeekStart(row.dateStr)), row)));
  (programs.mensileRows || []).forEach(row => monthlyRows.set(dateKey(mensileWeekStart(row.date)), row));
  monthlyRows.forEach(monthly);
  (programs.domenicaRows || []).forEach(row => {
    if (row.special || row.placeholder) return;
    [[row.presidente, 'Presidente'], [row.lettore, 'Lettore'], [row.oratore, 'Oratore']].filter(([name]) => validName(name))
      .forEach(([name, role]) => add('domenica', named(name), dateKey(row.date), role));
  });
  const vm = state.vitaEMinistero;
  const vmById = new Map((vm?.participants || []).map(p => [p.id, p]));
  (vm?.meetings || []).forEach(meeting => {
    if (meeting.isSpecialEvent) return;
    vitaAssignments(meeting).forEach(([id, role]) => add('vitaEMinistero', vmById.get(id), meeting.dateStr, role));
  });
  const field = state.servizioCampo;
  const fieldById = new Map((field?.conductors || []).map(p => [p.id, p]));
  (field?.schedule || []).forEach(meeting => {
    if (meeting.isActive !== false) add('servizioCampo', fieldById.get(meeting.conductorId), meeting.dateStr, 'Conduttore');
  });
  const publicWork = state.operaPubblica;
  const publicById = new Map((publicWork?.participants || []).map(p => [p.id, p]));
  (publicWork?.schedule || []).forEach(day => {
    if (day.turno1Active !== false) (day.turno1 || []).forEach(id => add('operaPubblica', publicById.get(id), day.dateStr, 'Turno'));
    if (day.turno2Active !== false) (day.turno2 || []).forEach(id => add('operaPubblica', publicById.get(id), day.dateStr, 'Turno'));
  });
  // Outgoing talks only feed the multi-role check: they never blocked automatic scheduling.
  (state.discorsiEsterni?.uscite || []).forEach(u => add('discorsiEsterni', named(u.oratore), u.data, 'Discorso esterno', false));
  return { identity, reserve, isBusy, isUnavailable, entries };
}

/** Monday (YYYY-MM-DD) of the Monday–Sunday week containing the date. */
export function weekKey(value) {
  const key = dateKey(value);
  if (!key) return '';
  const date = new Date(key + 'T12:00:00');
  date.setDate(date.getDate() - (date.getDay() + 6) % 7);
  return dateKey(date);
}

/** People holding two or more different roles (across all programmes) in the same week. */
export function weeklyRoleConflicts(state, programs = {}) {
  const groups = new Map();
  for (const entry of createSchedulingLedger(state, programs).entries) {
    const week = weekKey(entry.date);
    if (!week) continue;
    const id = week + '|' + entry.key;
    const group = groups.get(id) || { week, name: entry.name, roles: new Map() };
    group.roles.set(entry.program + '|' + entry.role, { program: entry.program, role: entry.role });
    groups.set(id, group);
  }
  return [...groups.values()].filter(g => g.roles.size > 1)
    .map(g => ({ week: g.week, name: g.name, roles: [...g.roles.values()] }))
    .sort((a, b) => a.week.localeCompare(b.week) || a.name.localeCompare(b.name, 'it'));
}

/** [id, label, rotation key]: the key groups slots of the same part (both prayers share one). */
function vitaAssignments(meeting) {
  if (meeting.isSpecialEvent) return [];
  return [
    [meeting.presidenteId, 'Presidente', 'presidente'], [meeting.preghieraInizialeId, 'Preghiera iniziale', 'preghiera'],
    [meeting.tesori1SpeakerId, 'Discorso Tesori', 'tesoriDiscorso'], [meeting.tesoriGemmeSpeakerId, 'Gemme spirituali', 'tesoriGemme'],
    [meeting.tesoriLetturaReaderId, 'Lettura biblica', 'tesoriLettura'],
    ...(meeting.ministeroParts || []).flatMap(p => [[p.studentId, 'Studente ministero', 'ministeroStudente'], [p.hasAssistant ? p.assistantId : '', 'Assistente ministero', 'ministeroAssistente']]),
    ...(meeting.vitaCristianaParts || []).map(p => [p.speakerId, 'Parte Vita cristiana', 'vitaCristianaParti']),
    ...(meeting.studioBiblicoType === 'discorsoSorvegliante' ? [] : [[meeting.studioBiblicoConductorId, 'Conduttore studio biblico', 'studioBiblicoConduttore'], [meeting.studioBiblicoReaderId, 'Lettore studio biblico', 'studioBiblicoLettore']]),
    [meeting.preghieraFinaleId, 'Preghiera finale', 'preghiera'],
  ].filter(([id]) => id);
}

export function vitaAssignmentIds(meeting) {
  return vitaAssignments(meeting).map(([id]) => id);
}

const MONTHLY_UNIQUE_ROLES = [['presidenteId', 'Presidente'], ['studioBiblicoConductorId', 'Conduttore studio biblico']];

/** Same person twice in one month as Presidente or as Conduttore studio biblico. */
export function vitaMonthlyRepeats(meetings, participants) {
  const names = new Map((participants || []).map(p => [p.id, p.name]));
  const seen = new Map();
  const warnings = [];
  [...(meetings || [])].filter(m => m.dateStr && !m.isSpecialEvent).sort((a, b) => a.dateStr.localeCompare(b.dateStr))
    .forEach(m => MONTHLY_UNIQUE_ROLES.forEach(([field, label]) => {
      const id = field === 'studioBiblicoConductorId' && m.studioBiblicoType === 'discorsoSorvegliante' ? '' : m[field];
      if (!id) return;
      const key = m.dateStr.slice(0, 7) + '|' + field + '|' + id;
      const first = seen.get(key);
      if (first) warnings.push(`${m.dateStr}: ${names.get(id) || id} è già ${label} il ${first} nello stesso mese.`);
      else seen.set(key, m.dateStr);
    }));
  return warnings;
}

const isBisogniLocali = part => /bisogni\s+locali/i.test(part.title || '');

/** Fill scarce roles first; relocate a previous choice only to fill a blocked role. */
export function assignUniqueSlots(tasks, identity) {
  const assignments = new Map();
  const owners = new Map();
  const ordered = tasks.map((task, index) => ({ ...task, index }))
    .sort((a, b) => a.candidates.length - b.candidates.length || a.index - b.index);
  const byId = new Map(ordered.map(task => [task.id, task]));
  function place(task, seen) {
    // Keep the oldest free candidate before trying to displace another role.
    for (const candidate of task.candidates) {
      const key = identity(candidate);
      if (!key || seen.has(key) || owners.has(key)) continue;
      owners.set(key, task.id);
      assignments.set(task.id, candidate);
      return true;
    }
    for (const candidate of task.candidates) {
      const key = identity(candidate);
      if (!key || seen.has(key)) continue;
      seen.add(key);
      const previous = owners.get(key);
      if (previous !== undefined && place(byId.get(previous), seen)) {
        owners.set(key, task.id);
        assignments.set(task.id, candidate);
        return true;
      }
    }
    return false;
  }
  ordered.forEach(task => place(task, new Set()));
  return assignments;
}

/**
 * Regenerate only the selected month. Rotation uses actual earlier VM assignments,
 * ignores future assignments, and updates after each generated meeting.
 * @param {import('../types').StateData} state
 * @param {import('../types').VitaEMinisteroData} data
 * @param {import('../types').SchedulingPrograms} programs
 */
export function generateVitaAssignments(state, data, year, month, programs = {}) {
  const prefix = year + '-' + String(month + 1).padStart(2, '0') + '-';
  const targets = (data.meetings || []).filter(m => m.dateStr?.startsWith(prefix) && !m.isSpecialEvent)
    .sort((a, b) => a.dateStr.localeCompare(b.dateStr) || a.id.localeCompare(b.id));
  const targetIds = new Set(targets.map(m => m.id));
  const ledger = createSchedulingLedger({ ...state, vitaEMinistero: data }, programs, {
    program: 'vitaEMinistero', dates: targets.map(m => m.dateStr),
  });
  const byId = new Map((data.participants || []).map(p => [p.id, p]));
  const history = [];
  (data.meetings || []).filter(m => !targetIds.has(m.id)).forEach(m =>
    vitaAssignments(m).forEach(([id, , role]) => {
      const key = ledger.identity(byId.get(id));
      if (key && m.dateStr) history.push({ key, role, date: m.dateStr });
    }));
  const generated = new Map();
  const warnings = [];
  // Presidente / conduttore studio già usati in questo mese: in coda ai candidati, non esclusi.
  const usedThisMonth = { presidente: new Set(), studio: new Set() };
  const markUsed = meeting => {
    if (meeting.presidenteId) usedThisMonth.presidente.add(ledger.identity(byId.get(meeting.presidenteId)));
    if (meeting.studioBiblicoConductorId && meeting.studioBiblicoType !== 'discorsoSorvegliante')
      usedThisMonth.studio.add(ledger.identity(byId.get(meeting.studioBiblicoConductorId)));
  };
  for (const original of targets) {
    if (state.special?.[original.dateStr]) {
      warnings.push(original.dateStr + ': data speciale, assegnazioni non rigenerate.');
      markUsed(original);
      continue;
    }
    const m = { ...original,
      ministeroParts: (original.ministeroParts || []).map(p => ({ ...p })),
      vitaCristianaParts: (original.vitaCristianaParts || []).map(p => ({ ...p })),
    };
    // Per-part rotation: whoever has gone longest without this part goes first,
    // then whoever has gone longest without any part, then fewest assignments.
    const stats = new Map();
    for (const assignment of history) {
      if (assignment.date >= m.dateStr) continue;
      const current = stats.get(assignment.key) || { last: '', count: 0, roles: new Map() };
      if (current.last < assignment.date) current.last = assignment.date;
      if ((current.roles.get(assignment.role) || '') < assignment.date) current.roles.set(assignment.role, assignment.date);
      current.count++;
      stats.set(assignment.key, current);
    }
    const empty = { last: '', count: 0, roles: new Map() };
    const compareFor = role => (a, b) => {
      const sa = stats.get(ledger.identity(a)) || empty;
      const sb = stats.get(ledger.identity(b)) || empty;
      return (sa.roles.get(role) || '').localeCompare(sb.roles.get(role) || '') ||
        sa.last.localeCompare(sb.last) || sa.count - sb.count || a.name.localeCompare(b.name, 'it');
    };
    if (!m.tesori1Title) m.tesori1SpeakerId = '';
    // Uscieri of the same evening (monthly programme, midweek date): not presidente too.
    const evening = weeklyDates(m.dateStr)[0];
    const uscieri = new Set(ledger.entries.filter(e => e.program === 'mensile' && e.date === evening && e.role.startsWith('Usciere')).map(e => e.key));
    const avoid = (id, p) => (usedThisMonth[id]?.has(ledger.identity(p)) ? 1 : 0) + (id === 'presidente' && uscieri.has(ledger.identity(p)) ? 2 : 0);
    const tasks = [];
    const addTask = (id, label, role, set, gender, partTypes) => {
      tasks.push({ id, label, role, set, candidates: (data.participants || []).filter(p =>
        p.roles?.[role] && (!gender || p.gender === gender) &&
        (!partTypes || isParticipantEligibleForPartTypes(p.roles, partTypes)) &&
        !ledger.isBusy(p, m.dateStr) && !ledger.isUnavailable(p, m.dateStr)).sort((a, b) =>
        avoid(id, a) - avoid(id, b) || compareFor(role)(a, b)) });
    };
    addTask('presidente', 'Presidente', 'presidente', p => { m.presidenteId = p; }, 'M');
    addTask('iniziale', 'Preghiera iniziale', 'preghiera', p => { m.preghieraInizialeId = p; }, 'M');
    if (m.tesori1Title) addTask('tesori', 'Discorso dei Tesori', 'tesoriDiscorso', p => { m.tesori1SpeakerId = p; }, 'M');
    addTask('gemme', 'Gemme spirituali', 'tesoriGemme', p => { m.tesoriGemmeSpeakerId = p; }, 'M');
    addTask('lettura', 'Lettura biblica', 'tesoriLettura', p => {
      if (p !== m.tesoriLetturaReaderId) m.tesoriLetturaSent = false;
      m.tesoriLetturaReaderId = p;
    }, 'M');
    m.ministeroParts.forEach((part, i) => {
      addTask('studente-' + i, 'Parte ' + part.number + ': studente', 'ministeroStudente', p => {
        if (p !== part.studentId) part.isSent = false;
        part.studentId = p;
      }, undefined, part.partTypeIds?.length ? part.partTypeIds : inferPartTypeIds(part.minutes, part.hasAssistant));
      if (part.hasAssistant) addTask('assistente-' + i, 'Parte ' + part.number + ': assistente', 'ministeroAssistente', p => {
        if (p !== part.assistantId) part.isSent = false;
        part.assistantId = p;
      });
    });
    m.vitaCristianaParts.forEach((part, i) => {
      // Local needs are assigned by hand.
      if (isBisogniLocali(part)) { part.speakerId = ''; return; }
      addTask('vita-' + i, 'Parte ' + part.number, 'vitaCristianaParti', p => { part.speakerId = p; }, 'M');
    });
    if (m.studioBiblicoType !== 'discorsoSorvegliante') {
      addTask('studio', 'Studio biblico: conduttore', 'studioBiblicoConduttore', p => { m.studioBiblicoConductorId = p; }, 'M');
      addTask('studio-lettore', 'Studio biblico: lettore', 'studioBiblicoLettore', p => { m.studioBiblicoReaderId = p; }, 'M');
    }
    addTask('finale', 'Preghiera finale', 'preghiera', p => { m.preghieraFinaleId = p; }, 'M');
    const assigned = assignUniqueSlots(tasks, ledger.identity);
    for (const task of tasks) {
      const person = assigned.get(task.id);
      task.set(person?.id || '');
      if (person) {
        ledger.reserve(person, m.dateStr);
        history.push({ key: ledger.identity(person), role: task.role, date: m.dateStr });
      } else warnings.push(m.dateStr + ': ' + task.label + ' da assegnare (nessun nominativo idoneo libero).');
    }
    const president = byId.get(m.presidenteId);
    if (president && uscieri.has(ledger.identity(president)))
      warnings.push(m.dateStr + ': ' + president.name + ' è presidente e anche usciere nella stessa sera (nessun altro presidente libero).');
    markUsed(m);
    generated.set(m.id, m);
  }
  return { meetings: (data.meetings || []).map(m => generated.get(m.id) || m), warnings, generatedCount: generated.size };
}
