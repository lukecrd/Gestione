import { mensileWeekStart, mondayDates } from '../src/utils/mensileWeek.js';
import { readFileSync } from 'node:fs';
import ts from 'typescript';
import test from 'node:test';
import assert from 'node:assert/strict';
import { createSchedulingLedger, weeklyDates, generateVitaAssignments, vitaAssignmentIds, assignUniqueSlots, weeklyRoleConflicts, weekKey, vitaMonthlyRepeats } from '../src/utils/automaticScheduling.js';

const person = (id, name = id) => ({ id, name, gender: 'M', spouseId: null, roles: {} });
const participant = (id, roles = { tesoriLettura: true }, personId) => ({ id, name: id, gender: 'M', personId, roles });
const meeting = (id, dateStr, patch = {}) => ({ id, dateStr, dateLabel: dateStr, bibleReading: '', ministeroParts: [], vitaCristianaParts: [], studioBiblicoType: 'discorsoSorvegliante', ...patch });
const base = patch => ({ people: [], unavail: {}, special: {}, groups: { riassetto: '1', pulizie: '1' }, ...patch });
const generate = (participants, meetings, patch = {}, programs = {}) =>
  generateVitaAssignments(base(patch), { participants, meetings }, 2026, 9, programs);

test('one identity across all five programmes and legacy names', () => {
  const state = base({
    people: [person('p', 'Luca')],
    vitaEMinistero: { participants: [participant('vm', {}, 'p')], meetings: [meeting('v', '2026-10-09', { tesoriLetturaReaderId: 'vm' })] },
    servizioCampo: { conductors: [{ id: 'f', personId: 'p', name: 'Old name' }], schedule: [{ dateStr: '2026-10-10', conductorId: 'f', isActive: true }] },
    operaPubblica: { participants: [{ id: 'o', personId: 'p', name: 'Luca' }], schedule: [{ dateStr: '2026-10-13', turno1: ['o'], turno2: [] }] },
  });
  const ledger = createSchedulingLedger(state, {
    mensileRows: [{ date: new Date('2026-10-04T12:00:00'), ingresso: [' Luca ', '—'] }],
    domenicaRows: [{ date: new Date('2026-10-11T12:00:00'), lettore: 'LUCA' }],
  });
  for (const date of ['2026-10-07', '2026-10-09', '2026-10-10', '2026-10-11', '2026-10-13']) {
    assert.ok(ledger.isBusy({ personId: 'p', name: 'Renamed' }, date), date);
  }
  assert.equal(ledger.isBusy({ name: 'Luca' }, '2026-10-12'), false);
});

test('regeneration excludes only target programme dates', () => {
  const state = base({ people: [person('p', 'Luca')], operaPubblica: {
    participants: [{ id: 'o', personId: 'p', name: 'Luca' }],
    schedule: [{ dateStr: '2026-10-07', turno1: ['o'] }, { dateStr: '2026-11-04', turno1: ['o'] }],
  }, servizioCampo: { conductors: [{ id: 'f', name: 'Luca' }], schedule: [{ dateStr: '2026-10-07', conductorId: 'f' }] } });
  const ledger = createSchedulingLedger(state, {}, { program: 'operaPubblica', dates: ['2026-10-07'] });
  assert.ok(ledger.isBusy({ name: 'Luca' }, '2026-10-07'));
  assert.ok(ledger.isBusy({ name: 'Luca' }, '2026-11-04'));
  const withoutField = createSchedulingLedger({ ...state, servizioCampo: undefined }, {}, { program: 'operaPubblica', dates: ['2026-10-07'] });
  assert.equal(withoutField.isBusy({ name: 'Luca' }, '2026-10-07'), false);
});

test('inactive meetings, shifts, events and placeholder rows do not occupy a person', () => {
  const state = base({
    vitaEMinistero: { participants: [participant('p')], meetings: [meeting('e', '2026-10-04', { isSpecialEvent: true, tesoriLetturaReaderId: 'p' })] },
    servizioCampo: { conductors: [{ id: 'f', name: 'p' }], schedule: [{ dateStr: '2026-10-04', conductorId: 'f', isActive: false }] },
    operaPubblica: { participants: [{ id: 'o', name: 'p' }], schedule: [{ dateStr: '2026-10-04', turno1: ['o'], turno1Active: false }] },
  });
  const ledger = createSchedulingLedger(state, { domenicaRows: [{ date: '2026-10-04', presidente: 'p', placeholder: true }] });
  assert.equal(ledger.isBusy({ name: 'p' }, '2026-10-04'), false);
});

test('latest archived monthly row wins and active row overrides it', () => {
  const state = base({ mensileArchives: [
    { savedAt: '2026-09-01', rows: [{ dateStr: '2026-10-04', auditorium: 'Old' }] },
    { savedAt: '2026-09-02', rows: [{ dateStr: '2026-10-04', auditorium: 'New' }] },
  ] });
  let ledger = createSchedulingLedger(state);
  assert.ok(ledger.isBusy({ name: 'New' }, '2026-10-07'));
  assert.equal(ledger.isBusy({ name: 'Old' }, '2026-10-07'), false);
  ledger = createSchedulingLedger(state, { mensileRows: [{ date: '2026-10-04', auditorium: 'Active' }] });
  assert.ok(ledger.isBusy({ name: 'Active' }, '2026-10-07'));
  assert.equal(ledger.isBusy({ name: 'New' }, '2026-10-07'), false);
});

test('weekly dates include Wednesday across month boundaries', () => {
  assert.deepEqual(weeklyDates('2026-05-31'), ['2026-06-03', '2026-06-07']);
});

test('unavailability resolves both personId and legacy name', () => {
  const ledger = createSchedulingLedger(base({ people: [person('p', 'Luca')], unavail: { p: ['2026-10-07'] } }));
  assert.ok(ledger.isUnavailable({ name: ' LUCA ' }, '2026-10-07'));
  assert.ok(ledger.isUnavailable({ personId: 'p', name: 'Old' }, '2026-10-07'));
  assert.equal(ledger.isUnavailable({ name: 'Luca' }, '2026-10-08'), false);
});

test('never assigned first, oldest assignment next, future history ignored; chronological month rotation', () => {
  const p = ['A', 'B', 'C'].map(id => participant(id));
  const oldA = meeting('old-a', '2026-09-01', { tesoriLetturaReaderId: 'A' });
  const oldB = meeting('old-b', '2026-09-20', { tesoriLetturaReaderId: 'B' });
  const futureC = meeting('future', '2026-11-01', { tesoriLetturaReaderId: 'C' });
  const targets = [meeting('t3', '2026-10-21'), meeting('t1', '2026-10-07'), meeting('t2', '2026-10-14')];
  const result = generate(p, [oldA, oldB, futureC, ...targets]);
  assert.equal(result.meetings.find(m => m.id === 't1').tesoriLetturaReaderId, 'C');
  assert.equal(result.meetings.find(m => m.id === 't2').tesoriLetturaReaderId, 'A');
  assert.equal(result.meetings.find(m => m.id === 't3').tesoriLetturaReaderId, 'B');
  assert.deepEqual(result.meetings.slice(0, 3), [oldA, oldB, futureC]);
});

test('same-day conflicts across programmes and calendar availability exclude the oldest candidates', () => {
  const p = ['A', 'B', 'C'].map(id => participant(id, { tesoriLettura: true }, id));
  const state = {
    people: ['A', 'B', 'C'].map(id => person(id)),
    unavail: { B: ['2026-10-07'] },
    operaPubblica: { participants: [{ id: 'o', personId: 'A', name: 'A' }], schedule: [{ dateStr: '2026-10-07', turno1: ['o'] }] },
  };
  const result = generate(p, [meeting('t', '2026-10-07')], state);
  assert.equal(result.meetings[0].tesoriLetturaReaderId, 'C');
});

test('VM does not duplicate a person across roles or duplicate linked profiles', () => {
  const roles = { tesoriLettura: true, presidente: true, preghiera: true };
  const p = [participant('first', roles, 'p'), participant('duplicate', roles, 'p')];
  const result = generate(p, [meeting('m', '2026-10-07')], { people: [person('p', 'Same')] });
  assert.equal(vitaAssignmentIds(result.meetings[0]).length, 1);
  assert.ok(result.warnings.length > 0);
});

test('shortages clear previous assignments and reset sent slips rather than reusing someone', () => {
  const result = generate([], [meeting('m', '2026-10-07', {
    tesoriLetturaReaderId: 'old', tesoriLetturaSent: true,
    ministeroParts: [{ id: 'mp', number: 4, title: 'Parte', minutes: 3, hasAssistant: true, studentId: 'old', assistantId: 'old', isSent: true }],
  })]);
  assert.equal(result.meetings[0].tesoriLetturaReaderId, '');
  assert.equal(result.meetings[0].tesoriLetturaSent, false);
  assert.equal(result.meetings[0].ministeroParts[0].studentId, '');
  assert.equal(result.meetings[0].ministeroParts[0].assistantId, '');
  assert.equal(result.meetings[0].ministeroParts[0].isSent, false);
});

test('eligibility uses ministero types and gender for Bible reading', () => {
  const p = [
    { ...participant('female', { tesoriLettura: true }), gender: 'F' },
    participant('wrong', { ministeroStudente: true, ministeroTipiAbilitati: ['lunga', 'discorso'] }),
    participant('right', { ministeroStudente: true, ministeroTipiAbilitati: ['breve', 'dimostrazione'] }),
  ];
  const result = generate(p, [meeting('m', '2026-10-07', { ministeroParts: [
    { id: 'mp', number: 4, title: 'Parte', minutes: 3, hasAssistant: true, studentId: '' },
  ] })]);
  assert.equal(result.meetings[0].tesoriLetturaReaderId, '');
  assert.equal(result.meetings[0].ministeroParts[0].studentId, 'right');
});

test('overseer talk preserves title and dormant Bible-study selections without occupying them', () => {
  const original = meeting('m', '2026-10-07', {
    discorsoSorveglianteTitle: 'Titolo', studioBiblicoConductorId: 'A', studioBiblicoReaderId: 'B',
  });
  const result = generate([participant('A'), participant('B')], [original]);
  assert.equal(result.meetings[0].discorsoSorveglianteTitle, 'Titolo');
  assert.equal(result.meetings[0].studioBiblicoConductorId, 'A');
  assert.equal(result.meetings[0].studioBiblicoReaderId, 'B');
  assert.equal(result.meetings[0].tesoriLetturaReaderId, 'A');
  assert.deepEqual(vitaAssignmentIds(result.meetings[0]), ['A']);
  assert.equal(original.tesoriLetturaReaderId, undefined);
});

test('two ordinary VM meetings on the same day cannot reuse a person', () => {
  const result = generate([participant('A'), participant('B')], [meeting('one', '2026-10-07'), meeting('two', '2026-10-07')]);
  assert.equal(result.meetings[0].tesoriLetturaReaderId, 'A');
  assert.equal(result.meetings[1].tesoriLetturaReaderId, 'B');
});

test('slot matching fills scarce roles without unnecessary empty slots', () => {
  const a = { id: 'a' }, b = { id: 'b' }, c = { id: 'c' };
  const result = assignUniqueSlots([
    { id: 'one', candidates: [a, b] },
    { id: 'two', candidates: [a, c] },
    { id: 'three', candidates: [a, c] },
  ], p => p.id);
  assert.equal(result.size, 3);
  assert.equal(new Set([...result.values()].map(p => p.id)).size, 3);
});

function invokeGenerator(path, name, context) {
  const source = readFileSync(new URL('../' + path, import.meta.url), 'utf8');
  const start = source.indexOf('  const ' + name + ' =');
  const end = source.indexOf('\n  };', start) + '\n  };'.length;
  assert.ok(start >= 0 && end > start, name + ' function found');
  const compiled = ts.transpileModule(source.slice(start, end), {
    compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ESNext },
  }).outputText;
  const fn = new Function(...Object.keys(context), compiled + '; return ' + name + ';')(...Object.values(context));
  fn(...(context.args || []));
}
const noop = () => {};
const iso = date => date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
const addDays = (date, count) => { const result = new Date(date); result.setDate(result.getDate() + count); return result; };

test('monthly generator leaves gaps instead of assigning duplicate roles or a person busy on Wednesday', () => {
  const people = ['A', 'B', 'C', 'D'].map(id => ({ ...person(id), roles: { uscieri: true, microfoni: true, console: true } }));
  const state = base({ people, vitaEMinistero: { participants: [participant('vm', {}, 'A')], meetings: [meeting('v', '2026-10-07', { tesoriLetturaReaderId: 'vm' })] } });
  let rows;
  invokeGenerator('src/App.tsx', 'generateMensile', {
    args: [2026, 9], state, menRows: null, allMensileRows: [], allDomenicaRows: null, domenicaProgramRows: null, checkAdminPermission: () => true,
    mondayDates: () => [new Date('2026-10-05T12:00:00')], iso, addDays, weeklyDates, createSchedulingLedger,
    fairPick: (pool, counts, last, index, exclude) => pool.find(p => !exclude.includes(p.id)),
    markUsed: noop, consolePairValid: () => true, MESI: Array(12).fill('Mese'), fmtShort: iso,
    setMenTitle: noop, setMenRows: noop, setMenWarn: noop, setActiveArchiveId: noop, syncActivePrograms: value => { rows = value.menRows; },
  });
  const assigned = [...rows[0].ingresso, rows[0].auditorium, ...rows[0].microfoni, ...rows[0].audioVideo].filter(n => n !== '—');
  assert.equal(assigned.includes('A'), false);
  assert.equal(new Set(assigned).size, assigned.length);
  assert.equal(assigned.length, 3);
  assert.equal(rows[0].date.getDay(), 1);
  assert.equal(rows[0].date.getDate(), 5);
});

test('Sunday generation respects existing monthly roles and removes the fixed sample reader override', () => {
  const state = base({ people: ['A', 'B'].map(id => ({ ...person(id), roles: { presidentePubblica: true, lettore: true } })) });
  let rows;
  invokeGenerator('src/App.tsx', 'generateDomenica', {
    args: [2026, 9], state, menRows: [], allMensileRows: [{ date: new Date('2026-10-05T12:00:00'), auditorium: 'A' }], allDomenicaRows: [], domenicaProgramRows: null,
    checkAdminPermission: () => true, sundayDates: () => [new Date('2026-10-11T12:00:00')],
    iso, addDays, createSchedulingLedger, OCTOBRE_2026_WEEKEND: { '2026-10-11': { lettore: 'A' } },
    fairPick: (pool, counts, last, index, exclude) => pool.find(p => !exclude.includes(p.id)),
    markUsed: noop, MESI: Array(12).fill('Mese'), fmtDate: iso,
    setDomTitle: noop, setDomRows: noop, setDomWarn: noop, syncActivePrograms: value => { rows = value.domRows; },
  });
  assert.equal(rows[0].presidente, 'B');
  assert.equal(rows[0].lettore, '—');
});

test('field service generator clears a conflicting old conductor and cannot reuse the first choice in a second meeting', () => {
  const conductors = ['A', 'B'].map(id => ({ id: 'f-' + id, name: id, personId: id, availability: {} }));
  const currentMonthSchedule = [
    { dateStr: '2026-10-07', slotKey: 'speciale', time: '09:00', conductorId: 'f-A', isActive: true },
    { dateStr: '2026-10-07', slotKey: 'speciale', time: '15:00', conductorId: 'f-A', isActive: true },
  ];
  const data = { conductors, schedule: currentMonthSchedule, locations: ['Sala'] };
  const state = base({ people: ['A', 'B'].map(id => person(id)), servizioCampo: data,
    vitaEMinistero: { participants: [participant('vm', {}, 'A')], meetings: [meeting('v', '2026-10-07', { tesoriLetturaReaderId: 'vm' })] } });
  let saved;
  invokeGenerator('src/components/ServizioCampoView.tsx', 'handleAutoGenerate', {
    state, activePrograms: {}, data, conductors, schedule: data.schedule, currentMonthSchedule,
    selectedYear: 2026, selectedMonth: 9, checkAdminPermission: () => true, createSchedulingLedger,
    onShowToast: noop, setActiveSubTab: noop, setGenerationWarnings: noop, defaultSettings: {},
    DEFAULT_SLOT_SETTINGS: { speciale: { time: '09:00', defaultLocation: 'Sala' } }, rotationLocations: ['Sala'],
    saveServizioData: value => { saved = value; },
  });
  assert.equal(saved.schedule[0].conductorId, 'f-B');
  assert.equal(saved.schedule[1].conductorId, null);
});

test('public-work generation cannot reuse someone across shifts and excludes field-service conductors', () => {
  const participants = ['A', 'B', 'C', 'D'].map(id => ({ id: 'o-' + id, name: id, personId: id, gender: 'M',
    availability: { ribollaMartedi: { turno1: true, turno2: true } } }));
  const operaData = { participants, schedule: [] };
  const state = base({ people: ['A', 'B', 'C', 'D'].map(id => person(id)), operaPubblica: operaData,
    servizioCampo: { conductors: [{ id: 'f', personId: 'A', name: 'A' }], schedule: [{ dateStr: '2026-10-06', conductorId: 'f', isActive: true }] } });
  let saved;
  invokeGenerator('src/components/OperaPubblicaView.tsx', 'handleAutoGenerateSchedule', {
    state, activePrograms: {}, operaData, participants, schedule: [], selectedYear: 2026, selectedMonth: 9,
    daysInMonthList: [{ dateStr: '2026-10-06', dayOfWeek: 'martedi', location: 'Mercato di Ribolla' }],
    checkAdminPermission: () => true, createSchedulingLedger, onShowToast: noop, setGenerationWarnings: noop,
    getParticipantGender: p => p.gender, updateOperaData: value => { saved = value; },
  });
  const ids = [...saved.schedule[0].turno1, ...saved.schedule[0].turno2];
  assert.equal(ids.includes('o-A'), false);
  assert.equal(new Set(ids).size, ids.length);
  assert.equal(ids.length, 3);
});

test('ordinary Bible study and every VM role are assigned uniquely with eligible participants', () => {
  const roleNames = ['presidente', 'preghiera', 'preghiera', 'tesoriDiscorso', 'tesoriGemme', 'tesoriLettura',
    'ministeroStudente', 'ministeroAssistente', 'vitaCristianaParti', 'studioBiblicoConduttore', 'studioBiblicoLettore'];
  const participants = roleNames.map((role, index) => participant('P' + index, { [role]: true }));
  const original = meeting('m', '2026-10-07', {
    studioBiblicoType: undefined, tesori1Title: 'Tesori',
    ministeroParts: [{ id: 'mp', number: 4, title: 'Parte', minutes: 3, hasAssistant: true, studentId: '' }],
    vitaCristianaParts: [{ id: 'vp', number: 5, title: 'Vita cristiana', minutes: 15, speakerId: '' }],
  });
  const result = generate(participants, [original]);
  const ids = vitaAssignmentIds(result.meetings[0]);
  assert.equal(ids.length, 11);
  assert.equal(new Set(ids).size, 11);
  assert.deepEqual(result.warnings, []);
  assert.equal(result.meetings[0].studioBiblicoConductorId, 'P9');
  assert.equal(result.meetings[0].studioBiblicoReaderId, 'P10');
});

test('new monthly programmes start on each Monday in the selected month', () => {
  assert.deepEqual(mondayDates(2026, 9).map(date => date.getDate()), [5, 12, 19, 26]);
  assert.deepEqual(mondayDates(2026, 10).map(date => date.getDate()), [2, 9, 16, 23, 30]);
  assert.ok(mondayDates(2026, 9).every(date => date.getDay() === 1));
});

test('legacy saved Sunday headings show the following Monday without changing the stored date', () => {
  const original = new Date('2026-10-04T12:00:00');
  const monday = mensileWeekStart(original);
  assert.equal(original.getDay(), 0);
  assert.equal(original.getDate(), 4);
  assert.equal(monday.getDay(), 1);
  assert.equal(monday.getDate(), 5);
  assert.equal(mensileWeekStart(monday).getTime(), monday.getTime());
});

test('Monday weekly duties reserve Wednesday and the Sunday of the same week', () => {
  assert.deepEqual(weeklyDates('2026-10-05'), ['2026-10-07', '2026-10-11']);
  assert.equal(mensileWeekStart('2026-05-31').getMonth(), 5);
  assert.equal(mensileWeekStart('2026-05-31').getDate(), 1);
  const ledger = createSchedulingLedger(base(), { mensileRows: [{ date: new Date('2026-10-05T12:00:00'), auditorium: 'Prova' }] });
  assert.ok(ledger.isBusy({ name: 'Prova' }, '2026-10-07'));
  assert.ok(ledger.isBusy({ name: 'Prova' }, '2026-10-11'));
  assert.equal(ledger.isBusy({ name: 'Prova' }, '2026-10-04'), false);
  assert.equal(ledger.isBusy({ name: 'Prova' }, '2026-10-14'), false);
});

test('active Monday programme overrides a legacy Sunday archive for the same week', () => {
  const ledger = createSchedulingLedger(base({ mensileArchives: [{ savedAt: '2026-10-01', rows: [
    { dateStr: '2026-10-04', auditorium: 'Archivio vecchio' },
  ] }] }), { mensileRows: [{ date: new Date('2026-10-05T12:00:00'), auditorium: 'Attivo nuovo' }] });
  assert.ok(ledger.isBusy({ name: 'Attivo nuovo' }, '2026-10-07'));
  assert.ok(ledger.isBusy({ name: 'Attivo nuovo' }, '2026-10-11'));
  assert.equal(ledger.isBusy({ name: 'Archivio vecchio' }, '2026-10-07'), false);
});

test('weekly multi-role check spans every programme, one week = Monday to Sunday', () => {
  assert.equal(weekKey('2026-10-11'), '2026-10-05'); // Sunday belongs to the week before
  assert.equal(weekKey('2026-10-12'), '2026-10-12');
  const state = base({
    people: [person('p', 'Luca')],
    vitaEMinistero: { participants: [participant('vm', {}, 'p')], meetings: [meeting('v', '2026-10-07', { tesoriLetturaReaderId: 'vm' })] },
    servizioCampo: { conductors: [{ id: 'f', personId: 'p', name: 'Luca' }], schedule: [
      { dateStr: '2026-10-13', conductorId: 'f', isActive: true }, { dateStr: '2026-10-15', conductorId: 'f', isActive: true }] },
    discorsiEsterni: { oratori: [], uscite: [{ id: 'u', data: '2026-10-25', oratore: 'Marco', titolo: 't', congregazione: 'c' }] },
  });
  const conflicts = weeklyRoleConflicts(state, {
    mensileRows: [{ date: new Date('2026-10-05T12:00:00'), ingresso: ['luca', '—'] }],
    domenicaRows: [{ date: new Date('2026-10-25T12:00:00'), presidente: 'Marco' }],
  });
  assert.deepEqual(conflicts.map(c => [c.week, c.name, c.roles.map(r => r.program + ':' + r.role)]), [
    ['2026-10-05', 'Luca', ['mensile:Usciere ingresso', 'vitaEMinistero:Lettura biblica']],
    ['2026-10-19', 'Marco', ['domenica:Presidente', 'discorsiEsterni:Discorso esterno']],
  ]); // same role twice in a week (two field-service meetings) is not a conflict
});

test('rotation is per part: who has gone longest without that part wins, even if busy elsewhere recently', () => {
  // A read on 09-01 but prayed on 09-30; B read on 09-20. A has gone longest without reading.
  const history = [
    meeting('h1', '2026-09-01', { tesoriLetturaReaderId: 'A' }),
    meeting('h2', '2026-09-20', { tesoriLetturaReaderId: 'B' }),
    meeting('h3', '2026-09-30', { preghieraInizialeId: 'A' }),
  ];
  const result = generate([participant('A', { tesoriLettura: true }), participant('B', { tesoriLettura: true })], [...history, meeting('t', '2026-10-07')]);
  assert.equal(result.meetings.find(m => m.id === 't').tesoriLetturaReaderId, 'A');
});

test('Bisogni locali is left empty and does not produce a warning', () => {
  const result = generate([participant('A', { vitaCristianaParti: true })], [meeting('m', '2026-10-07', {
    vitaCristianaParts: [
      { id: 'v1', number: 7, title: 'Bisogni locali', minutes: 15, speakerId: 'old' },
      { id: 'v2', number: 8, title: 'Altra parte', minutes: 15, speakerId: '' },
    ],
  })]);
  const parts = result.meetings[0].vitaCristianaParts;
  assert.equal(parts[0].speakerId, '');
  assert.equal(parts[1].speakerId, 'A');
  assert.equal(result.warnings.some(w => w.includes('Parte 7')), false);
});

test('presidente not repeated in the same month when someone else is free; repeats are flagged', () => {
  const participants = [participant('a', { presidente: true }), participant('b', { presidente: true, preghiera: true }), participant('c', { preghiera: true }), participant('d', { preghiera: true })];
  const { meetings } = generate(participants, [meeting('w1', '2026-10-05'), meeting('w2', '2026-10-12')]);
  assert.notEqual(meetings[0].presidenteId, meetings[1].presidenteId);
  assert.deepEqual(vitaMonthlyRepeats(meetings, participants), []);
  const repeated = [meeting('x', '2026-10-05', { presidenteId: 'a', studioBiblicoType: 'studio', studioBiblicoConductorId: 'b' }),
    meeting('y', '2026-10-12', { presidenteId: 'a', studioBiblicoType: 'studio', studioBiblicoConductorId: 'b' }),
    meeting('z', '2026-11-02', { presidenteId: 'a' })];
  assert.equal(vitaMonthlyRepeats(repeated, participants).length, 2);
});

test('presidente avoids the usciere of the same midweek evening', () => {
  const participants = [participant('Anna', { presidente: true }), participant('Bruno', { presidente: true })];
  const mensileRows = [{ date: '2026-10-05', ingresso: ['Anna', '—'], auditorium: '—', microfoni: [], audioVideo: [] }];
  const { meetings, warnings } = generate(participants, [meeting('w1', '2026-10-05')], {}, { mensileRows });
  assert.equal(meetings[0].presidenteId, 'Bruno');
  assert.ok(!warnings.some(w => w.includes('usciere')));
  const solo = generate([participants[0]], [meeting('w1', '2026-10-05')], {}, { mensileRows });
  assert.equal(solo.meetings[0].presidenteId, 'Anna');
  assert.ok(solo.warnings.some(w => w.includes('usciere')));
});
