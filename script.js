/* ==========================================================================
   Faculty PD Achievement Dashboard  |  Ajman University TLC
   Vanilla JS. Libraries: Chart.js, SheetJS (XLSX), jsPDF.
   Everything runs in the browser; no data leaves the page.

   SECTIONS
   1. Configuration (themes, badge-to-theme map, tier rules)  <- edit here
   2. Utilities
   3. File reading and row normalisation
   4. Model building and tier engine
   5. Rendering helpers
   6. Individual view
   7. College view
   8. Charts
   9. Exports (PDF, Excel, print)
   10. App wiring
   ========================================================================== */
(function () {
'use strict';

/* ==========================================================================
   1. CONFIGURATION
   ========================================================================== */

const THEMES = [
  { id: 'LD', name: 'Learning Design', short: 'Learning Design' },
  { id: 'AF', name: 'Assessment & Feedback for Learning', short: 'Assessment & Feedback' },
  { id: 'IT', name: 'Inclusive Teaching & Learner Support', short: 'Inclusive Teaching' },
  { id: 'DP', name: 'Digital Pedagogy & Information Literacy in Teaching', short: 'Digital Pedagogy' },
  { id: 'RP', name: 'Reflective Practice, Innovation & Professional Growth', short: 'Reflective Practice' }
];

/* DRAFT badge-to-theme map, inferred from badge names. The badge report does
   not carry themes. TLC should confirm or replace it (Settings > Badge-to-theme
   map, or edit below). Keys are canonical badge names (see canon()). */
const DEFAULT_THEME_MAP = {
  'self-paced course - learning design': 'LD',
  'learning battle cards': 'LD',
  'embedding education for sustainable development in the curriculum': 'LD',
  'innovative teaching: engaging adult learners with active learning': 'LD',
  'creating engaging learning content using pro.play': 'LD',
  'using the pro.play for exploring the role of ai in improving the student experience': 'LD',

  'self-paced courses - assessment & feedback for learning': 'AF',
  'pedagogy and assessment redesign': 'AF',
  'create and grade quizzes in moodle': 'AF',
  'beyond mcqs: design moodle quizzes using multimedia': 'AF',
  'emc 3 - create assessments and quizzes': 'AF',
  'the 5-phase journey of moodle workshops': 'AF',

  'inclusive teaching and learner support': 'IT',
  'the heart of teaching': 'IT',

  'digital pedagogy & information literacy in teaching': 'DP',
  'dec certificate in ai for higher education': 'DP',
  'emc 1 - introduction to moodle': 'DP',
  'emc 2 - create and manage content': 'DP',
  'emc 4 - manage students grades and progress': 'DP',
  'emc 5 - create interactive activities': 'DP',
  'teaching better with moodle book': 'DP',
  'teaching effectively on zoom: practical techniques for engaging online classes': 'DP',
  'imaging-based h5p interactive content in moodle': 'DP',
  'nearpod essentials: build your first interactive lesson': 'DP',
  'using nearpod to strengthen engagement in live online teaching': 'DP',

  'reflective practice, innovation and professional growth': 'RP',
  'drg briefing on a few research-related matters': 'RP',
  'case study teaching and writing': 'RP'
};

/* Tier rules (AU Faculty PD Framework). */
const RULES = {
  bronzeThemes: 3,            // engage with at least 3 themes ...
  bronzeSessionsPerTheme: 2,  // ... with at least 2 approved sessions in each
  disciplinePointsMax: 5      // not in the badge report: shown as a placeholder
};
/* Manual Silver awards present in the Moodle report. */
const SILVER_AWARD_KEYS = ['silver tier engagement', 'silver tier recognition'];
/* No Gold badge exists in the current report. Add canonical names here when it does. */
const GOLD_AWARD_KEYS = [];
/* Essentials of Moodle Certificate = EMC 1-5. A structured programme gives Silver automatically. */
const EMC_KEYS = [
  'emc 1 - introduction to moodle', 'emc 2 - create and manage content', 'emc 3 - create assessments and quizzes',
  'emc 4 - manage students grades and progress', 'emc 5 - create interactive activities'
];
const SPECIAL_BADGES = [
  { key: 'dec certificate in ai for higher education', name: 'DEC Certificate in AI for Higher Education' }
];
/* Test accounts and test badges removed on load. */
const EXCLUDE_EMAIL_DOMAINS = ['eummena.org'];
const EXCLUDE_BADGE_KEYS = ['test'];

const COLLEGE_RULES = [
  [/dentistry/i, 'College of Dentistry'],
  [/humanities/i, 'College of Humanities and Sciences'],
  [/medicine|medical/i, 'College of Medicine'],
  [/business/i, 'College of Business Administration'],
  [/engineering|eng and it|information technology/i, 'College of Engineering and Information Technology'],
  [/pharmacy/i, 'College of Pharmacy and Health Sciences'],
  [/\blaw\b/i, 'College of Law'],
  [/mass comm/i, 'College of Mass Communication'],
  [/architecture/i, 'College of Architecture, Art and Design'],
  [/research and graduate/i, 'Deanship of Research and Graduate Studies']
];
const TLC_NAME = 'Teaching & Learning Center';
const UNMAPPED_COLLEGE = 'Other (unmapped)';

const HEADER_ALIASES = {
  badge: ['name', 'badge name', 'badge'],
  email: ['email address', 'email', 'e-mail'],
  fullname: ['full name', 'fullname', 'user full name', 'faculty name'],
  date: ['date issued', 'issued', 'date'],
  criteria: ['criteria'],
  delivery: ['type of delivery', 'delivery'],
  college: ['college'],
  status: ['badge status', 'status']
};

/* ==========================================================================
   2. UTILITIES
   ========================================================================== */
const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pct = x => Math.round(x * 100);
const MONTH_NAMES = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
const fmtDate = d => d ? d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : 'Date not recorded';
const monthKey = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
const monthLabel = k => { const [y, m] = k.split('-'); return new Date(+y, +m - 1, 1).toLocaleDateString('en-GB', { month: 'short', year: 'numeric' }); };
const plural = (n, a, b) => n + ' ' + (n === 1 ? a : b);
const latinOnly = s => String(s || '').replace(/[^\x20-\x7E]/g, '').replace(/\s+/g, ' ').trim();
const ucfirst = s => s.charAt(0).toUpperCase() + s.slice(1);
function safeStore(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; } return null; }
function mulberry32(a) { return function () { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }

function tidyName(n) {
  return String(n || '').replace(/[–—]/g, '-').replace(/®/g, '').replace(/\s+/g, ' ')
    .replace(/\s*-\s*Instructor Approved$/i, '').trim();
}
function canon(n) {
  return tidyName(n).toLowerCase().replace('designing engaging moodle quizzes using multimedia', 'design moodle quizzes using multimedia');
}
function cleanCollege(raw) {
  const s = String(raw || '').trim();
  if (!s) return 'Unassigned';
  if (/^teaching\s*(and|&)?\s*learning/i.test(s)) return TLC_NAME;
  for (const [re, name] of COLLEGE_RULES) if (re.test(s)) return name;
  return UNMAPPED_COLLEGE;
}

/* ==========================================================================
   3. FILE READING AND ROW NORMALISATION
   ========================================================================== */
function parseHtmlTable(text) {
  const doc = new DOMParser().parseFromString(text, 'text/html');
  const table = doc.querySelector('table');
  if (!table) throw new Error('No table found in the HTML file.');
  const trs = Array.from(table.querySelectorAll('tr'));
  const head = Array.from(trs[0].children).map(c => c.textContent.trim());
  return trs.slice(1).map(tr => {
    const o = {}; const cells = Array.from(tr.children);
    head.forEach((h, i) => { o[h] = cells[i] ? cells[i].textContent.replace(/\s+/g, ' ').trim() : ''; });
    return o;
  });
}
async function readRowsFromFile(file) {
  const buf = await file.arrayBuffer();
  const head = new TextDecoder().decode(new Uint8Array(buf.slice(0, 400))).trim().toLowerCase();
  if (head.startsWith('<')) return parseHtmlTable(new TextDecoder().decode(buf));
  if (typeof XLSX === 'undefined') throw new Error('The Excel library did not load. Check your connection and reload.');
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
}
const cell = v => (v instanceof Date ? v : (v == null ? '' : String(v).trim()));

function normalizeRows(raw) {
  if (!raw.length) throw new Error('The file has no data rows.');
  const lk = {}; Object.keys(raw[0]).forEach(k => { lk[String(k).trim().toLowerCase()] = k; });
  const map = {};
  Object.keys(HEADER_ALIASES).forEach(f => { map[f] = HEADER_ALIASES[f].map(a => lk[a]).find(Boolean) || null; });
  if (!map.badge || !map.email) {
    throw new Error('Could not find the "Name" and "Email address" columns. Found: ' + Object.keys(raw[0]).join(', '));
  }
  return raw.map(r => {
    const o = {};
    Object.keys(map).forEach(f => { o[f] = map[f] ? cell(r[map[f]]) : ''; });
    return o;
  });
}
function parseDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400000));
  const s = String(v || '').trim(); if (!s) return null;
  const m = s.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})(?:,?\s*(\d{1,2}):(\d{2})\s*(AM|PM)?)?/i);
  if (m) {
    const mi = MONTH_NAMES.findIndex(x => x.startsWith(m[2].slice(0, 3).toLowerCase()));
    if (mi >= 0) {
      let h = +m[4] || 0;
      if (m[6]) h = (h % 12) + (m[6].toUpperCase() === 'PM' ? 12 : 0);
      return new Date(+m[3], mi, +m[1], h, +(m[5] || 0));
    }
  }
  const d = new Date(s); return isNaN(d) ? null : d;
}

/* ==========================================================================
   4. MODEL BUILDING AND TIER ENGINE
   ========================================================================== */
const S = {
  rows: [], people: new Map(), catalog: new Map(), catalogList: [], notes: {},
  roster: new Map(), themeMap: Object.assign({}, DEFAULT_THEME_MAP),
  cycleStart: 9, activeDays: 60, year: 'all', years: [],
  view: 'individual', selected: null, fileName: '', loaded: false,
  filters: { college: 'all', status: 'all' },
  analyses: [], charts: {}
};

function academicYear(d) {
  if (!d) return null;
  const y = d.getFullYear(), m = d.getMonth() + 1;
  if (S.cycleStart === 1) return String(y);
  const start = m >= S.cycleStart ? y : y - 1;
  return start + '–' + (start + 1);
}

function buildModel(rows) {
  const notes = { rowsRead: rows.length, definitions: 0, excluded: 0, duplicates: 0, unmappedColleges: new Map(), unmappedBadges: new Map(), undated: 0 };
  const people = new Map();
  const catalog = new Map();
  const nameVotes = new Map();

  rows.forEach(r => {
    const rawName = String(r.badge || '').trim();
    if (!rawName) return;
    const key = canon(rawName);
    if (EXCLUDE_BADGE_KEYS.includes(key)) { if (r.email) notes.excluded++; return; }
    const status = String(r.status || '');
    // catalogue entry (definitions and issuances)
    let c = catalog.get(key);
    if (!c) { c = { key, name: tidyName(rawName), count: 0, availableRow: false, theme: null }; catalog.set(key, c); }
    if (/^available/i.test(status) && !/not available/i.test(status)) c.availableRow = true;
    const vk = key + '||' + tidyName(rawName);
    nameVotes.set(vk, (nameVotes.get(vk) || 0) + 1);

    const email = String(r.email || '').trim().toLowerCase();
    if (!email) { notes.definitions++; return; }
    if (EXCLUDE_EMAIL_DOMAINS.some(d => email.endsWith('@' + d))) { notes.excluded++; return; }

    let p = people.get(email);
    if (!p) {
      p = { email, name: String(r.fullname || '').trim() || email, collegeRaw: String(r.college || '').trim(), badges: new Map() };
      p.college = cleanCollege(p.collegeRaw);
      if (p.college === UNMAPPED_COLLEGE) notes.unmappedColleges.set(p.collegeRaw, (notes.unmappedColleges.get(p.collegeRaw) || 0) + 1);
      people.set(email, p);
    }
    const date = parseDate(r.date);
    if (!date) notes.undated++;
    const prev = p.badges.get(key);
    if (prev) {
      notes.duplicates++;
      if (date && (!prev.date || date < prev.date)) { prev.date = date; }
      return;
    }
    c.count++;
    p.badges.set(key, { key, name: c.name, date, delivery: String(r.delivery || '').trim() });
  });

  // best display name per badge
  const best = new Map();
  nameVotes.forEach((n, vk) => {
    const [k, nm] = vk.split('||'); const cur = best.get(k);
    if (!cur || n > cur.n) best.set(k, { n, nm });
  });
  catalog.forEach((c, k) => { if (best.has(k)) c.name = best.get(k).nm; });
  people.forEach(p => p.badges.forEach(b => { const c = catalog.get(b.key); if (c) b.name = c.name; }));

  S.people = people; S.catalog = catalog; S.notes = notes;
  applyThemeMap(); applyRoster(); computeYears();
  S.loaded = true;
}

function applyThemeMap() {
  S.catalog.forEach(c => { c.theme = S.themeMap[c.key] || null; });
  const un = new Map();
  S.catalog.forEach(c => { if (!c.theme && c.count > 0 && !SILVER_AWARD_KEYS.includes(c.key) && !GOLD_AWARD_KEYS.includes(c.key)) un.set(c.name, c.count); });
  S.notes.unmappedBadges = un;
  S.catalogList = Array.from(S.catalog.values())
    .filter(c => c.theme && c.count > 0 && c.availableRow)
    .sort((a, b) => b.count - a.count);
}
function applyRoster() {
  S.people.forEach(p => {
    const r = S.roster.get(p.email);
    p.dept = r && r.dept ? r.dept : null;
    p.rank = r && r.rank ? r.rank : null;
    p.hire = r && r.hire ? r.hire : null;
  });
}
function computeYears() {
  const set = new Set();
  S.people.forEach(p => p.badges.forEach(b => { const y = academicYear(b.date); if (y) set.add(y); }));
  S.years = Array.from(set).sort();
  if (S.year !== 'all' && !S.years.includes(S.year)) S.year = 'all';
}

/* ---- tier evaluation for one person over a set of badges ---- */
function evaluate(p, badgeList) {
  const badges = badgeList.slice().sort((a, b) => (a.date ? a.date.getTime() : 0) - (b.date ? b.date.getTime() : 0));
  const keys = new Set(badges.map(b => b.key));
  const themes = THEMES.map(t => ({ id: t.id, name: t.name, short: t.short, badges: [], count: 0, done: false }));
  const themeById = {}; themes.forEach(t => { themeById[t.id] = t; });
  const unmapped = [];
  badges.forEach(b => {
    const tid = S.themeMap[b.key];
    if (tid && themeById[tid]) themeById[tid].badges.push(b); else if (!SILVER_AWARD_KEYS.includes(b.key) && !GOLD_AWARD_KEYS.includes(b.key)) unmapped.push(b);
  });
  themes.forEach(t => { t.count = t.badges.length; t.done = t.count >= RULES.bronzeSessionsPerTheme; });

  const capped = themes.map(t => Math.min(t.count, RULES.bronzeSessionsPerTheme)).sort((a, b) => b - a);
  const need = RULES.bronzeThemes * RULES.bronzeSessionsPerTheme;
  const pts = capped.slice(0, RULES.bronzeThemes).reduce((a, b) => a + b, 0);
  const themesDone = themes.filter(t => t.done).length;
  const bronzeEarned = themesDone >= RULES.bronzeThemes;
  // date bronze was reached: replay badges in order
  let bronzeDate = null;
  if (bronzeEarned) {
    const cnt = {};
    for (const b of badges) {
      const tid = S.themeMap[b.key]; if (!tid || !themeById[tid]) continue;
      cnt[tid] = (cnt[tid] || 0) + 1;
      if (Object.values(cnt).filter(n => n >= RULES.bronzeSessionsPerTheme).length >= RULES.bronzeThemes) { bronzeDate = b.date; break; }
    }
  }
  const bronze = { pts, need, progress: Math.min(1, pts / need), earned: bronzeEarned, themesDone, date: bronzeDate };

  // Silver
  const awardBadges = badges.filter(b => SILVER_AWARD_KEYS.includes(b.key));
  const emcHave = EMC_KEYS.filter(k => keys.has(k));
  const emcDone = emcHave.length === EMC_KEYS.length;
  const emcDate = emcDone ? badges.filter(b => EMC_KEYS.includes(b.key)).map(b => b.date).filter(Boolean).sort((a, b) => b - a)[0] || null : null;
  const silverVia = awardBadges.length ? 'award' : (emcDone ? 'programme' : null);
  const silverDates = [awardBadges.map(b => b.date).filter(Boolean).sort((a, b) => a - b)[0], emcDate].filter(Boolean).sort((a, b) => a - b);
  const silverEarned = !!silverVia;
  const silverProg = silverEarned ? 1 : Math.max(bronze.progress * 0.5, emcHave.length / EMC_KEYS.length);
  const silver = { earned: silverEarned, via: silverVia, progress: silverProg, emcHave: emcHave.length, emcMissing: EMC_KEYS.filter(k => !keys.has(k)), date: silverDates[0] || null };

  // Gold
  const goldBadges = badges.filter(b => GOLD_AWARD_KEYS.includes(b.key));
  const goldEarned = goldBadges.length > 0;
  const gold = { earned: goldEarned, progress: goldEarned ? 1 : (silverEarned ? 0.5 : silverProg * 0.5), date: goldBadges.length ? goldBadges[0].date : null };

  const status = (t, lockedWhen) => t.earned ? 'earned' : (lockedWhen || t.progress <= 0 ? 'locked' : 'progress');
  bronze.status = status(bronze);
  silver.status = status(silver);
  gold.status = status(gold, !silverEarned);

  const tier = goldEarned ? 'gold' : silverEarned ? 'silver' : bronzeEarned ? 'bronze' : 'none';
  const dated = badges.filter(b => b.date);
  return {
    p, badges, themes, unmapped, bronze, silver, gold, tier,
    tierRank: { none: 0, bronze: 1, silver: 2, gold: 3 }[tier],
    progress: goldEarned ? 1 : silverEarned ? 0.75 : bronze.progress * 0.5,
    totalCourses: badges.filter(b => !SILVER_AWARD_KEYS.includes(b.key) && !GOLD_AWARD_KEYS.includes(b.key)).length,
    lastDate: dated.length ? dated[dated.length - 1].date : null,
    firstDate: dated.length ? dated[0].date : null
  };
}

function runAnalyses() {
  const out = [];
  S.people.forEach(p => {
    let list = Array.from(p.badges.values());
    if (S.year !== 'all') list = list.filter(b => academicYear(b.date) === S.year);
    if (!list.length) return;
    out.push(evaluate(p, list));
  });
  S.analyses = out;
  S.maxDate = out.reduce((m, a) => (a.lastDate && (!m || a.lastDate > m) ? a.lastDate : m), null);
  out.forEach(a => {
    a.active = !!(a.lastDate && S.maxDate && (S.maxDate - a.lastDate) / 86400000 <= S.activeDays);
  });
}

/* ---- aggregations ---- */
const countTiers = list => list.reduce((c, a) => { c[a.tier]++; return c; }, { none: 0, bronze: 0, silver: 0, gold: 0 });
function groupBy(list, fn) {
  const m = new Map();
  list.forEach(a => { const k = fn(a); if (!m.has(k)) m.set(k, []); m.get(k).push(a); });
  return Array.from(m, ([name, items]) => {
    const t = countTiers(items); const n = items.length;
    return {
      name, items, n, tiers: t, bronzePlus: n - t.none, silverPlus: t.silver + t.gold, gold: t.gold,
      rate: n ? (n - t.none) / n : 0, avg: n ? items.reduce((s, a) => s + a.progress, 0) / n : 0,
      active: items.filter(a => a.active).length
    };
  });
}
function monthRange(list) {
  const ds = []; list.forEach(a => a.badges.forEach(b => { if (b.date) ds.push(b.date); }));
  if (!ds.length) return [];
  let lo = new Date(Math.min.apply(null, ds)), hi = new Date(Math.max.apply(null, ds));
  const out = []; let y = lo.getFullYear(), m = lo.getMonth();
  while (y < hi.getFullYear() || (y === hi.getFullYear() && m <= hi.getMonth())) {
    out.push(y + '-' + String(m + 1).padStart(2, '0')); m++; if (m > 11) { m = 0; y++; }
  }
  return out;
}
function monthlyStats(list) {
  const months = monthRange(list); const idx = {}; months.forEach((k, i) => { idx[k] = i; });
  const issued = months.map(() => 0), faculty = months.map(() => new Set());
  list.forEach(a => a.badges.forEach(b => { if (b.date) { const i = idx[monthKey(b.date)]; issued[i]++; faculty[i].add(a.p.email); } }));
  return { months, issued, faculty: faculty.map(s => s.size) };
}
const tierName = t => ({ none: 'No tier yet', bronze: 'Bronze', silver: 'Silver', gold: 'Gold' }[t]);
const hasRoster = () => S.roster.size > 0;

/* ---- individual guidance ---- */
function suggestions(a) {
  const owned = new Set(a.badges.map(b => b.key));
  const pool = S.catalogList.filter(c => !owned.has(c.key));
  const chosen = a.themes.slice().sort((x, y) => Math.min(y.count, 2) - Math.min(x.count, 2) || y.count - x.count).slice(0, RULES.bronzeThemes);
  const needs = chosen.filter(t => t.count < RULES.bronzeSessionsPerTheme).map(t => ({ theme: t, need: RULES.bronzeSessionsPerTheme - t.count }));
  needs.sort((x, y) => x.need - y.need);
  const picks = [];
  needs.forEach(n => {
    const opts = pool.filter(c => c.theme === n.theme.id).slice(0, n.need);
    opts.forEach(o => picks.push({ badge: o, theme: n.theme, gain: 1 / (RULES.bronzeThemes * RULES.bronzeSessionsPerTheme) }));
    n.available = opts.length;
  });
  return { picks, needs, remaining: a.bronze.need - a.bronze.pts };
}
function bronzeMissing(a) {
  const s = suggestions(a); const items = [];
  const startedOrChosen = a.themes.slice().sort((x, y) => Math.min(y.count, 2) - Math.min(x.count, 2) || y.count - x.count).slice(0, RULES.bronzeThemes);
  startedOrChosen.forEach(t => {
    if (t.count < RULES.bronzeSessionsPerTheme) items.push({ text: `${t.short}: ${t.count} of ${RULES.bronzeSessionsPerTheme} approved sessions`, theme: t.id });
  });
  items.push({ text: `Discipline-specific T&L activity, up to ${RULES.disciplinePointsMax} points`, pending: true });
  return items;
}
function silverMissing(a) {
  const items = [];
  if (a.silver.earned) return items;
  if (!a.bronze.earned) items.push({ text: 'Complete Bronze (see above), or complete the Essentials of Moodle Certificate (EMC 1-5)' });
  items.push({ text: 'Application example and reflective statement, submitted to the TLC (awarded as the Silver Tier badge)', pending: true });
  if (a.silver.emcMissing.length && a.silver.emcMissing.length < EMC_KEYS.length) {
    items.push({ text: 'Or finish the Essentials of Moodle Certificate: ' + a.silver.emcMissing.map(k => (S.catalog.get(k) || { name: k }).name.split(' - ')[0]).join(', ') + ' remaining' });
  }
  return items;
}
function goldMissing(a) {
  if (a.gold.earned) return [];
  const items = [];
  if (!a.silver.earned) items.push({ text: 'Reach Silver first' });
  items.push({ text: 'Contribution beyond your own teaching (e.g. leading workshops, mentoring, scholarship)', pending: true });
  items.push({ text: 'Or Advance HE Fellowship D3-D4 (not recorded in the badge report)', pending: true });
  return items;
}

/* ==========================================================================
   5. RENDERING HELPERS
   ========================================================================== */
function themeColor(id) { return `var(--theme-${THEMES.findIndex(t => t.id === id) + 1})`; }
function tierColor(t) { return `var(--tier-${t})`; }

function ring(p, color, label) {
  const r = 40, c = 2 * Math.PI * r, v = Math.max(0, Math.min(1, p));
  return `<svg class="ring" viewBox="0 0 100 100" role="img" aria-label="${esc(label)}: ${pct(v)} percent">
    <circle class="ring-bg" cx="50" cy="50" r="${r}"></circle>
    <circle class="ring-fg" cx="50" cy="50" r="${r}" stroke="${color}" stroke-dasharray="${c.toFixed(2)}" stroke-dashoffset="${(c * (1 - v)).toFixed(2)}" transform="rotate(-90 50 50)"></circle>
    <text class="ring-text" x="50" y="57">${pct(v)}%</text></svg>`;
}
function bar(p, color, label) {
  return `<div class="bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${pct(p)}" aria-label="${esc(label)}"><span style="width:${pct(p)}%;background:${color}"></span></div>`;
}
function medal(color, locked) {
  const c = locked ? 'var(--tier-none)' : color;
  return `<svg viewBox="0 0 64 64" aria-hidden="true" focusable="false">
    <path d="M20 4h10l4 18H24z" fill="${c}" opacity=".55"></path><path d="M44 4H34l-4 18h10z" fill="${c}" opacity=".85"></path>
    <circle cx="32" cy="40" r="19" fill="${c}"></circle><circle cx="32" cy="40" r="13" fill="none" stroke="var(--surface)" stroke-width="2.5" opacity=".8"></circle>
    ${locked ? '<rect x="26" y="38" width="12" height="9" rx="2" fill="var(--surface)"></rect><path d="M28.5 38v-3a3.5 3.5 0 0 1 7 0v3" fill="none" stroke="var(--surface)" stroke-width="2.4"></path>'
      : '<path d="M24 40l5.5 5.5L41 34" fill="none" stroke="var(--surface)" stroke-width="3.4" stroke-linecap="round" stroke-linejoin="round"></path>'}</svg>`;
}
const chip = (status, text) => `<span class="chip ${status}">${esc(text || ({ earned: 'Earned', progress: 'In progress', locked: 'Locked' }[status]))}</span>`;
const kpi = (label, value, sub, cls) => `<div class="card kpi ${cls || ''}"><span class="label">${esc(label)}</span><span class="value">${esc(value)}</span>${sub ? `<span class="sub">${esc(sub)}</span>` : ''}</div>`;
function table(cols, rows, opts) {
  opts = opts || {};
  return `<div class="table-wrap"><table class="data">${opts.caption ? `<caption>${esc(opts.caption)}</caption>` : ''}<thead><tr>${cols.map(c => `<th scope="col" class="${c.num ? 'num' : ''}">${esc(c.h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${r.map((v, i) => `<td class="${cols[i].num ? 'num' : ''}"${cols[i].dir ? ' dir="auto"' : ''}>${v}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
function setStatus(msg, isError) {
  const el = $('#status'); el.textContent = msg || ''; el.classList.toggle('error', !!isError);
}
function initials(name) {
  const words = latinOnly(name).replace(/^(dr|prof|professor|mr|mrs|ms|eng)\.?\s+/i, '').split(' ').filter(Boolean);
  return ((words[0] || '?')[0] + (words.length > 1 ? words[words.length - 1][0] : '')).toUpperCase();
}
function passFilter(a, college, status) {
  if (college && college !== 'all' && a.p.college !== college) return false;
  if (status && status !== 'all' && a.tier !== status) return false;
  return true;
}

/* ==========================================================================
   6. INDIVIDUAL VIEW
   ========================================================================== */
function renderSearchResults() {
  const q = $('#q').value.trim().toLowerCase();
  const ul = $('#results'); ul.innerHTML = '';
  if (!q || !S.loaded) return;
  const hits = Array.from(S.people.values()).filter(p => p.name.toLowerCase().includes(q) || p.email.includes(q)).slice(0, 8);
  if (!hits.length) { ul.innerHTML = '<li class="placeholder">No faculty match that name or email.</li>'; return; }
  ul.innerHTML = hits.map(p => `<li><button type="button" data-email="${esc(p.email)}"><strong dir="auto">${esc(p.name)}</strong><span class="sub">${esc(p.email)}</span><span class="sub">${esc(p.college)}</span></button></li>`).join('');
}
function selectPerson(email) {
  S.selected = email; $('#results').innerHTML = ''; renderIndividual();
  const h = $('#ind-content h2'); if (h) { h.setAttribute('tabindex', '-1'); h.focus({ preventScroll: false }); }
}

function renderIndividual() {
  const root = $('#ind-content');
  if (!S.selected) {
    root.innerHTML = `<div class="card"><h2>Find a faculty member</h2><p class="muted">Type a name or email above. ${plural(S.people.size, 'faculty member', 'faculty members')} in the loaded report.</p></div>`;
    return;
  }
  const a = S.analyses.find(x => x.p.email === S.selected);
  const person = S.people.get(S.selected);
  if (!a) {
    root.innerHTML = `<div class="card"><h2 dir="auto">${esc(person ? person.name : S.selected)}</h2><p class="muted">No badges recorded in academic year ${esc(S.year)}. Choose "All years" to see their full record.</p></div>`;
    return;
  }
  const p = a.p; const sug = suggestions(a);
  const earnedTiers = ['bronze', 'silver', 'gold'].filter(t => a[t].earned).length;
  const badgesEarned = a.totalCourses + earnedTiers;

  const tierCard = (t, label, sub) => `
    <div class="card tier-card">${ring(a[t].progress, tierColor(t), label + ' progress')}
      <div><h3>${label} ${chip(a[t].status)}</h3>
        ${bar(a[t].progress, tierColor(t), label + ' progress')}
        <p class="small muted">${esc(sub)}</p></div></div>`;
  const bronzeSub = a.bronze.earned
    ? `${a.bronze.themesDone} themes with 2+ sessions. Discipline points: not in report.`
    : `${a.bronze.pts} of ${a.bronze.need} sessions across your best ${RULES.bronzeThemes} themes. Discipline points: not in report.`;
  const silverSub = a.silver.earned ? (a.silver.via === 'award' ? 'Silver Tier badge awarded.' : 'Essentials of Moodle Certificate complete (EMC 1-5).')
    : `Bronze route ${pct(a.bronze.progress * 0.5)}% (evidence pending); EMC route ${a.silver.emcHave} of 5.`;
  const goldSub = a.gold.earned ? 'Gold badge awarded.' : (a.silver.earned ? 'Silver done. Contribution evidence is not in the report.' : 'Locked until Silver is earned.');

  // gallery
  const special = SPECIAL_BADGES.map(sb => {
    const own = a.badges.find(b => b.key === sb.key);
    return { name: sb.name, status: own ? 'earned' : 'locked', date: own && own.date };
  });
  const emcStatus = a.silver.emcHave === 5 ? 'earned' : a.silver.emcHave > 0 ? 'progress' : 'locked';
  special.push({ name: 'Essentials of Moodle Certificate', status: emcStatus, extra: `${a.silver.emcHave} of 5 courses` });
  const tileTier = (t, label) => `<div class="badge-tile ${a[t].status === 'locked' ? 'locked' : ''}">${medal(tierColor(t), a[t].status === 'locked')}<span class="name">${label}</span>${chip(a[t].status)}<span class="meta">${a[t].earned ? fmtDate(a[t].date) : (a[t].status === 'progress' ? pct(a[t].progress) + '% complete' : 'Not started')}</span></div>`;
  const tileSpecial = s => `<div class="badge-tile ${s.status === 'locked' ? 'locked' : ''}">${medal('var(--accent)', s.status === 'locked')}<span class="name">${esc(s.name)}</span>${chip(s.status)}<span class="meta">${s.extra ? esc(s.extra) : (s.date ? fmtDate(s.date) : '')}</span></div>`;
  const tileTheme = t => {
    const st = t.done ? 'earned' : t.count > 0 ? 'progress' : 'locked';
    return `<div class="badge-tile ${st === 'locked' ? 'locked' : ''}">${medal(themeColor(t.id), st === 'locked')}<span class="name">${esc(t.short)}</span>${chip(st)}<span class="meta">${t.count} of ${RULES.bronzeSessionsPerTheme} sessions</span></div>`;
  };

  // completed
  const completed = a.themes.filter(t => t.badges.length).map(t => `
    <div class="group-title"><span class="dot" style="background:${themeColor(t.id)}"></span>${esc(t.name)} <span class="chip ${t.done ? 'earned' : 'progress'}">${t.count}/${RULES.bronzeSessionsPerTheme}</span></div>
    <ul class="checklist">${t.badges.map(b => `<li><span class="mark done" aria-hidden="true">&#10003;</span><span>${esc(b.name)}<br><span class="small muted">${fmtDate(b.date)}${b.delivery ? ' &middot; ' + esc(b.delivery) : ''}</span></span></li>`).join('')}</ul>`).join('');
  const awards = a.badges.filter(b => SILVER_AWARD_KEYS.includes(b.key) || GOLD_AWARD_KEYS.includes(b.key));
  const other = a.unmapped;
  const completedHtml = (completed || '') +
    (awards.length ? `<div class="group-title">Tier awards</div><ul class="checklist">${awards.map(b => `<li><span class="mark done" aria-hidden="true">&#10003;</span><span>${esc(b.name)}<br><span class="small muted">${fmtDate(b.date)}</span></span></li>`).join('')}</ul>` : '') +
    (other.length ? `<div class="group-title">Not yet assigned to a theme</div><ul class="checklist">${other.map(b => `<li><span class="mark done" aria-hidden="true">&#10003;</span><span>${esc(b.name)}<br><span class="small muted">${fmtDate(b.date)}</span></span></li>`).join('')}</ul>` : '') || '<p class="muted">No completed activities yet.</p>';

  // missing & gap
  const mk = (label, items) => items.length ? `<div class="group-title">${label}</div><ul class="checklist">${items.map(i => `<li><span class="mark todo" aria-hidden="true">&#9675;</span><span>${esc(i.text)}${i.pending ? ' <span class="chip pending">Not in report</span>' : ''}</span></li>`).join('')}</ul>` : '';
  const bm = a.bronze.earned ? [{ text: `Discipline-specific T&L activity, up to ${RULES.disciplinePointsMax} points`, pending: true }] : bronzeMissing(a);
  const sm = silverMissing(a), gm = goldMissing(a);
  const missingHtml = mk('Bronze', bm) + mk('Silver', sm) + mk('Gold', gm);
  const gapCard = (label, t, items, lead) => `<div class="card"><h3>${label} ${chip(a[t].status)}</h3><p class="lead" style="font-size:1.05rem;font-weight:600">${esc(lead)}</p><ul class="small" style="margin:.3rem 0 0 1.1rem;padding:0">${items.map(i => `<li>${esc(i.text)}${i.pending ? ' (not in report)' : ''}</li>`).join('') || '<li>Nothing outstanding.</li>'}</ul></div>`;
  const bLead = a.bronze.earned ? 'Session requirement met' : `${plural(sug.remaining, 'session', 'sessions')} to go`;
  const sLead = a.silver.earned ? 'Awarded' : (a.bronze.earned ? 'Evidence needed' : (a.silver.emcHave ? `${5 - a.silver.emcHave} EMC courses or Bronze + evidence` : 'Bronze + evidence needed'));
  const gLead = a.gold.earned ? 'Awarded' : (a.silver.earned ? 'Contribution evidence needed' : 'Locked until Silver');

  // advisor
  const adv = [];
  const names = sug.picks.slice(0, 3).map(x => x.badge.name);
  if (!a.bronze.earned) {
    adv.push(`<p class="lead">You are ${plural(sug.remaining, 'activity', 'activities')} away from achieving Bronze Tier.</p>`);
    if (names.length) {
      adv.push(`<p>${names.length > 1 ? 'Complete ' + names.slice(0, -1).map(esc).join(', ') + ' and ' + esc(names[names.length - 1]) : 'Complete ' + esc(names[0])} to close the gap.</p>`);
      adv.push(`<p class="small muted">Each session adds about ${pct(sug.picks[0].gain)}% to Bronze progress. Finishing a theme you have already started gives the quickest result.</p>`);
      adv.push(`<ul>${sug.picks.slice(0, 5).map(x => `<li>${esc(x.badge.name)} <span class="chip theme"><span class="dot" style="background:${themeColor(x.theme.id)}"></span>${esc(x.theme.short)}</span> <span class="small muted">+${pct(x.gain)}% Bronze</span></li>`).join('')}</ul>`);
    } else {
      adv.push('<p>No unfinished sessions were found in the themes you have started. Check the TLC catalogue for another theme.</p>');
    }
    const short = sug.needs.filter(n => n.available < n.need);
    if (short.length) adv.push(`<p class="small muted">The report lists too few sessions for ${short.map(n => esc(n.theme.short)).join(', ')}; check the TLC catalogue for more.</p>`);
    if (a.silver.emcHave > 0) adv.push(`<p class="small">Another route: complete all five Essentials of Moodle courses (${a.silver.emcHave} of 5 done) for automatic Silver.</p>`);
  } else if (!a.silver.earned) {
    adv.push('<p class="lead">Bronze session requirement met. Silver is next.</p>');
    adv.push('<p>Submit your application example and reflective statement to the TLC (tlc@ajman.ac.ae), or complete the Essentials of Moodle Certificate.</p>');
    if (a.silver.emcMissing.length < 5) adv.push(`<p>Complete ${a.silver.emcMissing.slice(0, 2).map(k => esc((S.catalog.get(k) || { name: k }).name)).join(' and ')} to reach Silver through the programme route.</p>`);
    adv.push('<p class="small muted">Also confirm your discipline-specific activity points (up to 5), which the badge report does not record.</p>');
  } else if (!a.gold.earned) {
    adv.push('<p class="lead">Silver Tier achieved. Gold is the next step.</p>');
    adv.push('<p>Gold needs a contribution beyond your own teaching (for example leading a workshop or mentoring colleagues) or Advance HE Fellowship D3-D4. Speak to the TLC to record it; the badge report does not track this.</p>');
  } else {
    adv.push('<p class="lead">Gold Tier achieved. Thank you for your contribution.</p>');
  }

  // timeline
  const events = a.badges.map(b => ({ date: b.date, title: b.name, sub: (S.themeMap[b.key] ? THEMES.find(t => t.id === S.themeMap[b.key]).short : (SILVER_AWARD_KEYS.includes(b.key) ? 'Tier award' : 'Unassigned theme')) + (b.delivery ? ' · ' + b.delivery : '') }));
  if (a.bronze.earned) events.push({ date: a.bronze.date, title: 'Bronze Tier requirement met', sub: 'Milestone', ms: true });
  if (a.silver.earned) events.push({ date: a.silver.date, title: 'Silver Tier achieved', sub: 'Milestone', ms: true });
  if (a.gold.earned) events.push({ date: a.gold.date, title: 'Gold Tier achieved', sub: 'Milestone', ms: true });
  events.sort((x, y) => (x.date ? x.date.getTime() : 0) - (y.date ? y.date.getTime() : 0) || (x.ms ? 1 : 0) - (y.ms ? 1 : 0));
  const timeline = `<ol class="timeline">${events.map(e => `<li class="${e.ms ? 'milestone' : ''}"><time datetime="${e.date ? e.date.toISOString().slice(0, 10) : ''}">${fmtDate(e.date)}</time><span class="t">${esc(e.title)}</span><span class="small muted">${esc(e.sub)}</span></li>`).join('')}</ol>`;

  root.innerHTML = `
    <div class="card profile" style="margin-bottom:16px">
      <div class="avatar" aria-hidden="true">${esc(initials(p.name))}</div>
      <div><h2 dir="auto">${esc(p.name)}</h2>
        <dl><dt>Email</dt><dd>${esc(p.email)}</dd><dt>College</dt><dd>${esc(p.college)}</dd>
        <dt>Department</dt><dd>${p.dept ? esc(p.dept) : '<span class="chip pending">Not in report</span>'}</dd>
        ${p.rank ? `<dt>Rank</dt><dd>${esc(p.rank)}</dd>` : ''}</dl></div>
    </div>
    <div class="actions no-print">
      <button class="btn btn-primary" type="button" data-act="ind-pdf">Export PDF</button>
      <button class="btn" type="button" data-act="print">Print dashboard</button>
      <button class="btn" type="button" data-act="ind-xlsx">Download report</button>
    </div>
    <div class="grid g4">
      ${kpi('Total courses', a.totalCourses, 'Distinct course badges')}
      ${kpi('Total learning hours', '—', 'Not in badge report')}
      ${kpi('Badges earned', badgesEarned, `${a.totalCourses} course + ${earnedTiers} tier`)}
      ${kpi('Current tier', tierName(a.tier), a.tier === 'none' ? 'Bronze not yet met' : 'Highest tier reached', a.tier === 'none' ? '' : 'tier-' + a.tier)}
    </div>
    <h2 class="section-title">Professional development progress</h2>
    <div class="grid g3">
      ${tierCard('bronze', 'Bronze', bronzeSub)}${tierCard('silver', 'Silver', silverSub)}${tierCard('gold', 'Gold', goldSub)}
    </div>
    <div class="card" style="margin-bottom:16px"><h3>Progress by theme</h3>
      ${a.themes.map(t => `<div class="bar-row"><span><span class="dot" style="background:${themeColor(t.id)}"></span> ${esc(t.short)}</span>${bar(Math.min(1, t.count / RULES.bronzeSessionsPerTheme), themeColor(t.id), t.short + ' sessions')}<span>${t.count}/${RULES.bronzeSessionsPerTheme}</span></div>`).join('')}
      <p class="small muted">Bronze needs ${RULES.bronzeSessionsPerTheme} sessions in each of ${RULES.bronzeThemes} themes. Theme assignment is a draft until confirmed by the TLC.</p></div>
    <h2 class="section-title">Badge gallery</h2>
    <div class="card" style="margin-bottom:16px"><div class="gallery">${tileTier('bronze', 'Bronze')}${tileTier('silver', 'Silver')}${tileTier('gold', 'Gold')}${special.map(tileSpecial).join('')}</div>
      <h3 style="margin:16px 0 8px">Theme badges</h3><div class="gallery">${a.themes.map(tileTheme).join('')}</div></div>
    <div class="grid g2">
      <div class="card"><h2>Completed requirements</h2>${completedHtml}</div>
      <div class="card"><h2>Missing requirements</h2>${missingHtml || '<p class="muted">Nothing outstanding.</p>'}</div>
    </div>
    <h2 class="section-title">Gap analysis</h2>
    <div class="grid g3">${gapCard('Bronze', 'bronze', bm, bLead)}${gapCard('Silver', 'silver', sm, sLead)}${gapCard('Gold', 'gold', gm, gLead)}</div>
    <div class="card advisor" style="margin-bottom:16px"><h2>AI progress advisor</h2>${adv.join('')}
      <p class="small muted">Rule-based guidance computed from this report; it does not use an external AI service.</p></div>
    <div class="card"><h2>Achievement timeline</h2>${timeline}</div>`;
}

/* ==========================================================================
   7. COLLEGE VIEW
   ========================================================================== */
function collegeFiltered() {
  return S.analyses.filter(a => passFilter(a, S.filters.college, S.filters.status));
}
function fillSelect(sel, opts, value, allLabel) {
  sel.innerHTML = `<option value="all">${esc(allLabel)}</option>` + opts.map(o => `<option value="${esc(o)}">${esc(o)}</option>`).join('');
  sel.value = opts.includes(value) ? value : 'all';
  return sel.value;
}
function refreshFilters() {
  const colleges = Array.from(new Set(S.analyses.map(a => a.p.college))).sort();
  S.filters.college = fillSelect($('#f-college'), colleges, S.filters.college, 'All colleges');
  $('#f-status').value = S.filters.status;
  $$('.year-select').forEach(sel => {
    sel.innerHTML = '<option value="all">All years</option>' + S.years.map(y => `<option value="${esc(y)}">${esc(y)}</option>`).join('');
    sel.value = S.year;
  });
}

function chartCard(id, title, desc, heightPx) {
  return `<div class="card"><h3>${esc(title)}</h3><div class="chart-box" style="${heightPx ? 'height:' + heightPx + 'px' : ''}"><canvas id="${id}" role="img" aria-label="${esc(desc)}"></canvas></div><div id="${id}-alt"></div></div>`;
}
function leaderboardPeople(list) {
  const top = list.slice().sort((x, y) => y.tierRank - x.tierRank || y.progress - x.progress || y.totalCourses - x.totalCourses).slice(0, 10);
  return table([{ h: '#', num: 1 }, { h: 'Faculty', dir: 1 }, { h: 'College' }, { h: 'Tier' }, { h: 'Courses', num: 1 }],
    top.map((a, i) => [i + 1, esc(a.p.name), esc(a.p.college), `<span style="color:${tierColor(a.tier)};font-weight:700">${tierName(a.tier)}</span>`, a.totalCourses]),
    { caption: 'Top faculty (by tier, then progress, then courses)' });
}
function leaderboardGroups(groups, label) {
  const top = groups.slice().sort((x, y) => y.avg - x.avg || y.rate - x.rate).slice(0, 10);
  return table([{ h: '#', num: 1 }, { h: label }, { h: 'Faculty', num: 1 }, { h: 'Bronze+', num: 1 }, { h: 'Avg progress', num: 1 }],
    top.map((g, i) => [i + 1, esc(g.name), g.n, pct(g.rate) + '%', pct(g.avg) + '%']),
    { caption: 'Top ' + label.toLowerCase() + 's (by average progress)' });
}
function themeCompletion(list) {
  return THEMES.map(t => ({ t, rate: list.length ? list.filter(a => a.themes.find(x => x.id === t.id).done).length / list.length : 0 }));
}
function collegeInsights(list, groups, groupLabel) {
  const strengths = [], gaps = [], recs = [];
  if (!list.length) return { strengths: ['No faculty match the current filters.'], gaps: [], recs: [] };
  const tc = countTiers(list), n = list.length;
  const themeRates = themeCompletion(list).sort((a, b) => b.rate - a.rate);
  strengths.push(`${pct((n - tc.none) / n)}% of faculty (${n - tc.none} of ${n}) have met Bronze or higher.`);
  if (themeRates[0].rate > 0) strengths.push(`Strongest theme: ${themeRates[0].t.short}, completed by ${pct(themeRates[0].rate)}% of faculty.`);
  const gs = groups.filter(g => g.n >= 5).sort((a, b) => b.rate - a.rate);
  if (gs.length) strengths.push(`${groupLabel} with the highest Bronze+ rate: ${gs[0].name} (${pct(gs[0].rate)}%, ${gs[0].n} faculty).`);
  gaps.push(`${tc.none} faculty (${pct(tc.none / n)}%) have not yet met Bronze.`);
  const low = themeRates[themeRates.length - 1];
  gaps.push(`Weakest theme: ${low.t.short}, completed by only ${pct(low.rate)}% of faculty.`);
  const close = list.filter(a => a.tier === 'none' && a.bronze.pts >= 4).length;
  if (close) gaps.push(`${close} faculty are within 2 sessions of Bronze.`);
  if (tc.gold === 0) gaps.push('No Gold achievers recorded; the report holds no Gold badge or contribution evidence.');
  if (close) recs.push(`Send a targeted reminder to the ${close} faculty who are 1 to 2 sessions from Bronze.`);
  recs.push(`Offer an extra ${low.t.short} session: it is the least completed theme.`);
  if (tc.bronze > 0) recs.push(`Invite the ${tc.bronze} Bronze-only faculty to submit Silver evidence (application example and reflection).`);
  const quiet = gs.length ? gs[gs.length - 1] : null;
  if (quiet && quiet !== gs[0]) recs.push(`Prioritise outreach to ${quiet.name}, which has the lowest Bronze+ rate among larger groups (${pct(quiet.rate)}%).`);
  return { strengths, gaps, recs };
}
const insightCard = (cls, title, items) => `<div class="card insight ${cls}"><h3>${title}</h3><ul>${items.map(i => `<li>${esc(i)}</li>`).join('') || '<li>None identified.</li>'}</ul></div>`;

function renderCollege() {
  const list = collegeFiltered(); const root = $('#college-content');
  const tc = countTiers(list);
  const byGroup = groupBy(list, a => a.p.college);
  const groupLabel = 'College';
  const ins = collegeInsights(list, byGroup, groupLabel);
  S.collegeView = { list, tc, byGroup, groupLabel };
  root.innerHTML = `
    <div class="actions no-print">
      <button class="btn btn-primary" type="button" data-act="college-pdf">Export PDF</button>
      <button class="btn" type="button" data-act="college-xlsx">Export Excel</button>
      <button class="btn" type="button" data-act="print">Print</button>
    </div>
    <div class="grid g5">
      ${kpi('Total faculty', list.length, 'In the report, current filters')}
      ${kpi('Bronze achievers', tc.bronze, 'Highest tier is Bronze', 'tier-bronze')}
      ${kpi('Silver achievers', tc.silver, 'Highest tier is Silver', 'tier-silver')}
      ${kpi('Gold achievers', tc.gold, 'Highest tier is Gold', 'tier-gold')}
      ${kpi('In progress', tc.none, 'Badges earned, Bronze not yet met')}
    </div>
    <div class="grid g2">
      ${chartCard('ch-c-tier', 'Tier distribution', 'Doughnut chart of faculty by highest tier')}
      ${chartCard('ch-c-badge', 'Badge achievement', 'Bar chart of the most earned badges', 340)}
      ${chartCard('ch-c-part', 'Faculty participation', 'Column chart of faculty earning badges per month')}
      ${chartCard('ch-c-comp', 'Completion rates', 'Bar chart of the share of faculty who completed each theme and Bronze')}
    </div>
    <h2 class="section-title">Leaderboards</h2>
    <div class="grid g2"><div class="card">${leaderboardPeople(list)}</div><div class="card">${leaderboardGroups(byGroup, groupLabel)}</div></div>
    <h2 class="section-title">Insights</h2>
    <div class="grid g3">${insightCard('strength', 'Strengths', ins.strengths)}${insightCard('gap', 'Gaps', ins.gaps)}${insightCard('rec', 'Recommendations', ins.recs)}</div>`;
  drawCharts('ch-c', collegeCharts(list, tc));
}

/* ==========================================================================
   9. CHARTS
   ========================================================================== */
function palette(forceLight) {
  if (forceLight) {
    return { text: '#17232e', muted: '#465868', grid: '#dfe7ee', surface: '#ffffff', primary: '#0a5c8c', accent: '#3aa8dc', none: '#8795a3', bronze: '#9a5320', silver: '#5b6773', gold: '#8a6500', themes: ['#e8731a', '#e0a800', '#14877f', '#8c2f45', '#7a5230'] };
  }
  const cs = getComputedStyle(document.documentElement); const g = n => cs.getPropertyValue(n).trim();
  return { text: g('--text'), muted: g('--muted'), grid: g('--grid'), surface: g('--surface'), primary: g('--primary'), accent: g('--accent'), none: g('--tier-none'), bronze: g('--tier-bronze'), silver: g('--tier-silver'), gold: g('--tier-gold'), themes: [1, 2, 3, 4, 5].map(i => g('--theme-' + i)) };
}
function baseOptions(P, extra) {
  const o = { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: P.text, boxWidth: 14 } }, tooltip: { } }, color: P.text };
  return Object.assign(o, extra || {});
}
function axes(P, o) {
  o = o || {};
  const tick = { color: P.muted, font: { size: 11 } };
  const mk = (stacked, title, extra) => Object.assign({ stacked: !!stacked, ticks: tick, grid: { color: P.grid }, border: { color: P.grid }, title: title ? { display: true, text: title, color: P.muted } : { display: false } }, extra || {});
  return { x: mk(o.stacked, o.xTitle, o.xExtra), y: mk(o.stacked, o.yTitle, Object.assign({ beginAtZero: true }, o.yExtra)) };
}
const short = (s, n) => (s.length > n ? s.slice(0, n - 1) + '…' : s);

function collegeCharts(list, tc) {
  return P => {
    const ms = monthlyStats(list);
    const badgeCounts = new Map();
    list.forEach(a => a.badges.forEach(b => { if (!SILVER_AWARD_KEYS.includes(b.key)) badgeCounts.set(b.name, (badgeCounts.get(b.name) || 0) + 1); }));
    const top = Array.from(badgeCounts).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const tr = themeCompletion(list);
    const n = list.length || 1;
    return {
      tier: { type: 'doughnut', data: { labels: ['No tier yet', 'Bronze', 'Silver', 'Gold'], datasets: [{ data: [tc.none, tc.bronze, tc.silver, tc.gold], backgroundColor: [P.none, P.bronze, P.silver, P.gold], borderColor: P.surface, borderWidth: 2 }] }, options: baseOptions(P, { plugins: { legend: { position: 'bottom', labels: { color: P.text } } } }) },
      badge: { type: 'bar', data: { labels: top.map(t => short(t[0], 36)), datasets: [{ label: 'Faculty who earned it', data: top.map(t => t[1]), backgroundColor: P.accent }] }, options: baseOptions(P, { indexAxis: 'y', plugins: { legend: { display: false } }, scales: axes(P, { xTitle: 'Faculty' }) }) },
      part: { type: 'bar', data: { labels: ms.months.map(monthLabel), datasets: [{ label: 'Faculty earning a badge', data: ms.faculty, backgroundColor: P.primary }] }, options: baseOptions(P, { plugins: { legend: { display: false } }, scales: axes(P, { yTitle: 'Faculty' }) }) },
      comp: { type: 'bar', data: { labels: THEMES.map(t => short(t.short, 22)).concat(['Bronze']), datasets: [{ label: '% of faculty', data: tr.map(x => Math.round(x.rate * 100)).concat([Math.round((n - tc.none) / n * 100)]), backgroundColor: P.themes.concat([P.bronze]) }] }, options: baseOptions(P, { plugins: { legend: { display: false } }, scales: axes(P, { yTitle: '% of faculty', yExtra: { max: 100 } }) }) }
    };
  };
}
function finalizeCfg(cfg, exporting) {
  if (exporting) {
    cfg.options = Object.assign({}, cfg.options, { responsive: false, animation: false, devicePixelRatio: 1.5 });
  }
  return cfg;
}
function altTable(cfg) {
  const labels = cfg.data.labels; const ds = cfg.data.datasets;
  const head = ['Category'].concat(cfg.type === 'doughnut' ? ['Faculty'] : ds.map(d => d.label));
  const rows = labels.map((l, i) => [esc(l)].concat(cfg.type === 'doughnut' ? [ds[0].data[i]] : ds.map(d => d[i] !== undefined ? d[i] : d.data[i])));
  return `<details class="data-alt"><summary>View data as a table</summary>${table(head.map((h, i) => ({ h, num: i > 0 })), rows)}</details>`;
}
function drawCharts(prefix, builder) {
  if (typeof Chart === 'undefined') { $$('canvas[id^="' + prefix + '"]').forEach(c => { c.parentElement.innerHTML = '<p class="placeholder">The chart library did not load. Check your connection and reload.</p>'; }); return; }
  const cfgs = builder(palette(false));
  Object.keys(S.charts).forEach(id => { if (id.startsWith(prefix)) { S.charts[id].destroy(); delete S.charts[id]; } });
  Object.keys(cfgs).forEach(key => {
    const id = prefix + '-' + key; const el = document.getElementById(id); if (!el) return;
    S.charts[id] = new Chart(el, finalizeCfg(cfgs[key], false));
    const alt = document.getElementById(id + '-alt'); if (alt) alt.innerHTML = altTable(cfgs[key]);
  });
}
function renderToImages(builder) {
  const cfgs = builder(palette(true)); const out = {};
  Object.keys(cfgs).forEach(key => {
    const c = document.createElement('canvas'); c.width = 900; c.height = key === 'college' || key === 'badge' ? 520 : 400;
    c.style.width = c.width + 'px'; c.style.height = c.height + 'px';
    const white = { id: 'bg', beforeDraw(ch) { const x = ch.ctx; x.save(); x.fillStyle = '#fff'; x.fillRect(0, 0, ch.width, ch.height); x.restore(); } };
    const cfg = finalizeCfg(cfgs[key], true); cfg.plugins = [white];
    const ch = new Chart(c, cfg);
    out[key] = { url: c.toDataURL('image/jpeg', 0.9), w: c.width, h: c.height }; ch.destroy();
  });
  return out;
}

/* ==========================================================================
   10. EXPORTS
   ========================================================================== */
class Pdf {
  constructor(title, subtitle) {
    if (!window.jspdf) throw new Error('The PDF library did not load. Check your connection and reload.');
    this.d = new window.jspdf.jsPDF({ unit: 'pt', format: 'a4' });
    this.W = 595; this.H = 842; this.m = 40; this.y = 0; this.title = title;
    this.header(title, subtitle);
  }
  header(title, subtitle) {
    const d = this.d; d.setFillColor(10, 92, 140); d.rect(0, 0, this.W, 64, 'F');
    d.setTextColor(255); d.setFont('helvetica', 'bold'); d.setFontSize(16); d.text(title, this.m, 30);
    d.setFont('helvetica', 'normal'); d.setFontSize(9.5); d.text(subtitle, this.m, 48);
    d.setTextColor(23, 35, 46); this.y = 88;
  }
  need(h) { if (this.y + h > this.H - 50) { this.d.addPage(); this.y = 50; } }
  h2(t) { this.need(34); this.y += 6; this.d.setFont('helvetica', 'bold'); this.d.setFontSize(12.5); this.d.setTextColor(10, 92, 140); this.d.text(t, this.m, this.y); this.d.setDrawColor(211, 221, 230); this.d.line(this.m, this.y + 4, this.W - this.m, this.y + 4); this.d.setTextColor(23, 35, 46); this.y += 20; }
  p(t, o) {
    o = o || {}; const d = this.d; d.setFont('helvetica', o.bold ? 'bold' : 'normal'); d.setFontSize(o.size || 10);
    if (o.color) d.setTextColor.apply(d, o.color); else d.setTextColor(23, 35, 46);
    const lines = d.splitTextToSize(String(t), this.W - 2 * this.m - (o.indent || 0));
    lines.forEach(l => { this.need(14); d.text(l, this.m + (o.indent || 0), this.y); this.y += (o.size || 10) + 4; });
    d.setTextColor(23, 35, 46);
  }
  bullets(items, o) { items.forEach(t => this.p('• ' + t, Object.assign({ indent: 8 }, o || {}))); }
  kv(rows) { rows.forEach(([k, v]) => { this.need(14); this.d.setFont('helvetica', 'bold'); this.d.setFontSize(10); this.d.text(k, this.m, this.y); this.d.setFont('helvetica', 'normal'); this.d.text(String(v), this.m + 110, this.y); this.y += 14; }); }
  kpis(items) {
    const w = (this.W - 2 * this.m - (items.length - 1) * 8) / items.length; this.need(54);
    items.forEach((it, i) => { const x = this.m + i * (w + 8); this.d.setDrawColor(211, 221, 230); this.d.setFillColor(243, 246, 249); this.d.roundedRect(x, this.y, w, 46, 4, 4, 'FD'); this.d.setFont('helvetica', 'bold'); this.d.setFontSize(17); this.d.setTextColor(10, 92, 140); this.d.text(String(it[1]), x + 8, this.y + 22); this.d.setFont('helvetica', 'normal'); this.d.setFontSize(8); this.d.setTextColor(70, 88, 104); this.d.text(it[0], x + 8, this.y + 37); });
    this.d.setTextColor(23, 35, 46); this.y += 58;
  }
  table(cols, rows, widths) {
    const d = this.d; const tw = this.W - 2 * this.m; const tot = widths.reduce((a, b) => a + b, 0); const ws = widths.map(w => w / tot * tw);
    const row = (cells, head) => {
      d.setFont('helvetica', head ? 'bold' : 'normal'); d.setFontSize(8.5);
      const wrapped = cells.map((c, i) => d.splitTextToSize(String(c), ws[i] - 8)); const h = Math.max(...wrapped.map(w => w.length)) * 10.5 + 6;
      this.need(h + 2);
      if (head) { d.setFillColor(234, 240, 245); d.rect(this.m, this.y - 9, tw, h, 'F'); }
      let x = this.m; wrapped.forEach((w, i) => { d.text(w, x + 4, this.y); x += ws[i]; });
      d.setDrawColor(223, 231, 238); d.line(this.m, this.y - 9 + h, this.W - this.m, this.y - 9 + h); this.y += h;
    };
    row(cols, true); rows.forEach(r => row(r, false)); this.y += 6;
  }
  image(img, maxW) {
    const w = Math.min(maxW || this.W - 2 * this.m, this.W - 2 * this.m); const h = w * img.h / img.w; this.need(h + 8);
    this.d.addImage(img.url, 'JPEG', this.m, this.y, w, h); this.y += h + 10;
  }
  save(name) {
    const n = this.d.getNumberOfPages(); const stamp = 'Generated ' + new Date().toLocaleString('en-GB');
    for (let i = 1; i <= n; i++) { this.d.setPage(i); this.d.setFont('helvetica', 'normal'); this.d.setFontSize(8); this.d.setTextColor(110); this.d.text(stamp, this.m, this.H - 24); this.d.text('Page ' + i + ' of ' + n, this.W - this.m - 50, this.H - 24); }
    this.d.save(name);
  }
}
const fileStamp = () => new Date().toISOString().slice(0, 10);
const yearText = () => S.year === 'all' ? 'All years' : 'Academic year ' + S.year;

function exportIndividualPdf() {
  const a = S.analyses.find(x => x.p.email === S.selected); if (!a) return;
  const pdf = new Pdf('Faculty PD Achievement Report', 'Ajman University | Teaching & Learning Center | ' + yearText());
  const sug = suggestions(a);
  pdf.h2('Profile');
  pdf.kv([['Name', latinOnly(a.p.name) || '(name in Arabic only)'], ['Email', a.p.email], ['College', a.p.college], ['Department', a.p.dept || 'Not in report'], ['Current tier', tierName(a.tier)]]);
  pdf.kpis([['Total courses', a.totalCourses], ['Learning hours', 'n/a'], ['Badges earned', a.totalCourses + ['bronze', 'silver', 'gold'].filter(t => a[t].earned).length], ['Tier', tierName(a.tier)]]);
  pdf.h2('Tier progress');
  pdf.table(['Tier', 'Status', 'Progress', 'Note'], [
    ['Bronze', ucfirst(a.bronze.status), pct(a.bronze.progress) + '%', `${a.bronze.pts} of ${a.bronze.need} sessions; discipline points not in report`],
    ['Silver', ucfirst(a.silver.status), pct(a.silver.progress) + '%', a.silver.earned ? 'Awarded' : `EMC ${a.silver.emcHave} of 5; evidence pending`],
    ['Gold', ucfirst(a.gold.status), pct(a.gold.progress) + '%', a.gold.earned ? 'Awarded' : 'Contribution evidence not in report']], [1, 1, 1, 4]);
  pdf.h2('Progress by theme');
  pdf.table(['Theme', 'Sessions', 'Status'], a.themes.map(t => [t.name, t.count + ' of ' + RULES.bronzeSessionsPerTheme, t.done ? 'Complete' : t.count ? 'In progress' : 'Not started']), [5, 1.2, 1.5]);
  pdf.h2('Completed requirements');
  a.themes.filter(t => t.badges.length).forEach(t => { pdf.p(t.name, { bold: true }); pdf.bullets(t.badges.map(b => `${b.name} (${fmtDate(b.date)})`)); });
  const aw = a.badges.filter(b => SILVER_AWARD_KEYS.includes(b.key)); if (aw.length) { pdf.p('Tier awards', { bold: true }); pdf.bullets(aw.map(b => `${b.name} (${fmtDate(b.date)})`)); }
  pdf.h2('Missing requirements and gap analysis');
  const bm = a.bronze.earned ? [{ text: 'Discipline-specific T&L activity, up to 5 points (not in report)' }] : bronzeMissing(a);
  [['Bronze', bm], ['Silver', silverMissing(a)], ['Gold', goldMissing(a)]].forEach(([l, items]) => { pdf.p(l, { bold: true }); pdf.bullets(items.length ? items.map(i => i.text + (i.pending ? ' (not in report)' : '')) : ['Nothing outstanding']); });
  pdf.h2('Progress advisor');
  if (!a.bronze.earned) { pdf.p(`You are ${plural(sug.remaining, 'activity', 'activities')} away from achieving Bronze Tier.`, { bold: true }); if (sug.picks.length) pdf.bullets(sug.picks.slice(0, 5).map(x => `${x.badge.name} (${x.theme.short}, +${pct(x.gain)}% Bronze)`)); }
  else if (!a.silver.earned) pdf.p('Bronze session requirement met. Submit an application example and reflective statement to the TLC, or complete the Essentials of Moodle Certificate, for Silver.');
  else if (!a.gold.earned) pdf.p('Silver achieved. Gold needs a contribution beyond your own teaching; speak to the TLC to record it.');
  else pdf.p('Gold Tier achieved.');
  pdf.h2('Achievement timeline');
  pdf.table(['Date', 'Achievement', 'Theme'], a.badges.map(b => [fmtDate(b.date), b.name, S.themeMap[b.key] ? THEMES.find(t => t.id === S.themeMap[b.key]).short : (SILVER_AWARD_KEYS.includes(b.key) ? 'Tier award' : '')]), [1.3, 5, 2]);
  pdf.save('PD-report-' + (latinOnly(a.p.name).replace(/\s+/g, '-') || 'faculty') + '-' + fileStamp() + '.pdf');
}
function exportIndividualXlsx() {
  const a = S.analyses.find(x => x.p.email === S.selected); if (!a) return;
  const wb = XLSX.utils.book_new();
  const add = (name, aoa, widths) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = widths.map(w => ({ wch: w })); XLSX.utils.book_append_sheet(wb, ws, name); };
  add('Profile', [['Field', 'Value'], ['Name', a.p.name], ['Email', a.p.email], ['College', a.p.college], ['Department', a.p.dept || 'Not in report'], ['Academic year filter', yearText()], ['Current tier', tierName(a.tier)], ['Total courses', a.totalCourses], ['Learning hours', 'Not in report']], [24, 50]);
  add('Tiers', [['Tier', 'Status', 'Progress %', 'Note'], ['Bronze', a.bronze.status, pct(a.bronze.progress), `${a.bronze.pts} of ${a.bronze.need} sessions; discipline points not in report`], ['Silver', a.silver.status, pct(a.silver.progress), a.silver.earned ? 'Awarded' : 'Evidence pending'], ['Gold', a.gold.status, pct(a.gold.progress), a.gold.earned ? 'Awarded' : 'Evidence not in report']], [10, 12, 12, 60]);
  add('Themes', [['Theme', 'Sessions completed', 'Sessions needed', 'Complete']].concat(a.themes.map(t => [t.name, t.count, RULES.bronzeSessionsPerTheme, t.done ? 'Yes' : 'No'])), [55, 18, 16, 10]);
  add('Badges', [['Badge', 'Theme', 'Date earned', 'Delivery']].concat(a.badges.map(b => [b.name, S.themeMap[b.key] ? THEMES.find(t => t.id === S.themeMap[b.key]).name : '', b.date ? fmtDate(b.date) : '', b.delivery])), [60, 45, 16, 14]);
  const miss = []; [['Bronze', a.bronze.earned ? [{ text: 'Discipline-specific activity, up to 5 points', pending: true }] : bronzeMissing(a)], ['Silver', silverMissing(a)], ['Gold', goldMissing(a)]].forEach(([t, items]) => items.forEach(i => miss.push([t, i.text, i.pending ? 'Not in report' : ''])));
  add('Missing', [['Tier', 'Requirement', 'Data note']].concat(miss), [10, 80, 16]);
  XLSX.writeFile(wb, 'PD-report-' + (latinOnly(a.p.name).replace(/\s+/g, '-') || 'faculty') + '-' + fileStamp() + '.xlsx');
}
function facultyRows(list) {
  return [['Name', 'Email', 'College', 'Department', 'Tier', 'Bronze %', 'Silver %', 'Gold %', 'Courses', 'Active', 'Last badge date'].concat(THEMES.map(t => t.short + ' sessions'))]
    .concat(list.map(a => [a.p.name, a.p.email, a.p.college, a.p.dept || 'Not in report', tierName(a.tier), pct(a.bronze.progress), pct(a.silver.progress), pct(a.gold.progress), a.totalCourses, a.active ? 'Yes' : 'No', a.lastDate ? fmtDate(a.lastDate) : ''].concat(a.themes.map(t => t.count))));
}
function exportCollegePdf() {
  const v = S.collegeView; const pdf = new Pdf('College PD Achievement Report', `Ajman University | ${yearText()} | College: ${S.filters.college === 'all' ? 'All' : S.filters.college}`);
  pdf.kpis([['Total faculty', v.list.length], ['Bronze', v.tc.bronze], ['Silver', v.tc.silver], ['Gold', v.tc.gold], ['In progress', v.tc.none]]);
  const imgs = renderToImages(collegeCharts(v.list, v.tc));
  pdf.h2('Tier distribution and badge achievement'); pdf.image(imgs.tier, 300); pdf.image(imgs.badge);
  pdf.h2('Participation and completion rates'); pdf.image(imgs.part); pdf.image(imgs.comp);
  pdf.h2('Leaderboard: top faculty');
  pdf.table(['#', 'Faculty', 'College', 'Tier', 'Courses'], v.list.slice().sort((x, y) => y.tierRank - x.tierRank || y.progress - x.progress || y.totalCourses - x.totalCourses).slice(0, 10).map((a, i) => [i + 1, latinOnly(a.p.name) || a.p.email, a.p.college, tierName(a.tier), a.totalCourses]), [0.5, 3, 4, 1.4, 1]);
  pdf.h2('Leaderboard: top ' + v.groupLabel.toLowerCase() + 's');
  pdf.table(['#', v.groupLabel, 'Faculty', 'Bronze+', 'Avg progress'], v.byGroup.slice().sort((x, y) => y.avg - x.avg).slice(0, 10).map((g, i) => [i + 1, g.name, g.n, pct(g.rate) + '%', pct(g.avg) + '%']), [0.5, 5, 1.2, 1.2, 1.4]);
  const ins = collegeInsights(v.list, v.byGroup, v.groupLabel);
  pdf.h2('Insights'); pdf.p('Strengths', { bold: true }); pdf.bullets(ins.strengths); pdf.p('Gaps', { bold: true }); pdf.bullets(ins.gaps); pdf.p('Recommendations', { bold: true }); pdf.bullets(ins.recs);
  pdf.save('College-PD-report-' + fileStamp() + '.pdf');
}
function exportCollegeXlsx() {
  const v = S.collegeView; const wb = XLSX.utils.book_new();
  const ins = collegeInsights(v.list, v.byGroup, v.groupLabel);
  const sum = [['Measure', 'Value'], ['Academic year', yearText()], ['College filter', S.filters.college], ['Total faculty', v.list.length], ['Bronze achievers', v.tc.bronze], ['Silver achievers', v.tc.silver], ['Gold achievers', v.tc.gold], ['In progress', v.tc.none], [], ['Strengths'], ...ins.strengths.map(s => [s]), [], ['Gaps'], ...ins.gaps.map(s => [s]), [], ['Recommendations'], ...ins.recs.map(s => [s])];
  const add = (n, aoa, w) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = w.map(x => ({ wch: x })); XLSX.utils.book_append_sheet(wb, ws, n); };
  add('Summary', sum, [40, 60]);
  add('Faculty', facultyRows(v.list), [34, 30, 40, 24, 12, 10, 10, 10, 9, 8, 14, 12, 12, 12, 12, 12]);
  add(v.groupLabel + 's', [[v.groupLabel, 'Faculty', 'No tier', 'Bronze', 'Silver', 'Gold', 'Bronze+ %', 'Avg progress %']].concat(v.byGroup.map(g => [g.name, g.n, g.tiers.none, g.tiers.bronze, g.tiers.silver, g.tiers.gold, pct(g.rate), pct(g.avg)])), [44, 9, 9, 9, 9, 9, 11, 14]);
  const bc = new Map(); v.list.forEach(a => a.badges.forEach(b => bc.set(b.name, (bc.get(b.name) || 0) + 1)));
  add('Badges', [['Badge', 'Faculty']].concat(Array.from(bc).sort((a, b) => b[1] - a[1])), [70, 10]);
  XLSX.writeFile(wb, 'College-PD-report-' + fileStamp() + '.xlsx');
}
/* ==========================================================================
   11. APP WIRING
   ========================================================================== */
function showNotes() {
  const n = S.notes; const el = $('#notes'); const body = $('#notes-body');
  const items = [];
  items.push(`${n.rowsRead} rows read: ${n.definitions} badge definitions without an issue, ${n.excluded} test rows removed, ${n.duplicates} repeat issuances merged.`);
  if (n.undated) items.push(`${n.undated} issuances have no readable date and are left off the timeline.`);
  if (n.unmappedBadges.size) items.push(`${n.unmappedBadges.size} earned badges have no theme in the draft map and do not count toward Bronze: ${Array.from(n.unmappedBadges.keys()).slice(0, 8).join('; ')}${n.unmappedBadges.size > 8 ? '...' : ''}.`);
  if (n.unmappedColleges.size) items.push(`${Array.from(n.unmappedColleges.values()).reduce((a, b) => a + b, 0)} faculty have a college value that could not be matched (shown as "${UNMAPPED_COLLEGE}"): ${Array.from(n.unmappedColleges.keys()).slice(0, 6).join('; ')}.`);
  items.push('Department, rank, learning hours, discipline points, and Silver/Gold evidence are not in the Moodle badge report.');
  body.innerHTML = '<ul>' + items.map(i => `<li>${esc(i)}</li>`).join('') + '</ul>';
  el.hidden = false;
}

function refreshAll() {
  if (!S.loaded) return;
  runAnalyses(); refreshFilters();
  $('#empty').hidden = true;
  showView(S.view, true);
}

function showView(v, noFocus) {
  S.view = v;
  ['individual', 'college'].forEach(id => {
    const on = id === v;
    const tab = $('#tab-' + id); tab.setAttribute('aria-selected', on); tab.tabIndex = on ? 0 : -1;
    $('#panel-' + id).hidden = !on || !S.loaded;
  });
  if (!S.loaded) { $('#empty').hidden = false; return; }
  if (v === 'individual') renderIndividual(); else renderCollege();
  try { history.replaceState(null, '', '#' + v); } catch (e) { /* ignore */ }
}

async function loadFile(file) {
  try {
    setStatus('Reading ' + file.name + '...');
    const raw = await readRowsFromFile(file);
    S.rows = normalizeRows(raw); S.fileName = file.name; 
    S.selected = null; S.year = 'all';
    buildModel(S.rows);
    if (!S.people.size) throw new Error('No badge issuances (rows with an email address) were found in this file.');
    refreshAll(); showNotes();
    setStatus(`Loaded ${file.name}: ${S.people.size} faculty, ${Array.from(S.people.values()).reduce((s, p) => s + p.badges.size, 0)} badge issuances.`);
  } catch (e) {
    console.error(e); setStatus('Could not load the file. ' + e.message, true);
  }
}

async function loadSidecar(file, kind) {
  try {
    const raw = await readRowsFromFile(file); if (!raw.length) throw new Error('The file is empty.');
    const lk = {}; Object.keys(raw[0]).forEach(k => { lk[k.trim().toLowerCase()] = k; });
    const col = (...names) => names.map(n => lk[n]).find(Boolean);
    if (kind === 'roster') {
      const e = col('email', 'email address'), d = col('department', 'dept'), r = col('rank', 'academic rank'), h = col('hire date', 'hire_date', 'hired');
      if (!e) throw new Error('Roster needs an "email" column.');
      S.roster = new Map();
      raw.forEach(row => { const em = String(row[e] || '').trim().toLowerCase(); if (em) S.roster.set(em, { dept: d ? String(row[d] || '').trim() : '', rank: r ? String(row[r] || '').trim() : '', hire: h ? String(row[h] || '').trim() : '' }); });
      if (S.loaded) { applyRoster(); }
      setStatus(`Roster loaded: ${S.roster.size} people.`);
    } else {
      const b = col('badge', 'name', 'badge name'), t = col('theme');
      if (!b || !t) throw new Error('Theme map needs "badge" and "theme" columns.');
      let n = 0;
      raw.forEach(row => {
        const key = canon(row[b]); let th = String(row[t] || '').trim(); if (!key || !th) return;
        let id = THEMES.find(x => x.id.toLowerCase() === th.toLowerCase() || x.name.toLowerCase() === th.toLowerCase() || x.short.toLowerCase() === th.toLowerCase());
        if (!id && /^[1-5]$/.test(th)) id = THEMES[+th - 1];
        if (id) { S.themeMap[key] = id.id; n++; }
      });
      if (S.loaded) { applyThemeMap(); showNotes(); }
      setStatus(`Theme map loaded: ${n} badges assigned.`);
    }
    if (S.loaded) refreshAll();
  } catch (e) { console.error(e); setStatus('Could not load the file. ' + e.message, true); }
}

function downloadThemeTemplate() {
  const rows = [['badge', 'theme']];
  const cats = S.loaded ? Array.from(S.catalog.values()).filter(c => c.count > 0) : Object.keys(DEFAULT_THEME_MAP).map(k => ({ name: k, key: k }));
  cats.sort((a, b) => a.name.localeCompare(b.name)).forEach(c => rows.push([c.name, S.themeMap[c.key] || '']));
  rows.push([], ['Themes: ' + THEMES.map(t => t.id + ' = ' + t.name).join('; ')]);
  const csv = rows.map(r => r.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(',')).join('\r\n');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' })); a.download = 'badge-theme-map-template.csv'; document.body.appendChild(a); a.click(); a.remove();
}

/* ---- theme (light / dark) ---- */
function effectiveDark() { const t = document.documentElement.dataset.theme; return t === 'dark' || (t === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches); }
function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  const dark = effectiveDark(); const b = $('#btn-theme');
  b.setAttribute('aria-pressed', dark); $('#theme-icon').textContent = dark ? '☀' : '☾';
}

function runAction(act) {
  try {
    if (act === 'print') window.print();
    else if (act === 'ind-pdf') exportIndividualPdf();
    else if (act === 'ind-xlsx') exportIndividualXlsx();
    else if (act === 'college-pdf') exportCollegePdf();
    else if (act === 'college-xlsx') exportCollegeXlsx();
      } catch (e) { console.error(e); setStatus('Export failed. ' + e.message, true); }
}

function init() {
  const stored = safeStore('au-pd-theme'); applyTheme(stored === 'light' || stored === 'dark' ? stored : 'auto');
  $('#btn-theme').addEventListener('click', () => { const next = effectiveDark() ? 'light' : 'dark'; safeStore('au-pd-theme', next); applyTheme(next); if (S.loaded) showView(S.view, true); });
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { applyTheme(document.documentElement.dataset.theme); if (S.loaded) showView(S.view, true); });

  // tabs
  const tabs = $$('[role="tab"]');
  tabs.forEach((t, i) => {
    t.addEventListener('click', () => showView(t.id.replace('tab-', '')));
    t.addEventListener('keydown', e => {
      let j = null; if (e.key === 'ArrowRight') j = (i + 1) % tabs.length; else if (e.key === 'ArrowLeft') j = (i + tabs.length - 1) % tabs.length; else if (e.key === 'Home') j = 0; else if (e.key === 'End') j = tabs.length - 1;
      if (j !== null) { e.preventDefault(); tabs[j].focus(); showView(tabs[j].id.replace('tab-', '')); }
    });
  });

  // files
  $('#file-badges').addEventListener('change', e => { if (e.target.files[0]) loadFile(e.target.files[0]); e.target.value = ''; });
  $('#file-roster').addEventListener('change', e => { if (e.target.files[0]) loadSidecar(e.target.files[0], 'roster'); e.target.value = ''; });
  $('#file-map').addEventListener('change', e => { if (e.target.files[0]) loadSidecar(e.target.files[0], 'map'); e.target.value = ''; });
  $('#btn-map-template').addEventListener('click', downloadThemeTemplate);
  const dz = $('#dropzone');
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('dragging'); }));
  ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('dragging'); }));
  document.addEventListener('drop', e => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f) loadFile(f); });

  // settings
  $('#btn-settings').addEventListener('click', e => { const s = $('#settings'); const open = s.hidden; s.hidden = !open; e.currentTarget.setAttribute('aria-expanded', open); });
  $('#set-cycle').addEventListener('change', e => { S.cycleStart = +e.target.value; if (S.loaded) { computeYears(); refreshAll(); } });
  $('#set-active').addEventListener('change', e => { S.activeDays = Math.max(7, Math.min(365, +e.target.value || 60)); e.target.value = S.activeDays; if (S.loaded) refreshAll(); });

  // filters
  $$('.year-select').forEach(sel => sel.addEventListener('change', e => { S.year = e.target.value; refreshAll(); }));
  $('#f-college').addEventListener('change', e => { S.filters.college = e.target.value; renderCollege(); });
  $('#f-status').addEventListener('change', e => { S.filters.status = e.target.value; renderCollege(); });

  // search
  $('#q').addEventListener('input', renderSearchResults);
  $('#q').addEventListener('keydown', e => { if (e.key === 'Enter') { const b = $('#results button'); if (b) { e.preventDefault(); selectPerson(b.dataset.email); } } });
  $('#results').addEventListener('click', e => { const b = e.target.closest('button[data-email]'); if (b) selectPerson(b.dataset.email); });

  // actions (event delegation)
  document.addEventListener('click', e => { const b = e.target.closest('[data-act]'); if (b) runAction(b.dataset.act); });

  const hash = location.hash.replace('#', '');
  S.view = ['individual', 'college'].includes(hash) ? hash : 'individual';
  showView(S.view, true);
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
