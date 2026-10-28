/* ==========================================================================
   aTalent Training Plan Planner for Docebo  (POC)
   --------------------------------------------------------------------------
   Breaks an annual training plan into programs and monthly execution
   milestones, links each milestone to real Docebo ILT sessions or courses,
   and tracks plan vs actual from live Docebo data.

   One Cloudflare Worker that serves:
     /            Planner app (embed in Docebo via the Iframe widget)
     /api/...     JSON API used by the app
   Plus a scheduled() handler: add a Cron Trigger (e.g. "0 18 * * *") to sync
   actuals from Docebo every night.

   Storage: Cloudflare D1, bound as DB. Tables are created automatically.

   Variables (Settings > Variables and Secrets):
     DEMO_MODE         "1" = built-in construction machinery catalog, no Docebo calls
     DOCEBO_DOMAIN     e.g. atalentdemo2.docebosaas.com
     DOCEBO_CLIENT_ID, DOCEBO_CLIENT_SECRET, DOCEBO_USERNAME, DOCEBO_PASSWORD
     SALT_SECRET       same value as the Docebo Iframe widget Salt Secret
     APP_SECRET        random string used to sign app sessions
     ALLOWED_USER_IDS  optional comma list of Docebo user ids allowed in
     ALLOW_OPEN        "1" = allow opening the app outside Docebo (demos only)
     SYNC_BUDGET       optional max Docebo calls per sync run (default 45)
   ========================================================================== */

import APP_HTML from './app.html';
// SheetJS Community Edition 0.18.5 (Apache-2.0), served from this Worker so Excel
// import/export works without a third-party CDN.
import XLSX_JS from './vendor/xlsx.full.min.txt';

const TOKEN_TTL_SEC = 12 * 3600;

/* ----------------------------- helpers --------------------------------- */

class HttpError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

const json = (data, status = 200, extra = {}) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', ...extra },
  });

const uid = () => crypto.randomUUID().replace(/-/g, '').slice(0, 16);
const now = () => Date.now();
const isDemo = (env) => String(env.DEMO_MODE || '') === '1';

function pick(obj, keys) {
  if (!obj) return undefined;
  for (const k of keys) {
    const v = k.split('.').reduce((o, p) => (o == null ? undefined : o[p]), obj);
    if (v !== undefined && v !== null && v !== '') return v;
  }
  return undefined;
}

function num(v, d = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : d;
}

function clampStr(v, max = 300) {
  return v == null ? '' : String(v).slice(0, max);
}

/** "2026-03-14 09:00:00" or ISO -> "2026-03-14" (date part only). */
function dateOnly(s) {
  if (!s) return null;
  const m = String(s).match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[1]}-${m[2]}-${m[3]}` : null;
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

/* ----------------------------- crypto ---------------------------------- */

async function sha256Hex(input) {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(input));
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function hexEquals(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

const b64u = (s) => btoa(unescape(encodeURIComponent(s))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const unb64u = (s) => decodeURIComponent(escape(atob(s.replace(/-/g, '+').replace(/_/g, '/'))));

function appSecret(env) {
  if (env.APP_SECRET) return env.APP_SECRET;
  if (isDemo(env)) return 'demo-only-secret';
  throw new HttpError(500, 'APP_SECRET is not set.');
}

async function hmacHex(secret, data) {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const sig = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(data));
  return [...new Uint8Array(sig)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

async function issueToken(env, user) {
  const body = b64u(JSON.stringify({ ...user, exp: Math.floor(now() / 1000) + TOKEN_TTL_SEC }));
  return `${body}.${await hmacHex(appSecret(env), body)}`;
}

async function readToken(env, req) {
  const h = req.headers.get('authorization') || '';
  const tok = h.startsWith('Bearer ') ? h.slice(7) : '';
  const [body, sig] = tok.split('.');
  if (!body || !sig) throw new HttpError(401, 'Not signed in. Reopen the page from Docebo.');
  if (!hexEquals(await hmacHex(appSecret(env), body), sig)) throw new HttpError(401, 'Invalid session.');
  const user = JSON.parse(unb64u(body));
  if (!user.exp || user.exp < now() / 1000) throw new HttpError(401, 'Session expired. Reload the page in Docebo.');
  return user;
}

/**
 * Docebo Iframe widget handoff (Advanced Settings enabled). Docebo appends
 * user_id, username, auth_code and hash to the iframe URL on every load:
 *   hash = SHA-256 of "user_id,username,auth_code,<Salt Secret>"
 */
async function verifyFrameHandoff(env, params) {
  const userId = params.get('user_id');
  const username = params.get('username');
  const authCode = params.get('auth_code');
  const received = (params.get('hash') || '').toLowerCase();
  if (!userId || !username) return { ok: false, reason: 'no-params' };
  if (!env.SALT_SECRET) return { ok: false, reason: 'no-salt-configured' };
  if (!received) return { ok: false, reason: 'no-hash' };
  const expected = await sha256Hex([userId, username, authCode || '', env.SALT_SECRET].join(','));
  if (!hexEquals(expected, received)) return { ok: false, reason: 'hash-mismatch' };
  return { ok: true, userId: String(userId), username };
}

function userAllowed(env, userId) {
  const list = String(env.ALLOWED_USER_IDS || '').split(',').map((s) => s.trim()).filter(Boolean);
  return !list.length || list.includes(String(userId));
}

/* ----------------------------- schema ---------------------------------- */

const SCHEMA = [
  `CREATE TABLE IF NOT EXISTS plans (
     id TEXT PRIMARY KEY, name TEXT, name_zh TEXT, year INTEGER, owner TEXT,
     created_at INTEGER, updated_at INTEGER)`,
  `CREATE TABLE IF NOT EXISTS programs (
     id TEXT PRIMARY KEY, plan_id TEXT, name TEXT, name_zh TEXT,
     objective TEXT, objective_zh TEXT, target_sessions INTEGER DEFAULT 0,
     target_trainees INTEGER DEFAULT 0, owner TEXT, color INTEGER DEFAULT 0,
     sort INTEGER DEFAULT 0, updated_at INTEGER)`,
  `CREATE TABLE IF NOT EXISTS milestones (
     id TEXT PRIMARY KEY, plan_id TEXT, program_id TEXT, month INTEGER,
     title TEXT, title_zh TEXT, planned_sessions INTEGER DEFAULT 0,
     planned_trainees INTEGER DEFAULT 0, region TEXT, delivery TEXT,
     notes TEXT, updated_at INTEGER)`,
  `CREATE TABLE IF NOT EXISTS links (
     id TEXT PRIMARY KEY, plan_id TEXT, milestone_id TEXT, item_type TEXT,
     course_id TEXT, session_id TEXT, name TEXT, course_name TEXT, course_type TEXT,
     act_sessions INTEGER, act_delivered INTEGER, act_enrolled INTEGER,
     act_trainees INTEGER, first_date TEXT, last_date TEXT,
     sync_note TEXT, synced_at INTEGER, created_at INTEGER)`,
  `CREATE INDEX IF NOT EXISTS ms_plan ON milestones (plan_id)`,
  `CREATE INDEX IF NOT EXISTS links_plan ON links (plan_id)`,
  `CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)`,
];

let schemaReady = false;
async function ensureSchema(env) {
  if (schemaReady) return;
  if (!env.DB) throw new HttpError(500, 'D1 database is not bound. Add a D1 binding named DB.');
  await env.DB.batch(SCHEMA.map((s) => env.DB.prepare(s)));
  // Columns added after the first release; harmless if they already exist.
  const added = [
    ['milestones', 'audience_type TEXT'], ['milestones', 'audience_id TEXT'], ['milestones', 'audience_name TEXT'],
    ['milestones', 'audience_size INTEGER'], ['links', 'created_in_docebo INTEGER DEFAULT 0'], ['links', 'enrol_note TEXT'],
  ];
  for (const [table, col] of added) {
    try { await env.DB.prepare(`ALTER TABLE ${table} ADD COLUMN ${col}`).run(); } catch { /* exists */ }
  }
  schemaReady = true;
}

const all = async (env, sql, ...args) => (await env.DB.prepare(sql).bind(...args).all()).results || [];
const first = async (env, sql, ...args) => env.DB.prepare(sql).bind(...args).first();
const run = async (env, sql, ...args) => env.DB.prepare(sql).bind(...args).run();

async function setMeta(env, k, v) {
  await run(env, 'INSERT INTO meta (k, v) VALUES (?, ?) ON CONFLICT(k) DO UPDATE SET v = excluded.v', k, JSON.stringify(v));
}
async function getMeta(env, k) {
  const r = await first(env, 'SELECT v FROM meta WHERE k = ?', k);
  return r ? JSON.parse(r.v) : null;
}

/* ----------------------------- Docebo client ---------------------------- */

let tokenCache = { key: '', token: '', exp: 0 };

function doceboBase(env) {
  const d = String(env.DOCEBO_DOMAIN || '').replace(/^https?:\/\//, '').replace(/\/+$/, '');
  if (!d) throw new HttpError(500, 'DOCEBO_DOMAIN is not set.');
  return `https://${d}`;
}

async function doceboToken(env) {
  const base = doceboBase(env);
  const key = `${base}|${env.DOCEBO_CLIENT_ID}|${env.DOCEBO_USERNAME}`;
  if (tokenCache.key === key && tokenCache.exp > now() + 60000) return tokenCache.token;
  for (const k of ['DOCEBO_CLIENT_ID', 'DOCEBO_CLIENT_SECRET', 'DOCEBO_USERNAME', 'DOCEBO_PASSWORD']) {
    if (!env[k]) throw new HttpError(500, `${k} is not set.`);
  }
  const body = new URLSearchParams({
    grant_type: 'password', client_id: env.DOCEBO_CLIENT_ID, client_secret: env.DOCEBO_CLIENT_SECRET,
    username: env.DOCEBO_USERNAME, password: env.DOCEBO_PASSWORD, scope: 'api',
  });
  const r = await fetch(`${base}/oauth2/token`, { method: 'POST', body, headers: { 'content-type': 'application/x-www-form-urlencoded' } });
  const data = await r.json().catch(() => ({}));
  if (!r.ok || !data.access_token) {
    throw new HttpError(502, `Docebo sign-in failed (${r.status}): ${data.error_description || data.error || data.message || 'no token'}. Check the OAuth2 app has the Resource Owner Password Credentials grant enabled.`);
  }
  tokenCache = { key, token: data.access_token, exp: now() + num(data.expires_in, 3600) * 1000 };
  return data.access_token;
}

/** A per-request counter so syncs stay inside Workers subrequest limits. */
function makeBudget(max) {
  return { max, used: 0, take() { if (this.used >= this.max) throw new BudgetError(); this.used++; } };
}
class BudgetError extends Error { constructor() { super('budget'); } }

async function doceboGet(env, path, budget) {
  return doceboCall(env, 'GET', path, null, budget);
}

async function doceboCall(env, method, path, body, budget) {
  if (budget) budget.take();
  if (isDemo(env)) return method === 'GET' ? demoGet(path) : demoWrite(method, path, body);
  const tok = await doceboToken(env);
  const init = { method, headers: { authorization: `Bearer ${tok}`, accept: 'application/json' } };
  if (body) { init.body = JSON.stringify(body); init.headers['content-type'] = 'application/json'; }
  const r = await fetch(doceboBase(env) + path, init);
  const text = await r.text();
  let data;
  try { data = JSON.parse(text); } catch { data = { raw: text.slice(0, 500) }; }
  if (!r.ok) {
    let msg = Array.isArray(data.message) ? data.message.map((m) => (typeof m === 'string' ? m : `${m.field || ''} ${m.message || ''}`.trim())).join('; ') : pick(data, ['message', 'error_description', 'error']);
    msg = msg || text.slice(0, 200);
    throw new HttpError(r.status === 401 ? 502 : r.status, `Docebo ${r.status} on ${path.split('?')[0]}: ${typeof msg === 'string' ? msg : JSON.stringify(msg)}`);
  }
  return data;
}

/** Collect data.items across pages (Docebo: page, page_size, has_more_data). */
async function doceboItems(env, path, budget, maxPages = 5, pageSize = 200) {
  const out = [];
  for (let page = 1; page <= maxPages; page++) {
    const sep = path.includes('?') ? '&' : '?';
    const data = await doceboGet(env, `${path}${sep}page=${page}&page_size=${pageSize}`, budget);
    const items = pick(data, ['data.items', 'items']) || [];
    out.push(...items);
    const more = pick(data, ['data.has_more_data', 'has_more_data']);
    if (!more || !items.length) break;
  }
  return out;
}

/* ----------------------------- Docebo reads ----------------------------- */

function normCourse(c) {
  return {
    id: String(pick(c, ['id_course', 'id', 'course_id'])),
    name: pick(c, ['name', 'course_name', 'title']) || 'Course',
    code: pick(c, ['code', 'course_code']) || '',
    type: String(pick(c, ['course_type', 'type']) || '').toLowerCase(),
  };
}

function normSession(s) {
  return {
    id: String(pick(s, ['id', 'id_session', 'session_id'])),
    name: pick(s, ['name', 'session_name']) || 'Session',
    start: dateOnly(pick(s, ['date_start', 'date_begin', 'start_date'])),
    end: dateOnly(pick(s, ['date_end', 'end_date'])),
  };
}

function normEvent(e) {
  return {
    id: String(pick(e, ['id', 'id_event', 'event_id'])),
    name: pick(e, ['name']) || 'Event',
    date: dateOnly(pick(e, ['schedule.date', 'day', 'date', 'date_begin', 'start_date'])),
    venue: pick(e, ['location.name', 'location_name']) || (pick(e, ['webinar_tool', 'webinar.tool']) ? 'Online' : null),
  };
}

async function searchCourses(env, q) {
  const qs = q ? `?search_text=${encodeURIComponent(q)}&sort_attr=name&sort_dir=asc` : '?sort_attr=name&sort_dir=asc';
  const raw = await doceboItems(env, `/learn/v1/courses${qs}`, null, 2, 100);
  return raw.map(normCourse).filter((c) => c.id && c.id !== 'undefined');
}

async function courseSessions(env, courseId) {
  const raw = await doceboItems(env, `/course/v1/courses/${encodeURIComponent(courseId)}/sessions`, null, 2, 100);
  return raw.map(normSession).filter((s) => s.id && s.id !== 'undefined');
}

const isPresent = (a) => {
  const s = pick(a, ['status_id', 'status', 'attendance', 'present']);
  return s === 1 || s === '1' || s === true || /present|attended/i.test(String(s));
};

/** Actuals for an ILT session link: 1 scheduled session, delivered when its last event is past. */
async function syncSessionLink(env, link, budget) {
  const events = (await doceboItems(env, `/course/v1/sessions/${encodeURIComponent(link.session_id)}/events`, budget, 1, 100)).map(normEvent);
  const dates = events.map((e) => e.date).filter(Boolean).sort();
  const today = todayStr();
  const firstDate = dates[0] || null;
  const lastDate = dates[dates.length - 1] || null;
  const enrolled = new Set();
  const present = new Set();
  // Class list and attendance for up to 3 events (keeps the run inside subrequest limits).
  for (const ev of events.slice(0, 3)) {
    const rows = await doceboItems(env, `/course/v1/events/${encodeURIComponent(ev.id)}/attendance`, budget, 1, 200);
    for (const a of rows) {
      const u = String(pick(a, ['user_id', 'id_user', 'id']));
      enrolled.add(u);
      if (ev.date && ev.date <= today && isPresent(a)) present.add(u);
    }
  }
  return {
    act_sessions: events.length ? 1 : 0,
    act_delivered: lastDate && lastDate < today ? 1 : 0,
    act_enrolled: enrolled.size,
    act_trainees: present.size,
    first_date: firstDate,
    last_date: lastDate,
    sync_note: events.length ? ([...new Set(events.map((e) => e.venue).filter(Boolean))].length ? `Venue: ${[...new Set(events.map((e) => e.venue).filter(Boolean))].join(', ')}` : null) : 'Session has no events yet',
  };
}

/** Actuals for a course link: completions dated in the milestone month count as trainees. */
async function syncCourseLink(env, link, ms, year, budget) {
  let rows;
  try {
    rows = await doceboItems(env, `/course/v1/courses/enrollments?course_id[]=${encodeURIComponent(link.course_id)}`, budget, 10, 200);
  } catch (e) {
    if (e instanceof BudgetError) throw e;
    rows = await doceboItems(env, `/learn/v1/enrollments?id_course=${encodeURIComponent(link.course_id)}`, budget, 10, 200);
  }
  const prefix = `${year}-${String(ms.month).padStart(2, '0')}`;
  let completedInMonth = 0, completedAll = 0, firstDate = null, lastDate = null;
  for (const r of rows) {
    const status = String(pick(r, ['enrollment_status', 'status', 'course_status']) || '').toLowerCase();
    const done = dateOnly(pick(r, ['enrollment_completion_date', 'course_complete_date', 'date_complete', 'completion_date', 'completed_at']));
    if (status === 'completed' || status === '2' || done) {
      completedAll++;
      if (done && done.startsWith(prefix)) completedInMonth++;
      if (done) {
        if (!firstDate || done < firstDate) firstDate = done;
        if (!lastDate || done > lastDate) lastDate = done;
      }
    }
  }
  return {
    act_sessions: 0, act_delivered: 0,
    act_enrolled: rows.length,
    act_trainees: completedInMonth,
    first_date: firstDate, last_date: lastDate,
    sync_note: `${completedAll} completed in total, ${completedInMonth} in ${prefix}`,
  };
}

/** Sync the stalest links first, stopping cleanly when the call budget runs out. */
async function syncActuals(env, planId, milestoneId) {
  await ensureSchema(env);
  const plan = await first(env, 'SELECT * FROM plans WHERE id = ?', planId);
  if (!plan) throw new HttpError(404, 'Plan not found.');
  const links = milestoneId
    ? await all(env, 'SELECT * FROM links WHERE plan_id = ? AND milestone_id = ? ORDER BY COALESCE(synced_at, 0)', planId, milestoneId)
    : await all(env, 'SELECT * FROM links WHERE plan_id = ? ORDER BY COALESCE(synced_at, 0)', planId);
  const msMap = new Map((await all(env, 'SELECT * FROM milestones WHERE plan_id = ?', planId)).map((m) => [m.id, m]));
  const budget = makeBudget(isDemo(env) ? 1e6 : num(env.SYNC_BUDGET, 45));
  let done = 0, failed = 0;
  const errors = [];
  for (const link of links) {
    const ms = msMap.get(link.milestone_id);
    if (!ms) continue;
    try {
      const a = link.item_type === 'session'
        ? await syncSessionLink(env, link, budget)
        : await syncCourseLink(env, link, ms, plan.year, budget);
      await run(env, `UPDATE links SET act_sessions=?, act_delivered=?, act_enrolled=?, act_trainees=?,
        first_date=?, last_date=?, sync_note=?, synced_at=? WHERE id=?`,
        a.act_sessions, a.act_delivered, a.act_enrolled, a.act_trainees, a.first_date, a.last_date, a.sync_note, now(), link.id);
      done++;
    } catch (e) {
      if (e instanceof BudgetError) break;
      failed++;
      errors.push(`${link.name}: ${e.message}`);
      await run(env, 'UPDATE links SET sync_note=?, synced_at=? WHERE id=?', `Sync error: ${e.message}`.slice(0, 300), now(), link.id);
    }
  }
  const pending = links.length - done - failed;
  const result = { at: now(), synced: done, failed, pending, total: links.length, calls: budget.used, errors: errors.slice(0, 5) };
  if (!milestoneId) await setMeta(env, `sync:${planId}`, result);
  return result;
}

/* ----------------------------- audiences -------------------------------- */
// "To whom": a Docebo audience (formerly group) or a branch.
//   Audiences: GET /audiences/v1/audience, members GET /audiences/v1/audience/{uuid}/users
//   Branches:  GET /manage/v1/orgchart,   members GET /manage/v1/orgchart/{id}/users

async function searchAudiences(env, q) {
  const qs = q ? `?search_text=${encodeURIComponent(q)}` : '';
  const out = [];
  const errors = [];
  try {
    const aud = await doceboItems(env, `/audiences/v1/audience${qs}`, null, 1, 50);
    for (const a of aud) out.push({ type: 'audience', id: String(pick(a, ['uuid', 'id'])), name: pick(a, ['name']) || 'Audience', note: pick(a, ['type']) || '' });
  } catch (e) { errors.push(e.message); }
  try {
    const br = await doceboItems(env, `/manage/v1/orgchart${qs ? qs + '&' : '?'}flattened=true`, null, 1, 50);
    for (const b of br) {
      const id = pick(b, ['id', 'id_org', 'branch_id']);
      if (id == null) continue;
      out.push({ type: 'branch', id: String(id), name: pick(b, ['title', 'name', 'code']) || 'Branch', note: pick(b, ['code']) || '' });
    }
  } catch (e) { errors.push(e.message); }
  if (!out.length && errors.length) throw new HttpError(502, errors[0]);
  return out;
}

/** Members of an audience or branch, normalised to { user_id, username, name }. */
async function audienceMembers(env, type, id, budget) {
  const rows = type === 'branch'
    ? await doceboItems(env, `/manage/v1/orgchart/${encodeURIComponent(id)}/users`, budget, 10, 200)
    : await doceboItems(env, `/audiences/v1/audience/${encodeURIComponent(id)}/users`, budget, 10, 200);
  const seen = new Set();
  const out = [];
  for (const r of rows) {
    const uid = String(pick(r, ['user_id', 'idst', 'id']) ?? '');
    if (!uid || seen.has(uid)) continue;
    seen.add(uid);
    const username = String(pick(r, ['username', 'userid']) || '').replace(/^\//, '');
    const name = `${pick(r, ['first_name', 'firstname', 'firs_name']) || ''} ${pick(r, ['last_name', 'lastname']) || ''}`.trim();
    out.push({ user_id: uid, username, name: name || username });
  }
  return out;
}

/* ----------------------------- venues ----------------------------------- */
// GET /learn/v1/location (id, name, address, country); classrooms per location
// at GET /learn/v1/location/{id}/classroom (id, name, seats).

async function listLocations(env) {
  const raw = await doceboItems(env, '/learn/v1/location', null, 3, 100);
  return raw.map((l) => ({
    id: String(pick(l, ['id', 'id_location', 'location_id'])),
    name: pick(l, ['name']) || 'Location',
    country: pick(l, ['country']) || '',
    address: pick(l, ['address']) || '',
  })).filter((l) => l.id && l.id !== 'undefined');
}

async function listClassrooms(env, locationId) {
  const raw = await doceboItems(env, `/learn/v1/location/${encodeURIComponent(locationId)}/classroom`, null, 1, 100);
  return raw.map((c) => ({ id: String(pick(c, ['id', 'id_classroom'])), name: pick(c, ['name']) || 'Classroom', seats: num(pick(c, ['seats']), 0) }));
}

/* ----------------------------- create in Docebo ------------------------- */

function cleanTime(t, d) {
  const m = /^(\d{1,2}):(\d{2})/.exec(String(t || ''));
  return m ? `${m[1].padStart(2, '0')}:${m[2]}:00` : d;
}

/**
 * Create an ILT course (optional), then one session + one event per row, and
 * link each new session to the milestone. Docebo requires a venue or a
 * webinar tool on every event: a row with location_id (and optional
 * classroom_id) gets that venue, otherwise the "custom" webinar tool with
 * the meeting link given.
 */
async function createSessions(env, ms, plan, body) {
  const rows = Array.isArray(body.sessions) ? body.sessions.slice(0, 12) : [];
  if (!rows.length) throw new HttpError(400, 'Add at least one session.');
  const meetingUrl = String(body.meeting_url || env.DEFAULT_MEETING_URL || '').trim();
  for (const r of rows) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(r.date || ''))) throw new HttpError(400, 'Every session needs a date.');
    // Docebo requires a venue or a webinar tool on every event.
    if (!r.location_id && !/^https?:\/\//i.test(meetingUrl)) throw new HttpError(400, 'Each session needs a venue, or a meeting link (https://...) for online sessions.');
  }

  let courseId = body.course_id ? String(body.course_id) : null;
  let courseName = clampStr(body.course_name || '');
  const result = { course_id: null, course_created: false, created: [], failed: [] };

  if (!courseId) {
    const name = clampStr(body.new_course_name || ms.title || 'Training', 200).replace(/[<>]/g, '');
    const created = await doceboCall(env, 'POST', '/learn/v1/courses', {
      name, code: clampStr(body.new_course_code || '', 50) || undefined,
      description: `<p>${name}</p><p>Created by the aTalent Training Plan Planner.</p>`,
      course_type: 'classroom', type: 'classroom', status: 'published', language: 'english',
    });
    courseId = String(pick(created, ['data.course_id', 'data.id', 'course_id']) || '');
    if (!courseId) throw new HttpError(502, 'Docebo did not return the new course ID.');
    courseName = name;
    result.course_created = true;
  }
  result.course_id = courseId;

  for (const [i, r] of rows.entries()) {
    const sessName = clampStr(r.name || `${courseName || ms.title} | ${r.date}`, 200);
    let sid = null;
    try {
      const sess = await doceboCall(env, 'POST', `/course/v1/courses/${encodeURIComponent(courseId)}/sessions`, {
        name: sessName, description: `Planned under: ${ms.title} (${plan.year}-${String(ms.month).padStart(2, '0')})`,
        min_enroll: 0, max_enroll: Math.max(1, Math.round(num(r.max_enroll, 20))), evaluation_type: 2,
      });
      sid = String(pick(sess, ['data.id_session', 'data.session_id', 'data.id']) || '');
      if (!sid) throw new Error('Docebo did not return the new session ID.');
      await doceboCall(env, 'POST', `/course/v1/sessions/${encodeURIComponent(sid)}/events`, {
        name: sessName, day: r.date,
        time_begin: cleanTime(r.time_begin, '09:00:00'), time_end: cleanTime(r.time_end, '17:00:00'),
        timezone: clampStr(r.timezone || body.timezone || 'Asia/Shanghai', 60),
        ...(r.location_id
          ? { id_location: Number(r.location_id), ...(r.classroom_id ? { id_classroom: Number(r.classroom_id) } : {}) }
          : { webinar_tool: 'custom', custom_url: meetingUrl }),
      });
      const linkId = uid();
      await insertRow(env, 'links', {
        id: linkId, plan_id: ms.plan_id, milestone_id: ms.id, item_type: 'session', course_id: courseId, session_id: sid,
        name: sessName, course_name: courseName, course_type: 'classroom', created_in_docebo: 1, created_at: now(),
      });
      result.created.push({ session_id: sid, name: sessName, date: r.date });
    } catch (e) {
      result.failed.push({ row: i + 1, name: sessName, session_id: sid, error: e.message });
    }
  }
  return result;
}

/* ----------------------------- enrol audience --------------------------- */

/**
 * Who goes where: audience members are split evenly across the milestone's
 * linked ILT sessions (in date order); a course link gets everyone.
 */
async function enrolmentPlan(env, ms) {
  if (!ms.audience_type || !ms.audience_id) throw new HttpError(400, 'Choose an audience for this milestone first.');
  const links = await all(env, 'SELECT * FROM links WHERE milestone_id = ? ORDER BY COALESCE(first_date, \'9999\'), created_at', ms.id);
  if (!links.length) throw new HttpError(400, 'Link or create the Docebo sessions for this milestone first.');
  const members = await audienceMembers(env, ms.audience_type, ms.audience_id);
  const sessions = links.filter((l) => l.item_type === 'session');
  const courses = links.filter((l) => l.item_type !== 'session');
  const targets = [];
  if (sessions.length) {
    const per = Math.ceil(members.length / sessions.length);
    sessions.forEach((l, i) => targets.push({ link: l, users: members.slice(i * per, (i + 1) * per) }));
  }
  for (const l of courses) targets.push({ link: l, users: members });
  return { members, targets };
}

async function enrolAudience(env, ms, execute) {
  const { members, targets } = await enrolmentPlan(env, ms);
  const summary = {
    audience: ms.audience_name, members: members.length,
    targets: targets.map((t) => ({ link_id: t.link.id, name: t.link.name, type: t.link.item_type, count: t.users.length, sample: t.users.slice(0, 5).map((u) => u.name) })),
  };
  if (!execute) return { preview: true, ...summary };

  let enrolled = 0, already = 0;
  const failed = [];
  for (const t of targets) {
    for (let i = 0; i < t.users.length; i += 300) {
      const chunk = t.users.slice(i, i + 300);
      const items = chunk.map((u, k) => {
        const it = { course_id: Number(t.link.course_id), user_id: String(u.user_id), level: 'student', batch_item_id: `${i + k}` };
        if (t.link.item_type === 'session') it.session_id = Number(t.link.session_id);
        return it;
      });
      const r = await doceboCall(env, 'POST', '/learn/v1/enrollment/batch', {
        items, options: { status: 'subscribed', level: 'student', enrollment_date_now: true, enable_waitinglist: false, trigger_notifications: true },
        trigger_gamification: false,
      });
      const rows = Array.isArray(pick(r, ['data'])) ? r.data : [];
      if (!rows.length) { enrolled += chunk.length; continue; }
      for (const row of rows) {
        if (row.success !== false) { enrolled++; continue; }
        const msg = String(row.message || row.error_code || row.error || 'unknown error');
        if (/already|exist|duplicate/i.test(msg)) already++;
        else failed.push(`${chunk[Number(row.batch_item_id) - i]?.name || 'user'}: ${msg}`);
      }
    }
    await run(env, 'UPDATE links SET enrol_note = ? WHERE id = ?', `Enrolled ${t.users.length} from ${ms.audience_name} on ${todayStr()}`, t.link.id);
  }
  return { preview: false, ...summary, enrolled, already, failed: failed.slice(0, 10), failed_count: failed.length };
}

/* ----------------------------- plan data -------------------------------- */

async function planState(env, planId) {
  await ensureSchema(env);
  const plans = await all(env, 'SELECT * FROM plans ORDER BY year DESC, created_at DESC');
  const plan = plans.find((p) => p.id === planId) || plans[0] || null;
  if (!plan) return { plans: [], plan: null, programs: [], milestones: [], links: [], lastSync: null };
  const [programs, milestones, links, lastSync] = await Promise.all([
    all(env, 'SELECT * FROM programs WHERE plan_id = ? ORDER BY sort, name', plan.id),
    all(env, 'SELECT * FROM milestones WHERE plan_id = ? ORDER BY month, title', plan.id),
    all(env, 'SELECT * FROM links WHERE plan_id = ? ORDER BY created_at', plan.id),
    getMeta(env, `sync:${plan.id}`),
  ]);
  return { plans: plans.map((p) => ({ id: p.id, name: p.name, name_zh: p.name_zh, year: p.year })), plan, programs, milestones, links, lastSync };
}

const PROGRAM_FIELDS = ['name', 'name_zh', 'objective', 'objective_zh', 'target_sessions', 'target_trainees', 'owner', 'color', 'sort'];
const MILESTONE_FIELDS = ['program_id', 'month', 'title', 'title_zh', 'planned_sessions', 'planned_trainees', 'region', 'delivery', 'notes', 'audience_type', 'audience_id', 'audience_name'];
const INT_FIELDS = new Set(['target_sessions', 'target_trainees', 'color', 'sort', 'month', 'planned_sessions', 'planned_trainees', 'year']);

function cleanFields(body, fields) {
  const out = {};
  for (const f of fields) {
    if (!(f in body)) continue;
    out[f] = INT_FIELDS.has(f) ? Math.max(0, Math.round(num(body[f]))) : clampStr(body[f], f === 'notes' || f.startsWith('objective') ? 1000 : 200);
  }
  if ('month' in out && (out.month < 1 || out.month > 12)) throw new HttpError(400, 'month must be 1 to 12.');
  return out;
}

/** When a milestone's audience is set, store its current member count. */
async function audienceSize(env, b) {
  if (!('audience_id' in b)) return {};
  if (!b.audience_id) return { audience_type: '', audience_id: '', audience_name: '', audience_size: null };
  const type = b.audience_type === 'branch' ? 'branch' : 'audience';
  try { return { audience_type: type, audience_size: (await audienceMembers(env, type, String(b.audience_id))).length }; }
  catch { return { audience_type: type, audience_size: null }; }
}

async function insertRow(env, table, row) {
  const keys = Object.keys(row);
  await run(env, `INSERT INTO ${table} (${keys.join(',')}) VALUES (${keys.map(() => '?').join(',')})`, ...keys.map((k) => row[k]));
}

async function updateRow(env, table, id, row) {
  const keys = Object.keys(row);
  if (!keys.length) return;
  await run(env, `UPDATE ${table} SET ${keys.map((k) => `${k}=?`).join(',')} WHERE id=?`, ...keys.map((k) => row[k]), id);
}

/* ----------------------------- seed plan -------------------------------- */

// Construction machinery sample plan. Programs carry annual objectives;
// each milestone is one month of execution for one program.
const SEED = {
  name: 'Global Service & Dealer Training Plan', name_zh: '全球服务与经销商培训计划',
  programs: [
    {
      key: 'launch', color: 0, target_sessions: 12, target_trainees: 240,
      name: 'New Model Launch: EX750 Hydraulic Excavator', name_zh: '新机型上市：EX750液压挖掘机',
      objective: 'Every dealer sales and service team certified on the EX750 before machines reach their yard.',
      objective_zh: '在整机到达经销商之前，所有经销商销售与服务团队完成EX750认证。',
      months: [[2, 2, 40, 'China', 'Classroom'], [3, 3, 60, 'China', 'Hands-on workshop'], [4, 3, 60, 'Southeast Asia', 'Hands-on workshop'], [5, 2, 40, 'Middle East', 'Hands-on workshop'], [6, 2, 40, 'Africa', 'Virtual classroom']],
      title: 'EX750 launch training', title_zh: 'EX750上市培训',
    },
    {
      key: 'tech', color: 1, target_sessions: 24, target_trainees: 360,
      name: 'Service Technician Certification (L1 to L3)', name_zh: '服务技师认证（L1至L3）',
      objective: 'Grow certified technicians across the dealer network to cut first-time-fix failures.',
      objective_zh: '提升经销商网络认证技师数量，降低一次修复失败率。',
      months: [[1, 2, 30, 'China'], [2, 2, 30, 'China'], [3, 2, 30, 'Southeast Asia'], [4, 1, 15, 'China'], [5, 1, 15, 'Middle East'], [6, 2, 30, 'Southeast Asia'], [7, 3, 45, 'China'], [8, 3, 45, 'Latin America'], [9, 2, 30, 'Africa'], [10, 2, 30, 'China'], [11, 2, 30, 'Southeast Asia'], [12, 2, 30, 'China']],
      delivery: 'Hands-on workshop',
      title: 'Hydraulics, engine and electrical diagnostics workshop', title_zh: '液压、发动机与电气诊断实操工作坊',
    },
    {
      key: 'operator', color: 2, target_sessions: 20, target_trainees: 400,
      name: 'Operator Safety Certification (Customers & Rental Partners)', name_zh: '操作手安全认证（客户与租赁伙伴）',
      objective: 'Certify customer operators on safe operation; lighter load in peak construction season.',
      objective_zh: '为客户操作手提供安全操作认证；施工旺季减少排期。',
      months: [[1, 3, 60, 'China'], [2, 3, 60, 'China'], [3, 1, 20, 'China'], [6, 2, 40, 'Southeast Asia'], [7, 3, 60, 'Middle East'], [8, 3, 60, 'Africa'], [11, 2, 40, 'Latin America'], [12, 3, 60, 'China']],
      delivery: 'Classroom + machine practice',
      title: 'Operator safety certification', title_zh: '操作手安全认证',
    },
    {
      key: 'dealer', color: 3, target_sessions: 6, target_trainees: 150,
      name: 'Overseas Dealer Onboarding', name_zh: '海外经销商入职培训',
      objective: 'Onboard new distributors in SEA, Middle East, Africa and Latin America within 60 days of signing.',
      objective_zh: '新签约的东南亚、中东、非洲及拉美经销商在60天内完成入职培训。',
      months: [[3, 1, 25, 'Southeast Asia'], [5, 1, 25, 'Middle East'], [7, 1, 25, 'Africa'], [9, 1, 25, 'Latin America'], [10, 1, 25, 'Southeast Asia'], [11, 1, 25, 'Middle East']],
      delivery: 'Virtual classroom',
      title: 'Dealer onboarding cohort', title_zh: '经销商入职培训班',
    },
    {
      key: 'ehs', color: 4, target_sessions: 0, target_trainees: 1200,
      name: 'EHS & Workshop Safety Refresher', name_zh: 'EHS与车间安全复训',
      objective: 'All plant and workshop staff complete the quarterly EHS refresher.',
      objective_zh: '所有工厂与车间员工完成季度EHS复训。',
      months: [[3, 0, 300, 'China'], [6, 0, 300, 'China'], [9, 0, 300, 'China'], [12, 0, 300, 'China']],
      delivery: 'E-learning',
      title: 'Quarterly EHS refresher', title_zh: '季度EHS复训',
    },
  ],
};

async function seedPlan(env, user, year, withDemoLinks) {
  await ensureSchema(env);
  const planId = uid();
  const t = now();
  const stmts = [];
  const add = (table, row) => {
    const keys = Object.keys(row);
    stmts.push(env.DB.prepare(`INSERT INTO ${table} (${keys.join(',')}) VALUES (${keys.map(() => '?').join(',')})`).bind(...keys.map((k) => row[k])));
  };
  add('plans', { id: planId, name: `${SEED.name} ${year}`, name_zh: `${year}年${SEED.name_zh}`, year, owner: user.username || '', created_at: t, updated_at: t });
  const catalog = withDemoLinks ? demoCatalog() : null;
  SEED.programs.forEach((p, i) => {
    const programId = uid();
    add('programs', {
      id: programId, plan_id: planId, name: p.name, name_zh: p.name_zh, objective: p.objective, objective_zh: p.objective_zh,
      target_sessions: p.target_sessions, target_trainees: p.target_trainees, owner: '', color: p.color, sort: i, updated_at: t,
    });
    for (const [month, sessions, trainees, region, delivery] of p.months) {
      const msId = uid();
      add('milestones', {
        id: msId, plan_id: planId, program_id: programId, month, title: p.title, title_zh: p.title_zh,
        planned_sessions: sessions, planned_trainees: trainees, region, delivery: delivery || p.delivery || '', notes: '', updated_at: t,
      });
      if (catalog) {
        for (const l of catalog.linksFor(p.key, month, year)) {
          add('links', { id: uid(), plan_id: planId, milestone_id: msId, created_at: t, ...l });
        }
      }
    }
  });
  await env.DB.batch(stmts);
  return planId;
}

/* ----------------------------- demo catalog ----------------------------- */
// Deterministic fake Docebo tenant for DEMO_MODE. Answers the same endpoints
// the live code calls, so the sync path is identical in both modes.

let demoCache = null;
function demoCatalog() {
  if (demoCache) return demoCache;
  const year = new Date().getUTCFullYear();
  const today = todayStr();
  const courses = [];
  const sessions = new Map();   // courseId -> [session]
  const events = new Map();     // sessionId -> [event]
  const attendance = new Map(); // eventId -> [row]
  const enrollments = new Map(); // courseId -> [row]
  const plan = new Map();       // `${key}|${month}` -> [link rows]
  let seq = 1000;
  let rnd = 7;
  const rand = () => { rnd = (rnd * 16807) % 2147483647; return rnd / 2147483647; };
  const pad = (n) => String(n).padStart(2, '0');

  const ilt = [
    { key: 'launch', name: 'EX750 Excavator Launch Workshop', code: 'LAUNCH-EX750' },
    { key: 'tech', name: 'Technician Certification: Hydraulics & Diagnostics', code: 'TECH-CERT' },
    { key: 'operator', name: 'Operator Safety Certification', code: 'OPS-SAFE' },
    { key: 'dealer', name: 'Overseas Dealer Onboarding Cohort', code: 'DLR-ONB' },
  ];
  const elearn = [
    { key: 'ehs', name: 'EHS & Workshop Safety Refresher', code: 'EHS-Q' },
    { key: null, name: 'EX750 Product Walkaround (video)', code: 'EX750-VID' },
    { key: null, name: 'Parts Catalogue Navigation', code: 'PARTS-101' },
    { key: null, name: 'Dealer Warranty Claims Process', code: 'WARR-201' },
  ];
  const courseIds = {};
  for (const c of ilt) { const id = String(++seq); courseIds[c.key] = id; courses.push({ id_course: id, name: c.name, code: c.code, course_type: 'classroom' }); sessions.set(id, []); }
  for (const c of elearn) { const id = String(++seq); if (c.key) courseIds[c.key] = id; courses.push({ id_course: id, name: c.name, code: c.code, course_type: 'elearning' }); enrollments.set(id, []); }

  // Shape of reality vs plan, so the demo shows every status.
  const shortfall = { 'tech|4': 0, 'tech|11': 1, 'operator|7': 1, 'launch|5': 1 }; // sessions dropped
  const unlinked = new Set(['tech|9', 'operator|11', 'dealer|11']);
  const misaligned = new Set(['dealer|9']); // session actually runs the month after

  for (const p of SEED.programs) {
    for (const [month, plannedSessions, plannedTrainees] of p.months) {
      const k = `${p.key}|${month}`;
      if (unlinked.has(k)) continue;
      const rows = [];
      if (p.key === 'ehs') {
        const cid = courseIds.ehs;
        rows.push({ item_type: 'course', course_id: cid, session_id: null, name: 'EHS & Workshop Safety Refresher', course_name: 'EHS & Workshop Safety Refresher', course_type: 'elearning' });
        const done = month <= 9 ? Math.round(plannedTrainees * (0.82 + rand() * 0.2)) : 0;
        for (let i = 0; i < done; i++) {
          const d = Math.min(28, 1 + Math.floor(rand() * 28));
          const dt = `${year}-${pad(month)}-${pad(d)}`;
          enrollments.get(cid).push({ user_id: ++seq, enrollment_status: dt <= today ? 'completed' : 'in_progress', enrollment_completion_date: dt <= today ? `${dt} 10:00:00` : null });
        }
        plan.set(k, rows);
        continue;
      }
      const cid = courseIds[p.key];
      const n = k in shortfall ? shortfall[k] : plannedSessions;
      const per = plannedSessions ? plannedTrainees / plannedSessions : 15;
      for (let s = 0; s < n; s++) {
        const sid = String(++seq);
        const m = misaligned.has(k) ? month + 1 : month;
        const day = Math.min(26, 3 + s * 7 + Math.floor(rand() * 3));
        const start = `${year}-${pad(m)}-${pad(day)}`;
        const label = `${p.months.find((x) => x[0] === month)?.[3] || ''} ${start}`.trim();
        sessions.get(cid).push({ id: sid, name: `${ilt.find((c) => c.key === p.key).code} | ${label}`, date_start: `${start} 09:00:00`, date_end: `${year}-${pad(m)}-${pad(day + 1)} 17:00:00` });
        const evs = [0, 1].map((o) => ({ id: String(++seq), name: `Day ${o + 1}`, schedule: { date: `${year}-${pad(m)}-${pad(day + o)}`, time_begin: '09:00:00', time_end: '17:00:00' } }));
        events.set(sid, evs);
        const seats = Math.max(4, Math.round(per * (0.85 + rand() * 0.25)));
        const showRate = 0.8 + rand() * 0.18;
        for (const ev of evs) {
          const rowsA = [];
          for (let u = 0; u < seats; u++) {
            const past = ev.schedule.date <= today;
            rowsA.push({ user_id: 50000 + s * 100 + u + month * 1000, status_id: past && rand() < showRate ? 1 : (past ? 0 : null), time: 0 });
          }
          attendance.set(ev.id, rowsA);
        }
        rows.push({ item_type: 'session', course_id: cid, session_id: sid, name: sessions.get(cid).at(-1).name, course_name: ilt.find((c) => c.key === p.key).name, course_type: 'classroom' });
      }
      plan.set(k, rows);
    }
  }

  demoCache = {
    courses, sessions, events, attendance, enrollments,
    linksFor: (key, month) => (plan.get(`${key}|${month}`) || []).map((r) => ({ ...r })),
  };
  return demoCache;
}

const DEMO_AUDIENCES = [
  { type: 'audience', id: 'aud-sea-tech', name: 'SEA Dealer Technicians', size: 36 },
  { type: 'audience', id: 'aud-cn-svc', name: 'China Service Engineers', size: 48 },
  { type: 'audience', id: 'aud-me-sales', name: 'Middle East Dealer Sales', size: 22 },
  { type: 'audience', id: 'aud-rental-ops', name: 'Rental Partner Operators', size: 60 },
  { type: 'branch', id: '201', name: 'Dealers / Southeast Asia', size: 54 },
  { type: 'branch', id: '202', name: 'Dealers / Africa', size: 31 },
  { type: 'branch', id: '203', name: 'Plant / EHS Staff', size: 300 },
];
const DEMO_LOCATIONS = [
  { id: 11, name: 'Shanghai Training Centre', country: 'China', address: 'Lingang, Shanghai' },
  { id: 12, name: 'Jakarta Dealer Academy', country: 'Indonesia', address: 'Cikarang, Bekasi' },
  { id: 13, name: 'Dubai Service Hub', country: 'United Arab Emirates', address: 'Jebel Ali Free Zone' },
  { id: 14, name: 'Johannesburg Workshop', country: 'South Africa', address: 'Midrand' },
  { id: 15, name: 'Sao Paulo Parts Centre', country: 'Brazil', address: 'Guarulhos' },
];
const DEMO_CLASSROOMS = { 11: [{ id: 111, name: 'Hydraulics Lab', seats: 16 }, { id: 112, name: 'Machine Yard', seats: 24 }], 12: [{ id: 121, name: 'Workshop Bay 1', seats: 20 }] };
const DEMO_NAMES = ['Arif Rahman', 'Siti Aminah', 'Nguyen Van An', 'Tran Thi Mai', 'Somchai P.', 'Maria Santos', 'Li Wei', 'Zhang Hua', 'Wang Fang', 'Ahmed Hassan', 'Fatima Noor', 'Kwame Mensah', 'Amara Okafor', 'Carlos Rivera', 'Ana Souza'];
function demoMembers(id) {
  const a = DEMO_AUDIENCES.find((x) => x.id === String(id));
  const base = (DEMO_AUDIENCES.indexOf(a) + 1) * 10000;
  return Array.from({ length: a ? a.size : 0 }, (_, i) => ({ user_id: 80000 + base + i, username: `demo.user${base + i}`, first_name: DEMO_NAMES[i % DEMO_NAMES.length].split(' ')[0], last_name: DEMO_NAMES[i % DEMO_NAMES.length].split(' ').slice(1).join(' ') + ` ${i + 1}` }));
}

function demoWrite(method, path, body) {
  const cat = demoCatalog();
  const u = new URL(path, 'https://demo.local');
  let m;
  const id = () => String(900000 + Math.floor(Math.random() * 99999));
  if (method === 'POST' && u.pathname === '/learn/v1/courses') {
    const cid = id();
    cat.courses.push({ id_course: cid, name: body.name, code: body.code || '', course_type: 'classroom' });
    cat.sessions.set(cid, []);
    return { data: { course_id: Number(cid), success: true } };
  }
  if (method === 'POST' && (m = u.pathname.match(/^\/course\/v1\/courses\/(\d+)\/sessions$/))) {
    if (!cat.sessions.has(m[1])) throw new HttpError(404, 'Course not found');
    const sid = id();
    cat.sessions.get(m[1]).push({ id: sid, name: body.name });
    cat.events.set(sid, []);
    return { data: { id_session: Number(sid) } };
  }
  if (method === 'POST' && (m = u.pathname.match(/^\/course\/v1\/sessions\/(\d+)\/events$/))) {
    const eid = id();
    const evs = cat.events.get(m[1]) || [];
    if (!body.id_location && !(body.webinar_tool && body.custom_url)) throw new HttpError(400, 'The event must have a venue or a webinar tool');
    const loc = DEMO_LOCATIONS.find((l) => l.id === body.id_location);
    evs.push({ id: eid, name: body.name, schedule: { date: body.day, time_begin: body.time_begin, time_end: body.time_end }, location: loc ? { name: loc.name } : null, webinar_tool: body.webinar_tool || null });
    cat.events.set(m[1], evs);
    cat.attendance.set(eid, []);
    for (const s of cat.sessions.values()) { const ss = s.find((x) => x.id === m[1]); if (ss && !ss.date_start) ss.date_start = `${body.day} ${body.time_begin}`; }
    return { data: { event_id: Number(eid) } };
  }
  if (method === 'POST' && u.pathname === '/learn/v1/enrollment/batch') {
    const out = [];
    for (const it of body.items || []) {
      if (it.session_id) {
        const evs = cat.events.get(String(it.session_id)) || [];
        let dup = false;
        for (const ev of evs) {
          const list = cat.attendance.get(ev.id) || [];
          if (list.some((a) => String(a.user_id) === String(it.user_id))) dup = true;
          else list.push({ user_id: Number(it.user_id), status_id: null, time: 0 });
          cat.attendance.set(ev.id, list);
        }
        out.push(dup ? { batch_item_id: it.batch_item_id, success: false, message: 'User already enrolled' } : { batch_item_id: it.batch_item_id, success: true });
      } else {
        const list = cat.enrollments.get(String(it.course_id)) || [];
        if (list.some((a) => String(a.user_id) === String(it.user_id))) { out.push({ batch_item_id: it.batch_item_id, success: false, message: 'User already enrolled' }); continue; }
        list.push({ user_id: Number(it.user_id), enrollment_status: 'subscribed' });
        cat.enrollments.set(String(it.course_id), list);
        out.push({ batch_item_id: it.batch_item_id, success: true });
      }
    }
    return { data: out };
  }
  throw new HttpError(404, `Demo catalog cannot ${method} ${u.pathname}`);
}

function demoGet(path) {
  const cat = demoCatalog();
  const u = new URL(path, 'https://demo.local');
  const page = num(u.searchParams.get('page'), 1);
  const size = num(u.searchParams.get('page_size'), 200);
  const paged = (items) => ({ data: { items: items.slice((page - 1) * size, page * size), has_more_data: page * size < items.length } });
  let m;
  if (u.pathname === '/learn/v1/courses') {
    const q = (u.searchParams.get('search_text') || '').toLowerCase();
    return paged(cat.courses.filter((c) => !q || c.name.toLowerCase().includes(q) || c.code.toLowerCase().includes(q)));
  }
  if ((m = u.pathname.match(/^\/course\/v1\/courses\/(\d+)\/sessions$/))) return paged(cat.sessions.get(m[1]) || []);
  if ((m = u.pathname.match(/^\/course\/v1\/sessions\/(\d+)\/events$/))) return paged(cat.events.get(m[1]) || []);
  if ((m = u.pathname.match(/^\/course\/v1\/events\/(\d+)\/attendance$/))) return paged(cat.attendance.get(m[1]) || []);
  if (u.pathname === '/course/v1/courses/enrollments') return paged(cat.enrollments.get(u.searchParams.get('course_id[]')) || []);
  const q = (u.searchParams.get('search_text') || '').toLowerCase();
  if (u.pathname === '/learn/v1/location') return paged(DEMO_LOCATIONS);
  if ((m = u.pathname.match(/^\/learn\/v1\/location\/(\d+)\/classroom$/))) return paged(DEMO_CLASSROOMS[m[1]] || []);
  if (u.pathname === '/audiences/v1/audience') return paged(DEMO_AUDIENCES.filter((a) => a.type === 'audience' && (!q || a.name.toLowerCase().includes(q))).map((a) => ({ uuid: a.id, name: a.name, type: 'manual' })));
  if (u.pathname === '/manage/v1/orgchart') return paged(DEMO_AUDIENCES.filter((a) => a.type === 'branch' && (!q || a.name.toLowerCase().includes(q))).map((a) => ({ id: Number(a.id), title: a.name, code: a.name.split('/').pop().trim().toUpperCase().replace(/\W+/g, '-') })));
  if ((m = u.pathname.match(/^\/audiences\/v1\/audience\/([\w-]+)\/users$/))) return paged(demoMembers(m[1]).map((x) => ({ idst: String(x.user_id), userid: '/' + x.username, firstname: x.first_name, lastname: x.last_name })));
  if ((m = u.pathname.match(/^\/manage\/v1\/orgchart\/(\d+)\/users$/))) return paged(demoMembers(m[1]));
  throw new HttpError(404, `Demo catalog has no data for ${u.pathname}`);
}

/* ----------------------------- Excel import ----------------------------- */
// The browser parses the workbook (SheetJS) and posts clean rows here.
// Programs match existing ones by name (case-insensitive) and are updated;
// milestones match on program + month + title and are updated, else added.

async function importPlan(env, user, b) {
  const programs = Array.isArray(b.programs) ? b.programs.slice(0, 200) : [];
  const milestones = Array.isArray(b.milestones) ? b.milestones.slice(0, 2000) : [];
  if (!programs.length && !milestones.length) throw new HttpError(400, 'The file has no programs or milestones.');
  const t = now();
  let planId = b.plan_id ? String(b.plan_id) : null;
  if (planId) {
    if (!(await first(env, 'SELECT id FROM plans WHERE id = ?', planId))) throw new HttpError(404, 'Plan not found.');
  } else {
    const np = b.new_plan || {};
    planId = uid();
    await insertRow(env, 'plans', { id: planId, name: clampStr(np.name || 'Imported training plan'), name_zh: clampStr(np.name_zh), year: Math.round(num(np.year, new Date().getUTCFullYear())), owner: user.username, created_at: t, updated_at: t });
  }

  const stmts = [];
  const res = { plan_id: planId, programs_created: 0, programs_updated: 0, milestones_created: 0, milestones_updated: 0, warnings: [] };
  const existingProgs = await all(env, 'SELECT * FROM programs WHERE plan_id = ?', planId);
  const progByName = new Map(existingProgs.map((p) => [p.name.trim().toLowerCase(), p]));
  let sort = existingProgs.length;

  const ensureProgram = (row, fromMilestone) => {
    const key = String(row.name || '').trim().toLowerCase();
    if (!key) return null;
    let p = progByName.get(key);
    const fields = cleanFields(row, PROGRAM_FIELDS);
    if (p) {
      if (!fromMilestone) {
        const upd = Object.fromEntries(Object.entries(fields).filter(([k, v]) => k !== 'name' && v !== '' && k !== 'sort' && k !== 'color'));
        if (Object.keys(upd).length) {
          const keys = Object.keys(upd);
          stmts.push(env.DB.prepare(`UPDATE programs SET ${keys.map((k) => `${k}=?`).join(',')}, updated_at=? WHERE id=?`).bind(...keys.map((k) => upd[k]), t, p.id));
          if (!p._counted) { res.programs_updated++; p._counted = true; }
        }
      }
      return p;
    }
    p = { id: uid(), plan_id: planId, name: clampStr(row.name), name_zh: '', objective: '', objective_zh: '', target_sessions: 0, target_trainees: 0, owner: '', color: sort % 8, sort: sort++, ...fields, updated_at: t };
    p.name = clampStr(row.name); p.color = 'color' in fields ? fields.color : p.color;
    const keys = ['id', 'plan_id', 'name', 'name_zh', 'objective', 'objective_zh', 'target_sessions', 'target_trainees', 'owner', 'color', 'sort', 'updated_at'];
    stmts.push(env.DB.prepare(`INSERT INTO programs (${keys.join(',')}) VALUES (${keys.map(() => '?').join(',')})`).bind(...keys.map((k) => p[k] ?? null)));
    progByName.set(key, p);
    res.programs_created++;
    if (fromMilestone) res.warnings.push(`Program "${row.name}" was not on the Programs sheet, so it was created with no annual targets.`);
    return p;
  };

  for (const row of programs) ensureProgram(row, false);

  // Audience names -> Docebo audience or branch (exact name match).
  const audCache = new Map();
  const audNames = [...new Set(milestones.map((m) => String(m.audience || '').trim()).filter(Boolean))].slice(0, 25);
  for (const name of audNames) {
    try {
      const hits = await searchAudiences(env, name);
      const hit = hits.find((h) => h.name.trim().toLowerCase() === name.toLowerCase());
      if (hit) {
        let size = null;
        try { size = (await audienceMembers(env, hit.type, hit.id)).length; } catch { /* size optional */ }
        audCache.set(name.toLowerCase(), { ...hit, size });
      } else res.warnings.push(`Audience "${name}" was not found in Docebo, so those milestones have no audience yet.`);
    } catch (e) { res.warnings.push(`Could not look up audience "${name}": ${e.message}`); }
  }

  const existingMs = await all(env, 'SELECT * FROM milestones WHERE plan_id = ?', planId);
  const msKey = (pid, month, title) => `${pid}|${month}|${String(title).trim().toLowerCase()}`;
  const msByKey = new Map(existingMs.map((m) => [msKey(m.program_id, m.month, m.title), m]));
  for (const row of milestones) {
    const p = ensureProgram({ name: row.program }, true);
    if (!p) continue;
    const fields = cleanFields({ ...row, program_id: p.id }, MILESTONE_FIELDS.filter((f) => !f.startsWith('audience')));
    if (!fields.month) continue;
    const aud = row.audience ? audCache.get(String(row.audience).trim().toLowerCase()) : null;
    const audFields = aud ? { audience_type: aud.type, audience_id: aud.id, audience_name: aud.name, audience_size: aud.size } : {};
    const key = msKey(p.id, fields.month, fields.title || '');
    const ex = msByKey.get(key);
    if (ex) {
      const upd = { ...Object.fromEntries(Object.entries(fields).filter(([, v]) => v !== '')), ...audFields, updated_at: t };
      const keys = Object.keys(upd);
      stmts.push(env.DB.prepare(`UPDATE milestones SET ${keys.map((k) => `${k}=?`).join(',')} WHERE id=?`).bind(...keys.map((k) => upd[k]), ex.id));
      res.milestones_updated++;
    } else {
      const m = { id: uid(), plan_id: planId, title: '', title_zh: '', planned_sessions: 0, planned_trainees: 0, region: '', delivery: '', notes: '', ...fields, ...audFields, updated_at: t };
      const keys = Object.keys(m);
      stmts.push(env.DB.prepare(`INSERT INTO milestones (${keys.join(',')}) VALUES (${keys.map(() => '?').join(',')})`).bind(...keys.map((k) => m[k] ?? null)));
      msByKey.set(key, m);
      res.milestones_created++;
    }
  }
  for (let i = 0; i < stmts.length; i += 90) await env.DB.batch(stmts.slice(i, i + 90));
  return res;
}

/* ----------------------------- CSV export ------------------------------- */

async function exportCsv(env, planId) {
  const st = await planState(env, planId);
  if (!st.plan) throw new HttpError(404, 'No plan.');
  const prog = new Map(st.programs.map((p) => [p.id, p]));
  const rows = [['Program', 'Month', 'Milestone', 'Region', 'Delivery', 'Planned sessions', 'Scheduled sessions', 'Delivered sessions', 'Planned trainees', 'Actual trainees', 'Linked Docebo items']];
  for (const ms of st.milestones) {
    const ls = st.links.filter((l) => l.milestone_id === ms.id);
    const s = (k) => ls.reduce((a, l) => a + num(l[k]), 0);
    rows.push([prog.get(ms.program_id)?.name || '', `${st.plan.year}-${String(ms.month).padStart(2, '0')}`, ms.title, ms.region, ms.delivery,
      ms.planned_sessions, s('act_sessions'), s('act_delivered'), ms.planned_trainees, s('act_trainees'), ls.map((l) => l.name).join(' | ')]);
  }
  const csv = '﻿' + rows.map((r) => r.map((v) => `"${String(v ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
  return new Response(csv, { headers: { 'content-type': 'text/csv; charset=utf-8', 'content-disposition': `attachment; filename="training-plan-${st.plan.year}.csv"` } });
}

/* ----------------------------- router ----------------------------------- */

/** The user's Docebo language preference, mapped to the planner's en / zh. */
async function userLanguage(env, userId) {
  try {
    const d = await doceboGet(env, `/manage/v1/user/${encodeURIComponent(userId)}`);
    const code = String(pick(d, ['data.user_data.lang_code', 'data.user_data.language', 'data.user_data.lang_browsercode']) || '').toLowerCase();
    if (!code) return null;
    return /zh|chinese|mandarin/.test(code) ? 'zh' : 'en';
  } catch {
    return null; // fall back to the browser language
  }
}

async function serveApp(env, url) {
  let token = null, user = null, reason = null;
  const h = await verifyFrameHandoff(env, url.searchParams);
  if (h.ok) {
    if (!userAllowed(env, h.userId)) reason = 'not-allowed';
    else user = { uid: h.userId, username: h.username };
  } else if (isDemo(env) || String(env.ALLOW_OPEN || '') === '1') {
    user = { uid: 'demo', username: 'demo.admin' };
  } else {
    reason = h.reason;
  }
  let docLang = null;
  if (user && !isDemo(env) && user.uid !== 'demo') docLang = await userLanguage(env, user.uid);
  if (user) token = await issueToken(env, user);
  const boot = { token, user, reason, mode: isDemo(env) ? 'demo' : 'live', domain: isDemo(env) ? null : String(env.DOCEBO_DOMAIN || ''), meetingUrl: String(env.DEFAULT_MEETING_URL || ''), lang: docLang };
  const html = APP_HTML.replace('/*__BOOT__*/null', () => JSON.stringify(boot).replace(/</g, '\\u003c'));
  return new Response(html, {
    headers: {
      'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store',
      'content-security-policy': "frame-ancestors 'self' https://*.docebosaas.com https://*.docebo.com" + (env.EXTRA_FRAME_ANCESTORS ? ' ' + env.EXTRA_FRAME_ANCESTORS : ''),
    },
  });
}

async function readBody(req) {
  try { return await req.json(); } catch { return {}; }
}

async function api(req, env, url) {
  const user = await readToken(env, req);
  await ensureSchema(env);
  const p = url.pathname.replace(/^\/api/, '');
  const method = req.method;
  let m;

  if (p === '/state' && method === 'GET') return json(await planState(env, url.searchParams.get('plan')));

  if (p === '/seed' && method === 'POST') {
    const b = await readBody(req);
    const year = Math.round(num(b.year, new Date().getUTCFullYear()));
    const id = await seedPlan(env, user, year, isDemo(env));
    if (isDemo(env)) await syncActuals(env, id);
    return json({ id });
  }

  if (p === '/plans' && method === 'POST') {
    const b = await readBody(req);
    const id = uid();
    await insertRow(env, 'plans', { id, name: clampStr(b.name || 'Training Plan'), name_zh: clampStr(b.name_zh), year: Math.round(num(b.year, new Date().getUTCFullYear())), owner: user.username, created_at: now(), updated_at: now() });
    return json({ id });
  }
  if ((m = p.match(/^\/plans\/(\w+)$/))) {
    if (method === 'PUT') {
      const b = await readBody(req);
      const row = {};
      if ('name' in b) row.name = clampStr(b.name);
      if ('name_zh' in b) row.name_zh = clampStr(b.name_zh);
      if ('year' in b) row.year = Math.round(num(b.year));
      row.updated_at = now();
      await updateRow(env, 'plans', m[1], row);
      return json({ ok: true });
    }
    if (method === 'DELETE') {
      await env.DB.batch(['links', 'milestones', 'programs'].map((t) => env.DB.prepare(`DELETE FROM ${t} WHERE plan_id = ?`).bind(m[1]))
        .concat([env.DB.prepare('DELETE FROM plans WHERE id = ?').bind(m[1]), env.DB.prepare('DELETE FROM meta WHERE k = ?').bind(`sync:${m[1]}`)]));
      return json({ ok: true });
    }
  }

  if (p === '/programs' && method === 'POST') {
    const b = await readBody(req);
    if (!b.plan_id) throw new HttpError(400, 'plan_id is required.');
    const id = uid();
    await insertRow(env, 'programs', { id, plan_id: clampStr(b.plan_id), ...cleanFields(b, PROGRAM_FIELDS), updated_at: now() });
    return json({ id });
  }
  if ((m = p.match(/^\/programs\/(\w+)$/))) {
    if (method === 'PUT') { await updateRow(env, 'programs', m[1], { ...cleanFields(await readBody(req), PROGRAM_FIELDS), updated_at: now() }); return json({ ok: true }); }
    if (method === 'DELETE') {
      const ids = (await all(env, 'SELECT id FROM milestones WHERE program_id = ?', m[1])).map((r) => r.id);
      const stmts = ids.map((id) => env.DB.prepare('DELETE FROM links WHERE milestone_id = ?').bind(id));
      stmts.push(env.DB.prepare('DELETE FROM milestones WHERE program_id = ?').bind(m[1]), env.DB.prepare('DELETE FROM programs WHERE id = ?').bind(m[1]));
      await env.DB.batch(stmts);
      return json({ ok: true });
    }
  }

  if (p === '/milestones' && method === 'POST') {
    const b = await readBody(req);
    if (!b.plan_id || !b.program_id || !b.month) throw new HttpError(400, 'plan_id, program_id and month are required.');
    const id = uid();
    await insertRow(env, 'milestones', { id, plan_id: clampStr(b.plan_id), ...cleanFields(b, MILESTONE_FIELDS), ...(await audienceSize(env, b)), updated_at: now() });
    return json({ id });
  }
  if ((m = p.match(/^\/milestones\/(\w+)\/create-sessions$/)) && method === 'POST') {
    const ms = await first(env, 'SELECT * FROM milestones WHERE id = ?', m[1]);
    if (!ms) throw new HttpError(404, 'Milestone not found.');
    const plan = await first(env, 'SELECT * FROM plans WHERE id = ?', ms.plan_id);
    const result = await createSessions(env, ms, plan, await readBody(req));
    const sync = result.created.length ? await syncActuals(env, ms.plan_id, ms.id) : null;
    return json({ ...result, sync });
  }
  if ((m = p.match(/^\/milestones\/(\w+)\/enrol$/)) && method === 'POST') {
    const ms = await first(env, 'SELECT * FROM milestones WHERE id = ?', m[1]);
    if (!ms) throw new HttpError(404, 'Milestone not found.');
    const b = await readBody(req);
    const r = await enrolAudience(env, ms, b.confirm === true);
    if (!r.preview) r.sync = await syncActuals(env, ms.plan_id, ms.id);
    return json(r);
  }
  if ((m = p.match(/^\/milestones\/(\w+)$/))) {
    if (method === 'PUT') { const b = await readBody(req); await updateRow(env, 'milestones', m[1], { ...cleanFields(b, MILESTONE_FIELDS), ...(await audienceSize(env, b)), updated_at: now() }); return json({ ok: true }); }
    if (method === 'DELETE') {
      await env.DB.batch([env.DB.prepare('DELETE FROM links WHERE milestone_id = ?').bind(m[1]), env.DB.prepare('DELETE FROM milestones WHERE id = ?').bind(m[1])]);
      return json({ ok: true });
    }
  }

  if (p === '/links' && method === 'POST') {
    const b = await readBody(req);
    const ms = await first(env, 'SELECT * FROM milestones WHERE id = ?', clampStr(b.milestone_id));
    if (!ms) throw new HttpError(404, 'Milestone not found.');
    const type = b.item_type === 'session' ? 'session' : 'course';
    if (!b.course_id || (type === 'session' && !b.session_id)) throw new HttpError(400, 'course_id (and session_id for sessions) is required.');
    const dup = await first(env, 'SELECT id FROM links WHERE milestone_id = ? AND course_id = ? AND COALESCE(session_id, \'\') = ?', ms.id, String(b.course_id), String(b.session_id || ''));
    if (dup) throw new HttpError(409, 'That Docebo item is already linked to this milestone.');
    const id = uid();
    await insertRow(env, 'links', {
      id, plan_id: ms.plan_id, milestone_id: ms.id, item_type: type, course_id: String(b.course_id),
      session_id: type === 'session' ? String(b.session_id) : null, name: clampStr(b.name || b.course_name),
      course_name: clampStr(b.course_name), course_type: clampStr(b.course_type, 40), created_at: now(),
    });
    const sync = await syncActuals(env, ms.plan_id, ms.id);
    return json({ id, sync });
  }
  if ((m = p.match(/^\/links\/(\w+)$/)) && method === 'DELETE') {
    await run(env, 'DELETE FROM links WHERE id = ?', m[1]);
    return json({ ok: true });
  }

  if (p === '/sync' && method === 'POST') {
    const b = await readBody(req);
    if (!b.plan_id) throw new HttpError(400, 'plan_id is required.');
    return json(await syncActuals(env, b.plan_id, b.milestone_id || null));
  }

  if (p === '/docebo/locations' && method === 'GET') return json({ items: await listLocations(env) });
  if (p === '/docebo/classrooms' && method === 'GET') return json({ items: await listClassrooms(env, url.searchParams.get('location_id') || '') });
  if (p === '/docebo/audiences' && method === 'GET') return json({ items: await searchAudiences(env, url.searchParams.get('q') || '') });
  if (p === '/docebo/courses' && method === 'GET') return json({ items: await searchCourses(env, url.searchParams.get('q') || '') });
  if (p === '/docebo/sessions' && method === 'GET') {
    const cid = url.searchParams.get('course_id');
    if (!cid) throw new HttpError(400, 'course_id is required.');
    return json({ items: await courseSessions(env, cid) });
  }
  // Read-only passthrough for checking field names against a tenant.
  if (p === '/docebo/raw' && method === 'GET') {
    const path = url.searchParams.get('path') || '';
    if (!path.startsWith('/') || path.includes('..')) throw new HttpError(400, 'path must start with /');
    return json(await doceboGet(env, path));
  }

  if (p === '/import' && method === 'POST') return json(await importPlan(env, user, await readBody(req)));
  if (p === '/export.csv' && method === 'GET') return exportCsv(env, url.searchParams.get('plan'));

  throw new HttpError(404, 'Not found.');
}

export default {
  async fetch(req, env) {
    const url = new URL(req.url);
    try {
      if (url.pathname.startsWith('/api/')) return await api(req, env, url);
      if (url.pathname === '/' || url.pathname === '/index.html') return await serveApp(env, url);
      if (url.pathname === '/favicon.ico') return new Response(null, { status: 204 });
      if (url.pathname === '/vendor/xlsx.js') return new Response(XLSX_JS, { headers: { 'content-type': 'application/javascript; charset=utf-8', 'cache-control': 'public, max-age=604800, immutable' } });
      return new Response('Not found', { status: 404 });
    } catch (e) {
      const status = e instanceof HttpError ? e.status : 500;
      if (status === 500 && !(e instanceof HttpError)) console.error(e);
      return json({ error: e.message || 'Server error' }, status);
    }
  },

  /** Nightly: refresh actuals for every plan. */
  async scheduled(_event, env, ctx) {
    ctx.waitUntil((async () => {
      await ensureSchema(env);
      const plans = await all(env, 'SELECT id FROM plans');
      for (const p of plans) {
        try { await syncActuals(env, p.id); } catch (e) { console.error('sync failed', p.id, e.message); }
      }
    })());
  },
};
