# Training Plan Planner for Docebo (aTalent POC)

Breaks an annual training plan into programs and monthly execution milestones, links each milestone to real Docebo ILT sessions or courses, and tracks plan vs actual from live Docebo data. One Cloudflare Worker with D1, embedded in Docebo through the Iframe widget. English / 中文 UI.

- `src/worker.js`: API, Docebo client, sync, demo catalog, nightly cron
- `src/app.html`: the planner UI (bundled into the Worker by wrangler)
- `wrangler.toml`: name, D1 binding, cron, plain settings
- `build.mjs`: optional, makes a single `dist/worker.js` for pasting into the dashboard editor

Cloudflare builds and deploys this repo on every push to `main`.

## 1. Deploy (once)

1. Cloudflare dashboard → **Workers & Pages → Create → Import a repository** → pick `docebo-training-plan-planner`. If it isn't listed, use **Connect GitHub** / the GitHub app settings to give Cloudflare access to this repo.
2. Keep the name `docebo-training-plan-planner`. Build command: empty. Deploy command: `npx wrangler deploy`. **Deploy**.
3. The D1 database `training-plan-planner` (APAC) already exists and is bound as `DB` in `wrangler.toml`. Tables are created automatically on first request.

## 2. Secrets

Worker → **Settings → Variables and Secrets** → add each as type **Secret**:

| Name | Value |
|---|---|
| `APP_SECRET` | any long random string (signs app sessions) |
| `SALT_SECRET` | long random string; the same value goes in the Docebo Iframe widget |
| `DOCEBO_CLIENT_ID`, `DOCEBO_CLIENT_SECRET` | OAuth2 app with the **Resource Owner Password Credentials** grant (you can reuse the attendance add-on app) |
| `DOCEBO_USERNAME`, `DOCEBO_PASSWORD` | service account that can see the relevant courses, sessions and enrollments |

Plain settings live in `wrangler.toml` (edit on GitHub; each push redeploys):

| Name | Meaning |
|---|---|
| `DEMO_MODE` | `"1"` = built-in construction machinery catalog, no Docebo calls, opens outside Docebo too. `"0"` = live Docebo data. |
| `DOCEBO_DOMAIN` | e.g. `atalentdemo2.docebosaas.com` |
| `SYNC_BUDGET` | max Docebo calls per sync run. `45` on Workers Free (50 subrequest limit); `400` on Workers Paid. |
| `ALLOWED_USER_IDS` | optional comma list of Docebo user IDs allowed in; empty = anyone who can see the Docebo page |

## 3. Docebo: embed the planner

1. **Admin Menu → Manage pages → Add new page → Widget page** (not External Link).
2. Add an **iFrame** widget with URL `https://docebo-training-plan-planner.<your-subdomain>.workers.dev/`.
3. **Advanced Settings → Salt Secret** = the same value as `SALT_SECRET`. Height around 1100 px.
4. Make the page visible to the L&D admins / planners group and add it to the menu.

Docebo signs each load (`user_id`, `username`, `auth_code`, SHA-256 hash); the Worker verifies it, so there is no separate sign-in.

## 4. Nightly sync

`wrangler.toml` sets a cron at 18:00 UTC (02:00 Singapore) that refreshes actuals for every plan, stalest links first. The **Sync from Docebo** button does the same on demand.

## 5. Going live with real Docebo data

1. In Docebo, create a few ILT courses and sessions that match the plan (e.g. an EX750 launch workshop course with sessions in different months, a technician certification course, an e-learning EHS course). Enrol a few test users and mark attendance on a past event.
2. In `wrangler.toml` set `DEMO_MODE = "0"` and commit. Cloudflare redeploys on its own.
3. Open the planner page in Docebo → **Create sample plan**. In live mode this creates the programs and milestones without links.
4. Open a milestone → **Link a Docebo course or ILT session** → search → pick an ILT session (or **Link whole course** for e-learning). Actuals load straight away.

### Docebo endpoints used

| Purpose | Endpoint | Status |
|---|---|---|
| Course search | `GET /learn/v1/courses?search_text=` | Confirmed in earlier add-ons |
| ILT sessions | `GET /course/v1/courses/{id}/sessions` | Confirmed |
| Events | `GET /course/v1/sessions/{id}/events` (`schedule.date`) | Confirmed |
| Class list and attendance | `GET /course/v1/events/{id}/attendance` (`user_id`, `status_id`) | Confirmed |
| Course enrollments and completions | `GET /course/v1/courses/enrollments?course_id[]=` (fallback `/learn/v1/enrollments?id_course=`) | **Check on your tenant** |

To check field names, open the browser console on the planner page and run:

```js
fetch('/api/docebo/raw?path=' + encodeURIComponent('/course/v1/courses/enrollments?course_id[]=123&page_size=5'),
  { headers: { authorization: 'Bearer ' + BOOT.token } }).then(r => r.json()).then(console.log)
```

Completions are read from `enrollment_status = completed` and `enrollment_completion_date` (several common variants are accepted). If your tenant uses other names, adjust `syncCourseLink()` in `worker.js`.

### How actuals are counted

- **ILT session link**: 1 scheduled session once it has events; delivered once its last event date has passed. Trainees = distinct learners marked present on its events (first 3 events per session). Enrolled = distinct learners on the class list.
- **Course link** (e-learning or whole course): trainees = completions dated in the milestone month of the plan year. Enrolled = all enrollments.
- **Status**: past months are Delivered at 90% or more of plan (sessions and trainees), At risk at 70% to 89%, Behind below 70%. Current and next month are At risk when fewer sessions exist in Docebo than planned. Milestones due by next month with nothing linked are Not linked.
- **Alignment check** flags: past milestones below plan, milestones not linked, Docebo sessions running outside the planned month, upcoming milestones not fully scheduled, and programs whose monthly breakdown does not add up to the annual target.

---

## 6. Suggested POC demo flow

1. **Overview**: annual targets fully broken down (100%), YTD delivered vs planned, plan vs actual by month, progress by program. Switch to 中文.
2. **Monthly breakdown**: each program's annual objective spread across months, colour-coded status per milestone. Filter by region (e.g. Middle East) to show dealer-market view.
3. Open a **past milestone**: linked Docebo sessions, enrolled vs attended, delivered.
4. Open an **upcoming unlinked milestone** → link a real Docebo ILT session live → status changes.
5. Edit a milestone's planned sessions or month → **Alignment check** immediately shows the annual target no longer adds up.
6. **Alignment check**: the dealer onboarding session that slipped from September into October, technician workshops with no Docebo sessions behind them.
7. **Export CSV** from the plan menu for management reporting.

## Notes and limits (POC)

- Data is stored in Cloudflare D1 (plan structure and synced actual counts only; no learner names). Hosting region follows Cloudflare's D1 placement, which may come up in a security review.
- One planner role: anyone allowed in can edit. Approval of the annual plan, budget and session auto-creation are phase 2 items.
- Docebo API limits: syncs are capped by `SYNC_BUDGET`; large plans finish over several clicks or nightly runs.
