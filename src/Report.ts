// src/Report.ts

export type EmpStatus = "idle" | "active" | "paused";
export type TimeEvent = "start" | "pause" | "stop" | "deleted";

export interface Employee {
  id: number;
  name: string;
  status: EmpStatus;
  startTime: number | null;
  elapsedTime: number;
  pausedAccum: number;
  lastPausedAt: number | null;
  logs: string[];
  role?: string;
  skill?: string;
}

export interface TimeLogEntry {
  id: number;
  at: number;
  employeeId: number | null;
  employeeName: string;
  event: TimeEvent;
  reasonCode?: string;
  comment?: string;
}

export interface TaskEntry {
  id: number;
  at: number;
  text: string;
}

export interface Info {
  date: string;
  endDate: string;
  multiDay: boolean;
  location: string;
  procedure: string;
  workOrder: string;
  task: string;
  type?: string;      // Routine / Non-Routine / Customer Request / Cannibalization / Other
  workType?: string;  // Inspection / Remove & Replace / Setup / Rework / Test / Remove / Install / Other
  assetId?: string;
  station?: string;
  supervisor?: string;
  observer?: string;
}

type LiveTimesFn = (e: Employee) => { active: number; idle: number; total: number };
type MsToTimeFn = (ms: number) => string;
type FmtStampFn = (at: number, withDate: boolean) => string;

const pad2 = (n: number) => String(n).padStart(2, "0");
const msToHMS = (ms: number) => {
  const s = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = s % 60;
  return `${pad2(h)}:${pad2(m)}:${pad2(ss)}`;
};

// Helper to satisfy TS noUnused* checks without changing behavior
const _use = (..._args: unknown[]) => {};

function baseStyles() {
  return `
  :root{--ink:#e8eeff;--muted:#aabcdf;--line:#2a3560;--bg:#0b1228;--card:#0f1731}
  *{box-sizing:border-box}
  body{margin:0;font:14px/1.45 ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial;color:var(--ink);background:var(--bg)}
  .wrap{max-width:960px;margin:24px auto;padding:16px}
  h1{font-size:22px;margin:0 0 2px}
  h2{font-size:16px;margin:16px 0 8px}
  h3{font-size:13px;margin:12px 0 6px;color:#cfe0ff}
  .meta{color:var(--muted);font-size:12px}
  .card{background:linear-gradient(180deg,#0f1731,#0e152f);border:1px solid var(--line);border-radius:12px;padding:12px;margin:10px 0}
  .grid{display:grid;gap:10px}
  .grid.cols-2{grid-template-columns: 1fr 1fr}
  .mono{font-family:ui-monospace,Menlo,Consolas,monospace}
  table{width:100%;border-collapse:collapse}
  th,td{border-bottom:1px solid #202b54;padding:8px;text-align:left;vertical-align:top}
  th{font-size:12px;color:#cfe0ff}
  .kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
  .kpi{border:1px solid #2a3560;border-radius:10px;padding:10px;display:grid;gap:6px;justify-items:center;background:transparent}
  .kpi .label{font-size:12px;color:#cfe0ff}
  .kpi .num{font-size:20px;font-weight:800;font-family:ui-monospace,Menlo,Consolas,monospace}
  .defs{font-size:12px;color:var(--muted)}
  .defs dt{color:#cfe0ff;font-weight:600;margin-top:8px}
  .defs dd{margin:2px 0 6px 0}
  @media print{
    body{background:white;color:black}
    .card{border-color:#ccc}
    .kpi{border-color:#ccc}
    .meta{color:#444}
  }
  `;
}

function escapeHTML(s: string) {
  return s.replace(/[&<>"']/g, (c) =>
    c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : c === '"' ? "&quot;" : "&#39;",
  );
}

/* --------- KPI Calculation (with "any engaged" rule) --------- */
function calcKPIs(
  info: Info,
  employees: Employee[],
  timeLog: TimeLogEntry[],
  liveTimes: LiveTimesFn,
) {
  _use(info);
  const now = Date.now();
  const totalActive = employees.reduce((acc, e) => acc + liveTimes(e).active, 0);
  const totalIdle = employees.reduce((acc, e) => acc + liveTimes(e).idle, 0);
  const totalAll = totalActive + totalIdle;

  const starts = timeLog.filter((t) => t.event === "start").map((t) => t.at);
  const firstStartAt = starts.length ? Math.min(...starts) : null;
  const ends = timeLog.filter((t) => t.event === "stop" || t.event === "deleted").map((t) => t.at);
  const lastStopAt = ends.length ? Math.max(...ends) : null;

  // FIX: Actual time ends at "now" while anyone is Active/Paused; else at last stop.
  const anyEngagedNow = employees.some((e) => e.status === "active" || e.status === "paused");
  const endForActual = anyEngagedNow ? now : (lastStopAt ?? now);
  const actualClockMs = firstStartAt ? Math.max(0, endForActual - firstStartAt) : 0;

  const utilization = actualClockMs ? totalActive / actualClockMs : 0;
  const crewHours = totalActive / 3_600_000;
  const idleRatio = totalAll ? totalIdle / totalAll : 0;

  // Daily breakdown
  const daily: Record<string, { actualMs: number; touchMs: number; idleMs: number }> = {};
  const evs = [...timeLog].sort((a, b) => a.at - b.at);
  if (evs.length) {
    const snapshot: Record<number, EmpStatus> = {};
    const empSet = new Set(evs.map((e) => e.employeeId).filter((x): x is number => typeof x === "number"));
    empSet.forEach((id) => (snapshot[id] = "idle"));
    let tPrev = evs[0].at;

    const addSpan = (t0: number, t1: number) => {
      if (t1 <= t0) return;
      const activeCnt = Object.values(snapshot).filter((s) => s === "active").length;
      const pausedCnt = Object.values(snapshot).filter((s) => s === "paused").length;

      let a = t0;
      while (a < t1) {
        const dayEnd = new Date(new Date(a).toDateString()).getTime() + 24 * 3600 * 1000;
        const b = Math.min(t1, dayEnd);
        const key = new Date(a).toISOString().slice(0, 10);
        const dur = b - a;
        if (!daily[key]) daily[key] = { actualMs: 0, touchMs: 0, idleMs: 0 };
        daily[key].actualMs += dur;
        daily[key].touchMs += dur * activeCnt;
        daily[key].idleMs += dur * pausedCnt;
        a = b;
      }
    };

    for (const ev of evs) {
      addSpan(tPrev, ev.at);
      if (typeof ev.employeeId === "number") {
        if (ev.event === "start") snapshot[ev.employeeId] = "active";
        if (ev.event === "pause") snapshot[ev.employeeId] = "paused";
        if (ev.event === "stop" || ev.event === "deleted") snapshot[ev.employeeId] = "idle";
      }
      tPrev = ev.at;
    }
    if (lastStopAt == null) addSpan(tPrev, now);
  }

  return { totalActive, totalIdle, totalAll, actualClockMs, utilization, crewHours, idleRatio, daily };
}

/* ------------------------------ Render ------------------------------ */
function renderHTML(
  info: Info,
  employees: Employee[],
  timeLog: TimeLogEntry[],
  taskLog: TaskEntry[],
  liveTimes: LiveTimesFn,
  msToTime: MsToTimeFn,
  fmtStamp: FmtStampFn,
) {
  _use(msToTime);
  const genAt = new Date();
  const { totalActive, totalIdle, totalAll, actualClockMs, utilization, crewHours, idleRatio, daily } = calcKPIs(
    info,
    employees,
    timeLog,
    liveTimes,
  );
  _use(totalAll);

  const perfRows = employees
    .map((e) => {
      const { active, idle, total } = liveTimes(e);
      const status = e.status === "idle" && (active > 0 || idle > 0) ? "Completed" : e.status;
      return `<tr>
        <td>${escapeHTML(e.name)}</td>
        <td>${escapeHTML(e.role || "—")}</td>
        <td>${escapeHTML(e.skill || "—")}</td>
        <td>${escapeHTML(status)}</td>
        <td class="mono">${msToHMS(active)}</td>
        <td class="mono">${msToHMS(idle)}</td>
        <td class="mono">${msToHMS(total)}</td>
      </tr>`;
    })
    .join("");

  const timeRows = [...timeLog]
    .sort((a, b) => a.at - b.at)
    .map(
      (t) => `
    <tr>
      <td class="mono">${fmtStamp(t.at, info.multiDay)}</td>
      <td>${escapeHTML(t.employeeName)}</td>
      <td>${t.event}</td>
      <td>${t.reasonCode ? escapeHTML(t.reasonCode) : ""}</td>
      <td>${t.comment ? escapeHTML(t.comment) : ""}</td>
    </tr>
  `,
    )
    .join("");

  const taskRows = [...taskLog]
    .sort((a, b) => a.at - b.at)
    .map(
      (n) => `
    <tr>
      <td class="mono">${fmtStamp(n.at, info.multiDay)}</td>
      <td>${escapeHTML(n.text)}</td>
    </tr>
  `,
    )
    .join("");

  const dailyRows = Object.entries(daily)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(
      ([date, v]) => `
    <tr>
      <td class="mono">${date}</td>
      <td class="mono">${msToHMS(v.actualMs)}</td>
      <td class="mono">${msToHMS(v.touchMs)}</td>
      <td class="mono">${msToHMS(v.idleMs)}</td>
      <td class="mono">${v.actualMs ? ((v.touchMs / v.actualMs) * 100).toFixed(1) : "0.0"}%</td>
    </tr>
  `,
    )
    .join("");

  const metricDefs = `
    <dl class="defs">
      <dt>Type</dt>
      <dd>Priority category (Routine, Non-Routine, Customer Request, Cannibalization, or Other).</dd>

      <dt>Work Type</dt>
      <dd>Nature of work (Inspection, Remove & Replace, Setup, Rework, Test, Remove, Install, or Other).</dd>

      <dt>Actual Time</dt>
      <dd>Wall-clock time from first <em>Start</em> to now while anyone is Active/Paused; otherwise to the last <em>Stop</em>/<em>Delete</em>.</dd>

      <dt>Touch Labor</dt>
      <dd>Sum of time employees are <em>Active</em> (crew-weighted). Two people active for 1 hour = 2 touch-hours.</dd>

      <dt>Idle Time</dt>
      <dd>Sum of time employees are <em>Paused</em> (crew-weighted). Stopped employees are not counted as idle.</dd>

      <dt>Total Time</dt>
      <dd>Touch + Idle (crew-weighted time, not wall-clock).</dd>

      <dt>Utilization</dt>
      <dd>Touch Labor ÷ Actual Time.</dd>

      <dt>Crew-hours</dt>
      <dd>Touch Labor expressed in hours (Σ active ÷ 3600s).</dd>

      <dt>Idle Ratio</dt>
      <dd>Idle ÷ (Touch + Idle).</dd>

      <dt>Daily Breakdown</dt>
      <dd>Actual/Touch/Idle apportioned per calendar day; spans crossing midnight are split.</dd>
    </dl>
  `;

  return `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<title>Work Measurement Report</title>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<style>${baseStyles()}</style>
</head>
<body>
  <div class="wrap">
    <h1>Work Measurement Study — Report</h1>
    <div class="meta">Generated ${genAt.toLocaleString()}</div>

    <div class="card grid cols-2" style="margin-top:12px">
      <div>
        <div><span class="meta">Date:</span> ${escapeHTML(info.date)}</div>
        ${info.multiDay ? `<div><span class="meta">End Date:</span> ${escapeHTML(info.endDate || "")}</div>` : ""}
        ${info.type ? `<div><span class="meta">Type:</span> ${escapeHTML(info.type)}</div>` : ""}
        ${info.workType ? `<div><span class="meta">Work Type:</span> ${escapeHTML(info.workType)}</div>` : ""}
        ${info.assetId ? `<div><span class="meta">Asset ID:</span> ${escapeHTML(info.assetId)}</div>` : ""}
        ${info.station ? `<div><span class="meta">Station/Area:</span> ${escapeHTML(info.station)}</div>` : ""}
        ${info.supervisor ? `<div><span class="meta">Supervisor:</span> ${escapeHTML(info.supervisor)}</div>` : ""}
        ${info.observer ? `<div><span class="meta">Observer:</span> ${escapeHTML(info.observer)}</div>` : ""}
        <div><span class="meta">Location:</span> ${escapeHTML(info.location || "—")}</div>
        <div><span class="meta">Procedure:</span> ${escapeHTML(info.procedure || "—")}</div>
        <div><span class="meta">Work Order:</span> ${escapeHTML(info.workOrder || "—")}</div>
        <div><span class="meta">Task:</span> ${escapeHTML(info.task || "—")}</div>
      </div>
      <div>
        <div class="kpis">
          <div class="kpi"><div class="label">Actual</div><div class="num">${msToHMS(actualClockMs)}</div></div>
          <div class="kpi"><div class="label">Touch</div><div class="num">${msToHMS(totalActive)}</div></div>
          <div class="kpi"><div class="label">Idle</div><div class="num">${msToHMS(totalIdle)}</div></div>
          <div class="kpi"><div class="label">Utilization</div><div class="num">${(utilization * 100).toFixed(1)}%</div></div>
          <div class="kpi"><div class="label">Crew-hours</div><div class="num">${crewHours.toFixed(2)}</div></div>
          <div class="kpi"><div class="label">Idle Ratio</div><div class="num">${(idleRatio * 100).toFixed(1)}%</div></div>
        </div>
      </div>
    </div>

    <div class="card">
      <h2>Employee Performance</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr><th>Employee</th><th>Role</th><th>Skill</th><th>Status</th><th>Active (Touch)</th><th>Idle</th><th>Total</th></tr>
          </thead>
          <tbody>${perfRows || `<tr><td colspan="7" class="meta">No employees.</td></tr>`}</tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <h2>Time Log</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr><th>When</th><th>Employee</th><th>Event</th><th>Reason</th><th>Comment</th></tr>
          </thead>
          <tbody>${timeRows || `<tr><td colspan="5" class="meta">No time log entries.</td></tr>`}</tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <h2>Task Log</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>When</th><th>Note</th></tr></thead>
          <tbody>${taskRows || `<tr><td colspan="2" class="meta">No notes.</td></tr>`}</tbody>
        </table>
      </div>
    </div>

    ${Object.keys(daily).length ? `
      <div class="card">
        <h2>Daily Breakdown</h2>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Date</th><th>Actual</th><th>Touch</th><th>Idle</th><th>Utilization</th></tr></thead>
            <tbody>${dailyRows}</tbody>
          </table>
        </div>
      </div>
    ` : ""}

    <div class="card">
      <h2>Metric Definitions</h2>
      ${metricDefs}
    </div>

    <div class="meta" style="margin-top:12px">Report generated by Work Measurement App</div>
  </div>
</body>
</html>`;
}

/* ------------------------------ API ------------------------------ */
export function printReportHTML(
  info: Info,
  employees: Employee[],
  timeLog: TimeLogEntry[],
  taskLog: TaskEntry[],
  liveTimes: LiveTimesFn,
  msToTime: MsToTimeFn,
  fmtStamp: FmtStampFn,
) {
  const html = renderHTML(info, employees, timeLog, taskLog, liveTimes, msToTime, fmtStamp);
  const win = window.open("", "_blank");
  if (!win) return;
  win.document.open();
  win.document.write(html);
  win.document.close();
  setTimeout(() => {
    try {
      win.focus();
      win.print();
    } catch {}
  }, 200);
}

export function exportReportHTML(
  info: Info,
  employees: Employee[],
  timeLog: TimeLogEntry[],
  taskLog: TaskEntry[],
  liveTimes: LiveTimesFn,
  msToTime: MsToTimeFn,
  fmtStamp: FmtStampFn,
) {
  const html = renderHTML(info, employees, timeLog, taskLog, liveTimes, msToTime, fmtStamp);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "work_measurement_report.html";
  a.click();
  URL.revokeObjectURL(url);
}