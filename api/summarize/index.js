// Azure Functions (Node.js) — /api/summarize/index.js
// Uses metrics provided by the frontend to avoid re-deriving times.
// Ensures hours/minutes only (no seconds) in the narrative.

const DEFAULT_MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";

// Prefer env var key. Allow local dev override via request headers on localhost.
function resolveApiKey(req) {
  const envKey = process.env.OPENAI_API_KEY && String(process.env.OPENAI_API_KEY).trim();
  const isLocal =
    (req?.headers["origin"] && /localhost|127\.0\.0\.1/.test(String(req.headers["origin"]))) ||
    (req?.headers["referer"] && /localhost|127\.0\.0\.1/.test(String(req.headers["referer"])));

  if (isLocal) {
    const hdrAuth = req.headers["authorization"];
    if (hdrAuth && /^bearer\s+/i.test(hdrAuth)) {
      return hdrAuth.replace(/^bearer\s+/i, "").trim();
    }
    const hdrKey = req.headers["x-openai-key"];
    if (hdrKey && String(hdrKey).trim()) return String(hdrKey).trim();
  }
  return envKey || null;
}

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-OpenAI-Key, X-OpenAI-Model",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  };
}

module.exports = async function (context, req) {
  if (req.method === "OPTIONS") {
    context.res = { status: 204, headers: corsHeaders() };
    return;
  }

  const apiKey = resolveApiKey(req);
  if (!apiKey) {
    context.res = {
      status: 401,
      headers: corsHeaders(),
      jsonBody: { error: "Missing OpenAI API key. Set OPENAI_API_KEY in Azure (or send Authorization/X-OpenAI-Key in local dev)." },
    };
    return;
  }

  let body = {};
  try {
    body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
  } catch {
    body = {};
  }

  const {
    info = {},
    employees = [],
    taskLog = [],
    timeLog = [],
    photos = [],
    summaryText = "",
    metrics = {},
  } = body;

  // Helper to keep seconds out of the prose.
  const hm = (mins) => {
    const m = Math.max(0, Math.round(mins));
    const h = Math.floor(m / 60);
    const mm = m % 60;
    if (h === 0) return `${mm}m`;
    if (mm === 0) return `${h}h`;
    return `${h}h ${mm}m`;
  };

  // Pull metrics from the frontend; provide conservative defaults.
  const actualMinutes = Number.isFinite(metrics.actualMinutes) ? metrics.actualMinutes : 0;
  const touchMinutes  = Number.isFinite(metrics.touchMinutes)  ? metrics.touchMinutes  : 0;
  const idleMinutes   = Number.isFinite(metrics.idleMinutes)   ? metrics.idleMinutes   : 0;

  const actualHM = hm(actualMinutes);
  const touchHM  = hm(touchMinutes);
  const idleHM   = hm(idleMinutes);

  const utilizationPct = Number.isFinite(metrics.utilizationPct) ? metrics.utilizationPct : 0;
  const crewHours      = Number.isFinite(metrics.crewHours)      ? metrics.crewHours      : +(touchMinutes/60).toFixed(2);
  const idleRatioPct   = Number.isFinite(metrics.idleRatioPct)   ? metrics.idleRatioPct   : (touchMinutes + idleMinutes > 0 ? +(idleMinutes / (touchMinutes + idleMinutes) * 100).toFixed(1) : 0);

  const totalEmployees = Number.isFinite(metrics.totalEmployees) ? metrics.totalEmployees : (Array.isArray(employees) ? employees.length : 0);
  const totalSessions  = Number.isFinite(metrics.totalSessions)  ? metrics.totalSessions  : (Array.isArray(timeLog) ? timeLog.filter(t => t && t.event !== "deleted").length : 0);

  // Build a compact, reliable context (avoid huge logs).
  const recentNotes = (Array.isArray(taskLog) ? taskLog : []).slice(-12).map(n => n && n.text).filter(Boolean);
  const samplePhotos = (Array.isArray(photos) ? photos : []).slice(0, 5).map(p => ({
    name: p?.name || p?.caption || "photo",
  }));

  const model = req.headers["x-openai-model"] || DEFAULT_MODEL;

  const messages = [
    {
      role: "system",
      content:
        "You write concise, plain-English work-measurement summaries. NEVER mention seconds; ONLY use hours and minutes (e.g., 0h 3m, 18m, 2h 15m). Be neutral, factual, and readable to non-experts.",
    },
    {
      role: "user",
      content: JSON.stringify(
        {
          header: {
            date: info?.date || "",
            endDate: info?.multiDay ? (info?.endDate || "") : "",
            location: info?.location || "",
            procedure: info?.procedure || "",
            workOrder: info?.workOrder || "",
            task: info?.task || "",
            type: info?.type || "",
            workType: info?.workType || "",
            assetId: info?.assetId || "",
            station: info?.station || "",
            supervisor: info?.supervisor || "",
            observer: info?.observer || "",
            estimatedTime: info?.estimatedTime || "",
            observationScope: info?.observationScope || "Full",
          },
          crew: {
            totalEmployees,
            employees: (Array.isArray(employees) ? employees : []).map(e => ({
              name: e?.name || "",
              role: e?.role || "",
              skill: e?.skill || "",
            })),
          },
          metrics: {
            actualHM,
            touchHM,
            idleHM,
            utilizationPct,
            crewHours,
            idleRatioPct,
            totalSessions,
          },
          notes: recentNotes,
          photos: samplePhotos,
          operatorNotes: summaryText || "",
        },
        null,
        2
      ),
    },
    {
      role: "system",
      content:
        [
          "Write a concise 1–2 paragraph summary:",
          "• First paragraph: What task, where, and overall timing: Actual, Touch, Idle (hours/minutes only).",
          "• Second paragraph: Key observations (delays, roles, scope) and efficiency: utilization and crew-hours.",
          "Rules:",
          "- DO NOT use seconds; only hours/minutes.",
          "- Be readable for someone unfamiliar with the task; no jargon.",
          "- If estimatedTime is present, briefly compare expected vs actual.",
        ].join("\n"),
    },
  ];

  // Call OpenAI
  let summaryTextOut = "";
  try {
    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        messages,
        temperature: 0.3,
        max_tokens: 350,
      }),
    });

    if (!resp.ok) {
      const errTxt = await resp.text().catch(() => "");
      throw new Error(`OpenAI error ${resp.status}: ${errTxt}`);
    }
    const data = await resp.json();
    summaryTextOut = data?.choices?.[0]?.message?.content?.trim() || "";
  } catch (err) {
    context.log.error("OpenAI call failed", err);
    // Fallback templated summary using provided metrics
    summaryTextOut =
      `Observed task "${info?.task || ""}" at ${info?.location || "the specified location"}. ` +
      `Total Actual time ${actualHM}; Touch labor ${touchHM}; Idle ${idleHM}. ` +
      `Crew size ${totalEmployees} across ${totalSessions} session(s). ` +
      `Utilization ${utilizationPct}% with ${crewHours} crew-hours and Idle Ratio ${idleRatioPct}%.`;
  }

  context.res = {
    status: 200,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      ...corsHeaders(),
    },
    body: JSON.stringify({ summary: summaryTextOut }),
  };
};