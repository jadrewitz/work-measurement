// Azure Functions (Node 18+)
// GET  /api/summarize       -> health check JSON (no key required)
// POST /api/summarize       -> generate summary using OpenAI (requires key)

const MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";
const OPENAI_URL = "https://api.openai.com/v1/chat/completions";

// --- helpers ---------------------------------------------------------------
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type,Authorization"
};

function cors(res) {
  return { ...(res || {}), headers: { ...(res?.headers || {}), ...CORS } };
}

function trimJSON(obj, maxChars = 12000) {
  const s = JSON.stringify(obj);
  return s.length <= maxChars ? s : s.slice(0, maxChars) + " …(truncated)…";
}

// --- function entry --------------------------------------------------------
module.exports = async function (context, req) {
  try {
    const method = (req.method || "GET").toUpperCase();

    // Preflight for browsers
    if (method === "OPTIONS") {
      context.res = cors({ status: 204, body: "" });
      return;
    }

    // Health check
    if (method === "GET") {
      context.res = cors({
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: true,
          message: "summarize API is alive",
          model: MODEL
        }
      });
      return;
    }

    // Only POST below here
    if (method !== "POST") {
      context.res = cors({ status: 405, body: { error: "Method not allowed" } });
      return;
    }

    const key = process.env.OPENAI_API_KEY || process.env.OPENAI_KEY;
    if (!key) {
      context.res = cors({ status: 500, body: { error: "Server missing OPENAI_API_KEY" } });
      return;
    }

    const { info, employees, timeLog, taskLog, summaryText } = req.body || {};
    if (!info) {
      context.res = cors({ status: 400, body: { error: "Missing 'info' in request body" } });
      return;
    }

    const compact = {
      info,
      employees: (employees || []).map(e => ({
        name: e.name, role: e.role, skill: e.skill,
        status: e.status, elapsedTime: e.elapsedTime, pausedAccum: e.pausedAccum
      })),
      timeLog: (timeLog || []).slice(-300),
      taskLog: (taskLog || []).slice(-200)
    };
    const payloadStr = trimJSON(compact, 12000);

    const system = `You are an industrial engineering assistant.
Write a clear, objective summary for a work measurement study.
Use professional, neutral language.
If 'Observation Scope' is partial, mention that.
Include: scope, dates, crew makeup, notable pauses/idle causes, KPIs (utilization, crew-hours, idle ratio), and any risks/opportunities.
Keep it concise (6–10 sentences).`;

    const user = `\nStudy data (JSON):\n${payloadStr}\n\nIf the user has typed anything in "Summary" already, use it as hints (optional):\n${summaryText ? `"${summaryText}"` : "(none)"}\n`;

    const resp = await fetch(OPENAI_URL, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${key}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: MODEL,
        temperature: 0.2,
        messages: [
          { role: "system", content: system },
          { role: "user", content: user }
        ]
      })
    });

    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      context.res = cors({ status: 502, body: { error: "OpenAI error", detail: errText } });
      return;
    }

    const data = await resp.json();
    const summary = data.choices?.[0]?.message?.content?.trim() || "(no summary)";

    context.res = cors({
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { summary }
    });
  } catch (err) {
    context.res = cors({ status: 500, body: { error: err?.message || String(err) } });
  }
};
