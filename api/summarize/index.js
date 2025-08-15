// Azure Functions (Node 18+). Uses native fetch.
// POST /api/summarize with JSON: { info, employees, timeLog, taskLog, summaryText? }

const MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";
const OPENAI_URL = "https://api.openai.com/v1/chat/completions";

function trimJSON(obj, maxChars = 12000) {
  const s = JSON.stringify(obj);
  if (s.length <= maxChars) return s;
  return s.slice(0, maxChars) + " …(truncated)…";
}

module.exports = async function (context, req) {
  try {
    const key = process.env.OPENAI_API_KEY || process.env.OPENAI_KEY;
    if (!key) {
      context.res = {
        status: 500,
        body: { error: "Server missing OPENAI_API_KEY" }
      };
      return;
    }

    const { info, employees, timeLog, taskLog, summaryText } = req.body || {};
    if (!info) {
      context.res = { status: 400, body: { error: "Missing 'info' in request body" } };
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

    const user = `
Study data (JSON): 
${payloadStr}

If the user has typed anything in "Summary" already, use it as hints (optional):
${summaryText ? `"${summaryText}"` : "(none)"}
`;

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
      context.res = { status: 502, body: { error: "OpenAI error", detail: errText } };
      return;
    }

    const data = await resp.json();
    const summary = data.choices?.[0]?.message?.content?.trim() || "(no summary)";

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: { summary }
    };
  } catch (err) {
    context.res = { status: 500, body: { error: err?.message || String(err) } };
  }
};
