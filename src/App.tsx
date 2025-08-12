import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { printReportHTML, exportReportHTML, exportReportPDF } from "./Report";
import "./ui.css";
import heic2any from "heic2any";

async function toDataUrlIfNeeded(src: string): Promise<string> {
  if (!src) return "";
  if (src.startsWith("data:")) return src;
  try {
    const res = await fetch(src);
    const blob = await res.blob();
    const fr = new FileReader();
    return await new Promise<string>((resolve, reject) => {
      fr.onload = () => resolve(String(fr.result));
      fr.onerror = () => reject(fr.error);
      fr.readAsDataURL(blob);
    });
  } catch {
    return "";
  }
}

async function buildReportPhotos(raw: any[]): Promise<{data: string; name?: string}[]> {
  const mapped = raw.map(p => ({
    data: p.data || p.url || p.preview || p.src || "",
    name: p.name || ""
  }));
  const out: {data: string; name?: string}[] = [];
  for (const p of mapped) {
    const data = await toDataUrlIfNeeded(p.data);
    if (data) out.push({ data, name: p.name });
  }
  return out;
}

/* ---------- Types ---------- */
type EmpStatus = "idle" | "active" | "paused";
type TimeEvent = "start" | "pause" | "stop" | "deleted";

interface Employee {
  id: number;
  name: string;
  status: EmpStatus;
  startTime: number | null;
  elapsedTime: number;
  pausedAccum: number;
  lastPausedAt: number | null;
  logs: string[];
  role?: string;   // Mechanic / Inspector / Lead / Helper / Trainee / Other…
  skill?: string;  // A&P / Structures / Avionics / QA / NDT / Non-Certified / Cabin / Other…
}

interface TaskEntry {
  id: number;
  at: number;
  text: string;
}

interface TimeLogEntry {
  id: number;
  at: number;
  employeeId: number | null;
  employeeName: string;
  event: TimeEvent;
  reasonCode?: string;
  comment?: string;
}

interface PhotoItem {
  id: number;
  dataUrl: string;   // already resized and auto-rotated
  name?: string;
  caption?: string;
}

interface AppInfo {
  date: string;
  endDate: string;
  multiDay: boolean;
  location: string;
  procedure: string;
  workOrder: string;
  task: string;
  type?: string;       // Routine / Non-Routine / Customer Request / Cannibalization / Other
  workType?: string;   // Inspection / Remove & Replace / Setup / Rework / Test / Remove / Install / Other
  assetId?: string;
  station?: string;
  supervisor?: string;
  observer?: string;
}

interface AppState {
  info: AppInfo;
  employees: Employee[];
  taskLog: TaskEntry[];
  timeLog: TimeLogEntry[];
  photos: PhotoItem[];
}

/* ---------- Options ---------- */
const TYPE_OPTIONS = [
  "Routine",
  "Non-Routine",
  "Customer Request",
  "Cannibalization",
  "Other…",
] as const;

const WORKTYPE_OPTIONS = [
  "Inspection",
  "Remove & Replace",
  "Setup",
  "Rework",
  "Test",
  "Remove",
  "Install",
  "Other…",
] as const;

const ROLE_OPTIONS = ["Mechanic", "Inspector", "Lead", "Helper", "Trainee", "Other…"] as const;

const SKILL_OPTIONS = ["A&P", "Structures", "Avionics", "QA", "NDT", "Non-Certified", "Cabin", "Other…"] as const;

/* ---------- Helpers ---------- */
const STORAGE_KEY = "work-measurement:v1";
const LAST_OBSERVER_KEY = "work-measurement:last-observer";

const PAUSE_REASONS = [
  "Waiting on parts",
  "Waiting on tooling",
  "QA / Inspection",
  "Engineering review",
  "Break",
  "Shift change",
  "Personal",
  "Other",
];

const STOP_REASONS = [
  "Task Complete",
  "Reassigned",
  "Error",
  "Other",
];

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function msToTime(ms: number) {
  const s = Math.floor(Math.max(0, ms) / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = s % 60;
  return `${h}h ${m}m ${ss}s`;
}
const pad2 = (n: number) => String(n).padStart(2, "0");
function msToHMS(ms: number) {
  const s = Math.floor(Math.max(0, ms) / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = s % 60;
  return `${pad2(h)}:${pad2(m)}:${pad2(ss)}`;
}
function fmtStamp(at: number, withDate: boolean) {
  const d = new Date(at);
  return withDate ? `${d.toLocaleDateString()} ${d.toLocaleTimeString()}` : d.toLocaleTimeString();
}

// --- Image helpers (resize + auto-rotate) ---
function getJpegOrientation(arrayBuffer: ArrayBuffer): number | null {
  const view = new DataView(arrayBuffer);
  if (view.getUint16(0, false) !== 0xFFD8) return null; // not a JPEG
  let offset = 2;
  const length = view.byteLength;
  while (offset < length) {
    const marker = view.getUint16(offset, false);
    offset += 2;
    if (marker === 0xFFE1) {
      // APP1 segment (EXIF)
      // Skip EXIF segment length (we don't need the value)
      offset += 2;
      // "Exif" (0x45786966) + null
      if (view.getUint32(offset, false) !== 0x45786966) return null;
      offset += 6;
      const tiffOffset = offset;
      const little = view.getUint16(tiffOffset, false) === 0x4949;
      const firstIFDOffset = view.getUint32(tiffOffset + 4, little);
      let dirOffset = tiffOffset + firstIFDOffset;
      const entries = view.getUint16(dirOffset, little);
      dirOffset += 2;
      for (let i = 0; i < entries; i++) {
        const entryOffset = dirOffset + i * 12;
        const tag = view.getUint16(entryOffset, little);
        if (tag === 0x0112) { // Orientation
          const val = view.getUint16(entryOffset + 8, little);
          return val;
        }
      }
      return null;
    } else if ((marker & 0xFFF0) !== 0xFFE0) {
      break; // not an APPn marker; stop scanning
    } else {
      const size = view.getUint16(offset, false);
      offset += size;
    }
  }
  return null;
}

async function readFileAsDataURL(file: File, maxWidth = 1200, maxHeight = 900): Promise<string> {
  // Read as ArrayBuffer first so we can detect EXIF orientation
  const buffer = await file.arrayBuffer();
  const orientation = getJpegOrientation(buffer);

  const blob = new Blob([buffer], { type: file.type || "image/jpeg" });
  const objectUrl = URL.createObjectURL(blob);

  const img = await new Promise<HTMLImageElement>((resolve, reject) => {
    const _img = new Image();
    _img.onload = () => resolve(_img);
    _img.onerror = (e) => reject(e);
    _img.src = objectUrl;
  });

  const srcW = img.naturalWidth;
  const srcH = img.naturalHeight;

  // If EXIF orientation is 5–8, the image needs a 90° rotation (dimensions swap)
  const rotates90 = orientation != null && orientation >= 5 && orientation <= 8;

  // Compute scale so that the FINAL displayed orientation fits within maxWidth × maxHeight
  const scale = Math.min(
    maxWidth / (rotates90 ? srcH : srcW),
    maxHeight / (rotates90 ? srcW : srcH),
    1
  );

  // Destination draw size (before canvas rotation transforms)
  const dw = Math.round(srcW * scale);
  const dh = Math.round(srcH * scale);

  // Canvas size depends on whether we rotate 90°
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    URL.revokeObjectURL(objectUrl);
    throw new Error("Canvas not supported");
  }

  if (rotates90) {
    canvas.width = dh;
    canvas.height = dw;
  } else {
    canvas.width = dw;
    canvas.height = dh;
  }

  ctx.save();
  // Apply EXIF orientation transforms
  switch (orientation) {
    case 2: // flip X
      ctx.translate(dw, 0);
      ctx.scale(-1, 1);
      break;
    case 3: // 180°
      ctx.translate(dw, dh);
      ctx.rotate(Math.PI);
      break;
    case 4: // flip Y
      ctx.translate(0, dh);
      ctx.scale(1, -1);
      break;
    case 5: // transpose
      ctx.rotate(0.5 * Math.PI);
      ctx.scale(1, -1);
      ctx.translate(0, -dh);
      break;
    case 6: // rotate 90° CW
      ctx.rotate(0.5 * Math.PI);
      ctx.translate(0, -dh);
      break;
    case 7: // transverse
      ctx.rotate(0.5 * Math.PI);
      ctx.scale(-1, 1);
      ctx.translate(-dw, -dh);
      break;
    case 8: // rotate 270° CW
      ctx.rotate(-0.5 * Math.PI);
      ctx.translate(-dw, 0);
      break;
    default:
      // orientation 1 or unknown: no transform
      break;
  }

  // Draw scaled image
  ctx.drawImage(img, 0, 0, srcW, srcH, 0, 0, dw, dh);
  ctx.restore();

  URL.revokeObjectURL(objectUrl);
  // Export JPEG ~85% quality
  return canvas.toDataURL("image/jpeg", 0.85);
}

// Convert HEIC/HEIF to JPEG if needed
async function ensureJpeg(file: File): Promise<File> {
  const isHeic = /heic|heif/i.test(file.type) || /\.(heic|heif)$/i.test(file.name);
  if (!isHeic) return file;

  const out = await heic2any({ blob: file, toType: "image/jpeg", quality: 0.85 });
  const blob = Array.isArray(out) ? (out[0] as Blob) : (out as Blob);
  return new File([blob], file.name.replace(/\.(heic|heif)$/i, ".jpg"), { type: "image/jpeg" });
}

function safeLoad(): AppState | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const p = JSON.parse(raw);

    const info: AppInfo = {
      date: typeof p?.info?.date === "string" && p.info.date ? p.info.date : todayISO(),
      endDate: typeof p?.info?.endDate === "string" ? p.info.endDate : "",
      multiDay: Boolean(p?.info?.multiDay ?? false),
      location: String(p?.info?.location ?? ""),
      procedure: String(p?.info?.procedure ?? ""),
      workOrder: String(p?.info?.workOrder ?? ""),
      task: String(p?.info?.task ?? ""),
      type: typeof p?.info?.type === "string" ? p.info.type : "",
      workType: typeof p?.info?.workType === "string" ? p.info.workType : "Inspection",
      assetId: typeof p?.info?.assetId === "string" ? p.info.assetId : "",
      station: typeof p?.info?.station === "string" ? p.info.station : "",
      supervisor: typeof p?.info?.supervisor === "string" ? p.info.supervisor : "",
      observer: typeof p?.info?.observer === "string"
        ? p.info.observer
        : (localStorage.getItem(LAST_OBSERVER_KEY) || ""),
    };

    const employees: Employee[] = Array.isArray(p?.employees)
      ? p.employees.map((e: any) => ({
          id: Number(e?.id ?? Date.now()),
          name: String(e?.name ?? "Employee"),
          status: e?.status === "active" || e?.status === "paused" ? e.status : "idle",
          startTime: typeof e?.startTime === "number" ? e.startTime : null,
          elapsedTime: Number(e?.elapsedTime ?? 0),
          pausedAccum: Number(e?.pausedAccum ?? 0),
          lastPausedAt: typeof e?.lastPausedAt === "number" ? e.lastPausedAt : null,
          logs: Array.isArray(e?.logs) ? e.logs.map(String) : [],
          role: typeof e?.role === "string" ? e.role : "",
          skill: typeof e?.skill === "string" ? e.skill : "",
        }))
      : [];

    let taskLog: TaskEntry[] = [];
    if (Array.isArray(p?.taskLog)) {
      if (p.taskLog.length && typeof p.taskLog[0] === "string") {
        const base = Date.now();
        taskLog = (p.taskLog as string[]).map((t, i) => ({
          id: base - (p.taskLog.length - 1 - i),
          at: base - (p.taskLog.length - 1 - i),
          text: t.replace(/^[^:]+:\s*/, ""),
        }));
      } else {
        taskLog = p.taskLog
          .filter((x: any) => x && typeof x.at === "number" && typeof x.text === "string")
          .map((x: any) => ({ id: Number(x.id ?? x.at), at: x.at, text: x.text }));
      }
    }

    let timeLog: TimeLogEntry[] = [];
    if (Array.isArray(p?.timeLog)) {
      timeLog = p.timeLog
        .filter(
          (x: any) =>
            x &&
            typeof x.at === "number" &&
            (x.event === "start" || x.event === "pause" || x.event === "stop" || x.event === "deleted") &&
            typeof x.employeeName === "string",
        )
        .map((x: any) => ({
          id: Number(x.id ?? x.at),
          at: x.at,
          employeeId: typeof x.employeeId === "number" ? x.employeeId : null,
          employeeName: x.employeeName,
          event: x.event as TimeEvent,
          reasonCode: typeof x.reasonCode === "string" ? x.reasonCode : undefined,
          comment: typeof x.comment === "string" ? x.comment : undefined,
        }));
    }

    let photos: PhotoItem[] = [];
    if (Array.isArray(p?.photos)) {
      photos = p.photos
        .filter((ph: any) => ph && typeof ph.dataUrl === "string")
        .map((ph: any) => ({
          id: Number(ph.id ?? Date.now()),
          dataUrl: String(ph.dataUrl),
          name: typeof ph.name === "string" ? ph.name : undefined,
          caption: typeof ph.caption === "string" ? ph.caption : undefined,
        }));
    }

    return { info, employees, taskLog, timeLog, photos };
  } catch {
    return null;
  }
}

function safeSave(state: AppState) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {}
}

/* ---------- Visuals (no external libs) ---------- */
function ProgressBar({ value }: { value: number }) {
  const pct = Math.max(0, Math.min(100, value));
  return (
    <div style={{ background: "#111a34", border: "1px solid #26345a", borderRadius: 12, padding: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6, fontSize: 12, color: "#aabcdf" }}>
        <span>Utilization</span>
        <span>{pct.toFixed(1)}%</span>
      </div>
      <div style={{ height: 14, background: "#0b1228", borderRadius: 999, overflow: "hidden" }}>
        <div style={{ width: `${pct}%`, height: "100%", background: "linear-gradient(90deg,#2fe1a3,#64b5ff)" }} />
      </div>
    </div>
  );
}

function StackedBar({ touchMs, idleMs }: { touchMs: number; idleMs: number }) {
  const total = Math.max(1, touchMs + idleMs);
  const touchPct = (touchMs / total) * 100;
  const idlePct = (idleMs / total) * 100;
  return (
    <div style={{ background: "#111a34", border: "1px solid #26345a", borderRadius: 12, padding: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6, fontSize: 12, color: "#aabcdf" }}>
        <span>Touch vs Idle</span>
        <span>
          {touchPct.toFixed(1)}% / {idlePct.toFixed(1)}%
        </span>
      </div>
      <div style={{ height: 14, background: "#0b1228", borderRadius: 999, overflow: "hidden", display: "flex" }}>
        <div style={{ width: `${touchPct}%`, background: "#35c98e" }} />
        <div style={{ width: `${idlePct}%`, background: "#ffd166" }} />
      </div>
      <div style={{ display: "flex", gap: 12, marginTop: 6, fontSize: 12, color: "#aabcdf" }}>
        <span>
          <span style={{ display: "inline-block", width: 10, height: 10, background: "#35c98e", borderRadius: 2, marginRight: 6 }} />
          Touch {msToHMS(touchMs)}
        </span>
        <span>
          <span style={{ display: "inline-block", width: 10, height: 10, background: "#ffd166", borderRadius: 2, marginRight: 6 }} />
          Idle {msToHMS(idleMs)}
        </span>
      </div>
    </div>
  );
}

/* ---------- Modals ---------- */
function HelpModal({ open, onClose }: { open: boolean; onClose: () => void }) {
  if (!open) return null;
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <header>
          <h3>Help & Tips</h3>
        </header>
        <div className="body">
          <h4>General Info</h4>
          <p>
            Use <b>Type</b> for priority category and <b>Work Type</b> for the nature of work. Enable <b>Multi-day</b> to add End Date. You can
            also set an <b>Observer</b> name; we remember your last entry.
          </p>

          <h4>Time</h4>
          <p>
            <b>Actual Time</b> is wall-clock from first <i>Start</i> to last <i>Stop</i>/<i>Delete</i>; if anyone is still working/paused, it runs
            until now. <b>Touch</b> sums active work; <b>Idle</b> sums paused time.
          </p>

          <h4>Employees</h4>
          <p>Each card has Role and Skill (with “Other…”). Start, Pause, Stop record entries in the Time Log.</p>

          <h4>Task Log</h4>
          <p>Add notes. Sort newest/oldest. Notes can be deleted.</p>

          <h4>Time Log</h4>
          <p>Captures Start/Pause/Stop/Delete with timestamp and optional reason/comment. You can delete single items or all.</p>

          <h4>Exporting</h4>
          <p>CSV &amp; Excel include KPIs, Daily breakdown, and reasons. Export HTML opens a styled report for PDF.</p>

          <h4>Report Printing</h4>
          <p>Use Print Report. In Safari, Export HTML → <i>File → Export as PDF</i>.</p>

          <h4>Clearing Data</h4>
          <p><b>Clear Saved Data</b> wipes everything from this browser only (you’ll confirm first).</p>
        </div>
        <footer>
          <button className="btn" onClick={onClose}>
            Close
          </button>
        </footer>
      </div>
    </div>
  );
}

function ReasonModal({
  open,
  action,
  onCancel,
  onConfirm,
}: {
  open: boolean;
  action: "pause" | "stop";
  onCancel: () => void;
  onConfirm: (reasonCode: string, comment: string) => void;
}) {
  const options = action === "pause" ? PAUSE_REASONS : STOP_REASONS;
  const [reason, setReason] = useState(options[0]);
  const [comment, setComment] = useState("");

  useEffect(() => {
    if (open) {
      setReason(options[0]);
      setComment("");
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, action]);

  if (!open) return null;
  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <header>
          <h3>{action === "pause" ? "Pause reason" : "Stop reason"}</h3>
        </header>
        <div className="body">
          <h4>Reason code</h4>
          <select className="btn" value={reason} onChange={(e) => setReason(e.target.value)}>
            {options.map((r) => (
              <option key={r} value={r}>
                {r}
              </option>
            ))}
          </select>
          <h4>Comment (optional)</h4>
          <textarea
            rows={3}
            style={{
              width: "100%",
              background: "#0b1228",
              color: "var(--ink)",
              border: "1px solid #26345a",
              borderRadius: 10,
              padding: 10,
            }}
            placeholder="Add more detail…"
            value={comment}
            onChange={(e) => setComment(e.target.value)}
          />
        </div>
        <footer>
          <button className="btn" onClick={onCancel}>Cancel</button>
          <button className="btn yellow" onClick={() => onConfirm(reason, comment.trim())}>Confirm</button>
        </footer>
      </div>
    </div>
  );
}

/* ---------- App ---------- */
export default function WorkMeasurementApp() {
  const loaded = safeLoad();
  const [info, setInfo] = useState<AppInfo>(
    loaded?.info ?? {
      date: todayISO(),
      endDate: "",
      multiDay: false,
      location: "",
      procedure: "",
      workOrder: "",
      task: "",
      type: "Routine",
      workType: "Inspection",
      assetId: "",
      station: "",
      supervisor: "",observer: localStorage.getItem(LAST_OBSERVER_KEY) || "",
    },
  );
  const [employees, setEmployees] = useState<Employee[]>(loaded?.employees ?? []);
  const [taskLog, setTaskLog] = useState<TaskEntry[]>(loaded?.taskLog ?? []);
  const [timeLog, setTimeLog] = useState<TimeLogEntry[]>(loaded?.timeLog ?? []);
  const [photos, setPhotos] = useState<PhotoItem[]>(loaded?.photos ?? []);

  const [employeeName, setEmployeeName] = useState("");
  const [note, setNote] = useState("");

  // “Other…” reveal controls for General Info
  const [typeOther, setTypeOther] = useState("");
  const [workTypeOther, setWorkTypeOther] = useState("");

  // Sorting toggles
  const [sortNewestFirst, setSortNewestFirst] = useState(true);
  const [sortTimeNewestFirst, setSortTimeNewestFirst] = useState(true);

  // Help + reason modal state
  const [showHelp, setShowHelp] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const toggleExport = () => setExportOpen((v) => !v);
  const [pendingReason, setPendingReason] = useState<{ action: "pause" | "stop"; id: number } | null>(null);
  const menuRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (!exportOpen) return;
    const onDocClick = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        setExportOpen(false);
      }
    };
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, [exportOpen]);
  
  // Persist
  useEffect(() => {
    safeSave({ info, employees, taskLog, timeLog, photos });
  }, [info, employees, taskLog, timeLog, photos]);
  // --- Photos: handlers ---
  async function handlePhotoFiles(files: FileList | null) {
    if (!files || !files.length) return;

    const remaining = Math.max(0, 5 - photos.length);
    if (remaining === 0) {
      alert("You can attach up to 5 photos.");
      return;
    }

    const toProcess = Array.from(files).slice(0, remaining);
    const additions: PhotoItem[] = [];

    for (const f of toProcess) {
      try {
        const jpgFile = await ensureJpeg(f);
        const dataUrl = await readFileAsDataURL(jpgFile, 1200, 900);
        additions.push({ id: Date.now() + Math.random(), dataUrl, name: jpgFile.name });
      } catch (e) {
        console.error("Failed to process image", f.name, e);
      }
    }

    if (additions.length) {
      setPhotos(prev => [...prev, ...additions]);
    }

    if (files.length > remaining) {
      alert(`Only the first ${remaining} photo(s) were added. Limit is 5 photos total.`);
    }
  }

  function removePhoto(id: number) {
    setPhotos(prev => prev.filter(p => p.id !== id));
  }
  useEffect(() => {
  try {
    if (info.observer && info.observer.trim()) {
      localStorage.setItem(LAST_OBSERVER_KEY, info.observer.trim());
    }
  } catch {}
}, [info.observer]);

  // Ticking
  const [nowMs, setNowMs] = useState(() => Date.now());
  const anyRunningOrPaused = useMemo(
    () => employees.some((e) => e.status === "active" || e.status === "paused"),
    [employees],
  );
  useEffect(() => {
    if (!anyRunningOrPaused) return;
    const t = setInterval(() => setNowMs(Date.now()), 1000);
    return () => clearInterval(t);
  }, [anyRunningOrPaused]);

  // Time math
  const liveTimes = (e: Employee) => {
    const active = e.status === "active" && e.startTime ? e.elapsedTime + (nowMs - e.startTime) : e.elapsedTime;
    const idle = e.status === "paused" && e.lastPausedAt ? e.pausedAccum + (nowMs - e.lastPausedAt) : e.pausedAccum;
    return { active, idle, total: active + idle };
  };

  const appendTimeLog = (entry: Omit<TimeLogEntry, "id" | "at"> & { at?: number }) => {
    setTimeLog((prev) => {
      const at = entry.at ?? Date.now();
      const last = prev[prev.length - 1];
      if (last && last.employeeId === entry.employeeId && last.event === entry.event && at - last.at <= 400) return prev;
      const { employeeId, employeeName, event, reasonCode, comment } = entry;
      return [...prev, { id: at, at, employeeId, employeeName, event, reasonCode, comment }];
    });
  };

  /* ---------- General Info ---------- */
  const handleInfoChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value, type, checked } = e.target as HTMLInputElement;
    if (type === "checkbox" && name === "multiDay") {
      setInfo((prev) => ({ ...prev, multiDay: checked, endDate: checked ? prev.endDate || todayISO() : "" }));
      return;
    }
    setInfo((prev) => ({ ...prev, [name]: value }));
  };

  /* ---------- Employees ---------- */
  const addEmployee = () => {
    const name = employeeName.trim();
    if (!name) return;
    setEmployees((prev) => [
      ...prev,
      {
        id: Date.now(),
        name,
        status: "idle",
        startTime: null,
        elapsedTime: 0,
        pausedAccum: 0,
        lastPausedAt: null,
        logs: [],
        role: "Mechanic",
        skill: "A&P",
      },
    ]);
    setEmployeeName("");
  };

  const deleteEmployee = (id: number) => {
    const emp = employees.find((e) => e.id === id);
    if (!emp) return;
    if (!confirm(`Remove ${emp.name}? Their times will be removed from totals.`)) return;
    appendTimeLog({ employeeId: emp.id, employeeName: emp.name, event: "deleted" });
    setEmployees((prev) => prev.filter((e) => e.id !== id));
  };

  const startTimer = (id: number) => {
    const emp = employees.find((e) => e.id === id);
    if (!emp || emp.status === "active") return;
    appendTimeLog({ employeeId: emp.id, employeeName: emp.name, event: "start" });
    setEmployees((prev) =>
      prev.map((e) => {
        if (e.id !== id) return e;
        let pausedAccum = e.pausedAccum;
        if (e.status === "paused" && e.lastPausedAt) pausedAccum += Date.now() - e.lastPausedAt;
        return {
          ...e,
          status: "active",
          startTime: Date.now(),
          lastPausedAt: null,
          pausedAccum,
          logs: [...e.logs, `Started at ${new Date().toLocaleTimeString()}`],
        };
      }),
    );
  };

  const requestPause = (id: number) => setPendingReason({ action: "pause", id });
  const requestStop = (id: number) => setPendingReason({ action: "stop", id });

  const confirmReason = (reasonCode: string, comment: string) => {
    if (!pendingReason) return;
    const { action, id } = pendingReason;
    const emp = employees.find((e) => e.id === id);
    if (!emp) {
      setPendingReason(null);
      return;
    }
    if (action === "pause") {
      if (emp.status !== "active" || !emp.startTime) {
        setPendingReason(null);
        return;
      }
      appendTimeLog({ employeeId: emp.id, employeeName: emp.name, event: "pause", reasonCode, comment });
      setEmployees((prev) =>
        prev.map((e) =>
          e.id === id
            ? {
                ...e,
                status: "paused",
                elapsedTime: e.elapsedTime + (Date.now() - e.startTime!),
                startTime: null,
                lastPausedAt: Date.now(),
                logs: [
                  ...e.logs,
                  `Paused at ${new Date().toLocaleTimeString()} — ${reasonCode}${comment ? ` (${comment})` : ""}`,
                ],
              }
            : e,
        ),
      );
    } else {
      if (emp.status !== "active" && emp.status !== "paused") {
        setPendingReason(null);
        return;
      }
      appendTimeLog({ employeeId: emp.id, employeeName: emp.name, event: "stop", reasonCode, comment });
      setEmployees((prev) =>
        prev.map((e) => {
          if (e.id !== id) return e;
          if (e.status === "active" && e.startTime) {
            return {
              ...e,
              status: "idle",
              elapsedTime: e.elapsedTime + (Date.now() - e.startTime),
              startTime: null,
              logs: [
                ...e.logs,
                `Stopped at ${new Date().toLocaleTimeString()} — ${reasonCode}${comment ? ` (${comment})` : ""}`,
              ],
            };
          }
          if (e.status === "paused" && e.lastPausedAt) {
            return {
              ...e,
              status: "idle",
              pausedAccum: e.pausedAccum + (Date.now() - e.lastPausedAt),
              lastPausedAt: null,
              logs: [
                ...e.logs,
                `Stopped at ${new Date().toLocaleTimeString()} — ${reasonCode}${comment ? ` (${comment})` : ""}`,
              ],
            };
          }
          return e;
        }),
      );
    }
    setPendingReason(null);
  };

  const cancelReason = () => setPendingReason(null);

  /* ---------- Notes / Logs ---------- */
  const addTaskNote = () => {
    const t = note.trim();
    if (!t) return;
    const at = Date.now();
    setTaskLog((prev) => [...prev, { id: at, at, text: t }]);
    setNote("");
  };
  const deleteTaskNote = (id: number) => {
    if (!confirm("Delete this note?")) return;
    setTaskLog((p) => p.filter((n) => n.id !== id));
  };

  const deleteTimeLogEntry = (id: number) => {
    const entry = timeLog.find((t) => t.id === id);
    if (!entry) return;
    if (!confirm(`Delete this time entry for ${entry.employeeName}?`)) return;
    setTimeLog((prev) => prev.filter((t) => t.id !== id));
  };
  const clearTimeLog = () => {
    if (!timeLog.length) return;
    if (!confirm("Delete ALL time log entries? This cannot be undone.")) return;
    setTimeLog([]);
  };

  const clearSaved = () => {
    if (!confirm("Clear all saved data? This cannot be undone.")) return;
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch {}
    setInfo({
      date: todayISO(),
      endDate: "",
      multiDay: false,
      location: "",
      procedure: "",
      workOrder: "",
      task: "",
      type: "Routine",
      workType: "Inspection",
      assetId: "",
      station: "",
      supervisor: "",
    });
    setEmployees([]);
    setTaskLog([]);
    setTimeLog([]);
  };

  /* ---------- Totals / KPIs ---------- */
  const totalActive = useMemo(
    () =>
      employees.reduce(
        (sum, e) => sum + (e.status === "active" && e.startTime ? e.elapsedTime + (nowMs - e.startTime) : e.elapsedTime),
        0,
      ),
    [employees, nowMs],
  );
  const totalIdle = useMemo(
    () =>
      employees.reduce(
        (sum, e) =>
          sum + (e.status === "paused" && e.lastPausedAt ? e.pausedAccum + (nowMs - e.lastPausedAt) : e.pausedAccum),
        0,
      ),
    [employees, nowMs],
  );
  const totalAll = totalActive + totalIdle;

  const sortedTaskLog = useMemo(
    () => [...taskLog].sort((a, b) => (sortNewestFirst ? b.at - a.at : a.at - b.at)),
    [taskLog, sortNewestFirst],
  );
  const sortedTimeLog = useMemo(
    () => [...timeLog].sort((a, b) => (sortTimeNewestFirst ? b.at - a.at : a.at - b.at)),
    [timeLog, sortTimeNewestFirst],
  );

  const firstStartAt = useMemo(() => {
    const starts = timeLog.filter((t) => t.event === "start").map((t) => t.at);
    return starts.length ? Math.min(...starts) : null;
  }, [timeLog]);
  const lastStopAt = useMemo(() => {
    const ends = timeLog.filter((t) => t.event === "stop" || t.event === "deleted").map((t) => t.at);
    return ends.length ? Math.max(...ends) : null;
  }, [timeLog]);

  const anyEngaged = useMemo(
    () => employees.some((e) => e.status === "active" || e.status === "paused"),
    [employees],
  );
  const actualClockMs = useMemo(() => {
    if (!firstStartAt) return 0;
    const end = anyEngaged ? nowMs : (lastStopAt ?? nowMs);
    return Math.max(0, end - firstStartAt);
  }, [firstStartAt, lastStopAt, nowMs, anyEngaged]);

  const utilization = useMemo(() => (actualClockMs ? totalActive / actualClockMs : 0), [totalActive, actualClockMs]);
  const crewHours = useMemo(() => totalActive / 3_600_000, [totalActive]);
  const idleRatio = useMemo(() => (totalAll ? totalIdle / totalAll : 0), [totalIdle, totalAll]);

  // Daily breakdown
  const dailyBreakdown = useMemo(() => {
    const rows: Record<string, { actualMs: number; touchMs: number; idleMs: number }> = {};
    const events = [...timeLog].sort((a, b) => a.at - b.at);
    if (!events.length) return rows;

    const snapshot: Record<number, EmpStatus> = {};
    const empSet = new Set(events.map((e) => e.employeeId).filter((x): x is number => typeof x === "number"));
    empSet.forEach((id) => (snapshot[id] = "idle"));

    let tPrev = events[0].at;
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
        if (!rows[key]) rows[key] = { actualMs: 0, touchMs: 0, idleMs: 0 };
        rows[key].actualMs += dur;
        rows[key].touchMs += dur * activeCnt;
        rows[key].idleMs += dur * pausedCnt;
        a = b;
      }
    };

    for (const ev of events) {
      addSpan(tPrev, ev.at);
      if (typeof ev.employeeId === "number") {
        if (ev.event === "start") snapshot[ev.employeeId] = "active";
        if (ev.event === "pause") snapshot[ev.employeeId] = "paused";
        if (ev.event === "stop" || ev.event === "deleted") snapshot[ev.employeeId] = "idle";
      }
      tPrev = ev.at;
    }
    if (lastStopAt == null) addSpan(tPrev, nowMs);
    return rows;
  }, [timeLog, lastStopAt, nowMs]);

  /* ---------- Exports ---------- */
  const exportExcel = () => {
    const wb = XLSX.utils.book_new();

    // Summary
    const summaryHeader = [
      "Date",
      "End Date",
      "Type",
      "Work Type",
      "Location",
      "Procedure",
      "Work Order",
      "Task",
      "Asset ID",
      "Station/Area",
      "Supervisor",
      "Observer",
      "Actual Time (H:M:S)",
      "Utilization (%)",
      "Crew-hours",
      "Idle Ratio (%)",
      "Employees",
      "Total Time (H:M:S)",
      "Touch Labor (H:M:S)",
      "Idle Time (H:M:S)",
      "Task Log",
    ];

    const summaryRow = [
      info.date,
      info.multiDay ? info.endDate : "",
      info.type || "",
      info.workType || "",
      info.location,
      info.procedure,
      info.workOrder,
      info.task,
      info.assetId || "",
      info.station || "",
      info.supervisor || "",
      info.observer || "",
      msToHMS(actualClockMs),
      (utilization * 100).toFixed(1),
      crewHours.toFixed(2),
      (idleRatio * 100).toFixed(1),
      employees.map((e) => e.name).join("; "),
      msToTime(totalAll),
      msToTime(totalActive),
      msToTime(totalIdle),
      [...sortedTaskLog].map((n) => `${fmtStamp(n.at, info.multiDay)}: ${n.text}`).join(" | "),
    ];

    const wsSummary = XLSX.utils.aoa_to_sheet([summaryHeader, summaryRow]);
    wsSummary["!cols"] = [
      { wch: 12 }, // Date
      { wch: 12 }, // End Date
      { wch: 16 }, // Type
      { wch: 20 }, // Work Type
      { wch: 20 }, // Location
      { wch: 24 }, // Procedure
      { wch: 18 }, // Work Order
      { wch: 26 }, // Task
      { wch: 14 }, // Asset
      { wch: 16 }, // Station
      { wch: 16 }, // Supervisor
      { wch: 16 }, // Observer
      { wch: 16 }, // Actual
      { wch: 14 }, // Utilization
      { wch: 12 }, // Crew-hours
      { wch: 14 }, // Idle Ratio
      { wch: 28 }, // Employees
      { wch: 16 }, // Total
      { wch: 18 }, // Touch
      { wch: 16 }, // Idle
      { wch: 60 }, // Task Log
    ];
    (wsSummary as any)["!freeze"] = { xSplit: 0, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

    // Employee Performance
    const perfRows = employees.map((e) => {
      const { active, idle, total } = liveTimes(e);
      const sessionsForEmp = timeLog.filter((t) => t.employeeId === e.id && t.event !== "deleted").length;
      const status = e.status === "idle" && (active > 0 || idle > 0) ? "Completed" : e.status;
      return {
        "Employee Name": e.name,
        Role: e.role || "",
        Skill: e.skill || "",
        "Total Active (Touch) (H:M:S)": msToTime(active),
        "Total Idle (H:M:S)": msToTime(idle),
        "Total Time (H:M:S)": msToTime(total),
        Sessions: sessionsForEmp,
        Status: status,
      };
    });
    const wsPerf = XLSX.utils.json_to_sheet(perfRows, {
      header: [
        "Employee Name",
        "Role",
        "Skill",
        "Total Active (Touch) (H:M:S)",
        "Total Idle (H:M:S)",
        "Total Time (H:M:S)",
        "Sessions",
        "Status",
      ],
    });
    wsPerf["!cols"] = [
      { wch: 28 },
      { wch: 16 },
      { wch: 16 },
      { wch: 22 },
      { wch: 18 },
      { wch: 18 },
      { wch: 10 },
      { wch: 14 },
    ];
    (wsPerf as any)["!freeze"] = { xSplit: 0, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, wsPerf, "Employee Performance");

    // Time Log
    const timeRows = [...sortedTimeLog].map((t) => ({
      When: fmtStamp(t.at, info.multiDay),
      Employee: t.employeeName,
      Event: t.event,
      Reason: t.reasonCode || "",
      Comment: t.comment || "",
    }));
    const wsTime = XLSX.utils.json_to_sheet(timeRows, { header: ["When", "Employee", "Event", "Reason", "Comment"] });
    wsTime["!cols"] = [{ wch: 22 }, { wch: 22 }, { wch: 10 }, { wch: 22 }, { wch: 48 }];
    (wsTime as any)["!freeze"] = { xSplit: 0, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, wsTime, "Time Log");

    // Daily Breakdown
    const dayRows = Object.entries(dailyBreakdown)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([date, v]) => ({
        Date: date,
        "Actual (H:M:S)": msToHMS(v.actualMs),
        "Touch (H:M:S)": msToHMS(v.touchMs),
        "Idle (H:M:S)": msToHMS(v.idleMs),
        "Utilization (%)": (v.actualMs ? (v.touchMs / v.actualMs) * 100 : 0).toFixed(1),
      }));
    const wsDaily = XLSX.utils.json_to_sheet(dayRows, {
      header: ["Date", "Actual (H:M:S)", "Touch (H:M:S)", "Idle (H:M:S)", "Utilization (%)"],
    });
    wsDaily["!cols"] = [{ wch: 12 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    (wsDaily as any)["!freeze"] = { xSplit: 0, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, wsDaily, "Daily Breakdown");

    // Metrics Guide
    const guide = [
      { Metric: "Type", Explanation: "Priority category (Routine / Non-Routine / Customer Request / Cannibalization / Other)." },
      { Metric: "Work Type", Explanation: "Nature of work (Inspection / Remove & Replace / Setup / Rework / Test / Remove / Install / Other)." },
      { Metric: "Actual Time", Explanation: "Wall-clock from first Start to last Stop/Delete, or now if anyone is still Active/Paused." },
      { Metric: "Touch Labor", Explanation: "Sum of time employees are Active (crew-weighted)." },
      { Metric: "Idle Time", Explanation: "Sum of time employees are Paused (crew-weighted)." },
      { Metric: "Total Time", Explanation: "Touch + Idle (crew-weighted time, not wall-clock)." },
      { Metric: "Utilization (%)", Explanation: "Touch ÷ Actual × 100." },
      { Metric: "Crew-hours", Explanation: "Touch time converted to hours (Σ Active / 3600s)." },
      { Metric: "Idle Ratio (%)", Explanation: "Idle ÷ (Touch + Idle) × 100." },
      { Metric: "Daily Breakdown", Explanation: "Actual/Touch/Idle apportioned per calendar day." },
    ];
    const wsGuide = XLSX.utils.json_to_sheet(guide, { header: ["Metric", "Explanation"] });
    wsGuide["!cols"] = [{ wch: 22 }, { wch: 70 }];
    XLSX.utils.book_append_sheet(wb, wsGuide, "Metrics Guide");

    XLSX.writeFile(wb, "work_measurement.xlsx");
  };

  const csvEscape = (v: string | number) => {
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const toCSV = (rows: (string | number)[][]) => rows.map((r) => r.map(csvEscape).join(",")).join("\n");
  const download = (filename: string, content: string) => {
    const blob = new Blob([content], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportSummaryCSV = () => {
    const headers = [
      "Date",
      "End Date",
      "Type",
      "Work Type",
      "Location",
      "Procedure",
      "Work Order",
      "Task",
      "Asset ID",
      "Station/Area",
      "Supervisor",
      "Observer",
      "Actual Time (H:M:S)",
      "Utilization (%)",
      "Crew-hours",
      "Idle Ratio (%)",
      "Employees",
      "Total Time (H:M:S)",
      "Touch Labor (H:M:S)",
      "Idle Time (H:M:S)",
      "Task Log",
    ];
    const row = [
      info.date,
      info.multiDay ? info.endDate : "",
      info.type || "",
      info.workType || "",
      info.location,
      info.procedure,
      info.workOrder,
      info.task,
      info.assetId || "",
      info.station || "",
      info.supervisor || "",
      info.observer || "",
      msToHMS(actualClockMs),
      (utilization * 100).toFixed(1),
      crewHours.toFixed(2),
      (idleRatio * 100).toFixed(1),
      employees.map((e) => e.name).join("; "),
      msToTime(totalAll),
      msToTime(totalActive),
      msToTime(totalIdle),
      sortedTaskLog.map((n) => `${fmtStamp(n.at, info.multiDay)}: ${n.text}`).join(" | "),
    ];
    download("work_measurement_summary.csv", toCSV([headers, row]));
  };

  const printReport = () => {
    printReportHTML(info, employees, timeLog, taskLog, liveTimes, msToTime, fmtStamp, photos);
  };

  /* ---------- UI ---------- */
  return (
    <div className="app">
      <header className="app-header">
        <h1 className="app-title">
          <img src="/WorkMeasurmentIcon.png" alt="App icon" className="app-icon" />
          Work Measurement Analysis <span className="badge">v2</span>
        </h1>
        <div className="toolbar">
          <div
            className="export-menu"
            ref={menuRef}
            data-open={exportOpen ? "true" : "false"}
            style={{ position: "relative" }}
          >
            <button className="btn" onClick={toggleExport}>
              Export ▾
            </button>
            <div
              className="menu"
              style={{ display: exportOpen ? "grid" : "none" }}
              onClick={() => setExportOpen(false)}
            >
              <button className="btn ghost" onClick={exportSummaryCSV}>
                CSV (Summary)
              </button>
              <button className="btn ghost" onClick={exportExcel}>
                Excel (Full)
              </button>
              <button
                className="btn ghost"
                onClick={() =>
                  exportReportHTML(info, employees, timeLog, taskLog, liveTimes, msToTime, fmtStamp, photos)
                }
              >
                HTML Report
              </button>
              <button
                className="btn ghost"
                onClick={() =>
                  exportReportPDF(
                    info,
                    employees,
                    timeLog,
                    taskLog,
                    liveTimes,
                    msToTime,
                    fmtStamp,
                    photos
                  )
                }
              >
                PDF (offline)
              </button>
              <button className="btn ghost" onClick={printReport}>
                Print Report
              </button>
            </div>
          </div>
          <button className="btn red" onClick={clearSaved}>
            Clear Saved Data
          </button>
          <button className="help-btn" onClick={() => setShowHelp(true)}>
            Help
          </button>
        </div>
      </header>

<section className="section">
        <h2>Employees <span className="meta">(add each employee involved with the task)</span></h2>
        <div className="grid-auto" style={{ marginTop: 8 }}>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              type="text"
              value={employeeName}
              onChange={(e) => setEmployeeName(e.target.value)}
              placeholder="Employee name"
              style={{ flex: 1 }}
            />
            <button className="btn blue" onClick={addEmployee}>
              Add
            </button>
          </div>
        </div>

        <ul className="card-list">
          {employees.map((emp) => {
            const { active, idle, total } = liveTimes(emp);
            const hasAnyTime = active > 0 || idle > 0 || emp.logs.length > 0;
            const statusClass =
              emp.status === "active"
                ? "emp-active"
                : emp.status === "paused"
                ? "emp-paused"
                : hasAnyTime
                ? "emp-stopped"
                : "emp-neutral";

            const setRole = (v: string) =>
              setEmployees((prev) => prev.map((e) => (e.id === emp.id ? { ...e, role: v } : e)));
            const setSkill = (v: string) =>
              setEmployees((prev) => prev.map((e) => (e.id === emp.id ? { ...e, skill: v } : e)));

            const roleIsPreset = emp.role && ROLE_OPTIONS.includes(emp.role as any);
            const skillIsPreset = emp.skill && SKILL_OPTIONS.includes(emp.skill as any);

            return (
              <li key={emp.id} className={`emp ${statusClass}`}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 16 }}>{emp.name}</div>

                  {/* Role + Skill */}
                  <div style={{ display: "grid", gridTemplateColumns: "minmax(140px, 1fr) minmax(140px, 1fr)", gap: 8, margin: "6px 0 8px" }}>
                    <div>
                      <div className="meta" style={{ marginBottom: 4 }}>
                        Role
                      </div>
                      <select
                        className="btn"
                        value={roleIsPreset ? (emp.role as string) : "__OTHER__"}
                        onChange={(e) => {
                          const v = e.target.value;
                          if (v === "__OTHER__") {
                            setRole("");
                          } else {
                            setRole(v);
                          }
                        }}
                      >
                        {ROLE_OPTIONS.filter(r => r !== "Other…").map((r) => (
                          <option key={r} value={r}>
                            {r}
                          </option>
                        ))}
                        <option value="__OTHER__">Other…</option>
                      </select>
                      {!roleIsPreset && (
                        <input
                          className="other-input"
                          style={{ marginTop: 6, width: "100%" }}
                          placeholder="Role (free text)"
                          value={emp.role || ""}
                          onChange={(e) => setRole(e.target.value)}
                          autoFocus
                        />
                      )}
                    </div>

                    <div>
                      <div className="meta" style={{ marginBottom: 4 }}>
                        Skill
                      </div>
                      <select
                        className="btn"
                        value={skillIsPreset ? (emp.skill as string) : "__OTHER__"}
                        onChange={(e) => {
                          const v = e.target.value;
                          if (v === "__OTHER__") {
                            setSkill("");
                          } else {
                            setSkill(v);
                          }
                        }}
                      >
                        {SKILL_OPTIONS.filter(s => s !== "Other…").map((s) => (
                          <option key={s} value={s}>
                            {s}
                          </option>
                        ))}
                        <option value="__OTHER__">Other…</option>
                      </select>
                      {!skillIsPreset && (
                        <input
                          className="other-input"
                          style={{ marginTop: 6, width: "100%" }}
                          placeholder="Skill (free text)"
                          value={emp.skill || ""}
                          onChange={(e) => setSkill(e.target.value)}
                          autoFocus
                        />
                      )}
                    </div>
                  </div>

                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button
                    className="btn green"
                    onClick={() => startTimer(emp.id)}
                    disabled={emp.status === "active"}
                    aria-disabled={emp.status === "active"}
                    title={emp.status === "active" ? "Already running" : "Start timer"}
                  >
                    Start
                  </button>
                  <button
                    className="btn yellow"
                    onClick={() => requestPause(emp.id)}
                    disabled={emp.status !== "active"}
                    aria-disabled={emp.status !== "active"}
                    title={emp.status !== "active" ? "Nothing to pause" : "Pause timer"}
                  >
                    Pause
                  </button>
                  <button
                    className="btn red"
                    onClick={() => requestStop(emp.id)}
                    disabled={emp.status === "idle"}
                    aria-disabled={emp.status === "idle"}
                    title={emp.status === "idle" ? "Nothing to stop" : "Stop timer"}
                  >
                    Stop
                  </button>
                  <button className="btn ghost" onClick={() => deleteEmployee(emp.id)} title="Remove employee">
                    Remove
                  </button>
                </div>
                <div
                  className="card-status"
                  style={{ flexBasis: "100%" }}
                >
                  <span
                    className={
                      "state " +
                      (emp.status === "active"
                        ? "state-active"
                        : emp.status === "paused"
                        ? "state-paused"
                        : "state-stopped")
                    }
                  >
                    {emp.status === "idle"
                      ? "Stopped"
                      : emp.status === "paused"
                      ? "Paused"
                      : "Active"}
                  </span>
                  <span>Total: {msToTime(total)}</span>
                  <span>Touch: {msToTime(active)}</span>
                  <span>Idle: {msToTime(idle)}</span>
                  <span>Logs: {emp.logs.length}</span>
                </div>
              </li>
            );
          })}
        </ul>
      </section>
      
      {/* PHOTOS SECTION MOVED BELOW TASK LOG */}

      <section className="section card">
        <h2>General Info</h2>

        {/* Big counters */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, margin: "8px 0 12px" }}>
          <div
            style={{
              border: "1px solid #aab2c833",
              borderRadius: 12,
              padding: "10px 14px",
              display: "grid",
              gap: 6,
              justifyItems: "center",
            }}
          >
            <div style={{ fontWeight: 800 }}>ACTUAL TIME</div>
            <div style={{ fontFamily: "ui-monospace, Menlo, Consolas, monospace", fontSize: 36, fontWeight: 800 }}>
              {msToHMS(actualClockMs)}
            </div>
          </div>
          <div
            style={{
              border: "1px solid #aab2c833",
              borderRadius: 12,
              padding: "10px 14px",
              display: "grid",
              gap: 6,
              justifyItems: "center",
            }}
          >
            <div style={{ fontWeight: 800 }}>TOUCH LABOR</div>
            <div style={{ fontFamily: "ui-monospace, Menlo, Consolas, monospace", fontSize: 36, fontWeight: 800 }}>
              {msToHMS(totalActive)}
            </div>
          </div>
          <div
            style={{
              border: "1px solid #aab2c833",
              borderRadius: 12,
              padding: "10px 14px",
              display: "grid",
              gap: 6,
              justifyItems: "center",
            }}
          >
            <div style={{ fontWeight: 800 }}>IDLE TIME</div>
            <div style={{ fontFamily: "ui-monospace, Menlo, Consolas, monospace", fontSize: 36, fontWeight: 800 }}>
              {msToHMS(totalIdle)}
            </div>
          </div>
        </div>

        {/* Visuals */}
        <div style={{ display: "grid", gap: 12, gridTemplateColumns: "1fr 1fr", marginTop: 8 }}>
          <ProgressBar value={utilization * 100} />
          <StackedBar touchMs={totalActive} idleMs={totalIdle} />
        </div>

        {/* Form */}
        <div className="grid-two" style={{ marginTop: 8 }}>
          <label className="stack">
            <span>Start Date</span>
            <input type="date" name="date" value={info.date} onChange={handleInfoChange} />
          </label>
          
          <label className="stack">
            <span>Observer</span>
            <input
            type="text"
            name="observer"
            value={info.observer || ""}
            onChange={handleInfoChange}
            placeholder="e.g., Your name"
            />
            </label>

          <label className="switch">
            <input type="checkbox" name="multiDay" checked={info.multiDay} onChange={handleInfoChange} />
            <span>Multi-day study</span>
          </label>

          {info.multiDay && (
            <label className="stack">
              <span>End Date</span>
              <input type="date" name="endDate" value={info.endDate} min={info.date} onChange={handleInfoChange} />
            </label>
          )}

          <label className="stack">
            <span>Type</span>
            <select
              name="type"
              value={TYPE_OPTIONS.includes((info.type || "") as any) ? info.type : "Other…"}
              onChange={(e) => {
                const v = e.target.value;
                if (v === "Other…") {
                  setTypeOther(info.type && !TYPE_OPTIONS.includes(info.type as any) ? info.type : "");
                  setInfo((prev) => ({ ...prev, type: "" }));
                } else {
                  setTypeOther("");
                  setInfo((prev) => ({ ...prev, type: v }));
                }
              }}
            >
              {TYPE_OPTIONS.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </label>
          {(!info.type || !TYPE_OPTIONS.includes(info.type as any)) && (
            <label className="stack">
              <span>Type — Other</span>
              <input
                type="text"
                className="other-input"
                value={typeOther}
                placeholder="Describe type (e.g., AOG, Special Project)"
                onChange={(e) => {
                  const v = e.target.value;
                  setTypeOther(v);
                  setInfo((prev) => ({ ...prev, type: v }));
                }}
              />
            </label>
          )}

          <label className="stack">
            <span>Work Type</span>
            <select
              name="workType"
              value={WORKTYPE_OPTIONS.includes((info.workType || "") as any) ? info.workType : "Other…"}
              onChange={(e) => {
                const v = e.target.value;
                if (v === "Other…") {
                  setWorkTypeOther(info.workType && !WORKTYPE_OPTIONS.includes(info.workType as any) ? info.workType : "");
                  setInfo((prev) => ({ ...prev, workType: "" }));
                } else {
                  setWorkTypeOther("");
                  setInfo((prev) => ({ ...prev, workType: v }));
                }
              }}
            >
              {WORKTYPE_OPTIONS.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </label>
          {(!info.workType || !WORKTYPE_OPTIONS.includes(info.workType as any)) && (
            <label className="stack">
              <span>Work Type — Other</span>
              <input
                type="text"
                className="other-input"
                value={workTypeOther}
                placeholder="Describe work type"
                onChange={(e) => {
                  const v = e.target.value;
                  setWorkTypeOther(v);
                  setInfo((prev) => ({ ...prev, workType: v }));
                }}
              />
            </label>
          )}

          <label className="stack">
            <span>Location</span>
            <input
              type="text"
              name="location"
              value={info.location}
              onChange={handleInfoChange}
              placeholder="e.g., Hangar 5A; Backshop"
            />
          </label>

          <label className="stack">
            <span>Procedure</span>
            <input
              type="text"
              name="procedure"
              value={info.procedure}
              onChange={handleInfoChange}
              placeholder="e.g., Manual; Repair Procedure"
            />
          </label>

          <label className="stack">
            <span>Work Order</span>
            <input
              type="text"
              name="workOrder"
              value={info.workOrder}
              onChange={handleInfoChange}
              placeholder="e.g., 55555-1000-0001"
            />
          </label>

          <label className="stack">
            <span>Task</span>
            <input
              type="text"
              name="task"
              value={info.task}
              onChange={handleInfoChange}
              placeholder="e.g., Landing Gear Removal; Screw Extraction"
            />
          </label>

          <label className="stack">
            <span>Asset ID</span>
            <input type="text" name="assetId" value={info.assetId || ""} onChange={handleInfoChange} placeholder="e.g., Tail # / Serial" />
          </label>

          <label className="stack">
            <span>Station/Area</span>
            <input type="text" name="station" value={info.station || ""} onChange={handleInfoChange} placeholder="e.g., Bay 12 / Area B" />
          </label>

          <label className="stack">
            <span>Supervisor</span>
            <input type="text" name="supervisor" value={info.supervisor || ""} onChange={handleInfoChange} placeholder="e.g., J. Smith" />
          </label>
        </div>

        <div className="kpis" style={{ marginTop: 12 }}>
          <div className="kpi">
            <div className="label">Total Employees</div>
            <div className="num">{employees.length}</div>
          </div>
          <div className="kpi">
            <div className="label">Total Sessions</div>
            <div className="num">{timeLog.filter((t) => t.event !== "deleted").length}</div>
          </div>
          <div className="kpi">
            <div className="label">Combined Time</div>
            <div className="num">{msToTime(totalAll)}</div>
          </div>
          <div className="kpi">
            <div className="label">Utilization</div>
            <div className="num">{(utilization * 100).toFixed(1)}%</div>
          </div>
          <div className="kpi">
            <div className="label">Crew-hours</div>
            <div className="num">{crewHours.toFixed(2)}</div>
          </div>
          <div className="kpi">
            <div className="label">Idle Ratio</div>
            <div className="num">{(idleRatio * 100).toFixed(1)}%</div>
          </div>
        </div>
      </section>

      

      <section className="section card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <h2>Task Log</h2>
          <button
            className="sort-toggle"
            onClick={() => setSortNewestFirst((v) => !v)}
            title={sortNewestFirst ? "Newest → Oldest" : "Oldest → Newest"}
          >
            {sortNewestFirst ? "▼ Newest first" : "▲ Oldest first"}
          </button>
        </div>

        <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
          <input
            type="text"
            value={note}
            onChange={(e) => setNote(e.target.value)}
            placeholder="Add note"
            style={{ flex: 1 }}
          />
          <button className="btn blue" onClick={addTaskNote}>
            Add
          </button>
        </div>

        <ul className="grid-auto" style={{ marginTop: 10 }}>
          {sortedTaskLog.map((n) => (
            <li key={n.id} style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ whiteSpace: "pre-wrap", flex: 1 }}>
                {fmtStamp(n.at, info.multiDay)}: {n.text}
              </span>
              <button className="btn ghost" onClick={() => deleteTaskNote(n.id)} title="Delete note">
                Delete
              </button>
            </li>
          ))}
          {sortedTaskLog.length === 0 && <li className="meta">(no notes)</li>}
        </ul>
      </section>

      {/* PHOTOS SECTION MOVED HERE */}
      <section className="section card">
        <h2>Photos <span className="meta">({photos.length}/5)</span></h2>
        <div style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 8, flexWrap: "wrap" }}>
          <label className="btn blue" style={{ cursor: "pointer" }}>
            Upload Photos
            <input
              type="file"
              accept="image/*,.heic,.heif"
              multiple
              onChange={(e) => handlePhotoFiles(e.target.files)}
              style={{ display: "none" }}
            />
          </label>
          <span className="meta">Images are resized and auto-rotated on upload.</span>
        </div>

        {photos.length > 0 && (
          <ul className="card-list" style={{ gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", display: "grid" }}>
            {photos.map((p) => (
              <li key={p.id} className="emp" style={{ padding: 8, display: "grid", gap: 8 }}>
                <div className="photo-thumb">
                  <img src={p.dataUrl} alt={p.name || "photo"} />
                  <button
                    type="button"
                    className="photo-remove"
                    aria-label="Remove photo"
                    title="Remove photo"
                    onClick={() => removePhoto(p.id)}
                  >
                    ×
                  </button>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span className="meta" style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {p.name || "photo"}
                  </span>
                </div>
              </li>
            ))}
          </ul>
        )}
      </section>

      <section className="section card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <h2>Time Log</h2>
          <div style={{ display: "flex", gap: 8 }}>
            <button
              className="sort-toggle"
              onClick={() => setSortTimeNewestFirst((v) => !v)}
              title={sortTimeNewestFirst ? "Newest → Oldest" : "Oldest → Newest"}
            >
              {sortTimeNewestFirst ? "▼ Newest first" : "▲ Oldest first"}
            </button>
            <button className="btn ghost" onClick={clearTimeLog} title="Delete all time log entries">
              Delete All
            </button>
          </div>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>When</th>
                <th>Employee</th>
                <th>Event</th>
                <th>Reason</th>
                <th>Comment</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {sortedTimeLog.map((row) => (
                <tr key={row.id}>
                  <td className="mono">{fmtStamp(row.at, info.multiDay)}</td>
                  <td>{row.employeeName}</td>
                  <td className="cap">{row.event}</td>
                  <td>{row.reasonCode || ""}</td>
                  <td style={{ whiteSpace: "pre-wrap" }}>{row.comment || ""}</td>
                  <td>
                    <button className="btn ghost" onClick={() => deleteTimeLogEntry(row.id)} title="Delete this entry">
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
              {sortedTimeLog.length === 0 && (
                <tr>
                  <td colSpan={6} className="meta">
                    No time log entries.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Daily breakdown table */}
        {Object.keys(dailyBreakdown).length > 0 && (
          <div style={{ marginTop: 12 }}>
            <h3 style={{ margin: "6px 0 8px" }}>Daily Breakdown</h3>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>Actual</th>
                    <th>Touch</th>
                    <th>Idle</th>
                    <th>Utilization</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(dailyBreakdown)
                    .sort(([a], [b]) => a.localeCompare(b))
                    .map(([date, v]) => (
                      <tr key={date}>
                        <td className="mono">{date}</td>
                        <td className="mono">{msToHMS(v.actualMs)}</td>
                        <td className="mono">{msToHMS(v.touchMs)}</td>
                        <td className="mono">{msToHMS(v.idleMs)}</td>
                        <td className="mono">{v.actualMs ? ((v.touchMs / v.actualMs) * 100).toFixed(1) : "0.0"}%</td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <p className="footer-hint">Tip: Export HTML and “Save as PDF” if Safari blocks direct print.</p>
      </section>


      <HelpModal open={showHelp} onClose={() => setShowHelp(false)} />
      <ReasonModal
        open={!!pendingReason}
        action={pendingReason?.action ?? "pause"}
        onCancel={cancelReason}
        onConfirm={confirmReason}
      />
    </div>
  );
}