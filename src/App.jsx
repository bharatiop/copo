import React, { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { Upload, Plus, Trash2, ChevronRight, ChevronLeft, LogOut, LayoutGrid, BookOpen, Grid3x3, ClipboardList, FileBarChart2, Building2, Save, X, AlertTriangle, CheckCircle2, Printer } from "lucide-react";
import { sget, sset } from "./storage.js";

/* ---------------------------------------------------------------
   TOKENS
   Ink:      #14213D  (institute navy — authority, letterhead)
   Oxblood:  #7A2E2E  (accreditation seal accent)
   Brass:    #A9812F  (seal ring / gold foil)
   Paper:    #F7F5F0  (ledger paper)
   Ink-soft: #48506B  (secondary text)
   Good:     #2F6F52
   Warn:     #B4762B
   Bad:      #A23B3B
----------------------------------------------------------------*/

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,500;8..60,600;8..60,700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');`;

const DEFAULT_POS = [
  "Engineering knowledge", "Problem analysis", "Design/development of solutions",
  "Conduct investigations of complex problems", "Modern tool usage", "The engineer and society",
  "Environment and sustainability", "Ethics", "Individual and team work",
  "Communication", "Project management and finance", "Life-long learning",
];
const DEFAULT_PSOS = ["Domain-specific specialization", "Application to real-world systems"];

const BL_LEVELS = ["Remember", "Understand", "Apply", "Analyze", "Evaluate", "Create"];
const uid = () => Math.random().toString(36).slice(2, 10);

const LEVEL_BANDS_DEFAULT = [
  { level: 3, min: 80 },
  { level: 2, min: 70 },
  { level: 1, min: 60 },
  { level: 0, min: 0 },
];
// Model-2 (NITR, average-class-score method) bands
const AVG_BANDS_DEFAULT = { high: 70, med: 50 }; // >=high -> 3, >=med -> 2, else -> 1

const DEFAULT_SETTINGS = {
  method: "crossing",        // "crossing" (Model 1: % of students crossing target) or "average" (Model 2/NITR: class average score)
  targetInternal: 60,        // Model 1 default: 60% target for internal assessments
  targetExternal: 40,        // Model 1 default: 40% target for the end-semester exam
  bands: LEVEL_BANDS_DEFAULT,
  avgBands: AVG_BANDS_DEFAULT,
  internalWeight: 40,        // final CO blend: 40% internal / 60% external (Model 1). Use 50/50 for Model 2.
  externalWeight: 60,
  directWeight: 80,          // final CO blend: 80% direct / 20% indirect (course-end survey) — both models agree here
  indirectWeight: 20,
  indirectByCo: {},          // { [coId]: percent of students rating >=3 "good" on the course-end survey }
};

// Shared calculation used by both the per-course report and the admin rollup, so the
// two screens can never drift apart on methodology.
function computeFinalCoAttainments(cos, components, marksByComponent, settings) {
  const s = { ...DEFAULT_SETTINGS, ...settings };
  const out = {};
  cos.forEach((co) => {
    const groups = { internal: { sum: 0, w: 0 }, external: { sum: 0, w: 0 } };
    components.forEach((comp) => {
      const qs = comp.questions.filter((q) => q.co === co.id);
      if (qs.length === 0) return;
      const target = comp.targetPct != null ? comp.targetPct : (comp.type === "external" ? s.targetExternal : s.targetInternal);
      const res = computeCoAttainment(qs, marksByComponent[comp.id] || {}, { method: s.method, targetPct: target, bands: s.bands, avgBands: s.avgBands });
      if (!res) return;
      const g = comp.type === "external" ? groups.external : groups.internal;
      g.sum += res.level * (comp.weight || 100);
      g.w += (comp.weight || 100);
    });
    const intAvg = groups.internal.w > 0 ? groups.internal.sum / groups.internal.w : null;
    const extAvg = groups.external.w > 0 ? groups.external.sum / groups.external.w : null;
    let directLevel = null;
    if (intAvg != null && extAvg != null) {
      directLevel = (intAvg * s.internalWeight + extAvg * s.externalWeight) / (s.internalWeight + s.externalWeight);
    } else if (intAvg != null) directLevel = intAvg;
    else if (extAvg != null) directLevel = extAvg;
    if (directLevel == null) { out[co.id] = null; return; }

    const indirectPct = s.indirectByCo ? s.indirectByCo[co.id] : null;
    let finalLevel = directLevel;
    let indirectLevel = null;
    if (indirectPct != null && indirectPct !== "") {
      const band = s.bands.find((b) => Number(indirectPct) >= b.min);
      indirectLevel = band ? band.level : 0;
      finalLevel = (directLevel * s.directWeight + indirectLevel * s.indirectWeight) / (s.directWeight + s.indirectWeight);
    }
    out[co.id] = { level: +finalLevel.toFixed(2), directLevel: +directLevel.toFixed(2), indirectLevel };
  });
  return out;
}
// Method A ("crossing"): % of students crossing a target score -> banded 0-3 (60/70/80% of class)
// Method B ("average"): the class's average score itself -> banded 0-3 (NITR Model-2, target bands 50/70%)
function computeCoAttainment(questions, marksByRoll, opts) {
  const { method = "crossing", targetPct = 60, bands = LEVEL_BANDS_DEFAULT, avgBands = AVG_BANDS_DEFAULT } = opts || {};
  const rolls = Object.keys(marksByRoll);
  if (rolls.length === 0 || questions.length === 0) return null;
  const maxTotal = questions.reduce((s, q) => s + Number(q.max || 0), 0);
  if (maxTotal === 0) return null;

  const pcts = rolls.map((r) => {
    const rec = marksByRoll[r] || {};
    const got = questions.reduce((s, q) => s + Number(rec[q.id] || 0), 0);
    return (got / maxTotal) * 100;
  });

  if (method === "average") {
    const avg = pcts.reduce((s, p) => s + p, 0) / pcts.length;
    const level = avg >= avgBands.high ? 3 : avg >= avgBands.med ? 2 : 1;
    return { level, avgPct: avg.toFixed(1), studentsConsidered: rolls.length, method };
  }
  // default: students-crossing-target method
  const meeting = pcts.filter((p) => p >= targetPct).length;
  const pctMeeting = (meeting / rolls.length) * 100;
  const band = bands.find((b) => pctMeeting >= b.min);
  return { level: band ? band.level : 0, pctMeeting: pctMeeting.toFixed(1), studentsConsidered: rolls.length, method };
}

// Official formula (OBE training deck, Model 1 & Model 2 both use this):
// PO attainment = Σ (CO attainment level × mapping strength) / (3 × count of non-zero mapping strengths)
function poAttainmentFromCos(coAttainments, matrixRow) {
  let num = 0, count = 0;
  Object.entries(matrixRow || {}).forEach(([coId, corr]) => {
    const c = corr || 0;
    if (c > 0 && coAttainments[coId] != null) {
      num += coAttainments[coId].level * c;
      count += 1;
    }
  });
  if (count === 0) return null;
  return +(num / (3 * count)).toFixed(2);
}

// ================================================================
export default function CoPoApp() {
  const [ready, setReady] = useState(false);
  const [user, setUser] = useState(null); // {name, role, dept}
  const [pos, setPos] = useState(DEFAULT_POS);
  const [psos, setPsos] = useState(DEFAULT_PSOS);
  const [courses, setCourses] = useState([]);
  const [view, setView] = useState("dashboard");
  const [activeCourseId, setActiveCourseId] = useState(null);
  const [toast, setToast] = useState(null);

  useEffect(() => {
    (async () => {
      const u = await sget("app:lastuser", null);
      const p = await sget("app:pos", DEFAULT_POS);
      const ps = await sget("app:psos", DEFAULT_PSOS);
      const c = await sget("app:courses", []);
      setPos(p); setPsos(ps); setCourses(c);
      if (u) setUser(u);
      setReady(true);
    })();
  }, []);

  const flash = useCallback((msg, kind = "good") => {
    setToast({ msg, kind });
    setTimeout(() => setToast(null), 2600);
  }, []);

  const login = async (u) => {
    setUser(u);
    await sset("app:lastuser", u);
    const allUsers = await sget("app:users", []);
    if (!allUsers.find((x) => x.name === u.name && x.role === u.role)) {
      await sset("app:users", [...allUsers, u]);
    }
  };
  const logout = async () => { setUser(null); await sset("app:lastuser", null); setView("dashboard"); };

  const saveCourses = async (next) => { setCourses(next); await sset("app:courses", next); };
  const savePos = async (next) => { setPos(next); await sset("app:pos", next); };
  const savePsos = async (next) => { setPsos(next); await sset("app:psos", next); };

  if (!ready) return <Shell><LoadingScreen /></Shell>;
  if (!user) return <Shell><Login onLogin={login} /></Shell>;

  const activeCourse = courses.find((c) => c.id === activeCourseId);
  const visibleCourses = user.role === "admin" ? courses : courses.filter((c) => c.faculty === user.name);

  return (
    <Shell>
      <div style={{ display: "flex", height: "100%", minHeight: 640 }}>
        <Sidebar
          user={user} view={view} setView={setView} logout={logout}
          courses={visibleCourses} activeCourseId={activeCourseId}
          setActiveCourseId={(id) => { setActiveCourseId(id); setView("course"); }}
        />
        <main style={{ flex: 1, overflow: "auto", background: "var(--paper)" }}>
          {toast && <Toast msg={toast.msg} kind={toast.kind} />}
          {view === "dashboard" && (
            <Dashboard user={user} courses={visibleCourses} pos={pos} psos={psos} />
          )}
          {view === "pos" && user.role === "admin" && (
            <PoEditor pos={pos} psos={psos} savePos={savePos} savePsos={savePsos} flash={flash} />
          )}
          {view === "courses" && (
            <CourseList
              courses={visibleCourses} allCourses={courses} user={user}
              saveCourses={saveCourses} openCourse={(id) => { setActiveCourseId(id); setView("course"); }}
              flash={flash}
            />
          )}
          {view === "admin-report" && user.role === "admin" && (
            <AdminReport courses={courses} pos={pos} psos={psos} />
          )}
          {view === "course" && activeCourse && (
            <CourseWorkspace
              key={activeCourse.id}
              course={activeCourse}
              pos={pos} psos={psos}
              onUpdateCourse={(patch) => saveCourses(courses.map((c) => c.id === activeCourse.id ? { ...c, ...patch } : c))}
              flash={flash}
            />
          )}
        </main>
      </div>
    </Shell>
  );
}

// ---------------- Shell / chrome ----------------
function Shell({ children }) {
  return (
    <div style={{
      fontFamily: "'IBM Plex Sans', sans-serif", color: "var(--ink)", height: "100vh",
      "--ink": "#14213D", "--ink-soft": "#48506B", "--oxblood": "#7A2E2E", "--brass": "#A9812F",
      "--paper": "#F7F5F0", "--paper-deep": "#EFEBE0", "--good": "#2F6F52", "--warn": "#B4762B", "--bad": "#A23B3B",
      "--line": "#DCD6C6",
    }}>
      <style>{`
        ${FONT_IMPORT}
        * { box-sizing: border-box; }
        h1,h2,h3,.serif { font-family: 'Source Serif 4', serif; }
        .mono { font-family: 'IBM Plex Mono', monospace; }
        input, select, button, textarea { font-family: inherit; }
        input, select, textarea {
          border: 1px solid var(--line); border-radius: 3px; padding: 7px 9px; font-size: 13.5px;
          background: #fff; color: var(--ink); outline: none;
        }
        input:focus, select:focus, textarea:focus { border-color: var(--oxblood); box-shadow: 0 0 0 3px rgba(122,46,46,0.08); }
        button { cursor: pointer; }
        button:focus-visible, a:focus-visible, [tabindex]:focus-visible { outline: 2px solid var(--brass); outline-offset: 2px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid var(--line); padding: 6px 8px; font-size: 12.5px; text-align: center; }
        th { background: var(--paper-deep); font-weight: 600; color: var(--ink); }
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-thumb { background: #D8D2C0; border-radius: 6px; }
        @media print {
          .no-print { display: none !important; }
          main { overflow: visible !important; }
        }
      `}</style>
      {children}
    </div>
  );
}

function LoadingScreen() {
  return <div style={{ display: "grid", placeItems: "center", height: "100vh", color: "var(--ink-soft)" }}>Loading register…</div>;
}

function Toast({ msg, kind }) {
  const color = kind === "good" ? "var(--good)" : kind === "bad" ? "var(--bad)" : "var(--warn)";
  return (
    <div style={{
      position: "sticky", top: 0, zIndex: 20, margin: "10px 16px 0", padding: "9px 14px",
      background: "#fff", border: `1px solid ${color}`, borderLeft: `4px solid ${color}`,
      borderRadius: 3, fontSize: 13, color: "var(--ink)", boxShadow: "0 2px 8px rgba(20,33,61,0.08)",
    }}>{msg}</div>
  );
}

// ---------------- Seal (signature element) ----------------
function AttainmentSeal({ pct, label = "Overall", size = 92 }) {
  const r = 40, c = 2 * Math.PI * r;
  const val = Math.max(0, Math.min(100, pct || 0));
  const dash = (val / 100) * c;
  const color = val >= 70 ? "var(--good)" : val >= 50 ? "var(--warn)" : "var(--bad)";
  return (
    <svg width={size} height={size} viewBox="0 0 100 100">
      <circle cx="50" cy="50" r="46" fill="none" stroke="var(--brass)" strokeWidth="1" strokeDasharray="1.5 2.2" opacity="0.6" />
      <circle cx="50" cy="50" r={r} fill="none" stroke="var(--paper-deep)" strokeWidth="7" />
      <circle cx="50" cy="50" r={r} fill="none" stroke={color} strokeWidth="7" strokeLinecap="round"
        strokeDasharray={`${dash} ${c}`} transform="rotate(-90 50 50)" />
      <text x="50" y="48" textAnchor="middle" fontSize="19" fontWeight="700" fontFamily="'Source Serif 4', serif" fill="var(--ink)">{val.toFixed(0)}%</text>
      <text x="50" y="63" textAnchor="middle" fontSize="7.5" fontFamily="'IBM Plex Sans', sans-serif" fill="var(--ink-soft)">{label}</text>
    </svg>
  );
}

// ---------------- Login ----------------
function Login({ onLogin }) {
  const [name, setName] = useState("");
  const [dept, setDept] = useState("");
  const [role, setRole] = useState("faculty");
  return (
    <div style={{ display: "grid", placeItems: "center", height: "100vh", background: "var(--ink)" }}>
      <div style={{ width: 400, background: "var(--paper)", borderRadius: 4, padding: "34px 32px", boxShadow: "0 20px 60px rgba(0,0,0,0.35)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
          <Building2 size={22} color="var(--oxblood)" />
          <div className="serif" style={{ fontSize: 21, fontWeight: 700 }}>Attainment Register</div>
        </div>
        <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginBottom: 24, letterSpacing: 0.3 }}>
          CO–PO–PSO mapping &amp; attainment, department-wide
        </div>
        <label style={{ fontSize: 12, color: "var(--ink-soft)", display: "block", marginBottom: 4 }}>Your name</label>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Dr. S. Mishra" style={{ width: "100%", marginBottom: 14 }} />
        <label style={{ fontSize: 12, color: "var(--ink-soft)", display: "block", marginBottom: 4 }}>Department</label>
        <input value={dept} onChange={(e) => setDept(e.target.value)} placeholder="e.g. Electronics & Comm. Engg." style={{ width: "100%", marginBottom: 14 }} />
        <label style={{ fontSize: 12, color: "var(--ink-soft)", display: "block", marginBottom: 6 }}>Sign in as</label>
        <div style={{ display: "flex", gap: 8, marginBottom: 22 }}>
          {["faculty", "admin"].map((r) => (
            <button key={r} onClick={() => setRole(r)} style={{
              flex: 1, padding: "9px 0", borderRadius: 3, border: `1px solid ${role === r ? "var(--oxblood)" : "var(--line)"}`,
              background: role === r ? "var(--oxblood)" : "#fff", color: role === r ? "#fff" : "var(--ink)", fontSize: 13, fontWeight: 600, textTransform: "capitalize",
            }}>{r === "admin" ? "Department Admin" : "Faculty"}</button>
          ))}
        </div>
        <button
          disabled={!name.trim()}
          onClick={() => onLogin({ name: name.trim(), dept: dept.trim() || "General", role })}
          style={{
            width: "100%", padding: "11px 0", background: name.trim() ? "var(--ink)" : "#B9B9B9", color: "#fff",
            border: "none", borderRadius: 3, fontSize: 14, fontWeight: 600, letterSpacing: 0.3,
          }}>Enter</button>
        <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 16, lineHeight: 1.5 }}>
          Note: this is a lightweight name-based sign-in for a department team, not a secured institutional login.
        </div>
      </div>
    </div>
  );
}

// ---------------- Sidebar ----------------
function Sidebar({ user, view, setView, logout, courses, activeCourseId, setActiveCourseId }) {
  const NavBtn = ({ id, icon: Icon, label }) => (
    <button onClick={() => setView(id)} style={{
      display: "flex", alignItems: "center", gap: 9, width: "100%", padding: "9px 14px",
      background: view === id ? "rgba(255,255,255,0.09)" : "transparent", border: "none",
      borderLeft: view === id ? "3px solid var(--brass)" : "3px solid transparent",
      color: view === id ? "#fff" : "#C6CCDC", fontSize: 13.5, textAlign: "left",
    }}><Icon size={15} />{label}</button>
  );
  return (
    <aside className="no-print" style={{ width: 240, background: "var(--ink)", color: "#fff", display: "flex", flexDirection: "column", flexShrink: 0 }}>
      <div style={{ padding: "20px 16px 14px", borderBottom: "1px solid rgba(255,255,255,0.1)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <Building2 size={17} color="var(--brass)" />
          <span className="serif" style={{ fontWeight: 700, fontSize: 15.5 }}>Attainment Register</span>
        </div>
        <div style={{ fontSize: 11.5, color: "#9BA3BD", marginTop: 6 }}>{user.name} · <span style={{ textTransform: "capitalize" }}>{user.role}</span></div>
        <div style={{ fontSize: 11, color: "#7C86A6" }}>{user.dept}</div>
      </div>
      <nav style={{ padding: "10px 0", flex: "0 0 auto" }}>
        <NavBtn id="dashboard" icon={LayoutGrid} label="Dashboard" />
        <NavBtn id="courses" icon={BookOpen} label="Courses" />
        {user.role === "admin" && <NavBtn id="pos" icon={Grid3x3} label="Program Outcomes" />}
        {user.role === "admin" && <NavBtn id="admin-report" icon={FileBarChart2} label="Program Attainment" />}
      </nav>
      <div style={{ padding: "6px 14px", fontSize: 10.5, letterSpacing: 0.8, color: "#7C86A6", marginTop: 8 }}>YOUR COURSES</div>
      <div style={{ overflow: "auto", flex: 1 }}>
        {courses.length === 0 && <div style={{ padding: "6px 16px", fontSize: 12, color: "#7C86A6" }}>None yet</div>}
        {courses.map((c) => (
          <button key={c.id} onClick={() => setActiveCourseId(c.id)} style={{
            display: "block", width: "100%", textAlign: "left", padding: "8px 16px",
            background: activeCourseId === c.id ? "rgba(255,255,255,0.09)" : "transparent",
            borderLeft: activeCourseId === c.id ? "3px solid var(--brass)" : "3px solid transparent",
            border: "none", color: "#DADEEA", fontSize: 12.5,
          }}>
            <div style={{ fontWeight: 600 }}>{c.code || "—"}</div>
            <div style={{ fontSize: 11, color: "#9BA3BD", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c.name}</div>
          </button>
        ))}
      </div>
      <button onClick={logout} style={{
        display: "flex", alignItems: "center", gap: 8, padding: "13px 16px", background: "transparent",
        border: "none", borderTop: "1px solid rgba(255,255,255,0.1)", color: "#C6CCDC", fontSize: 13,
      }}><LogOut size={14} /> Sign out</button>
    </aside>
  );
}

// ---------------- Dashboard ----------------
function Dashboard({ user, courses, pos, psos }) {
  return (
    <div style={{ padding: "28px 34px" }}>
      <div className="serif" style={{ fontSize: 24, fontWeight: 700 }}>Welcome, {user.name.split(" ")[0]}</div>
      <div style={{ color: "var(--ink-soft)", fontSize: 13.5, marginTop: 4, maxWidth: 620 }}>
        Upload marks, tag questions to Course Outcomes and Bloom's levels, and let the register compute CO and PO/PSO attainment automatically — the way NBA/OBE accreditation expects it.
      </div>
      <div style={{ display: "flex", gap: 18, marginTop: 26, flexWrap: "wrap" }}>
        <StatCard label="Courses" value={courses.length} />
        <StatCard label="Program Outcomes" value={pos.length} />
        <StatCard label="Program-Specific Outcomes" value={psos.length} />
      </div>
      <div className="serif" style={{ fontSize: 16, fontWeight: 700, marginTop: 32, marginBottom: 10 }}>How the numbers are computed</div>
      <ol style={{ fontSize: 13, color: "var(--ink-soft)", lineHeight: 1.9, paddingLeft: 18, maxWidth: 680 }}>
        <li>Each question in a test/assignment is tagged with a <b>CO</b> and a <b>Bloom's level</b>, with a max mark.</li>
        <li>You upload the marks each student scored per question (Excel/CSV).</li>
        <li><b>CO attainment (direct)</b> — pick either official method per course: <b>Model 1</b>, % of students crossing a target score (Level 1: 60–69%, Level 2: 70–79%, Level 3: ≥80%, internal target usually 60%, external/end-sem 40%); or <b>Model 2</b>, the class's average score banded (Low &lt;50%, Medium 50–69%, High ≥70%). Internal and external components are then blended by weight (40/60 or 50/50).</li>
        <li>An optional <b>course-end survey</b> (indirect) can blend in at 20%, alongside 80% direct.</li>
        <li><b>PO/PSO attainment</b> = Σ(CO level × mapping strength) ÷ (3 × number of non-zero mappings) — the exact equation used in NBA/OBE accreditation training.</li>
      </ol>
    </div>
  );
}
function StatCard({ label, value }) {
  return (
    <div style={{ background: "#fff", border: "1px solid var(--line)", borderRadius: 4, padding: "16px 22px", minWidth: 150 }}>
      <div className="serif" style={{ fontSize: 26, fontWeight: 700 }}>{value}</div>
      <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{label}</div>
    </div>
  );
}

// ---------------- PO editor ----------------
function PoEditor({ pos, psos, savePos, savePsos, flash }) {
  const [localPos, setLocalPos] = useState(pos);
  const [localPsos, setLocalPsos] = useState(psos);
  useEffect(() => { setLocalPos(pos); }, [pos]);
  useEffect(() => { setLocalPsos(psos); }, [psos]);
  return (
    <div style={{ padding: "28px 34px", maxWidth: 780 }}>
      <div className="serif" style={{ fontSize: 22, fontWeight: 700 }}>Program Outcomes &amp; PSOs</div>
      <div style={{ color: "var(--ink-soft)", fontSize: 13, marginTop: 4, marginBottom: 20 }}>Shared across every course in the department.</div>
      <ListEditor
        title="Program Outcomes (POs)" prefix="PO" items={localPos} setItems={setLocalPos}
      />
      <div style={{ height: 22 }} />
      <ListEditor
        title="Program-Specific Outcomes (PSOs)" prefix="PSO" items={localPsos} setItems={setLocalPsos}
      />
      <button onClick={async () => { await savePos(localPos); await savePsos(localPsos); flash("Program outcomes saved."); }} style={{
        marginTop: 20, display: "flex", alignItems: "center", gap: 7, padding: "9px 18px", background: "var(--ink)",
        color: "#fff", border: "none", borderRadius: 3, fontSize: 13, fontWeight: 600,
      }}><Save size={14} /> Save</button>
    </div>
  );
}
function ListEditor({ title, prefix, items, setItems }) {
  return (
    <div>
      <div style={{ fontWeight: 600, fontSize: 13.5, marginBottom: 8 }}>{title}</div>
      {items.map((it, i) => (
        <div key={i} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "center" }}>
          <div className="mono" style={{ width: 46, fontSize: 12, color: "var(--ink-soft)" }}>{prefix}{i + 1}</div>
          <input value={it} onChange={(e) => setItems(items.map((x, j) => j === i ? e.target.value : x))} style={{ flex: 1 }} />
          <button onClick={() => setItems(items.filter((_, j) => j !== i))} style={{ background: "none", border: "none", color: "var(--bad)" }}><Trash2 size={14} /></button>
        </div>
      ))}
      <button onClick={() => setItems([...items, ""])} style={{
        display: "flex", alignItems: "center", gap: 5, marginTop: 4, background: "none", border: "1px dashed var(--line)",
        borderRadius: 3, padding: "6px 10px", fontSize: 12.5, color: "var(--ink-soft)",
      }}><Plus size={13} /> Add {prefix}</button>
    </div>
  );
}

// ---------------- Course list ----------------
function CourseList({ courses, allCourses, user, saveCourses, openCourse, flash }) {
  const [showNew, setShowNew] = useState(false);
  const [name, setName] = useState(""); const [code, setCode] = useState("");
  const [faculty, setFaculty] = useState(user.role === "faculty" ? user.name : "");
  const [semester, setSemester] = useState("");

  const addCourse = async () => {
    if (!name.trim() || !code.trim()) { flash("Course name and code are required.", "bad"); return; }
    const c = { id: uid(), name: name.trim(), code: code.trim().toUpperCase(), faculty: faculty.trim() || user.name, semester: semester.trim() };
    await saveCourses([...allCourses, c]);
    setShowNew(false); setName(""); setCode(""); setSemester("");
    flash("Course created.");
  };
  const removeCourse = async (id) => { await saveCourses(allCourses.filter((c) => c.id !== id)); flash("Course removed.", "warn"); };

  return (
    <div style={{ padding: "28px 34px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div className="serif" style={{ fontSize: 22, fontWeight: 700 }}>Courses</div>
        <button onClick={() => setShowNew((s) => !s)} style={{
          display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: "var(--oxblood)", color: "#fff",
          border: "none", borderRadius: 3, fontSize: 13, fontWeight: 600,
        }}><Plus size={14} /> New course</button>
      </div>
      {showNew && (
        <div style={{ marginTop: 16, background: "#fff", border: "1px solid var(--line)", borderRadius: 4, padding: 18, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, maxWidth: 640 }}>
          <Field label="Course code"><input value={code} onChange={(e) => setCode(e.target.value)} placeholder="EC301" /></Field>
          <Field label="Semester"><input value={semester} onChange={(e) => setSemester(e.target.value)} placeholder="V" /></Field>
          <Field label="Course name" full><input value={name} onChange={(e) => setName(e.target.value)} placeholder="Digital Signal Processing" /></Field>
          <Field label="Faculty in charge" full><input value={faculty} onChange={(e) => setFaculty(e.target.value)} placeholder={user.name} /></Field>
          <div style={{ gridColumn: "1 / -1", display: "flex", gap: 8 }}>
            <button onClick={addCourse} style={{ padding: "8px 16px", background: "var(--ink)", color: "#fff", border: "none", borderRadius: 3, fontSize: 13, fontWeight: 600 }}>Create</button>
            <button onClick={() => setShowNew(false)} style={{ padding: "8px 16px", background: "none", border: "1px solid var(--line)", borderRadius: 3, fontSize: 13 }}>Cancel</button>
          </div>
        </div>
      )}
      <div style={{ marginTop: 20, display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(230px,1fr))", gap: 14 }}>
        {courses.map((c) => (
          <div key={c.id} style={{ background: "#fff", border: "1px solid var(--line)", borderRadius: 4, padding: 16, position: "relative" }}>
            <div className="mono" style={{ fontSize: 11.5, color: "var(--brass)", fontWeight: 700 }}>{c.code}{c.semester ? ` · SEM ${c.semester}` : ""}</div>
            <div className="serif" style={{ fontSize: 15.5, fontWeight: 700, marginTop: 3 }}>{c.name}</div>
            <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 3 }}>{c.faculty}</div>
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button onClick={() => openCourse(c.id)} style={{
                display: "flex", alignItems: "center", gap: 4, background: "none", border: "1px solid var(--ink)", borderRadius: 3,
                padding: "6px 10px", fontSize: 12, fontWeight: 600,
              }}>Open <ChevronRight size={13} /></button>
              {user.role === "admin" && (
                <button onClick={() => removeCourse(c.id)} style={{ background: "none", border: "none", color: "var(--bad)" }}><Trash2 size={14} /></button>
              )}
            </div>
          </div>
        ))}
        {courses.length === 0 && !showNew && <div style={{ color: "var(--ink-soft)", fontSize: 13 }}>No courses yet — create one to begin.</div>}
      </div>
    </div>
  );
}
function Field({ label, children, full }) {
  return (
    <div style={{ gridColumn: full ? "1 / -1" : "auto" }}>
      <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginBottom: 3 }}>{label}</div>
      {children}
    </div>
  );
}

// ---------------- Course workspace ----------------
const COURSE_TABS = [
  { id: "cos", label: "Course Outcomes", icon: ClipboardList },
  { id: "matrix", label: "CO–PO/PSO Matrix", icon: Grid3x3 },
  { id: "assess", label: "Assessments & Marks", icon: Upload },
  { id: "report", label: "Attainment Report", icon: FileBarChart2 },
];

function CourseWorkspace({ course, pos, psos, onUpdateCourse, flash }) {
  const [tab, setTab] = useState("cos");
  const [cos, setCos] = useState([]);
  const [matrix, setMatrix] = useState({});
  const [components, setComponents] = useState([]);
  const [marksByComponent, setMarksByComponent] = useState({});
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    (async () => {
      setCos(await sget(`course:${course.id}:cos`, []));
      setMatrix(await sget(`course:${course.id}:matrix`, {}));
      setComponents(await sget(`course:${course.id}:components`, []));
      setMarksByComponent(await sget(`course:${course.id}:marks`, {}));
      setSettings(await sget(`course:${course.id}:settings`, DEFAULT_SETTINGS));
      setLoaded(true);
    })();
  }, [course.id]);

  const persistCos = async (next) => { setCos(next); await sset(`course:${course.id}:cos`, next); };
  const persistMatrix = async (next) => { setMatrix(next); await sset(`course:${course.id}:matrix`, next); };
  const persistComponents = async (next) => { setComponents(next); await sset(`course:${course.id}:components`, next); };
  const persistMarks = async (next) => { setMarksByComponent(next); await sset(`course:${course.id}:marks`, next); };
  const persistSettings = async (next) => { setSettings(next); await sset(`course:${course.id}:settings`, next); };

  if (!loaded) return <div style={{ padding: 34 }}>Loading course…</div>;

  return (
    <div>
      <div className="no-print" style={{ padding: "22px 34px 0" }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
          <div className="serif" style={{ fontSize: 21, fontWeight: 700 }}>{course.code} — {course.name}</div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{course.faculty}{course.semester ? ` · Sem ${course.semester}` : ""}</div>
        </div>
        <div style={{ display: "flex", gap: 4, marginTop: 16, borderBottom: "1px solid var(--line)" }}>
          {COURSE_TABS.map((t) => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              display: "flex", alignItems: "center", gap: 6, padding: "9px 14px", background: "none", border: "none",
              borderBottom: tab === t.id ? "2px solid var(--oxblood)" : "2px solid transparent",
              color: tab === t.id ? "var(--ink)" : "var(--ink-soft)", fontWeight: tab === t.id ? 600 : 500, fontSize: 13,
            }}><t.icon size={14} />{t.label}</button>
          ))}
        </div>
      </div>
      <div style={{ padding: "22px 34px 40px" }}>
        {tab === "cos" && <CoEditor cos={cos} setCos={persistCos} flash={flash} />}
        {tab === "matrix" && <MatrixEditor cos={cos} pos={pos} psos={psos} matrix={matrix} setMatrix={persistMatrix} />}
        {tab === "assess" && (
          <AssessmentTab
            cos={cos} components={components} setComponents={persistComponents}
            marksByComponent={marksByComponent} setMarksByComponent={persistMarks} flash={flash}
            method={settings.method}
          />
        )}
        {tab === "report" && (
          <ReportTab
            course={course} cos={cos} pos={pos} psos={psos} matrix={matrix}
            components={components} marksByComponent={marksByComponent}
            settings={settings} setSettings={persistSettings}
          />
        )}
      </div>
    </div>
  );
}

// ---- CO editor ----
function CoEditor({ cos, setCos, flash }) {
  const addCo = () => setCos([...cos, { id: uid(), code: `CO${cos.length + 1}`, text: "" }]);
  const update = (id, patch) => setCos(cos.map((c) => c.id === id ? { ...c, ...patch } : c));
  const remove = (id) => setCos(cos.filter((c) => c.id !== id));
  return (
    <div style={{ maxWidth: 780 }}>
      <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 14 }}>Define what a student should be able to do by the end of this course.</div>
      {cos.map((c) => (
        <div key={c.id} style={{ display: "flex", gap: 8, marginBottom: 8, alignItems: "flex-start" }}>
          <input value={c.code} onChange={(e) => update(c.id, { code: e.target.value })} style={{ width: 66, fontWeight: 600 }} />
          <input value={c.text} onChange={(e) => update(c.id, { text: e.target.value })} placeholder="e.g. Analyze discrete-time signals using Z-transform" style={{ flex: 1 }} />
          <button onClick={() => remove(c.id)} style={{ background: "none", border: "none", color: "var(--bad)", padding: 8 }}><Trash2 size={14} /></button>
        </div>
      ))}
      <button onClick={addCo} style={{ display: "flex", alignItems: "center", gap: 5, marginTop: 6, background: "none", border: "1px dashed var(--line)", borderRadius: 3, padding: "7px 12px", fontSize: 12.5, color: "var(--ink-soft)" }}>
        <Plus size={13} /> Add Course Outcome
      </button>
    </div>
  );
}

// ---- Matrix editor ----
function MatrixEditor({ cos, pos, psos, matrix, setMatrix }) {
  const all = [...pos.map((p, i) => ({ key: `PO${i + 1}`, label: p })), ...psos.map((p, i) => ({ key: `PSO${i + 1}`, label: p }))];
  const setVal = (coId, key, val) => {
    const v = Math.max(0, Math.min(3, Number(val) || 0));
    setMatrix({ ...matrix, [coId]: { ...(matrix[coId] || {}), [key]: v } });
  };
  if (cos.length === 0) return <div style={{ color: "var(--ink-soft)", fontSize: 13 }}>Add Course Outcomes first.</div>;
  return (
    <div>
      <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 12 }}>Correlation strength: 0 = none, 1 = low, 2 = medium, 3 = high.</div>
      <div style={{ overflow: "auto" }}>
        <table style={{ minWidth: 720 }}>
          <thead>
            <tr>
              <th style={{ textAlign: "left" }}>CO</th>
              {all.map((a) => <th key={a.key} title={a.label} className="mono">{a.key}</th>)}
            </tr>
          </thead>
          <tbody>
            {cos.map((c) => (
              <tr key={c.id}>
                <td style={{ textAlign: "left", fontWeight: 600 }} className="mono">{c.code}</td>
                {all.map((a) => (
                  <td key={a.key}>
                    <select value={(matrix[c.id] || {})[a.key] || 0} onChange={(e) => setVal(c.id, a.key, e.target.value)} style={{ width: 46, padding: "3px 2px", textAlign: "center" }}>
                      {[0, 1, 2, 3].map((n) => <option key={n} value={n}>{n === 0 ? "-" : n}</option>)}
                    </select>
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---- Assessment tab: components, questions, mark upload ----
function AssessmentTab({ cos, components, setComponents, marksByComponent, setMarksByComponent, flash, method }) {
  const [openComp, setOpenComp] = useState(components[0]?.id || null);

  const addComponent = () => {
    const comp = { id: uid(), name: `Assessment ${components.length + 1}`, type: "internal", weight: 100, targetPct: 60, questions: [] };
    setComponents([...components, comp]);
    setOpenComp(comp.id);
  };
  const updateComp = (id, patch) => setComponents(components.map((c) => c.id === id ? { ...c, ...patch } : c));
  const removeComp = (id) => {
    setComponents(components.filter((c) => c.id !== id));
    const nextMarks = { ...marksByComponent }; delete nextMarks[id]; setMarksByComponent(nextMarks);
  };

  const addQuestion = (compId) => {
    const comp = components.find((c) => c.id === compId);
    const q = { id: uid(), label: `Q${(comp.questions.length || 0) + 1}`, co: cos[0]?.id || "", bl: "Understand", max: 10 };
    updateComp(compId, { questions: [...comp.questions, q] });
  };
  const updateQuestion = (compId, qId, patch) => {
    const comp = components.find((c) => c.id === compId);
    updateComp(compId, { questions: comp.questions.map((q) => q.id === qId ? { ...q, ...patch } : q) });
  };
  const removeQuestion = (compId, qId) => {
    const comp = components.find((c) => c.id === compId);
    updateComp(compId, { questions: comp.questions.filter((q) => q.id !== qId) });
  };

  const handleFile = (compId, file) => {
    const comp = components.find((c) => c.id === compId);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (rows.length === 0) { flash("The file appears empty.", "bad"); return; }
        const headers = Object.keys(rows[0]);
        const rollKey = headers.find((h) => /roll|reg|student/i.test(h)) || headers[0];
        const parsed = {};
        rows.forEach((row) => {
          const roll = String(row[rollKey]).trim();
          if (!roll) return;
          const rec = {};
          comp.questions.forEach((q) => {
            const hit = headers.find((h) => h.trim().toLowerCase() === q.label.trim().toLowerCase());
            rec[q.id] = hit ? Number(row[hit]) || 0 : 0;
          });
          parsed[roll] = rec;
        });
        setMarksByComponent({ ...marksByComponent, [compId]: parsed });
        flash(`Imported marks for ${Object.keys(parsed).length} students.`);
      } catch (err) {
        flash("Could not read this file — check it's a valid .xlsx/.csv.", "bad");
      }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div>
      <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 14, maxWidth: 680 }}>
        Add each assessment (Test 1, Test 2, Assignment, End-Sem…), tag every question with a CO and Bloom's level, then upload the marks sheet.
        Excel columns should be one column per question label (e.g. <span className="mono">Q1, Q2, Q3</span>) plus a roll-number column.
      </div>
      <button onClick={addComponent} style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 16, background: "var(--oxblood)", color: "#fff", border: "none", borderRadius: 3, padding: "8px 14px", fontSize: 13, fontWeight: 600 }}>
        <Plus size={14} /> Add assessment component
      </button>
      {components.map((comp) => (
        <div key={comp.id} style={{ border: "1px solid var(--line)", borderRadius: 4, marginBottom: 14, background: "#fff" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", cursor: "pointer" }} onClick={() => setOpenComp(openComp === comp.id ? null : comp.id)}>
            {openComp === comp.id ? <ChevronLeft size={14} style={{ transform: "rotate(-90deg)" }} /> : <ChevronRight size={14} />}
            <input onClick={(e) => e.stopPropagation()} value={comp.name} onChange={(e) => updateComp(comp.id, { name: e.target.value })} style={{ fontWeight: 600, width: 200 }} />
            <select onClick={(e) => e.stopPropagation()} value={comp.type} onChange={(e) => updateComp(comp.id, { type: e.target.value, targetPct: comp.targetPct === undefined ? undefined : (e.target.value === "external" ? 40 : 60) })} style={{ width: 110 }}>
              <option value="internal">Internal</option>
              <option value="external">External</option>
            </select>
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>weight</div>
            <input onClick={(e) => e.stopPropagation()} type="number" value={comp.weight} onChange={(e) => updateComp(comp.id, { weight: Number(e.target.value) })} style={{ width: 60 }} />
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>%</div>
            {method !== "average" && (
              <>
                <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>target</div>
                <input onClick={(e) => e.stopPropagation()} type="number" value={comp.targetPct ?? (comp.type === "external" ? 40 : 60)} onChange={(e) => updateComp(comp.id, { targetPct: Number(e.target.value) })} style={{ width: 56 }} title="% marks a student must score to count as attaining" />
                <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>%</div>
              </>
            )}
            <div style={{ flex: 1 }} />
            <div style={{ fontSize: 12, color: marksByComponent[comp.id] ? "var(--good)" : "var(--ink-soft)" }}>
              {marksByComponent[comp.id] ? `${Object.keys(marksByComponent[comp.id]).length} students loaded` : "No marks yet"}
            </div>
            <button onClick={(e) => { e.stopPropagation(); removeComp(comp.id); }} style={{ background: "none", border: "none", color: "var(--bad)" }}><Trash2 size={14} /></button>
          </div>
          {openComp === comp.id && (
            <div style={{ padding: "0 14px 16px", borderTop: "1px solid var(--line)" }}>
              <div style={{ fontWeight: 600, fontSize: 12.5, margin: "12px 0 6px" }}>Questions</div>
              <table style={{ marginBottom: 10 }}>
                <thead><tr><th>Label</th><th>CO</th><th>Bloom's level</th><th>Max marks</th><th></th></tr></thead>
                <tbody>
                  {comp.questions.map((q) => (
                    <tr key={q.id}>
                      <td><input value={q.label} onChange={(e) => updateQuestion(comp.id, q.id, { label: e.target.value })} style={{ width: 60, textAlign: "center" }} /></td>
                      <td>
                        <select value={q.co} onChange={(e) => updateQuestion(comp.id, q.id, { co: e.target.value })}>
                          {cos.map((c) => <option key={c.id} value={c.id}>{c.code}</option>)}
                        </select>
                      </td>
                      <td>
                        <select value={q.bl} onChange={(e) => updateQuestion(comp.id, q.id, { bl: e.target.value })}>
                          {BL_LEVELS.map((b) => <option key={b} value={b}>{b}</option>)}
                        </select>
                      </td>
                      <td><input type="number" value={q.max} onChange={(e) => updateQuestion(comp.id, q.id, { max: Number(e.target.value) })} style={{ width: 60, textAlign: "center" }} /></td>
                      <td><button onClick={() => removeQuestion(comp.id, q.id)} style={{ background: "none", border: "none", color: "var(--bad)" }}><Trash2 size={13} /></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <button onClick={() => addQuestion(comp.id)} disabled={cos.length === 0} style={{ display: "flex", alignItems: "center", gap: 5, background: "none", border: "1px dashed var(--line)", borderRadius: 3, padding: "6px 10px", fontSize: 12, color: "var(--ink-soft)", marginBottom: 14 }}>
                <Plus size={12} /> Add question
              </button>
              <div>
                <label style={{
                  display: "inline-flex", alignItems: "center", gap: 7, padding: "8px 14px", background: "var(--paper-deep)",
                  border: "1px solid var(--line)", borderRadius: 3, fontSize: 12.5, cursor: "pointer",
                }}>
                  <Upload size={14} /> Upload marks (.xlsx / .csv)
                  <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
                    onChange={(e) => { if (e.target.files[0]) handleFile(comp.id, e.target.files[0]); e.target.value = ""; }} />
                </label>
                {marksByComponent[comp.id] && (
                  <MarksPreview questions={comp.questions} marks={marksByComponent[comp.id]} />
                )}
              </div>
            </div>
          )}
        </div>
      ))}
    </div>
  );
}
function MarksPreview({ questions, marks }) {
  const rolls = Object.keys(marks).slice(0, 6);
  return (
    <div style={{ marginTop: 12, overflow: "auto" }}>
      <table style={{ minWidth: 400 }}>
        <thead><tr><th>Roll</th>{questions.map((q) => <th key={q.id}>{q.label}</th>)}</tr></thead>
        <tbody>
          {rolls.map((r) => (
            <tr key={r}><td>{r}</td>{questions.map((q) => <td key={q.id}>{marks[r][q.id]}</td>)}</tr>
          ))}
        </tbody>
      </table>
      {Object.keys(marks).length > 6 && <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 4 }}>+ {Object.keys(marks).length - 6} more students</div>}
    </div>
  );
}

// ---- Report tab ----
function ReportTab({ course, cos, pos, psos, matrix, components, marksByComponent, settings, setSettings }) {
  const s = { ...DEFAULT_SETTINGS, ...settings };
  const coAttainments = useMemo(
    () => computeFinalCoAttainments(cos, components, marksByComponent, s),
    [cos, components, marksByComponent, s]
  );

  const all = [...pos.map((p, i) => ({ key: `PO${i + 1}`, label: p })), ...psos.map((p, i) => ({ key: `PSO${i + 1}`, label: p }))];
  const poAttainments = useMemo(() => {
    const out = {};
    all.forEach((a) => {
      const row = {};
      cos.forEach((co) => { row[co.id] = (matrix[co.id] || {})[a.key] || 0; });
      const cAtt = {}; cos.forEach((co) => { if (coAttainments[co.id]) cAtt[co.id] = coAttainments[co.id]; });
      out[a.key] = poAttainmentFromCos(cAtt, row);
    });
    return out;
  }, [all, cos, matrix, coAttainments]);

  const overallPct = useMemo(() => {
    const vals = Object.values(poAttainments).filter((v) => v != null);
    if (vals.length === 0) return 0;
    return (vals.reduce((s2, v) => s2 + v, 0) / vals.length / 3) * 100;
  }, [poAttainments]);

  const applyPreset = (name) => {
    if (name === "model1") setSettings({ ...s, method: "crossing", targetInternal: 60, targetExternal: 40, internalWeight: 40, externalWeight: 60, directWeight: 80, indirectWeight: 20 });
    if (name === "model2") setSettings({ ...s, method: "average", avgBands: { high: 70, med: 50 }, internalWeight: 50, externalWeight: 50, directWeight: 80, indirectWeight: 20 });
  };

  return (
    <div>
      <div className="no-print" style={{ background: "#fff", border: "1px solid var(--line)", borderRadius: 4, padding: 16, marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
          <button onClick={() => applyPreset("model1")} style={{ padding: "7px 12px", fontSize: 12.5, borderRadius: 3, border: "1px solid var(--line)", background: s.method === "crossing" ? "var(--ink)" : "#fff", color: s.method === "crossing" ? "#fff" : "var(--ink)" }}>Model 1 (40–60, % students crossing target)</button>
          <button onClick={() => applyPreset("model2")} style={{ padding: "7px 12px", fontSize: 12.5, borderRadius: 3, border: "1px solid var(--line)", background: s.method === "average" ? "var(--ink)" : "#fff", color: s.method === "average" ? "#fff" : "var(--ink)" }}>Model 2 (50–50, class average score)</button>
          <button onClick={() => window.print()} style={{
            display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", background: "var(--oxblood)", color: "#fff",
            border: "none", borderRadius: 3, fontSize: 12.5, fontWeight: 600, marginLeft: "auto",
          }}><Printer size={14} /> Print / Save as PDF</button>
        </div>
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
          {s.method === "average" && (
            <>
              <Field label="High (level 3) if class avg ≥ %"><input type="number" value={s.avgBands.high} onChange={(e) => setSettings({ ...s, avgBands: { ...s.avgBands, high: Number(e.target.value) } })} style={{ width: 80 }} /></Field>
              <Field label="Medium (level 2) if class avg ≥ %"><input type="number" value={s.avgBands.med} onChange={(e) => setSettings({ ...s, avgBands: { ...s.avgBands, med: Number(e.target.value) } })} style={{ width: 80 }} /></Field>
            </>
          )}
          <Field label="Internal weight in final CO (%)"><input type="number" value={s.internalWeight} onChange={(e) => setSettings({ ...s, internalWeight: Number(e.target.value) })} style={{ width: 70 }} /></Field>
          <Field label="External weight in final CO (%)"><input type="number" value={s.externalWeight} onChange={(e) => setSettings({ ...s, externalWeight: Number(e.target.value) })} style={{ width: 70 }} /></Field>
          <Field label="Direct weight (%)"><input type="number" value={s.directWeight} onChange={(e) => setSettings({ ...s, directWeight: Number(e.target.value) })} style={{ width: 70 }} /></Field>
          <Field label="Indirect / survey weight (%)"><input type="number" value={s.indirectWeight} onChange={(e) => setSettings({ ...s, indirectWeight: Number(e.target.value) })} style={{ width: 70 }} /></Field>
        </div>
      </div>

      <div style={{ display: "flex", gap: 24, alignItems: "center", background: "#fff", border: "1px solid var(--line)", borderRadius: 4, padding: "18px 22px", marginBottom: 22 }}>
        <AttainmentSeal pct={overallPct} label="Program-level" />
        <div>
          <div className="serif" style={{ fontWeight: 700, fontSize: 16 }}>{course.code} — {course.name}</div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{course.faculty}{course.semester ? ` · Semester ${course.semester}` : ""}</div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 4 }}>
            {components.length} assessment component(s) · {s.method === "average" ? "class-average method" : "students-crossing-target method"} · {s.internalWeight}% internal / {s.externalWeight}% external
          </div>
        </div>
      </div>

      <div className="serif" style={{ fontWeight: 700, fontSize: 15, marginBottom: 8 }}>Course Outcome attainment</div>
      <table style={{ marginBottom: 10, maxWidth: 760 }}>
        <thead><tr><th style={{ textAlign: "left" }}>CO</th><th style={{ textAlign: "left" }}>Description</th><th>Direct level</th><th>Course-end survey<br />(% rating ≥3 "good")</th><th>Final level</th></tr></thead>
        <tbody>
          {cos.map((co) => (
            <tr key={co.id}>
              <td className="mono" style={{ fontWeight: 600 }}>{co.code}</td>
              <td style={{ textAlign: "left" }}>{co.text}</td>
              <td>{coAttainments[co.id] ? coAttainments[co.id].directLevel : <span style={{ color: "var(--ink-soft)" }}>no data</span>}</td>
              <td className="no-print">
                <input type="number" min="0" max="100" placeholder="—" value={(s.indirectByCo || {})[co.id] ?? ""}
                  onChange={(e) => setSettings({ ...s, indirectByCo: { ...(s.indirectByCo || {}), [co.id]: e.target.value === "" ? undefined : Number(e.target.value) } })}
                  style={{ width: 60, textAlign: "center" }} />
              </td>
              <td>{coAttainments[co.id] ? <LevelBadge level={coAttainments[co.id].level} /> : <span style={{ color: "var(--ink-soft)" }}>no data</span>}</td>
            </tr>
          ))}
          {cos.length === 0 && <tr><td colSpan={5} style={{ color: "var(--ink-soft)" }}>No COs defined.</td></tr>}
        </tbody>
      </table>
      <div className="no-print" style={{ fontSize: 11.5, color: "var(--ink-soft)", marginBottom: 24, maxWidth: 700 }}>
        Course-end survey is optional — leave blank to use direct assessment only. When filled in, it blends in at {s.indirectWeight}% (direct assessment keeps the other {s.directWeight}%), per the standard course-end-survey method (Level 1 ≥60% rating "good" or better, Level 2 ≥70%, Level 3 ≥80%).
      </div>

      <div className="serif" style={{ fontWeight: 700, fontSize: 15, marginBottom: 8 }}>PO / PSO attainment (direct)</div>
      <div style={{ overflow: "auto" }}>
        <table style={{ minWidth: 700, maxWidth: 900 }}>
          <thead><tr>{all.map((a) => <th key={a.key} className="mono" title={a.label}>{a.key}</th>)}</tr></thead>
          <tbody>
            <tr>{all.map((a) => (
              <td key={a.key}>{poAttainments[a.key] != null ? <LevelBadge level={poAttainments[a.key]} /> : <span style={{ color: "var(--ink-soft)" }}>—</span>}</td>
            ))}</tr>
          </tbody>
        </table>
      </div>
      <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginTop: 16, maxWidth: 700, lineHeight: 1.6 }}>
        Formula used — CO attainment: {s.method === "average"
          ? `class average score banded (≥${s.avgBands.high}%→3, ≥${s.avgBands.med}%→2, else→1)`
          : `% of students crossing the target score banded (≥80%→3, ≥70%→2, ≥60%→1, else→0)`}, blended {s.internalWeight}% internal / {s.externalWeight}% external, then {s.directWeight}% direct / {s.indirectWeight}% indirect if a survey value is entered.
        PO/PSO attainment: Σ(CO level × mapping strength) ÷ (3 × count of non-zero mappings) — the standard NBA/OBE equation.
      </div>
    </div>
  );
}
function LevelBadge({ level }) {
  const color = level >= 2.5 ? "var(--good)" : level >= 1.5 ? "var(--warn)" : level > 0 ? "var(--bad)" : "var(--ink-soft)";
  return <span className="mono" style={{ fontWeight: 700, color }}>{level}</span>;
}

// ---------------- Admin department-wide report ----------------
function AdminReport({ courses, pos, psos }) {
  const [rows, setRows] = useState([]);
  const [indirectPO, setIndirectPO] = useState({});
  const [indirectWeight, setIndirectWeight] = useState(20);
  const all = [...pos.map((p, i) => ({ key: `PO${i + 1}`, label: p })), ...psos.map((p, i) => ({ key: `PSO${i + 1}`, label: p }))];

  useEffect(() => {
    (async () => {
      setIndirectPO(await sget("app:indirectPO", {}));
      const results = [];
      for (const course of courses) {
        const cos = await sget(`course:${course.id}:cos`, []);
        const matrix = await sget(`course:${course.id}:matrix`, {});
        const components = await sget(`course:${course.id}:components`, []);
        const marks = await sget(`course:${course.id}:marks`, {});
        const settings = await sget(`course:${course.id}:settings`, DEFAULT_SETTINGS);
        const coAtt = computeFinalCoAttainments(cos, components, marks, settings);
        const poRow = {};
        all.forEach((a) => {
          const row = {}; cos.forEach((co) => { row[co.id] = (matrix[co.id] || {})[a.key] || 0; });
          poRow[a.key] = poAttainmentFromCos(coAtt, row);
        });
        results.push({ course, poRow });
      }
      setRows(results);
    })();
  }, [courses]);

  const persistIndirect = async (next) => { setIndirectPO(next); await sset("app:indirectPO", next); };

  const directAvg = {};
  all.forEach((a) => {
    const vals = rows.map((r) => r.poRow[a.key]).filter((v) => v != null);
    directAvg[a.key] = vals.length ? +(vals.reduce((s, v) => s + v, 0) / vals.length).toFixed(2) : null;
  });
  const finalAvg = {};
  all.forEach((a) => {
    const dv = directAvg[a.key];
    const iv = indirectPO[a.key];
    if (dv == null) { finalAvg[a.key] = null; return; }
    if (iv == null || iv === "") { finalAvg[a.key] = dv; return; }
    finalAvg[a.key] = +(((dv * (100 - indirectWeight)) + (Number(iv) * indirectWeight)) / 100).toFixed(2);
  });
  const overallPct = (() => {
    const vals = Object.values(finalAvg).filter((v) => v != null);
    return vals.length ? (vals.reduce((s, v) => s + v, 0) / vals.length / 3) * 100 : 0;
  })();

  return (
    <div style={{ padding: "28px 34px" }}>
      <div className="serif" style={{ fontSize: 22, fontWeight: 700 }}>Program-level attainment</div>
      <div style={{ color: "var(--ink-soft)", fontSize: 13, margin: "4px 0 20px" }}>Aggregated across every course in the department.</div>
      <div style={{ display: "flex", gap: 24, alignItems: "center", background: "#fff", border: "1px solid var(--line)", borderRadius: 4, padding: "18px 22px", marginBottom: 24 }}>
        <AttainmentSeal pct={overallPct} label="Program average" />
        <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Based on {rows.length} course(s) with recorded attainment, {100 - indirectWeight}% direct + {indirectWeight}% survey-indirect.</div>
      </div>
      <div style={{ overflow: "auto", marginBottom: 20 }}>
        <table style={{ minWidth: 900 }}>
          <thead><tr><th style={{ textAlign: "left" }}>Course</th>{all.map((a) => <th key={a.key} className="mono" title={a.label}>{a.key}</th>)}</tr></thead>
          <tbody>
            {rows.map(({ course, poRow }) => (
              <tr key={course.id}>
                <td style={{ textAlign: "left" }}>{course.code}</td>
                {all.map((a) => <td key={a.key}>{poRow[a.key] != null ? <LevelBadge level={poRow[a.key]} /> : <span style={{ color: "var(--ink-soft)" }}>—</span>}</td>)}
              </tr>
            ))}
            <tr style={{ background: "var(--paper-deep)", fontWeight: 700 }}>
              <td style={{ textAlign: "left" }}>Direct average</td>
              {all.map((a) => <td key={a.key}>{directAvg[a.key] != null ? <LevelBadge level={directAvg[a.key]} /> : "—"}</td>)}
            </tr>
            <tr className="no-print">
              <td style={{ textAlign: "left", fontSize: 12, color: "var(--ink-soft)" }}>Survey / indirect (0–3, optional)</td>
              {all.map((a) => (
                <td key={a.key}>
                  <input type="number" min="0" max="3" step="0.1" placeholder="—" value={indirectPO[a.key] ?? ""}
                    onChange={(e) => persistIndirect({ ...indirectPO, [a.key]: e.target.value === "" ? undefined : Number(e.target.value) })}
                    style={{ width: 50, textAlign: "center" }} />
                </td>
              ))}
            </tr>
            <tr style={{ background: "var(--paper-deep)", fontWeight: 700 }}>
              <td style={{ textAlign: "left" }}>Final PO/PSO attainment</td>
              {all.map((a) => <td key={a.key}>{finalAvg[a.key] != null ? <LevelBadge level={finalAvg[a.key]} /> : "—"}</td>)}
            </tr>
          </tbody>
        </table>
      </div>
      <div className="no-print" style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
        <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>Survey/indirect weight in final PO/PSO:</div>
        <input type="number" min="0" max="100" value={indirectWeight} onChange={(e) => setIndirectWeight(Number(e.target.value))} style={{ width: 60 }} />
        <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>%</div>
      </div>
      <div style={{ fontSize: 11.5, color: "var(--ink-soft)", maxWidth: 700, lineHeight: 1.6 }}>
        Direct average = mean of each course's direct PO/PSO attainment. Survey/indirect values (Graduate Exit, Alumni, Employer, Parent, Faculty surveys) are entered manually here on a 0–3 scale, per the standard NBA practice of surveying stakeholders on a 1–5 Likert scale and banding the results. Final = direct×{100 - indirectWeight}% + indirect×{indirectWeight}%.
      </div>
    </div>
  );
}
