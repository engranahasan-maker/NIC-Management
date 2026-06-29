import { useState, useEffect } from "react";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from "xlsx";

// ============================================================
// SUPABASE CONFIG
// ============================================================
const SUPABASE_URL = "https://jijxnycopycsysugppnw.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImppanhueWNvcHljc3lzdWdwcG53Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI3MDAyNTAsImV4cCI6MjA5ODI3NjI1MH0.7kXRGGnW4VdWU9VT1XEBKp5oC9V5Z21KA_PBqZtvjJA";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// ============================================================
// ADMIN CREDENTIALS (সহজ login system)
// ============================================================
const ADMIN_USER = "admin@noksha.com";
const ADMIN_PASS = "noksha2024";

// ============================================================
// COLOR TOKENS
// ============================================================
const C = {
  navy: "#1A1A2E", navyLight: "#16213E", navyMid: "#0F3460",
  gold: "#C9A84C", goldLight: "#E8C97A", goldPale: "#FDF6E3",
  white: "#FFFFFF", gray50: "#F8F9FA", gray100: "#F1F3F5",
  gray200: "#E9ECEF", gray400: "#ADB5BD", gray600: "#6C757D", gray800: "#343A40",
  green: "#28A745", greenLight: "#D4EDDA", red: "#DC3545", redLight: "#F8D7DA",
  yellow: "#FFC107", yellowLight: "#FFF3CD", blue: "#007BFF", blueLight: "#D0E8FF",
};

const fmt = (n) => "৳" + Number(n || 0).toLocaleString("bn-BD");
const fmtNum = (n) => Number(n || 0).toLocaleString("bn-BD");

// ============================================================
// NOKSHA PAD HEADER (print এ দেখাবে)
// ============================================================
const NokshaPad = ({ title }) => (
  <div style={{ display: "none" }} className="noksha-pad">
    <div style={{ textAlign: "center", borderBottom: `3px solid ${C.gold}`, paddingBottom: 12, marginBottom: 16 }}>
      <div style={{ fontSize: 28, fontWeight: 800, color: C.navy }}>নকশা ইন্টেরিয়র অ্যান্ড কনস্ট্রাকশন</div>
      <div style={{ fontSize: 14, color: C.gold, fontWeight: 600 }}>Noksha Interior & Construction (NIC)</div>
      <div style={{ fontSize: 12, color: C.gray600, marginTop: 4 }}>নীলটুলী, ফরিদপুর | পল্লবী, ঢাকা</div>
      <div style={{ fontSize: 12, color: C.gray600 }}>📞 01XXX-XXXXXX | ✉ info@noksha.com</div>
      <div style={{ fontSize: 16, fontWeight: 700, color: C.navy, marginTop: 8, padding: "4px 20px", background: C.goldPale, borderRadius: 6, display: "inline-block" }}>{title}</div>
      <div style={{ fontSize: 11, color: C.gray400, marginTop: 4 }}>তারিখ: {new Date().toLocaleDateString("bn-BD")}</div>
    </div>
  </div>
);

// ============================================================
// UI COMPONENTS
// ============================================================
function Badge({ label, color }) {
  const colors = { green: { bg: C.greenLight, text: C.green }, red: { bg: C.redLight, text: C.red }, yellow: { bg: C.yellowLight, text: "#856404" }, blue: { bg: C.blueLight, text: C.blue }, gray: { bg: C.gray200, text: C.gray600 }, gold: { bg: "#FDF3D8", text: "#8a6a00" } };
  const s = colors[color] || colors.gray;
  return <span style={{ background: s.bg, color: s.text, padding: "2px 10px", borderRadius: 20, fontSize: 12, fontWeight: 600 }}>{label}</span>;
}

function StatusBadge({ status }) {
  const map = { "চলমান": "blue", "সম্পন্ন": "green", "নতুন": "gold", "পেন্ডিং": "yellow", "সক্রিয়": "green", "কর্মরত": "green", "উপস্থিত": "green", "অনুপস্থিত": "red", "অর্ধদিন": "yellow", "ছুটি": "gray", "আয়": "green", "ব্যয়": "red" };
  return <Badge label={status} color={map[status] || "gray"} />;
}

function Card({ children, style }) {
  return <div style={{ background: C.white, borderRadius: 12, padding: 20, boxShadow: "0 2px 12px rgba(0,0,0,0.07)", border: `1px solid ${C.gray200}`, ...style }}>{children}</div>;
}

function StatCard({ icon, label, value, sub, color }) {
  return (
    <Card style={{ display: "flex", alignItems: "center", gap: 16 }}>
      <div style={{ width: 52, height: 52, borderRadius: 12, background: color || C.goldPale, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24, flexShrink: 0 }}>{icon}</div>
      <div>
        <div style={{ fontSize: 22, fontWeight: 700, color: C.navy }}>{value}</div>
        <div style={{ fontSize: 13, color: C.gray600, marginTop: 2 }}>{label}</div>
        {sub && <div style={{ fontSize: 12, color: C.gold, marginTop: 2, fontWeight: 600 }}>{sub}</div>}
      </div>
    </Card>
  );
}

function ProgressBar({ value, color }) {
  return <div style={{ background: C.gray200, borderRadius: 99, height: 8, overflow: "hidden" }}><div style={{ width: `${value}%`, height: "100%", background: color || C.gold, borderRadius: 99 }} /></div>;
}

function SectionHeader({ title, action, onAction, onPrint, onExport }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 8 }}>
      <h2 style={{ margin: 0, color: C.navy, fontSize: 18, fontWeight: 700 }}>{title}</h2>
      <div style={{ display: "flex", gap: 8 }}>
        {onExport && <button onClick={onExport} style={{ background: C.green, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 12 }}>📊 Excel</button>}
        {onPrint && <button onClick={onPrint} style={{ background: C.blue, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 12 }}>🖨️ Print</button>}
        {action && <button onClick={onAction} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 8, padding: "8px 16px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ {action}</button>}
      </div>
    </div>
  );
}

function Table({ headers, rows }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
        <thead>
          <tr style={{ background: C.gray50 }}>
            {headers.map((h, i) => <th key={i} style={{ padding: "10px 14px", textAlign: "left", color: C.gray600, fontWeight: 600, borderBottom: `2px solid ${C.gray200}`, whiteSpace: "nowrap" }}>{h}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={e => e.currentTarget.style.background = C.gray50} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
              {row.map((cell, j) => <td key={j} style={{ padding: "10px 14px", color: C.gray800, verticalAlign: "middle" }}>{cell}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 16, width: "100%", maxWidth: 520, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 24px", borderBottom: `1px solid ${C.gray200}`, position: "sticky", top: 0, background: C.white }}>
          <h3 style={{ margin: 0, color: C.navy, fontSize: 16, fontWeight: 700 }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: C.gray600 }}>×</button>
        </div>
        <div style={{ padding: 24 }}>{children}</div>
      </div>
    </div>
  );
}

function FormField({ label, children }) {
  return <div style={{ marginBottom: 16 }}><label style={{ display: "block", fontSize: 13, fontWeight: 600, color: C.gray800, marginBottom: 6 }}>{label}</label>{children}</div>;
}

const inputStyle = { width: "100%", padding: "9px 12px", border: `1px solid ${C.gray200}`, borderRadius: 8, fontSize: 13, color: C.gray800, outline: "none", boxSizing: "border-box", fontFamily: "inherit" };
const btnStyle = { width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 10, padding: "12px", fontWeight: 700, cursor: "pointer", fontSize: 14, marginTop: 8, fontFamily: "inherit" };

// ============================================================
// PRINT UTILITY
// ============================================================
const printSection = (title, contentId) => {
  const content = document.getElementById(contentId);
  if (!content) return;
  const win = window.open("", "_blank");
  win.document.write(`
    <html><head><title>${title} — NIC</title>
    <style>
      body { font-family: 'Arial', sans-serif; padding: 20px; color: #1A1A2E; }
      .header { text-align: center; border-bottom: 3px solid #C9A84C; padding-bottom: 12px; margin-bottom: 20px; }
      .company-name { font-size: 24px; font-weight: 800; color: #1A1A2E; }
      .company-sub { font-size: 13px; color: #C9A84C; font-weight: 600; }
      .company-info { font-size: 11px; color: #6C757D; }
      .doc-title { font-size: 16px; font-weight: 700; background: #FDF6E3; padding: 4px 20px; border-radius: 6px; display: inline-block; margin-top: 8px; }
      table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 12px; }
      th { background: #F1F3F5; padding: 8px 12px; text-align: left; border: 1px solid #E9ECEF; font-weight: 600; }
      td { padding: 8px 12px; border: 1px solid #E9ECEF; }
      tr:nth-child(even) { background: #F8F9FA; }
      .footer { margin-top: 30px; border-top: 1px solid #C9A84C; padding-top: 10px; font-size: 11px; color: #6C757D; text-align: center; }
    </style></head><body>
    <div class="header">
      <div class="company-name">নকশা ইন্টেরিয়র অ্যান্ড কনস্ট্রাকশন</div>
      <div class="company-sub">Noksha Interior & Construction (NIC)</div>
      <div class="company-info">নীলটুলী, ফরিদপুর | পল্লবী, ঢাকা | 📞 01XXX-XXXXXX</div>
      <div class="doc-title">${title}</div>
      <div class="company-info">তারিখ: ${new Date().toLocaleDateString("bn-BD")}</div>
    </div>
    ${content.innerHTML}
    <div class="footer">Noksha Interior & Construction — সকল অধিকার সংরক্ষিত</div>
    </body></html>
  `);
  win.document.close();
  win.print();
};

// ============================================================
// EXCEL EXPORT UTILITY
// ============================================================
const exportToExcel = (data, sheetName, fileName) => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = Object.keys(data[0] || {}).map(() => ({ wch: 20 }));
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `NIC_${fileName}_${new Date().toISOString().split("T")[0]}.xlsx`);
};

// ============================================================
// LOGIN PAGE
// ============================================================
function LoginPage({ onLogin }) {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    setLoading(true);
    setError("");
    if (email === ADMIN_USER && pass === ADMIN_PASS) {
      localStorage.setItem("nic_logged_in", "true");
      onLogin();
    } else {
      setError("❌ ভুল ইমেইল বা পাসওয়ার্ড!");
    }
    setLoading(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: `linear-gradient(135deg, ${C.navy} 0%, ${C.navyMid} 100%)`, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 20, padding: 40, width: "100%", maxWidth: 400, boxShadow: "0 20px 60px rgba(0,0,0,0.3)" }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ width: 64, height: 64, background: `linear-gradient(135deg, ${C.navy}, ${C.navyMid})`, borderRadius: 16, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 32, margin: "0 auto 16px" }}>🏛️</div>
          <div style={{ fontSize: 22, fontWeight: 800, color: C.navy }}>নকশা</div>
          <div style={{ fontSize: 13, color: C.gold, fontWeight: 600 }}>Noksha Interior & Construction</div>
          <div style={{ fontSize: 12, color: C.gray400, marginTop: 4 }}>Management System</div>
        </div>

        {error && <div style={{ background: C.redLight, color: C.red, padding: "10px 14px", borderRadius: 8, fontSize: 13, marginBottom: 16, textAlign: "center" }}>{error}</div>}

        <FormField label="ইমেইল">
          <input style={inputStyle} type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="admin@noksha.com" onKeyDown={e => e.key === "Enter" && handleLogin()} />
        </FormField>
        <FormField label="পাসওয়ার্ড">
          <input style={inputStyle} type="password" value={pass} onChange={e => setPass(e.target.value)} placeholder="••••••••" onKeyDown={e => e.key === "Enter" && handleLogin()} />
        </FormField>

        <button onClick={handleLogin} disabled={loading} style={{ ...btnStyle, background: loading ? C.gray400 : C.navy, marginTop: 8 }}>
          {loading ? "লগইন হচ্ছে..." : "🔐 লগইন করুন"}
        </button>

        <div style={{ textAlign: "center", marginTop: 16, fontSize: 11, color: C.gray400 }}>
          Demo: admin@noksha.com / noksha2024
        </div>
      </div>
    </div>
  );
}

// ============================================================
// DASHBOARD
// ============================================================
function Dashboard({ projects, clients, employees, transactions, materials }) {
  const totalIncome = transactions.filter(t => t.type === "আয়").reduce((s, t) => s + (t.amount || 0), 0);
  const totalExpense = transactions.filter(t => t.type === "ব্যয়").reduce((s, t) => s + (t.amount || 0), 0);
  const activeProjects = projects.filter(p => p.status === "চলমান").length;
  const lowStock = materials.filter(m => m.stock < m.min_stock).length;

  const handlePrint = () => printSection("ড্যাশবোর্ড সারসংক্ষেপ", "dashboard-content");
  const handleExport = () => {
    exportToExcel([
      { বিভাগ: "সক্রিয় প্রজেক্ট", মান: activeProjects },
      { বিভাগ: "মোট ক্লায়েন্ট", মান: clients.length },
      { বিভাগ: "মোট কর্মী", মান: employees.length },
      { বিভাগ: "মাসিক আয়", মান: totalIncome },
      { বিভাগ: "মাসিক ব্যয়", মান: totalExpense },
      { বিভাগ: "নিট লাভ", মান: totalIncome - totalExpense },
    ], "Dashboard", "Dashboard");
  };

  return (
    <div>
      <SectionHeader title="ড্যাশবোর্ড" onPrint={handlePrint} onExport={handleExport} />
      <div id="dashboard-content">
        <div style={{ marginBottom: 24, padding: "16px 20px", background: `linear-gradient(135deg, ${C.navy}, ${C.navyMid})`, borderRadius: 14, color: C.white }}>
          <div style={{ fontSize: 13, opacity: 0.8 }}>স্বাগতম,</div>
          <div style={{ fontSize: 20, fontWeight: 700, marginTop: 4 }}>Noksha Interior & Construction</div>
          <div style={{ fontSize: 12, opacity: 0.7, marginTop: 4 }}>আজকের তারিখ: {new Date().toLocaleDateString("bn-BD")}</div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14, marginBottom: 24 }}>
          <StatCard icon="🏗️" label="সক্রিয় প্রজেক্ট" value={fmtNum(activeProjects)} sub={`মোট ${projects.length}টি`} color="#EEF2FF" />
          <StatCard icon="👥" label="মোট ক্লায়েন্ট" value={fmtNum(clients.length)} color="#F0FFF4" />
          <StatCard icon="👷" label="কর্মী সংখ্যা" value={fmtNum(employees.length)} color="#FFF8E1" />
          <StatCard icon="💰" label="মোট আয়" value={fmt(totalIncome)} color="#F0FFF4" />
          <StatCard icon="💸" label="মোট ব্যয়" value={fmt(totalExpense)} color="#FFF5F5" />
          <StatCard icon="📊" label="নিট লাভ" value={fmt(totalIncome - totalExpense)} color={totalIncome >= totalExpense ? "#F0FFF4" : "#FFF5F5"} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
          <Card>
            <h3 style={{ margin: "0 0 14px", color: C.navy, fontSize: 15 }}>চলমান প্রজেক্ট</h3>
            {projects.filter(p => p.status === "চলমান").map(p => (
              <div key={p.id} style={{ marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: C.navy }}>{p.name}</span>
                  <span style={{ fontSize: 12, color: C.gold, fontWeight: 700 }}>{p.progress}%</span>
                </div>
                <ProgressBar value={p.progress} />
                <div style={{ fontSize: 11, color: C.gray400, marginTop: 4 }}>{p.client}</div>
              </div>
            ))}
            {projects.filter(p => p.status === "চলমান").length === 0 && <div style={{ color: C.gray400, fontSize: 13 }}>কোনো চলমান প্রজেক্ট নেই</div>}
          </Card>
          <Card>
            <h3 style={{ margin: "0 0 14px", color: C.navy, fontSize: 15 }}>সর্বশেষ লেনদেন</h3>
            {[...transactions].slice(-5).reverse().map(t => (
              <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.gray100}` }}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600 }}>{t.description}</div>
                  <div style={{ fontSize: 11, color: C.gray400 }}>{t.date}</div>
                </div>
                <div style={{ fontWeight: 700, color: t.type === "আয়" ? C.green : C.red, fontSize: 13 }}>
                  {t.type === "আয়" ? "+" : "-"}{fmt(t.amount)}
                </div>
              </div>
            ))}
          </Card>
        </div>
        {lowStock > 0 && (
          <Card>
            <h3 style={{ margin: "0 0 14px", color: C.navy, fontSize: 15 }}>কম স্টক সতর্কতা ⚠️</h3>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
              {materials.filter(m => m.stock < m.min_stock).map(m => (
                <div key={m.id} style={{ background: C.redLight, border: `1px solid #F5C6CB`, borderRadius: 10, padding: "10px 14px" }}>
                  <div style={{ fontWeight: 700, color: C.red, fontSize: 13 }}>{m.name}</div>
                  <div style={{ fontSize: 12, color: C.gray600 }}>বর্তমান: {fmtNum(m.stock)} {m.unit}</div>
                </div>
              ))}
            </div>
          </Card>
        )}
      </div>
    </div>
  );
}

// ============================================================
// PROJECTS
// ============================================================
function Projects({ data, onAdd }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", client: "", budget: "", start_date: "", end_date: "", location: "", type: "আবাসিক নির্মাণ", status: "নতুন", progress: 0 });

  const save = async () => {
    if (!form.name || !form.client || !form.budget) return alert("নাম, ক্লায়েন্ট ও বাজেট আবশ্যক");
    const { error } = await supabase.from("projects").insert([{ ...form, budget: +form.budget, progress: +form.progress, spent: 0 }]);
    if (error) return alert("Error: " + error.message);
    onAdd(); setShowModal(false);
    setForm({ name: "", client: "", budget: "", start_date: "", end_date: "", location: "", type: "আবাসিক নির্মাণ", status: "নতুন", progress: 0 });
  };

  const handleExport = () => exportToExcel(data.map(p => ({ নাম: p.name, ক্লায়েন্ট: p.client, বাজেট: p.budget, খরচ: p.spent, অগ্রগতি: p.progress + "%", স্ট্যাটাস: p.status, অবস্থান: p.location, শুরু: p.start_date, শেষ: p.end_date })), "Projects", "Projects");
  const handlePrint = () => printSection("প্রজেক্ট তালিকা", "projects-content");

  return (
    <div>
      <SectionHeader title="প্রজেক্ট ব্যবস্থাপনা" action="নতুন প্রজেক্ট" onAction={() => setShowModal(true)} onPrint={handlePrint} onExport={handleExport} />
      <div id="projects-content" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 16 }}>
        {data.map(p => (
          <Card key={p.id}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
              <div>
                <div style={{ fontWeight: 700, color: C.navy, fontSize: 14 }}>{p.name}</div>
                <div style={{ fontSize: 12, color: C.gray400, marginTop: 3 }}>📍 {p.location}</div>
              </div>
              <StatusBadge status={p.status} />
            </div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}>ক্লায়েন্ট: <strong>{p.client}</strong></div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 12 }}>ধরন: {p.type}</div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ fontSize: 12, color: C.gray600 }}>অগ্রগতি</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.gold }}>{p.progress}%</span>
            </div>
            <ProgressBar value={p.progress} color={p.progress === 100 ? C.green : C.gold} />
            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 12, paddingTop: 12, borderTop: `1px solid ${C.gray100}` }}>
              <div><div style={{ fontSize: 11, color: C.gray400 }}>বাজেট</div><div style={{ fontWeight: 700, color: C.navy, fontSize: 13 }}>{fmt(p.budget)}</div></div>
              <div style={{ textAlign: "right" }}><div style={{ fontSize: 11, color: C.gray400 }}>খরচ</div><div style={{ fontWeight: 700, color: C.green, fontSize: 13 }}>{fmt(p.spent)}</div></div>
            </div>
          </Card>
        ))}
      </div>
      {showModal && (
        <Modal title="নতুন প্রজেক্ট" onClose={() => setShowModal(false)}>
          <FormField label="প্রজেক্টের নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <FormField label="ক্লায়েন্ট *"><input style={inputStyle} value={form.client} onChange={e => setForm({ ...form, client: e.target.value })} /></FormField>
          <FormField label="বাজেট (টাকা) *"><input style={inputStyle} type="number" value={form.budget} onChange={e => setForm({ ...form, budget: e.target.value })} /></FormField>
          <FormField label="অবস্থান"><input style={inputStyle} value={form.location} onChange={e => setForm({ ...form, location: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="শুরুর তারিখ"><input style={inputStyle} type="date" value={form.start_date} onChange={e => setForm({ ...form, start_date: e.target.value })} /></FormField>
            <FormField label="শেষের তারিখ"><input style={inputStyle} type="date" value={form.end_date} onChange={e => setForm({ ...form, end_date: e.target.value })} /></FormField>
          </div>
          <FormField label="ধরন">
            <select style={inputStyle} value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}>
              {["আবাসিক নির্মাণ", "কমার্শিয়াল ইন্টেরিয়র", "মডুলার কিচেন", "থ্রিডি ভিজুয়ালাইজেশন", "সিভিল কনস্ট্রাকশন"].map(t => <option key={t}>{t}</option>)}
            </select>
          </FormField>
          <button onClick={save} style={btnStyle}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// CLIENTS
// ============================================================
function Clients({ data, onAdd }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", phone: "", email: "", address: "", status: "সক্রিয়" });

  const save = async () => {
    if (!form.name || !form.phone) return alert("নাম ও ফোন আবশ্যক");
    const { error } = await supabase.from("clients").insert([{ ...form, join_date: new Date().toISOString().split("T")[0] }]);
    if (error) return alert("Error: " + error.message);
    onAdd(); setShowModal(false);
    setForm({ name: "", phone: "", email: "", address: "", status: "সক্রিয়" });
  };

  const handleExport = () => exportToExcel(data.map(c => ({ নাম: c.name, ফোন: c.phone, ইমেইল: c.email, ঠিকানা: c.address, স্ট্যাটাস: c.status, যোগদান: c.join_date })), "Clients", "Clients");
  const handlePrint = () => printSection("ক্লায়েন্ট তালিকা", "clients-content");

  return (
    <div>
      <SectionHeader title="ক্লায়েন্ট ব্যবস্থাপনা" action="নতুন ক্লায়েন্ট" onAction={() => setShowModal(true)} onPrint={handlePrint} onExport={handleExport} />
      <Card>
        <div id="clients-content">
          <Table
            headers={["নাম", "ফোন", "ইমেইল", "ঠিকানা", "স্ট্যাটাস", "যোগদান"]}
            rows={data.map(c => [<strong style={{ color: C.navy }}>{c.name}</strong>, c.phone, c.email || "—", c.address, <StatusBadge status={c.status} />, c.join_date])}
          />
        </div>
      </Card>
      {showModal && (
        <Modal title="নতুন ক্লায়েন্ট" onClose={() => setShowModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <FormField label="ফোন *"><input style={inputStyle} value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} /></FormField>
          <FormField label="ইমেইল"><input style={inputStyle} value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} /></FormField>
          <FormField label="ঠিকানা"><input style={inputStyle} value={form.address} onChange={e => setForm({ ...form, address: e.target.value })} /></FormField>
          <button onClick={save} style={btnStyle}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// EMPLOYEES
// ============================================================
function Employees({ data, onAdd }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", role: "", dept: "ডিজাইন", phone: "", salary: "", join_date: "", status: "কর্মরত" });

  const save = async () => {
    if (!form.name || !form.role || !form.salary) return alert("নাম, পদবি ও বেতন আবশ্যক");
    const { error } = await supabase.from("employees").insert([{ ...form, salary: +form.salary }]);
    if (error) return alert("Error: " + error.message);
    onAdd(); setShowModal(false);
    setForm({ name: "", role: "", dept: "ডিজাইন", phone: "", salary: "", join_date: "", status: "কর্মরত" });
  };

  const totalSalary = data.reduce((s, e) => s + (e.salary || 0), 0);
  const handleExport = () => exportToExcel(data.map(e => ({ নাম: e.name, পদবি: e.role, বিভাগ: e.dept, ফোন: e.phone, বেতন: e.salary, যোগদান: e.join_date, স্ট্যাটাস: e.status })), "Employees", "Employees");
  const handlePrint = () => printSection("কর্মী তালিকা", "employees-content");

  return (
    <div>
      <SectionHeader title="কর্মী ব্যবস্থাপনা (HR)" action="নতুন কর্মী" onAction={() => setShowModal(true)} onPrint={handlePrint} onExport={handleExport} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="👷" label="মোট কর্মী" value={fmtNum(data.length)} color="#EEF2FF" />
        <StatCard icon="💰" label="মাসিক বেতন বিল" value={fmt(totalSalary)} color="#FFF8E1" />
        <StatCard icon="🏢" label="বিভাগ" value={fmtNum([...new Set(data.map(e => e.dept))].length)} color="#F0FFF4" />
      </div>
      <Card>
        <div id="employees-content">
          <Table
            headers={["নাম", "পদবি", "বিভাগ", "ফোন", "বেতন", "যোগদান", "স্ট্যাটাস"]}
            rows={data.map(e => [<strong style={{ color: C.navy }}>{e.name}</strong>, e.role, <Badge label={e.dept} color="blue" />, e.phone, <span style={{ fontWeight: 700, color: C.green }}>{fmt(e.salary)}</span>, e.join_date, <StatusBadge status={e.status} />])}
          />
        </div>
      </Card>
      {showModal && (
        <Modal title="নতুন কর্মী" onClose={() => setShowModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <FormField label="পদবি *"><input style={inputStyle} value={form.role} onChange={e => setForm({ ...form, role: e.target.value })} /></FormField>
          <FormField label="বিভাগ"><select style={inputStyle} value={form.dept} onChange={e => setForm({ ...form, dept: e.target.value })}>{["ডিজাইন", "নির্মাণ", "প্রশাসন", "বিপণন"].map(d => <option key={d}>{d}</option>)}</select></FormField>
          <FormField label="ফোন"><input style={inputStyle} value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} /></FormField>
          <FormField label="বেতন (টাকা) *"><input style={inputStyle} type="number" value={form.salary} onChange={e => setForm({ ...form, salary: e.target.value })} /></FormField>
          <FormField label="যোগদান"><input style={inputStyle} type="date" value={form.join_date} onChange={e => setForm({ ...form, join_date: e.target.value })} /></FormField>
          <button onClick={save} style={btnStyle}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// ATTENDANCE
// ============================================================
function Attendance({ employees }) {
  const today = new Date().toISOString().split("T")[0];
  const [selDate, setSelDate] = useState(today);
  const [attData, setAttData] = useState({});
  const [loading, setLoading] = useState(false);

  useEffect(() => { loadAtt(); }, [selDate]);

  const loadAtt = async () => {
    const { data } = await supabase.from("attendance").select("*").eq("date", selDate);
    const map = {};
    (data || []).forEach(a => { map[a.employee_id] = a.status; });
    setAttData(map);
  };

  const setAtt = async (empId, status) => {
    setAttData(prev => ({ ...prev, [empId]: status }));
    const existing = await supabase.from("attendance").select("id").eq("employee_id", empId).eq("date", selDate).single();
    if (existing.data) {
      await supabase.from("attendance").update({ status }).eq("id", existing.data.id);
    } else {
      await supabase.from("attendance").insert([{ employee_id: empId, date: selDate, status }]);
    }
  };

  const statuses = ["উপস্থিত", "অনুপস্থিত", "অর্ধদিন", "ছুটি"];
  const counts = statuses.map(s => ({ s, n: Object.values(attData).filter(v => v === s).length }));

  const handleExport = () => exportToExcel(employees.map(e => ({ নাম: e.name, পদবি: e.role, তারিখ: selDate, উপস্থিতি: attData[e.id] || "চিহ্নিত নয়" })), "Attendance", "Attendance");
  const handlePrint = () => printSection(`উপস্থিতি রিপোর্ট — ${selDate}`, "attendance-content");

  return (
    <div>
      <SectionHeader title="উপস্থিতি ব্যবস্থাপনা" onPrint={handlePrint} onExport={handleExport} />
      <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 20 }}>
        <div style={{ fontWeight: 600, color: C.gray800, fontSize: 14 }}>তারিখ:</div>
        <input type="date" value={selDate} onChange={e => setSelDate(e.target.value)} style={{ ...inputStyle, width: "auto" }} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 20 }}>
        {counts.map(({ s, n }) => (
          <Card key={s} style={{ textAlign: "center", padding: 14 }}>
            <div style={{ fontSize: 22, fontWeight: 700, color: s === "উপস্থিত" ? C.green : s === "অনুপস্থিত" ? C.red : s === "অর্ধদিন" ? "#856404" : C.gray600 }}>{fmtNum(n)}</div>
            <div style={{ fontSize: 12, color: C.gray600, marginTop: 4 }}>{s}</div>
          </Card>
        ))}
      </div>
      <Card>
        <div id="attendance-content">
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: C.gray50 }}>
                <th style={{ padding: "10px 14px", textAlign: "left", color: C.gray600, fontWeight: 600, borderBottom: `2px solid ${C.gray200}` }}>কর্মীর নাম</th>
                <th style={{ padding: "10px 14px", textAlign: "left", color: C.gray600, fontWeight: 600, borderBottom: `2px solid ${C.gray200}` }}>পদবি</th>
                {statuses.map(s => <th key={s} style={{ padding: "10px 14px", textAlign: "center", color: C.gray600, fontWeight: 600, borderBottom: `2px solid ${C.gray200}` }}>{s}</th>)}
              </tr>
            </thead>
            <tbody>
              {employees.map(e => (
                <tr key={e.id} style={{ borderBottom: `1px solid ${C.gray100}` }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.navy }}>{e.name}</td>
                  <td style={{ padding: "10px 14px", color: C.gray600, fontSize: 12 }}>{e.role}</td>
                  {statuses.map(s => (
                    <td key={s} style={{ padding: "6px 14px", textAlign: "center" }}>
                      <input type="radio" name={`att-${e.id}`} checked={attData[e.id] === s} onChange={() => setAtt(e.id, s)} style={{ accentColor: s === "উপস্থিত" ? C.green : s === "অনুপস্থিত" ? C.red : C.gold, width: 16, height: 16, cursor: "pointer" }} />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 12, fontSize: 12, color: C.green, fontWeight: 600 }}>✅ পরিবর্তন স্বয়ংক্রিয়ভাবে সংরক্ষিত হচ্ছে</div>
      </Card>
    </div>
  );
}

// ============================================================
// FINANCE
// ============================================================
function Finance({ data, onAdd }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ date: new Date().toISOString().split("T")[0], type: "আয়", category: "প্রজেক্ট পেমেন্ট", description: "", amount: "", project: "—" });

  const save = async () => {
    if (!form.description || !form.amount) return alert("বিবরণ ও পরিমাণ আবশ্যক");
    const { error } = await supabase.from("transactions").insert([{ ...form, amount: +form.amount }]);
    if (error) return alert("Error: " + error.message);
    onAdd(); setShowModal(false);
    setForm({ date: new Date().toISOString().split("T")[0], type: "আয়", category: "প্রজেক্ট পেমেন্ট", description: "", amount: "", project: "—" });
  };

  const totalIncome = data.filter(t => t.type === "আয়").reduce((s, t) => s + (t.amount || 0), 0);
  const totalExpense = data.filter(t => t.type === "ব্যয়").reduce((s, t) => s + (t.amount || 0), 0);

  const handleExport = () => exportToExcel(data.map(t => ({ তারিখ: t.date, ধরন: t.type, বিভাগ: t.category, বিবরণ: t.description, প্রজেক্ট: t.project, পরিমাণ: t.amount })), "Finance", "Finance");
  const handlePrint = () => printSection("আর্থিক লেনদেন রিপোর্ট", "finance-content");

  return (
    <div>
      <SectionHeader title="আর্থিক ব্যবস্থাপনা" action="নতুন লেনদেন" onAction={() => setShowModal(true)} onPrint={handlePrint} onExport={handleExport} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="💰" label="মোট আয়" value={fmt(totalIncome)} color="#F0FFF4" />
        <StatCard icon="💸" label="মোট ব্যয়" value={fmt(totalExpense)} color="#FFF5F5" />
        <StatCard icon="📊" label="নিট লাভ" value={fmt(totalIncome - totalExpense)} color={totalIncome >= totalExpense ? "#F0FFF4" : "#FFF5F5"} />
      </div>
      <Card>
        <div id="finance-content">
          <Table
            headers={["তারিখ", "ধরন", "বিভাগ", "বিবরণ", "পরিমাণ"]}
            rows={[...data].reverse().map(t => [t.date, <StatusBadge status={t.type} />, t.category, t.description, <span style={{ fontWeight: 700, color: t.type === "আয়" ? C.green : C.red }}>{t.type === "আয়" ? "+" : "-"}{fmt(t.amount)}</span>])}
          />
        </div>
      </Card>
      {showModal && (
        <Modal title="নতুন লেনদেন" onClose={() => setShowModal(false)}>
          <FormField label="তারিখ"><input style={inputStyle} type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} /></FormField>
          <FormField label="ধরন"><select style={inputStyle} value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}><option>আয়</option><option>ব্যয়</option></select></FormField>
          <FormField label="বিভাগ"><select style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{["প্রজেক্ট পেমেন্ট", "কর্মী বেতন", "নির্মাণ সামগ্রী", "অফিস খরচ", "পরিবহন", "বিবিধ"].map(c => <option key={c}>{c}</option>)}</select></FormField>
          <FormField label="বিবরণ *"><input style={inputStyle} value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></FormField>
          <FormField label="পরিমাণ (টাকা) *"><input style={inputStyle} type="number" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} /></FormField>
          <button onClick={save} style={btnStyle}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// SITE PROGRESS
// ============================================================
function SiteProgress({ data, projects, onAdd }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ project: "", date: new Date().toISOString().split("T")[0], work: "", phase: "ফাউন্ডেশন", workers: "", note: "", status: "চলমান" });

  const save = async () => {
    if (!form.project || !form.work) return alert("প্রজেক্ট ও কাজের বিবরণ আবশ্যক");
    const { error } = await supabase.from("site_progress").insert([{ ...form, workers: +form.workers }]);
    if (error) return alert("Error: " + error.message);
    onAdd(); setShowModal(false);
  };

  const handleExport = () => exportToExcel(data.map(s => ({ প্রজেক্ট: s.project, তারিখ: s.date, কাজ: s.work, পর্যায়: s.phase, শ্রমিক: s.workers, নোট: s.note, স্ট্যাটাস: s.status })), "SiteProgress", "Site_Progress");
  const handlePrint = () => printSection("সাইট অগ্রগতি রিপোর্ট", "site-content");

  return (
    <div>
      <SectionHeader title="সাইট অগ্রগতি" action="নতুন আপডেট" onAction={() => setShowModal(true)} onPrint={handlePrint} onExport={handleExport} />
      <div id="site-content" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 16 }}>
        {[...data].reverse().map(s => (
          <Card key={s.id}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
              <Badge label={s.phase} color="blue" />
              <StatusBadge status={s.status} />
            </div>
            <div style={{ fontWeight: 700, color: C.navy, fontSize: 14, marginBottom: 6 }}>{s.work}</div>
            <div style={{ fontSize: 12, color: C.gold, fontWeight: 600, marginBottom: 8 }}>📋 {s.project}</div>
            <div style={{ fontSize: 12, color: C.gray600 }}>📅 {s.date} | 👷 {fmtNum(s.workers)} জন</div>
            {s.note && <div style={{ fontSize: 12, color: C.gray400, marginTop: 8, padding: 8, background: C.gray50, borderRadius: 6 }}>📝 {s.note}</div>}
          </Card>
        ))}
      </div>
      {showModal && (
        <Modal title="সাইট আপডেট" onClose={() => setShowModal(false)}>
          <FormField label="প্রজেক্ট *"><select style={inputStyle} value={form.project} onChange={e => setForm({ ...form, project: e.target.value })}><option value="">— নির্বাচন করুন —</option>{projects.map(p => <option key={p.id} value={p.name}>{p.name}</option>)}</select></FormField>
          <FormField label="তারিখ"><input style={inputStyle} type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} /></FormField>
          <FormField label="কাজের বিবরণ *"><input style={inputStyle} value={form.work} onChange={e => setForm({ ...form, work: e.target.value })} /></FormField>
          <FormField label="পর্যায়"><select style={inputStyle} value={form.phase} onChange={e => setForm({ ...form, phase: e.target.value })}>{["ডিজাইন", "ফাউন্ডেশন", "স্ট্রাকচার", "ওয়াল", "ছাদ", "ফিনিশিং", "ইন্টেরিয়র", "ইন্সটলেশন"].map(p => <option key={p}>{p}</option>)}</select></FormField>
          <FormField label="শ্রমিক সংখ্যা"><input style={inputStyle} type="number" value={form.workers} onChange={e => setForm({ ...form, workers: e.target.value })} /></FormField>
          <FormField label="নোট"><input style={inputStyle} value={form.note} onChange={e => setForm({ ...form, note: e.target.value })} /></FormField>
          <button onClick={save} style={btnStyle}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// MATERIALS
// ============================================================
function Materials({ data, onAdd }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", unit: "পিস", stock: "", min_stock: "", unit_price: "", supplier: "", last_purchase: new Date().toISOString().split("T")[0] });

  const save = async () => {
    if (!form.name || !form.stock) return alert("নাম ও স্টক আবশ্যক");
    const { error } = await supabase.from("materials").insert([{ ...form, stock: +form.stock, min_stock: +form.min_stock, unit_price: +form.unit_price }]);
    if (error) return alert("Error: " + error.message);
    onAdd(); setShowModal(false);
  };

  const handleExport = () => exportToExcel(data.map(m => ({ নাম: m.name, একক: m.unit, স্টক: m.stock, ন্যূনতম: m.min_stock, একক_মূল্য: m.unit_price, মোট_মূল্য: m.stock * m.unit_price, সাপ্লায়ার: m.supplier })), "Materials", "Materials");
  const handlePrint = () => printSection("সামগ্রী ও ইনভেন্টরি রিপোর্ট", "materials-content");

  return (
    <div>
      <SectionHeader title="সামগ্রী ও ইনভেন্টরি" action="নতুন সামগ্রী" onAction={() => setShowModal(true)} onPrint={handlePrint} onExport={handleExport} />
      <Card>
        <div id="materials-content">
          <Table
            headers={["সামগ্রী", "একক", "স্টক", "ন্যূনতম", "একক মূল্য", "মোট মূল্য", "সাপ্লায়ার", "স্ট্যাটাস"]}
            rows={data.map(m => [
              <strong style={{ color: C.navy }}>{m.name}</strong>,
              m.unit,
              <span style={{ fontWeight: 600, color: m.stock < m.min_stock ? C.red : C.gray800 }}>{fmtNum(m.stock)}</span>,
              fmtNum(m.min_stock), fmt(m.unit_price),
              <span style={{ fontWeight: 700, color: C.green }}>{fmt(m.stock * m.unit_price)}</span>,
              m.supplier,
              m.stock < m.min_stock ? <Badge label="কম স্টক ⚠️" color="red" /> : <Badge label="পর্যাপ্ত ✅" color="green" />,
            ])}
          />
        </div>
      </Card>
      {showModal && (
        <Modal title="নতুন সামগ্রী" onClose={() => setShowModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="একক"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{["পিস", "কেজি", "ব্যাগ", "লিটার", "সিএফটি", "ফুট", "মিটার"].map(u => <option key={u}>{u}</option>)}</select></FormField>
            <FormField label="স্টক *"><input style={inputStyle} type="number" value={form.stock} onChange={e => setForm({ ...form, stock: e.target.value })} /></FormField>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="ন্যূনতম স্টক"><input style={inputStyle} type="number" value={form.min_stock} onChange={e => setForm({ ...form, min_stock: e.target.value })} /></FormField>
            <FormField label="একক মূল্য"><input style={inputStyle} type="number" value={form.unit_price} onChange={e => setForm({ ...form, unit_price: e.target.value })} /></FormField>
          </div>
          <FormField label="সাপ্লায়ার"><input style={inputStyle} value={form.supplier} onChange={e => setForm({ ...form, supplier: e.target.value })} /></FormField>
          <button onClick={save} style={btnStyle}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// MAIN APP
// ============================================================
const MENU = [
  { id: "dashboard", icon: "🏠", label: "ড্যাশবোর্ড" },
  { id: "projects", icon: "🏗️", label: "প্রজেক্ট" },
  { id: "clients", icon: "👥", label: "ক্লায়েন্ট" },
  { id: "employees", icon: "👷", label: "কর্মী (HR)" },
  { id: "attendance", icon: "📋", label: "উপস্থিতি" },
  { id: "finance", icon: "💰", label: "আর্থিক" },
  { id: "site", icon: "📍", label: "সাইট প্রগ্রেস" },
  { id: "materials", icon: "📦", label: "সামগ্রী" },
];

export default function App() {
  const [loggedIn, setLoggedIn] = useState(localStorage.getItem("nic_logged_in") === "true");
  const [active, setActive] = useState("dashboard");
  const [sideOpen, setSideOpen] = useState(true);
  const [data, setData] = useState({ projects: [], clients: [], employees: [], transactions: [], siteProgress: [], materials: [] });
  const [loading, setLoading] = useState(true);

  useEffect(() => { if (loggedIn) loadAll(); }, [loggedIn]);

  const loadAll = async () => {
    setLoading(true);
    const [p, c, e, t, s, m] = await Promise.all([
      supabase.from("projects").select("*").order("created_at", { ascending: false }),
      supabase.from("clients").select("*").order("created_at", { ascending: false }),
      supabase.from("employees").select("*").order("created_at", { ascending: false }),
      supabase.from("transactions").select("*").order("created_at", { ascending: false }),
      supabase.from("site_progress").select("*").order("created_at", { ascending: false }),
      supabase.from("materials").select("*").order("created_at", { ascending: false }),
    ]);
    setData({ projects: p.data || [], clients: c.data || [], employees: e.data || [], transactions: t.data || [], siteProgress: s.data || [], materials: m.data || [] });
    setLoading(false);
  };

  const logout = () => { localStorage.removeItem("nic_logged_in"); setLoggedIn(false); };

  if (!loggedIn) return <LoginPage onLogin={() => setLoggedIn(true)} />;

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: "#F1F3F5", fontFamily: "'Hind Siliguri', Arial, sans-serif" }}>
      {/* SIDEBAR */}
      <div style={{ width: sideOpen ? 220 : 64, background: C.navy, display: "flex", flexDirection: "column", transition: "width 0.3s", flexShrink: 0, position: "sticky", top: 0, height: "100vh", overflow: "hidden" }}>
        <div style={{ padding: sideOpen ? "20px 20px 16px" : "20px 12px 16px", borderBottom: `1px solid rgba(255,255,255,0.1)` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 36, height: 36, background: C.gold, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>🏛️</div>
            {sideOpen && <div><div style={{ color: C.gold, fontWeight: 800, fontSize: 14 }}>নকশা</div><div style={{ color: "rgba(255,255,255,0.6)", fontSize: 10 }}>NIC Management</div></div>}
          </div>
        </div>
        <nav style={{ flex: 1, padding: "12px 0", overflowY: "auto" }}>
          {MENU.map(m => (
            <button key={m.id} onClick={() => setActive(m.id)} style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", padding: sideOpen ? "11px 20px" : "11px 0", justifyContent: sideOpen ? "flex-start" : "center", background: active === m.id ? `rgba(201,168,76,0.15)` : "none", border: "none", borderLeft: active === m.id ? `3px solid ${C.gold}` : "3px solid transparent", color: active === m.id ? C.gold : "rgba(255,255,255,0.65)", cursor: "pointer", fontSize: 13, fontWeight: active === m.id ? 700 : 400, fontFamily: "inherit" }}>
              <span style={{ fontSize: 18, flexShrink: 0 }}>{m.icon}</span>
              {sideOpen && <span style={{ whiteSpace: "nowrap" }}>{m.label}</span>}
            </button>
          ))}
        </nav>
        <button onClick={logout} style={{ background: "rgba(220,53,69,0.2)", border: "none", color: "#ff8080", padding: "12px", cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>
          {sideOpen ? "🚪 লগআউট" : "🚪"}
        </button>
        <button onClick={() => setSideOpen(o => !o)} style={{ background: "rgba(255,255,255,0.1)", border: "none", color: C.white, padding: "10px", cursor: "pointer", fontSize: 14, fontFamily: "inherit" }}>
          {sideOpen ? "◀" : "▶"}
        </button>
      </div>

      {/* MAIN */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <div style={{ background: C.white, borderBottom: `1px solid ${C.gray200}`, padding: "14px 24px", display: "flex", justifyContent: "space-between", alignItems: "center", position: "sticky", top: 0, zIndex: 100 }}>
          <div>
            <div style={{ fontWeight: 700, color: C.navy, fontSize: 16 }}>{MENU.find(m => m.id === active)?.label}</div>
            <div style={{ fontSize: 11, color: C.gray400 }}>Noksha Interior & Construction · ফরিদপুর & ঢাকা</div>
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <button onClick={loadAll} style={{ background: C.gray100, border: "none", borderRadius: 8, padding: "6px 12px", cursor: "pointer", fontSize: 12, color: C.gray600 }}>🔄 রিফ্রেশ</button>
            <div style={{ width: 36, height: 36, background: `linear-gradient(135deg, ${C.navy}, ${C.navyMid})`, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", color: C.gold, fontWeight: 700, fontSize: 14 }}>R</div>
            <div style={{ fontSize: 12, color: C.gray600 }}><div style={{ fontWeight: 600, color: C.navy }}>মোঃ রানা</div><div>Admin</div></div>
          </div>
        </div>

        <div style={{ flex: 1, padding: "24px", overflowY: "auto" }}>
          {loading ? (
            <div style={{ textAlign: "center", padding: 60, color: C.gray400 }}>
              <div style={{ fontSize: 40, marginBottom: 16 }}>⏳</div>
              <div>ডেটা লোড হচ্ছে...</div>
            </div>
          ) : (
            <>
              {active === "dashboard" && <Dashboard projects={data.projects} clients={data.clients} employees={data.employees} transactions={data.transactions} materials={data.materials} />}
              {active === "projects" && <Projects data={data.projects} onAdd={loadAll} />}
              {active === "clients" && <Clients data={data.clients} onAdd={loadAll} />}
              {active === "employees" && <Employees data={data.employees} onAdd={loadAll} />}
              {active === "attendance" && <Attendance employees={data.employees} />}
              {active === "finance" && <Finance data={data.transactions} onAdd={loadAll} />}
              {active === "site" && <SiteProgress data={data.siteProgress} projects={data.projects} onAdd={loadAll} />}
              {active === "materials" && <Materials data={data.materials} onAdd={loadAll} />}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
