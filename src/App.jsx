import { useState, useEffect, useRef } from "react";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://jijxnycopycsysugppnw.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImppanhueWNvcHljc3lzdWdwcG53Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI3MDAyNTAsImV4cCI6MjA5ODI3NjI1MH0.7kXRGGnW4VdWU9VT1XEBKp5oC9V5Z21KA_PBqZtvjJA";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// ============================================================
// NOKSHA BRAND COLORS
// ============================================================
const C = {
  primary: "#3F5F45",
  primaryDark: "#2A3F2E",
  primaryLight: "#6B8F6B",
  primaryPale: "#F0F5F1",
  primaryBg: "#E8F0E9",
  white: "#FFFFFF",
  gray50: "#F8F9FA",
  gray100: "#F1F3F5",
  gray200: "#E9ECEF",
  gray400: "#ADB5BD",
  gray600: "#6C757D",
  gray800: "#343A40",
  green: "#28A745",
  greenLight: "#D4EDDA",
  red: "#DC3545",
  redLight: "#F8D7DA",
  yellow: "#FFC107",
  yellowLight: "#FFF3CD",
  blue: "#007BFF",
  blueLight: "#D0E8FF",
};

const fmt = (n) => "৳" + Number(n || 0).toLocaleString("bn-BD");
const fmtNum = (n) => Number(n || 0).toLocaleString("bn-BD");

// ============================================================
// AUTH
// ============================================================
const ADMIN_USER = "admin@noksha.com";
let ADMIN_PASS = localStorage.getItem("nic_password") || "noksha2024";

// ============================================================
// PRINT
// ============================================================
const printSection = (title, contentId) => {
  const content = document.getElementById(contentId);
  if (!content) return;
  const win = window.open("", "_blank");
  win.document.write(`
    <html><head><title>${title} — NIC</title>
    <style>
      body { font-family: Arial, sans-serif; padding: 24px; color: #1A1A1A; }
      .header { text-align: center; border-bottom: 3px solid #3F5F45; padding-bottom: 14px; margin-bottom: 20px; }
      .logo-area { display: flex; align-items: center; justify-content: center; gap: 12px; margin-bottom: 8px; }
      .logo-box { width: 48px; height: 48px; background: #3F5F45; border-radius: 10px; display: flex; align-items: center; justify-content: center; color: white; font-size: 22px; }
      .company-name { font-size: 22px; font-weight: 800; color: #3F5F45; }
      .company-sub { font-size: 12px; color: #6B8F6B; font-weight: 600; }
      .company-info { font-size: 11px; color: #6C757D; margin-top: 4px; }
      .doc-title { font-size: 15px; font-weight: 700; background: #E8F0E9; color: #2A3F2E; padding: 5px 20px; border-radius: 6px; display: inline-block; margin-top: 8px; }
      table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 12px; }
      th { background: #3F5F45; color: white; padding: 9px 12px; text-align: left; }
      td { padding: 8px 12px; border-bottom: 1px solid #E9ECEF; }
      tr:nth-child(even) { background: #F8F9FA; }
      .footer { margin-top: 30px; border-top: 1px solid #3F5F45; padding-top: 10px; font-size: 11px; color: #6C757D; text-align: center; }
    </style></head><body>
    <div class="header">
      <div class="logo-area">
        <div class="logo-box">🏠</div>
        <div>
          <div class="company-name">NOKSHA</div>
          <div class="company-sub">INTERIOR & CONSTRUCTION</div>
        </div>
      </div>
      <div class="company-info">নীলটুলী, ফরিদপুর | পল্লবী, ঢাকা | 📞 01XXX-XXXXXX</div>
      <div class="doc-title">${title}</div>
      <div class="company-info">তারিখ: ${new Date().toLocaleDateString("bn-BD")}</div>
    </div>
    ${content.innerHTML}
    <div class="footer">Noksha Interior & Construction — সকল অধিকার সংরক্ষিত © ${new Date().getFullYear()}</div>
    </body></html>
  `);
  win.document.close();
  setTimeout(() => win.print(), 500);
};

// ============================================================
// EXCEL EXPORT
// ============================================================
const exportToExcel = (data, sheetName, fileName) => {
  if (!data || data.length === 0) return alert("কোনো ডেটা নেই!");
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = Object.keys(data[0]).map(() => ({ wch: 22 }));
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `NIC_${fileName}_${new Date().toISOString().split("T")[0]}.xlsx`);
};

// ============================================================
// EXCEL UPLOAD & PARSE
// ============================================================
const parseExcelFile = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    resolve(XLSX.utils.sheet_to_json(ws));
  };
  reader.onerror = reject;
  reader.readAsBinaryString(file);
});

// ============================================================
// UI COMPONENTS
// ============================================================
function Badge({ label, color }) {
  const colors = {
    green: { bg: C.greenLight, text: C.green },
    red: { bg: C.redLight, text: C.red },
    yellow: { bg: C.yellowLight, text: "#856404" },
    blue: { bg: C.blueLight, text: C.blue },
    gray: { bg: C.gray200, text: C.gray600 },
    primary: { bg: C.primaryBg, text: C.primaryDark },
  };
  const s = colors[color] || colors.gray;
  return <span style={{ background: s.bg, color: s.text, padding: "2px 10px", borderRadius: 20, fontSize: 12, fontWeight: 600 }}>{label}</span>;
}

function StatusBadge({ status }) {
  const map = { "চলমান": "blue", "সম্পন্ন": "green", "নতুন": "primary", "পেন্ডিং": "yellow", "সক্রিয়": "green", "কর্মরত": "green", "উপস্থিত": "green", "অনুপস্থিত": "red", "অর্ধদিন": "yellow", "ছুটি": "gray", "আয়": "green", "ব্যয়": "red" };
  return <Badge label={status} color={map[status] || "gray"} />;
}

function Card({ children, style }) {
  return <div style={{ background: C.white, borderRadius: 12, padding: 20, boxShadow: "0 2px 12px rgba(63,95,69,0.08)", border: `1px solid ${C.gray200}`, ...style }}>{children}</div>;
}

function StatCard({ icon, label, value, sub, color }) {
  return (
    <Card style={{ display: "flex", alignItems: "center", gap: 16 }}>
      <div style={{ width: 52, height: 52, borderRadius: 12, background: color || C.primaryBg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24, flexShrink: 0 }}>{icon}</div>
      <div>
        <div style={{ fontSize: 22, fontWeight: 700, color: C.primaryDark }}>{value}</div>
        <div style={{ fontSize: 13, color: C.gray600, marginTop: 2 }}>{label}</div>
        {sub && <div style={{ fontSize: 12, color: C.primaryLight, marginTop: 2, fontWeight: 600 }}>{sub}</div>}
      </div>
    </Card>
  );
}

function ProgressBar({ value, color }) {
  return <div style={{ background: C.gray200, borderRadius: 99, height: 8, overflow: "hidden" }}><div style={{ width: `${Math.min(value, 100)}%`, height: "100%", background: color || C.primary, borderRadius: 99, transition: "width 0.5s" }} /></div>;
}

function SectionHeader({ title, action, onAction, onPrint, onExport, onUpload, uploadRef }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 8 }}>
      <h2 style={{ margin: 0, color: C.primaryDark, fontSize: 18, fontWeight: 700 }}>{title}</h2>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        {onUpload && <>
          <input ref={uploadRef} type="file" accept=".xlsx,.xls" onChange={onUpload} style={{ display: "none" }} />
          <button onClick={() => uploadRef.current.click()} style={{ background: C.primary, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 12 }}>📤 Excel Upload</button>
        </>}
        {onExport && <button onClick={onExport} style={{ background: C.green, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 12 }}>📊 Excel Download</button>}
        {onPrint && <button onClick={onPrint} style={{ background: C.blue, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 12 }}>🖨️ Print</button>}
        {action && <button onClick={onAction} style={{ background: C.primary, color: C.white, border: "none", borderRadius: 8, padding: "8px 16px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ {action}</button>}
      </div>
    </div>
  );
}

function Modal({ title, onClose, children, size }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 16, width: "100%", maxWidth: size || 520, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 24px", borderBottom: `2px solid ${C.primary}`, position: "sticky", top: 0, background: C.white, zIndex: 1 }}>
          <h3 style={{ margin: 0, color: C.primaryDark, fontSize: 16, fontWeight: 700 }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: C.gray600 }}>×</button>
        </div>
        <div style={{ padding: 24 }}>{children}</div>
      </div>
    </div>
  );
}

function FormField({ label, children }) {
  return <div style={{ marginBottom: 14 }}><label style={{ display: "block", fontSize: 13, fontWeight: 600, color: C.gray800, marginBottom: 5 }}>{label}</label>{children}</div>;
}

const inputStyle = { width: "100%", padding: "9px 12px", border: `1px solid ${C.gray200}`, borderRadius: 8, fontSize: 13, color: C.gray800, outline: "none", boxSizing: "border-box", fontFamily: "inherit", transition: "border 0.2s" };
const btnPrimary = { background: C.primary, color: C.white, border: "none", borderRadius: 10, padding: "11px", fontWeight: 700, cursor: "pointer", fontSize: 14, fontFamily: "inherit", width: "100%", marginTop: 8 };
const btnDanger = { background: C.red, color: C.white, border: "none", borderRadius: 8, padding: "6px 12px", fontWeight: 600, cursor: "pointer", fontSize: 12, fontFamily: "inherit" };
const btnEdit = { background: C.primaryBg, color: C.primaryDark, border: `1px solid ${C.primaryLight}`, borderRadius: 8, padding: "5px 10px", fontWeight: 600, cursor: "pointer", fontSize: 12, fontFamily: "inherit" };

// ============================================================
// LOGIN
// ============================================================
function LoginPage({ onLogin }) {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    setLoading(true); setError("");
    if (!email || !pass) { setError("❌ ইমেইল ও পাসওয়ার্ড দিন!"); setLoading(false); return; }
    // Check Supabase app_users table
    const { data, error: dbErr } = await supabase.from("app_users").select("*").eq("email", email.trim().toLowerCase()).eq("password_hash", pass).eq("is_active", true).single();
    if (dbErr || !data) {
      // Fallback: check admin local password
      ADMIN_PASS = localStorage.getItem("nic_password") || "noksha2024";
      if (email === ADMIN_USER && pass === ADMIN_PASS) {
        localStorage.setItem("nic_logged_in", "true");
        localStorage.setItem("nic_user", JSON.stringify({ email, name: "মোঃ রানা", role: "admin", assigned_projects: [] }));
        onLogin({ email, name: "মোঃ রানা", role: "admin", assigned_projects: [] });
      } else {
        setError("❌ ভুল ইমেইল বা পাসওয়ার্ড!");
      }
      setLoading(false); return;
    }
    localStorage.setItem("nic_logged_in", "true");
    localStorage.setItem("nic_user", JSON.stringify(data));
    onLogin(data);
    setLoading(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: `linear-gradient(135deg, ${C.primaryDark} 0%, ${C.primary} 60%, ${C.primaryLight} 100%)`, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 20, padding: 40, width: "100%", maxWidth: 400, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ width: 70, height: 70, background: C.primary, borderRadius: 18, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px", fontSize: 36 }}>🏠</div>
          <div style={{ fontSize: 28, fontWeight: 900, color: C.primary, letterSpacing: 2 }}>NOKSHA</div>
          <div style={{ fontSize: 12, color: C.primaryLight, fontWeight: 700, letterSpacing: 3, marginTop: 2 }}>INTERIOR & CONSTRUCTION</div>
          <div style={{ fontSize: 12, color: C.gray400, marginTop: 8 }}>Management System v3.0</div>
        </div>
        {error && <div style={{ background: C.redLight, color: C.red, padding: "10px 14px", borderRadius: 8, fontSize: 13, marginBottom: 16, textAlign: "center" }}>{error}</div>}
        <FormField label="ইমেইল"><input style={inputStyle} type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="your@email.com" onKeyDown={e => e.key === "Enter" && handleLogin()} /></FormField>
        <FormField label="পাসওয়ার্ড"><input style={inputStyle} type="password" value={pass} onChange={e => setPass(e.target.value)} placeholder="••••••••" onKeyDown={e => e.key === "Enter" && handleLogin()} /></FormField>
        <button onClick={handleLogin} disabled={loading} style={{ ...btnPrimary, background: loading ? C.gray400 : C.primary }}>
          {loading ? "লগইন হচ্ছে..." : "🔐 লগইন করুন"}
        </button>
        <div style={{ textAlign: "center", marginTop: 14, fontSize: 11, color: C.gray400 }}>Noksha Interior & Construction © {new Date().getFullYear()}</div>
      </div>
    </div>
  );
}

// ============================================================
// ANALYTICS
// ============================================================
function Analytics({ transactions, projects, employees }) {
  const income = transactions.filter(t => t.type === "আয়").reduce((s, t) => s + (t.amount || 0), 0);
  const expense = transactions.filter(t => t.type === "ব্যয়").reduce((s, t) => s + (t.amount || 0), 0);
  const profit = income - expense;
  const profitPct = income > 0 ? Math.round((profit / income) * 100) : 0;
  const completed = projects.filter(p => p.status === "সম্পন্ন").length;
  const active = projects.filter(p => p.status === "চলমান").length;
  const totalBudget = projects.reduce((s, p) => s + (p.budget || 0), 0);
  const totalSpent = projects.reduce((s, p) => s + (p.spent || 0), 0);
  const salaryBill = employees.reduce((s, e) => s + (e.salary || 0), 0);

  const catMap = {};
  transactions.forEach(t => { catMap[t.category] = (catMap[t.category] || 0) + (t.amount || 0); });
  const cats = Object.entries(catMap).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const maxCat = cats[0]?.[1] || 1;

  const barColors = [C.primary, C.primaryLight, "#5B8F5F", "#7FAF83", "#A3C9A7"];

  return (
    <div>
      <h2 style={{ color: C.primaryDark, fontSize: 18, fontWeight: 700, marginBottom: 18 }}>রিপোর্ট ও বিশ্লেষণ</h2>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 20 }}>
        <StatCard icon="💰" label="মোট আয়" value={fmt(income)} color="#E8F5E9" />
        <StatCard icon="💸" label="মোট ব্যয়" value={fmt(expense)} color="#FFEBEE" />
        <StatCard icon="📈" label="নিট লাভ" value={fmt(profit)} color={profit >= 0 ? "#E8F5E9" : "#FFEBEE"} />
        <StatCard icon="📊" label="লাভের হার" value={`${profitPct}%`} color={C.primaryBg} />
        <StatCard icon="👷" label="মাসিক বেতন" value={fmt(salaryBill)} color="#FFF8E1" />
        <StatCard icon="🏗️" label="সম্পন্ন প্রজেক্ট" value={fmtNum(completed)} sub={`${active}টি চলমান`} color={C.primaryBg} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <Card>
          <h3 style={{ margin: "0 0 16px", color: C.primaryDark, fontSize: 14, fontWeight: 700 }}>আয় বনাম ব্যয়</h3>
          <div style={{ display: "flex", gap: 12, alignItems: "flex-end", height: 140, paddingBottom: 8 }}>
            {[{ label: "আয়", val: income, color: C.green }, { label: "ব্যয়", val: expense, color: C.red }, { label: "লাভ", val: Math.max(profit, 0), color: C.primary }].map((item, i) => {
              const max = Math.max(income, expense) || 1;
              const h = Math.round((item.val / max) * 120);
              return (
                <div key={i} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                  <div style={{ fontSize: 10, color: C.gray600, fontWeight: 600 }}>{fmt(item.val)}</div>
                  <div style={{ width: "100%", height: h, background: item.color, borderRadius: "4px 4px 0 0", minHeight: 4 }} />
                  <div style={{ fontSize: 11, color: C.gray600 }}>{item.label}</div>
                </div>
              );
            })}
          </div>
        </Card>

        <Card>
          <h3 style={{ margin: "0 0 16px", color: C.primaryDark, fontSize: 14, fontWeight: 700 }}>প্রজেক্ট বাজেট ব্যবহার</h3>
          <div style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ fontSize: 12, color: C.gray600 }}>মোট বাজেট</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark }}>{fmt(totalBudget)}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ fontSize: 12, color: C.gray600 }}>ব্যয় হয়েছে</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.red }}>{fmt(totalSpent)}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
              <span style={{ fontSize: 12, color: C.gray600 }}>বাকি আছে</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.green }}>{fmt(totalBudget - totalSpent)}</span>
            </div>
            <ProgressBar value={totalBudget > 0 ? Math.round((totalSpent / totalBudget) * 100) : 0} color={totalSpent > totalBudget ? C.red : C.primary} />
            <div style={{ fontSize: 11, color: C.gray400, marginTop: 4, textAlign: "right" }}>{totalBudget > 0 ? Math.round((totalSpent / totalBudget) * 100) : 0}% ব্যবহৃত</div>
          </div>
        </Card>
      </div>

      <Card>
        <h3 style={{ margin: "0 0 16px", color: C.primaryDark, fontSize: 14, fontWeight: 700 }}>বিভাগ অনুযায়ী ব্যয়</h3>
        {cats.map(([cat, val], i) => (
          <div key={i} style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
              <span style={{ fontSize: 13, color: C.gray800 }}>{cat}</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark }}>{fmt(val)}</span>
            </div>
            <div style={{ background: C.gray100, borderRadius: 99, height: 8, overflow: "hidden" }}>
              <div style={{ width: `${Math.round((val / maxCat) * 100)}%`, height: "100%", background: barColors[i % barColors.length], borderRadius: 99 }} />
            </div>
          </div>
        ))}
        {cats.length === 0 && <div style={{ color: C.gray400, fontSize: 13 }}>কোনো ডেটা নেই</div>}
      </Card>
    </div>
  );
}

// ============================================================
// PASSWORD CHANGE
// ============================================================
function PasswordChange() {
  const [form, setForm] = useState({ current: "", newPass: "", confirm: "" });
  const [msg, setMsg] = useState("");
  const [error, setError] = useState("");

  const handleChange = () => {
    setMsg(""); setError("");
    const currentPass = localStorage.getItem("nic_password") || "noksha2024";
    if (form.current !== currentPass) return setError("❌ বর্তমান পাসওয়ার্ড ভুল!");
    if (form.newPass.length < 6) return setError("❌ নতুন পাসওয়ার্ড কমপক্ষে ৬ অক্ষরের হতে হবে!");
    if (form.newPass !== form.confirm) return setError("❌ নতুন পাসওয়ার্ড মিলছে না!");
    localStorage.setItem("nic_password", form.newPass);
    ADMIN_PASS = form.newPass;
    setMsg("✅ পাসওয়ার্ড সফলভাবে পরিবর্তন হয়েছে!");
    setForm({ current: "", newPass: "", confirm: "" });
  };

  return (
    <div>
      <h2 style={{ color: C.primaryDark, fontSize: 18, fontWeight: 700, marginBottom: 18 }}>পাসওয়ার্ড পরিবর্তন</h2>
      <Card style={{ maxWidth: 450 }}>
        <div style={{ marginBottom: 20, padding: "12px 16px", background: C.primaryBg, borderRadius: 10, fontSize: 13, color: C.primaryDark }}>
          🔐 আপনার অ্যাকাউন্ট: <strong>{ADMIN_USER}</strong>
        </div>
        {error && <div style={{ background: C.redLight, color: C.red, padding: "10px", borderRadius: 8, fontSize: 13, marginBottom: 14 }}>{error}</div>}
        {msg && <div style={{ background: C.greenLight, color: C.green, padding: "10px", borderRadius: 8, fontSize: 13, marginBottom: 14 }}>{msg}</div>}
        <FormField label="বর্তমান পাসওয়ার্ড"><input style={inputStyle} type="password" value={form.current} onChange={e => setForm({ ...form, current: e.target.value })} placeholder="বর্তমান পাসওয়ার্ড" /></FormField>
        <FormField label="নতুন পাসওয়ার্ড"><input style={inputStyle} type="password" value={form.newPass} onChange={e => setForm({ ...form, newPass: e.target.value })} placeholder="কমপক্ষে ৬ অক্ষর" /></FormField>
        <FormField label="নতুন পাসওয়ার্ড নিশ্চিত করুন"><input style={inputStyle} type="password" value={form.confirm} onChange={e => setForm({ ...form, confirm: e.target.value })} placeholder="পুনরায় টাইপ করুন" /></FormField>
        <button onClick={handleChange} style={btnPrimary}>🔑 পাসওয়ার্ড পরিবর্তন করুন</button>
      </Card>
    </div>
  );
}

// ============================================================
// PROJECTS
// ============================================================
function Projects({ data, onRefresh }) {
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ name: "", client: "", budget: "", start_date: "", end_date: "", location: "", type: "আবাসিক নির্মাণ", status: "নতুন", progress: 0, spent: 0 });
  const uploadRef = useRef();

  const openAdd = () => { setEditItem(null); setForm({ name: "", client: "", budget: "", start_date: "", end_date: "", location: "", type: "আবাসিক নির্মাণ", status: "নতুন", progress: 0, spent: 0 }); setShowModal(true); };
  const openEdit = (item) => { setEditItem(item); setForm({ ...item, budget: item.budget || "", spent: item.spent || "", progress: item.progress || 0 }); setShowModal(true); };

  const save = async () => {
    if (!form.name || !form.client) return alert("নাম ও ক্লায়েন্ট আবশ্যক");
    const payload = { ...form, budget: +form.budget || 0, spent: +form.spent || 0, progress: +form.progress || 0 };
    if (editItem) {
      await supabase.from("projects").update(payload).eq("id", editItem.id);
    } else {
      await supabase.from("projects").insert([payload]);
    }
    onRefresh(); setShowModal(false);
  };

  const deleteItem = async (id) => {
    if (!confirm("এই প্রজেক্ট মুছে ফেলবেন?")) return;
    await supabase.from("projects").delete().eq("id", id);
    onRefresh();
  };

  const handleExport = () => exportToExcel(data.map(p => ({ নাম: p.name, ক্লায়েন্ট: p.client, বাজেট: p.budget, খরচ: p.spent, অগ্রগতি: p.progress + "%", স্ট্যাটাস: p.status, অবস্থান: p.location, ধরন: p.type, শুরু: p.start_date, শেষ: p.end_date })), "Projects", "Projects");

  const handleUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      if (!row["নাম"]) continue;
      await supabase.from("projects").upsert([{ name: row["নাম"], client: row["ক্লায়েন্ট"] || "", budget: +row["বাজেট"] || 0, spent: +row["খরচ"] || 0, progress: parseInt(row["অগ্রগতি"]) || 0, status: row["স্ট্যাটাস"] || "নতুন", location: row["অবস্থান"] || "", type: row["ধরন"] || "আবাসিক নির্মাণ" }]);
      count++;
    }
    alert(`✅ ${count}টি প্রজেক্ট আপলোড হয়েছে!`);
    onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <SectionHeader title="প্রজেক্ট ব্যবস্থাপনা" action="নতুন প্রজেক্ট" onAction={openAdd} onExport={handleExport} onPrint={() => printSection("প্রজেক্ট তালিকা", "projects-content")} onUpload={handleUpload} uploadRef={uploadRef} />
      <div id="projects-content" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 16 }}>
        {data.map(p => (
          <Card key={p.id}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
              <div style={{ flex: 1, marginRight: 8 }}>
                <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14 }}>{p.name}</div>
                <div style={{ fontSize: 11, color: C.gray400, marginTop: 2 }}>📍 {p.location}</div>
              </div>
              <StatusBadge status={p.status} />
            </div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}>ক্লায়েন্ট: <strong>{p.client}</strong></div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 10 }}>ধরন: {p.type}</div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
              <span style={{ fontSize: 12, color: C.gray600 }}>অগ্রগতি</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.primary }}>{p.progress}%</span>
            </div>
            <ProgressBar value={p.progress} color={p.progress === 100 ? C.green : C.primary} />
            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 12, paddingTop: 12, borderTop: `1px solid ${C.gray100}` }}>
              <div><div style={{ fontSize: 11, color: C.gray400 }}>বাজেট</div><div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 13 }}>{fmt(p.budget)}</div></div>
              <div style={{ textAlign: "right" }}><div style={{ fontSize: 11, color: C.gray400 }}>খরচ</div><div style={{ fontWeight: 700, color: C.green, fontSize: 13 }}>{fmt(p.spent)}</div></div>
            </div>
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button onClick={() => openEdit(p)} style={{ ...btnEdit, flex: 1 }}>✏️ Edit</button>
              <button onClick={() => deleteItem(p.id)} style={{ ...btnDanger }}>🗑️</button>
            </div>
          </Card>
        ))}
        {data.length === 0 && <div style={{ color: C.gray400, fontSize: 13, padding: 20 }}>কোনো প্রজেক্ট নেই। নতুন যোগ করুন!</div>}
      </div>

      {showModal && (
        <Modal title={editItem ? "প্রজেক্ট সম্পাদনা" : "নতুন প্রজেক্ট"} onClose={() => setShowModal(false)}>
          <FormField label="প্রজেক্টের নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <FormField label="ক্লায়েন্ট *"><input style={inputStyle} value={form.client} onChange={e => setForm({ ...form, client: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="বাজেট (৳)"><input style={inputStyle} type="number" value={form.budget} onChange={e => setForm({ ...form, budget: e.target.value })} /></FormField>
            <FormField label="খরচ (৳)"><input style={inputStyle} type="number" value={form.spent} onChange={e => setForm({ ...form, spent: e.target.value })} /></FormField>
          </div>
          <FormField label="অগ্রগতি (%)"><input style={inputStyle} type="number" min="0" max="100" value={form.progress} onChange={e => setForm({ ...form, progress: e.target.value })} /></FormField>
          <FormField label="অবস্থান"><input style={inputStyle} value={form.location} onChange={e => setForm({ ...form, location: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="শুরুর তারিখ"><input style={inputStyle} type="date" value={form.start_date} onChange={e => setForm({ ...form, start_date: e.target.value })} /></FormField>
            <FormField label="শেষের তারিখ"><input style={inputStyle} type="date" value={form.end_date} onChange={e => setForm({ ...form, end_date: e.target.value })} /></FormField>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="ধরন"><select style={inputStyle} value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}>{["আবাসিক নির্মাণ", "কমার্শিয়াল ইন্টেরিয়র", "মডুলার কিচেন", "থ্রিডি ভিজুয়ালাইজেশন", "সিভিল কনস্ট্রাকশন"].map(t => <option key={t}>{t}</option>)}</select></FormField>
            <FormField label="স্ট্যাটাস"><select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>{["নতুন", "চলমান", "সম্পন্ন", "বাতিল"].map(s => <option key={s}>{s}</option>)}</select></FormField>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// CLIENTS
// ============================================================
function Clients({ data, onRefresh }) {
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ name: "", phone: "", email: "", address: "", status: "সক্রিয়" });
  const uploadRef = useRef();

  const openAdd = () => { setEditItem(null); setForm({ name: "", phone: "", email: "", address: "", status: "সক্রিয়" }); setShowModal(true); };
  const openEdit = (item) => { setEditItem(item); setForm({ ...item }); setShowModal(true); };

  const save = async () => {
    if (!form.name || !form.phone) return alert("নাম ও ফোন আবশ্যক");
    if (editItem) {
      await supabase.from("clients").update(form).eq("id", editItem.id);
    } else {
      await supabase.from("clients").insert([{ ...form, join_date: new Date().toISOString().split("T")[0] }]);
    }
    onRefresh(); setShowModal(false);
  };

  const deleteItem = async (id) => {
    if (!confirm("এই ক্লায়েন্ট মুছে ফেলবেন?")) return;
    await supabase.from("clients").delete().eq("id", id);
    onRefresh();
  };

  const handleExport = () => exportToExcel(data.map(c => ({ নাম: c.name, ফোন: c.phone, ইমেইল: c.email, ঠিকানা: c.address, স্ট্যাটাস: c.status, যোগদান: c.join_date })), "Clients", "Clients");

  const handleUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      if (!row["নাম"]) continue;
      await supabase.from("clients").insert([{ name: row["নাম"], phone: row["ফোন"] || "", email: row["ইমেইল"] || "", address: row["ঠিকানা"] || "", status: row["স্ট্যাটাস"] || "সক্রিয়", join_date: new Date().toISOString().split("T")[0] }]);
      count++;
    }
    alert(`✅ ${count}টি ক্লায়েন্ট আপলোড হয়েছে!`); onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <SectionHeader title="ক্লায়েন্ট ব্যবস্থাপনা" action="নতুন ক্লায়েন্ট" onAction={openAdd} onExport={handleExport} onPrint={() => printSection("ক্লায়েন্ট তালিকা", "clients-content")} onUpload={handleUpload} uploadRef={uploadRef} />
      <Card>
        <div id="clients-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["নাম", "ফোন", "ইমেইল", "ঠিকানা", "স্ট্যাটাস", "যোগদান", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>{h}</th>)}</tr></thead>
            <tbody>
              {data.map(c => (
                <tr key={c.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "10px 14px" }}><strong style={{ color: C.primaryDark }}>{c.name}</strong></td>
                  <td style={{ padding: "10px 14px" }}>{c.phone}</td>
                  <td style={{ padding: "10px 14px" }}>{c.email || "—"}</td>
                  <td style={{ padding: "10px 14px" }}>{c.address}</td>
                  <td style={{ padding: "10px 14px" }}><StatusBadge status={c.status} /></td>
                  <td style={{ padding: "10px 14px" }}>{c.join_date}</td>
                  <td style={{ padding: "10px 14px" }}><div style={{ display: "flex", gap: 6 }}><button onClick={() => openEdit(c)} style={btnEdit}>✏️</button><button onClick={() => deleteItem(c.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "ক্লায়েন্ট সম্পাদনা" : "নতুন ক্লায়েন্ট"} onClose={() => setShowModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <FormField label="ফোন *"><input style={inputStyle} value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} /></FormField>
          <FormField label="ইমেইল"><input style={inputStyle} value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} /></FormField>
          <FormField label="ঠিকানা"><input style={inputStyle} value={form.address} onChange={e => setForm({ ...form, address: e.target.value })} /></FormField>
          <FormField label="স্ট্যাটাস"><select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>{["সক্রিয়", "নিষ্ক্রিয়", "সম্পন্ন"].map(s => <option key={s}>{s}</option>)}</select></FormField>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// EMPLOYEES
// ============================================================
function Employees({ data, onRefresh }) {
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ name: "", role: "", dept: "ডিজাইন", phone: "", salary: "", join_date: "", status: "কর্মরত" });
  const uploadRef = useRef();

  const openAdd = () => { setEditItem(null); setForm({ name: "", role: "", dept: "ডিজাইন", phone: "", salary: "", join_date: "", status: "কর্মরত" }); setShowModal(true); };
  const openEdit = (item) => { setEditItem(item); setForm({ ...item, salary: item.salary || "" }); setShowModal(true); };

  const save = async () => {
    if (!form.name || !form.role) return alert("নাম ও পদবি আবশ্যক");
    const payload = { ...form, salary: +form.salary || 0 };
    if (editItem) { await supabase.from("employees").update(payload).eq("id", editItem.id); }
    else { await supabase.from("employees").insert([payload]); }
    onRefresh(); setShowModal(false);
  };

  const deleteItem = async (id) => {
    if (!confirm("এই কর্মী মুছে ফেলবেন?")) return;
    await supabase.from("employees").delete().eq("id", id); onRefresh();
  };

  const handleExport = () => exportToExcel(data.map(e => ({ নাম: e.name, পদবি: e.role, বিভাগ: e.dept, ফোন: e.phone, বেতন: e.salary, যোগদান: e.join_date, স্ট্যাটাস: e.status })), "Employees", "Employees");
  const totalSalary = data.reduce((s, e) => s + (e.salary || 0), 0);

  const handleUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      if (!row["নাম"]) continue;
      await supabase.from("employees").insert([{ name: row["নাম"], role: row["পদবি"] || "", dept: row["বিভাগ"] || "ডিজাইন", phone: row["ফোন"] || "", salary: +row["বেতন"] || 0, join_date: row["যোগদান"] || "", status: row["স্ট্যাটাস"] || "কর্মরত" }]);
      count++;
    }
    alert(`✅ ${count}জন কর্মী আপলোড হয়েছে!`); onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="👷" label="মোট কর্মী" value={fmtNum(data.length)} color={C.primaryBg} />
        <StatCard icon="💰" label="মাসিক বেতন" value={fmt(totalSalary)} color="#FFF8E1" />
        <StatCard icon="🏢" label="বিভাগ" value={fmtNum([...new Set(data.map(e => e.dept))].length)} color="#F0FFF4" />
      </div>
      <SectionHeader title="কর্মী ব্যবস্থাপনা (HR)" action="নতুন কর্মী" onAction={openAdd} onExport={handleExport} onPrint={() => printSection("কর্মী তালিকা", "employees-content")} onUpload={handleUpload} uploadRef={uploadRef} />
      <Card>
        <div id="employees-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["নাম", "পদবি", "বিভাগ", "ফোন", "বেতন", "যোগদান", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>{h}</th>)}</tr></thead>
            <tbody>
              {data.map(e => (
                <tr key={e.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={ev => ev.currentTarget.style.background = C.primaryBg} onMouseLeave={ev => ev.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "10px 14px" }}><strong style={{ color: C.primaryDark }}>{e.name}</strong></td>
                  <td style={{ padding: "10px 14px" }}>{e.role}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={e.dept} color="primary" /></td>
                  <td style={{ padding: "10px 14px" }}>{e.phone}</td>
                  <td style={{ padding: "10px 14px" }}><span style={{ fontWeight: 700, color: C.green }}>{fmt(e.salary)}</span></td>
                  <td style={{ padding: "10px 14px" }}>{e.join_date}</td>
                  <td style={{ padding: "10px 14px" }}><StatusBadge status={e.status} /></td>
                  <td style={{ padding: "10px 14px" }}><div style={{ display: "flex", gap: 6 }}><button onClick={() => openEdit(e)} style={btnEdit}>✏️</button><button onClick={() => deleteItem(e.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "কর্মী সম্পাদনা" : "নতুন কর্মী"} onClose={() => setShowModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <FormField label="পদবি *"><input style={inputStyle} value={form.role} onChange={e => setForm({ ...form, role: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="বিভাগ"><select style={inputStyle} value={form.dept} onChange={e => setForm({ ...form, dept: e.target.value })}>{["ডিজাইন", "নির্মাণ", "প্রশাসন", "বিপণন"].map(d => <option key={d}>{d}</option>)}</select></FormField>
            <FormField label="স্ট্যাটাস"><select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>{["কর্মরত", "ছুটিতে", "বরখাস্ত"].map(s => <option key={s}>{s}</option>)}</select></FormField>
          </div>
          <FormField label="ফোন"><input style={inputStyle} value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="বেতন (৳)"><input style={inputStyle} type="number" value={form.salary} onChange={e => setForm({ ...form, salary: e.target.value })} /></FormField>
            <FormField label="যোগদান"><input style={inputStyle} type="date" value={form.join_date} onChange={e => setForm({ ...form, join_date: e.target.value })} /></FormField>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
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
    if (existing.data) { await supabase.from("attendance").update({ status }).eq("id", existing.data.id); }
    else { await supabase.from("attendance").insert([{ employee_id: empId, date: selDate, status }]); }
  };

  const statuses = ["উপস্থিত", "অনুপস্থিত", "অর্ধদিন", "ছুটি"];
  const counts = statuses.map(s => ({ s, n: Object.values(attData).filter(v => v === s).length }));
  const handlePrint = () => printSection(`উপস্থিতি — ${selDate}`, "attendance-content");
  const handleExport = () => exportToExcel(employees.map(e => ({ নাম: e.name, পদবি: e.role, তারিখ: selDate, উপস্থিতি: attData[e.id] || "চিহ্নিত নয়" })), "Attendance", "Attendance");

  return (
    <div>
      <SectionHeader title="উপস্থিতি ব্যবস্থাপনা" onPrint={handlePrint} onExport={handleExport} />
      <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 16 }}>
        <div style={{ fontWeight: 600, color: C.gray800, fontSize: 14 }}>তারিখ:</div>
        <input type="date" value={selDate} onChange={e => setSelDate(e.target.value)} style={{ ...inputStyle, width: "auto" }} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginBottom: 16 }}>
        {counts.map(({ s, n }) => (
          <Card key={s} style={{ textAlign: "center", padding: 14 }}>
            <div style={{ fontSize: 22, fontWeight: 700, color: s === "উপস্থিত" ? C.green : s === "অনুপস্থিত" ? C.red : s === "অর্ধদিন" ? "#856404" : C.gray600 }}>{fmtNum(n)}</div>
            <div style={{ fontSize: 12, color: C.gray600, marginTop: 4 }}>{s}</div>
          </Card>
        ))}
      </div>
      <Card>
        <div id="attendance-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>
              <th style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>কর্মীর নাম</th>
              <th style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>পদবি</th>
              {statuses.map(s => <th key={s} style={{ padding: "10px 14px", textAlign: "center", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>{s}</th>)}
            </tr></thead>
            <tbody>
              {employees.map(e => (
                <tr key={e.id} style={{ borderBottom: `1px solid ${C.gray100}` }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{e.name}</td>
                  <td style={{ padding: "10px 14px", color: C.gray600, fontSize: 12 }}>{e.role}</td>
                  {statuses.map(s => (
                    <td key={s} style={{ padding: "6px 14px", textAlign: "center" }}>
                      <input type="radio" name={`att-${e.id}`} checked={attData[e.id] === s} onChange={() => setAtt(e.id, s)} style={{ accentColor: s === "উপস্থিত" ? C.green : s === "অনুপস্থিত" ? C.red : C.primary, width: 16, height: 16, cursor: "pointer" }} />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 10, fontSize: 12, color: C.green, fontWeight: 600 }}>✅ পরিবর্তন স্বয়ংক্রিয়ভাবে সংরক্ষিত হচ্ছে</div>
      </Card>
    </div>
  );
}

// ============================================================
// FINANCE
// ============================================================
function Finance({ data, onRefresh }) {
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ date: new Date().toISOString().split("T")[0], type: "আয়", category: "প্রজেক্ট পেমেন্ট", description: "", amount: "", project: "—" });
  const uploadRef = useRef();

  const openAdd = () => { setEditItem(null); setForm({ date: new Date().toISOString().split("T")[0], type: "আয়", category: "প্রজেক্ট পেমেন্ট", description: "", amount: "", project: "—" }); setShowModal(true); };
  const openEdit = (item) => { setEditItem(item); setForm({ ...item, amount: item.amount || "" }); setShowModal(true); };

  const save = async () => {
    if (!form.description || !form.amount) return alert("বিবরণ ও পরিমাণ আবশ্যক");
    const payload = { ...form, amount: +form.amount };
    if (editItem) { await supabase.from("transactions").update(payload).eq("id", editItem.id); }
    else { await supabase.from("transactions").insert([payload]); }
    onRefresh(); setShowModal(false);
  };

  const deleteItem = async (id) => {
    if (!confirm("এই লেনদেন মুছে ফেলবেন?")) return;
    await supabase.from("transactions").delete().eq("id", id); onRefresh();
  };

  const totalIncome = data.filter(t => t.type === "আয়").reduce((s, t) => s + (t.amount || 0), 0);
  const totalExpense = data.filter(t => t.type === "ব্যয়").reduce((s, t) => s + (t.amount || 0), 0);

  const handleExport = () => exportToExcel(data.map(t => ({ তারিখ: t.date, ধরন: t.type, বিভাগ: t.category, বিবরণ: t.description, প্রজেক্ট: t.project, পরিমাণ: t.amount })), "Finance", "Finance");
  const handleUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      if (!row["বিবরণ"] || !row["পরিমাণ"]) continue;
      await supabase.from("transactions").insert([{ date: row["তারিখ"] || new Date().toISOString().split("T")[0], type: row["ধরন"] || "আয়", category: row["বিভাগ"] || "বিবিধ", description: row["বিবরণ"], amount: +row["পরিমাণ"] || 0, project: row["প্রজেক্ট"] || "—" }]);
      count++;
    }
    alert(`✅ ${count}টি লেনদেন আপলোড হয়েছে!`); onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="💰" label="মোট আয়" value={fmt(totalIncome)} color="#F0FFF4" />
        <StatCard icon="💸" label="মোট ব্যয়" value={fmt(totalExpense)} color="#FFF5F5" />
        <StatCard icon="📊" label="নিট লাভ" value={fmt(totalIncome - totalExpense)} color={totalIncome >= totalExpense ? "#F0FFF4" : "#FFF5F5"} />
      </div>
      <SectionHeader title="আর্থিক ব্যবস্থাপনা" action="নতুন লেনদেন" onAction={openAdd} onExport={handleExport} onPrint={() => printSection("আর্থিক রিপোর্ট", "finance-content")} onUpload={handleUpload} uploadRef={uploadRef} />
      <Card>
        <div id="finance-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "ধরন", "বিভাগ", "বিবরণ", "পরিমাণ", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>{h}</th>)}</tr></thead>
            <tbody>
              {[...data].reverse().map(t => (
                <tr key={t.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "10px 14px" }}>{t.date}</td>
                  <td style={{ padding: "10px 14px" }}><StatusBadge status={t.type} /></td>
                  <td style={{ padding: "10px 14px" }}>{t.category}</td>
                  <td style={{ padding: "10px 14px" }}>{t.description}</td>
                  <td style={{ padding: "10px 14px" }}><span style={{ fontWeight: 700, color: t.type === "আয়" ? C.green : C.red }}>{t.type === "আয়" ? "+" : "-"}{fmt(t.amount)}</span></td>
                  <td style={{ padding: "10px 14px" }}><div style={{ display: "flex", gap: 6 }}><button onClick={() => openEdit(t)} style={btnEdit}>✏️</button><button onClick={() => deleteItem(t.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "লেনদেন সম্পাদনা" : "নতুন লেনদেন"} onClose={() => setShowModal(false)}>
          <FormField label="তারিখ"><input style={inputStyle} type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="ধরন"><select style={inputStyle} value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}><option>আয়</option><option>ব্যয়</option></select></FormField>
            <FormField label="বিভাগ"><select style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{["প্রজেক্ট পেমেন্ট", "কর্মী বেতন", "নির্মাণ সামগ্রী", "অফিস খরচ", "পরিবহন", "বিবিধ"].map(c => <option key={c}>{c}</option>)}</select></FormField>
          </div>
          <FormField label="বিবরণ *"><input style={inputStyle} value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></FormField>
          <FormField label="পরিমাণ (৳) *"><input style={inputStyle} type="number" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} /></FormField>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// MATERIALS
// ============================================================
function Materials({ data, onRefresh }) {
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ name: "", unit: "পিস", stock: "", min_stock: "", unit_price: "", supplier: "", last_purchase: new Date().toISOString().split("T")[0] });
  const uploadRef = useRef();

  const openAdd = () => { setEditItem(null); setForm({ name: "", unit: "পিস", stock: "", min_stock: "", unit_price: "", supplier: "", last_purchase: new Date().toISOString().split("T")[0] }); setShowModal(true); };
  const openEdit = (item) => { setEditItem(item); setForm({ ...item }); setShowModal(true); };

  const save = async () => {
    if (!form.name) return alert("নাম আবশ্যক");
    const payload = { ...form, stock: +form.stock || 0, min_stock: +form.min_stock || 0, unit_price: +form.unit_price || 0 };
    if (editItem) { await supabase.from("materials").update(payload).eq("id", editItem.id); }
    else { await supabase.from("materials").insert([payload]); }
    onRefresh(); setShowModal(false);
  };

  const deleteItem = async (id) => {
    if (!confirm("এই সামগ্রী মুছে ফেলবেন?")) return;
    await supabase.from("materials").delete().eq("id", id); onRefresh();
  };

  const handleExport = () => exportToExcel(data.map(m => ({ নাম: m.name, একক: m.unit, স্টক: m.stock, ন্যূনতম: m.min_stock, একক_মূল্য: m.unit_price, মোট_মূল্য: m.stock * m.unit_price, সাপ্লায়ার: m.supplier })), "Materials", "Materials");

  const handleUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      if (!row["নাম"]) continue;
      await supabase.from("materials").insert([{ name: row["নাম"], unit: row["একক"] || "পিস", stock: +row["স্টক"] || 0, min_stock: +row["ন্যূনতম"] || 0, unit_price: +row["একক_মূল্য"] || 0, supplier: row["সাপ্লায়ার"] || "" }]);
      count++;
    }
    alert(`✅ ${count}টি সামগ্রী আপলোড হয়েছে!`); onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <SectionHeader title="সামগ্রী ও ইনভেন্টরি" action="নতুন সামগ্রী" onAction={openAdd} onExport={handleExport} onPrint={() => printSection("সামগ্রী রিপোর্ট", "materials-content")} onUpload={handleUpload} uploadRef={uploadRef} />
      <Card>
        <div id="materials-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["সামগ্রী", "একক", "স্টক", "ন্যূনতম", "মূল্য/একক", "মোট মূল্য", "সাপ্লায়ার", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>{h}</th>)}</tr></thead>
            <tbody>
              {data.map(m => (
                <tr key={m.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "10px 14px" }}><strong style={{ color: C.primaryDark }}>{m.name}</strong></td>
                  <td style={{ padding: "10px 14px" }}>{m.unit}</td>
                  <td style={{ padding: "10px 14px" }}><span style={{ fontWeight: 600, color: m.stock < m.min_stock ? C.red : C.gray800 }}>{fmtNum(m.stock)}</span></td>
                  <td style={{ padding: "10px 14px" }}>{fmtNum(m.min_stock)}</td>
                  <td style={{ padding: "10px 14px" }}>{fmt(m.unit_price)}</td>
                  <td style={{ padding: "10px 14px" }}><span style={{ fontWeight: 700, color: C.green }}>{fmt(m.stock * m.unit_price)}</span></td>
                  <td style={{ padding: "10px 14px" }}>{m.supplier}</td>
                  <td style={{ padding: "10px 14px" }}>{m.stock < m.min_stock ? <Badge label="কম ⚠️" color="red" /> : <Badge label="পর্যাপ্ত ✅" color="green" />}</td>
                  <td style={{ padding: "10px 14px" }}><div style={{ display: "flex", gap: 6 }}><button onClick={() => openEdit(m)} style={btnEdit}>✏️</button><button onClick={() => deleteItem(m.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "সামগ্রী সম্পাদনা" : "নতুন সামগ্রী"} onClose={() => setShowModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="একক"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{["পিস", "কেজি", "ব্যাগ", "লিটার", "সিএফটি", "ফুট", "মিটার"].map(u => <option key={u}>{u}</option>)}</select></FormField>
            <FormField label="স্টক"><input style={inputStyle} type="number" value={form.stock} onChange={e => setForm({ ...form, stock: e.target.value })} /></FormField>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="ন্যূনতম স্টক"><input style={inputStyle} type="number" value={form.min_stock} onChange={e => setForm({ ...form, min_stock: e.target.value })} /></FormField>
            <FormField label="একক মূল্য (৳)"><input style={inputStyle} type="number" value={form.unit_price} onChange={e => setForm({ ...form, unit_price: e.target.value })} /></FormField>
          </div>
          <FormField label="সাপ্লায়ার"><input style={inputStyle} value={form.supplier} onChange={e => setForm({ ...form, supplier: e.target.value })} /></FormField>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// SITE PROGRESS
// ============================================================
function SiteProgress({ data, projects, onRefresh }) {
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ project: "", date: new Date().toISOString().split("T")[0], work: "", phase: "ফাউন্ডেশন", workers: "", note: "", status: "চলমান" });

  const openAdd = () => { setEditItem(null); setForm({ project: "", date: new Date().toISOString().split("T")[0], work: "", phase: "ফাউন্ডেশন", workers: "", note: "", status: "চলমান" }); setShowModal(true); };
  const openEdit = (item) => { setEditItem(item); setForm({ ...item }); setShowModal(true); };

  const save = async () => {
    if (!form.project || !form.work) return alert("প্রজেক্ট ও কাজের বিবরণ আবশ্যক");
    const payload = { ...form, workers: +form.workers || 0 };
    if (editItem) { await supabase.from("site_progress").update(payload).eq("id", editItem.id); }
    else { await supabase.from("site_progress").insert([payload]); }
    onRefresh(); setShowModal(false);
  };

  const deleteItem = async (id) => {
    if (!confirm("এই আপডেট মুছে ফেলবেন?")) return;
    await supabase.from("site_progress").delete().eq("id", id); onRefresh();
  };

  const handleExport = () => exportToExcel(data.map(s => ({ প্রজেক্ট: s.project, তারিখ: s.date, কাজ: s.work, পর্যায়: s.phase, শ্রমিক: s.workers, নোট: s.note, স্ট্যাটাস: s.status })), "SiteProgress", "Site_Progress");

  return (
    <div>
      <SectionHeader title="সাইট অগ্রগতি" action="নতুন আপডেট" onAction={openAdd} onExport={handleExport} onPrint={() => printSection("সাইট অগ্রগতি রিপোর্ট", "site-content")} />
      <div id="site-content" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 16 }}>
        {[...data].reverse().map(s => (
          <Card key={s.id}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
              <Badge label={s.phase} color="primary" />
              <StatusBadge status={s.status} />
            </div>
            <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14, marginBottom: 6 }}>{s.work}</div>
            <div style={{ fontSize: 12, color: C.primary, fontWeight: 600, marginBottom: 8 }}>📋 {s.project}</div>
            <div style={{ fontSize: 12, color: C.gray600 }}>📅 {s.date} | 👷 {fmtNum(s.workers)} জন</div>
            {s.note && <div style={{ fontSize: 12, color: C.gray400, marginTop: 8, padding: 8, background: C.gray50, borderRadius: 6 }}>📝 {s.note}</div>}
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button onClick={() => openEdit(s)} style={{ ...btnEdit, flex: 1 }}>✏️ Edit</button>
              <button onClick={() => deleteItem(s.id)} style={btnDanger}>🗑️</button>
            </div>
          </Card>
        ))}
      </div>
      {showModal && (
        <Modal title={editItem ? "আপডেট সম্পাদনা" : "নতুন সাইট আপডেট"} onClose={() => setShowModal(false)}>
          <FormField label="প্রজেক্ট *"><select style={inputStyle} value={form.project} onChange={e => setForm({ ...form, project: e.target.value })}><option value="">— নির্বাচন করুন —</option>{projects.map(p => <option key={p.id} value={p.name}>{p.name}</option>)}</select></FormField>
          <FormField label="তারিখ"><input style={inputStyle} type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} /></FormField>
          <FormField label="কাজের বিবরণ *"><input style={inputStyle} value={form.work} onChange={e => setForm({ ...form, work: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="পর্যায়"><select style={inputStyle} value={form.phase} onChange={e => setForm({ ...form, phase: e.target.value })}>{["ডিজাইন", "ফাউন্ডেশন", "স্ট্রাকচার", "ওয়াল", "ছাদ", "ফিনিশিং", "ইন্টেরিয়র", "ইন্সটলেশন"].map(p => <option key={p}>{p}</option>)}</select></FormField>
            <FormField label="স্ট্যাটাস"><select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>{["চলমান", "সম্পন্ন", "পেন্ডিং"].map(s => <option key={s}>{s}</option>)}</select></FormField>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="শ্রমিক সংখ্যা"><input style={inputStyle} type="number" value={form.workers} onChange={e => setForm({ ...form, workers: e.target.value })} /></FormField>
          </div>
          <FormField label="নোট"><input style={inputStyle} value={form.note} onChange={e => setForm({ ...form, note: e.target.value })} /></FormField>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
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

  return (
    <div>
      <h2 style={{ color: C.primaryDark, fontSize: 18, fontWeight: 700, marginBottom: 18 }}>ড্যাশবোর্ড</h2>
      <div id="dashboard-content">
        <div style={{ marginBottom: 20, padding: "18px 22px", background: `linear-gradient(135deg, ${C.primaryDark}, ${C.primary})`, borderRadius: 14, color: C.white }}>
          <div style={{ fontSize: 13, opacity: 0.8 }}>স্বাগতম,</div>
          <div style={{ fontSize: 20, fontWeight: 700, marginTop: 4 }}>Noksha Interior & Construction</div>
          <div style={{ fontSize: 12, opacity: 0.7, marginTop: 4 }}>আজকের তারিখ: {new Date().toLocaleDateString("bn-BD")} | Supabase Connected ✅</div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 20 }}>
          <StatCard icon="🏗️" label="সক্রিয় প্রজেক্ট" value={fmtNum(activeProjects)} sub={`মোট ${projects.length}টি`} color="#E8F5E9" />
          <StatCard icon="👥" label="মোট ক্লায়েন্ট" value={fmtNum(clients.length)} color="#F0FFF4" />
          <StatCard icon="👷" label="কর্মী সংখ্যা" value={fmtNum(employees.length)} color="#FFF8E1" />
          <StatCard icon="💰" label="মোট আয়" value={fmt(totalIncome)} color="#F0FFF4" />
          <StatCard icon="💸" label="মোট ব্যয়" value={fmt(totalExpense)} color="#FFF5F5" />
          <StatCard icon="📦" label="কম স্টক" value={fmtNum(lowStock)} color={lowStock > 0 ? "#FFEBEE" : "#F0FFF4"} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
          <Card>
            <h3 style={{ margin: "0 0 14px", color: C.primaryDark, fontSize: 14, fontWeight: 700 }}>চলমান প্রজেক্ট</h3>
            {projects.filter(p => p.status === "চলমান").map(p => (
              <div key={p.id} style={{ marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: C.primaryDark }}>{p.name}</span>
                  <span style={{ fontSize: 12, color: C.primary, fontWeight: 700 }}>{p.progress}%</span>
                </div>
                <ProgressBar value={p.progress} />
                <div style={{ fontSize: 11, color: C.gray400, marginTop: 3 }}>{p.client}</div>
              </div>
            ))}
            {projects.filter(p => p.status === "চলমান").length === 0 && <div style={{ color: C.gray400, fontSize: 13 }}>কোনো চলমান প্রজেক্ট নেই</div>}
          </Card>
          <Card>
            <h3 style={{ margin: "0 0 14px", color: C.primaryDark, fontSize: 14, fontWeight: 700 }}>সর্বশেষ লেনদেন</h3>
            {[...transactions].slice(-5).reverse().map(t => (
              <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: `1px solid ${C.gray100}` }}>
                <div><div style={{ fontSize: 13, fontWeight: 600 }}>{t.description}</div><div style={{ fontSize: 11, color: C.gray400 }}>{t.date}</div></div>
                <div style={{ fontWeight: 700, color: t.type === "আয়" ? C.green : C.red, fontSize: 13 }}>{t.type === "আয়" ? "+" : "-"}{fmt(t.amount)}</div>
              </div>
            ))}
            {transactions.length === 0 && <div style={{ color: C.gray400, fontSize: 13 }}>কোনো লেনদেন নেই</div>}
          </Card>
        </div>
        {lowStock > 0 && (
          <Card>
            <h3 style={{ margin: "0 0 12px", color: C.primaryDark, fontSize: 14, fontWeight: 700 }}>⚠️ কম স্টক সতর্কতা</h3>
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
// BOQ SYSTEM
// ============================================================
const BOQ_ROOMS = ["Master Bedroom", "Son's Bedroom", "Daughter's Bedroom", "Kitchen", "Drawing Room", "Dining Room", "Bathroom", "Others"];
const BOQ_UNITS = ["sft", "rft", "nos", "set", "ls", "kg", "sqm", "rmt"];
const fmtBOQ = (n) => Number(n || 0).toLocaleString("en-BD", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const genProjId = () => "NIC-BOQ-" + Date.now().toString(36).toUpperCase();

function BOQSystem() {
  const [tab, setTab] = useState("boq");
  const [projects, setProjects] = useState([]);
  const [selProj, setSelProj] = useState(null);
  const [settings, setSettings] = useState(null);
  const [boqItems, setBoqItems] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [stdRates, setStdRates] = useState([]);
  const [loading, setLoading] = useState(false);
  const [showProjModal, setShowProjModal] = useState(false);
  const [showItemModal, setShowItemModal] = useState(false);
  const [editItem, setEditItem] = useState(null);

  useEffect(() => { loadProjects(); loadStdRates(); }, []);
  useEffect(() => { if (selProj) { loadBOQ(selProj); loadExpenses(selProj); loadSettings(selProj); } }, [selProj]);

  const loadProjects = async () => { const { data } = await supabase.from("project_settings").select("*").order("created_at", { ascending: false }); setProjects(data || []); };
  const loadSettings = async (pid) => { const { data } = await supabase.from("project_settings").select("*").eq("project_id", pid).single(); setSettings(data); };
  const loadBOQ = async (pid) => { setLoading(true); const { data } = await supabase.from("project_boq").select("*").eq("project_id", pid).order("room_name").order("item_no"); setBoqItems(data || []); setLoading(false); };
  const loadExpenses = async (pid) => { const { data } = await supabase.from("project_expenses").select("*").eq("project_id", pid).order("expense_date"); setExpenses(data || []); };
  const loadStdRates = async () => { const { data } = await supabase.from("standard_rates").select("*").order("category"); setStdRates(data || []); };

  const roomGroups = boqItems.reduce((acc, item) => { if (!acc[item.room_name]) acc[item.room_name] = []; acc[item.room_name].push(item); return acc; }, {});
  const grandTotal = boqItems.reduce((s, i) => s + Number(i.amount || 0), 0);
  const deliveryCharge = Number(settings?.delivery_charge || 0);
  const subTotal = grandTotal + deliveryCharge;
  const totalExpenses = expenses.reduce((s, e) => s + Number(e.amount || 0), 0);
  const netProfit = subTotal - totalExpenses;

  const saveProject = async (form) => {
    const pid = genProjId();
    await supabase.from("project_settings").insert({ project_id: pid, project_name: form.project_name, client_name: form.client_name, client_address: form.client_address, client_phone: form.client_phone, delivery_charge: Number(form.delivery_charge) || 0, exclusions: form.exclusions ? form.exclusions.split("\n").filter(Boolean) : [], terms_conditions: form.terms_conditions ? form.terms_conditions.split("\n").filter(Boolean) : [] });
    await loadProjects(); setSelProj(pid); setShowProjModal(false);
  };

  const saveBOQItem = async (form) => {
    if (!selProj) return alert("আগে একটি Project select করুন!");
    const qty = Number(form.qty) || 0;
    const rate = Number(form.rate) || 0;
    const amount = qty * rate;
    const payload = {
      project_id: selProj,
      room_name: form.room_name || "Master Bedroom",
      code_no: form.code_no || "",
      item_no: Number(form.item_no) || 1,
      item_name: form.item_name || "",
      work_description: form.work_description || "",
      specification: form.specification || "",
      unit: form.unit || "sft",
      qty,
      rate,
      amount,
      is_rate_fixed: form.is_rate_fixed || false
    };
    let result;
    if (editItem) {
      result = await supabase.from("project_boq").update(payload).eq("id", editItem.id);
    } else {
      result = await supabase.from("project_boq").insert([payload]);
    }
    if (result.error) {
      alert("সংরক্ষণ ব্যর্থ: " + result.error.message);
      return;
    }
    await loadBOQ(selProj);
    setShowItemModal(false);
    setEditItem(null);
  };

  const deleteBOQItem = async (id) => { if (!confirm("Item মুছবেন?")) return; await supabase.from("project_boq").delete().eq("id", id); await loadBOQ(selProj); };

  const saveExpense = async (form) => { if (!selProj) return alert("আগে Project select করুন!"); const qty2 = Number(form.qty) || 1; const rate2 = Number(form.rate) || 0; const r = await supabase.from("project_expenses").insert([{ project_id: selProj, expense_date: form.expense_date || new Date().toISOString().split("T")[0], item_name: form.item_name || "", description: form.description || "", qty: qty2, rate: rate2, amount: qty2 * rate2, category: form.category || "material" }]); if (r.error) { alert("Error: " + r.error.message); return; } await loadExpenses(selProj); };
  const deleteExpense = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("project_expenses").delete().eq("id", id); await loadExpenses(selProj); };

  const updateDelivery = async (val) => { await supabase.from("project_settings").update({ delivery_charge: Number(val) || 0 }).eq("project_id", selProj); await loadSettings(selProj); };

  const handlePrint = () => {
    const content = document.getElementById("boq-print-area");
    if (!content) return;
    const win = window.open("", "_blank");
    win.document.write(`<html><head><title>BOQ — NIC</title><style>body{font-family:Arial,sans-serif;padding:20px;color:#1A1A1A}.header{text-align:center;border-bottom:3px solid #3F5F45;padding-bottom:12px;margin-bottom:16px}.company-name{font-size:22px;font-weight:800;color:#3F5F45}.doc-title{font-size:15px;font-weight:700;background:#E8F0E9;color:#2A3F2E;padding:4px 16px;border-radius:6px;display:inline-block;margin-top:6px}table{width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px}th{background:#3F5F45;color:white;padding:8px 10px;text-align:center}td{padding:7px 10px;border:1px solid #ddd;text-align:center}.room-header{background:#3F5F45;color:white;padding:8px 14px;font-weight:bold;margin-top:14px}.subtotal{background:#E8F0E9;font-weight:bold}.total-box{max-width:400px;margin-left:auto;border:2px solid #3F5F45;padding:14px;border-radius:6px}.exclusions{margin-top:16px;background:#fff8e1;border:1px solid #C9A84C;padding:12px;border-radius:6px}.signatures{display:flex;justify-content:space-between;margin-top:40px;padding-top:14px;border-top:1px solid #ddd}.sig-line{border-top:1px solid #333;padding-top:4px;width:150px;text-align:center;font-size:12px}@media print{body{padding:10px}}</style></head><body>`);
    win.document.write(`<div class="header"><div class="company-name">NOKSHA INTERIOR & CONSTRUCTION</div><div style="font-size:12px;color:#6B8F6B">নীলটুলী, ফরিদপুর | পল্লবী, ঢাকা</div><div class="doc-title">BILL OF QUANTITIES (BOQ)</div><div style="font-size:12px;margin-top:6px">Project: ${settings?.project_name || ""} | Client: ${settings?.client_name || ""} | তারিখ: ${new Date().toLocaleDateString("en-BD")}</div></div>`);
    win.document.write(content.innerHTML);
    win.document.write(`</body></html>`);
    win.document.close();
    setTimeout(() => win.print(), 600);
  };

  const thS = { padding: "8px 10px", border: "1px solid #ddd", background: "#3F5F45", color: "#fff", fontSize: 12 };
  const tdS = { padding: "7px 10px", border: "1px solid #eee", fontSize: 13, textAlign: "center" };

  return (
    <div>
      {/* Header row */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, color: C.primaryDark, fontSize: 18, fontWeight: 700 }}>BOQ সিস্টেম</h2>
        <button onClick={() => setShowProjModal(true)} style={{ background: C.primary, color: "#fff", border: "none", borderRadius: 8, padding: "8px 16px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}>+ নতুন Project</button>
      </div>

      {/* Project selector */}
      <Card style={{ marginBottom: 16, padding: "12px 16px" }}>
        <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
          <label style={{ fontWeight: 700, color: C.primaryDark, fontSize: 13 }}>Project:</label>
          <select value={selProj || ""} onChange={e => setSelProj(e.target.value || null)} style={{ ...inputStyle, maxWidth: 320, padding: "7px 10px" }}>
            <option value="">— Project বেছে নিন —</option>
            {projects.map(p => <option key={p.project_id} value={p.project_id}>{p.project_id} — {p.project_name} ({p.client_name})</option>)}
          </select>
          {settings && <span style={{ fontSize: 12, color: C.gray600 }}>📍 {settings.client_address} | 📞 {settings.client_phone}</span>}
        </div>
      </Card>

      {!selProj ? (
        <Card style={{ textAlign: "center", padding: "60px 20px" }}>
          <div style={{ fontSize: 48, marginBottom: 12 }}>📋</div>
          <div style={{ color: C.primaryDark, fontWeight: 700, fontSize: 16, marginBottom: 8 }}>BOQ Management System</div>
          <div style={{ color: C.gray600, fontSize: 13 }}>{projects.length > 0 ? `${projects.length}টি project আছে। উপর থেকে বেছে নিন।` : "শুরু করতে নতুন project তৈরি করুন।"}</div>
        </Card>
      ) : (
        <>
          {/* Tabs */}
          <div style={{ display: "flex", borderBottom: `2px solid ${C.primary}`, marginBottom: 20 }}>
            {[{ k: "boq", l: "📋 BOQ" }, { k: "expenses", l: "💸 Daily Expenses" }, { k: "compare", l: "📊 Profit Analysis" }].map(t => (
              <button key={t.k} onClick={() => setTab(t.k)} style={{ padding: "10px 20px", border: "none", borderBottom: tab === t.k ? `3px solid #C9A84C` : "3px solid transparent", background: "none", color: tab === t.k ? C.primaryDark : C.gray600, fontWeight: tab === t.k ? 700 : 400, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>{t.l}</button>
            ))}
          </div>

          {/* BOQ TAB */}
          {tab === "boq" && (
            <div>
              <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
                <button onClick={() => { setEditItem(null); setShowItemModal(true); }} style={{ background: C.primary, color: "#fff", border: "none", borderRadius: 8, padding: "8px 16px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}>+ Item যোগ করুন</button>
                <button onClick={handlePrint} style={{ background: C.blue, color: "#fff", border: "none", borderRadius: 8, padding: "8px 16px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}>🖨️ Print / PDF</button>
              </div>

              {loading ? <div style={{ textAlign: "center", padding: 40, color: C.gray400 }}>লোড হচ্ছে...</div> : (
                <div id="boq-print-area">
                  {Object.keys(roomGroups).length === 0 ? (
                    <Card style={{ textAlign: "center", padding: 50 }}><div style={{ fontSize: 36, marginBottom: 10 }}>📋</div><div style={{ color: C.gray400 }}>কোনো item নেই। "+ Item যোগ করুন" ক্লিক করুন।</div></Card>
                  ) : (
                    Object.entries(roomGroups).map(([room, items]) => {
                      const roomTotal = items.reduce((s, i) => s + Number(i.amount || 0), 0);
                      return (
                        <div key={room} style={{ marginBottom: 20 }}>
                          <div className="room-header" style={{ background: C.primary, color: "#fff", padding: "8px 16px", fontWeight: 700, borderRadius: "4px 4px 0 0", fontSize: 14 }}>{room}</div>
                          <div style={{ overflowX: "auto", background: "#fff", border: "1px solid #eee", borderTop: "none" }}>
                            <table style={{ width: "100%", borderCollapse: "collapse" }}>
                              <thead>
                                <tr>
                                  {["Code No", "Item No", "Work Description & Specification", "Unit", "Qty", "Rate (৳)", "Amount (৳)", ""].map((h, i) => (
                                    <th key={i} style={{ ...thS, textAlign: i === 2 ? "left" : "center", minWidth: i === 2 ? 200 : 60 }}>{h}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {items.map((item, idx) => (
                                  <tr key={item.id} style={{ background: idx % 2 === 0 ? "#fff" : C.gray50 }}>
                                    <td style={tdS}>{item.code_no}</td>
                                    <td style={tdS}>{item.item_no}</td>
                                    <td style={{ ...tdS, textAlign: "left" }}>
                                      <div style={{ fontWeight: 600, color: C.primaryDark }}>{item.item_name}</div>
                                      {item.work_description && <div style={{ fontSize: 11, color: C.gray600, marginTop: 2 }}>{item.work_description}</div>}
                                      {item.specification && <div style={{ fontSize: 11, color: C.gray400, fontStyle: "italic" }}>{item.specification}</div>}
                                    </td>
                                    <td style={tdS}>{item.unit}</td>
                                    <td style={tdS}>{fmtBOQ(item.qty)}</td>
                                    <td style={tdS}>{fmtBOQ(item.rate)}{item.is_rate_fixed && <span style={{ color: "#C9A84C", marginLeft: 4, fontSize: 10 }}>🔒</span>}</td>
                                    <td style={{ ...tdS, fontWeight: 700, color: C.primaryDark }}>৳ {fmtBOQ(item.amount)}</td>
                                    <td style={{ ...tdS, width: 70 }} className="no-print">
                                      <button onClick={() => { setEditItem(item); setShowItemModal(true); }} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 15 }}>✏️</button>
                                      <button onClick={() => deleteBOQItem(item.id)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 15 }}>🗑️</button>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                              <tfoot>
                                <tr style={{ background: C.primaryBg, fontWeight: 700 }}>
                                  <td colSpan={6} style={{ ...tdS, textAlign: "right", fontWeight: 700 }}>Sub Total ({room}):</td>
                                  <td style={{ ...tdS, fontWeight: 700, color: C.primaryDark }}>৳ {fmtBOQ(roomTotal)}</td>
                                  <td />
                                </tr>
                              </tfoot>
                            </table>
                          </div>
                        </div>
                      );
                    })
                  )}

                  {Object.keys(roomGroups).length > 0 && (
                    <>
                      {/* Grand Total box */}
                      <Card style={{ maxWidth: 480, marginLeft: "auto", marginTop: 8 }}>
                        <table style={{ width: "100%", borderCollapse: "collapse" }}>
                          <tbody>
                            <tr><td style={{ padding: "7px 0", fontWeight: 700 }}>Grand Total:</td><td style={{ textAlign: "right", fontWeight: 700 }}>৳ {fmtBOQ(grandTotal)}</td></tr>
                            <tr>
                              <td style={{ padding: "7px 0" }}>Delivery Charge:</td>
                              <td style={{ textAlign: "right" }}>
                                <BOQDeliveryEdit value={deliveryCharge} onSave={updateDelivery} />
                              </td>
                            </tr>
                            <tr style={{ borderTop: `2px solid ${C.primary}` }}>
                              <td style={{ padding: "10px 0", fontWeight: 700, color: C.primaryDark, fontSize: 16 }}>Sub Total:</td>
                              <td style={{ textAlign: "right", fontWeight: 700, color: C.primaryDark, fontSize: 16 }}>৳ {fmtBOQ(subTotal)}</td>
                            </tr>
                          </tbody>
                        </table>
                        {/* Payment Terms */}
                        <div style={{ marginTop: 14, paddingTop: 12, borderTop: `1px solid ${C.gray200}` }}>
                          <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 8, fontSize: 13 }}>Payment Terms:</div>
                          {[[settings?.payment_term_1 || 50, "কাজ শুরুর আগে"], [settings?.payment_term_2 || 40, "কাজ চলাকালীন"], [settings?.payment_term_3 || 10, "কাজ সম্পন্নে"]].map(([pct, label], i) => (
                            <div key={i} style={{ fontSize: 13, marginBottom: 4 }}>• {i + 1}ম কিস্তি ({label}): <strong>{pct}%</strong> = ৳ {fmtBOQ(subTotal * pct / 100)}</div>
                          ))}
                        </div>
                      </Card>

                      {/* Exclusions */}
                      {settings?.exclusions?.length > 0 && (
                        <Card style={{ marginTop: 16, background: "#fff8e1", border: `1px solid #C9A84C` }}>
                          <div style={{ fontWeight: 700, color: "#856404", marginBottom: 8 }}>Exclusions (BOQ তে অন্তর্ভুক্ত নয়):</div>
                          {settings.exclusions.map((ex, i) => <div key={i} style={{ fontSize: 13 }}>• {ex}</div>)}
                        </Card>
                      )}

                      {/* Terms */}
                      {settings?.terms_conditions?.length > 0 && (
                        <Card style={{ marginTop: 12 }}>
                          <div style={{ fontWeight: 700, marginBottom: 8, color: C.primaryDark }}>Terms & Conditions:</div>
                          {settings.terms_conditions.map((t, i) => <div key={i} style={{ fontSize: 12, marginBottom: 3 }}>{i + 1}. {t}</div>)}
                        </Card>
                      )}

                      {/* Signatures */}
                      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 40, paddingTop: 14, borderTop: `1px solid ${C.gray200}` }}>
                        <div style={{ textAlign: "center" }}><div style={{ borderTop: `1px solid ${C.gray800}`, paddingTop: 4, width: 150, fontSize: 12 }}>Client Signature</div></div>
                        <div style={{ textAlign: "center" }}><div style={{ borderTop: `1px solid ${C.gray800}`, paddingTop: 4, width: 150, fontSize: 12 }}>Authorized by NIC</div></div>
                      </div>
                    </>
                  )}
                </div>
              )}
            </div>
          )}

          {/* EXPENSES TAB */}
          {tab === "expenses" && <BOQExpenses expenses={expenses} onSave={saveExpense} onDelete={deleteExpense} />}

          {/* COMPARISON TAB */}
          {tab === "compare" && <BOQComparison grandTotal={grandTotal} deliveryCharge={deliveryCharge} subTotal={subTotal} totalExpenses={totalExpenses} netProfit={netProfit} roomGroups={roomGroups} />}
        </>
      )}

      {/* Modals */}
      {showProjModal && <BOQProjectModal onSave={saveProject} onClose={() => setShowProjModal(false)} />}
      {showItemModal && <BOQItemModal item={editItem} stdRates={stdRates} existingRooms={Object.keys(roomGroups)} onSave={saveBOQItem} onClose={() => { setShowItemModal(false); setEditItem(null); }} />}
    </div>
  );
}

function BOQDeliveryEdit({ value, onSave }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(value);
  useEffect(() => { setVal(value); }, [value]);
  if (editing) return (
    <span style={{ display: "flex", gap: 6, justifyContent: "flex-end", alignItems: "center" }}>
      <input type="number" value={val} onChange={e => setVal(e.target.value)} style={{ width: 90, padding: "3px 6px", border: `1px solid ${C.gray200}`, borderRadius: 4 }} />
      <button onClick={() => { onSave(val); setEditing(false); }} style={{ background: C.primary, color: "#fff", border: "none", borderRadius: 4, padding: "3px 8px", cursor: "pointer", fontSize: 12 }}>✓</button>
    </span>
  );
  return <span>৳ {fmtBOQ(value)} <button onClick={() => setEditing(true)} style={{ background: "none", border: "none", cursor: "pointer", color: C.primary, fontSize: 13 }}>✏️</button></span>;
}

function BOQExpenses({ expenses, onSave, onDelete }) {
  const [form, setForm] = useState({ expense_date: new Date().toISOString().slice(0, 10), item_name: "", description: "", qty: 1, rate: "", category: "material" });
  const total = expenses.reduce((s, e) => s + Number(e.amount || 0), 0);
  const cats = ["material", "labor", "transport", "tools", "miscellaneous"];

  return (
    <div>
      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 12 }}>নতুন Expense যোগ করুন</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10 }}>
          <FormField label="তারিখ"><input type="date" value={form.expense_date} onChange={e => setForm({ ...form, expense_date: e.target.value })} style={inputStyle} /></FormField>
          <FormField label="Item নাম *"><input value={form.item_name} onChange={e => setForm({ ...form, item_name: e.target.value })} placeholder="রড, সিমেন্ট..." style={inputStyle} /></FormField>
          <FormField label="বিবরণ"><input value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} style={inputStyle} /></FormField>
          <FormField label="Category"><select value={form.category} onChange={e => setForm({ ...form, category: e.target.value })} style={inputStyle}>{cats.map(c => <option key={c}>{c}</option>)}</select></FormField>
          <FormField label="Qty"><input type="number" value={form.qty} onChange={e => setForm({ ...form, qty: e.target.value })} style={inputStyle} /></FormField>
          <FormField label="Rate (৳) *"><input type="number" value={form.rate} onChange={e => setForm({ ...form, rate: e.target.value })} style={inputStyle} /></FormField>
          <FormField label="Amount"><div style={{ padding: "9px 12px", background: C.primaryBg, borderRadius: 8, fontWeight: 700, color: C.primaryDark }}>৳ {fmtBOQ(Number(form.qty || 0) * Number(form.rate || 0))}</div></FormField>
        </div>
        <button onClick={async () => { if (!form.item_name || !form.rate) return alert("Item ও Rate দিন"); await onSave(form); setForm({ ...form, item_name: "", description: "", qty: 1, rate: "" }); }} style={{ ...btnPrimary, marginTop: 12, width: "auto", padding: "9px 20px" }}>+ যোগ করুন</button>
      </Card>

      <Card>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "Item", "বিবরণ", "Category", "Qty", "Rate", "Amount", ""].map(h => <th key={h} style={{ padding: "10px 12px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>{h}</th>)}</tr></thead>
            <tbody>
              {expenses.length === 0 ? <tr><td colSpan={8} style={{ textAlign: "center", padding: 40, color: C.gray400 }}>কোনো expense নেই</td></tr> : expenses.map((e, i) => (
                <tr key={e.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={ev => ev.currentTarget.style.background = C.primaryBg} onMouseLeave={ev => ev.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "9px 12px" }}>{e.expense_date}</td>
                  <td style={{ padding: "9px 12px", fontWeight: 600, color: C.primaryDark }}>{e.item_name}</td>
                  <td style={{ padding: "9px 12px", color: C.gray600 }}>{e.description}</td>
                  <td style={{ padding: "9px 12px" }}><Badge label={e.category} color="primary" /></td>
                  <td style={{ padding: "9px 12px" }}>{e.qty}</td>
                  <td style={{ padding: "9px 12px" }}>৳ {fmtBOQ(e.rate)}</td>
                  <td style={{ padding: "9px 12px", fontWeight: 700, color: C.primaryDark }}>৳ {fmtBOQ(e.amount)}</td>
                  <td style={{ padding: "9px 12px" }}><button onClick={() => onDelete(e.id)} style={btnDanger}>🗑️</button></td>
                </tr>
              ))}
            </tbody>
            <tfoot><tr style={{ background: C.primaryBg, fontWeight: 700 }}><td colSpan={6} style={{ padding: "10px 12px", textAlign: "right" }}>মোট Expense:</td><td style={{ padding: "10px 12px", color: C.red, fontWeight: 700 }}>৳ {fmtBOQ(total)}</td><td /></tr></tfoot>
          </table>
        </div>
      </Card>
    </div>
  );
}

function BOQComparison({ grandTotal, deliveryCharge, subTotal, totalExpenses, netProfit, roomGroups }) {
  const isProfit = netProfit >= 0;
  const profitPct = subTotal > 0 ? ((netProfit / subTotal) * 100).toFixed(1) : 0;
  const expensePct = subTotal > 0 ? Math.min(100, (totalExpenses / subTotal) * 100) : 0;

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 20 }}>
        <StatCard icon="📋" label="BOQ Contract Value" value={`৳ ${fmtBOQ(subTotal)}`} color={C.primaryBg} />
        <StatCard icon="💸" label="মোট Actual Expense" value={`৳ ${fmtBOQ(totalExpenses)}`} color="#FFEBEE" />
        <StatCard icon={isProfit ? "📈" : "📉"} label={`Net ${isProfit ? "Profit" : "Loss"} (${profitPct}%)`} value={`৳ ${fmtBOQ(Math.abs(netProfit))}`} color={isProfit ? "#E8F5E9" : "#FFEBEE"} />
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 12 }}>Expense vs Revenue</div>
        <div style={{ background: C.gray100, borderRadius: 10, height: 32, overflow: "hidden", position: "relative" }}>
          <div style={{ width: `${expensePct}%`, height: "100%", background: C.red, borderRadius: 10, display: "flex", alignItems: "center", paddingLeft: 10, color: "#fff", fontSize: 12, fontWeight: 700, transition: "width 0.5s", minWidth: expensePct > 0 ? 80 : 0 }}>
            {expensePct > 5 && `Expense: ${expensePct.toFixed(1)}%`}
          </div>
        </div>
        <div style={{ fontSize: 12, color: C.gray600, marginTop: 6 }}>Profit margin: <strong style={{ color: isProfit ? C.green : C.red }}>{profitPct}%</strong></div>
      </Card>

      <Card>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 12 }}>Room-wise BOQ Breakdown</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead><tr style={{ background: C.primaryBg }}>{["Room / Area", "BOQ Amount (৳)", "% of Total"].map(h => <th key={h} style={{ padding: "9px 12px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>{h}</th>)}</tr></thead>
          <tbody>
            {Object.entries(roomGroups).map(([room, items]) => {
              const rt = items.reduce((s, i) => s + Number(i.amount || 0), 0);
              const pct = grandTotal > 0 ? ((rt / grandTotal) * 100).toFixed(1) : 0;
              return (
                <tr key={room} style={{ borderBottom: `1px solid ${C.gray100}` }}>
                  <td style={{ padding: "9px 12px", fontWeight: 600, color: C.primaryDark }}>{room}</td>
                  <td style={{ padding: "9px 12px", fontWeight: 700 }}>৳ {fmtBOQ(rt)}</td>
                  <td style={{ padding: "9px 12px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ flex: 1, background: C.gray100, borderRadius: 4, height: 8 }}><div style={{ width: `${pct}%`, height: "100%", background: C.primary, borderRadius: 4 }} /></div>
                      <span style={{ minWidth: 35, fontSize: 12 }}>{pct}%</span>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

function BOQProjectModal({ onSave, onClose }) {
  const [form, setForm] = useState({ project_name: "", client_name: "", client_address: "", client_phone: "", delivery_charge: 0, exclusions: "Civil works\nElectrical main wiring\nPlumbing main line\nFurniture (existing)\nAC & installation", terms_conditions: "Payment as per schedule\nWork starts after 1st installment\nExtra work charged separately\nMaterials as per specification" });
  return (
    <Modal title="নতুন BOQ Project" onClose={onClose} size={600}>
      <FormField label="Project Name *"><input style={inputStyle} value={form.project_name} onChange={e => setForm({ ...form, project_name: e.target.value })} placeholder="যেমন: Rahman Villa Interior" /></FormField>
      <FormField label="Client Name *"><input style={inputStyle} value={form.client_name} onChange={e => setForm({ ...form, client_name: e.target.value })} /></FormField>
      <FormField label="Project Address"><input style={inputStyle} value={form.client_address} onChange={e => setForm({ ...form, client_address: e.target.value })} /></FormField>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <FormField label="Client Phone"><input style={inputStyle} value={form.client_phone} onChange={e => setForm({ ...form, client_phone: e.target.value })} /></FormField>
        <FormField label="Delivery Charge (৳)"><input style={inputStyle} type="number" value={form.delivery_charge} onChange={e => setForm({ ...form, delivery_charge: e.target.value })} /></FormField>
      </div>
      <FormField label="Exclusions (প্রতি লাইনে একটি)"><textarea style={{ ...inputStyle, height: 80, resize: "vertical" }} value={form.exclusions} onChange={e => setForm({ ...form, exclusions: e.target.value })} /></FormField>
      <FormField label="Terms & Conditions (প্রতি লাইনে একটি)"><textarea style={{ ...inputStyle, height: 80, resize: "vertical" }} value={form.terms_conditions} onChange={e => setForm({ ...form, terms_conditions: e.target.value })} /></FormField>
      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
        <button onClick={onClose} style={{ ...btnPrimary, width: "auto", background: C.gray400, padding: "9px 16px" }}>বাতিল</button>
        <button onClick={() => { if (!form.project_name || !form.client_name) return alert("Project ও Client name দিন"); onSave(form); }} style={{ ...btnPrimary, width: "auto", padding: "9px 20px" }}>✅ তৈরি করুন</button>
      </div>
    </Modal>
  );
}

function BOQItemModal({ item, onSave, onClose, stdRates, existingRooms }) {
  const allRooms = [...new Set([...existingRooms, ...BOQ_ROOMS])];
  const [form, setForm] = useState(item || { room_name: allRooms[0] || "Master Bedroom", code_no: "", item_no: 1, item_name: "", work_description: "", specification: "", unit: "sft", qty: "", rate: "", is_rate_fixed: false });
  const amt = fmtBOQ(Number(form.qty || 0) * Number(form.rate || 0));

  return (
    <Modal title={item ? "Item সম্পাদনা" : "নতুন BOQ Item"} onClose={onClose} size={680}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <div style={{ gridColumn: "1/-1" }}><FormField label="Room / Area *"><select style={inputStyle} value={form.room_name} onChange={e => setForm({ ...form, room_name: e.target.value })}>{allRooms.map(r => <option key={r}>{r}</option>)}</select></FormField></div>
        <FormField label="Code No"><input style={inputStyle} value={form.code_no} onChange={e => setForm({ ...form, code_no: e.target.value })} placeholder="MB-01" /></FormField>
        <FormField label="Item No"><input style={inputStyle} type="number" value={form.item_no} onChange={e => setForm({ ...form, item_no: e.target.value })} /></FormField>
        <div style={{ gridColumn: "1/-1" }}><FormField label="Item Name *"><input style={inputStyle} value={form.item_name} onChange={e => setForm({ ...form, item_name: e.target.value })} placeholder="যেমন: False Ceiling" /></FormField></div>
        <div style={{ gridColumn: "1/-1" }}><FormField label="Work Description"><input style={inputStyle} value={form.work_description} onChange={e => setForm({ ...form, work_description: e.target.value })} /></FormField></div>
        <div style={{ gridColumn: "1/-1" }}><FormField label="Specification"><input style={inputStyle} value={form.specification} onChange={e => setForm({ ...form, specification: e.target.value })} placeholder="Brand, grade, model" /></FormField></div>
        <FormField label="Unit"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{BOQ_UNITS.map(u => <option key={u}>{u}</option>)}</select></FormField>
        <FormField label="Qty *"><input style={inputStyle} type="number" value={form.qty} onChange={e => setForm({ ...form, qty: e.target.value })} /></FormField>
        <FormField label="Rate (৳) *"><input style={inputStyle} type="number" value={form.rate} onChange={e => setForm({ ...form, rate: e.target.value })} /></FormField>
        <div style={{ display: "flex", alignItems: "center", gap: 8, paddingTop: 18 }}><input type="checkbox" id="fixed" checked={form.is_rate_fixed} onChange={e => setForm({ ...form, is_rate_fixed: e.target.checked })} /><label htmlFor="fixed" style={{ cursor: "pointer", fontSize: 13 }}>🔒 Rate Fixed রাখুন</label></div>
        <div style={{ gridColumn: "1/-1", background: C.primaryBg, padding: "10px 14px", borderRadius: 8, textAlign: "center", fontWeight: 700, color: C.primaryDark }}>Amount: ৳ {amt}</div>
        {stdRates.length > 0 && (
          <div style={{ gridColumn: "1/-1" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.gray600, marginBottom: 6 }}>Standard Rate থেকে নিন:</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {stdRates.map(sr => <button key={sr.id} onClick={() => setForm(f => ({ ...f, item_name: sr.item_name, unit: sr.unit, rate: sr.rate, is_rate_fixed: true }))} style={{ background: C.primaryBg, border: `1px solid ${C.primaryLight}`, borderRadius: 4, padding: "4px 10px", cursor: "pointer", fontSize: 11, color: C.primaryDark }}>{sr.item_name} ({sr.rate}/{sr.unit})</button>)}
            </div>
          </div>
        )}
      </div>
      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
        <button onClick={onClose} style={{ ...btnPrimary, width: "auto", background: C.gray400, padding: "9px 16px" }}>বাতিল</button>
        <button onClick={() => { if (!form.item_name || !form.qty || !form.rate) return alert("Item, Qty ও Rate দিন"); onSave(form); }} style={{ ...btnPrimary, width: "auto", padding: "9px 20px" }}>{item ? "✅ Update করুন" : "✅ যোগ করুন"}</button>
      </div>
    </Modal>
  );
}

// ============================================================
// CONSTRUCTION PROJECTS MODULE
// ============================================================

function ConstructionProjects({ currentUser }) {
  const [projects, setProjects] = useState([]);
  const [selProject, setSelProject] = useState(null);
  const [tab, setTab] = useState("daily");
  const [showNewProject, setShowNewProject] = useState(false);
  const isAdmin = currentUser?.role === "admin";

  useEffect(() => { loadProjects(); }, []);

  const loadProjects = async () => {
    const { data } = await supabase.from("construction_projects").select("*").order("created_at", { ascending: false });
    let all = data || [];
    // Site Engineer শুধু assigned projects দেখবে
    if (currentUser?.role === "site_engineer" && currentUser?.assigned_projects?.length > 0) {
      all = all.filter(p => currentUser.assigned_projects.includes(p.id));
    }
    setProjects(all);
  };

  const CP_TABS = [
    { id: "daily", icon: "📅", label: "Daily Updates" },
    { id: "expenses", icon: "💸", label: "Expenses" },
    { id: "stock", icon: "🧱", label: "Stock Register" },
    { id: "payments", icon: "💰", label: "Office থেকে টাকা" },
    { id: "summary", icon: "📊", label: "Project Summary" },
  ];

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, color: C.primaryDark, fontSize: 18, fontWeight: 700 }}>🏚️ Construction Projects</h2>
        {isAdmin && <button onClick={() => setShowNewProject(true)} style={{ background: C.primary, color: C.white, border: "none", borderRadius: 8, padding: "8px 16px", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>+ নতুন প্রজেক্ট</button>}
      </div>
      <Card style={{ marginBottom: 16, padding: "14px 18px" }}>
        <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
          <label style={{ fontWeight: 700, color: C.primaryDark, fontSize: 13, flexShrink: 0 }}>প্রজেক্ট বেছে নিন:</label>
          <select value={selProject?.id || ""} onChange={e => { const p = projects.find(x => x.id === e.target.value); setSelProject(p || null); }} style={{ ...inputStyle, maxWidth: 360, padding: "8px 12px" }}>
            <option value="">— প্রজেক্ট সিলেক্ট করুন —</option>
            {projects.map(p => <option key={p.id} value={p.id}>{p.name} — {p.client_name}</option>)}
          </select>
          {selProject && <span style={{ fontSize: 12, color: C.primaryLight, fontWeight: 600 }}>📍 {selProject.site_address}</span>}
        </div>
      </Card>
      {!selProject ? (
        <Card style={{ textAlign: "center", padding: "50px 20px" }}>
          <div style={{ fontSize: 48, marginBottom: 12 }}>🏗️</div>
          <div style={{ color: C.primaryDark, fontWeight: 700, fontSize: 16, marginBottom: 8 }}>Construction Site Management</div>
          <div style={{ color: C.gray600, fontSize: 13 }}>{projects.length > 0 ? `${projects.length}টি প্রজেক্ট আছে। উপর থেকে বেছে নিন।` : "শুরু করতে নতুন প্রজেক্ট তৈরি করুন।"}</div>
        </Card>
      ) : (
        <>
          <div style={{ background: `linear-gradient(135deg, ${C.primaryDark}, ${C.primary})`, borderRadius: 12, padding: "14px 20px", marginBottom: 16, color: C.white, display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
            <div>
              <div style={{ fontSize: 16, fontWeight: 800 }}>{selProject.name}</div>
              <div style={{ fontSize: 12, opacity: 0.8, marginTop: 3 }}>ক্লায়েন্ট: {selProject.client_name} | সাইট: {selProject.site_address}</div>
            </div>
            <div style={{ display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap" }}>
              <div style={{ textAlign: "center" }}><div style={{ fontSize: 11, opacity: 0.7 }}>শুরু</div><div style={{ fontSize: 13, fontWeight: 700 }}>{selProject.start_date || "—"}</div></div>
              <div style={{ textAlign: "center" }}><div style={{ fontSize: 11, opacity: 0.7 }}>শেষ</div><div style={{ fontSize: 13, fontWeight: 700 }}>{selProject.end_date || "—"}</div></div>
              <div style={{ textAlign: "center" }}><div style={{ fontSize: 11, opacity: 0.7 }}>চুক্তি</div><div style={{ fontSize: 13, fontWeight: 700 }}>{fmt(selProject.deal_amount)}</div></div>
              <span style={{ background: "rgba(255,255,255,0.2)", padding: "4px 12px", borderRadius: 20, fontSize: 12, fontWeight: 600 }}>{selProject.status}</span>
            </div>
          </div>
          <div style={{ display: "flex", gap: 4, marginBottom: 20, borderBottom: `2px solid ${C.gray200}`, overflowX: "auto" }}>
            {CP_TABS.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)} style={{ padding: "10px 16px", border: "none", borderBottom: tab === t.id ? `3px solid ${C.primary}` : "3px solid transparent", background: "none", color: tab === t.id ? C.primaryDark : C.gray600, fontWeight: tab === t.id ? 700 : 400, cursor: "pointer", fontSize: 13, fontFamily: "inherit", whiteSpace: "nowrap" }}>
                {t.icon} {t.label}
              </button>
            ))}
          </div>
          {tab === "daily" && <CPDailyUpdates projectId={selProject.id} />}
          {tab === "expenses" && <CPExpenses projectId={selProject.id} />}
          {tab === "stock" && <CPStock projectId={selProject.id} />}
          {tab === "payments" && <CPPayments projectId={selProject.id} />}
          {tab === "summary" && <CPSummary project={selProject} />}
        </>
      )}
      {showNewProject && <CPNewProjectModal onSave={async (form) => {
        const { error } = await supabase.from("construction_projects").insert([{ ...form, total_budget: +form.total_budget || 0, deal_amount: +form.deal_amount || 0, start_date: form.start_date || null, end_date: form.end_date || null }]);
        if (error) return alert("Error: " + error.message);
        await loadProjects(); setShowNewProject(false);
      }} onClose={() => setShowNewProject(false)} />}
    </div>
  );
}

function CPNewProjectModal({ onSave, onClose }) {
  const [form, setForm] = useState({ name: "", client_name: "", site_address: "", project_type: "আবাসিক নির্মাণ", start_date: "", end_date: "", chief_engineer: "", total_budget: "", deal_amount: "", status: "চলমান" });
  return (
    <Modal title="নতুন Construction প্রজেক্ট" onClose={onClose} size={580}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <div style={{ gridColumn: "1/-1" }}><FormField label="প্রজেক্টের নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="যেমন: রহিম সাহেবের বাড়ি নির্মাণ" /></FormField></div>
        <FormField label="ক্লায়েন্টের নাম *"><input style={inputStyle} value={form.client_name} onChange={e => setForm({ ...form, client_name: e.target.value })} /></FormField>
        <FormField label="প্রজেক্টের ধরন"><select style={inputStyle} value={form.project_type} onChange={e => setForm({ ...form, project_type: e.target.value })}>{["আবাসিক নির্মাণ", "বাণিজ্যিক নির্মাণ", "রিনোভেশন", "এক্সটেনশন", "অন্যান্য"].map(t => <option key={t}>{t}</option>)}</select></FormField>
        <div style={{ gridColumn: "1/-1" }}><FormField label="সাইটের ঠিকানা"><input style={inputStyle} value={form.site_address} onChange={e => setForm({ ...form, site_address: e.target.value })} placeholder="পূর্ণ ঠিকানা" /></FormField></div>
        <FormField label="প্রধান সাইট ইঞ্জিনিয়ার"><input style={inputStyle} value={form.chief_engineer} onChange={e => setForm({ ...form, chief_engineer: e.target.value })} /></FormField>
        <FormField label="স্ট্যাটাস"><select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>{["চলমান", "বিরতি", "সম্পন্ন", "বাতিল"].map(s => <option key={s}>{s}</option>)}</select></FormField>
        <FormField label="শুরুর তারিখ"><input style={inputStyle} type="date" value={form.start_date} onChange={e => setForm({ ...form, start_date: e.target.value })} /></FormField>
        <FormField label="শেষের তারিখ"><input style={inputStyle} type="date" value={form.end_date} onChange={e => setForm({ ...form, end_date: e.target.value })} /></FormField>
        <FormField label="চুক্তি মূল্য (৳)"><input style={inputStyle} type="number" value={form.deal_amount} onChange={e => setForm({ ...form, deal_amount: e.target.value })} placeholder="0" /></FormField>
        <FormField label="মোট বাজেট (৳)"><input style={inputStyle} type="number" value={form.total_budget} onChange={e => setForm({ ...form, total_budget: e.target.value })} placeholder="0" /></FormField>
      </div>
      <button onClick={() => { if (!form.name || !form.client_name) return alert("নাম ও ক্লায়েন্ট আবশ্যক"); onSave({ ...form, start_date: form.start_date || null, end_date: form.end_date || null }); }} style={btnPrimary}>✅ প্রজেক্ট তৈরি করুন</button>
    </Modal>
  );
}

function CPDailyUpdates({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { update_date: new Date().toISOString().split("T")[0], work_done: "", category: "নির্মাণ কাজ", workers_count: "", weather: "স্বাভাবিক", progress_pct: "", issues: "", next_plan: "", reported_by: "" };
  const [form, setForm] = useState(emptyForm);
  const categories = ["নির্মাণ কাজ", "ফাউন্ডেশন", "কলাম/বিম", "ছাদ ঢালাই", "ইটের গাঁথুনি", "প্লাস্টার", "ফ্লোর টাইলস", "ওয়াল টাইলস", "রঙের কাজ", "ইলেকট্রিক", "প্লাম্বিং", "দরজা-জানালা", "পরিষ্কার", "অন্যান্য"];
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("site_daily_updates").select("*").eq("project_id", projectId).order("update_date", { ascending: false }); setItems(data || []); };
  const save = async () => {
    if (!form.work_done) return alert("কাজের বিবরণ আবশ্যক");
    const payload = { ...form, project_id: projectId, workers_count: +form.workers_count || 0, progress_pct: +form.progress_pct || 0 };
    if (editItem) { await supabase.from("site_daily_updates").update(payload).eq("id", editItem.id); } else { await supabase.from("site_daily_updates").insert([payload]); }
    await load(); setShowModal(false); setEditItem(null);
  };
  const del = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("site_daily_updates").delete().eq("id", id); await load(); };
  return (
    <div>
      <SectionHeader title="📅 Daily Work Updates" action="নতুন আপডেট" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => printSection("Daily Work Updates", "daily-print")} onExport={() => exportToExcel(items.map(i => ({ তারিখ: i.update_date, কাজ: i.work_done, ক্যাটাগরি: i.category, শ্রমিক: i.workers_count, আবহাওয়া: i.weather, অগ্রগতি: i.progress_pct + "%", সমস্যা: i.issues, পরবর্তী_পরিকল্পনা: i.next_plan, রিপোর্টকারী: i.reported_by })), "Daily", "Daily_Updates")} />
      <div id="daily-print">
        {items.length === 0 ? <Card style={{ textAlign: "center", padding: 40, color: C.gray400 }}>কোনো আপডেট নেই!</Card> : items.map(item => (
          <Card key={item.id} style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8, flexWrap: "wrap" }}>
                  <span style={{ fontWeight: 700, color: C.primary, fontSize: 14 }}>📅 {item.update_date}</span>
                  <Badge label={item.category} color="primary" />
                  <Badge label={`👷 ${item.workers_count} জন`} color="blue" />
                  <Badge label={`☁️ ${item.weather}`} color="gray" />
                  {item.progress_pct > 0 && <Badge label={`📈 ${item.progress_pct}%`} color="green" />}
                </div>
                <div style={{ fontWeight: 600, color: C.primaryDark, fontSize: 14, marginBottom: 6 }}>{item.work_done}</div>
                {item.issues && <div style={{ fontSize: 12, color: C.red, background: C.redLight, padding: "6px 10px", borderRadius: 6, marginBottom: 6 }}>⚠️ {item.issues}</div>}
                {item.next_plan && <div style={{ fontSize: 12, color: C.gray600, background: C.gray50, padding: "6px 10px", borderRadius: 6 }}>📋 {item.next_plan}</div>}
                {item.reported_by && <div style={{ fontSize: 11, color: C.gray400, marginTop: 4 }}>রিপোর্টকারী: {item.reported_by}</div>}
              </div>
              <div style={{ display: "flex", gap: 6, marginLeft: 12 }}>
                <button onClick={() => { setEditItem(item); setForm({ ...item }); setShowModal(true); }} style={btnEdit}>✏️</button>
                <button onClick={() => del(item.id)} style={btnDanger}>🗑️</button>
              </div>
            </div>
          </Card>
        ))}
      </div>
      {showModal && (
        <Modal title={editItem ? "আপডেট সম্পাদনা" : "নতুন Daily Update"} onClose={() => { setShowModal(false); setEditItem(null); }} size={600}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="তারিখ *"><input style={inputStyle} type="date" value={form.update_date} onChange={e => setForm({ ...form, update_date: e.target.value })} /></FormField>
            <FormField label="ক্যাটাগরি"><select style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{categories.map(c => <option key={c}>{c}</option>)}</select></FormField>
            <div style={{ gridColumn: "1/-1" }}><FormField label="আজকের কাজের বিবরণ *"><textarea style={{ ...inputStyle, height: 80, resize: "vertical" }} value={form.work_done} onChange={e => setForm({ ...form, work_done: e.target.value })} placeholder="আজকে কী কী কাজ হয়েছে..." /></FormField></div>
            <FormField label="শ্রমিক সংখ্যা"><input style={inputStyle} type="number" value={form.workers_count} onChange={e => setForm({ ...form, workers_count: e.target.value })} /></FormField>
            <FormField label="আবহাওয়া"><select style={inputStyle} value={form.weather} onChange={e => setForm({ ...form, weather: e.target.value })}>{["স্বাভাবিক", "রৌদ্রোজ্জ্বল", "মেঘলা", "বৃষ্টি", "ঝড়"].map(w => <option key={w}>{w}</option>)}</select></FormField>
            <FormField label="অগ্রগতি (%)"><input style={inputStyle} type="number" min="0" max="100" value={form.progress_pct} onChange={e => setForm({ ...form, progress_pct: e.target.value })} /></FormField>
            <FormField label="রিপোর্টকারী"><input style={inputStyle} value={form.reported_by} onChange={e => setForm({ ...form, reported_by: e.target.value })} placeholder="ইঞ্জিনিয়ারের নাম" /></FormField>
            <div style={{ gridColumn: "1/-1" }}><FormField label="সমস্যা / ইস্যু"><textarea style={{ ...inputStyle, height: 60, resize: "vertical" }} value={form.issues} onChange={e => setForm({ ...form, issues: e.target.value })} /></FormField></div>
            <div style={{ gridColumn: "1/-1" }}><FormField label="পরবর্তী পরিকল্পনা"><textarea style={{ ...inputStyle, height: 60, resize: "vertical" }} value={form.next_plan} onChange={e => setForm({ ...form, next_plan: e.target.value })} /></FormField></div>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

function CPExpenses({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { expense_date: new Date().toISOString().split("T")[0], category: "নির্মাণ সামগ্রী", item_name: "", description: "", unit: "পিস", quantity: 1, unit_price: "", amount: 0, supplier: "", payment_method: "নগদ", payment_status: "পরিশোধিত", received_by: "", note: "" };
  const [form, setForm] = useState(emptyForm);
  const categories = ["নির্মাণ সামগ্রী", "শ্রমিক মজুরি", "ইলেকট্রিক সামগ্রী", "প্লাম্বিং সামগ্রী", "যন্ত্রপাতি ভাড়া", "পরিবহন", "টাইলস/পাথর", "রঙ সামগ্রী", "দরজা-জানালা", "হার্ডওয়্যার", "অন্যান্য"];
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("site_expenses").select("*").eq("project_id", projectId).order("expense_date", { ascending: false }); setItems(data || []); };
  const save = async () => {
    if (!form.item_name || !form.unit_price) return alert("আইটেম ও মূল্য আবশ্যক");
    const qty = +form.quantity || 1; const price = +form.unit_price || 0;
    const payload = { ...form, project_id: projectId, quantity: qty, unit_price: price, amount: qty * price };
    if (editItem) { await supabase.from("site_expenses").update(payload).eq("id", editItem.id); } else { await supabase.from("site_expenses").insert([payload]); }
    await load(); setShowModal(false); setEditItem(null);
  };
  const del = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("site_expenses").delete().eq("id", id); await load(); };
  const totalExpense = items.reduce((s, i) => s + (i.amount || 0), 0);
  const catGroups = items.reduce((acc, i) => { acc[i.category] = (acc[i.category] || 0) + (i.amount || 0); return acc; }, {});
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 16 }}>
        <StatCard icon="💸" label="মোট খরচ" value={fmt(totalExpense)} color="#FFF5F5" />
        <StatCard icon="✅" label="পরিশোধিত" value={fmt(items.filter(i => i.payment_status === "পরিশোধিত").reduce((s, i) => s + (i.amount || 0), 0))} color="#F0FFF4" />
        <StatCard icon="⏳" label="বকেয়া" value={fmt(items.filter(i => i.payment_status !== "পরিশোধিত").reduce((s, i) => s + (i.amount || 0), 0))} color={C.yellowLight} />
        <StatCard icon="📋" label="মোট এন্ট্রি" value={fmtNum(items.length)} color={C.primaryBg} />
      </div>
      <SectionHeader title="💸 Project Expenses" action="নতুন খরচ" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => printSection("Project Expenses", "expenses-print")} onExport={() => exportToExcel(items.map(i => ({ তারিখ: i.expense_date, ক্যাটাগরি: i.category, আইটেম: i.item_name, বিবরণ: i.description, একক: i.unit, পরিমাণ: i.quantity, একক_মূল্য: i.unit_price, মোট: i.amount, সাপ্লায়ার: i.supplier, পেমেন্ট: i.payment_method, স্ট্যাটাস: i.payment_status })), "Expenses", "Site_Expenses")} />
      {Object.keys(catGroups).length > 0 && (
        <Card style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 13, marginBottom: 8 }}>ক্যাটাগরি অনুযায়ী:</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {Object.entries(catGroups).sort((a, b) => b[1] - a[1]).map(([cat, total]) => (
              <div key={cat} style={{ background: C.primaryBg, border: `1px solid ${C.primaryLight}`, borderRadius: 8, padding: "6px 12px" }}>
                <div style={{ fontSize: 11, color: C.gray600 }}>{cat}</div>
                <div style={{ fontSize: 13, fontWeight: 700, color: C.primaryDark }}>{fmt(total)}</div>
              </div>
            ))}
          </div>
        </Card>
      )}
      <Card>
        <div id="expenses-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "ক্যাটাগরি", "আইটেম", "বিবরণ", "পরিমাণ", "একক মূল্য", "মোট", "সাপ্লায়ার", "পেমেন্ট", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}`, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
            <tbody>
              {items.map(item => (
                <tr key={item.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{item.expense_date}</td>
                  <td style={{ padding: "8px 10px" }}><Badge label={item.category} color="primary" /></td>
                  <td style={{ padding: "8px 10px", fontWeight: 600, color: C.primaryDark }}>{item.item_name}</td>
                  <td style={{ padding: "8px 10px", fontSize: 11, color: C.gray600, maxWidth: 150 }}>{item.description || "—"}</td>
                  <td style={{ padding: "8px 10px" }}>{item.quantity} {item.unit}</td>
                  <td style={{ padding: "8px 10px" }}>{fmt(item.unit_price)}</td>
                  <td style={{ padding: "8px 10px", fontWeight: 700, color: C.red }}>{fmt(item.amount)}</td>
                  <td style={{ padding: "8px 10px", fontSize: 11 }}>{item.supplier || "—"}</td>
                  <td style={{ padding: "8px 10px", fontSize: 11 }}>{item.payment_method}</td>
                  <td style={{ padding: "8px 10px" }}><Badge label={item.payment_status} color={item.payment_status === "পরিশোধিত" ? "green" : "yellow"} /></td>
                  <td style={{ padding: "8px 10px" }}><div style={{ display: "flex", gap: 4 }}><button onClick={() => { setEditItem(item); setForm({ ...item }); setShowModal(true); }} style={btnEdit}>✏️</button><button onClick={() => del(item.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
              {items.length > 0 && <tr style={{ background: C.primaryBg, fontWeight: 700 }}><td colSpan={6} style={{ padding: "10px", textAlign: "right", color: C.primaryDark }}>সর্বমোট:</td><td style={{ padding: "10px", color: C.red, fontSize: 15 }}>{fmt(totalExpense)}</td><td colSpan={4}></td></tr>}
            </tbody>
          </table>
          {items.length === 0 && <div style={{ textAlign: "center", padding: 30, color: C.gray400 }}>কোনো খরচ নেই!</div>}
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "খরচ সম্পাদনা" : "নতুন খরচ"} onClose={() => { setShowModal(false); setEditItem(null); }} size={620}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="তারিখ *"><input style={inputStyle} type="date" value={form.expense_date} onChange={e => setForm({ ...form, expense_date: e.target.value })} /></FormField>
            <FormField label="ক্যাটাগরি *"><select style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{categories.map(c => <option key={c}>{c}</option>)}</select></FormField>
            <div style={{ gridColumn: "1/-1" }}><FormField label="আইটেমের নাম *"><input style={inputStyle} value={form.item_name} onChange={e => setForm({ ...form, item_name: e.target.value })} placeholder="সিমেন্ট, রড, মিস্ত্রি মজুরি..." /></FormField></div>
            <div style={{ gridColumn: "1/-1" }}><FormField label="বিবরণ / স্পেসিফিকেশন"><input style={inputStyle} value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} placeholder="ব্র্যান্ড, গ্রেড, কাজের বিস্তারিত..." /></FormField></div>
            <FormField label="একক"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{["পিস", "ব্যাগ", "কেজি", "টন", "সিএফটি", "বর্গফুট", "রানিংফুট", "লিটার", "গ্যালন", "দিন", "লট", "সেট"].map(u => <option key={u}>{u}</option>)}</select></FormField>
            <FormField label="পরিমাণ *"><input style={inputStyle} type="number" value={form.quantity} onChange={e => { const q = e.target.value; setForm(f => ({ ...f, quantity: q, amount: (q * (+f.unit_price || 0)).toFixed(2) })); }} /></FormField>
            <FormField label="একক মূল্য (৳) *"><input style={inputStyle} type="number" value={form.unit_price} onChange={e => { const p = e.target.value; setForm(f => ({ ...f, unit_price: p, amount: ((+f.quantity || 1) * p).toFixed(2) })); }} /></FormField>
            <div style={{ background: C.primaryBg, borderRadius: 8, padding: "10px", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column" }}>
              <div style={{ fontSize: 11, color: C.gray600 }}>মোট</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: C.primaryDark }}>{fmt(+form.amount || (+form.quantity || 1) * (+form.unit_price || 0))}</div>
            </div>
            <FormField label="সাপ্লায়ার"><input style={inputStyle} value={form.supplier} onChange={e => setForm({ ...form, supplier: e.target.value })} /></FormField>
            <FormField label="পেমেন্ট পদ্ধতি"><select style={inputStyle} value={form.payment_method} onChange={e => setForm({ ...form, payment_method: e.target.value })}>{["নগদ", "বিকাশ", "নগদ (মোবাইল)", "ব্যাংক ট্রান্সফার", "চেক", "বকেয়া"].map(m => <option key={m}>{m}</option>)}</select></FormField>
            <FormField label="পেমেন্ট স্ট্যাটাস"><select style={inputStyle} value={form.payment_status} onChange={e => setForm({ ...form, payment_status: e.target.value })}>{["পরিশোধিত", "বকেয়া", "আংশিক"].map(s => <option key={s}>{s}</option>)}</select></FormField>
            <FormField label="গ্রহণকারী"><input style={inputStyle} value={form.received_by} onChange={e => setForm({ ...form, received_by: e.target.value })} /></FormField>
            <div style={{ gridColumn: "1/-1" }}><FormField label="নোট"><input style={inputStyle} value={form.note} onChange={e => setForm({ ...form, note: e.target.value })} /></FormField></div>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

function CPStock({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { item_name: "", category: "নির্মাণ সামগ্রী", unit: "পিস", opening_stock: 0, received: 0, used: 0, unit_price: "", supplier: "", last_updated: new Date().toISOString().split("T")[0], min_stock: 0, note: "" };
  const [form, setForm] = useState(emptyForm);
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("site_stock").select("*").eq("project_id", projectId).order("category"); setItems(data || []); };
  const save = async () => {
    if (!form.item_name) return alert("আইটেমের নাম আবশ্যক");
    const received = +form.received || 0; const opening = +form.opening_stock || 0; const used = +form.used || 0;
    const closing = opening + received - used; const price = +form.unit_price || 0;
    const payload = { ...form, project_id: projectId, opening_stock: opening, received, used, closing_stock: closing, unit_price: price, total_value: closing * price, min_stock: +form.min_stock || 0 };
    if (editItem) { await supabase.from("site_stock").update(payload).eq("id", editItem.id); } else { await supabase.from("site_stock").insert([payload]); }
    await load(); setShowModal(false); setEditItem(null);
  };
  const del = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("site_stock").delete().eq("id", id); await load(); };
  const totalValue = items.reduce((s, i) => s + (i.total_value || 0), 0);
  const lowStock = items.filter(i => i.closing_stock <= i.min_stock && i.min_stock > 0);
  return (
    <div>
      {lowStock.length > 0 && <Card style={{ background: C.redLight, border: `1px solid #F5C6CB`, marginBottom: 14 }}><div style={{ fontWeight: 700, color: C.red, marginBottom: 8 }}>⚠️ কম স্টক:</div><div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>{lowStock.map(i => <span key={i.id} style={{ background: C.white, padding: "4px 10px", borderRadius: 6, fontSize: 12, color: C.red, fontWeight: 600 }}>{i.item_name}: {i.closing_stock} {i.unit}</span>)}</div></Card>}
      <SectionHeader title="🧱 Stock Register" action="নতুন আইটেম" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => printSection("Stock Register", "stock-print")} onExport={() => exportToExcel(items.map(i => ({ আইটেম: i.item_name, ক্যাটাগরি: i.category, একক: i.unit, প্রারম্ভিক: i.opening_stock, প্রাপ্ত: i.received, ব্যবহৃত: i.used, অবশিষ্ট: i.closing_stock, একক_মূল্য: i.unit_price, মোট_মূল্য: i.total_value, সাপ্লায়ার: i.supplier })), "Stock", "Stock_Register")} />
      <Card>
        <div id="stock-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["আইটেম", "ক্যাটাগরি", "একক", "প্রারম্ভিক", "প্রাপ্ত", "ব্যবহৃত", "অবশিষ্ট", "একক মূল্য", "মোট মূল্য", "সাপ্লায়ার", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}`, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
            <tbody>
              {items.map(item => (
                <tr key={item.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "8px 10px", fontWeight: 600, color: C.primaryDark }}>{item.item_name}</td>
                  <td style={{ padding: "8px 10px" }}><Badge label={item.category} color="primary" /></td>
                  <td style={{ padding: "8px 10px" }}>{item.unit}</td>
                  <td style={{ padding: "8px 10px", textAlign: "center" }}>{item.opening_stock}</td>
                  <td style={{ padding: "8px 10px", textAlign: "center", color: C.green, fontWeight: 600 }}>+{item.received}</td>
                  <td style={{ padding: "8px 10px", textAlign: "center", color: C.red, fontWeight: 600 }}>-{item.used}</td>
                  <td style={{ padding: "8px 10px", textAlign: "center", fontWeight: 700, color: item.closing_stock <= item.min_stock && item.min_stock > 0 ? C.red : C.primaryDark }}>{item.closing_stock}</td>
                  <td style={{ padding: "8px 10px" }}>{fmt(item.unit_price)}</td>
                  <td style={{ padding: "8px 10px", fontWeight: 700, color: C.primary }}>{fmt(item.total_value)}</td>
                  <td style={{ padding: "8px 10px", fontSize: 11 }}>{item.supplier || "—"}</td>
                  <td style={{ padding: "8px 10px" }}>{item.closing_stock <= item.min_stock && item.min_stock > 0 ? <Badge label="কম ⚠️" color="red" /> : <Badge label="পর্যাপ্ত ✅" color="green" />}</td>
                  <td style={{ padding: "8px 10px" }}><div style={{ display: "flex", gap: 4 }}><button onClick={() => { setEditItem(item); setForm({ ...item }); setShowModal(true); }} style={btnEdit}>✏️</button><button onClick={() => del(item.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
              {items.length > 0 && <tr style={{ background: C.primaryBg, fontWeight: 700 }}><td colSpan={8} style={{ padding: "10px", textAlign: "right", color: C.primaryDark }}>মোট স্টক মূল্য:</td><td style={{ padding: "10px", color: C.primary, fontSize: 15 }}>{fmt(totalValue)}</td><td colSpan={3}></td></tr>}
            </tbody>
          </table>
          {items.length === 0 && <div style={{ textAlign: "center", padding: 30, color: C.gray400 }}>কোনো স্টক নেই!</div>}
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "স্টক সম্পাদনা" : "নতুন স্টক আইটেম"} onClose={() => { setShowModal(false); setEditItem(null); }} size={560}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <div style={{ gridColumn: "1/-1" }}><FormField label="আইটেমের নাম *"><input style={inputStyle} value={form.item_name} onChange={e => setForm({ ...form, item_name: e.target.value })} placeholder="সিমেন্ট, রড, বালি..." /></FormField></div>
            <FormField label="ক্যাটাগরি"><select style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{["নির্মাণ সামগ্রী", "ইলেকট্রিক", "প্লাম্বিং", "ফিনিশিং", "যন্ত্রপাতি", "অন্যান্য"].map(c => <option key={c}>{c}</option>)}</select></FormField>
            <FormField label="একক"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{["পিস", "ব্যাগ", "কেজি", "টন", "সিএফটি", "বর্গফুট", "রানিংফুট", "লিটার", "গ্যালন", "লট"].map(u => <option key={u}>{u}</option>)}</select></FormField>
            <FormField label="প্রারম্ভিক স্টক"><input style={inputStyle} type="number" value={form.opening_stock} onChange={e => setForm({ ...form, opening_stock: e.target.value })} /></FormField>
            <FormField label="নতুন প্রাপ্তি"><input style={inputStyle} type="number" value={form.received} onChange={e => setForm({ ...form, received: e.target.value })} /></FormField>
            <FormField label="ব্যবহৃত"><input style={inputStyle} type="number" value={form.used} onChange={e => setForm({ ...form, used: e.target.value })} /></FormField>
            <div style={{ background: C.primaryBg, borderRadius: 8, padding: "10px", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column" }}>
              <div style={{ fontSize: 11, color: C.gray600 }}>অবশিষ্ট</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: C.primaryDark }}>{(+form.opening_stock || 0) + (+form.received || 0) - (+form.used || 0)} {form.unit}</div>
            </div>
            <FormField label="একক মূল্য (৳)"><input style={inputStyle} type="number" value={form.unit_price} onChange={e => setForm({ ...form, unit_price: e.target.value })} /></FormField>
            <FormField label="সর্বনিম্ন স্টক (সতর্কতা)"><input style={inputStyle} type="number" value={form.min_stock} onChange={e => setForm({ ...form, min_stock: e.target.value })} /></FormField>
            <FormField label="সাপ্লায়ার"><input style={inputStyle} value={form.supplier} onChange={e => setForm({ ...form, supplier: e.target.value })} /></FormField>
            <FormField label="তারিখ"><input style={inputStyle} type="date" value={form.last_updated} onChange={e => setForm({ ...form, last_updated: e.target.value })} /></FormField>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

function CPPayments({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { receive_date: new Date().toISOString().split("T")[0], amount: "", payment_type: "১ম কিস্তি", payment_method: "নগদ", received_by: "", sender_name: "", bank_ref: "", bkash_ref: "", note: "" };
  const [form, setForm] = useState(emptyForm);
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("site_received_payments").select("*").eq("project_id", projectId).order("receive_date", { ascending: false }); setItems(data || []); };
  const save = async () => {
    if (!form.amount) return alert("পরিমাণ আবশ্যক");
    const payload = { ...form, project_id: projectId, amount: +form.amount };
    if (editItem) { await supabase.from("site_received_payments").update(payload).eq("id", editItem.id); } else { await supabase.from("site_received_payments").insert([payload]); }
    await load(); setShowModal(false); setEditItem(null);
  };
  const del = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("site_received_payments").delete().eq("id", id); await load(); };
  const totalReceived = items.reduce((s, i) => s + (i.amount || 0), 0);
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 16 }}>
        <StatCard icon="💰" label="মোট প্রাপ্ত" value={fmt(totalReceived)} color="#F0FFF4" />
        <StatCard icon="📋" label="মোট কিস্তি" value={fmtNum(items.length)} color={C.primaryBg} />
      </div>
      <SectionHeader title="💰 Office থেকে প্রাপ্ত টাকা" action="নতুন প্রাপ্তি" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => printSection("Office Payment Register", "payments-print")} onExport={() => exportToExcel(items.map(i => ({ তারিখ: i.receive_date, পরিমাণ: i.amount, কিস্তি: i.payment_type, পদ্ধতি: i.payment_method, প্রেরক: i.sender_name, গ্রহণকারী: i.received_by, রেফারেন্স: i.bank_ref || i.bkash_ref || "", নোট: i.note })), "Payments", "Office_Payments")} />
      <Card>
        <div id="payments-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "কিস্তির ধরন", "পরিমাণ", "পেমেন্ট পদ্ধতি", "প্রেরক", "গ্রহণকারী", "রেফারেন্স", "নোট", "Action"].map(h => <th key={h} style={{ padding: "9px 12px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}`, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
            <tbody>
              {items.map(item => (
                <tr key={item.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "9px 12px", whiteSpace: "nowrap" }}>{item.receive_date}</td>
                  <td style={{ padding: "9px 12px" }}><Badge label={item.payment_type} color="primary" /></td>
                  <td style={{ padding: "9px 12px", fontWeight: 700, color: C.green, fontSize: 14 }}>+{fmt(item.amount)}</td>
                  <td style={{ padding: "9px 12px" }}>{item.payment_method}</td>
                  <td style={{ padding: "9px 12px" }}>{item.sender_name || "—"}</td>
                  <td style={{ padding: "9px 12px" }}>{item.received_by || "—"}</td>
                  <td style={{ padding: "9px 12px", fontSize: 11, color: C.gray400 }}>{item.bank_ref || item.bkash_ref || "—"}</td>
                  <td style={{ padding: "9px 12px", fontSize: 11 }}>{item.note || "—"}</td>
                  <td style={{ padding: "9px 12px" }}><div style={{ display: "flex", gap: 4 }}><button onClick={() => { setEditItem(item); setForm({ ...item, amount: item.amount || "" }); setShowModal(true); }} style={btnEdit}>✏️</button><button onClick={() => del(item.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
              {items.length > 0 && <tr style={{ background: C.primaryBg, fontWeight: 700 }}><td colSpan={2} style={{ padding: "10px", textAlign: "right", color: C.primaryDark }}>সর্বমোট:</td><td style={{ padding: "10px", color: C.green, fontSize: 15 }}>{fmt(totalReceived)}</td><td colSpan={6}></td></tr>}
            </tbody>
          </table>
          {items.length === 0 && <div style={{ textAlign: "center", padding: 30, color: C.gray400 }}>কোনো প্রাপ্তি নেই!</div>}
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "প্রাপ্তি সম্পাদনা" : "নতুন টাকা প্রাপ্তি"} onClose={() => { setShowModal(false); setEditItem(null); }}>
          <FormField label="তারিখ *"><input style={inputStyle} type="date" value={form.receive_date} onChange={e => setForm({ ...form, receive_date: e.target.value })} /></FormField>
          <FormField label="কিস্তির ধরন"><select style={inputStyle} value={form.payment_type} onChange={e => setForm({ ...form, payment_type: e.target.value })}>{["১ম কিস্তি", "২য় কিস্তি", "৩য় কিস্তি", "৪র্থ কিস্তি", "চূড়ান্ত কিস্তি", "জরুরি অগ্রিম", "বোনাস", "অন্যান্য"].map(t => <option key={t}>{t}</option>)}</select></FormField>
          <FormField label="পরিমাণ (৳) *"><input style={inputStyle} type="number" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} placeholder="০" /></FormField>
          <FormField label="পেমেন্ট পদ্ধতি"><select style={inputStyle} value={form.payment_method} onChange={e => setForm({ ...form, payment_method: e.target.value })}>{["নগদ", "বিকাশ", "নগদ (মোবাইল)", "ব্যাংক ট্রান্সফার", "চেক"].map(m => <option key={m}>{m}</option>)}</select></FormField>
          <FormField label="প্রেরকের নাম"><input style={inputStyle} value={form.sender_name} onChange={e => setForm({ ...form, sender_name: e.target.value })} placeholder="কে পাঠিয়েছেন" /></FormField>
          <FormField label="গ্রহণকারীর নাম"><input style={inputStyle} value={form.received_by} onChange={e => setForm({ ...form, received_by: e.target.value })} placeholder="কে গ্রহণ করেছেন" /></FormField>
          <FormField label="বিকাশ/ব্যাংক রেফারেন্স"><input style={inputStyle} value={form.bkash_ref} onChange={e => setForm({ ...form, bkash_ref: e.target.value })} placeholder="Transaction ID" /></FormField>
          <FormField label="নোট"><input style={inputStyle} value={form.note} onChange={e => setForm({ ...form, note: e.target.value })} /></FormField>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

function CPSummary({ project }) {
  const [expenses, setExpenses] = useState([]);
  const [payments, setPayments] = useState([]);
  const [stock, setStock] = useState([]);
  const [dailyUpdates, setDailyUpdates] = useState([]);
  useEffect(() => {
    const load = async () => {
      const [e, p, s, d] = await Promise.all([
        supabase.from("site_expenses").select("*").eq("project_id", project.id),
        supabase.from("site_received_payments").select("*").eq("project_id", project.id),
        supabase.from("site_stock").select("*").eq("project_id", project.id),
        supabase.from("site_daily_updates").select("*").eq("project_id", project.id),
      ]);
      setExpenses(e.data || []); setPayments(p.data || []); setStock(s.data || []); setDailyUpdates(d.data || []);
    };
    load();
  }, [project.id]);
  const totalExpense = expenses.reduce((s, i) => s + (i.amount || 0), 0);
  const totalReceived = payments.reduce((s, i) => s + (i.amount || 0), 0);
  const totalStockValue = stock.reduce((s, i) => s + (i.total_value || 0), 0);
  const balance = totalReceived - totalExpense;
  const dealAmount = project.deal_amount || 0;
  const projectedProfit = dealAmount - totalExpense;
  const catGroups = expenses.reduce((acc, i) => { acc[i.category] = (acc[i.category] || 0) + (i.amount || 0); return acc; }, {});
  return (
    <div>
      <SectionHeader title="📊 Project Summary" onPrint={() => printSection("Project Summary", "summary-print")} />
      <div id="summary-print">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 20 }}>
          <StatCard icon="💰" label="চুক্তি মূল্য" value={fmt(dealAmount)} color={C.primaryBg} />
          <StatCard icon="📥" label="অফিস থেকে প্রাপ্ত" value={fmt(totalReceived)} color="#F0FFF4" />
          <StatCard icon="💸" label="মোট খরচ" value={fmt(totalExpense)} color="#FFF5F5" />
          <StatCard icon="🏦" label="হাতে ব্যালেন্স" value={fmt(balance)} color={balance >= 0 ? "#F0FFF4" : "#FFF5F5"} />
          <StatCard icon="📦" label="স্টক মূল্য" value={fmt(totalStockValue)} color="#FFF8E1" />
          <StatCard icon="📈" label="প্রজেক্টেড লাভ" value={fmt(projectedProfit)} color={projectedProfit >= 0 ? "#F0FFF4" : "#FFF5F5"} />
        </div>
        <Card style={{ marginBottom: 16 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14, marginBottom: 12 }}>বাজেট ব্যবহার</div>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
            <span style={{ fontSize: 12 }}>মোট খরচ: {fmt(totalExpense)}</span>
            <span style={{ fontSize: 12, fontWeight: 700, color: dealAmount > 0 && totalExpense > dealAmount ? C.red : C.primary }}>{dealAmount > 0 ? Math.round((totalExpense / dealAmount) * 100) : 0}%</span>
          </div>
          <ProgressBar value={dealAmount > 0 ? Math.min(Math.round((totalExpense / dealAmount) * 100), 100) : 0} color={totalExpense > dealAmount ? C.red : C.primary} />
          <div style={{ fontSize: 11, color: C.gray400, marginTop: 6, textAlign: "right" }}>চুক্তি মূল্য: {fmt(dealAmount)}</div>
        </Card>
        <Card style={{ marginBottom: 16 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14, marginBottom: 12 }}>ক্যাটাগরি অনুযায়ী খরচ</div>
          {Object.entries(catGroups).sort((a, b) => b[1] - a[1]).map(([cat, total]) => (
            <div key={cat} style={{ marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                <span style={{ fontSize: 13 }}>{cat}</span>
                <span style={{ fontSize: 13, fontWeight: 700, color: C.primaryDark }}>{fmt(total)}</span>
              </div>
              <ProgressBar value={totalExpense > 0 ? Math.round((total / totalExpense) * 100) : 0} />
            </div>
          ))}
          {Object.keys(catGroups).length === 0 && <div style={{ color: C.gray400, fontSize: 13 }}>কোনো খরচ নেই</div>}
        </Card>
        <Card>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14, marginBottom: 12 }}>প্রজেক্টের তথ্য</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, fontSize: 13 }}>
            {[["প্রজেক্টের নাম", project.name], ["ক্লায়েন্ট", project.client_name], ["সাইট ঠিকানা", project.site_address], ["ধরন", project.project_type], ["প্রধান ইঞ্জিনিয়ার", project.chief_engineer], ["শুরুর তারিখ", project.start_date], ["শেষের তারিখ", project.end_date], ["স্ট্যাটাস", project.status], ["Daily Updates", dailyUpdates.length + "টি"], ["Stock আইটেম", stock.length + "টি"]].map(([label, val]) => (
              <div key={label}><div style={{ fontSize: 11, color: C.gray400 }}>{label}</div><div style={{ fontWeight: 600, color: C.primaryDark }}>{val || "—"}</div></div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

// ============================================================
// MAIN APP
// ============================================================
const ALL_MENU = [
  { id: "dashboard", icon: "🏠", label: "ড্যাশবোর্ড", roles: ["admin"] },
  { id: "projects", icon: "🏗️", label: "প্রজেক্ট", roles: ["admin"] },
  { id: "construction", icon: "🏚️", label: "Construction Projects", roles: ["admin", "site_engineer"] },
  { id: "boq", icon: "📋", label: "BOQ সিস্টেম", roles: ["admin"] },
  { id: "clients", icon: "👥", label: "ক্লায়েন্ট", roles: ["admin"] },
  { id: "employees", icon: "👷", label: "কর্মী (HR)", roles: ["admin"] },
  { id: "attendance", icon: "📋", label: "উপস্থিতি", roles: ["admin"] },
  { id: "finance", icon: "💰", label: "আর্থিক", roles: ["admin"] },
  { id: "site", icon: "📍", label: "সাইট প্রগ্রেস", roles: ["admin"] },
  { id: "materials", icon: "📦", label: "সামগ্রী", roles: ["admin"] },
  { id: "analytics", icon: "📊", label: "রিপোর্ট & Analytics", roles: ["admin"] },
  { id: "users", icon: "👤", label: "User Management", roles: ["admin"] },
  { id: "password", icon: "🔑", label: "পাসওয়ার্ড", roles: ["admin", "site_engineer"] },
];

// ============================================================
// USER MANAGEMENT (Admin only)
// ============================================================
function UserManagement() {
  const [users, setUsers] = useState([]);
  const [projects, setProjects] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ name: "", email: "", password_hash: "", role: "site_engineer", assigned_projects: [], is_active: true });

  useEffect(() => { loadUsers(); loadProjects(); }, []);
  const loadUsers = async () => { const { data } = await supabase.from("app_users").select("*").order("created_at"); setUsers(data || []); };
  const loadProjects = async () => { const { data } = await supabase.from("construction_projects").select("id, name, client_name").order("created_at"); setProjects(data || []); };

  const save = async () => {
    if (!form.name || !form.email || !form.password_hash) return alert("নাম, ইমেইল ও পাসওয়ার্ড আবশ্যক");
    if (editItem) {
      await supabase.from("app_users").update({ ...form }).eq("id", editItem.id);
    } else {
      const { error } = await supabase.from("app_users").insert([{ ...form }]);
      if (error) return alert("Error: " + error.message);
    }
    await loadUsers(); setShowModal(false); setEditItem(null);
  };

  const toggleActive = async (user) => {
    await supabase.from("app_users").update({ is_active: !user.is_active }).eq("id", user.id);
    await loadUsers();
  };

  const del = async (id) => {
    if (!confirm("এই user মুছবেন?")) return;
    await supabase.from("app_users").delete().eq("id", id);
    await loadUsers();
  };

  const toggleProject = (projId) => {
    setForm(f => {
      const arr = f.assigned_projects || [];
      return { ...f, assigned_projects: arr.includes(projId) ? arr.filter(x => x !== projId) : [...arr, projId] };
    });
  };

  return (
    <div>
      <SectionHeader title="👤 User Management" action="নতুন User" onAction={() => { setEditItem(null); setForm({ name: "", email: "", password_hash: "", role: "site_engineer", assigned_projects: [], is_active: true }); setShowModal(true); }} />
      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead><tr style={{ background: C.primaryBg }}>
            {["নাম", "ইমেইল", "Role", "Assigned Projects", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: `2px solid ${C.primary}` }}>{h}</th>)}
          </tr></thead>
          <tbody>
            {users.map(u => (
              <tr key={u.id} style={{ borderBottom: `1px solid ${C.gray100}` }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{u.name}</td>
                <td style={{ padding: "10px 14px" }}>{u.email}</td>
                <td style={{ padding: "10px 14px" }}><Badge label={u.role === "admin" ? "🔑 Admin" : "👷 Site Engineer"} color={u.role === "admin" ? "green" : "blue"} /></td>
                <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600 }}>
                  {u.role === "admin" ? "সব প্রজেক্ট" : (u.assigned_projects?.length > 0 ? projects.filter(p => u.assigned_projects.includes(p.id)).map(p => p.name).join(", ") : "কোনো প্রজেক্ট নেই")}
                </td>
                <td style={{ padding: "10px 14px" }}><Badge label={u.is_active ? "✅ সক্রিয়" : "❌ নিষ্ক্রিয়"} color={u.is_active ? "green" : "red"} /></td>
                <td style={{ padding: "10px 14px" }}>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button onClick={() => { setEditItem(u); setForm({ ...u, password_hash: "" }); setShowModal(true); }} style={btnEdit}>✏️</button>
                    <button onClick={() => toggleActive(u)} style={{ ...btnEdit, background: u.is_active ? C.yellowLight : C.greenLight }}>{u.is_active ? "⏸️" : "▶️"}</button>
                    {u.role !== "admin" && <button onClick={() => del(u.id)} style={btnDanger}>🗑️</button>}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>

      {showModal && (
        <Modal title={editItem ? "User সম্পাদনা" : "নতুন User তৈরি করুন"} onClose={() => { setShowModal(false); setEditItem(null); }} size={520}>
          <FormField label="পূর্ণ নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="ইঞ্জিনিয়ারের নাম" /></FormField>
          <FormField label="ইমেইল *"><input style={inputStyle} type="email" value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} placeholder="engineer@noksha.com" /></FormField>
          <FormField label={editItem ? "নতুন পাসওয়ার্ড (পরিবর্তন করতে চাইলে)" : "পাসওয়ার্ড *"}><input style={inputStyle} type="text" value={form.password_hash} onChange={e => setForm({ ...form, password_hash: e.target.value })} placeholder="password123" /></FormField>
          <FormField label="Role"><select style={inputStyle} value={form.role} onChange={e => setForm({ ...form, role: e.target.value })}><option value="site_engineer">👷 Site Engineer</option><option value="admin">🔑 Admin</option></select></FormField>
          {form.role === "site_engineer" && projects.length > 0 && (
            <FormField label="Assigned Projects (যে প্রজেক্টে access পাবে)">
              <div style={{ border: `1px solid ${C.gray200}`, borderRadius: 8, padding: 12, maxHeight: 200, overflowY: "auto" }}>
                {projects.map(p => (
                  <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, cursor: "pointer" }} onClick={() => toggleProject(p.id)}>
                    <input type="checkbox" checked={(form.assigned_projects || []).includes(p.id)} onChange={() => toggleProject(p.id)} style={{ accentColor: C.primary, width: 16, height: 16 }} />
                    <span style={{ fontSize: 13 }}>{p.name} — {p.client_name}</span>
                  </div>
                ))}
              </div>
            </FormField>
          )}
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
            <input type="checkbox" checked={form.is_active} onChange={e => setForm({ ...form, is_active: e.target.checked })} style={{ accentColor: C.primary, width: 16, height: 16 }} />
            <label style={{ fontSize: 13, fontWeight: 600, color: C.gray800 }}>সক্রিয় করুন</label>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

export default function App() {
  const [loggedIn, setLoggedIn] = useState(localStorage.getItem("nic_logged_in") === "true");
  const [currentUser, setCurrentUser] = useState(() => {
    try { return JSON.parse(localStorage.getItem("nic_user") || "null"); } catch { return null; }
  });
  const [active, setActive] = useState("dashboard");
  const [sideOpen, setSideOpen] = useState(true);
  const [data, setData] = useState({ projects: [], clients: [], employees: [], transactions: [], siteProgress: [], materials: [] });
  const [loading, setLoading] = useState(true);

  const isAdmin = currentUser?.role === "admin";
  const isSiteEngineer = currentUser?.role === "site_engineer";

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

  const logout = () => { localStorage.removeItem("nic_logged_in"); localStorage.removeItem("nic_user"); setLoggedIn(false); setCurrentUser(null); };

  if (!loggedIn) return <LoginPage onLogin={(user) => { setCurrentUser(user); setLoggedIn(true); if (user.role === "site_engineer") setActive("construction"); }} />;

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.gray50, fontFamily: "'Hind Siliguri', Arial, sans-serif" }}>
      {/* SIDEBAR */}
      <div style={{ width: sideOpen ? 230 : 60, background: C.primaryDark, display: "flex", flexDirection: "column", transition: "width 0.3s", flexShrink: 0, position: "sticky", top: 0, height: "100vh", overflow: "hidden" }}>
        <div style={{ padding: sideOpen ? "18px 18px 14px" : "18px 10px 14px", borderBottom: `1px solid rgba(255,255,255,0.1)` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 36, height: 36, background: C.primary, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>🏠</div>
            {sideOpen && <div><div style={{ color: C.white, fontWeight: 900, fontSize: 15, letterSpacing: 1 }}>NOKSHA</div><div style={{ color: "rgba(255,255,255,0.5)", fontSize: 9, letterSpacing: 2 }}>INTERIOR & CONSTRUCTION</div></div>}
          </div>
        </div>
        <nav style={{ flex: 1, padding: "8px 0", overflowY: "auto" }}>
          {ALL_MENU.filter(m => m.roles.includes(currentUser?.role || "admin")).map(m => (
            <button key={m.id} onClick={() => setActive(m.id)} style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", padding: sideOpen ? "10px 18px" : "10px 0", justifyContent: sideOpen ? "flex-start" : "center", background: active === m.id ? `rgba(255,255,255,0.12)` : "none", border: "none", borderLeft: active === m.id ? `3px solid ${C.primaryLight}` : "3px solid transparent", color: active === m.id ? C.white : "rgba(255,255,255,0.6)", cursor: "pointer", fontSize: 12, fontWeight: active === m.id ? 700 : 400, fontFamily: "inherit", transition: "all 0.15s" }}>
              <span style={{ fontSize: 16, flexShrink: 0 }}>{m.icon}</span>
              {sideOpen && <span style={{ whiteSpace: "nowrap", fontSize: 13 }}>{m.label}</span>}
            </button>
          ))}
        </nav>
        <button onClick={logout} style={{ background: "rgba(220,53,69,0.15)", border: "none", color: "#ff8080", padding: "10px", cursor: "pointer", fontSize: sideOpen ? 12 : 16, fontFamily: "inherit" }}>
          {sideOpen ? "🚪 লগআউট" : "🚪"}
        </button>
        <button onClick={() => setSideOpen(o => !o)} style={{ background: "rgba(255,255,255,0.08)", border: "none", color: C.white, padding: "10px", cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>
          {sideOpen ? "◀" : "▶"}
        </button>
      </div>

      {/* MAIN */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <div style={{ background: C.white, borderBottom: `1px solid ${C.gray200}`, padding: "12px 24px", display: "flex", justifyContent: "space-between", alignItems: "center", position: "sticky", top: 0, zIndex: 100, borderTop: `3px solid ${C.primary}` }}>
          <div>
            <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 16 }}>{ALL_MENU.find(m => m.id === active)?.label}</div>
            <div style={{ fontSize: 11, color: C.gray400 }}>Noksha Interior & Construction · ফরিদপুর & ঢাকা</div>
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            {isAdmin && <button onClick={loadAll} style={{ background: C.primaryBg, border: `1px solid ${C.primaryLight}`, borderRadius: 8, padding: "6px 12px", cursor: "pointer", fontSize: 12, color: C.primaryDark, fontFamily: "inherit" }}>🔄 রিফ্রেশ</button>}
            <div style={{ width: 36, height: 36, background: C.primary, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", color: C.white, fontWeight: 700, fontSize: 14 }}>{currentUser?.name?.[0] || "R"}</div>
            <div style={{ fontSize: 12, color: C.gray600 }}>
              <div style={{ fontWeight: 600, color: C.primaryDark }}>{currentUser?.name || "মোঃ রানা"}</div>
              <div>{currentUser?.role === "admin" ? "Admin" : "Site Engineer"}</div>
            </div>
          </div>
        </div>

        <div style={{ flex: 1, padding: "20px 24px", overflowY: "auto" }}>
          {loading && isAdmin ? (
            <div style={{ textAlign: "center", padding: 60, color: C.gray400 }}>
              <div style={{ fontSize: 40, marginBottom: 16 }}>⏳</div>
              <div style={{ color: C.primary, fontWeight: 600 }}>ডেটা লোড হচ্ছে...</div>
            </div>
          ) : (
            <>
              {active === "dashboard" && isAdmin && <Dashboard projects={data.projects} clients={data.clients} employees={data.employees} transactions={data.transactions} materials={data.materials} />}
              {active === "projects" && isAdmin && <Projects data={data.projects} onRefresh={loadAll} />}
              {active === "construction" && <ConstructionProjects currentUser={currentUser} />}
              {active === "boq" && isAdmin && <BOQSystem />}
              {active === "clients" && isAdmin && <Clients data={data.clients} onRefresh={loadAll} />}
              {active === "employees" && isAdmin && <Employees data={data.employees} onRefresh={loadAll} />}
              {active === "attendance" && isAdmin && <Attendance employees={data.employees} />}
              {active === "finance" && isAdmin && <Finance data={data.transactions} onRefresh={loadAll} />}
              {active === "site" && isAdmin && <SiteProgress data={data.siteProgress} projects={data.projects} onRefresh={loadAll} />}
              {active === "materials" && isAdmin && <Materials data={data.materials} onRefresh={loadAll} />}
              {active === "analytics" && isAdmin && <Analytics transactions={data.transactions} projects={data.projects} employees={data.employees} />}
              {active === "users" && isAdmin && <UserManagement />}
              {active === "password" && <PasswordChange />}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
