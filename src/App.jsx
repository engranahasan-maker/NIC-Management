import { useState, useEffect } from "react";

// ============================================================
// COLOR TOKENS — NIC Brand: Navy #1A1A2E, Gold #C9A84C
// ============================================================
const C = {
  navy: "#1A1A2E",
  navyLight: "#16213E",
  navyMid: "#0F3460",
  gold: "#C9A84C",
  goldLight: "#E8C97A",
  goldPale: "#FDF6E3",
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

// ============================================================
// SAMPLE DATA
// ============================================================
const initialData = {
  projects: [
    { id: 1, name: "রহিম সাহেবের ডুপ্লেক্স বাড়ি", client: "আব্দুর রহিম", budget: 3500000, spent: 1200000, progress: 35, status: "চলমান", startDate: "2024-01-15", endDate: "2024-12-31", location: "ধানমন্ডি, ঢাকা", type: "আবাসিক নির্মাণ" },
    { id: 2, name: "করিম ট্রেডার্স অফিস ইন্টেরিয়র", client: "করিম সাহেব", budget: 850000, spent: 850000, progress: 100, status: "সম্পন্ন", startDate: "2023-10-01", endDate: "2024-01-30", location: "মতিঝিল, ঢাকা", type: "কমার্শিয়াল ইন্টেরিয়র" },
    { id: 3, name: "নতুন রেস্তোরাঁ ডিজাইন", client: "সালমা বেগম", budget: 1200000, spent: 150000, progress: 12, status: "নতুন", startDate: "2024-06-01", endDate: "2024-10-31", location: "গুলশান, ঢাকা", type: "কমার্শিয়াল ইন্টেরিয়র" },
    { id: 4, name: "মডুলার কিচেন ইন্সটলেশন", client: "রফিক সাহেব", budget: 450000, spent: 200000, progress: 45, status: "চলমান", startDate: "2024-05-01", endDate: "2024-07-31", location: "উত্তরা, ঢাকা", type: "মডুলার কিচেন" },
  ],
  clients: [
    { id: 1, name: "আব্দুর রহিম", phone: "01711-123456", email: "rahim@gmail.com", address: "ধানমন্ডি, ঢাকা", projects: 2, totalValue: 4200000, status: "সক্রিয়", joinDate: "2023-06-01" },
    { id: 2, name: "করিম সাহেব", phone: "01811-234567", email: "karim@gmail.com", address: "মতিঝিল, ঢাকা", projects: 1, totalValue: 850000, status: "সম্পন্ন", joinDate: "2023-09-15" },
    { id: 3, name: "সালমা বেগম", phone: "01911-345678", email: "salma@gmail.com", address: "গুলশান, ঢাকা", projects: 1, totalValue: 1200000, status: "সক্রিয়", joinDate: "2024-05-20" },
    { id: 4, name: "রফিক সাহেব", phone: "01611-456789", email: "rafiq@gmail.com", address: "উত্তরা, ঢাকা", projects: 1, totalValue: 450000, status: "সক্রিয়", joinDate: "2024-04-10" },
  ],
  employees: [
    { id: 1, name: "মোঃ রাকিব হোসেন", role: "সিনিয়র আর্কিটেক্ট", dept: "ডিজাইন", phone: "01712-111111", salary: 45000, joinDate: "2022-01-01", status: "কর্মরত" },
    { id: 2, name: "নাসরিন আক্তার", role: "ইন্টেরিয়র ডিজাইনার", dept: "ডিজাইন", phone: "01812-222222", salary: 35000, joinDate: "2022-06-15", status: "কর্মরত" },
    { id: 3, name: "কামাল হোসেন", role: "সাইট ইঞ্জিনিয়ার", dept: "নির্মাণ", phone: "01912-333333", salary: 30000, joinDate: "2023-03-01", status: "কর্মরত" },
    { id: 4, name: "সুমাইয়া খানম", role: "থ্রিডি ভিজুয়ালাইজার", dept: "ডিজাইন", phone: "01612-444444", salary: 28000, joinDate: "2023-07-01", status: "কর্মরত" },
    { id: 5, name: "জলিল মিয়া", role: "ফোরম্যান", dept: "নির্মাণ", phone: "01512-555555", salary: 22000, joinDate: "2022-09-01", status: "কর্মরত" },
    { id: 6, name: "রেজাউল করিম", role: "একাউন্টেন্ট", dept: "প্রশাসন", phone: "01412-666666", salary: 25000, joinDate: "2023-01-15", status: "কর্মরত" },
  ],
  attendance: {
    "2024-06-01": { 1: "উপস্থিত", 2: "উপস্থিত", 3: "অনুপস্থিত", 4: "উপস্থিত", 5: "উপস্থিত", 6: "উপস্থিত" },
    "2024-06-02": { 1: "উপস্থিত", 2: "অর্ধদিন", 3: "উপস্থিত", 4: "উপস্থিত", 5: "উপস্থিত", 6: "উপস্থিত" },
    "2024-06-03": { 1: "উপস্থিত", 2: "উপস্থিত", 3: "উপস্থিত", 4: "ছুটি", 5: "উপস্থিত", 6: "অনুপস্থিত" },
  },
  transactions: [
    { id: 1, date: "2024-06-01", type: "আয়", category: "প্রজেক্ট পেমেন্ট", desc: "রহিম সাহেবের অগ্রিম", amount: 500000, project: "রহিম সাহেবের ডুপ্লেক্স বাড়ি" },
    { id: 2, date: "2024-06-03", type: "ব্যয়", category: "কর্মী বেতন", desc: "জুন মাসের বেতন", amount: 185000, project: "—" },
    { id: 3, date: "2024-06-05", type: "ব্যয়", category: "নির্মাণ সামগ্রী", desc: "সিমেন্ট ও রড ক্রয়", amount: 320000, project: "রহিম সাহেবের ডুপ্লেক্স বাড়ি" },
    { id: 4, date: "2024-06-08", type: "আয়", category: "প্রজেক্ট পেমেন্ট", desc: "মডুলার কিচেন অগ্রিম", amount: 150000, project: "মডুলার কিচেন ইন্সটলেশন" },
    { id: 5, date: "2024-06-10", type: "ব্যয়", category: "অফিস খরচ", desc: "বিদ্যুৎ ও ইন্টারনেট বিল", amount: 8500, project: "—" },
    { id: 6, date: "2024-06-12", type: "আয়", category: "প্রজেক্ট পেমেন্ট", desc: "সালমা বেগমের অগ্রিম", amount: 200000, project: "নতুন রেস্তোরাঁ ডিজাইন" },
  ],
  siteProgress: [
    { id: 1, project: "রহিম সাহেবের ডুপ্লেক্স বাড়ি", date: "2024-06-10", work: "ফাউন্ডেশন কাজ সম্পন্ন", phase: "ফাউন্ডেশন", workers: 12, note: "মাটি পরীক্ষা রিপোর্ট পজিটিভ", status: "সম্পন্ন" },
    { id: 2, project: "রহিম সাহেবের ডুপ্লেক্স বাড়ি", date: "2024-06-12", work: "কলাম ঢালাই ১ম তলা", phase: "স্ট্রাকচার", workers: 15, note: "কংক্রিট মিক্স রেশিও ১:১.৫:৩", status: "চলমান" },
    { id: 3, project: "মডুলার কিচেন ইন্সটলেশন", date: "2024-06-11", work: "ক্যাবিনেট ফ্রেম ইন্সটল", phase: "ইন্সটলেশন", workers: 4, note: "জার্মান ফিটিংস ব্যবহার", status: "সম্পন্ন" },
    { id: 4, project: "নতুন রেস্তোরাঁ ডিজাইন", date: "2024-06-13", work: "থ্রিডি ডিজাইন অ্যাপ্রুভাল", phase: "ডিজাইন", workers: 2, note: "ক্লায়েন্ট রিভিশন চেয়েছেন", status: "পেন্ডিং" },
  ],
  materials: [
    { id: 1, name: "সিমেন্ট (৫০কেজি ব্যাগ)", unit: "ব্যাগ", stock: 145, minStock: 50, unitPrice: 520, supplier: "মেঘনা সিমেন্ট", lastPurchase: "2024-06-05" },
    { id: 2, name: "রড (১২মিমি)", unit: "কেজি", stock: 2800, minStock: 1000, unitPrice: 85, supplier: "BSRM", lastPurchase: "2024-06-05" },
    { id: 3, name: "ইট", unit: "পিস", stock: 8500, minStock: 3000, unitPrice: 12, supplier: "স্থানীয় সাপ্লায়ার", lastPurchase: "2024-06-01" },
    { id: 4, name: "বালি", unit: "সিএফটি", stock: 450, minStock: 200, unitPrice: 35, supplier: "নদী বালু সাপ্লায়ার", lastPurchase: "2024-05-28" },
    { id: 5, name: "কাঠ (ইঞ্চি)", unit: "কেজি", stock: 320, minStock: 500, unitPrice: 150, supplier: "ঢাকা উড মার্ট", lastPurchase: "2024-05-20" },
    { id: 6, name: "রং (লিটার)", unit: "লিটার", stock: 85, minStock: 30, unitPrice: 380, supplier: "বার্জার পেইন্টস", lastPurchase: "2024-06-08" },
  ],
};

// ============================================================
// UTILITY
// ============================================================
const fmt = (n) => "৳" + Number(n).toLocaleString("bn-BD");
const fmtNum = (n) => Number(n).toLocaleString("bn-BD");

function Badge({ label, color }) {
  const colors = {
    green: { bg: C.greenLight, text: C.green },
    red: { bg: C.redLight, text: C.red },
    yellow: { bg: C.yellowLight, text: "#856404" },
    blue: { bg: C.blueLight, text: C.blue },
    gray: { bg: C.gray200, text: C.gray600 },
    gold: { bg: "#FDF3D8", text: "#8a6a00" },
  };
  const s = colors[color] || colors.gray;
  return (
    <span style={{ background: s.bg, color: s.text, padding: "2px 10px", borderRadius: 20, fontSize: 12, fontWeight: 600 }}>
      {label}
    </span>
  );
}

function StatusBadge({ status }) {
  const map = {
    "চলমান": "blue", "সম্পন্ন": "green", "নতুন": "gold", "পেন্ডিং": "yellow",
    "সক্রিয়": "green", "কর্মরত": "green",
    "উপস্থিত": "green", "অনুপস্থিত": "red", "অর্ধদিন": "yellow", "ছুটি": "gray",
    "আয়": "green", "ব্যয়": "red",
  };
  return <Badge label={status} color={map[status] || "gray"} />;
}

function Card({ children, style }) {
  return (
    <div style={{ background: C.white, borderRadius: 12, padding: 20, boxShadow: "0 2px 12px rgba(0,0,0,0.07)", border: `1px solid ${C.gray200}`, ...style }}>
      {children}
    </div>
  );
}

function StatCard({ icon, label, value, sub, color }) {
  return (
    <Card style={{ display: "flex", alignItems: "center", gap: 16 }}>
      <div style={{ width: 52, height: 52, borderRadius: 12, background: color || C.goldPale, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24, flexShrink: 0 }}>
        {icon}
      </div>
      <div>
        <div style={{ fontSize: 22, fontWeight: 700, color: C.navy }}>{value}</div>
        <div style={{ fontSize: 13, color: C.gray600, marginTop: 2 }}>{label}</div>
        {sub && <div style={{ fontSize: 12, color: C.gold, marginTop: 2, fontWeight: 600 }}>{sub}</div>}
      </div>
    </Card>
  );
}

function ProgressBar({ value, color }) {
  return (
    <div style={{ background: C.gray200, borderRadius: 99, height: 8, overflow: "hidden" }}>
      <div style={{ width: `${value}%`, height: "100%", background: color || C.gold, borderRadius: 99, transition: "width 0.5s" }} />
    </div>
  );
}

function SectionHeader({ title, action, onAction }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
      <h2 style={{ margin: 0, color: C.navy, fontSize: 18, fontWeight: 700 }}>{title}</h2>
      {action && (
        <button onClick={onAction} style={{ background: C.gold, color: C.white, border: "none", borderRadius: 8, padding: "8px 16px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>
          + {action}
        </button>
      )}
    </div>
  );
}

function Table({ headers, rows }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
        <thead>
          <tr style={{ background: C.gray50 }}>
            {headers.map((h, i) => (
              <th key={i} style={{ padding: "10px 14px", textAlign: "left", color: C.gray600, fontWeight: 600, borderBottom: `2px solid ${C.gray200}`, whiteSpace: "nowrap" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{ borderBottom: `1px solid ${C.gray100}`, transition: "background 0.15s" }}
              onMouseEnter={e => e.currentTarget.style.background = C.gray50}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
              {row.map((cell, j) => (
                <td key={j} style={{ padding: "10px 14px", color: C.gray800, verticalAlign: "middle" }}>{cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ============================================================
// MODAL
// ============================================================
function Modal({ title, onClose, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 16, width: "100%", maxWidth: 520, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 24px", borderBottom: `1px solid ${C.gray200}`, position: "sticky", top: 0, background: C.white }}>
          <h3 style={{ margin: 0, color: C.navy, fontSize: 16, fontWeight: 700 }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: C.gray600, lineHeight: 1 }}>×</button>
        </div>
        <div style={{ padding: 24 }}>{children}</div>
      </div>
    </div>
  );
}

function FormField({ label, children }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <label style={{ display: "block", fontSize: 13, fontWeight: 600, color: C.gray800, marginBottom: 6 }}>{label}</label>
      {children}
    </div>
  );
}

const inputStyle = { width: "100%", padding: "9px 12px", border: `1px solid ${C.gray200}`, borderRadius: 8, fontSize: 13, color: C.gray800, outline: "none", boxSizing: "border-box", fontFamily: "inherit" };

// ============================================================
// DASHBOARD
// ============================================================
function Dashboard({ data }) {
  const totalIncome = data.transactions.filter(t => t.type === "আয়").reduce((s, t) => s + t.amount, 0);
  const totalExpense = data.transactions.filter(t => t.type === "ব্যয়").reduce((s, t) => s + t.amount, 0);
  const activeProjects = data.projects.filter(p => p.status === "চলমান").length;
  const lowStock = data.materials.filter(m => m.stock < m.minStock).length;

  return (
    <div>
      <div style={{ marginBottom: 24, padding: "16px 20px", background: `linear-gradient(135deg, ${C.navy}, ${C.navyMid})`, borderRadius: 14, color: C.white }}>
        <div style={{ fontSize: 13, opacity: 0.8 }}>স্বাগতম,</div>
        <div style={{ fontSize: 20, fontWeight: 700, marginTop: 4 }}>Noksha Interior & Construction</div>
        <div style={{ fontSize: 12, opacity: 0.7, marginTop: 4 }}>ড্যাশবোর্ড ওভারভিউ — আজকের তারিখ: {new Date().toLocaleDateString("bn-BD")}</div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 14, marginBottom: 24 }}>
        <StatCard icon="🏗️" label="সক্রিয় প্রজেক্ট" value={fmtNum(activeProjects)} sub={`মোট ${data.projects.length}টি`} color="#EEF2FF" />
        <StatCard icon="👥" label="মোট ক্লায়েন্ট" value={fmtNum(data.clients.length)} sub="নিবন্ধিত" color="#F0FFF4" />
        <StatCard icon="👷" label="কর্মী সংখ্যা" value={fmtNum(data.employees.length)} sub="কর্মরত" color="#FFF8E1" />
        <StatCard icon="💰" label="মাসিক আয়" value={fmt(totalIncome)} sub="এই মাসে" color="#F0FFF4" />
        <StatCard icon="📦" label="কম স্টক" value={fmtNum(lowStock)} sub="সামগ্রী" color={lowStock > 0 ? C.redLight : C.greenLight} />
        <StatCard icon="💸" label="মাসিক ব্যয়" value={fmt(totalExpense)} sub="এই মাসে" color="#FFF5F5" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
        <Card>
          <SectionHeader title="চলমান প্রজেক্ট" />
          {data.projects.filter(p => p.status === "চলমান").map(p => (
            <div key={p.id} style={{ marginBottom: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                <span style={{ fontSize: 13, fontWeight: 600, color: C.navy }}>{p.name}</span>
                <span style={{ fontSize: 12, color: C.gold, fontWeight: 700 }}>{p.progress}%</span>
              </div>
              <ProgressBar value={p.progress} />
              <div style={{ fontSize: 11, color: C.gray400, marginTop: 4 }}>{p.client}</div>
            </div>
          ))}
        </Card>

        <Card>
          <SectionHeader title="সর্বশেষ লেনদেন" />
          {data.transactions.slice(-5).reverse().map(t => (
            <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.gray100}` }}>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.gray800 }}>{t.desc}</div>
                <div style={{ fontSize: 11, color: C.gray400 }}>{t.date}</div>
              </div>
              <div style={{ fontWeight: 700, color: t.type === "আয়" ? C.green : C.red, fontSize: 13 }}>
                {t.type === "আয়" ? "+" : "-"}{fmt(t.amount)}
              </div>
            </div>
          ))}
        </Card>
      </div>

      <Card>
        <SectionHeader title="কম স্টক সতর্কতা ⚠️" />
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
          {data.materials.filter(m => m.stock < m.minStock).map(m => (
            <div key={m.id} style={{ background: C.redLight, border: `1px solid #F5C6CB`, borderRadius: 10, padding: "10px 14px", minWidth: 160 }}>
              <div style={{ fontWeight: 700, color: C.red, fontSize: 13 }}>{m.name}</div>
              <div style={{ fontSize: 12, color: C.gray600 }}>বর্তমান: {fmtNum(m.stock)} {m.unit}</div>
              <div style={{ fontSize: 12, color: C.gray600 }}>ন্যূনতম: {fmtNum(m.minStock)} {m.unit}</div>
            </div>
          ))}
          {data.materials.filter(m => m.stock < m.minStock).length === 0 && (
            <div style={{ color: C.green, fontWeight: 600, fontSize: 13 }}>✅ সকল সামগ্রীর পর্যাপ্ত স্টক আছে</div>
          )}
        </div>
      </Card>
    </div>
  );
}

// ============================================================
// PROJECTS
// ============================================================
function Projects({ data, setData }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", client: "", budget: "", startDate: "", endDate: "", location: "", type: "আবাসিক নির্মাণ", status: "নতুন", progress: 0 });

  const save = () => {
    if (!form.name || !form.client || !form.budget) return alert("নাম, ক্লায়েন্ট ও বাজেট আবশ্যক");
    setData(d => ({ ...d, projects: [...d.projects, { ...form, id: Date.now(), spent: 0, budget: +form.budget, progress: +form.progress }] }));
    setShowModal(false);
    setForm({ name: "", client: "", budget: "", startDate: "", endDate: "", location: "", type: "আবাসিক নির্মাণ", status: "নতুন", progress: 0 });
  };

  return (
    <div>
      <SectionHeader title="প্রজেক্ট ব্যবস্থাপনা" action="নতুন প্রজেক্ট" onAction={() => setShowModal(true)} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 16 }}>
        {data.projects.map(p => (
          <Card key={p.id}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
              <div>
                <div style={{ fontWeight: 700, color: C.navy, fontSize: 14 }}>{p.name}</div>
                <div style={{ fontSize: 12, color: C.gray400, marginTop: 3 }}>📍 {p.location}</div>
              </div>
              <StatusBadge status={p.status} />
            </div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}>ক্লায়েন্ট: <strong>{p.client}</strong></div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}>ধরন: {p.type}</div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 12 }}>সময়: {p.startDate} → {p.endDate}</div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ fontSize: 12, color: C.gray600 }}>অগ্রগতি</span>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.gold }}>{p.progress}%</span>
            </div>
            <ProgressBar value={p.progress} color={p.progress === 100 ? C.green : C.gold} />
            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 12, paddingTop: 12, borderTop: `1px solid ${C.gray100}` }}>
              <div>
                <div style={{ fontSize: 11, color: C.gray400 }}>বাজেট</div>
                <div style={{ fontWeight: 700, color: C.navy, fontSize: 13 }}>{fmt(p.budget)}</div>
              </div>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 11, color: C.gray400 }}>খরচ হয়েছে</div>
                <div style={{ fontWeight: 700, color: p.spent > p.budget ? C.red : C.green, fontSize: 13 }}>{fmt(p.spent)}</div>
              </div>
            </div>
          </Card>
        ))}
      </div>

      {showModal && (
        <Modal title="নতুন প্রজেক্ট যোগ করুন" onClose={() => setShowModal(false)}>
          <FormField label="প্রজেক্টের নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="প্রজেক্টের নাম" /></FormField>
          <FormField label="ক্লায়েন্টের নাম *"><input style={inputStyle} value={form.client} onChange={e => setForm({ ...form, client: e.target.value })} placeholder="ক্লায়েন্টের নাম" /></FormField>
          <FormField label="বাজেট (টাকা) *"><input style={inputStyle} type="number" value={form.budget} onChange={e => setForm({ ...form, budget: e.target.value })} placeholder="০" /></FormField>
          <FormField label="অবস্থান"><input style={inputStyle} value={form.location} onChange={e => setForm({ ...form, location: e.target.value })} placeholder="ঠিকানা" /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="শুরুর তারিখ"><input style={inputStyle} type="date" value={form.startDate} onChange={e => setForm({ ...form, startDate: e.target.value })} /></FormField>
            <FormField label="শেষের তারিখ"><input style={inputStyle} type="date" value={form.endDate} onChange={e => setForm({ ...form, endDate: e.target.value })} /></FormField>
          </div>
          <FormField label="ধরন">
            <select style={inputStyle} value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}>
              {["আবাসিক নির্মাণ", "কমার্শিয়াল ইন্টেরিয়র", "মডুলার কিচেন", "থ্রিডি ভিজুয়ালাইজেশন", "সিভিল কনস্ট্রাকশন"].map(t => <option key={t}>{t}</option>)}
            </select>
          </FormField>
          <button onClick={save} style={{ width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 10, padding: "12px", fontWeight: 700, cursor: "pointer", fontSize: 14, marginTop: 8 }}>
            ✅ সংরক্ষণ করুন
          </button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// CLIENTS
// ============================================================
function Clients({ data, setData }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", phone: "", email: "", address: "", status: "সক্রিয়" });

  const save = () => {
    if (!form.name || !form.phone) return alert("নাম ও ফোন নম্বর আবশ্যক");
    setData(d => ({ ...d, clients: [...d.clients, { ...form, id: Date.now(), projects: 0, totalValue: 0, joinDate: new Date().toISOString().split("T")[0] }] }));
    setShowModal(false);
    setForm({ name: "", phone: "", email: "", address: "", status: "সক্রিয়" });
  };

  return (
    <div>
      <SectionHeader title="ক্লায়েন্ট ব্যবস্থাপনা" action="নতুন ক্লায়েন্ট" onAction={() => setShowModal(true)} />
      <Card>
        <Table
          headers={["নাম", "ফোন", "ইমেইল", "ঠিকানা", "প্রজেক্ট", "মোট মূল্য", "স্ট্যাটাস"]}
          rows={data.clients.map(c => [
            <strong style={{ color: C.navy }}>{c.name}</strong>,
            c.phone,
            c.email || "—",
            c.address,
            <Badge label={`${c.projects}টি`} color="blue" />,
            <span style={{ fontWeight: 700, color: C.green }}>{fmt(c.totalValue)}</span>,
            <StatusBadge status={c.status} />,
          ])}
        />
      </Card>
      {showModal && (
        <Modal title="নতুন ক্লায়েন্ট যোগ করুন" onClose={() => setShowModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="পূর্ণ নাম" /></FormField>
          <FormField label="ফোন *"><input style={inputStyle} value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} placeholder="০১X-XXXXXXXX" /></FormField>
          <FormField label="ইমেইল"><input style={inputStyle} value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} placeholder="email@example.com" /></FormField>
          <FormField label="ঠিকানা"><input style={inputStyle} value={form.address} onChange={e => setForm({ ...form, address: e.target.value })} placeholder="পূর্ণ ঠিকানা" /></FormField>
          <button onClick={save} style={{ width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 10, padding: "12px", fontWeight: 700, cursor: "pointer", fontSize: 14, marginTop: 8 }}>
            ✅ সংরক্ষণ করুন
          </button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// EMPLOYEES
// ============================================================
function Employees({ data, setData }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", role: "", dept: "ডিজাইন", phone: "", salary: "", joinDate: "", status: "কর্মরত" });

  const save = () => {
    if (!form.name || !form.role || !form.salary) return alert("নাম, পদবি ও বেতন আবশ্যক");
    setData(d => ({ ...d, employees: [...d.employees, { ...form, id: Date.now(), salary: +form.salary }] }));
    setShowModal(false);
    setForm({ name: "", role: "", dept: "ডিজাইন", phone: "", salary: "", joinDate: "", status: "কর্মরত" });
  };

  const totalSalary = data.employees.reduce((s, e) => s + e.salary, 0);

  return (
    <div>
      <SectionHeader title="কর্মী ব্যবস্থাপনা (HR)" action="নতুন কর্মী" onAction={() => setShowModal(true)} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="👷" label="মোট কর্মী" value={fmtNum(data.employees.length)} color="#EEF2FF" />
        <StatCard icon="💰" label="মাসিক বেতন বিল" value={fmt(totalSalary)} color="#FFF8E1" />
        <StatCard icon="🏢" label="বিভাগ সংখ্যা" value={fmtNum([...new Set(data.employees.map(e => e.dept))].length)} color="#F0FFF4" />
      </div>
      <Card>
        <Table
          headers={["নাম", "পদবি", "বিভাগ", "ফোন", "বেতন", "যোগদান", "স্ট্যাটাস"]}
          rows={data.employees.map(e => [
            <strong style={{ color: C.navy }}>{e.name}</strong>,
            e.role,
            <Badge label={e.dept} color="blue" />,
            e.phone,
            <span style={{ fontWeight: 700, color: C.green }}>{fmt(e.salary)}</span>,
            e.joinDate,
            <StatusBadge status={e.status} />,
          ])}
        />
      </Card>
      {showModal && (
        <Modal title="নতুন কর্মী যোগ করুন" onClose={() => setShowModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="পূর্ণ নাম" /></FormField>
          <FormField label="পদবি *"><input style={inputStyle} value={form.role} onChange={e => setForm({ ...form, role: e.target.value })} placeholder="পদবি" /></FormField>
          <FormField label="বিভাগ">
            <select style={inputStyle} value={form.dept} onChange={e => setForm({ ...form, dept: e.target.value })}>
              {["ডিজাইন", "নির্মাণ", "প্রশাসন", "বিপণন"].map(d => <option key={d}>{d}</option>)}
            </select>
          </FormField>
          <FormField label="ফোন"><input style={inputStyle} value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} placeholder="০১X-XXXXXXXX" /></FormField>
          <FormField label="মাসিক বেতন (টাকা) *"><input style={inputStyle} type="number" value={form.salary} onChange={e => setForm({ ...form, salary: e.target.value })} placeholder="০" /></FormField>
          <FormField label="যোগদানের তারিখ"><input style={inputStyle} type="date" value={form.joinDate} onChange={e => setForm({ ...form, joinDate: e.target.value })} /></FormField>
          <button onClick={save} style={{ width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 10, padding: "12px", fontWeight: 700, cursor: "pointer", fontSize: 14, marginTop: 8 }}>
            ✅ সংরক্ষণ করুন
          </button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// ATTENDANCE
// ============================================================
function Attendance({ data, setData }) {
  const today = new Date().toISOString().split("T")[0];
  const [selDate, setSelDate] = useState(today);

  const todayAtt = data.attendance[selDate] || {};

  const setAtt = (empId, val) => {
    setData(d => ({
      ...d,
      attendance: {
        ...d.attendance,
        [selDate]: { ...(d.attendance[selDate] || {}), [empId]: val }
      }
    }));
  };

  const statuses = ["উপস্থিত", "অনুপস্থিত", "অর্ধদিন", "ছুটি"];
  const counts = statuses.map(s => ({ s, n: Object.values(todayAtt).filter(v => v === s).length }));

  return (
    <div>
      <SectionHeader title="উপস্থিতি ব্যবস্থাপনা" />
      <div style={{ display: "flex", gap: 12, alignItems: "center", marginBottom: 20 }}>
        <div style={{ fontWeight: 600, color: C.gray800, fontSize: 14 }}>তারিখ নির্বাচন:</div>
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
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: C.gray50 }}>
              <th style={{ padding: "10px 14px", textAlign: "left", color: C.gray600, fontWeight: 600, borderBottom: `2px solid ${C.gray200}` }}>কর্মীর নাম</th>
              <th style={{ padding: "10px 14px", textAlign: "left", color: C.gray600, fontWeight: 600, borderBottom: `2px solid ${C.gray200}` }}>পদবি</th>
              {statuses.map(s => (
                <th key={s} style={{ padding: "10px 14px", textAlign: "center", color: C.gray600, fontWeight: 600, borderBottom: `2px solid ${C.gray200}` }}>{s}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.employees.map(e => (
              <tr key={e.id} style={{ borderBottom: `1px solid ${C.gray100}` }}>
                <td style={{ padding: "10px 14px", fontWeight: 600, color: C.navy }}>{e.name}</td>
                <td style={{ padding: "10px 14px", color: C.gray600, fontSize: 12 }}>{e.role}</td>
                {statuses.map(s => (
                  <td key={s} style={{ padding: "6px 14px", textAlign: "center" }}>
                    <input type="radio" name={`att-${e.id}`} checked={todayAtt[e.id] === s} onChange={() => setAtt(e.id, s)}
                      style={{ accentColor: s === "উপস্থিত" ? C.green : s === "অনুপস্থিত" ? C.red : C.gold, width: 16, height: 16, cursor: "pointer" }} />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ marginTop: 16, display: "flex", justifyContent: "flex-end" }}>
          <div style={{ fontSize: 13, color: C.green, fontWeight: 600 }}>✅ পরিবর্তন স্বয়ংক্রিয়ভাবে সংরক্ষিত হচ্ছে</div>
        </div>
      </Card>
    </div>
  );
}

// ============================================================
// FINANCE
// ============================================================
function Finance({ data, setData }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ date: new Date().toISOString().split("T")[0], type: "আয়", category: "প্রজেক্ট পেমেন্ট", desc: "", amount: "", project: "—" });

  const totalIncome = data.transactions.filter(t => t.type === "আয়").reduce((s, t) => s + t.amount, 0);
  const totalExpense = data.transactions.filter(t => t.type === "ব্যয়").reduce((s, t) => s + t.amount, 0);

  const save = () => {
    if (!form.desc || !form.amount) return alert("বিবরণ ও পরিমাণ আবশ্যক");
    setData(d => ({ ...d, transactions: [...d.transactions, { ...form, id: Date.now(), amount: +form.amount }] }));
    setShowModal(false);
    setForm({ date: new Date().toISOString().split("T")[0], type: "আয়", category: "প্রজেক্ট পেমেন্ট", desc: "", amount: "", project: "—" });
  };

  return (
    <div>
      <SectionHeader title="আর্থিক ব্যবস্থাপনা" action="নতুন লেনদেন" onAction={() => setShowModal(true)} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="💰" label="মোট আয়" value={fmt(totalIncome)} color="#F0FFF4" />
        <StatCard icon="💸" label="মোট ব্যয়" value={fmt(totalExpense)} color="#FFF5F5" />
        <StatCard icon="📊" label="নিট লাভ" value={fmt(totalIncome - totalExpense)} color={totalIncome >= totalExpense ? "#F0FFF4" : "#FFF5F5"} />
      </div>
      <Card>
        <Table
          headers={["তারিখ", "ধরন", "বিভাগ", "বিবরণ", "প্রজেক্ট", "পরিমাণ"]}
          rows={[...data.transactions].reverse().map(t => [
            t.date,
            <StatusBadge status={t.type} />,
            t.category,
            t.desc,
            <span style={{ fontSize: 12, color: C.gray400 }}>{t.project}</span>,
            <span style={{ fontWeight: 700, color: t.type === "আয়" ? C.green : C.red }}>
              {t.type === "আয়" ? "+" : "-"}{fmt(t.amount)}
            </span>,
          ])}
        />
      </Card>
      {showModal && (
        <Modal title="নতুন লেনদেন যোগ করুন" onClose={() => setShowModal(false)}>
          <FormField label="তারিখ"><input style={inputStyle} type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} /></FormField>
          <FormField label="ধরন">
            <select style={inputStyle} value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}>
              <option>আয়</option><option>ব্যয়</option>
            </select>
          </FormField>
          <FormField label="বিভাগ">
            <select style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>
              {["প্রজেক্ট পেমেন্ট", "কর্মী বেতন", "নির্মাণ সামগ্রী", "অফিস খরচ", "পরিবহন", "বিবিধ"].map(c => <option key={c}>{c}</option>)}
            </select>
          </FormField>
          <FormField label="বিবরণ *"><input style={inputStyle} value={form.desc} onChange={e => setForm({ ...form, desc: e.target.value })} placeholder="লেনদেনের বিবরণ" /></FormField>
          <FormField label="পরিমাণ (টাকা) *"><input style={inputStyle} type="number" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} placeholder="০" /></FormField>
          <FormField label="সংশ্লিষ্ট প্রজেক্ট">
            <select style={inputStyle} value={form.project} onChange={e => setForm({ ...form, project: e.target.value })}>
              <option value="—">— নির্বাচন করুন —</option>
              {data.projects.map(p => <option key={p.id} value={p.name}>{p.name}</option>)}
            </select>
          </FormField>
          <button onClick={save} style={{ width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 10, padding: "12px", fontWeight: 700, cursor: "pointer", fontSize: 14, marginTop: 8 }}>
            ✅ সংরক্ষণ করুন
          </button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// SITE PROGRESS
// ============================================================
function SiteProgress({ data, setData }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ project: "", date: new Date().toISOString().split("T")[0], work: "", phase: "ফাউন্ডেশন", workers: "", note: "", status: "চলমান" });

  const save = () => {
    if (!form.project || !form.work) return alert("প্রজেক্ট ও কাজের বিবরণ আবশ্যক");
    setData(d => ({ ...d, siteProgress: [...d.siteProgress, { ...form, id: Date.now(), workers: +form.workers }] }));
    setShowModal(false);
    setForm({ project: "", date: new Date().toISOString().split("T")[0], work: "", phase: "ফাউন্ডেশন", workers: "", note: "", status: "চলমান" });
  };

  return (
    <div>
      <SectionHeader title="সাইট অগ্রগতি ট্র্যাকিং" action="নতুন আপডেট" onAction={() => setShowModal(true)} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 16 }}>
        {[...data.siteProgress].reverse().map(s => (
          <Card key={s.id}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
              <Badge label={s.phase} color="blue" />
              <StatusBadge status={s.status} />
            </div>
            <div style={{ fontWeight: 700, color: C.navy, fontSize: 14, marginBottom: 6 }}>{s.work}</div>
            <div style={{ fontSize: 12, color: C.gold, fontWeight: 600, marginBottom: 8 }}>📋 {s.project}</div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}>📅 তারিখ: {s.date}</div>
            <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}>👷 শ্রমিক: {fmtNum(s.workers)} জন</div>
            {s.note && <div style={{ fontSize: 12, color: C.gray400, marginTop: 8, padding: "8px", background: C.gray50, borderRadius: 6 }}>📝 {s.note}</div>}
          </Card>
        ))}
      </div>
      {showModal && (
        <Modal title="সাইট আপডেট যোগ করুন" onClose={() => setShowModal(false)}>
          <FormField label="প্রজেক্ট *">
            <select style={inputStyle} value={form.project} onChange={e => setForm({ ...form, project: e.target.value })}>
              <option value="">— নির্বাচন করুন —</option>
              {data.projects.map(p => <option key={p.id} value={p.name}>{p.name}</option>)}
            </select>
          </FormField>
          <FormField label="তারিখ"><input style={inputStyle} type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} /></FormField>
          <FormField label="কাজের বিবরণ *"><input style={inputStyle} value={form.work} onChange={e => setForm({ ...form, work: e.target.value })} placeholder="আজকের কাজ" /></FormField>
          <FormField label="নির্মাণ পর্যায়">
            <select style={inputStyle} value={form.phase} onChange={e => setForm({ ...form, phase: e.target.value })}>
              {["ডিজাইন", "ফাউন্ডেশন", "স্ট্রাকচার", "ওয়াল", "ছাদ", "ফিনিশিং", "ইন্টেরিয়র", "ইন্সটলেশন"].map(p => <option key={p}>{p}</option>)}
            </select>
          </FormField>
          <FormField label="শ্রমিক সংখ্যা"><input style={inputStyle} type="number" value={form.workers} onChange={e => setForm({ ...form, workers: e.target.value })} placeholder="০" /></FormField>
          <FormField label="নোট/মন্তব্য"><input style={inputStyle} value={form.note} onChange={e => setForm({ ...form, note: e.target.value })} placeholder="অতিরিক্ত তথ্য" /></FormField>
          <FormField label="স্ট্যাটাস">
            <select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>
              {["চলমান", "সম্পন্ন", "পেন্ডিং"].map(s => <option key={s}>{s}</option>)}
            </select>
          </FormField>
          <button onClick={save} style={{ width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 10, padding: "12px", fontWeight: 700, cursor: "pointer", fontSize: 14, marginTop: 8 }}>
            ✅ সংরক্ষণ করুন
          </button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// MATERIALS
// ============================================================
function Materials({ data, setData }) {
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", unit: "পিস", stock: "", minStock: "", unitPrice: "", supplier: "", lastPurchase: new Date().toISOString().split("T")[0] });

  const save = () => {
    if (!form.name || !form.stock) return alert("নাম ও স্টক পরিমাণ আবশ্যক");
    setData(d => ({ ...d, materials: [...d.materials, { ...form, id: Date.now(), stock: +form.stock, minStock: +form.minStock, unitPrice: +form.unitPrice }] }));
    setShowModal(false);
    setForm({ name: "", unit: "পিস", stock: "", minStock: "", unitPrice: "", supplier: "", lastPurchase: new Date().toISOString().split("T")[0] });
  };

  const totalValue = data.materials.reduce((s, m) => s + m.stock * m.unitPrice, 0);

  return (
    <div>
      <SectionHeader title="সামগ্রী ও ইনভেন্টরি" action="নতুন সামগ্রী" onAction={() => setShowModal(true)} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="📦" label="মোট সামগ্রী" value={fmtNum(data.materials.length)} color="#EEF2FF" />
        <StatCard icon="⚠️" label="কম স্টক" value={fmtNum(data.materials.filter(m => m.stock < m.minStock).length)} color={data.materials.filter(m => m.stock < m.minStock).length > 0 ? C.redLight : C.greenLight} />
        <StatCard icon="💰" label="মোট মূল্যমান" value={fmt(totalValue)} color="#FFF8E1" />
      </div>
      <Card>
        <Table
          headers={["সামগ্রী", "একক", "স্টক", "ন্যূনতম", "একক মূল্য", "মোট মূল্য", "সাপ্লায়ার", "স্ট্যাটাস"]}
          rows={data.materials.map(m => [
            <strong style={{ color: C.navy }}>{m.name}</strong>,
            m.unit,
            <span style={{ fontWeight: 600, color: m.stock < m.minStock ? C.red : C.gray800 }}>{fmtNum(m.stock)}</span>,
            fmtNum(m.minStock),
            fmt(m.unitPrice),
            <span style={{ fontWeight: 700, color: C.green }}>{fmt(m.stock * m.unitPrice)}</span>,
            <span style={{ fontSize: 12 }}>{m.supplier}</span>,
            m.stock < m.minStock ? <Badge label="কম স্টক ⚠️" color="red" /> : <Badge label="পর্যাপ্ত ✅" color="green" />,
          ])}
        />
      </Card>
      {showModal && (
        <Modal title="নতুন সামগ্রী যোগ করুন" onClose={() => setShowModal(false)}>
          <FormField label="সামগ্রীর নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="সামগ্রীর নাম" /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="একক">
              <select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>
                {["পিস", "কেজি", "ব্যাগ", "লিটার", "সিএফটি", "ফুট", "মিটার"].map(u => <option key={u}>{u}</option>)}
              </select>
            </FormField>
            <FormField label="বর্তমান স্টক *"><input style={inputStyle} type="number" value={form.stock} onChange={e => setForm({ ...form, stock: e.target.value })} placeholder="০" /></FormField>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="ন্যূনতম স্টক"><input style={inputStyle} type="number" value={form.minStock} onChange={e => setForm({ ...form, minStock: e.target.value })} placeholder="০" /></FormField>
            <FormField label="একক মূল্য (৳)"><input style={inputStyle} type="number" value={form.unitPrice} onChange={e => setForm({ ...form, unitPrice: e.target.value })} placeholder="০" /></FormField>
          </div>
          <FormField label="সাপ্লায়ার"><input style={inputStyle} value={form.supplier} onChange={e => setForm({ ...form, supplier: e.target.value })} placeholder="সাপ্লায়ারের নাম" /></FormField>
          <button onClick={save} style={{ width: "100%", background: C.gold, color: C.white, border: "none", borderRadius: 10, padding: "12px", fontWeight: 700, cursor: "pointer", fontSize: 14, marginTop: 8 }}>
            ✅ সংরক্ষণ করুন
          </button>
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
  const [active, setActive] = useState("dashboard");
  const [data, setData] = useState(initialData);
  const [sideOpen, setSideOpen] = useState(true);

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.gray100, fontFamily: "'Noto Sans Bengali', 'Hind Siliguri', Arial, sans-serif" }}>
      {/* SIDEBAR */}
      <div style={{
        width: sideOpen ? 220 : 64,
        background: C.navy,
        display: "flex",
        flexDirection: "column",
        transition: "width 0.3s",
        flexShrink: 0,
        overflow: "hidden",
        position: "sticky",
        top: 0,
        height: "100vh",
      }}>
        {/* Logo */}
        <div style={{ padding: sideOpen ? "20px 20px 16px" : "20px 12px 16px", borderBottom: `1px solid rgba(255,255,255,0.1)` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 36, height: 36, background: C.gold, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>🏛️</div>
            {sideOpen && (
              <div>
                <div style={{ color: C.gold, fontWeight: 800, fontSize: 14, lineHeight: 1.2 }}>নকশা</div>
                <div style={{ color: "rgba(255,255,255,0.6)", fontSize: 10 }}>NIC Management</div>
              </div>
            )}
          </div>
        </div>

        {/* Nav */}
        <nav style={{ flex: 1, padding: "12px 0", overflowY: "auto" }}>
          {MENU.map(m => (
            <button key={m.id} onClick={() => setActive(m.id)}
              style={{
                display: "flex", alignItems: "center", gap: 12,
                width: "100%", padding: sideOpen ? "11px 20px" : "11px 0", justifyContent: sideOpen ? "flex-start" : "center",
                background: active === m.id ? `rgba(201,168,76,0.15)` : "none",
                border: "none", borderLeft: active === m.id ? `3px solid ${C.gold}` : "3px solid transparent",
                color: active === m.id ? C.gold : "rgba(255,255,255,0.65)",
                cursor: "pointer", fontSize: 13, fontWeight: active === m.id ? 700 : 400,
                transition: "all 0.15s", fontFamily: "inherit",
              }}>
              <span style={{ fontSize: 18, flexShrink: 0 }}>{m.icon}</span>
              {sideOpen && <span style={{ whiteSpace: "nowrap" }}>{m.label}</span>}
            </button>
          ))}
        </nav>

        {/* Toggle */}
        <button onClick={() => setSideOpen(o => !o)}
          style={{ background: "rgba(255,255,255,0.1)", border: "none", color: C.white, padding: "12px", cursor: "pointer", fontSize: 16, fontFamily: "inherit" }}>
          {sideOpen ? "◀" : "▶"}
        </button>
      </div>

      {/* MAIN */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        {/* Topbar */}
        <div style={{ background: C.white, borderBottom: `1px solid ${C.gray200}`, padding: "14px 24px", display: "flex", justifyContent: "space-between", alignItems: "center", position: "sticky", top: 0, zIndex: 100 }}>
          <div>
            <div style={{ fontWeight: 700, color: C.navy, fontSize: 16 }}>{MENU.find(m => m.id === active)?.label}</div>
            <div style={{ fontSize: 11, color: C.gray400, marginTop: 2 }}>Noksha Interior & Construction · ফরিদপুর & ঢাকা</div>
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <div style={{ width: 36, height: 36, background: `linear-gradient(135deg, ${C.navy}, ${C.navyMid})`, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", color: C.gold, fontWeight: 700, fontSize: 14 }}>R</div>
            <div style={{ fontSize: 12, color: C.gray600 }}>
              <div style={{ fontWeight: 600, color: C.navy }}>মোঃ রানা</div>
              <div>প্রতিষ্ঠাতা ও পরিচালক</div>
            </div>
          </div>
        </div>

        {/* Content */}
        <div style={{ flex: 1, padding: "24px", overflowY: "auto" }}>
          {active === "dashboard" && <Dashboard data={data} />}
          {active === "projects" && <Projects data={data} setData={setData} />}
          {active === "clients" && <Clients data={data} setData={setData} />}
          {active === "employees" && <Employees data={data} setData={setData} />}
          {active === "attendance" && <Attendance data={data} setData={setData} />}
          {active === "finance" && <Finance data={data} setData={setData} />}
          {active === "site" && <SiteProgress data={data} setData={setData} />}
          {active === "materials" && <Materials data={data} setData={setData} />}
        </div>
      </div>
    </div>
  );
}
