import { useState, useEffect, useRef, Fragment } from "react";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://jijxnycopycsysugppnw.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImppanhueWNvcHljc3lzdWdwcG53Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI3MDAyNTAsImV4cCI6MjA5ODI3NjI1MH0.7kXRGGnW4VdWU9VT1XEBKp5oC9V5Z21KA_PBqZtvjJA";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
const NOKSHA_ICON_URL = "https://jijxnycopycsysugppnw.supabase.co/storage/v1/object/public/Upload%20images/icon.png";


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

const numToWordsTaka = (num) => {
  num = Math.round(Number(num) || 0);
  if (num <= 0) return "Zero Taka Only";
  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  const twoDigits = (n) => n < 20 ? ones[n] : tens[Math.floor(n / 10)] + (n % 10 ? " " + ones[n % 10] : "");
  const threeDigits = (n) => { const h = Math.floor(n / 100), r = n % 100; return (h ? ones[h] + " Hundred " : "") + (r ? twoDigits(r) : ""); };
  let n = num, result = "";
  const crore = Math.floor(n / 10000000); n %= 10000000;
  const lakh = Math.floor(n / 100000); n %= 100000;
  const thousand = Math.floor(n / 1000); n %= 1000;
  const rest = n;
  if (crore) result += threeDigits(crore) + " Crore ";
  if (lakh) result += twoDigits(lakh) + " Lakh ";
  if (thousand) result += twoDigits(thousand) + " Thousand ";
  if (rest) result += threeDigits(rest);
  return result.trim() + " Taka Only";
};

// ============================================================
// TRANSLATIONS (BN/EN)
// ============================================================
const TXT = {
  bn: {
    dashboard: "ড্যাশবোর্ড", projects: "প্রজেক্ট", construction: "Construction Projects", interior: "Interior Projects",
    boq: "Estimate Project", clients: "ক্লায়েন্ট", employees: "কর্মী (HR)", attendance: "উপস্থিতি",
    smart_attendance: "স্মার্ট অ্যাটেন্ডেন্স", finance: "আর্থিক", site: "সাইট প্রগ্রেস", materials: "সামগ্রী",
    analytics: "রিপোর্ট & Analytics", users: "User Management", password: "পাসওয়ার্ড",
    good_morning: "শুভ সকাল", have_productive_day: "আপনার দিনটি প্রোডাক্টিভ হোক!",
    todays_attendance: "আজকের উপস্থিতি", live_status: "লাইভ স্ট্যাটাস",
    present: "উপস্থিত", absent: "অনুপস্থিত", late: "দেরি", on_leave: "ছুটিতে", total: "মোট",
    department_wise_summary: "বিভাগ ভিত্তিক সারাংশ", recent_checkins: "সাম্প্রতিক চেক-ইন",
    quick_checkin: "দ্রুত চেক-ইন", not_marked: "চিহ্নিত হয়নি", mark: "চিহ্নিত করুন",
    refresh: "রিফ্রেশ", search_employee: "কর্মী খুঁজুন...", no_records: "কোনো রেকর্ড নেই",
    live: "লাইভ", employees_label: "জন কর্মী", of_total: "মোট এর মধ্যে",
    my_attendance: "আমার হাজিরা", check_in: "চেক-ইন করুন", checked_in_at: "চেক-ইন সময়",
    take_photo: "📸 ছবি তুলুন", retake: "🔄 আবার তুলুন", use_photo: "✅ ব্যবহার করুন",
    capturing_location: "লোকেশন সংগ্রহ হচ্ছে...", location_captured: "লোকেশন সংগৃহীত",
    view_location: "📍 লোকেশন দেখুন", work_update: "কাজের আপডেট",
    next_update_in: "পরবর্তী আপডেট", give_update: "➕ আপডেট দিন", update_note_placeholder: "এই মুহূর্তে কী কাজ করছেন লিখুন...",
    todays_updates: "আজকের আপডেটসমূহ", already_checked_in: "✅ আজ চেক-ইন সম্পন্ন হয়েছে",
    camera_permission_note: "ক্যামেরা ও লোকেশন পারমিশন দিন", cancel: "বাতিল",
    live_work_updates: "লাইভ কাজের আপডেট", no_updates_yet: "এখনো কোনো আপডেট নেই",
    uploading: "আপলোড হচ্ছে...", link_employee: "কর্মীর সাথে লিংক করুন", select_employee: "কর্মী বাছাই করুন",
    permissions_label: "কোন কোন Menu Access পাবে (টিক দিন)", role_employee: "👤 Employee",
    leave: "ছুটি ব্যবস্থাপনা", payroll: "পে-রোল", recruitment: "নিয়োগ", hr_system: "HR ও পে-রোল সিস্টেম",
    hr_reports: "HR রিপোর্ট", disbursement: "বিতরণ চ্যানেল", today: "আজ", this_week: "এই সপ্তাহ", this_month: "এই মাস", custom: "কাস্টম", documents: "রশিদ ও চালান",
  },
  en: {
    dashboard: "Dashboard", projects: "Projects", construction: "Construction Projects", interior: "Interior Projects",
    boq: "Estimate Project", clients: "Clients", employees: "Personnel (HR)", attendance: "Attendance",
    smart_attendance: "Smart Attendance", finance: "Financial", site: "Site Progress", materials: "Materials",
    analytics: "Reports & Analytics", users: "User Management", password: "Password",
    good_morning: "Good Morning", have_productive_day: "Have a productive day!",
    todays_attendance: "Today's Attendance", live_status: "Live Status",
    present: "Present", absent: "Absent", late: "Late", on_leave: "On Leave", total: "Total",
    department_wise_summary: "Department Wise Summary", recent_checkins: "Recent Check-ins",
    quick_checkin: "Quick Check-in", not_marked: "Not marked", mark: "Mark",
    refresh: "Refresh", search_employee: "Search employee...", no_records: "No records yet",
    live: "Live", employees_label: "employees", of_total: "of total",
    my_attendance: "My Attendance", check_in: "Check In", checked_in_at: "Checked in at",
    take_photo: "📸 Take Photo", retake: "🔄 Retake", use_photo: "✅ Use Photo",
    capturing_location: "Capturing location...", location_captured: "Location captured",
    view_location: "📍 View Location", work_update: "Work Update",
    next_update_in: "Next update in", give_update: "➕ Give Update", update_note_placeholder: "What are you working on right now...",
    todays_updates: "Today's Updates", already_checked_in: "✅ Checked in for today",
    camera_permission_note: "Allow camera & location permission", cancel: "Cancel",
    live_work_updates: "Live Work Updates", no_updates_yet: "No updates yet",
    uploading: "Uploading...", link_employee: "Link to Employee", select_employee: "Select employee",
    permissions_label: "Menu Access (check to allow)", role_employee: "👤 Employee",
    leave: "Leave Management", payroll: "Payroll", recruitment: "Recruitment", hr_system: "HR & Payroll System",
    hr_reports: "HR Reports", disbursement: "Disbursement Channel", today: "Today", this_week: "This Week", this_month: "This Month", custom: "Custom", documents: "Receipt & Invoice",
  },
};

// ============================================================
// AUTH
// ============================================================
const ADMIN_USER = "admin@noksha.com";
let ADMIN_PASS = localStorage.getItem("nic_password") || "noksha2024";

// ============================================================
// PRINT
// ============================================================
const NOKSHA_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAABHEAAAPyCAYAAAD7a0O5AAAACXBIWXMAABcRAAAXEQHKJvM/AAAgAElEQVR4nOzdTW6cR57g4Sij9tQNxAZ6Lx5ggGTDyN0AYu9mkxALPIA5y/bGrA23VqHXhmjkelDyOkGYiTlAyScY6QbSCTx4VUFXSiKpTOb7Ef+I5wEI2+WC/SpCJjN/GR9/+v333xMA05gt5sf5X/wkpXS08RCH+WtXNzv8/9+klN7v8Qt/u16u3vqts73ZYv75PE/puJDnmNrb/LUv/z0AAIMTcQAGtvHG/XgjzsyMOzTt3Rbx6H2OrY/5+6ISAFRIxAHo2WwxP8zB5vbrqTEGCvHhs/CzGYJuVyW9Xy9XD8UjAGAiIg5AD2aLebfS5jRHm2fGFKjAbfC5jTs3VvgAwLREHIBHytukTvOXcAO0ZJ0Dz8cvK3cAYBwiDsCO8napi5TSSUrpwPgB/LFqp1utc7NernY5ZB0A2JKIA7ClvPKmizffGTOAr1qLOgDQLxEHYAuzxfw8BxwrbwB29+E26KSUXjtXBwAeR8QBeEDeOnXlSnCAXv3WxZwcdJynAwBbEnEA7jFbzE9ywLH6BmA473LQuRJ0AOBhIg7AHWaLebd16gdjAzCq33I8t+UKAO4g4gB8ZraYd28gXhgXgEn9kmPOlWkAgH8ScQA2CDgAxfmQV+e8tDoHgNaJOACZgANQvO7a8gtXlgPQKhEH4F9XiP9oLABCeJdjjq1WADRFxAGaN1vMj1NKv7Y+DgABdVutXuatVu9NIAC1E3GAps0W8ycppbeuEQcITcwBoAkiDtC02WL+OqX0vPVxAKiEmANA1UQcoFm2UQFU60MOORemGICaiDhAs2aLebeN6qnfAQDVcgAyAFURcYAmuY0KoCmuJgegCiIO0KTZYv7eYcYAzflbjjnOywEgpG9MG9Ca2WJ+KuAANOm77kbC/HMAAMKxEgdojrNwAMhbrE7Xy9VbgwFAFFbiAE3JN1IJOADMUkpv8hlpABCCiAO0xhJ6AG51W2t/nC3mr2eL+ROjAkDpbKcCmpFfoL91Hg4Ad/iQUjpxgxUAJbMSB2jJiYADwD26nw+/zhbzCwMEQKlEHKAlx2YbgK/4YbaY39heBUCJRBygJSdmG4At3B56fGSwACiJiAM0Ib8Qt5UKgG11Nxl2K3J8AABAMUQcoBW2UgGwqy7+/322mLvZEIAiiDhAK0QcAB7rlQOPASiBiAO0QsQBYB/dgcdXRhCAKYk4QPVmi/mh83AA6MELIQeAKYk4QAsOzTIAPRFyAJiMiAO0wFYqAPok5AAwCREHaIGVOAD0TcgBYHQiDtACEQeAIXwMObPF/InRBWAMIg7QgiOzDMBAXqSUboQcAMYg4gAtcDMVAEN6JuQAMAYRB6jabDG3CgeAMQg5AAxOxAFq58U0AGMRcgAYlIgD1M4LaQDGJOQAMBgRB6id7VQAjE3IAWAQIg4AAPRPyAGgdyIOAAAMows5bxyyD0BfRBygdl44AzClp3lFjp9HAOxNxAFqZxk7AFM7EHIA6IOIAwAAwxNyANibiAMAAOMQcgDYi4gDAADjEXIAeDQRBwAAxiXkAPAoIg4AAIxPyAFgZyIOAABM4zbknBh/ALYh4gAAwHS6kPP32WJ+ag4A+BoRBwAApvdKyAHga0QcAAAog5ADwINEHAAAKIeQA8C9RBwAACiLkAPAnUQcAAAoj5ADwBdEHAAAKJOQA8AnRBwAACiXkAPAH0QcAAAoWxdyrswRACIOAACU74WQA4CIAwAAMQg5AI0TcQAAIA4hB6BhIg4AAMQi5AA0SsQBAIB4hByABok4AAAQk5AD0BgRBwAA4hJyABoi4gAAQGxCDkAjRBwAAIivCzlvZov5E3MJUC8RBwAA6vAspXQj5ADUS8QBAIB6CDkAFRNxAACgLkIOQKVEHAAAqI+QA1AhEQcAAOok5ABURsQBAIB6CTkAFRFxAACgbkIOQCVEHAAAqJ+QA1ABEQcAANpwG3KOzDdATCIOAAC0Q8gBCEzEAQCAthwIOQAxiTgAANAeIQcgIBEHAADaJOQABCPiAABAu4QcgEBEHAAAaJuQAxCEiAMAAAg5AAGIOAAAQBJyAMon4gAAALeEHICCiTgAAMCm25BzbFQAyiLiAAAAn+tCzq+zxfzUyACUQ8QBAADu80rIASiHiAMAADxEyAEohIgDAAB8jZADUAARBwAA2IaQAzAxEQcAANiWkAMwIREHAADYhZADMBERBwAA2JWQAzABEQcAAHgMIQdgZCIOAADwWEIOwIhEHAAAYB9dyDk3ggDDE3EAAIB9/ThbzK+MIsCwRBwAAKAPL4QcgGGJOAAAQF+EHIABiTgAAECfhByAgYg4AABA34QcgAGIOAAAwBCEHICeiTgAAMBQhByAHok4AADAkIQcgJ6IOAAAwNCEHIAeiDgAAMAYhByAPYk4AADAWIQcgD2IOAAAwJiEHIBHEnEAAICxdSHnzWwxf2LkAbYn4gAAAFN4llK6EXIAtifiAAAAUxFyAHYg4gAAAFMScgC2JOIAAABTE3IAtiDiAAAAJRByAL5CxAEAAEoh5AA8QMQBAABKIuQA3EPEAQAASiPkANxBxAEAAEok5AB8RsQBAABKJeQAbBBxAACAkgk5AJmIAwAAlE7IAZqXRBwAACCI25BzaMKAVok4AABAFF3IeTNbzI/MGNAiEQcAAIjkIK/IEXKA5og4AABANEIO0CQRBwAAiEjIAZoj4gAAAFEJOUBTRBwAACAyIQdohogDAABEJ+QATRBxgNp5MQcAbRBygOqJOEDtDswwADRDyAGqJuIAAAA1EXKAaok4AABAbYQcoEoiDgAAUCMhB6iOiAMAANRKyAGqIuIAAAA1E3KAaog4AABA7YQcoAoiDgAA0ILbkHNstoGoRBwAAKAVXcj5dbaYn5pxICIRBwAAaM0rIQeISMQBAABaJOQA4Yg4AABAq4QcIBQRBwAAaJmQA4Qh4gAAAK0TcoAQRBwAAAAhBwhAxAEAAPgnIQcomogDAADwL0IOUCwRBwAA4FNCDlAkEQcAAOBLQg5QHBEHAADgbkIOUBQRBwAA4H5CDlAMEQeo1mwxPzK7AEAPhBygCCIOULMnZhcA6ImQA0xOxAEAANiOkANMSsQBAADYnpADTEbEAQAA2I2QA0xCxAEAANidkAOMTsQBAAB4nC7knBs7YCwiDgAAwOP9OFvMr4wfMAYRBwAAYD8vhBxgDCIOAADA/oQcYHAiDgAAQD+EHGBQIg4AAEB/hBxgMCIOAABAv4QcYBAiDgAAQP+EHKB3Ig4AAMAwhBygVyIOAADAcIQcoDciDgAAwLCEHKAXIg5Qs0OzCwAUQsgB9ibiADUTcQCAkgg5wF5EHAAAgPEIOcCjiTgAAADjEnKARxFxAAAAxifkADsTcQAAAKYh5AA7EXEAAACmI+QAWxNxAAAApiXkAFsRcQAAAKYn5ABfJeIAAACUQcgBHiTiAAAAlEPIAe4l4gAAAJRFyAHuJOIAAACUR8gBviDiAAAAlEnIAT4h4gAAAJRLyAH+IOIAAACUTcgBPhJxgJodmV0AoBJCDiDiAFV7YnoBgIoIOdA4EQcAACAOIQcaJuIAAADEIuRAo0QcAACAeIQcaJCIAwAAEFMXcl7PFnPnAEIjRBwAAIC4nqeUboQcaIOIAwAAENszIQfaIOIAAADEJ+RAA0QcAACAOgg5UDkRBwAAoB5CDlRMxAEAAKiLkAOVEnEAAADqI+RAhUQcAACAOgk5UBkRB6iZFywAQOuEHKiIiAPU7JnZBQAQcqAWIg4AAED9hByogIgDAADQBiEHghNxAAAA2iHkQGAiDgAAQFuEHAhKxAEAAGiPkAMBiTgAAABtEnIgGBEHAACgXUIOBCLiAAAAtE3IgSBEHAAAAIQcCEDEAQAAIAk5UD4RBwAAgFtCDhRMxAEAAGCTkAOFEnGAKs0W80MzCwDwaEIOFEjEAWol4gAA7EfIgcKIOAAAANxHyIGCiDgAAAA8RMiBQog4AAAAfI2QAwUQcQAAANiGkAMTE3EAAADYlpADExJxAAAA2IWQAxMRcQAAANiVkAMTEHEAAAB4DCEHRibiAAAA8FhCDoxIxAEAAGAfQg6MRMQBAABgX0IOjEDEAQAAoA9CDgxMxAFqdWxmAQBGJ+TAgEQcAAAA+iTkwEBEHAAAAPom5MAARBwAAACGIORAz0QcAAAAhiLkQI9EHAAAAIYk5EBPRBwAAACGJuRAD0QcAAAAxiDkwJ5EHAAAAMYi5MAeRBwAAADGJOTAI4k4AAAAjE3IgUcQcQAAAJiCkAM7EnEAAACYipADOxBxgFodmlkAgBCEHNiSiAPUSsQBAIhDyIEtiDgAAACUQMiBrxBxAAAAKIWQAw8QcQAAACiJkAP3EHEAAAAojZADdxBxAAAAKJGQA58RcQAAACiVkAMbRBwAAABKJuRAJuIAAABQuo8hxyzROhEHAACACJ7NFvMrM0XLRBwAAACieCHk0DIRBwAAgEiEHJol4gAAABCNkEOTRBygVkdmFgCgakIOzRFxgFodmFkAgOoJOTRFxAEAACAyIYdmiDgAAABEJ+TQBBEHAACAGgg5VE/EAQAAoBZCDlUTcQAAAKiJkEO1RBwAAABqI+RQJREHAACAGgk5VEfEAQAAoFZCDlURcQAAAKiZkEM1RBwAAABqJ+RQBREHAACAFgg5hCfiAAAA0Aohh9BEHKA6s8X8yKwCAHAPIYewRBygRk/MKgAADxByCEnEAQAAoEVCDuGIOAAAALRKyCEUEQcAAICWCTmEIeIAAADQOiGHEEQcAAAAEHIIQMQBAACAfxJyKJqIAwAAAP8i5FAsEQcAAAA+JeRQJBEHAAAAviTkUBwRBwAAAO4m5FAUEQcAAADuJ+RQDBEHAAAAHtaFnAtjxNREHAAAAPi6H2aL+alxYkoiDlCjQ7MKAMAAXgk5TEnEAWok4gAAMBQhh8mIOAAAALAbIYdJiDgAAACwOyGH0Yk4AAAA8DhCDqMScQAAAODxhBxGI+IAAADAfoQcRiHiAAAAwP6EHAYn4gAAAEA/hBwGJeIAAABAf4QcBiPiAAAAQL+EHAYh4gAAAED/hBx6J+IAAADAMIQceiXiAAAAwHCEHHoj4gAAAMCwhBx6IeIAAADA8IQc9ibiAAAAwDiEHPYi4gA1OjarAAAUSsjh0f70+++/Gz2gCrPF/H+klP5nSul/pZSemlUAAAr2n+vl6rUJYhciDlCN2WLercD5Pyml/14vVxdT/bpmi/mTlNLRRP/6KVYhHeavsc0m+HcCAPTlQ/fabb1cvTGibEvEASa3R/T4PB50f/7vKaX/u16u/svMMoQcC6cyZSD8XEnP8pDuGQ/KfTwAGifksBMRByrwiAhylN+AbeMxb9QmWyGxXq7+NNW/G2BKs8V8m+/td31P//x/6/76mckEGI2Qw9ZEHNjwiE/Yd4khacetLj49fgQRB6Bfn/1svP3zzfBjayPA/oQcthI+4swW817PYlgvVzd9/bMi2fLTu8fY95+7z3J9LyobJOIATGPjtcTta7OjjZ/jPpQA+Dohh68qMuJsvAj4/I/JCwHgISIOQJnyip7b13dHOfTYtgXwKSGHB00ecfIP9GM/zIE+iDgAseQP7w434k739dQ0Ag0TcrjX6BEnR5uTHG4EG6BXIg5AfBsH9h8LO0CjhBzuNErEmS3mJzncnNgKBQxJxAGoUw47xxtxx9l3QO2EHL4wWMTJBw6f5i+fnACjEHEA2pG3Yh1vfPmwEKiNkMMneo84Od5cpJReGGpgbCIOQLs+izrP/VYAKiHk8IfeIk5e4trFm+8MLzAVEQeAW85iBCoi5PBRLxFntph3W6ZeWsIKTE3EAeAuebX4bdSxSgeISMhhv4iTV99c+UEIlELEAeBrNg5JdvEGEI2Q07hHR5y85/i1Q4uBkog4AOzKTapAMEJOwx4VcWyfAkol4gCwj42g45IOoGRCTqN2jjg54LxqfeCAMok4APQhb7nqYk732ndmUIEC/ZZDznuT046dIo6AA5ROxAGgb/lQ5NP85SgBoCRCTmO2jjj5isZfWx8woGwiDgBDytutTl3sARREyGnIVhEnf/rwxhk4QOlEHADGYHUOUBghpxHbRpwu4DxrfbCA8ok4AIwtHzlw7vUyMDEhpwHffO2XOFvML/xAAgCAu62Xq6v1cnWUUvqPlNIvhgmYSPe+/SYfzE6lHlyJM1vMux9G/zD5QBRW4gAwtbzV6sI15cBErMip2NdW4rxsfYAAAGAX6+Xq7Xq56rZY/VtK6a8ppQ8GEBiRFTkVu3cljuvEgYisxAGgNPmN1Hn+clEIMBYrcir00Eqci9YHBwAA9tW9gVovV91r60Mrc4ARWZFToTsjTl6F46pEAADoyUbM6c6d/Nm4AiMQcipz30qc89YHBgAAhvDZmTluswKGJuRU5IuIk2+kcqU4AAAMKMeck3w1+dpYAwMScipx10ocq3AAAGAk6+XqZr1cHaeU/pJSemfcgYEIORW4K+KctD4oAAAwtvVydZXPy/mrwQcG0oWclwY3rk8izmwxP3HtIQAATGPj8ON/s8UKGMiL2WJ+ZXBj+nwljlU4AAAwsXxeTrfF6j9dSQ4MQMgJ6vOIc9z6gAAAQCnWy9XrlNKhK8mBAQg5Af0RcfKtVE9bHxAAAChJ3mJ1mm+xcvAx0CchJ5jNlThW4QAAQKG6W6zywcd/M0dAj4ScQEQcAAAIIq/KOc+rcpyVA/RFyAliM+IctT4YAAAQQV6V052V84sJA3oi5ATwMeLMFvMnzsMBAIA48qqc7nbZv1iVA/REyCnc7Uocq3AAACCg9XJ1lY9G+M38AT0Qcgom4gAAQHDr5epNDjmuIgf6IOQU6jbiHLY+EAAAENnGVeS2VwF9EHIKZCUOAABUZGN71TvzCuypCznnBrEcVuIAAEBl8vaq7oPatbkF9vTjbDE/NYhluI04bqYCAICK5O1V3Yqcv5lXYE+vhJwyfJOvFwcAACq0Xq7O8zk5APsQcgrwjfNwAACgbvmcnP9w4DGwJyFnYt8U/XQAAEAv1svVjQOPgR4IORPqIo7tVAAA0ICNA49/M9/AHoScidhOBQAADekOPM4rcoQcYB9CzgRspwIAgMbkm6u6D3N/NvfAHoSckYk4AADQqPVydSrkAHsSckYk4gAAQMOEHKAHQs5IRBwAAGickAP0oAs5xwZyWA42BgAAhBygD69ni7nGMCBXjAMAAB8JOcCeDlJKN0LOcGynAgAA/iDkAHsScgYk4gAAAJ8QcoA9CTkDEXEAAIAvCDnAnoScAYg4AADAfc5TSr8ZHeCRhJyeiTgAAMCd1svV+5TSsZAD7EHI6ZGIAwAA3Gsj5HwwSsAjCTk9EXEAAIAHCTlAD4ScHog4AADAV62XqzcppRMjBexByNlTF3HehP4VAAAAo1gvVzcppb8YbWAPXci5mi3mTwzi7rqI8z7aQwMAANNYL1dXrh4H9vQsr8gRcnZkOxUAALCT9XJ16sYqYE9CziOIOAAAwGM46BjYl5CzIxEHAADY2caNVQD7EHJ2IOIAAACPkm+s+t9GD9iTkLMlEQcAAHi09XL1MqX0ixEE9iTkbEHEAQAA9tUddPzOKAJ7EnK+QsQBAAD2ks/HOTGKQA+EnAeIOAAAwN6cjwP0qAs5rw3ol0QcAACgF/l8nLXRBHowmy3mVwbyUyIOAADQp+58nA9GFOjBCyHnUyIOAADQm/Vy9TaHHIA+CDkbRBwAAKBX6+XqtWvHgR4JOZmIAwAADMG2KqBPzYecJOIAAABDyNeO21YF9Kn5kCPiAAAAg7CtChhA0yFHxAEAAIZ0blsV0LNmQ46IAwAADCbfVnVhhIGedSHnZWuDKuIAAACDWi9X3Rut34wy0LPvZot5U2dviTgAAMAYzo0yMIBXLYUcEQcAABjcerm6SSn9bKSBATQTckQcAABgLA45BobSRMgRcQAAgFGsl6v3KaXmDiIFRlN9yBFxAACA0ayXq+6mqndGHBhI1SFHxAEAAMbmynFgSNWGHBEHAAAY1Xq5uur+YNSBAVUZckQcAABgClbjAEPrQs5xTaMs4gAAAKPLV45bjQMM7fVsMT+qZZRFHAAAYCpW4wBDO0gp3dQSckQcAABgElbjACOpJuSIOAAAwJSqvQoYKEoVIUfEAQAAJrNert6mlH42A8AIwoccEQcAAJjalRkARhI65HQR520BzwEAADTK2TjAyMKGHBEHAAAogZuqgDHdhpzDSKNuOxUAADA5q3GACXQh5/VsMX8SZfBFHAAAoBQvzQQwsmd5RU6IkCPiAAAARVgvV69TSu/MBjCyMCFHxAEAAEribBxgCiFCjogDAACUpFuN88GMABMoPuSIOAAAQDHWy9X7lNKVGQEmUnTIEXEAAIDSOOAYmFKxIUfEAQAAirJert6mlH4xK8CEnuXtnUURcQAAgBLZUgVMbTZbzIv6XiTiAAAAxXHdOFCIFyWFHBEHAAAoldU4QAmKCTkiDgAAUCoRByhFESFHxAEAAIqUDzhemx2gEF3IOZ/yUUQcAACgZFbjACX5cbaYn071PCIOAABQsu6A4w9mCCjIy9lifjTF44g4AABAsdbL1fsccgBKcdB9X5ot5k/Gfh4RBwAAKJ2IA5Tm6RTfm0QcAACgaOvlqnuj9M4sAYWZzRbzizEfScQBAAAisBoHKNEPs8X8eKznEnEAAIAI3FIFlGq083FEHAAAoHjr5eqNLVVAoQ7GCs0iDgAAEIUtVUCpns8W8/Ohn03EAQAAohBxgJJdzBbzwyGfT8QBAABCWC9XNymlD2YLKNTg26pEHAAAIBKrcYCSddeOnwz1fCIOAAAQiYgDlO5qqNuqRBwAACCSG7MFFK7bVnUxxCOKOAAAQBjr5ep9SukXMwYU7rvZYn7U9yOKOAAAQDRW4wARvOz7GUUcAAAgGhEHiKD3Q45FHAAAIJT1cvUmpfTOrAEB9LoaR8QBAAAishoHiODpbDE/7+s5RRwAACAiEQeI4qKvK8dFHAAAIKLXZg0IortyvJfVOCIOAAAQTr5q/DczBwRx3sdqHBEHAACIypYqIIpeVuOIOAAAQFQiDhDJ3qtxRBwAACAqEQeIZO/VOCIOAAAQUj4X553ZAwLZO+K8N9sAAEBQVuMAkRzMFvPTxz7vN+vl6o3pBgAAghJxgGguHvu8tlMBAACR+VAaiObpbDE/fswzizgAAEBYeWfBBzMIBPOo1TgiDgAAEJ3VOEA0s9lifrjrM4s4AABAdM7FASLa+aYqEQcAAIjOShwgop1vqRJxAACA6EQcIKKdrxsXcQAAgNDWy9VbhxsDQZ3s8tgiDgAAUAOrcYCInu9ywLGIAwAA1EDEAaLaejWOiAMAANRAxAGi2vqWKhEHAACowVuzCAT1dLaYH23z6CIOAAAQ3nq5ujGLQGBb3VIl4gAAALV4ZyaBoLY6F0fEAQAAamFLFRDVVluqRBwAAKAWtlQBkX11S5WIAwAA1MJKHCCy4689u4gDAADUQsQBIns2W8wPH3p+EQcAAKjFGzMJBPfgAcciDgAAUIX1cvXeTALBPbilSsQBAABqsjabQGDPH3p0EQcAAKiJ1ThAaLPF/N7VOCIOAABQE+fiANGJOAAAQBOsxAGiE3EAAIAmWIkDRDe77/lFHAAAoCZW4gDh3XcujogDAABUY71cWYkD1EDEAQAAAAhAxAEAAJqwNs1AcEd3Pb6IAwAAAFCWg9lifvj5E4k4AABAbZyLA9Tgi9U4Ig4AAFAbN1QBNRBxAACA6ok4QA2+ONxYxAEAAGpjOxVQA2fiAAAAAATw9PNHFHEAAIDa2E4FVGG2mH+ypUrEAapy1zV8AEBb1suV7VRALT55fyPiALURcQAAgFqIOAAAAAABfHLNuIgDAADUaG1WgQo82fwliDgAAAAAZbISBwAAACCAg81HFHEAAAAACjVbzP/YUiXiAAAANboxq0Al/thSJeIAAAAABCDiAAAAAJTLdioAAACAAGynAgAAAIhExAEAAGrkYGOgOiIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAB/NkkAwBAur88OU0qH9/yj337/7U9vSxv4y+uz7grPJ/f87Tfff/vT+5EfCQDgjyvGbyPOOqU0a35YAIAHbYSZ268nGy8sur9+uuUI/jWldFHgaL986DXR5fXZ5l+uN/789hacNyml94IPANCjPz5gshIHAPjC5fXZ8UaoOc4vHp4ZqU9sxp4vwk8OPuvbqHP7VeIKJAAgBhEHABqXg83RxpdY05/buPP89p+4EXduw86NsAMAbEPEAYDG5Ghz+2U79TRmm2N/eX32IW/JuslR501j4wEAbEHEAYDK5cN6T0Sboh3k1TofV+xsRJ3XVurAo/nvBqiOiAMAFbq8PjvJ4eYkBwJi+Tzq/JajzpVVOrCd9XL1draYGy2gKiIOAFTg8vrsyUa0eW5Oq/Msf313eX32Lq/QEXQAoDEiDgAEZsVNk7pr3L8TdACgPSIOAARzeX3WXft9nlI6FW6atxl0ui1XVznovG99YACgRiIOAARxeX12msONw4m5S7fd6sfu6/L67Occc26MFADUQ8QBgILlVTeneeWNVTds60X3lVfnvPz+25+ujBwAxCfiAECB8rXg5/nNODxWtzrn1eX12csu5uSgY6sVAAQl4gBAQS6vz45TShe2TNGzbhXXD10YvLw+6w5Cvvj+25/eGmQAiEXEAYACiDeM5GBjq9XPYg4AhHB4+5AiDgBMSLxhQmIOAMTw9PYpRRwAmIB4Q0FuY87fcsxxZg4AFOobEwMA4+lum7q8PutuCvpVwKEw36WU3l5en12YGAAok5U4ADCCy+uzJ/m2qR+MNwX7eADy5fXZx2vtv//2p9cmCwDKYSUOAAzs8vrsJKX0RsAhkG7v/d8vr89u8nX3AEABrMQBgIF0W6dSSle2TRFY92SmuPEAABZnSURBVHv3H5fXZ39NKb10Xg4ATMtKHAAYQD5X5P8JOFSiW0X2Jh/IDQBMxEocAOhR3nrSrb55ZlypTLfF6tfL67NfUkqnVuUAwPisxAGAnuTVN/8QcKjc83yL1YmJBoBxWYkDAHuy+oYGHeSDj3/Ot1hZlQMAI7ASBwD2cHl9dm71DQ174awcABiPiAMAj3B5ffaku345pfSj8aNxt2flXLQ+EAAwNBEHAHaUVx28dfMUfOKHLmx2gdOwAMAwRBwA2EFebfBrPhME+NQsH3psexUADEDEAYAt5O1Tr7vVBsYLHnSQt1edGyYA6JeIAwBfkW+fuslXKwPb+fHy+uzK9ioA2N9sMf/481TEAYAHXF6fneSA4/Yp2F13e5VzcgBgf92HiiIOANzn8vrsNKX0d+ffwF6e5XNyjgwjAOxHxAGAO3TbQFJKr4wN9OIgr8g5MZwA8HgiDgBsyAcYX+VtIEB/upDz97zCDQB4BBEHALJ8bseNgAODepWv6gcAdiTiAMCnAccBxjC8H/KKNwBgByIOAM0TcGASL4QcANiNiANA0/KNOW8FHJiEkAMAOxBxAGhWDjg3rhCHSQk5ALAlEQeAJgk4UBQhBwC2IOIA0BwBB4ok5ADAV4g4ADQlH2J8JeBAkYQcAHiAiANAM9xCBSEIOQBwDxEHgCYIOBBKF3IuTBkAfErEAaAVVwIOhPLD5fXZqSkDgH8RcQCoXt6a8dxMQzivLq/PTkwbAPyTiANA1S6vz867rRlmGcK6yjfKAUDzRBwAqpW3YvxohiG07ia5m3yuFQA0TcQBoEr5k/uXZheqIOQA0Lwk4gBQo42bqA5MMFTjmTALQMMOk4gDQKUEHKjTi3zOFQC0RsQBoD6X12cvXSUOVfvx8vrs2BQD0CIRB4Bq5KuIvzOjUL3XzscBoEUiDgBVuLw+65aYXplNaEK3XfK1qQagNSIOALV47RwcaMrs8vrswpQD0BIRB4DwnIMDzfrB+TgAtETEASC0/AbOOTjQrivn4wDQChEHgLDyGzfn4EDbnqaUXrY+CAC0QcQBILKr/AYOaNuLfDsdAFRNxAEgpPyG7bnZAzLbqgConogDQDi2UQF3OPB9AYDaiTgARPTSdeLAHZ67rQqAmok4AISS36C9MGvAPWyrAqBaIg4A0dguATykO+z83AgBUCMRB4AwLq/PLtxGBWzhh8vrsyMDBUBtRBwAQri8Pjv06Tqwg5cGC4DaiDgAROEwY2AXs8vrsxMjBkBNRBwAipcPM35upoAdWY0DQFVEHAAiuDBLwCM8zWdpAUAVRBwAinZ5fXbabYswS8AjnbtyHIBaiDgAlM6n6MA+DnwfAaAWIg4AxcqrcFwpDuzru3zDHQCEJuIAUKS8/cGhpEBfrMYBIDwRB4BSnbtSHOjRC6txAIhOxAGgOHkVzrmZAXpmNQ4AUX38IELEAaBEJ1bhAAOwGgeAqEQcAIrl03JgKL6/ABCWiANAUdxIBQzsJG/ZBIBwRBwASuNTcmBIB87cAiAqEQeAYlxenx1bhQOMQMQBICQRB4CSWIUDjOEgb90EgFBEHACKkG+MmZkNYCRW4wAQjogDQCm8oQLG9Cxv4QSAMEQcAEphawMwNt93AAhFxAFgcvlsigMzAYzshevGAYhExAGgBD4NB6bi+w8AYYg4AEzKgcbAxEQcAMIQcQCYmjdQwJS6A46PzAAAEYg4AExNxAGm5vsQACGIOABMJn/6/dQMABM7MQEARCDiADAln34DJXhqSxUAEYg4AEzJp99AKURlAIon4gAwCVupgMKIygAUT8QBYCo+9QZKYksVAMUTcQCYik+9gdKIywAUTcQBYHSX12eHtlIBBTo2KQCUTMQBYApW4QAlepYjMwAUScQBYAo+7QZK5fsTAMUScQCYwnOjDhTKSkEAiiXiADCqy+szn3IDJfM9CoBiiTgAjM0bJKBkB64aB6BUIg4AYxNxgNL5PgVAkUQcAMY2M+JA4UQcAIok4gAwGufhAEH4XgVAkUQcAMbkjREQQXcuzqGZAqA0Ig4AY3JYKBCF6AxAcUQcAMbkTREQhegMQHFEHABGkbcmHBhtIAgRB4DiiDgAjMX5EkAkbtIDoDgiDgBjsZUKCOXy+sxqHACKIuIAMBZvhoBorCAEoCgiDgBj8WYIiEZ8BqAotxHnrWkBYGDPDDAQjIgDQFFEHAAGl2+mAojmiRkDoBAfP1iwnQqAMYg4QERuqIrvXesDAFTjIIk4AIzEzVRASJfXZ1bjxGbHAVAVEQcAAO7nXBwAiiHiADAGK3GAqKzEAaAYIg4AANzPShwAiiHiADAGb4IAAGBPIg4AYzgwykBQtoMCUAwRB4BBudkFAAD6IeIAMDRbqYDIhGgAiiHiAADA/Z4ZGwBKIeIAAAAABCDiADA026kAAKAHIg4AQ3OeBBCaA9oBKIWIAwAAD7OiEIAiiDgAAAAAAYg4AAAAAAGIOAAAAAABiDgAAAAAAYg4AAzNrS4AANCDPxtEAEawNsh85m2hA/KmgGegPO/NCQAlEHEAGNT33/50boSJwu9XAKBktlMBAAAABCDiAAAAAAQg4gAAAAAEIOIAAAAABCDiAAAAAAQg4gAAAAAEIOIAAAAABCDiAAAAAAQg4gAAAAAEIOIAAAAABCDiAAAAAAQg4gAAAAAEIOIAAAAABCDiAAAAAAQg4gAAAAAEIOIAAAAABCDiAAAAAAQg4gAAAAAEIOIAAAAABCDiAAAAAAQg4gAAAAAEIOIAAAAABCDiAAAAAAQg4gAAAAAEIOIAAAAABCDiAAAAAAQg4gC1OTajAABAjUQcAAAAgABEHAAAAIAARBwAAACAAEQcAAAAgABEHAAAAIAARBwAAACAAEQcAAAAgABEHAAAAIAARBwAAACAAG4jzrHJAgAAACjXn80NAEO6vD47TSkdGmQ+c/P9tz/dlDYofr9yj6vvv/3prcEBYGoiDgBD694Uz4wydygu4vj9yj2636siDgCTcyYOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOAAAAQAAiDgAAAEAAIg4AAABAACIOUJtDMwoAANRIxAFqI+IAAABVEnEAAAAAAhBxAAAAAAIQcQAAAAACEHEAAAAAAhBxAAAAAAIQcQAAAAACEHEAAAAAAhBxAAAAAAIQcQAAAAACEHH4/+3dzW0bSQKG4V5h78MMVhmYGVDAgud1AsRoM/Bc5+TNQBvAAhKYgH3mhcxAykDKwIrAg7artdQPRVIkzf5KzwMIMxdppC6yp+tldTUAAAAQQMQBAAAACCDiAAAAAAQQcQAAAAACiDgAAAAAAUQcAAAAgAAiDgAAAEAAEQcAAAAggIgDAAAAEEDEAQAAAAgg4gAAAAAEEHEAAAAAAog4AAAAAP1234g4AAAAAL133Yg4AAAAABlEHAAAAIAAIg4AAABAABEHAAAAIEAXcQYGCwAAAKC/uojzwRgBAAAA9JfbqQAAAAACiDgAAAAAAUQcAAAAgAAiDgAAAEAAEQcAAAAggIgDAAAAEEDEAQAAAAgg4gAAAAAEEHEAAAAAAog4AAAAAAFEHAAAAIAAfzdIABzYZdM0cweZJ/r6mvB65SW3jgoAfSDiAHBQf/7zf5eOMCm8XgGAPnM7FQAAAEAAEQcAAAAggIgDAAAAEEDEAQAAAAgg4gAAAAAEEHEAAAAAAog4AAAAAAFEHAAAAIAAIg4AAABAABEHAAAAIICIAwAAABBAxAEAAAAIIOIAAAAABBBxAAAAAAKIOAAAQK0GRhaoiYgDAADU6oORBWoi4gAAAAAEEHEAAAAAAog4AAAAAAFEHAAAAIAAIg4AAABAABEHAAAAIICIAwAAABBAxAEAAAAIIOIAAAAABBBxAAAAAAKIOAAAAAABRBwAAACAACIOUJuhEQUAAGok4gC1+c2IAgAANRJxAAAAAAKIOAAAAAABRBwAAACAACIOAAAAQAARBwAAACCAiAMAAAAQQMQBAAAACCDiAAAAAAQQcQAAAAACiDgAAAAAAUQcAAAAgAAiDgAAAEAAEQcAAAAggIgDAAAAEEDEAQAAAAgg4gAAAAAEEHEAAAAAAog4AAAAAAFEHAAAAIAAIg4AAABAgJPRZDw0UAAAAAD91q7EGRgjAAAAgH5zOxUAAABAABEHAAAAIICIAwAAABBAxAEAAAAIIOIAAAAABBBxAAAAAAKIOAAAAAABRBwAAACAACIOAAAAQAARBwAAACCAiAMAAAAQQMQBAAAACCDiAAAAAAQQcQAAAAD6bdCIOAAAAAC996ERcQAAAAAyiDgAAAAAAUQcAAAAgAAiDgAAAEAAEQcAAAAggIgDAAAAEEDEAQAAqjOajE+NKlAbEQcAAKiRiANUR8QBAAAACCDiAAAAAAQQcQAAAAACiDgAAAAAAUQcAAAAgAAiDgAAAEAAEQcAAAAggIgDAAAAEEDEAQAAAAgg4gDVGU3GQ6MKAADURsQBajQwqgAAQG1EHAAAAIAAIg4AAABAABEHAAAAIICIAwAAABBAxAEAAAAIIOIAAAAABBBxAAAAAAKIOAAAAAABRBwAAACAACIOAAAAQAARBwAAACCAiAMAAAAQQMQBAAAACCDiAAAAAAQQcQAAAAACtBFnYKAAAAAA+q2NOENjBAAAANBvbqcCAAAACCDiAAAAAAQQcQAAAAACiDgAAAAAAUQcAAAAgAAiDgAAAEAAEQcAAAAggIgDAAAAEEDEAQAAajQwqkBtRBwAAKBGQ6MK1EbEAQAAAAgg4gAAAAAEEHEAAAAAAog4AAAAAAFEHAAAAIAAIg4AAABAABEHAAAAIICIAwAAANBzo8l4IOIAAAAA9N9QxAEAAAAIIOIAAAAABBBxAAAAAAKIOAAAAAABRBwAAACAACIOAAAAQAARBwAAACCAiAMAAAAQQMQBanRqVAEAgNqIOECNRBwAAKA6Ig4AAABAABEHAAAAIICIAwAAABBAxAEAAAAIIOIAAAAABBBxAAAAAAKIOAAAAAABRBwAAACAACIOAAAAQAARBwAAACCAiAMAANTozKgCtRFxAAAAAAKIOAAAAAABRBwAAACAAG3EGRooAAAAgH5rI87AGAEAAAD0m9upAAAAAAKIOAAAAAABRBwAAKBGc6MK1KaNONdGFQAAAKDf2ojzzRgBAAAA9JvbqQAAAAACiDhAjW6NKgAAUBsRB6iRiAMAAFRlMZ3NRRwAAKBGnk4FVOfEJ9YAAAAA/SfiADXy1D0AAKAmd409cYAaLaazawMLAO+e6wGgJj8W4FiJAwAAVGcxnVmZC9TkxzntZDGdiThATW6MJgBQ3DsQQCV+rC7sbqdycgNq4VM3AKDjliqgFj9X4pQ/xskNqIXVhQBAx3UBUItHK3F8cg3UwsUaANBxXQDU4mFj48ZKHKAizmcAQMd1AVCD+24/YxEHqI1P3ACAjusCoAYPzUbEAaqymM6czwCAH1wXAJV4HHHKshxPqALSLYwgAPCE6wMg3bz7/U+W/hCVGkjnPAYAPOX6AEj37HaqZrnsAIRyHgMAnnJ9ACS76zY1bkQcoDLOYwDAU64PgGRfln/3h4izmM7m9sUBgt0sprNvBhAAWFauD24cFCDUoxB98uRvUKmBVF+MHACwwqUDAwS6X0xnL6/EKUyCgFTOXwDAKj6sBhI9m+OIOEAN2s2+PHkCAHhRuU64c3SAMM9WET6KOOV+0SujCoS5MGAAwBquF4Akd2Xv4keersRp3C8KBHLeAgDWcb0AJHkxPD+LOKX02L0dSHHlqVQAwDruOgCC3K8Kzy+txGksNQSCfDZYAMCGXDcACS5WfVD9YsRZTGeXNv4CArSrcG4NFACwiXLdYDUO0Gf3ry2sWbUSp/XJsAI959M0AGBbrh+APvv02nYRKyPOYjprHze+MLRAT/3HKhwAYFvl+uG/DhzQQ4tyZ9RKr63EaZ2XpTwAfXJn7y4AYAefbR8B9Mx9aTCvejXilEptuSHQNx89kQoAeKtyHbF2sgTwC33a5E6DdStx2hPchc2/gB75YzGdXRsQAGAXi+ls3t6e7SACPXC17jaqztqIU7SbHN8YWeDIrkpYBgDY2WI6a+86+OpIAkfU7oOz8crAjSJOWW54JuQAR3S1zckNAGBD5+Y5wJG0556P2/ynN12JI+QAxyTgAAAHYZ4DHEl7zjnbdq/Pv33//n2rX3c0GQ+apmnvH/1gpIFfQMABAA6uzHO+tP/qaAMH9uY5ztYRp/n/Ca7dl+J3Iwsc0L833eALAGAfRpPxpXkOcEB/7LLP55siTmc0GZ+XmPObEQb2qF1aeO4pVADAMYwm43aPikvzHGCP9jLH2SniND9PcKflBGfZIbCr+zYMlydFAAAcTbn7oJ3n/MsoADvY6xxn54jTGU3GZ+Uk94+9/EDgvblqmubzYjq7NfIAQF+Uec6FPUGBN9j7HGdvEadTbrE6tzIH2MB9ib8X4g0A0GfmOcCGDjrH2XvE6Ywm42HTNJ/KM8/dSwos+1qe/vBl20fqAQAck3kOsMLXMr856INZDhZxlpWNwc7Kic7tVvD+tDV6LtwAADUp85yPgg68S0eZ4/ySiLOsbITcBp1h+bIc8TDaF5Qn+xyO1+1q3Wvv4ctTpgCA2pUVOmdLcx0fXv+02PH7b8sXz53t6ZgMRciN3CzNcebHmuP88ojzkhJ2TsuLZ7D0z2bPk+X2oG9ax769IYK85XseLKaz+Vu/l0zlqQfDN/zy235f9x7b5Oe+tGnfS++d5f+hXnevf6tsAAB+Kpsinz75ag7wgeDdmtBxvWYetOk85psP5+iUcDnY4oBsGp02/bn7fh91wXH5/TDv2+u+FxEHAADgvSqxZ51bD4KA7awITbnvpaZp/gJ24d4M8Gx/6gAAAABJRU5ErkJggg==";

const printSection = async (title, contentId, customDate) => {
  const content = document.getElementById(contentId);
  if (!content) return;
  // Fetch logo as base64
  let logoB64 = "";
  try {
    const resp = await fetch("https://jijxnycopycsysugppnw.supabase.co/storage/v1/object/public/Upload%20images/icon.png");
    const blob = await resp.blob();
    logoB64 = await new Promise(res => { const r = new FileReader(); r.onload = e => res(e.target.result); r.readAsDataURL(blob); });
  } catch(e) { logoB64 = ""; }

  const win = window.open("", "_blank");
  win.document.write(`<!DOCTYPE html><html><head><title>${title}</title>
  <style>
    @page { size: A4; margin: 0; }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; font-size: 10pt; color: #222; }
    html, body { height: 100%; margin: 0; padding: 0; }
    .page { width: 210mm; min-height: 297mm; padding: 12mm 14mm 26mm 14mm; box-sizing: border-box; position: relative; }
    /* HEADER - exact Noksha Pad layout */
    .pad-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      padding-bottom: 6px;
      border-bottom: 1.5px solid #3F5F45;
      margin-bottom: 10px;
    }
    .pad-left {
      font-size: 9pt;
      color: #333;
      line-height: 1.7;
    }
    .pad-right {
      text-align: right;
      display: flex;
      flex-direction: column;
      align-items: flex-end;
    }
    .pad-logo-img {
      height: 55px;
      width: auto;
      object-fit: contain;
    }
    /* Document title bar */
    .doc-title {
      background: #3F5F45;
      color: white;
      text-align: center;
      padding: 4px 10px;
      font-size: 10pt;
      font-weight: bold;
      margin-bottom: 6px;
    }
    .doc-date {
      display: flex;
      justify-content: space-between;
      font-size: 8.5pt;
      color: #555;
      margin-bottom: 8px;
    }
    /* Content */
    .pad-content { }
    table { width: 100%; border-collapse: collapse; font-size: 10pt !important; }
    th { background: #3F5F45 !important; color: white !important; padding: 4px 6px !important; text-align: left; border: 0.5px solid #2A3F2E; font-size: 9.5pt !important; font-weight: 600 !important; white-space: nowrap; }
    td { padding: 4px 6px !important; border-bottom: 0.5px solid #E0E0E0; font-size: 10pt !important; }
    tr:nth-child(even) td { background: #F5F8F5; }
    /* Columns/buttons that should never appear on paper (Action, Edit/Delete, on-screen-only badges) */
    .no-print { display: none !important; }
    button { display: none !important; }
    /* FOOTER - fixed at bottom of every printed page */
    .pad-footer-fixed {
      display: none;
    }
    @media print {
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .screen-footer { display: none; }
      .pad-footer-fixed {
        display: block;
        position: fixed;
        bottom: 6mm;
        left: 14mm;
        right: 14mm;
        padding: 5px 0 0;
        border-top: 1px solid #aaa;
        font-size: 7.5pt;
        color: #555;
        text-align: center;
        background: white;
      }
    }
    .screen-footer {
      margin-top: 30px;
      border-top: 1px solid #aaa;
      padding: 5px 0 8px;
      font-size: 7.5pt;
      color: #555;
      text-align: center;
    }
  </style></head><body>
  <div class="page">
    <div class="pad-header">
      <div class="pad-left">
        Address: Arju Super Market (3rd Floor)<br/>
        Niltuli Mujib Sarak, Faridpur.<br/>
        Cell: +88 01619-677070 &nbsp; E-mail: noksha.ltd@gmail.com
      </div>
      <div class="pad-right">
        <img src="${logoB64}" style="height:55px;width:auto;object-fit:contain" alt="NOKSHA" />
      </div>
    </div>
    <div class="doc-title">${title}</div>
    <div class="doc-date">
      ${customDate ? `<span></span><span>Date: ${new Date(customDate).toLocaleDateString("en-GB")}</span>` : `<span>তারিখ: ${new Date().toLocaleDateString("bn-BD")}</span><span>Date: ${new Date().toLocaleDateString("en-GB")}</span>`}
    </div>
    <div class="pad-content">
      ${content.innerHTML}
    </div>
    <div class="screen-footer">
      (We provide all kind of Building Design, 3D View, Exterior/Interior 3D Visualization, Structural Design, Electrical Design, Plumbing Design, Pouroshova/Rajuk Sheet, Estimating &amp; Costing, Building Construction &amp; Supervision)
    </div>
  </div>
  <div class="pad-footer-fixed">
    (We provide all kind of Building Design, 3D View, Exterior/Interior 3D Visualization, Structural Design, Electrical Design, Plumbing Design, Pouroshova/Rajuk Sheet, Estimating &amp; Costing, Building Construction &amp; Supervision)
  </div>
  <script>setTimeout(() => { window.print(); }, 800 );<\/script>
  </body></html>`);
  win.document.close();
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
  XLSX.writeFile(wb, "NIC_" + fileName + "_" + new Date().toISOString().split("T")[0] + ".xlsx");
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
// IMAGE UPLOAD TO SUPABASE STORAGE
// ============================================================
const uploadImage = async (file, folder) => {
  if (!file) return null;
  if (file.size > 20 * 1024 * 1024) { alert("❌ ছবি সর্বোচ্চ 20MB হতে হবে!"); return null; }
  const ext = file.name.split(".").pop().toLowerCase();
  const fileName = folder + "/" + Date.now() + "_" + Math.random().toString(36).substr(2, 6) + "." + ext;
  const { data, error } = await supabase.storage.from("Upload images").upload(fileName, file, { cacheControl: "3600", upsert: true, contentType: file.type });
  if (error) { alert("Image upload error: " + error.message); return null; }
  const { data: urlData } = supabase.storage.from("Upload images").getPublicUrl(fileName);
  return urlData.publicUrl;
};

// ============================================================
// LOCATION + LIVE CAMERA CAPTURE UPLOAD (for attendance / work updates)
// ============================================================
const getLocation = () => new Promise((resolve) => {
  if (!navigator.geolocation) return resolve(null);
  navigator.geolocation.getCurrentPosition(
    (pos) => resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
    () => resolve(null),
    { enableHighAccuracy: true, timeout: 10000 }
  );
});

const uploadCapturedPhoto = async (dataUrl, folder) => {
  try {
    const res = await fetch(dataUrl);
    const blob = await res.blob();
    const fileName = folder + "/" + Date.now() + "_" + Math.random().toString(36).substr(2, 6) + ".jpg";
    const { error } = await supabase.storage.from("Upload images").upload(fileName, blob, { cacheControl: "3600", upsert: true, contentType: "image/jpeg" });
    if (error) { console.error(error); return null; }
    const { data: urlData } = supabase.storage.from("Upload images").getPublicUrl(fileName);
    return urlData.publicUrl;
  } catch (e) { console.error(e); return null; }
};

function CameraCapture({ lang, onCapture, onCancel }) {
  const T = TXT[lang];
  const videoRef = useRef();
  const streamRef = useRef();
  const [snap, setSnap] = useState(null);
  const [err, setErr] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "user" }, audio: false });
        streamRef.current = stream;
        if (videoRef.current) videoRef.current.srcObject = stream;
      } catch (e) { setErr(T.camera_permission_note); }
    })();
    return () => { streamRef.current?.getTracks().forEach(t => t.stop()); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const capture = () => {
    const video = videoRef.current;
    const canvas = document.createElement("canvas");
    canvas.width = video.videoWidth; canvas.height = video.videoHeight;
    canvas.getContext("2d").drawImage(video, 0, 0);
    setSnap(canvas.toDataURL("image/jpeg", 0.85));
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.75)", zIndex: 1200, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 16, padding: 20, width: "100%", maxWidth: 420 }}>
        {err ? (
          <div style={{ color: C.red, fontSize: 13, textAlign: "center", padding: 20 }}>{err}</div>
        ) : snap ? (
          <img src={snap} alt="capture" style={{ width: "100%", borderRadius: 12, marginBottom: 14, transform: "scaleX(-1)" }} />
        ) : (
          <video ref={videoRef} autoPlay playsInline muted style={{ width: "100%", borderRadius: 12, marginBottom: 14, transform: "scaleX(-1)", background: "#000" }} />
        )}
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={onCancel} style={{ ...btnEdit, flex: 1, textAlign: "center" }}>{T.cancel}</button>
          {snap ? (
            <>
              <button onClick={() => setSnap(null)} style={{ ...btnEdit, flex: 1, textAlign: "center" }}>{T.retake}</button>
              <button onClick={() => onCapture(snap)} style={{ ...btnPrimary, flex: 1, marginTop: 0 }}>{T.use_photo}</button>
            </>
          ) : (
            !err && <button onClick={capture} style={{ ...btnPrimary, flex: 1, marginTop: 0 }}>{T.take_photo}</button>
          )}
        </div>
      </div>
    </div>
  );
}

function ImageUploadField({ label, value, onChange, folder }) {
  const [uploading, setUploading] = useState(false);
  const handleFile = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setUploading(true);
    const url = await uploadImage(file, folder);
    if (url) onChange(url);
    setUploading(false);
  };
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 13, fontWeight: 600, color: C.gray800, marginBottom: 5 }}>{label}</label>
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <label style={{ background: C.primaryBg, color: C.primaryDark, border: "1px solid " + C.primaryLight, borderRadius: 8, padding: "8px 14px", cursor: "pointer", fontWeight: 600, fontSize: 12, fontFamily: "inherit", flexShrink: 0 }}>
          {uploading ? "⏳ আপলোড হচ্ছে..." : "📷 ছবি বেছে নিন"}
          <input type="file" accept="image/*" onChange={handleFile} style={{ display: "none" }} disabled={uploading} />
        </label>
        <span style={{ fontSize: 11, color: C.gray400 }}>সর্বোচ্চ 20MB</span>
      </div>
      {value && (
        <div style={{ marginTop: 8, position: "relative", display: "inline-block" }}>
          <img src={value} alt="Preview" style={{ maxWidth: 200, maxHeight: 120, borderRadius: 8, border: "1px solid " + C.gray200, objectFit: "cover" }} />
          <button onClick={() => onChange("")} style={{ position: "absolute", top: -8, right: -8, background: C.red, color: C.white, border: "none", borderRadius: "50%", width: 22, height: 22, cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", justifyContent: "center" }}>×</button>
        </div>
      )}
    </div>
  );
}

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
  return <div style={{ background: C.white, borderRadius: 12, padding: 20, boxShadow: "0 2px 12px rgba(63,95,69,0.08)", border: "1px solid " + C.gray200, ...style }}>{children}</div>;
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
  return <div style={{ background: C.gray200, borderRadius: 99, height: 8, overflow: "hidden" }}><div style={{ width: (Math.min(value, 100)) + "%", height: "100%", background: color || C.primary, borderRadius: 99, transition: "width 0.5s" }} /></div>;
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
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 24px", borderBottom: "2px solid " + C.primary, position: "sticky", top: 0, background: C.white, zIndex: 1 }}>
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

const inputStyle = { width: "100%", padding: "9px 12px", border: "1px solid " + C.gray200, borderRadius: 8, fontSize: 13, color: C.gray800, outline: "none", boxSizing: "border-box", fontFamily: "inherit", transition: "border 0.2s" };
const btnPrimary = { background: C.primary, color: C.white, border: "none", borderRadius: 10, padding: "11px", fontWeight: 700, cursor: "pointer", fontSize: 14, fontFamily: "inherit", width: "100%", marginTop: 8 };
const btnDanger = { background: C.red, color: C.white, border: "none", borderRadius: 8, padding: "6px 12px", fontWeight: 600, cursor: "pointer", fontSize: 12, fontFamily: "inherit" };
const btnEdit = { background: C.primaryBg, color: C.primaryDark, border: "1px solid " + C.primaryLight, borderRadius: 8, padding: "5px 10px", fontWeight: 600, cursor: "pointer", fontSize: 12, fontFamily: "inherit" };

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
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, " + C.primaryDark + " 0%, " + C.primary + " 60%, " + C.primaryLight + " 100%)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: C.white, borderRadius: 20, padding: 40, width: "100%", maxWidth: 400, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <img src="https://jijxnycopycsysugppnw.supabase.co/storage/v1/object/public/Upload%20images/icon.png" alt="Noksha" style={{ width: 80, height: 80, objectFit: "contain", margin: "0 auto 16px", display: "block" }} />
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
        <StatCard icon="📊" label="লাভের হার" value={profitPct + "%"} color={C.primaryBg} />
        <StatCard icon="👷" label="মাসিক বেতন" value={fmt(salaryBill)} color="#FFF8E1" />
        <StatCard icon="🏗️" label="সম্পন্ন প্রজেক্ট" value={fmtNum(completed)} sub={active + "টি চলমান"} color={C.primaryBg} />
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
              <div style={{ width: Math.round((val / maxCat) * 100) + "%", height: "100%", background: barColors[i % barColors.length], borderRadius: 99 }} />
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
    alert("✅ " + count + "টি প্রজেক্ট আপলোড হয়েছে!");
    onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <SectionHeader title="প্রজেক্ট ব্যবস্থাপনা" action="নতুন প্রজেক্ট" onAction={openAdd} onExport={handleExport} onPrint={() => { printSection("প্রজেক্ট তালিকা", "projects-content"); }} onUpload={handleUpload} uploadRef={uploadRef} />
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
            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 12, paddingTop: 12, borderTop: "1px solid " + C.gray100 }}>
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
    alert("✅ " + count + "টি ক্লায়েন্ট আপলোড হয়েছে!"); onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <SectionHeader title="ক্লায়েন্ট ব্যবস্থাপনা" action="নতুন ক্লায়েন্ট" onAction={openAdd} onExport={handleExport} onPrint={() => { printSection("ক্লায়েন্ট তালিকা", "clients-content"); }} onUpload={handleUpload} uploadRef={uploadRef} />
      <Card>
        <div id="clients-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["নাম", "ফোন", "ইমেইল", "ঠিকানা", "স্ট্যাটাস", "যোগদান", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>{h}</th>)}</tr></thead>
            <tbody>
              {data.map(c => (
                <tr key={c.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
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
// MONEY RECEIPT
// ============================================================
function MoneyReceipts({ clients }) {
  const [receipts, setReceipts] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [showMultiModal, setShowMultiModal] = useState(false);
  const [printRow, setPrintRow] = useState(null);
  const blankForm = { client_id: "", amount: "", payment_for: "", payment_method: "নগদ", received_date: new Date().toISOString().split("T")[0], received_by: "" };
  const [form, setForm] = useState(blankForm);
  const blankRow = () => ({ client_id: "", amount: "", payment_for: "", payment_method: "নগদ", received_date: new Date().toISOString().split("T")[0], received_by: "" });
  const [rows, setRows] = useState([blankRow()]);

  const load = async () => {
    setLoading(true);
    const { data, error } = await supabase.from("money_receipts").select("*").order("created_at", { ascending: false });
    if (error) { console.error(error); setLoading(false); return; }
    setReceipts(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, []);

  const clientById = {}; clients.forEach(c => { clientById[c.id] = c; });

  const save = async () => {
    if (!form.client_id || !form.amount) return alert("ক্লায়েন্ট ও পরিমাণ আবশ্যক");
    const receiptNo = "MR-" + String(receipts.length + 1).padStart(4, "0");
    const { error } = await supabase.from("money_receipts").insert([{ ...form, amount: +form.amount, received_date: form.received_date || null, receipt_no: receiptNo }]);
    if (error) return alert("❌ সংরক্ষণ ব্যর্থ: " + error.message);
    setShowModal(false); setForm(blankForm); load();
  };

  const addRow = () => setRows(r => [...r, blankRow()]);
  const removeRow = (idx) => setRows(r => r.filter((_, i) => i !== idx));
  const updateRow = (idx, field, val) => setRows(r => r.map((row, i) => i === idx ? { ...row, [field]: val } : row));

  const saveMultiRows = async () => {
    const valid = rows.filter(r => r.client_id && r.amount);
    if (valid.length === 0) return alert("কমপক্ষে একটি সারিতে ক্লায়েন্ট ও পরিমাণ দিন!");
    const payloads = valid.map((r, i) => ({ ...r, amount: +r.amount || 0, receipt_no: "MR-" + String(receipts.length + i + 1).padStart(4, "0") }));
    const { error } = await supabase.from("money_receipts").insert(payloads);
    if (error) return alert("❌ সংরক্ষণ ব্যর্থ: " + error.message);
    setShowMultiModal(false); setRows([blankRow()]); load();
  };

  const del = async (id) => { if (!confirm("এই রশিদ মুছবেন?")) return; await supabase.from("money_receipts").delete().eq("id", id); load(); };
  const doPrint = (r) => { setPrintRow(r); setTimeout(() => printSection("Money Receipt — " + r.receipt_no, "receipt-print", r.received_date), 100); };

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button onClick={() => setShowModal(true)} style={{ background: C.primary, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ একটি রশিদ</button>
        <button onClick={() => setShowMultiModal(true)} style={{ background: "#2A5C8F", color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ একসাথে একাধিক রশিদ</button>
      </div>
      <SectionHeader title="🧾 Money Receipt" />
      <Card>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : receipts.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো রশিদ তৈরি হয়নি</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["রশিদ নং", "ক্লায়েন্ট", "পরিমাণ", "বাবদ", "তারিখ", "পদ্ধতি", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {receipts.map(r => (
                <tr key={r.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 700, color: C.primaryDark }}>{r.receipt_no}</td>
                  <td style={{ padding: "10px 14px" }}>{clientById[r.client_id]?.name || "—"}</td>
                  <td style={{ padding: "10px 14px", fontWeight: 700, color: C.green }}>{fmt(r.amount)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600 }}>{r.payment_for}</td>
                  <td style={{ padding: "10px 14px" }}>{r.received_date}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={r.payment_method} color="primary" /></td>
                  <td style={{ padding: "10px 14px" }}><div style={{ display: "flex", gap: 6 }}><button onClick={() => doPrint(r)} style={btnEdit}>🖨️</button><button onClick={() => del(r.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>

      {showModal && (
        <Modal title="নতুন Money Receipt" onClose={() => setShowModal(false)}>
          <FormField label="ক্লায়েন্ট *">
            <select style={inputStyle} value={form.client_id} onChange={e => setForm({ ...form, client_id: e.target.value })}>
              <option value="">— বাছাই করুন —</option>
              {clients.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </FormField>
          <FormField label="পরিমাণ (৳) *"><input type="number" style={inputStyle} value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} /></FormField>
          <FormField label="কী বাবদ (Payment For)"><input style={inputStyle} value={form.payment_for} onChange={e => setForm({ ...form, payment_for: e.target.value })} placeholder="যেমন: এডভান্স পেমেন্ট - ইন্টেরিয়র ডিজাইন" /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="পেমেন্ট পদ্ধতি">
              <select style={inputStyle} value={form.payment_method} onChange={e => setForm({ ...form, payment_method: e.target.value })}>
                {["নগদ", "ব্যাংক", "বিকাশ", "নগদ (App)", "চেক"].map(m => <option key={m}>{m}</option>)}
              </select>
            </FormField>
            <FormField label="তারিখ"><input type="date" style={inputStyle} value={form.received_date} onChange={e => setForm({ ...form, received_date: e.target.value })} /></FormField>
          </div>
          <FormField label="গ্রহণকারী (Received By)"><input style={inputStyle} value={form.received_by} onChange={e => setForm({ ...form, received_by: e.target.value })} /></FormField>
          <button onClick={save} style={btnPrimary}>✅ তৈরি করুন</button>
        </Modal>
      )}

      {showMultiModal && (
        <Modal title="একসাথে একাধিক Money Receipt যোগ করুন" onClose={() => setShowMultiModal(false)} size={950}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: C.primaryBg }}>
                  {["তারিখ", "ক্লায়েন্ট *", "পরিমাণ *", "বাবদ", "পদ্ধতি", "গ্রহণকারী", ""].map(h => (
                    <th key={h} style={{ padding: "8px 6px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, idx) => (
                  <tr key={idx} style={{ borderBottom: "1px solid " + C.gray100 }}>
                    <td style={{ padding: 4 }}><input type="date" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 120 }} value={row.received_date} onChange={e => updateRow(idx, "received_date", e.target.value)} /></td>
                    <td style={{ padding: 4 }}>
                      <select style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 140 }} value={row.client_id} onChange={e => updateRow(idx, "client_id", e.target.value)}>
                        <option value="">— বাছাই —</option>
                        {clients.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: 4 }}><input type="number" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, width: 90 }} value={row.amount} onChange={e => updateRow(idx, "amount", e.target.value)} /></td>
                    <td style={{ padding: 4 }}><input style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 140 }} value={row.payment_for} onChange={e => updateRow(idx, "payment_for", e.target.value)} /></td>
                    <td style={{ padding: 4 }}>
                      <select style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 100 }} value={row.payment_method} onChange={e => updateRow(idx, "payment_method", e.target.value)}>
                        {["নগদ", "ব্যাংক", "বিকাশ", "নগদ (App)", "চেক"].map(m => <option key={m}>{m}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: 4 }}><input style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 100 }} value={row.received_by} onChange={e => updateRow(idx, "received_by", e.target.value)} /></td>
                    <td style={{ padding: 4 }}>{rows.length > 1 && <button onClick={() => removeRow(idx)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 14 }}>🗑️</button>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button onClick={addRow} style={{ ...btnEdit, marginTop: 10 }}>➕ সারি যোগ করুন</button>
          <div style={{ marginTop: 16 }}>
            <button onClick={saveMultiRows} style={btnPrimary}>✅ সব রশিদ সংরক্ষণ করুন ({rows.filter(r => r.client_id && r.amount).length})</button>
          </div>
        </Modal>
      )}

      <div style={{ display: "none" }}>
        <div id="receipt-print">
          {printRow && (
            <div>
              <table style={{ marginBottom: 16 }}>
                <tbody>
                  <tr><td style={{ fontWeight: 700, width: "25%" }}>Receipt No:</td><td colSpan={3}>{printRow.receipt_no}</td></tr>
                </tbody>
              </table>
              <div style={{ fontSize: "11pt", lineHeight: 2, marginBottom: 20 }}>
                <div>Received with thanks from <strong>{clientById[printRow.client_id]?.name}</strong>{clientById[printRow.client_id]?.address ? (", " + clientById[printRow.client_id].address) : ""}</div>
                <div>the sum of Taka <strong>{fmt(printRow.amount)}</strong> ({numToWordsTaka(printRow.amount)})</div>
                <div>being payment for <strong>{printRow.payment_for || "—"}</strong>.</div>
                <div>Payment Method: <strong>{printRow.payment_method}</strong></div>
              </div>
              <table>
                <tbody><tr><td style={{ fontWeight: 700 }}>মোট (Total)</td><td style={{ fontWeight: 700, color: "#2e7d32" }}>৳ {fmt(printRow.amount)}</td></tr></tbody>
              </table>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 60 }}>
                <div style={{ textAlign: "center" }}><div style={{ borderTop: "1px solid #333", width: 150, paddingTop: 4, fontSize: "9pt" }}>Client Signature</div></div>
                <div style={{ textAlign: "center" }}><div style={{ borderTop: "1px solid #333", width: 150, paddingTop: 4, fontSize: "9pt" }}>{printRow.received_by || "Authorized Signature"}</div></div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ============================================================
// INVOICE
// ============================================================
function Invoices({ clients }) {
  const [invoices, setInvoices] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [showMultiModal, setShowMultiModal] = useState(false);
  const [printRow, setPrintRow] = useState(null);
  const blankForm = { client_id: "", project_name: "", invoice_date: new Date().toISOString().split("T")[0], due_date: "", items: [{ description: "", quantity: 1, rate: 0 }], discount: 0, notes: "", status: "Unpaid" };
  const [form, setForm] = useState(blankForm);
  const blankRow = () => ({ client_id: "", project_name: "", description: "", quantity: 1, rate: "", invoice_date: new Date().toISOString().split("T")[0], due_date: "" });
  const [rows, setRows] = useState([blankRow()]);

  const load = async () => {
    setLoading(true);
    const { data, error } = await supabase.from("invoices").select("*").order("created_at", { ascending: false });
    if (error) { console.error(error); setLoading(false); return; }
    setInvoices(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, []);

  const clientById = {}; clients.forEach(c => { clientById[c.id] = c; });
  const subtotalOf = (items) => (items || []).reduce((s, it) => s + ((+it.quantity || 0) * (+it.rate || 0)), 0);
  const totalOf = (inv) => subtotalOf(inv.items) - (+inv.discount || 0);

  const updateItem = (i, field, val) => { const next = form.items.slice(); next[i] = { ...next[i], [field]: val }; setForm({ ...form, items: next }); };
  const addItem = () => setForm({ ...form, items: [...form.items, { description: "", quantity: 1, rate: 0 }] });
  const removeItem = (i) => setForm({ ...form, items: form.items.filter((_, idx) => idx !== i) });

  const save = async () => {
    if (!form.client_id) return alert("ক্লায়েন্ট আবশ্যক");
    const invoiceNo = "INV-" + String(invoices.length + 1).padStart(4, "0");
    const total = subtotalOf(form.items) - (+form.discount || 0);
    const { error } = await supabase.from("invoices").insert([{ ...form, due_date: form.due_date || null, invoice_date: form.invoice_date || null, discount: +form.discount || 0, total, invoice_no: invoiceNo }]);
    if (error) return alert("❌ সংরক্ষণ ব্যর্থ: " + error.message);
    setShowModal(false); setForm(blankForm); load();
  };

  const addRow = () => setRows(r => [...r, blankRow()]);
  const removeRow = (idx) => setRows(r => r.filter((_, i) => i !== idx));
  const updateRow = (idx, field, val) => setRows(r => r.map((row, i) => i === idx ? { ...row, [field]: val } : row));

  const saveMultiRows = async () => {
    const valid = rows.filter(r => r.client_id && r.description && r.rate);
    if (valid.length === 0) return alert("কমপক্ষে একটি সারিতে ক্লায়েন্ট, বিবরণ ও দর দিন!");
    const payloads = valid.map((r, i) => {
      const items = [{ description: r.description, quantity: +r.quantity || 1, rate: +r.rate || 0 }];
      return {
        client_id: r.client_id, project_name: r.project_name, invoice_date: r.invoice_date, due_date: r.due_date || null,
        items, discount: 0, notes: "", status: "Unpaid", total: subtotalOf(items),
        invoice_no: "INV-" + String(invoices.length + i + 1).padStart(4, "0"),
      };
    });
    const { error } = await supabase.from("invoices").insert(payloads);
    if (error) return alert("❌ সংরক্ষণ ব্যর্থ: " + error.message);
    setShowMultiModal(false); setRows([blankRow()]); load();
  };

  const markStatus = async (inv, status) => { await supabase.from("invoices").update({ status }).eq("id", inv.id); load(); };
  const del = async (id) => { if (!confirm("এই ইনভয়েস মুছবেন?")) return; await supabase.from("invoices").delete().eq("id", id); load(); };
  const doPrint = (inv) => { setPrintRow(inv); setTimeout(() => printSection("Invoice — " + inv.invoice_no, "invoice-print", inv.invoice_date), 100); };

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button onClick={() => setShowModal(true)} style={{ background: C.primary, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ একটি Invoice (বিস্তারিত)</button>
        <button onClick={() => setShowMultiModal(true)} style={{ background: "#2A5C8F", color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ একসাথে একাধিক Invoice</button>
      </div>
      <SectionHeader title="📄 Invoice" />
      <Card>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : invoices.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো Invoice তৈরি হয়নি</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["Invoice নং", "ক্লায়েন্ট", "প্রজেক্ট", "মোট", "তারিখ", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {invoices.map(inv => (
                <tr key={inv.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 700, color: C.primaryDark }}>{inv.invoice_no}</td>
                  <td style={{ padding: "10px 14px" }}>{clientById[inv.client_id]?.name || "—"}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600 }}>{inv.project_name || "—"}</td>
                  <td style={{ padding: "10px 14px", fontWeight: 700, color: C.green }}>{fmt(totalOf(inv))}</td>
                  <td style={{ padding: "10px 14px" }}>{inv.invoice_date}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={inv.status === "Paid" ? "✅ Paid" : "⏳ Unpaid"} color={inv.status === "Paid" ? "green" : "yellow"} /></td>
                  <td style={{ padding: "10px 14px" }}>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button onClick={() => doPrint(inv)} style={btnEdit}>🖨️</button>
                      {inv.status !== "Paid" && <button onClick={() => markStatus(inv, "Paid")} style={{ ...btnEdit, background: C.greenLight, color: C.green }}>✓ Paid</button>}
                      <button onClick={() => del(inv.id)} style={btnDanger}>🗑️</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>

      {showModal && (
        <Modal title="নতুন Invoice তৈরি করুন" onClose={() => setShowModal(false)} size={620}>
          <FormField label="ক্লায়েন্ট *">
            <select style={inputStyle} value={form.client_id} onChange={e => setForm({ ...form, client_id: e.target.value })}>
              <option value="">— বাছাই করুন —</option>
              {clients.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </FormField>
          <FormField label="প্রজেক্ট নাম"><input style={inputStyle} value={form.project_name} onChange={e => setForm({ ...form, project_name: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="Invoice তারিখ"><input type="date" style={inputStyle} value={form.invoice_date} onChange={e => setForm({ ...form, invoice_date: e.target.value })} /></FormField>
            <FormField label="Due তারিখ"><input type="date" style={inputStyle} value={form.due_date} onChange={e => setForm({ ...form, due_date: e.target.value })} /></FormField>
          </div>

          <div style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.primaryDark }}>আইটেম সমূহ</div>
              <button onClick={addItem} style={{ ...btnEdit, padding: "4px 10px", fontSize: 11 }}>➕ লাইন যোগ করুন</button>
            </div>
            <div style={{ border: "1px solid " + C.gray200, borderRadius: 8, overflow: "hidden" }}>
              <div style={{ display: "grid", gridTemplateColumns: "1.6fr 0.6fr 0.8fr 0.8fr 30px", gap: 0, background: C.primaryBg, padding: "6px 10px", fontSize: 11, fontWeight: 700, color: C.primaryDark }}>
                <div>বিবরণ</div><div>পরিমাণ</div><div>দর</div><div>মোট</div><div></div>
              </div>
              {form.items.map((it, i) => (
                <div key={i} style={{ display: "grid", gridTemplateColumns: "1.6fr 0.6fr 0.8fr 0.8fr 30px", gap: 6, padding: "6px 10px", borderTop: "1px solid " + C.gray100, alignItems: "center" }}>
                  <input value={it.description} onChange={e => updateItem(i, "description", e.target.value)} style={{ ...inputStyle, padding: "5px 8px", fontSize: 12 }} />
                  <input type="number" value={it.quantity} onChange={e => updateItem(i, "quantity", e.target.value)} style={{ ...inputStyle, padding: "5px 8px", fontSize: 12 }} />
                  <input type="number" value={it.rate} onChange={e => updateItem(i, "rate", e.target.value)} style={{ ...inputStyle, padding: "5px 8px", fontSize: 12 }} />
                  <div style={{ fontSize: 12, fontWeight: 600 }}>{fmt((+it.quantity || 0) * (+it.rate || 0))}</div>
                  <button onClick={() => removeItem(i)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer" }}>🗑️</button>
                </div>
              ))}
              <div style={{ padding: "8px 10px", background: C.gray50, fontSize: 12, fontWeight: 700, textAlign: "right" }}>সাবটোটাল: {fmt(subtotalOf(form.items))}</div>
            </div>
          </div>

          <FormField label="ডিসকাউন্ট (৳)"><input type="number" style={inputStyle} value={form.discount} onChange={e => setForm({ ...form, discount: e.target.value })} /></FormField>
          <FormField label="নোট"><textarea style={{ ...inputStyle, minHeight: 50 }} value={form.notes} onChange={e => setForm({ ...form, notes: e.target.value })} /></FormField>
          <div style={{ fontSize: 15, fontWeight: 700, color: C.primaryDark, marginBottom: 14, padding: "10px 14px", background: C.primaryBg, borderRadius: 8 }}>
            মোট প্রদেয়: {fmt(subtotalOf(form.items) - (+form.discount || 0))}
          </div>
          <button onClick={save} style={btnPrimary}>✅ Invoice তৈরি করুন</button>
        </Modal>
      )}

      {showMultiModal && (
        <Modal title="একসাথে একাধিক Invoice যোগ করুন (এক-লাইন প্রতিটি)" onClose={() => setShowMultiModal(false)} size={1000}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: C.primaryBg }}>
                  {["তারিখ", "ক্লায়েন্ট *", "প্রজেক্ট", "বিবরণ *", "পরিমাণ", "দর *", "মোট", "Due তারিখ", ""].map(h => (
                    <th key={h} style={{ padding: "8px 6px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, idx) => (
                  <tr key={idx} style={{ borderBottom: "1px solid " + C.gray100 }}>
                    <td style={{ padding: 4 }}><input type="date" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 120 }} value={row.invoice_date} onChange={e => updateRow(idx, "invoice_date", e.target.value)} /></td>
                    <td style={{ padding: 4 }}>
                      <select style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 140 }} value={row.client_id} onChange={e => updateRow(idx, "client_id", e.target.value)}>
                        <option value="">— বাছাই —</option>
                        {clients.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: 4 }}><input style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 120 }} value={row.project_name} onChange={e => updateRow(idx, "project_name", e.target.value)} /></td>
                    <td style={{ padding: 4 }}><input style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 160 }} value={row.description} onChange={e => updateRow(idx, "description", e.target.value)} /></td>
                    <td style={{ padding: 4 }}><input type="number" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, width: 70 }} value={row.quantity} onChange={e => updateRow(idx, "quantity", e.target.value)} /></td>
                    <td style={{ padding: 4 }}><input type="number" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, width: 90 }} value={row.rate} onChange={e => updateRow(idx, "rate", e.target.value)} /></td>
                    <td style={{ padding: 4, fontSize: 12, fontWeight: 600 }}>{fmt((+row.quantity || 0) * (+row.rate || 0))}</td>
                    <td style={{ padding: 4 }}><input type="date" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 120 }} value={row.due_date} onChange={e => updateRow(idx, "due_date", e.target.value)} /></td>
                    <td style={{ padding: 4 }}>{rows.length > 1 && <button onClick={() => removeRow(idx)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 14 }}>🗑️</button>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button onClick={addRow} style={{ ...btnEdit, marginTop: 10 }}>➕ সারি যোগ করুন</button>
          <div style={{ marginTop: 16 }}>
            <button onClick={saveMultiRows} style={btnPrimary}>✅ সব Invoice সংরক্ষণ করুন ({rows.filter(r => r.client_id && r.description && r.rate).length})</button>
          </div>
        </Modal>
      )}

      <div style={{ display: "none" }}>
        <div id="invoice-print">
          {printRow && (
            <div>
              <table style={{ marginBottom: 14 }}>
                <tbody>
                  <tr><td style={{ fontWeight: 700, width: "25%" }}>Invoice No:</td><td>{printRow.invoice_no}</td><td style={{ fontWeight: 700 }}>Due Date:</td><td>{printRow.due_date || "—"}</td></tr>
                  <tr><td style={{ fontWeight: 700 }}>Bill To:</td><td colSpan={3}>{clientById[printRow.client_id]?.name}</td></tr>
                  {printRow.project_name && <tr><td style={{ fontWeight: 700 }}>Project:</td><td colSpan={3}>{printRow.project_name}</td></tr>}
                </tbody>
              </table>
              <table style={{ marginBottom: 12 }}>
                <thead><tr><th>#</th><th>বিবরণ</th><th>পরিমাণ</th><th>দর (৳)</th><th>মোট (৳)</th></tr></thead>
                <tbody>
                  {(printRow.items || []).map((it, i) => (
                    <tr key={i}><td>{i + 1}</td><td>{it.description}</td><td>{it.quantity}</td><td>{fmt(it.rate)}</td><td>{fmt((+it.quantity || 0) * (+it.rate || 0))}</td></tr>
                  ))}
                  <tr><td colSpan={4} style={{ textAlign: "right", fontWeight: 700 }}>সাবটোটাল</td><td style={{ fontWeight: 700 }}>{fmt(subtotalOf(printRow.items))}</td></tr>
                  {printRow.discount > 0 && <tr><td colSpan={4} style={{ textAlign: "right", fontWeight: 700 }}>ডিসকাউন্ট</td><td style={{ fontWeight: 700, color: "#c0392b" }}>-{fmt(printRow.discount)}</td></tr>}
                  <tr><td colSpan={4} style={{ textAlign: "right", fontWeight: 700, fontSize: "11pt" }}>মোট প্রদেয়</td><td style={{ fontWeight: 700, fontSize: "11pt", color: "#2e7d32" }}>{fmt(totalOf(printRow))}</td></tr>
                </tbody>
              </table>
              {printRow.notes && <div style={{ fontSize: "9pt", color: "#555", marginBottom: 20 }}><strong>নোট:</strong> {printRow.notes}</div>}
              <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 50 }}>
                <div style={{ textAlign: "center" }}><div style={{ borderTop: "1px solid #333", width: 160, paddingTop: 4, fontSize: "9pt" }}>Authorized Signature</div></div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ============================================================
// DOCUMENTS HUB (Money Receipt + Invoice)
// ============================================================

function DocumentsHub({ clients }) {
  const [sub, setSub] = useState("receipt");
  return (
    <div>
      <div style={{ display: "flex", gap: 6, marginBottom: 20 }}>
        {[["receipt", "🧾 Money Receipt"], ["invoice", "📄 Invoice"]].map(([id, label]) => (
          <button key={id} onClick={() => setSub(id)} style={{ background: sub === id ? C.primary : C.white, color: sub === id ? C.white : C.gray800, border: "1px solid " + (sub === id ? C.primary : C.gray200), borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
        ))}
      </div>
      {sub === "receipt" && <MoneyReceipts clients={clients} />}
      {sub === "invoice" && <Invoices clients={clients} />}
    </div>
  );
}

// ============================================================
// EMPLOYEES
// ============================================================
function Employees({ data, onRefresh }) {
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [viewItem, setViewItem] = useState(null);
  const blankForm = { name: "", role: "", dept: "ডিজাইন", phone: "", salary: "", join_date: "", status: "কর্মরত", email: "", dob: "", gender: "পুরুষ", blood_group: "", nid: "", address: "", emergency_contact: "", bank_name: "", bank_account: "", photo_url: "" };
  const [form, setForm] = useState(blankForm);
  const uploadRef = useRef();

  const openAdd = () => { setEditItem(null); setForm(blankForm); setShowModal(true); };
  const openEdit = (item) => { setEditItem(item); setForm({ ...blankForm, ...item, salary: item.salary || "" }); setShowModal(true); };

  const save = async () => {
    if (!form.name || !form.role) return alert("নাম ও পদবি আবশ্যক");
    const payload = { ...form, salary: +form.salary || 0 };
    if (editItem) {
      const { error } = await supabase.from("employees").update(payload).eq("id", editItem.id);
      if (error) return alert("❌ সংরক্ষণ ব্যর্থ: " + error.message);
    } else {
      const maxOrder = data.reduce((m, e) => Math.max(m, e.sort_order || 0), 0);
      const { error } = await supabase.from("employees").insert([{ ...payload, sort_order: maxOrder + 1 }]);
      if (error) return alert("❌ যোগ করতে ব্যর্থ: " + error.message);
    }
    onRefresh(); setShowModal(false);
  };

  const deleteItem = async (id) => {
    if (!confirm("এই কর্মী মুছে ফেলবেন?")) return;
    await supabase.from("employees").delete().eq("id", id); onRefresh();
  };

  const moveEmployee = async (index, dir) => {
    const targetIndex = index + dir;
    if (targetIndex < 0 || targetIndex >= data.length) return;
    const a = data[index], b = data[targetIndex];
    const aOrder = a.sort_order ?? index, bOrder = b.sort_order ?? targetIndex;
    await Promise.all([
      supabase.from("employees").update({ sort_order: bOrder }).eq("id", a.id),
      supabase.from("employees").update({ sort_order: aOrder }).eq("id", b.id),
    ]);
    onRefresh();
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
    alert("✅ " + count + "জন কর্মী আপলোড হয়েছে!"); onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="👷" label="মোট কর্মী" value={fmtNum(data.length)} color={C.primaryBg} />
        <StatCard icon="💰" label="মাসিক বেতন" value={fmt(totalSalary)} color="#FFF8E1" />
        <StatCard icon="🏢" label="বিভাগ" value={fmtNum([...new Set(data.map(e => e.dept))].length)} color="#F0FFF4" />
      </div>
      <SectionHeader title="কর্মী ব্যবস্থাপনা (HR)" action="নতুন কর্মী" onAction={openAdd} onExport={handleExport} onPrint={() => { printSection("কর্মী তালিকা", "employees-content"); }} onUpload={handleUpload} uploadRef={uploadRef} />
      <Card>
        <div id="employees-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["", "নাম", "পদবি", "বিভাগ", "ফোন", "বেতন", "যোগদান", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>{h}</th>)}</tr></thead>
            <tbody>
              {data.map((e, i) => (
                <tr key={e.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={ev => ev.currentTarget.style.background = C.primaryBg} onMouseLeave={ev => ev.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "10px 6px" }}>
                    <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
                      <button onClick={() => moveEmployee(i, -1)} disabled={i === 0} style={{ background: "none", border: "none", cursor: i === 0 ? "default" : "pointer", color: i === 0 ? C.gray200 : C.gray600, fontSize: 12, lineHeight: 1, padding: 2 }}>▲</button>
                      <button onClick={() => moveEmployee(i, 1)} disabled={i === data.length - 1} style={{ background: "none", border: "none", cursor: i === data.length - 1 ? "default" : "pointer", color: i === data.length - 1 ? C.gray200 : C.gray600, fontSize: 12, lineHeight: 1, padding: 2 }}>▼</button>
                    </div>
                  </td>
                  <td style={{ padding: "10px 14px" }}>
                    <a onClick={() => setViewItem(e)} style={{ cursor: "pointer", color: C.primaryDark, fontWeight: 700, textDecoration: "none" }}>
                      {e.photo_url ? <img src={e.photo_url} alt="" style={{ width: 24, height: 24, borderRadius: "50%", objectFit: "cover", verticalAlign: "middle", marginRight: 8 }} /> : null}
                      {e.name}
                    </a>
                  </td>
                  <td style={{ padding: "10px 14px" }}>{e.role}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={e.dept} color="primary" /></td>
                  <td style={{ padding: "10px 14px" }}>{e.phone}</td>
                  <td style={{ padding: "10px 14px" }}><span style={{ fontWeight: 700, color: C.green }}>{fmt(e.salary)}</span></td>
                  <td style={{ padding: "10px 14px" }}>{e.join_date}</td>
                  <td style={{ padding: "10px 14px" }}><StatusBadge status={e.status} /></td>
                  <td style={{ padding: "10px 14px" }}><div style={{ display: "flex", gap: 6 }}><button onClick={() => setViewItem(e)} style={btnEdit}>👁️</button><button onClick={() => openEdit(e)} style={btnEdit}>✏️</button><button onClick={() => deleteItem(e.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "কর্মী সম্পাদনা" : "নতুন কর্মী"} onClose={() => setShowModal(false)} size={560}>
          <FormField label="ছবি"><ImageUploadField label="" value={form.photo_url} onChange={url => setForm({ ...form, photo_url: url })} folder="employees" /></FormField>
          <FormField label="নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></FormField>
          <FormField label="পদবি *"><input style={inputStyle} value={form.role} onChange={e => setForm({ ...form, role: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="বিভাগ"><select style={inputStyle} value={form.dept} onChange={e => setForm({ ...form, dept: e.target.value })}>{["ডিজাইন", "নির্মাণ", "প্রশাসন", "বিপণন"].map(d => <option key={d}>{d}</option>)}</select></FormField>
            <FormField label="স্ট্যাটাস"><select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>{["কর্মরত", "ছুটিতে", "বরখাস্ত"].map(s => <option key={s}>{s}</option>)}</select></FormField>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="ফোন"><input style={inputStyle} value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} /></FormField>
            <FormField label="ইমেইল"><input style={inputStyle} type="email" value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} /></FormField>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="বেতন (৳)"><input style={inputStyle} type="number" value={form.salary} onChange={e => setForm({ ...form, salary: e.target.value })} /></FormField>
            <FormField label="যোগদান"><input style={inputStyle} type="date" value={form.join_date} onChange={e => setForm({ ...form, join_date: e.target.value })} /></FormField>
          </div>
          <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, margin: "16px 0 8px", borderTop: "1px solid " + C.gray100, paddingTop: 14 }}>👤 ব্যক্তিগত তথ্য</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="জন্ম তারিখ"><input style={inputStyle} type="date" value={form.dob} onChange={e => setForm({ ...form, dob: e.target.value })} /></FormField>
            <FormField label="লিঙ্গ"><select style={inputStyle} value={form.gender} onChange={e => setForm({ ...form, gender: e.target.value })}>{["পুরুষ", "মহিলা", "অন্যান্য"].map(g => <option key={g}>{g}</option>)}</select></FormField>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="রক্তের গ্রুপ"><input style={inputStyle} value={form.blood_group} onChange={e => setForm({ ...form, blood_group: e.target.value })} placeholder="B+" /></FormField>
            <FormField label="NID নম্বর"><input style={inputStyle} value={form.nid} onChange={e => setForm({ ...form, nid: e.target.value })} /></FormField>
          </div>
          <FormField label="ঠিকানা"><input style={inputStyle} value={form.address} onChange={e => setForm({ ...form, address: e.target.value })} /></FormField>
          <FormField label="জরুরী যোগাযোগ নম্বর"><input style={inputStyle} value={form.emergency_contact} onChange={e => setForm({ ...form, emergency_contact: e.target.value })} /></FormField>
          <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, margin: "16px 0 8px", borderTop: "1px solid " + C.gray100, paddingTop: 14 }}>🏦 ব্যাংক তথ্য</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="ব্যাংকের নাম"><input style={inputStyle} value={form.bank_name} onChange={e => setForm({ ...form, bank_name: e.target.value })} /></FormField>
            <FormField label="অ্যাকাউন্ট নম্বর"><input style={inputStyle} value={form.bank_account} onChange={e => setForm({ ...form, bank_account: e.target.value })} /></FormField>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}

      {viewItem && <EmployeeProfileModal employee={viewItem} onClose={() => setViewItem(null)} onEdit={() => { setViewItem(null); openEdit(viewItem); }} />}
    </div>
  );
}

// ============================================================
// WEEKLY TIMELINE (check-in/check-out visual, per employee)
// ============================================================
function WeeklyTimeline({ employeeId, lang }) {
  const [weekOffset, setWeekOffset] = useState(0);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const weekdayNames = ["রবি", "সোম", "মঙ্গল", "বুধ", "বৃহঃ", "শুক্র", "শনি"];
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  const fmtDate = (d) => d.toISOString().split("T")[0];
  const getWeekStart = () => {
    const now = new Date();
    const sunday = new Date(now);
    sunday.setDate(now.getDate() - now.getDay() + weekOffset * 7);
    sunday.setHours(0, 0, 0, 0);
    return sunday;
  };
  const weekStart = getWeekStart();
  const days = Array.from({ length: 7 }, (_, i) => { const d = new Date(weekStart); d.setDate(weekStart.getDate() + i); return d; });

  const load = async () => {
    if (!employeeId) return;
    setLoading(true);
    const { data } = await supabase.from("attendance").select("*").eq("employee_id", employeeId).gte("date", fmtDate(days[0])).lte("date", fmtDate(days[6]));
    setRows(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, [weekOffset, employeeId]);

  const rowByDate = {}; rows.forEach(r => { rowByDate[r.date] = r; });
  const todayStr = new Date().toISOString().split("T")[0];

  const fmtHours = (checkIn, checkOut) => {
    if (!checkIn || !checkOut) return null;
    const diffMs = new Date(checkOut) - new Date(checkIn);
    if (diffMs <= 0) return null;
    const h = Math.floor(diffMs / 3600000), m = Math.round((diffMs % 3600000) / 60000);
    return String(h).padStart(2, "0") + ":" + String(m).padStart(2, "0");
  };
  const fmtTime = (t) => t ? new Date(t).toLocaleTimeString(lang === "bn" ? "bn-BD" : "en-GB", { hour: "2-digit", minute: "2-digit" }) : "--:--";

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <button onClick={() => setWeekOffset(w => w - 1)} style={btnEdit}>◀</button>
        <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 13 }}>
          {days[0].getDate()} {monthNames[days[0].getMonth()]} {days[0].getFullYear()} — {days[6].getDate()} {monthNames[days[6].getMonth()]} {days[6].getFullYear()}
        </div>
        <button onClick={() => setWeekOffset(w => w + 1)} disabled={weekOffset >= 0} style={{ ...btnEdit, opacity: weekOffset >= 0 ? 0.4 : 1, cursor: weekOffset >= 0 ? "default" : "pointer" }}>▶</button>
      </div>

      {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {days.map((d) => {
            const dateStr = fmtDate(d);
            const row = rowByDate[dateStr];
            const isWeekend = d.getDay() === 5; // Friday
            const isFuture = dateStr > todayStr;
            let label = null, dotColor = C.gray200, bg = "transparent";
            if (isWeekend) { label = "সাপ্তাহিক ছুটি"; dotColor = C.gray400; bg = C.gray50; }
            else if (isFuture) { label = null; dotColor = C.gray200; bg = "transparent"; }
            else if (!row) { label = "অনুপস্থিত"; dotColor = C.red; bg = "#FFF5F5"; }
            else if (row.status === "ছুটি") { label = "ছুটিতে"; dotColor = C.blue; bg = "#F0F7FF"; }
            else if (row.status === "অনুপস্থিত") { label = "অনুপস্থিত"; dotColor = C.red; bg = "#FFF5F5"; }
            else { label = "উপস্থিত" + (row.source === "manual_employee" || row.source === "manual_admin" ? " (Manual)" : ""); dotColor = C.green; bg = "#F0FFF4"; }

            const hours = row ? fmtHours(row.check_in_time, row.check_out_time) : null;
            const ongoing = row?.check_in_time && !row?.check_out_time;

            return (
              <div key={dateStr} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", borderRadius: 10, background: bg, border: "1px solid " + C.gray100, flexWrap: "wrap" }}>
                <div style={{ width: 44, textAlign: "center", flexShrink: 0 }}>
                  <div style={{ fontSize: 15, fontWeight: 700, color: C.primaryDark }}>{d.getDate()}</div>
                  <div style={{ fontSize: 10, color: C.gray400 }}>{weekdayNames[d.getDay()]}</div>
                </div>
                <div style={{ width: 60, fontSize: 11, color: C.gray600, flexShrink: 0 }}>{row?.check_in_time ? fmtTime(row.check_in_time) : "--:--"}</div>
                <div style={{ flex: 1, minWidth: 80, position: "relative", height: 2, background: (isFuture || isWeekend) ? C.gray100 : dotColor, borderRadius: 2, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <span style={{ position: "absolute", left: -3, width: 8, height: 8, borderRadius: "50%", background: dotColor }} />
                  {label && <span style={{ background: C.white, border: "1px solid " + C.gray200, borderRadius: 20, padding: "3px 10px", fontSize: 10, fontWeight: 600, color: dotColor, whiteSpace: "nowrap" }}>{label}</span>}
                  <span style={{ position: "absolute", right: -3, width: 8, height: 8, borderRadius: "50%", background: dotColor }} />
                </div>
                <div style={{ width: 60, fontSize: 11, color: C.gray600, textAlign: "right", flexShrink: 0 }}>{row?.check_out_time ? fmtTime(row.check_out_time) : "--:--"}</div>
                <div style={{ width: 70, fontSize: 11, fontWeight: 700, color: C.primaryDark, textAlign: "right", flexShrink: 0 }}>{hours ? hours + " Hrs" : ongoing ? "চলছে…" : "—"}</div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function ProfileField({ label, value }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ fontSize: 11, color: C.gray400, marginBottom: 2 }}>{label}</div>
      <div style={{ fontSize: 13, color: C.gray800, fontWeight: 600 }}>{value || "—"}</div>
    </div>
  );
}

function EmployeeProfileModal({ employee, onClose, onEdit }) {
  const [tab, setTab] = useState("overview");
  const [attSummary, setAttSummary] = useState(null);
  const [payHistory, setPayHistory] = useState([]);

  useEffect(() => {
    (async () => {
      const monthStart = new Date().toISOString().slice(0, 7) + "-01";
      const { data: att } = await supabase.from("attendance").select("status").eq("employee_id", employee.id).gte("date", monthStart);
      const s = { "উপস্থিত": 0, "অনুপস্থিত": 0, "অর্ধদিন": 0, "ছুটি": 0 };
      (att || []).forEach(a => { if (s[a.status] != null) s[a.status]++; });
      setAttSummary(s);
      const { data: pay } = await supabase.from("payroll_runs").select("*").eq("employee_id", employee.id).order("month", { ascending: false }).limit(6);
      setPayHistory(pay || []);
    })();
  }, [employee.id]);

  const tabs = [["overview", "👤 Overview"], ["employment", "💼 চাকরির তথ্য"], ["timeline", "📅 সাপ্তাহিক সময়রেখা"], ["bank", "🏦 ব্যাংক ও ডকুমেন্ট"], ["salary", "💰 বেতন ইতিহাস"]];

  return (
    <Modal title="কর্মীর প্রোফাইল" onClose={onClose} size={620}>
      <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 18, paddingBottom: 18, borderBottom: "1px solid " + C.gray100 }}>
        {employee.photo_url ? <img src={employee.photo_url} alt="" style={{ width: 64, height: 64, borderRadius: "50%", objectFit: "cover" }} /> :
          <div style={{ width: 64, height: 64, borderRadius: "50%", background: C.primaryBg, color: C.primaryDark, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 24 }}>{(employee.name || "?")[0]}</div>}
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 17, fontWeight: 700, color: C.primaryDark }}>{employee.name}</div>
          <div style={{ fontSize: 13, color: C.gray600 }}>{employee.role} · {employee.dept}</div>
          <StatusBadge status={employee.status} />
        </div>
        <button onClick={onEdit} style={btnEdit}>✏️ Edit</button>
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 18, borderBottom: "1px solid " + C.gray100, flexWrap: "wrap" }}>
        {tabs.map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{ background: "none", border: "none", borderBottom: tab === id ? "2px solid " + C.primary : "2px solid transparent", padding: "8px 12px", fontSize: 12, fontWeight: 600, color: tab === id ? C.primaryDark : C.gray600, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
        ))}
      </div>

      {tab === "overview" && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
            <ProfileField label="ফোন" value={employee.phone} />
            <ProfileField label="ইমেইল" value={employee.email} />
            <ProfileField label="জন্ম তারিখ" value={employee.dob} />
            <ProfileField label="লিঙ্গ" value={employee.gender} />
            <ProfileField label="রক্তের গ্রুপ" value={employee.blood_group} />
            <ProfileField label="জরুরী যোগাযোগ" value={employee.emergency_contact} />
          </div>
          <ProfileField label="ঠিকানা" value={employee.address} />
          {attSummary && (
            <div style={{ marginTop: 8 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 8 }}>এই মাসের উপস্থিতি সারাংশ</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>
                <StatCard icon="✅" label="উপস্থিত" value={fmtNum(attSummary["উপস্থিত"])} color={C.greenLight} />
                <StatCard icon="❌" label="অনুপস্থিত" value={fmtNum(attSummary["অনুপস্থিত"])} color={C.redLight} />
                <StatCard icon="⏰" label="অর্ধদিন" value={fmtNum(attSummary["অর্ধদিন"])} color={C.yellowLight} />
                <StatCard icon="🌴" label="ছুটি" value={fmtNum(attSummary["ছুটি"])} color={C.blueLight} />
              </div>
            </div>
          )}
        </div>
      )}

      {tab === "employment" && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
          <ProfileField label="পদবি" value={employee.role} />
          <ProfileField label="বিভাগ" value={employee.dept} />
          <ProfileField label="যোগদানের তারিখ" value={employee.join_date} />
          <ProfileField label="স্ট্যাটাস" value={employee.status} />
          <ProfileField label="মাসিক বেতন" value={fmt(employee.salary)} />
          <ProfileField label="NID নম্বর" value={employee.nid} />
        </div>
      )}

      {tab === "timeline" && <WeeklyTimeline employeeId={employee.id} lang="bn" />}

      {tab === "bank" && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
          <ProfileField label="ব্যাংকের নাম" value={employee.bank_name} />
          <ProfileField label="অ্যাকাউন্ট নম্বর" value={employee.bank_account} />
          <ProfileField label="NID নম্বর" value={employee.nid} />
        </div>
      )}

      {tab === "salary" && (
        <div>
          {payHistory.length === 0 ? (
            <div style={{ color: C.gray400, textAlign: "center", padding: 20, fontSize: 13 }}>এখনো কোনো পে-রোল রেকর্ড নেই</div>
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead><tr style={{ background: C.primaryBg }}>{["মাস", "মূল বেতন", "ভাতা", "কর্তন", "নীট প্রদেয়", "স্ট্যাটাস"].map(h => <th key={h} style={{ padding: "8px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
              <tbody>
                {payHistory.map(p => (
                  <tr key={p.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                    <td style={{ padding: "8px 10px" }}>{p.month}</td>
                    <td style={{ padding: "8px 10px" }}>{fmt(p.basic)}</td>
                    <td style={{ padding: "8px 10px" }}>{fmt(p.allowances)}</td>
                    <td style={{ padding: "8px 10px" }}>{fmt(p.deductions)}</td>
                    <td style={{ padding: "8px 10px", fontWeight: 700, color: C.green }}>{fmt(p.net_pay)}</td>
                    <td style={{ padding: "8px 10px" }}><StatusBadge status={p.status} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
    </Modal>
  );
}

// ============================================================
// LEAVE MANAGEMENT
// ============================================================
function LeaveManagement({ employees }) {
  const [requests, setRequests] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [tab, setTab] = useState("pending");
  const [form, setForm] = useState({ employee_id: "", leave_type: "casual", start_date: "", end_date: "", reason: "", status: "Approved" });

  const load = async () => {
    const { data } = await supabase.from("leave_requests").select("*").order("applied_at", { ascending: false });
    setRequests(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, []);

  const empById = {}; employees.forEach(e => { empById[e.id] = e; });
  const calcDays = (s, e) => { if (!s || !e) return 0; const d = (new Date(e) - new Date(s)) / 86400000 + 1; return d > 0 ? d : 0; };

  const decide = async (id, status) => {
    await supabase.from("leave_requests").update({ status, decided_at: new Date().toISOString() }).eq("id", id);
    load();
  };

  const submitAdminLeave = async () => {
    if (!form.employee_id || !form.start_date || !form.end_date) return alert("কর্মী ও তারিখ আবশ্যক");
    const days = calcDays(form.start_date, form.end_date);
    await supabase.from("leave_requests").insert([{ ...form, days, applied_at: new Date().toISOString(), decided_at: new Date().toISOString() }]);
    setShowModal(false); setForm({ employee_id: "", leave_type: "casual", start_date: "", end_date: "", reason: "", status: "Approved" }); load();
  };

  const thisYear = new Date().getFullYear();
  const usedByEmpType = {};
  requests.filter(r => r.status === "Approved" && new Date(r.start_date).getFullYear() === thisYear).forEach(r => {
    const key = r.employee_id + "_" + r.leave_type;
    usedByEmpType[key] = (usedByEmpType[key] || 0) + (r.days || 0);
  });

  const pending = requests.filter(r => r.status === "Pending");
  const approved = requests.filter(r => r.status === "Approved");
  const rejected = requests.filter(r => r.status === "Rejected");
  const onLeaveToday = requests.filter(r => r.status === "Approved" && r.start_date <= new Date().toISOString().split("T")[0] && r.end_date >= new Date().toISOString().split("T")[0]).length;
  const list = tab === "pending" ? pending : tab === "approved" ? approved : tab === "rejected" ? rejected : requests;

  const statusColor = { Pending: "yellow", Approved: "green", Rejected: "red" };

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="⏳" label="অপেক্ষমাণ আবেদন" value={fmtNum(pending.length)} color={C.yellowLight} />
        <StatCard icon="✅" label="এই বছর অনুমোদিত" value={fmtNum(approved.length)} color={C.greenLight} />
        <StatCard icon="🌴" label="আজ ছুটিতে আছে" value={fmtNum(onLeaveToday)} color={C.blueLight} />
      </div>

      <SectionHeader title="🌴 ছুটি ব্যবস্থাপনা" action="নতুন ছুটি এন্ট্রি" onAction={() => setShowModal(true)} />

      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {[["pending", "⏳ অপেক্ষমাণ (" + pending.length + ")"], ["approved", "✅ অনুমোদিত"], ["rejected", "❌ বাতিল"], ["all", "সব"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{ background: tab === id ? C.primary : C.white, color: tab === id ? C.white : C.gray800, border: "1px solid " + (tab === id ? C.primary : C.gray200), borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
        ))}
      </div>

      <Card>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : list.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো রেকর্ড নেই</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "ধরন", "তারিখ", "দিন", "কারণ", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {list.map(r => (
                <tr key={r.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{empById[r.employee_id]?.name || "—"}</td>
                  <td style={{ padding: "10px 14px" }}>{LEAVE_TYPES.find(t => t.id === r.leave_type)?.label || r.leave_type}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12 }}>{r.start_date} → {r.end_date}</td>
                  <td style={{ padding: "10px 14px" }}>{fmtNum(r.days)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600, maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.reason}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={r.status} color={statusColor[r.status]} /></td>
                  <td style={{ padding: "10px 14px" }}>
                    {r.status === "Pending" && (
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => decide(r.id, "Approved")} style={{ ...btnEdit, background: C.greenLight, color: C.green }}>✓</button>
                        <button onClick={() => decide(r.id, "Rejected")} style={btnDanger}>✕</button>
                      </div>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>

      <Card style={{ marginTop: 16 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>📊 ছুটির ব্যালেন্স ({thisYear})</div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ background: C.primaryBg }}>
              <th style={{ padding: "8px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>কর্মী</th>
              {LEAVE_TYPES.map(t => <th key={t.id} style={{ padding: "8px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{t.label} ({t.quota})</th>)}
            </tr></thead>
            <tbody>
              {employees.map(e => (
                <tr key={e.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "8px 10px", fontWeight: 600, color: C.primaryDark }}>{e.name}</td>
                  {LEAVE_TYPES.map(t => {
                    const used = usedByEmpType[e.id + "_" + t.id] || 0;
                    const remaining = t.quota - used;
                    return <td key={t.id} style={{ padding: "8px 10px", color: remaining <= 0 ? C.red : C.gray800 }}>{fmtNum(remaining)} / {fmtNum(t.quota)}</td>;
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      {showModal && (
        <Modal title="নতুন ছুটি এন্ট্রি (Admin)" onClose={() => setShowModal(false)}>
          <FormField label="কর্মী *">
            <select style={inputStyle} value={form.employee_id} onChange={e => setForm({ ...form, employee_id: e.target.value })}>
              <option value="">— বাছাই করুন —</option>
              {employees.map(e => <option key={e.id} value={e.id}>{e.name}</option>)}
            </select>
          </FormField>
          <FormField label="ছুটির ধরন">
            <select style={inputStyle} value={form.leave_type} onChange={e => setForm({ ...form, leave_type: e.target.value })}>
              {LEAVE_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
            </select>
          </FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="শুরুর তারিখ *"><input type="date" style={inputStyle} value={form.start_date} onChange={e => setForm({ ...form, start_date: e.target.value })} /></FormField>
            <FormField label="শেষের তারিখ *"><input type="date" style={inputStyle} value={form.end_date} onChange={e => setForm({ ...form, end_date: e.target.value })} /></FormField>
          </div>
          <FormField label="কারণ"><textarea style={{ ...inputStyle, minHeight: 60 }} value={form.reason} onChange={e => setForm({ ...form, reason: e.target.value })} /></FormField>
          <FormField label="স্ট্যাটাস"><select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}><option value="Approved">✅ Approved</option><option value="Pending">⏳ Pending</option></select></FormField>
          <button onClick={submitAdminLeave} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
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
  const uploadRef = useRef();
  const handleUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      const emp = employees.find(em => em.name === (row["নাম"] || row["name"]));
      if (!emp) continue;
      const date = row["তারিখ"] || row["date"] || selDate;
      const status = row["উপস্থিতি"] || row["status"] || "উপস্থিত";
      const existing = await supabase.from("attendance").select("id").eq("employee_id", emp.id).eq("date", date).single();
      if (existing.data) { await supabase.from("attendance").update({ status }).eq("id", existing.data.id); }
      else { await supabase.from("attendance").insert([{ employee_id: emp.id, date, status }]); }
      count++;
    }
    alert("✅ " + count + "জনের উপস্থিতি আপলোড হয়েছে!"); e.target.value = "";
  };

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
  const handlePrint = () => printSection("উপস্থিতি — " + selDate, "attendance-content");
  const handleExport = () => exportToExcel(employees.map(e => ({ নাম: e.name, পদবি: e.role, তারিখ: selDate, উপস্থিতি: attData[e.id] || "চিহ্নিত নয়" })), "Attendance", "Attendance");

  return (
    <div>
      <SectionHeader title="উপস্থিতি ব্যবস্থাপনা" onPrint={handlePrint} onExport={handleExport} onUpload={handleUpload} uploadRef={uploadRef} />
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
              <th style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>কর্মীর নাম</th>
              <th style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>পদবি</th>
              {statuses.map(s => <th key={s} style={{ padding: "10px 14px", textAlign: "center", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>{s}</th>)}
            </tr></thead>
            <tbody>
              {employees.map(e => (
                <tr key={e.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{e.name}</td>
                  <td style={{ padding: "10px 14px", color: C.gray600, fontSize: 12 }}>{e.role}</td>
                  {statuses.map(s => (
                    <td key={s} style={{ padding: "6px 14px", textAlign: "center" }}>
                      <input type="radio" name={"att-" + e.id} checked={attData[e.id] === s} onChange={() => setAtt(e.id, s)} style={{ accentColor: s === "উপস্থিত" ? C.green : s === "অনুপস্থিত" ? C.red : C.primary, width: 16, height: 16, cursor: "pointer" }} />
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
// DONUT CHART (pure SVG, no libraries)
// ============================================================
function DonutChart({ segments, size = 170, thickness = 24, centerLabel, centerSub }) {
  const total = segments.reduce((s, x) => s + x.value, 0) || 1;
  const r = (size - thickness) / 2;
  const c = 2 * Math.PI * r;
  let offset = 0;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ flexShrink: 0 }}>
      <g transform={`rotate(-90 ${size / 2} ${size / 2})`}>
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={C.gray100} strokeWidth={thickness} />
        {segments.filter(s => s.value > 0).map((seg, i) => {
          const frac = seg.value / total;
          const dash = frac * c;
          const el = (
            <circle key={i} cx={size / 2} cy={size / 2} r={r} fill="none" stroke={seg.color} strokeWidth={thickness}
              strokeDasharray={`${dash} ${c - dash}`} strokeDashoffset={-offset} strokeLinecap="butt" />
          );
          offset += dash;
          return el;
        })}
      </g>
      <text x="50%" y="46%" textAnchor="middle" fontSize="26" fontWeight="800" fill={C.primaryDark}>{centerLabel}</text>
      <text x="50%" y="61%" textAnchor="middle" fontSize="11" fill={C.gray600}>{centerSub}</text>
    </svg>
  );
}

// ============================================================
// HR SMART ATTENDANCE SYSTEM
// ============================================================
// ============================================================
// SMART ATTENDANCE SUMMARY (Weekly / Monthly / Yearly)
// ============================================================
// ============================================================
// ATTENDANCE REQUESTS (Overtime / Regularization / On Duty / Hourly Permission / Shift Change)
// ============================================================
function AttendanceRequestsPanel({ employees, type, shifts }) {
  const cfg = REQUEST_TYPES[type];
  const [requests, setRequests] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("pending");

  const load = async () => {
    setLoading(true);
    const { data } = await supabase.from("attendance_requests").select("*").eq("type", type).order("applied_at", { ascending: false });
    setRequests(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, [type]);

  const empById = {}; employees.forEach(e => { empById[String(e.id)] = e; });
  const shiftById = {}; (shifts || []).forEach(s => { shiftById[s.id] = s; });

  const decide = async (id, status) => {
    await supabase.from("attendance_requests").update({ status, decided_at: new Date().toISOString() }).eq("id", id);
    load();
  };

  const pending = requests.filter(r => r.status === "Pending");
  const approved = requests.filter(r => r.status === "Approved");
  const rejected = requests.filter(r => r.status === "Rejected");
  const list = tab === "pending" ? pending : tab === "approved" ? approved : rejected;

  return (
    <div>
      <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>{cfg.icon} {cfg.label}</div>
      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {[["pending", "⏳ অপেক্ষমাণ (" + pending.length + ")"], ["approved", "✅ অনুমোদিত"], ["rejected", "❌ বাতিল"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{ background: tab === id ? C.primary : C.white, color: tab === id ? C.white : C.gray800, border: "1px solid " + (tab === id ? C.primary : C.gray200), borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
        ))}
      </div>
      <Card>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : list.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো রেকর্ড নেই</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "তারিখ", cfg.needsShift ? "চাহিত Shift" : "সময়", "কারণ", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {list.map(r => (
                <tr key={r.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{empById[String(r.employee_id)]?.name || "—"}</td>
                  <td style={{ padding: "10px 14px" }}>{r.date}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12 }}>
                    {cfg.needsShift ? (shiftById[r.requested_shift_id]?.name || "—") : (r.start_time && r.end_time ? r.start_time + " - " + r.end_time : "—")}
                  </td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.reason}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={r.status === "Pending" ? "⏳ Pending" : r.status === "Approved" ? "✅ Approved" : "❌ Rejected"} color={r.status === "Pending" ? "yellow" : r.status === "Approved" ? "green" : "red"} /></td>
                  <td style={{ padding: "10px 14px" }}>
                    {r.status === "Pending" && (
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => decide(r.id, "Approved")} style={{ ...btnEdit, background: C.greenLight, color: C.green }}>✓</button>
                        <button onClick={() => decide(r.id, "Rejected")} style={btnDanger}>✕</button>
                      </div>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

// ============================================================
// SHIFT MANAGEMENT
// ============================================================
function ShiftManagement({ employees }) {
  const [shifts, setShifts] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", start_time: "09:00", end_time: "18:00" });

  const load = async () => { setLoading(true); const { data } = await supabase.from("shifts").select("*").order("start_time"); setShifts(data || []); setLoading(false); };
  useEffect(() => { load(); }, []);

  const save = async () => {
    if (!form.name) return alert("Shift নাম আবশ্যক");
    await supabase.from("shifts").insert([form]);
    setShowModal(false); setForm({ name: "", start_time: "09:00", end_time: "18:00" }); load();
  };

  const assignShift = async (empId, shiftId) => {
    await supabase.from("employees").update({ shift_id: shiftId || null }).eq("id", empId);
  };

  const del = async (id) => { if (!confirm("এই শিফট মুছবেন?")) return; await supabase.from("shifts").delete().eq("id", id); load(); };

  return (
    <div>
      <SectionHeader title="🔄 Shift ব্যবস্থাপনা" action="নতুন Shift" onAction={() => setShowModal(true)} />
      <Card style={{ marginBottom: 16 }}>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : shifts.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো Shift তৈরি হয়নি</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["Shift নাম", "শুরু", "শেষ", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {shifts.map(s => (
                <tr key={s.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{s.name}</td>
                  <td style={{ padding: "10px 14px" }}>{s.start_time}</td>
                  <td style={{ padding: "10px 14px" }}>{s.end_time}</td>
                  <td style={{ padding: "10px 14px" }}><button onClick={() => del(s.id)} style={btnDanger}>🗑️</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>

      <Card>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>👷 কর্মী-ভিত্তিক Shift নির্ধারণ</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "বর্তমান Shift"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
          <tbody>
            {employees.map(e => (
              <tr key={e.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{e.name}</td>
                <td style={{ padding: "10px 14px" }}>
                  <select defaultValue={e.shift_id || ""} onChange={ev => assignShift(e.id, ev.target.value || null)} style={{ ...inputStyle, width: "auto" }}>
                    <option value="">— নির্ধারিত নয় —</option>
                    {shifts.map(s => <option key={s.id} value={s.id}>{s.name} ({s.start_time}-{s.end_time})</option>)}
                  </select>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>

      {showModal && (
        <Modal title="নতুন Shift তৈরি করুন" onClose={() => setShowModal(false)}>
          <FormField label="Shift নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="Day Shift" /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="শুরুর সময়"><input type="time" style={inputStyle} value={form.start_time} onChange={e => setForm({ ...form, start_time: e.target.value })} /></FormField>
            <FormField label="শেষের সময়"><input type="time" style={inputStyle} value={form.end_time} onChange={e => setForm({ ...form, end_time: e.target.value })} /></FormField>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

function SmartAttendanceSummary({ employees, lang }) {
  const [range, setRange] = useState("week");
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);

  const getRange = () => {
    const today = new Date();
    const todayStr = today.toISOString().split("T")[0];
    if (range === "week") { const d = new Date(today); d.setDate(d.getDate() - 6); return { start: d.toISOString().split("T")[0], end: todayStr }; }
    if (range === "month") { const d = new Date(today.getFullYear(), today.getMonth(), 1); return { start: d.toISOString().split("T")[0], end: todayStr }; }
    const d = new Date(today.getFullYear(), 0, 1); return { start: d.toISOString().split("T")[0], end: todayStr };
  };

  const load = async () => {
    setLoading(true);
    const { start, end } = getRange();
    const { data } = await supabase.from("attendance").select("*").gte("date", start).lte("date", end).eq("approved", true);
    setRows(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, [range]);

  const { start, end } = getRange();
  const summary = employees.map(e => {
    const empRows = rows.filter(r => r.employee_id === e.id);
    const present = empRows.filter(r => r.status === "উপস্থিত").length;
    const absent = empRows.filter(r => r.status === "অনুপস্থিত").length;
    const late = empRows.filter(r => r.status === "অর্ধদিন").length;
    const leave = empRows.filter(r => r.status === "ছুটি").length;
    const manual = empRows.filter(r => r.source === "manual_employee" || r.source === "manual_admin").length;
    return { employee: e, present, absent, late, leave, manual };
  });

  return (
    <div>
      <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
        {[["week", "এই সপ্তাহ"], ["month", "এই মাস"], ["year", "এই বছর"]].map(([id, label]) => (
          <button key={id} onClick={() => setRange(id)} style={{ background: range === id ? C.primary : C.white, color: range === id ? C.white : C.gray800, border: "1px solid " + (range === id ? C.primary : C.gray200), borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
        ))}
      </div>
      <div style={{ fontSize: 12, color: C.gray400, marginBottom: 12 }}>{start} → {end}</div>

      <Card>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "বিভাগ", "উপস্থিত", "অনুপস্থিত", "অর্ধদিন", "ছুটি", "ম্যানুয়াল এন্ট্রি"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
              <tbody>
                {summary.map(s => (
                  <tr key={s.employee.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                    <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{s.employee.name}</td>
                    <td style={{ padding: "10px 14px" }}><Badge label={s.employee.dept} color="primary" /></td>
                    <td style={{ padding: "10px 14px", color: C.green, fontWeight: 700 }}>{fmtNum(s.present)}</td>
                    <td style={{ padding: "10px 14px", color: C.red, fontWeight: 700 }}>{fmtNum(s.absent)}</td>
                    <td style={{ padding: "10px 14px", color: "#856404", fontWeight: 700 }}>{fmtNum(s.late)}</td>
                    <td style={{ padding: "10px 14px", color: C.blue, fontWeight: 700 }}>{fmtNum(s.leave)}</td>
                    <td style={{ padding: "10px 14px", color: C.gray600 }}>{s.manual > 0 ? fmtNum(s.manual) + " দিন" : "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        {summary.some(s => s.manual > 0) && (
          <div style={{ marginTop: 12, fontSize: 11, color: C.gray400, fontStyle: "italic" }}>* ম্যানুয়াল এন্ট্রি: Attendance by Manual System</div>
        )}
      </Card>
    </div>
  );
}

function SmartAttendance({ employees, lang }) {
  const T = TXT[lang];
  const today = new Date().toISOString().split("T")[0];
  const [rows, setRows] = useState([]);
  const [workUpdates, setWorkUpdates] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [tab, setTab] = useState("today");
  const [shifts, setShifts] = useState([]);

  useEffect(() => { supabase.from("shifts").select("*").then(({ data }) => setShifts(data || [])); }, []);

  const load = async () => {
    const { data } = await supabase.from("attendance").select("*").eq("date", today).order("id", { ascending: false });
    setRows(data || []);
    const { data: wu } = await supabase.from("work_updates").select("*").eq("date", today).order("created_at", { ascending: false }).limit(20);
    setWorkUpdates(wu || []);
    setLoading(false);
  };

  useEffect(() => {
    load();
    const channel = supabase.channel("attendance-live-" + today)
      .on("postgres_changes", { event: "*", schema: "public", table: "attendance", filter: `date=eq.${today}` }, () => load())
      .on("postgres_changes", { event: "*", schema: "public", table: "work_updates", filter: `date=eq.${today}` }, () => load())
      .subscribe();
    return () => { supabase.removeChannel(channel); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const statusMap = { "উপস্থিত": "present", "অনুপস্থিত": "absent", "অর্ধদিন": "late", "ছুটি": "on_leave" };
  const reverseStatus = { present: "উপস্থিত", absent: "অনুপস্থিত", late: "অর্ধদিন", on_leave: "ছুটি" };
  const empById = {}; employees.forEach(e => { empById[e.id] = e; });

  const pendingRows = rows.filter(r => r.source === "manual_employee" && r.approved === false);
  const visibleRows = rows.filter(r => !(r.source === "manual_employee" && r.approved === false));
  const attByEmp = {}; visibleRows.forEach(r => { attByEmp[String(r.employee_id)] = r; });

  const approveManual = async (id) => { await supabase.from("attendance").update({ approved: true }).eq("id", id); load(); };
  const rejectManual = async (id) => { await supabase.from("attendance").delete().eq("id", id); load(); };

  const counts = { present: 0, absent: 0, late: 0, on_leave: 0 };
  visibleRows.forEach(r => { const k = statusMap[r.status]; if (k) counts[k]++; });
  const total = employees.length;

  const donutSegments = [
    { value: counts.present, color: C.green },
    { value: counts.absent, color: C.red },
    { value: counts.late, color: C.yellow },
    { value: counts.on_leave, color: C.blue },
  ];

  const depts = [...new Set(employees.map(e => e.dept || "—"))];
  const deptSummary = depts.map(d => {
    const deptEmps = employees.filter(e => (e.dept || "—") === d);
    const present = deptEmps.filter(e => statusMap[attByEmp[String(e.id)]?.status] === "present").length;
    return { dept: d, present, total: deptEmps.length };
  });

  const recentSorted = [...visibleRows].sort((a, b) => {
    if (a.created_at && b.created_at) return new Date(b.created_at) - new Date(a.created_at);
    return (b.id || 0) - (a.id || 0);
  });
  const recentList = recentSorted.slice(0, 8);

  const unmarked = employees.filter(e => !attByEmp[String(e.id)]).filter(e => (e.name || "").toLowerCase().includes(search.toLowerCase()));

  const quickMark = async (empId, statusKey) => {
    const status = reverseStatus[statusKey];
    setRows(prev => [...prev, { id: "temp-" + empId, employee_id: empId, date: today, status, created_at: new Date().toISOString(), source: "manual_admin", approved: true }]);
    const existing = await supabase.from("attendance").select("id").eq("employee_id", empId).eq("date", today).single();
    if (existing.data) { await supabase.from("attendance").update({ status, source: "manual_admin", approved: true }).eq("id", existing.data.id); }
    else { await supabase.from("attendance").insert([{ employee_id: empId, date: today, status, source: "manual_admin", approved: true }]); }
    load();
  };

  const statusColor = { present: C.green, absent: C.red, late: "#856404", on_leave: C.blue };
  const statusBg = { present: C.greenLight, absent: C.redLight, late: C.yellowLight, on_leave: C.blueLight };

  return (
    <div>
      <style>{`@keyframes nicPulse { 0% { box-shadow: 0 0 0 0 rgba(40,167,69,0.5); } 70% { box-shadow: 0 0 0 6px rgba(40,167,69,0); } 100% { box-shadow: 0 0 0 0 rgba(40,167,69,0); } }`}</style>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 8 }}>
        <h2 style={{ margin: 0, color: C.primaryDark, fontSize: 18, fontWeight: 700 }}>{T.smart_attendance}</h2>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, background: C.greenLight, color: C.green, padding: "5px 12px", borderRadius: 20, fontSize: 12, fontWeight: 700 }}>
            <span style={{ width: 8, height: 8, borderRadius: "50%", background: C.green, display: "inline-block", animation: "nicPulse 1.6s infinite" }} />
            {T.live_status}
          </div>
          <button onClick={load} style={{ background: C.primaryBg, border: "1px solid " + C.primaryLight, borderRadius: 8, padding: "6px 12px", cursor: "pointer", fontSize: 12, color: C.primaryDark, fontFamily: "inherit" }}>🔄 {T.refresh}</button>
        </div>
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 18, flexWrap: "wrap" }}>
        {[["today", "🗓️ আজ"], ["summary", "📅 সারাংশ (Weekly/Monthly/Yearly)"], ["overtime", "⏱️ Overtime"], ["regularization", "✏️ Regularization"], ["on_duty", "🚗 On Duty"], ["hourly_permission", "⏳ Hourly Permission"], ["shift", "🔄 Shift"], ["shift_change", "🔁 Shift Change"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{ background: tab === id ? C.primary : C.white, color: tab === id ? C.white : C.gray800, border: "1px solid " + (tab === id ? C.primary : C.gray200), borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>{label}</button>
        ))}
      </div>

      {tab === "summary" && <SmartAttendanceSummary employees={employees} lang={lang} />}
      {tab === "overtime" && <AttendanceRequestsPanel employees={employees} type="overtime" shifts={shifts} />}
      {tab === "regularization" && <AttendanceRequestsPanel employees={employees} type="regularization" shifts={shifts} />}
      {tab === "on_duty" && <AttendanceRequestsPanel employees={employees} type="on_duty" shifts={shifts} />}
      {tab === "hourly_permission" && <AttendanceRequestsPanel employees={employees} type="hourly_permission" shifts={shifts} />}
      {tab === "shift" && <ShiftManagement employees={employees} />}
      {tab === "shift_change" && <AttendanceRequestsPanel employees={employees} type="shift_change" shifts={shifts} />}

      {tab === "today" && (<>

      {pendingRows.length > 0 && (
        <Card style={{ marginBottom: 16, border: "1px solid #FFE58F", background: "#FFFBE6" }}>
          <div style={{ fontWeight: 700, color: "#856404", marginBottom: 12, fontSize: 14 }}>⏳ ম্যানুয়াল হাজিরা — অনুমোদনের অপেক্ষায় ({fmtNum(pendingRows.length)})</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {pendingRows.map(r => (
              <div key={r.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: C.white, borderRadius: 8, padding: "8px 12px", flexWrap: "wrap", gap: 8 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.primaryDark }}>{empById[r.employee_id]?.name || "—"}</div>
                <div style={{ fontSize: 11, color: C.gray600 }}>{r.check_in_time ? new Date(r.check_in_time).toLocaleTimeString(lang === "bn" ? "bn-BD" : "en-GB", { hour: "2-digit", minute: "2-digit" }) : ""}</div>
                <div style={{ display: "flex", gap: 6 }}>
                  <button onClick={() => approveManual(r.id)} style={{ ...btnEdit, background: C.greenLight, color: C.green }}>✓ Approve</button>
                  <button onClick={() => rejectManual(r.id)} style={btnDanger}>✕ Reject</button>
                </div>
              </div>
            ))}
          </div>
        </Card>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="✅" label={T.present} value={fmtNum(counts.present)} color={C.greenLight} />
        <StatCard icon="❌" label={T.absent} value={fmtNum(counts.absent)} color={C.redLight} />
        <StatCard icon="⏰" label={T.late} value={fmtNum(counts.late)} color={C.yellowLight} />
        <StatCard icon="🌴" label={T.on_leave} value={fmtNum(counts.on_leave)} color={C.blueLight} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.1fr 1fr", gap: 16, marginBottom: 16 }}>
        <Card>
          <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>{T.department_wise_summary}</div>
          <div style={{ display: "flex", gap: 20, alignItems: "center", flexWrap: "wrap" }}>
            <DonutChart segments={donutSegments} centerLabel={fmtNum(total)} centerSub={T.total} />
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {[["present", T.present, C.green], ["absent", T.absent, C.red], ["late", T.late, "#E0A800"], ["on_leave", T.on_leave, C.blue]].map(([k, label, color]) => (
                <div key={k} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                  <span style={{ width: 10, height: 10, borderRadius: "50%", background: color, display: "inline-block" }} />
                  <span style={{ color: C.gray800, minWidth: 90 }}>{label}</span>
                  <span style={{ fontWeight: 700, color: C.primaryDark }}>{fmtNum(counts[k])}</span>
                </div>
              ))}
            </div>
          </div>
          {depts.length > 0 && (
            <div style={{ marginTop: 16, borderTop: "1px solid " + C.gray100, paddingTop: 12 }}>
              {deptSummary.map(d => (
                <div key={d.dept} style={{ marginBottom: 10 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.gray600, marginBottom: 4 }}>
                    <span>{d.dept}</span><span>{fmtNum(d.present)}/{fmtNum(d.total)} {T.present}</span>
                  </div>
                  <ProgressBar value={d.total ? (d.present / d.total) * 100 : 0} color={C.primary} />
                </div>
              ))}
            </div>
          )}
        </Card>

        <Card>
          <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>{T.recent_checkins}</div>
          {loading ? (
            <div style={{ color: C.gray400, textAlign: "center", padding: 20 }}>⏳</div>
          ) : recentList.length === 0 ? (
            <div style={{ color: C.gray400, textAlign: "center", padding: 20, fontSize: 13 }}>{T.no_records}</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10, maxHeight: 360, overflowY: "auto" }}>
              {recentList.map(r => {
                const emp = empById[r.employee_id];
                const key = statusMap[r.status] || "present";
                const time = r.created_at ? new Date(r.created_at).toLocaleTimeString(lang === "bn" ? "bn-BD" : "en-GB", { hour: "2-digit", minute: "2-digit" }) : "—";
                return (
                  <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 10px", borderRadius: 10, background: C.gray50 }}>
                    {r.photo_url ? (
                      <img src={r.photo_url} alt="" style={{ width: 36, height: 36, borderRadius: "50%", objectFit: "cover", flexShrink: 0 }} />
                    ) : (
                      <div style={{ width: 36, height: 36, borderRadius: "50%", background: C.primaryBg, color: C.primaryDark, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13, flexShrink: 0 }}>
                        {(emp?.name || "?")[0]}
                      </div>
                    )}
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, color: C.primaryDark, fontSize: 13, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{emp?.name || "—"}</div>
                      <div style={{ fontSize: 11, color: C.gray600, display: "flex", gap: 6, alignItems: "center" }}>
                        <span>{time}</span>
                        {r.lat && <a href={`https://www.google.com/maps?q=${r.lat},${r.lng}`} target="_blank" rel="noreferrer" style={{ color: C.blue, textDecoration: "none" }}>📍</a>}
                      </div>
                    </div>
                    <span style={{ background: statusBg[key], color: statusColor[key], padding: "3px 10px", borderRadius: 20, fontSize: 11, fontWeight: 600, flexShrink: 0 }}>{T[key]}</span>
                  </div>
                );
              })}
            </div>
          )}
        </Card>
      </div>

      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 15 }}>{T.quick_checkin} ({fmtNum(unmarked.length)} {T.not_marked})</div>
          <input placeholder={T.search_employee} value={search} onChange={e => setSearch(e.target.value)} style={{ ...inputStyle, width: 220 }} />
        </div>
        {unmarked.length === 0 ? (
          <div style={{ color: C.green, textAlign: "center", padding: 16, fontSize: 13, fontWeight: 600 }}>✅ {fmtNum(total)} {T.employees_label} — {T.present}</div>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(240px,1fr))", gap: 10 }}>
            {unmarked.map(e => (
              <div key={e.id} style={{ border: "1px solid " + C.gray200, borderRadius: 10, padding: 10, display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 32, height: 32, borderRadius: "50%", background: C.primaryBg, color: C.primaryDark, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12, flexShrink: 0 }}>{(e.name || "?")[0]}</div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, fontSize: 12, color: C.primaryDark, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{e.name}</div>
                  <div style={{ fontSize: 10, color: C.gray600 }}>{e.dept}</div>
                </div>
                <div style={{ display: "flex", gap: 4 }}>
                  <button onClick={() => quickMark(e.id, "present")} title={T.present} style={{ background: C.greenLight, color: C.green, border: "none", borderRadius: 6, width: 26, height: 26, cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>✓</button>
                  <button onClick={() => quickMark(e.id, "late")} title={T.late} style={{ background: C.yellowLight, color: "#856404", border: "none", borderRadius: 6, width: 26, height: 26, cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>⏰</button>
                  <button onClick={() => quickMark(e.id, "absent")} title={T.absent} style={{ background: C.redLight, color: C.red, border: "none", borderRadius: 6, width: 26, height: 26, cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>✕</button>
                  <button onClick={() => quickMark(e.id, "on_leave")} title={T.on_leave} style={{ background: C.blueLight, color: C.blue, border: "none", borderRadius: 6, width: 26, height: 26, cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>🌴</button>
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card style={{ marginTop: 16 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>🔴 {T.live_work_updates}</div>
        {workUpdates.length === 0 ? (
          <div style={{ color: C.gray400, textAlign: "center", padding: 16, fontSize: 13 }}>{T.no_updates_yet}</div>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(260px,1fr))", gap: 10 }}>
            {workUpdates.map(u => {
              const emp = empById[u.employee_id];
              const time = u.created_at ? new Date(u.created_at).toLocaleTimeString(lang === "bn" ? "bn-BD" : "en-GB", { hour: "2-digit", minute: "2-digit" }) : "—";
              return (
                <div key={u.id} style={{ border: "1px solid " + C.gray200, borderRadius: 10, padding: 10, display: "flex", gap: 10, alignItems: "flex-start" }}>
                  {u.photo_url ? <img src={u.photo_url} alt="" style={{ width: 40, height: 40, borderRadius: 8, objectFit: "cover", flexShrink: 0 }} /> : <div style={{ width: 40, height: 40, borderRadius: 8, background: C.primaryBg, flexShrink: 0 }} />}
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 6 }}>
                      <span style={{ fontWeight: 600, fontSize: 12, color: C.primaryDark, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{emp?.name || "—"}</span>
                      <span style={{ fontSize: 11, color: C.gray600, flexShrink: 0 }}>{time}</span>
                    </div>
                    {u.note && <div style={{ fontSize: 11, color: C.gray800, marginTop: 3 }}>{u.note}</div>}
                    {u.lat && <a href={`https://www.google.com/maps?q=${u.lat},${u.lng}`} target="_blank" rel="noreferrer" style={{ fontSize: 11, color: C.blue }}>{T.view_location}</a>}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </Card>
      </>)}
    </div>
  );
}

// ============================================================
// MY ATTENDANCE (Employee self check-in + 15-min work updates)
// ============================================================
function MyAttendance({ currentUser, lang }) {
  const T = TXT[lang];
  const today = new Date().toISOString().split("T")[0];
  const [employee, setEmployee] = useState(null);
  const [todayAtt, setTodayAtt] = useState(null);
  const [updates, setUpdates] = useState([]);
  const [showCamera, setShowCamera] = useState(false); // false | 'checkin' | 'update'
  const [note, setNote] = useState("");
  const [busy, setBusy] = useState(false);
  const [myLeaves, setMyLeaves] = useState([]);
  const [showLeaveModal, setShowLeaveModal] = useState(false);
  const [leaveForm, setLeaveForm] = useState({ leave_type: "casual", start_date: "", end_date: "", reason: "" });
  const [myRequests, setMyRequests] = useState([]);
  const [shiftsList, setShiftsList] = useState([]);
  const [showReqModal, setShowReqModal] = useState(false);
  const [reqType, setReqType] = useState("overtime");
  const [reqForm, setReqForm] = useState({ date: today, start_time: "", end_time: "", reason: "", requested_shift_id: "" });
  const [myReviews, setMyReviews] = useState([]);
  const [selfReviewDraft, setSelfReviewDraft] = useState({});
  const [myTimeLogs, setMyTimeLogs] = useState([]);
  const [showTimeLogModal, setShowTimeLogModal] = useState(false);
  const [timeLogForm, setTimeLogForm] = useState({ task: "", project_name: "", date: today, start_time: "", end_time: "", note: "" });
  const [myTickets, setMyTickets] = useState([]);
  const [showTicketModal, setShowTicketModal] = useState(false);
  const [ticketForm, setTicketForm] = useState({ subject: "", category: TICKET_CATEGORIES[0], description: "" });

  const loadReviews = async () => {
    if (!currentUser?.employee_id) return;
    const { data } = await supabase.from("performance_reviews").select("*").eq("employee_id", currentUser.employee_id).order("created_at", { ascending: false });
    setMyReviews(data || []);
  };
  const submitSelfReview = async (id) => {
    await supabase.from("performance_reviews").update({ self_review: selfReviewDraft[id] || "", status: "Submitted" }).eq("id", id);
    loadReviews();
  };

  const loadTimeLogs = async () => {
    if (!currentUser?.employee_id) return;
    const { data } = await supabase.from("time_logs").select("*").eq("employee_id", currentUser.employee_id).order("date", { ascending: false }).limit(20);
    setMyTimeLogs(data || []);
  };
  const submitTimeLog = async () => {
    if (!timeLogForm.task || !timeLogForm.date) return alert("কাজ ও তারিখ দিন");
    await supabase.from("time_logs").insert([{ employee_id: currentUser.employee_id, ...timeLogForm }]);
    setShowTimeLogModal(false); setTimeLogForm({ task: "", project_name: "", date: today, start_time: "", end_time: "", note: "" }); loadTimeLogs();
  };

  const loadTickets = async () => {
    if (!currentUser?.employee_id) return;
    const { data } = await supabase.from("hr_tickets").select("*").eq("employee_id", currentUser.employee_id).order("created_at", { ascending: false });
    setMyTickets(data || []);
  };
  const submitTicket = async () => {
    if (!ticketForm.subject) return alert("বিষয় লিখুন");
    await supabase.from("hr_tickets").insert([{ employee_id: currentUser.employee_id, ...ticketForm, status: "Open" }]);
    setShowTicketModal(false); setTicketForm({ subject: "", category: TICKET_CATEGORIES[0], description: "" }); loadTickets();
  };

  const loadRequests = async () => {
    if (!currentUser?.employee_id) return;
    const { data } = await supabase.from("attendance_requests").select("*").eq("employee_id", currentUser.employee_id).order("applied_at", { ascending: false });
    setMyRequests(data || []);
  };

  const submitRequest = async () => {
    if (!reqForm.date) return alert("তারিখ দিন");
    if (REQUEST_TYPES[reqType].needsShift && !reqForm.requested_shift_id) return alert("Shift বাছাই করুন");
    await supabase.from("attendance_requests").insert([{ employee_id: currentUser.employee_id, type: reqType, date: reqForm.date, start_time: reqForm.start_time || null, end_time: reqForm.end_time || null, reason: reqForm.reason, requested_shift_id: reqForm.requested_shift_id || null, status: "Pending" }]);
    setShowReqModal(false); setReqForm({ date: today, start_time: "", end_time: "", reason: "", requested_shift_id: "" }); loadRequests();
  };

  const loadLeaves = async () => {
    if (!currentUser?.employee_id) return;
    const { data } = await supabase.from("leave_requests").select("*").eq("employee_id", currentUser.employee_id).order("applied_at", { ascending: false });
    setMyLeaves(data || []);
  };

  const submitLeave = async () => {
    if (!leaveForm.start_date || !leaveForm.end_date) return alert("তারিখ দিন");
    const days = Math.max((new Date(leaveForm.end_date) - new Date(leaveForm.start_date)) / 86400000 + 1, 1);
    await supabase.from("leave_requests").insert([{ employee_id: currentUser.employee_id, ...leaveForm, days, status: "Pending", applied_at: new Date().toISOString() }]);
    setShowLeaveModal(false); setLeaveForm({ leave_type: "casual", start_date: "", end_date: "", reason: "" }); loadLeaves();
  };

  const load = async () => {
    if (!currentUser?.employee_id) return;
    const { data: emp } = await supabase.from("employees").select("*").eq("id", currentUser.employee_id).single();
    setEmployee(emp || null);
    const { data: att } = await supabase.from("attendance").select("*").eq("employee_id", currentUser.employee_id).eq("date", today).single();
    setTodayAtt(att || null);
    const { data: upd } = await supabase.from("work_updates").select("*").eq("employee_id", currentUser.employee_id).eq("date", today).order("created_at", { ascending: false });
    setUpdates(upd || []);
  };

  useEffect(() => {
    load();
    loadLeaves();
    loadRequests();
    loadReviews();
    loadTimeLogs();
    loadTickets();
    supabase.from("shifts").select("*").then(({ data }) => setShiftsList(data || []));
    if (!currentUser?.employee_id) return;
    const channel = supabase.channel("my-att-" + currentUser.employee_id)
      .on("postgres_changes", { event: "*", schema: "public", table: "work_updates", filter: `employee_id=eq.${currentUser.employee_id}` }, load)
      .on("postgres_changes", { event: "*", schema: "public", table: "leave_requests", filter: `employee_id=eq.${currentUser.employee_id}` }, loadLeaves)
      .subscribe();
    return () => { supabase.removeChannel(channel); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser?.employee_id]);

  const doCheckIn = async (photoDataUrl) => {
    setBusy(true); setShowCamera(false);
    const loc = await getLocation();
    const photoUrl = await uploadCapturedPhoto(photoDataUrl, "attendance");
    await supabase.from("attendance").insert([{ employee_id: currentUser.employee_id, date: today, status: "উপস্থিত", photo_url: photoUrl, lat: loc?.lat || null, lng: loc?.lng || null, check_in_time: new Date().toISOString(), source: "live", approved: true }]);
    await load(); setBusy(false);
  };

  const doManualCheckIn = async () => {
    if (!confirm("ম্যানুয়াল হাজিরা জমা দিতে চান? এটা Admin অনুমোদন না করা পর্যন্ত চূড়ান্ত হবে না।")) return;
    setBusy(true);
    const loc = await getLocation();
    await supabase.from("attendance").insert([{ employee_id: currentUser.employee_id, date: today, status: "উপস্থিত", lat: loc?.lat || null, lng: loc?.lng || null, check_in_time: new Date().toISOString(), source: "manual_employee", approved: false }]);
    await load(); setBusy(false);
  };

  const doCheckOut = async () => {
    if (!todayAtt) return;
    setBusy(true);
    await supabase.from("attendance").update({ check_out_time: new Date().toISOString() }).eq("id", todayAtt.id);
    await load(); setBusy(false);
  };

  const doWorkUpdate = async (photoDataUrl) => {
    setBusy(true); setShowCamera(false);
    const loc = await getLocation();
    const photoUrl = photoDataUrl ? await uploadCapturedPhoto(photoDataUrl, "work-updates") : null;
    await supabase.from("work_updates").insert([{ employee_id: currentUser.employee_id, date: today, note, photo_url: photoUrl, lat: loc?.lat || null, lng: loc?.lng || null }]);
    setNote(""); await load(); setBusy(false);
  };

  if (!currentUser?.employee_id) {
    return <Card><div style={{ textAlign: "center", padding: 30, color: C.gray600, fontSize: 13 }}>⚠️ আপনার অ্যাকাউন্ট কোনো Employee রেকর্ডের সাথে লিংক করা নেই। Admin-কে User Management থেকে লিংক করতে বলুন।</div></Card>;
  }

  const isFirstUpdate = updates.length === 0;

  return (
    <div>
      <div style={{ marginBottom: 18 }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: C.primaryDark }}>{T.good_morning}, {employee?.name || "..."}</div>
        <div style={{ fontSize: 12, color: C.gray400 }}>{new Date().toLocaleDateString(lang === "bn" ? "bn-BD" : "en-GB", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}</div>
      </div>

      {!todayAtt ? (
        <Card style={{ textAlign: "center", padding: 30 }}>
          <div style={{ fontSize: 40, marginBottom: 10 }}>🙋</div>
          <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 16, fontSize: 16 }}>{T.check_in}</div>
          <button disabled={busy} onClick={() => setShowCamera("checkin")} style={{ ...btnPrimary, maxWidth: 260, margin: "0 auto" }}>{busy ? T.uploading : "📸 " + T.check_in}</button>
          <div style={{ marginTop: 10 }}>
            <button disabled={busy} onClick={doManualCheckIn} style={{ background: "none", border: "none", color: C.gray600, fontSize: 12, textDecoration: "underline", cursor: "pointer", fontFamily: "inherit" }}>ক্যামেরা কাজ করছে না? ম্যানুয়ালি হাজিরা দিন</button>
          </div>
        </Card>
      ) : todayAtt.source === "manual_employee" && !todayAtt.approved ? (
        <Card style={{ marginBottom: 16, display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap", background: C.yellowLight }}>
          <div style={{ fontSize: 30 }}>⏳</div>
          <div style={{ flex: 1 }}>
            <div style={{ color: "#856404", fontWeight: 700, fontSize: 14, marginBottom: 4 }}>ম্যানুয়াল হাজিরা জমা হয়েছে — Admin অনুমোদনের অপেক্ষায়</div>
            <div style={{ fontSize: 12, color: C.gray600 }}>{T.checked_in_at}: {new Date(todayAtt.check_in_time || todayAtt.created_at || Date.now()).toLocaleTimeString(lang === "bn" ? "bn-BD" : "en-GB", { hour: "2-digit", minute: "2-digit" })}</div>
          </div>
        </Card>
      ) : (
        <>
          <Card style={{ marginBottom: 16, display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap" }}>
            {todayAtt.photo_url && <img src={todayAtt.photo_url} alt="checkin" style={{ width: 70, height: 70, borderRadius: 12, objectFit: "cover" }} />}
            <div style={{ flex: 1 }}>
              <div style={{ color: C.green, fontWeight: 700, fontSize: 14, marginBottom: 4 }}>{T.already_checked_in}</div>
              <div style={{ fontSize: 12, color: C.gray600 }}>{T.checked_in_at}: {new Date(todayAtt.check_in_time || todayAtt.created_at || Date.now()).toLocaleTimeString(lang === "bn" ? "bn-BD" : "en-GB", { hour: "2-digit", minute: "2-digit" })}</div>
              {todayAtt.check_out_time && <div style={{ fontSize: 12, color: C.gray600 }}>Check-out: {new Date(todayAtt.check_out_time).toLocaleTimeString(lang === "bn" ? "bn-BD" : "en-GB", { hour: "2-digit", minute: "2-digit" })}</div>}
              {todayAtt.lat && <a href={`https://www.google.com/maps?q=${todayAtt.lat},${todayAtt.lng}`} target="_blank" rel="noreferrer" style={{ fontSize: 12, color: C.blue }}>{T.view_location}</a>}
            </div>
            {!todayAtt.check_out_time ? (
              <button disabled={busy} onClick={doCheckOut} style={{ ...btnEdit, background: C.red, color: C.white, border: "none" }}>🚪 Check Out</button>
            ) : (
              <Badge label="✅ কাজ শেষ" color="green" />
            )}
          </Card>

          <Card>
            <div style={{ marginBottom: 14 }}>
              <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 15 }}>{T.work_update}</div>
            </div>
            <textarea value={note} onChange={e => setNote(e.target.value)} placeholder={T.update_note_placeholder} style={{ ...inputStyle, minHeight: 60, resize: "vertical", marginBottom: 10 }} />
            {isFirstUpdate ? (
              <button disabled={busy} onClick={() => setShowCamera("update")} style={{ ...btnPrimary, marginTop: 0 }}>{busy ? T.uploading : "📸 " + T.give_update}</button>
            ) : (
              <div style={{ display: "flex", gap: 10 }}>
                <button disabled={busy} onClick={() => doWorkUpdate(null)} style={{ ...btnPrimary, marginTop: 0, flex: 1 }}>{busy ? T.uploading : T.give_update}</button>
                <button disabled={busy} onClick={() => setShowCamera("update")} style={{ ...btnEdit, flex: 1, textAlign: "center" }}>{T.take_photo}</button>
              </div>
            )}

            {updates.length > 0 && (
              <div style={{ marginTop: 18, borderTop: "1px solid " + C.gray100, paddingTop: 14 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.primaryDark, marginBottom: 10 }}>{T.todays_updates}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 10, maxHeight: 320, overflowY: "auto" }}>
                  {updates.map(u => (
                    <div key={u.id} style={{ display: "flex", gap: 10, alignItems: "flex-start", background: C.gray50, borderRadius: 10, padding: 10 }}>
                      {u.photo_url && <img src={u.photo_url} alt="update" style={{ width: 44, height: 44, borderRadius: 8, objectFit: "cover", flexShrink: 0 }} />}
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 11, color: C.gray600, marginBottom: 2 }}>{new Date(u.created_at).toLocaleTimeString(lang === "bn" ? "bn-BD" : "en-GB", { hour: "2-digit", minute: "2-digit" })}</div>
                        {u.note && <div style={{ fontSize: 12, color: C.gray800 }}>{u.note}</div>}
                        {u.lat && <a href={`https://www.google.com/maps?q=${u.lat},${u.lng}`} target="_blank" rel="noreferrer" style={{ fontSize: 11, color: C.blue }}>{T.view_location}</a>}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </Card>
        </>
      )}

      {showCamera && (
        <CameraCapture lang={lang} onCancel={() => setShowCamera(false)} onCapture={(img) => showCamera === "checkin" ? doCheckIn(img) : doWorkUpdate(img)} />
      )}

      <Card style={{ marginTop: 16 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>📅 সাপ্তাহিক সময়রেখা</div>
        <WeeklyTimeline employeeId={currentUser.employee_id} lang={lang} />
      </Card>

      <Card style={{ marginTop: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 15 }}>🌴 ছুটির আবেদন</div>
          <button onClick={() => setShowLeaveModal(true)} style={{ ...btnEdit, background: C.primary, color: C.white }}>➕ নতুন আবেদন</button>
        </div>
        {myLeaves.length === 0 ? (
          <div style={{ color: C.gray400, textAlign: "center", padding: 16, fontSize: 13 }}>এখনো কোনো আবেদন করেননি</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {myLeaves.map(r => (
              <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: C.gray50, borderRadius: 8, padding: "8px 12px", flexWrap: "wrap", gap: 6 }}>
                <div>
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.primaryDark }}>{LEAVE_TYPES.find(t => t.id === r.leave_type)?.label || r.leave_type}</div>
                  <div style={{ fontSize: 11, color: C.gray600 }}>{r.start_date} → {r.end_date} ({fmtNum(r.days)} দিন)</div>
                </div>
                <Badge label={r.status === "Pending" ? "⏳ Pending" : r.status === "Approved" ? "✅ Approved" : "❌ Rejected"} color={r.status === "Pending" ? "yellow" : r.status === "Approved" ? "green" : "red"} />
              </div>
            ))}
          </div>
        )}
      </Card>

      {showLeaveModal && (
        <Modal title="ছুটির আবেদন করুন" onClose={() => setShowLeaveModal(false)}>
          <FormField label="ছুটির ধরন">
            <select style={inputStyle} value={leaveForm.leave_type} onChange={e => setLeaveForm({ ...leaveForm, leave_type: e.target.value })}>
              {LEAVE_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
            </select>
          </FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="শুরুর তারিখ *"><input type="date" style={inputStyle} value={leaveForm.start_date} onChange={e => setLeaveForm({ ...leaveForm, start_date: e.target.value })} /></FormField>
            <FormField label="শেষের তারিখ *"><input type="date" style={inputStyle} value={leaveForm.end_date} onChange={e => setLeaveForm({ ...leaveForm, end_date: e.target.value })} /></FormField>
          </div>
          <FormField label="কারণ"><textarea style={{ ...inputStyle, minHeight: 60 }} value={leaveForm.reason} onChange={e => setLeaveForm({ ...leaveForm, reason: e.target.value })} /></FormField>
          <button onClick={submitLeave} style={btnPrimary}>✅ আবেদন জমা দিন</button>
        </Modal>
      )}

      <Card style={{ marginTop: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 15 }}>📝 বিশেষ অনুরোধ (Overtime / Regularization / On Duty / Permission / Shift)</div>
          <button onClick={() => setShowReqModal(true)} style={{ ...btnEdit, background: C.primary, color: C.white }}>➕ নতুন অনুরোধ</button>
        </div>
        {myRequests.length === 0 ? (
          <div style={{ color: C.gray400, textAlign: "center", padding: 16, fontSize: 13 }}>এখনো কোনো অনুরোধ করেননি</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {myRequests.map(r => (
              <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: C.gray50, borderRadius: 8, padding: "8px 12px", flexWrap: "wrap", gap: 6 }}>
                <div>
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.primaryDark }}>{REQUEST_TYPES[r.type]?.icon} {REQUEST_TYPES[r.type]?.label}</div>
                  <div style={{ fontSize: 11, color: C.gray600 }}>{r.date}{r.start_time ? " · " + r.start_time + "-" + r.end_time : ""}</div>
                </div>
                <Badge label={r.status === "Pending" ? "⏳ Pending" : r.status === "Approved" ? "✅ Approved" : "❌ Rejected"} color={r.status === "Pending" ? "yellow" : r.status === "Approved" ? "green" : "red"} />
              </div>
            ))}
          </div>
        )}
      </Card>

      {showReqModal && (
        <Modal title="নতুন অনুরোধ জমা দিন" onClose={() => setShowReqModal(false)}>
          <FormField label="অনুরোধের ধরন">
            <select style={inputStyle} value={reqType} onChange={e => setReqType(e.target.value)}>
              {Object.entries(REQUEST_TYPES).map(([id, cfg]) => <option key={id} value={id}>{cfg.icon} {cfg.label}</option>)}
            </select>
          </FormField>
          <FormField label="তারিখ *"><input type="date" style={inputStyle} value={reqForm.date} onChange={e => setReqForm({ ...reqForm, date: e.target.value })} /></FormField>
          {REQUEST_TYPES[reqType].needsTime && (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <FormField label="শুরুর সময়"><input type="time" style={inputStyle} value={reqForm.start_time} onChange={e => setReqForm({ ...reqForm, start_time: e.target.value })} /></FormField>
              <FormField label="শেষের সময়"><input type="time" style={inputStyle} value={reqForm.end_time} onChange={e => setReqForm({ ...reqForm, end_time: e.target.value })} /></FormField>
            </div>
          )}
          {REQUEST_TYPES[reqType].needsShift && (
            <FormField label="চাহিত Shift *">
              <select style={inputStyle} value={reqForm.requested_shift_id} onChange={e => setReqForm({ ...reqForm, requested_shift_id: e.target.value })}>
                <option value="">— বাছাই করুন —</option>
                {shiftsList.map(s => <option key={s.id} value={s.id}>{s.name} ({s.start_time}-{s.end_time})</option>)}
              </select>
            </FormField>
          )}
          <FormField label="কারণ"><textarea style={{ ...inputStyle, minHeight: 60 }} value={reqForm.reason} onChange={e => setReqForm({ ...reqForm, reason: e.target.value })} /></FormField>
          <button onClick={submitRequest} style={btnPrimary}>✅ জমা দিন</button>
        </Modal>
      )}

      <Card style={{ marginTop: 16 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>🎯 আমার পারফরম্যান্স</div>
        {myReviews.length === 0 ? (
          <div style={{ color: C.gray400, textAlign: "center", padding: 16, fontSize: 13 }}>এখনো কোনো Review Cycle নেই</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {myReviews.map(r => (
              <div key={r.id} style={{ border: "1px solid " + C.gray100, borderRadius: 10, padding: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: C.primaryDark }}>{r.period}</div>
                  <Badge label={r.status} color={r.status === "Completed" ? "green" : r.status === "Submitted" ? "yellow" : "gray"} />
                </div>
                {r.goals && <div style={{ fontSize: 12, color: C.gray600, marginBottom: 8 }}><strong>লক্ষ্য:</strong> {r.goals}</div>}
                {r.status === "Draft" ? (
                  <>
                    <textarea value={selfReviewDraft[r.id] ?? r.self_review ?? ""} onChange={e => setSelfReviewDraft({ ...selfReviewDraft, [r.id]: e.target.value })} placeholder="নিজের performance সম্পর্কে লিখুন..." style={{ ...inputStyle, minHeight: 60, marginBottom: 8 }} />
                    <button onClick={() => submitSelfReview(r.id)} style={{ ...btnEdit, background: C.primary, color: C.white }}>✅ Self Review জমা দিন</button>
                  </>
                ) : (
                  <>
                    <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}><strong>Self Review:</strong> {r.self_review || "—"}</div>
                    {r.manager_review && <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}><strong>Manager Review:</strong> {r.manager_review}</div>}
                    {r.rating && <div style={{ fontSize: 13 }}>{"⭐".repeat(r.rating)}</div>}
                  </>
                )}
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card style={{ marginTop: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 15 }}>⏲️ টাইম ট্র্যাকার</div>
          <button onClick={() => setShowTimeLogModal(true)} style={{ ...btnEdit, background: C.primary, color: C.white }}>➕ সময় লগ করুন</button>
        </div>
        {myTimeLogs.length === 0 ? (
          <div style={{ color: C.gray400, textAlign: "center", padding: 16, fontSize: 13 }}>এখনো কোনো লগ নেই</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {myTimeLogs.map(l => (
              <div key={l.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: C.gray50, borderRadius: 8, padding: "8px 12px", flexWrap: "wrap", gap: 6 }}>
                <div>
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.primaryDark }}>{l.task}{l.project_name ? " — " + l.project_name : ""}</div>
                  <div style={{ fontSize: 11, color: C.gray600 }}>{l.date}{l.start_time ? " · " + l.start_time + "-" + l.end_time : ""}</div>
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card style={{ marginTop: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 15 }}>🎫 HR Help Desk</div>
          <button onClick={() => setShowTicketModal(true)} style={{ ...btnEdit, background: C.primary, color: C.white }}>➕ নতুন টিকেট</button>
        </div>
        {myTickets.length === 0 ? (
          <div style={{ color: C.gray400, textAlign: "center", padding: 16, fontSize: 13 }}>এখনো কোনো টিকেট নেই</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {myTickets.map(t => (
              <div key={t.id} style={{ border: "1px solid " + C.gray100, borderRadius: 8, padding: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.primaryDark }}>{t.subject}</div>
                  <Badge label={t.status} color={t.status === "Open" ? "yellow" : t.status === "In Progress" ? "primary" : "green"} />
                </div>
                <div style={{ fontSize: 11, color: C.gray600, marginBottom: t.admin_reply ? 6 : 0 }}>{t.category}</div>
                {t.admin_reply && <div style={{ fontSize: 11, color: C.primaryDark, background: C.primaryBg, borderRadius: 6, padding: 6 }}><strong>Admin:</strong> {t.admin_reply}</div>}
              </div>
            ))}
          </div>
        )}
      </Card>

      {showTimeLogModal && (
        <Modal title="সময় লগ করুন" onClose={() => setShowTimeLogModal(false)}>
          <FormField label="কাজ *"><input style={inputStyle} value={timeLogForm.task} onChange={e => setTimeLogForm({ ...timeLogForm, task: e.target.value })} placeholder="যেমন: BOQ তৈরি" /></FormField>
          <FormField label="প্রজেক্ট (ঐচ্ছিক)"><input style={inputStyle} value={timeLogForm.project_name} onChange={e => setTimeLogForm({ ...timeLogForm, project_name: e.target.value })} /></FormField>
          <FormField label="তারিখ *"><input type="date" style={inputStyle} value={timeLogForm.date} onChange={e => setTimeLogForm({ ...timeLogForm, date: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="শুরুর সময়"><input type="time" style={inputStyle} value={timeLogForm.start_time} onChange={e => setTimeLogForm({ ...timeLogForm, start_time: e.target.value })} /></FormField>
            <FormField label="শেষের সময়"><input type="time" style={inputStyle} value={timeLogForm.end_time} onChange={e => setTimeLogForm({ ...timeLogForm, end_time: e.target.value })} /></FormField>
          </div>
          <FormField label="নোট"><textarea style={{ ...inputStyle, minHeight: 50 }} value={timeLogForm.note} onChange={e => setTimeLogForm({ ...timeLogForm, note: e.target.value })} /></FormField>
          <button onClick={submitTimeLog} style={btnPrimary}>✅ লগ করুন</button>
        </Modal>
      )}

      {showTicketModal && (
        <Modal title="নতুন HR টিকেট" onClose={() => setShowTicketModal(false)}>
          <FormField label="বিষয় *"><input style={inputStyle} value={ticketForm.subject} onChange={e => setTicketForm({ ...ticketForm, subject: e.target.value })} /></FormField>
          <FormField label="Category">
            <select style={inputStyle} value={ticketForm.category} onChange={e => setTicketForm({ ...ticketForm, category: e.target.value })}>
              {TICKET_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </FormField>
          <FormField label="বিবরণ"><textarea style={{ ...inputStyle, minHeight: 80 }} value={ticketForm.description} onChange={e => setTicketForm({ ...ticketForm, description: e.target.value })} /></FormField>
          <button onClick={submitTicket} style={btnPrimary}>✅ জমা দিন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// PAYROLL / PAYSLIP
// ============================================================
function LineItemsEditor({ title, items, onChange, color }) {
  const update = (i, field, val) => {
    const next = items.slice();
    next[i] = { ...next[i], [field]: val };
    onChange(next);
  };
  const addRow = () => onChange([...items, { label: "", amount: 0, note: "" }]);
  const removeRow = (i) => onChange(items.filter((_, idx) => idx !== i));
  const subtotal = items.reduce((s, it) => s + (+it.amount || 0), 0);

  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: color || C.primaryDark }}>{title}</div>
        <button onClick={addRow} style={{ ...btnEdit, padding: "4px 10px", fontSize: 11 }}>➕ লাইন যোগ করুন</button>
      </div>
      <div style={{ border: "1px solid " + C.gray200, borderRadius: 8, overflow: "hidden" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1.3fr 0.8fr 1.3fr 30px", gap: 0, background: C.primaryBg, padding: "6px 10px", fontSize: 11, fontWeight: 700, color: C.primaryDark }}>
          <div>বিবরণ</div><div>পরিমাণ (৳)</div><div>নোট</div><div></div>
        </div>
        {items.length === 0 ? (
          <div style={{ padding: 12, fontSize: 12, color: C.gray400, textAlign: "center" }}>কোনো লাইন নেই — "লাইন যোগ করুন" চাপুন</div>
        ) : items.map((it, i) => (
          <div key={i} style={{ display: "grid", gridTemplateColumns: "1.3fr 0.8fr 1.3fr 30px", gap: 6, padding: "6px 10px", borderTop: "1px solid " + C.gray100, alignItems: "center" }}>
            <input value={it.label} onChange={e => update(i, "label", e.target.value)} placeholder="বিবরণ" style={{ ...inputStyle, padding: "5px 8px", fontSize: 12 }} />
            <input type="number" value={it.amount} onChange={e => update(i, "amount", e.target.value)} style={{ ...inputStyle, padding: "5px 8px", fontSize: 12 }} />
            <input value={it.note || ""} onChange={e => update(i, "note", e.target.value)} placeholder="নোট" style={{ ...inputStyle, padding: "5px 8px", fontSize: 12 }} />
            <button onClick={() => removeRow(i)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 14 }}>🗑️</button>
          </div>
        ))}
        <div style={{ padding: "8px 10px", background: C.gray50, fontSize: 12, fontWeight: 700, textAlign: "right", color: C.primaryDark }}>সাবটোটাল: {fmt(subtotal)}</div>
      </div>
    </div>
  );
}

function Payroll({ employees }) {
  const [month, setMonth] = useState(new Date().toISOString().slice(0, 7));
  const [runs, setRuns] = useState([]);
  const [loading, setLoading] = useState(true);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState(null);
  const [printRow, setPrintRow] = useState(null);
  const [tab, setTab] = useState("list");
  const uploadRef = useRef();

  const load = async () => {
    setLoading(true);
    const { data } = await supabase.from("payroll_runs").select("*").eq("month", month);
    const empOrder = {}; employees.forEach((e, i) => { empOrder[e.id] = i; });
    const sorted = (data || []).slice().sort((a, b) => (empOrder[a.employee_id] ?? 999999) - (empOrder[b.employee_id] ?? 999999));
    setRuns(sorted);
    setLoading(false);
  };
  useEffect(() => { load(); }, [month, employees]);

  const empById = {}; employees.forEach(e => { empById[e.id] = e; });
  const subtotalOf = (items) => (items || []).reduce((s, it) => s + (+it.amount || 0), 0);
  const netOf = (r) => subtotalOf(r.fixed_items) + subtotalOf(r.kpi_items) - subtotalOf(r.penalty_items);

  const generateRun = async () => {
    const existingIds = new Set(runs.map(r => r.employee_id));
    const toCreate = employees.filter(e => e.status === "কর্মরত" && !existingIds.has(e.id)).map(e => ({
      employee_id: e.id, month, position: e.role,
      fixed_items: [{ label: "Fixed Salary", amount: e.salary || 0, note: "মাসিক মূল বেতন" }],
      kpi_items: [], penalty_items: [],
      net_pay: e.salary || 0, status: "Unpaid", disbursement_channel: "Bank",
    }));
    if (toCreate.length === 0) return alert("এই মাসের জন্য সব কর্মীর পে-রোল ইতিমধ্যে তৈরি আছে");
    await supabase.from("payroll_runs").insert(toCreate);
    load();
  };

  const openEdit = (r) => {
    setEditItem(r);
    let penaltyItems = r.penalty_items && r.penalty_items.length ? r.penalty_items : [];
    if (penaltyItems.length === 0 && (r.penalty_days || 0) * (r.penalty_rate || 0) > 0) {
      penaltyItems = [{ label: `Penalty (${r.penalty_days} দিন × ৳${r.penalty_rate})`, amount: r.penalty_days * r.penalty_rate, note: "ধারা ১২.১" }];
    }
    setForm({
      position: r.position || empById[r.employee_id]?.role || "",
      fixed_items: r.fixed_items && r.fixed_items.length ? r.fixed_items : [{ label: "Fixed Salary", amount: r.basic || 0, note: "" }],
      kpi_items: r.kpi_items || [],
      penalty_items: penaltyItems,
      quick_days: "", quick_rate: "",
      disbursement_channel: r.disbursement_channel || "Bank",
      disbursement_date: r.disbursement_date || "",
    });
  };

  const addQuickPenalty = () => {
    const days = +form.quick_days || 0, rate = +form.quick_rate || 0;
    if (days <= 0 || rate <= 0) return;
    const line = { label: `Penalty (${days} দিন × ৳${rate})`, amount: days * rate, note: "ধারা ১২.১ অনুযায়ী" };
    setForm({ ...form, penalty_items: [...form.penalty_items, line], quick_days: "", quick_rate: "" });
  };

  const saveEdit = async () => {
    const net = subtotalOf(form.fixed_items) + subtotalOf(form.kpi_items) - subtotalOf(form.penalty_items);
    await supabase.from("payroll_runs").update({
      position: form.position, fixed_items: form.fixed_items, kpi_items: form.kpi_items,
      penalty_items: form.penalty_items,
      net_pay: net, disbursement_channel: form.disbursement_channel, disbursement_date: form.disbursement_date || null,
    }).eq("id", editItem.id);
    setEditItem(null); setForm(null); load();
  };

  const markPaid = async (r) => {
    await supabase.from("payroll_runs").update({ status: "Paid", paid_at: new Date().toISOString() }).eq("id", r.id);
    load();
  };

  const setChannelDate = async (channel, date) => {
    const ids = runs.filter(r => (r.disbursement_channel || "Bank") === channel).map(r => r.id);
    if (ids.length === 0) return;
    await supabase.from("payroll_runs").update({ disbursement_date: date }).in("id", ids);
    load();
  };

  const doPrint = (r) => {
    setPrintRow(r);
    setTimeout(() => printSection("Salary Sheet — " + (empById[r.employee_id]?.name || ""), "payslip-content"), 100);
  };

  const handleExport = () => exportToExcel(runs.map(r => ({
    নাম: empById[r.employee_id]?.name, পদবি: r.position || empById[r.employee_id]?.role,
    "Fixed সাবটোটাল": subtotalOf(r.fixed_items), "KPI সাবটোটাল": subtotalOf(r.kpi_items),
    "মোট পেনাল্টি": subtotalOf(r.penalty_items),
    "নীট প্রদেয়": netOf(r), চ্যানেল: r.disbursement_channel || "Bank", স্ট্যাটাস: r.status,
  })), "Payroll", "Payroll_" + month);

  const handleUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      const emp = employees.find(em => em.name === (row["নাম"] || row["name"]));
      if (!emp) continue;
      const fixed_items = [
        { label: "Fixed Salary", amount: +row["Fixed Salary"] || +row["মূল বেতন"] || emp.salary || 0, note: "মাসিক মূল বেতন" },
        { label: "House Rent", amount: +row["House Rent"] || +row["বাড়ি ভাড়া"] || 0, note: "ঘর ভাড়া" },
        { label: "Site Visit / Field Allowance", amount: +row["Site Visit Allowance"] || +row["সাইট ভাতা"] || 0, note: "সাইট পরিদর্শন ভাতা" },
      ].filter(it => it.amount > 0);
      const kpiAmt = +row["KPI Bonus"] || +row["কেপিআই বোনাস"] || 0;
      const kpi_items = kpiAmt > 0 ? [{ label: "Overall Project Delivery KPI Bonus", amount: kpiAmt, note: "" }] : [];
      const penalty_days = +row["Penalty Days"] || +row["পেনাল্টি দিন"] || 0;
      const penalty_rate = +row["Penalty Rate"] || +row["পেনাল্টি হার"] || 0;
      const penalty_items = penalty_days * penalty_rate > 0 ? [{ label: `Penalty (${penalty_days} দিন × ৳${penalty_rate})`, amount: penalty_days * penalty_rate, note: "ধারা ১২.১" }] : [];
      const net_pay = subtotalOf(fixed_items) + subtotalOf(kpi_items) - subtotalOf(penalty_items);

      const existing = runs.find(r => r.employee_id === emp.id);
      const payload = { employee_id: emp.id, month, position: emp.role, fixed_items, kpi_items, penalty_items, net_pay, status: existing?.status || "Unpaid", disbursement_channel: existing?.disbursement_channel || "Bank" };
      if (existing) await supabase.from("payroll_runs").update(payload).eq("id", existing.id);
      else await supabase.from("payroll_runs").insert([payload]);
      count++;
    }
    alert("✅ " + count + " জনের পে-রোল আপলোড/আপডেট হয়েছে!");
    e.target.value = ""; load();
  };

  const totalNet = runs.reduce((s, r) => s + netOf(r), 0);
  const totalFixed = runs.reduce((s, r) => s + subtotalOf(r.fixed_items), 0);
  const totalKpi = runs.reduce((s, r) => s + subtotalOf(r.kpi_items), 0);
  const totalPenalty = runs.reduce((s, r) => s + subtotalOf(r.penalty_items), 0);
  const paidCount = runs.filter(r => r.status === "Paid").length;

  const channelGroups = DISBURSEMENT_CHANNELS.map(ch => {
    const items = runs.filter(r => (r.disbursement_channel || "Bank") === ch.id);
    return { ...ch, count: items.length, totalDeduction: items.reduce((s, r) => s + subtotalOf(r.penalty_items), 0), totalNet: items.reduce((s, r) => s + netOf(r), 0), date: items[0]?.disbursement_date || "" };
  }).filter(g => g.count > 0);
  const channelDonut = channelGroups.map(g => ({ value: g.count, color: g.color }));
  const earningsDonut = [
    { value: totalFixed, color: C.primary },
    { value: totalKpi, color: C.green },
    { value: totalPenalty, color: C.red },
  ];

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="👷" label="মোট কর্মী (এই মাসে)" value={fmtNum(runs.length)} color={C.primaryBg} />
        <StatCard icon="✅" label="পরিশোধিত" value={fmtNum(paidCount) + " / " + fmtNum(runs.length)} color={C.greenLight} />
        <StatCard icon="💰" label="মোট নীট বেতন" value={fmt(totalNet)} color="#FFF8E1" />
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ fontWeight: 600, color: C.gray800, fontSize: 14 }}>মাস:</div>
          <input type="month" value={month} onChange={e => setMonth(e.target.value)} style={{ ...inputStyle, width: "auto" }} />
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={generateRun} style={{ ...btnPrimary, width: "auto", margin: 0 }}>⚙️ এই মাসের পে-রোল তৈরি করুন</button>
          <button onClick={handleExport} style={btnEdit}>⬇️ Excel Download</button>
          <button onClick={() => uploadRef.current?.click()} style={btnEdit}>⬆️ Excel Upload</button>
          <button onClick={() => printSection("মোট বেতন খরচ — " + month, "payroll-total-content")} style={btnEdit}>🖨️ মোট খরচ প্রিন্ট</button>
          <input type="file" ref={uploadRef} onChange={handleUpload} accept=".xlsx,.xls,.csv" style={{ display: "none" }} />
        </div>
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {[["list", "📋 তালিকা"], ["disbursement", "🏦 বিতরণ চ্যানেল"], ["analytics", "📊 Analytics"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{ background: tab === id ? C.primary : C.white, color: tab === id ? C.white : C.gray800, border: "1px solid " + (tab === id ? C.primary : C.gray200), borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
        ))}
      </div>

      {tab === "list" && (
        <Card>
          {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : runs.length === 0 ? (
            <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>এই মাসের জন্য এখনো পে-রোল তৈরি হয়নি। উপরের বাটনে ক্লিক করুন অথবা Excel Upload করুন।</div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "পদবি", "Fixed", "KPI", "পেনাল্টি", "নীট প্রদেয়", "চ্যানেল", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
                <tbody>
                  {runs.map(r => (
                    <tr key={r.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                      <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{empById[r.employee_id]?.name || "—"}</td>
                      <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600 }}>{r.position || empById[r.employee_id]?.role}</td>
                      <td style={{ padding: "10px 14px", color: C.green }}>{fmt(subtotalOf(r.fixed_items))}</td>
                      <td style={{ padding: "10px 14px", color: C.green }}>{fmt(subtotalOf(r.kpi_items))}</td>
                      <td style={{ padding: "10px 14px", color: C.red }}>-{fmt(subtotalOf(r.penalty_items))}</td>
                      <td style={{ padding: "10px 14px", fontWeight: 700, color: C.primaryDark }}>{fmt(netOf(r))}</td>
                      <td style={{ padding: "10px 14px" }}><Badge label={DISBURSEMENT_CHANNELS.find(c => c.id === (r.disbursement_channel || "Bank"))?.label} color="primary" /></td>
                      <td style={{ padding: "10px 14px" }}><Badge label={r.status === "Paid" ? "✅ Paid" : "⏳ Unpaid"} color={r.status === "Paid" ? "green" : "yellow"} /></td>
                      <td style={{ padding: "10px 14px" }}>
                        <div style={{ display: "flex", gap: 6 }}>
                          <button onClick={() => openEdit(r)} style={btnEdit}>✏️</button>
                          <button onClick={() => doPrint(r)} style={btnEdit}>🖨️</button>
                          {r.status !== "Paid" && <button onClick={() => markPaid(r)} style={{ ...btnEdit, background: C.greenLight, color: C.green }}>✓ Paid</button>}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </Card>
      )}

      {tab === "disbursement" && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(280px,1fr))", gap: 14 }}>
          {channelGroups.length === 0 ? (
            <Card><div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>এই মাসে কোনো পে-রোল নেই</div></Card>
          ) : channelGroups.map(g => (
            <Card key={g.id}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
                <span style={{ width: 10, height: 10, borderRadius: "50%", background: g.color, display: "inline-block" }} />
                <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14 }}>{g.label}</div>
              </div>
              <ProfileField label="মোট কর্মী" value={fmtNum(g.count) + " জন"} />
              <ProfileField label="মোট পেনাল্টি/কর্তন" value={fmt(g.totalDeduction)} />
              <ProfileField label="মোট নীট প্রদেয়" value={fmt(g.totalNet)} />
              <FormField label="বিতরণের তারিখ">
                <input type="date" style={inputStyle} value={g.date || ""} onChange={e => setChannelDate(g.id, e.target.value)} />
              </FormField>
            </Card>
          ))}
        </div>
      )}

      {tab === "analytics" && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
          <Card>
            <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>🏦 Disbursement Channel</div>
            {channelGroups.length === 0 ? <div style={{ color: C.gray400, textAlign: "center", padding: 20, fontSize: 13 }}>ডেটা নেই</div> : (
              <div style={{ display: "flex", gap: 20, alignItems: "center", flexWrap: "wrap" }}>
                <DonutChart segments={channelDonut} centerLabel={fmtNum(runs.length)} centerSub="মোট কর্মী" />
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {channelGroups.map(g => (
                    <div key={g.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                      <span style={{ width: 10, height: 10, borderRadius: "50%", background: g.color, display: "inline-block" }} />
                      <span style={{ color: C.gray800, minWidth: 90 }}>{g.label}</span>
                      <span style={{ fontWeight: 700, color: C.primaryDark }}>{fmtNum(g.count)}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </Card>
          <Card>
            <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>💰 Earnings Overview</div>
            {runs.length === 0 ? <div style={{ color: C.gray400, textAlign: "center", padding: 20, fontSize: 13 }}>ডেটা নেই</div> : (
              <div style={{ display: "flex", gap: 20, alignItems: "center", flexWrap: "wrap" }}>
                <DonutChart segments={earningsDonut} centerLabel={fmt(totalNet)} centerSub="নীট বেতন" />
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}><span style={{ width: 10, height: 10, borderRadius: "50%", background: C.primary, display: "inline-block" }} /><span style={{ color: C.gray800, minWidth: 90 }}>Fixed Part</span><span style={{ fontWeight: 700, color: C.primaryDark }}>{fmt(totalFixed)}</span></div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}><span style={{ width: 10, height: 10, borderRadius: "50%", background: C.green, display: "inline-block" }} /><span style={{ color: C.gray800, minWidth: 90 }}>KPI Part</span><span style={{ fontWeight: 700, color: C.primaryDark }}>{fmt(totalKpi)}</span></div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}><span style={{ width: 10, height: 10, borderRadius: "50%", background: C.red, display: "inline-block" }} /><span style={{ color: C.gray800, minWidth: 90 }}>পেনাল্টি</span><span style={{ fontWeight: 700, color: C.primaryDark }}>{fmt(totalPenalty)}</span></div>
                </div>
              </div>
            )}
          </Card>
        </div>
      )}

      {editItem && form && (
        <Modal title={"পে-রোল সম্পাদনা — " + (empById[editItem.employee_id]?.name || "")} onClose={() => { setEditItem(null); setForm(null); }} size={620}>
          <FormField label="পদবি (Position For)"><input style={inputStyle} value={form.position} onChange={e => setForm({ ...form, position: e.target.value })} /></FormField>

          <LineItemsEditor title="Fixed Part" items={form.fixed_items} onChange={items => setForm({ ...form, fixed_items: items })} color={C.primary} />
          <LineItemsEditor title="KPI Part" items={form.kpi_items} onChange={items => setForm({ ...form, kpi_items: items })} color={C.green} />
          <LineItemsEditor title="Penalty (ধারা ১২.১ অনুযায়ী)" items={form.penalty_items} onChange={items => setForm({ ...form, penalty_items: items })} color={C.red} />

          <div style={{ border: "1px dashed " + C.gray200, borderRadius: 8, padding: 10, marginBottom: 16, background: C.gray50 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: C.gray600, marginBottom: 6 }}>⚡ দ্রুত হিসাব (দিন × দৈনিক হার) — যোগ করলে উপরের Penalty লিস্টে লাইন হিসেবে যুক্ত হবে</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 8, alignItems: "end" }}>
              <FormField label="দিন সংখ্যা"><input type="number" style={inputStyle} value={form.quick_days} onChange={e => setForm({ ...form, quick_days: e.target.value })} /></FormField>
              <FormField label="দৈনিক হার (৳)"><input type="number" style={inputStyle} value={form.quick_rate} onChange={e => setForm({ ...form, quick_rate: e.target.value })} /></FormField>
              <button onClick={addQuickPenalty} style={{ ...btnEdit, marginBottom: 14 }}>➕ যোগ করুন</button>
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 8 }}>
            <FormField label="বিতরণ চ্যানেল">
              <select style={inputStyle} value={form.disbursement_channel} onChange={e => setForm({ ...form, disbursement_channel: e.target.value })}>
                {DISBURSEMENT_CHANNELS.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
              </select>
            </FormField>
            <FormField label="বিতরণের তারিখ"><input type="date" style={inputStyle} value={form.disbursement_date} onChange={e => setForm({ ...form, disbursement_date: e.target.value })} /></FormField>
          </div>

          <div style={{ fontSize: 15, fontWeight: 700, color: C.primaryDark, marginBottom: 14, padding: "10px 14px", background: C.primaryBg, borderRadius: 8 }}>
            মোট প্রদেয় (Net Payable): {fmt(subtotalOf(form.fixed_items) + subtotalOf(form.kpi_items) - subtotalOf(form.penalty_items))}
          </div>
          <button onClick={saveEdit} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}

      {/* Hidden total cost content for printing */}
      <div style={{ display: "none" }}>
        <div id="payroll-total-content">
          <table>
            <thead><tr><th>কর্মী</th><th>পদবি</th><th>Fixed</th><th>KPI</th><th>পেনাল্টি</th><th>নীট প্রদেয়</th><th>চ্যানেল</th></tr></thead>
            <tbody>
              {runs.map(r => (
                <tr key={r.id}>
                  <td>{empById[r.employee_id]?.name}</td>
                  <td>{r.position || empById[r.employee_id]?.role}</td>
                  <td>{fmt(subtotalOf(r.fixed_items))}</td>
                  <td>{fmt(subtotalOf(r.kpi_items))}</td>
                  <td>-{fmt(subtotalOf(r.penalty_items))}</td>
                  <td style={{ fontWeight: 700 }}>{fmt(netOf(r))}</td>
                  <td>{DISBURSEMENT_CHANNELS.find(c => c.id === (r.disbursement_channel || "Bank"))?.label}</td>
                </tr>
              ))}
              <tr><td colSpan={5} style={{ fontWeight: 700 }}>সর্বমোট (Total Office Salary Cost)</td><td style={{ fontWeight: 700 }}>{fmt(totalNet)}</td><td></td></tr>
            </tbody>
          </table>
        </div>
      </div>

      {/* Hidden payslip content for printing — matches Salary Sheet layout */}
      <div style={{ display: "none" }}>
        <div id="payslip-content">
          {printRow && (
            <div>
              <table style={{ marginBottom: 10 }}>
                <tbody>
                  <tr><td style={{ fontWeight: 700, width: "25%" }}>Position For:</td><td colSpan={3}>{printRow.position || empById[printRow.employee_id]?.role}</td></tr>
                  <tr><td style={{ fontWeight: 700 }}>কর্মচারীর নাম:</td><td>{empById[printRow.employee_id]?.name}</td><td style={{ fontWeight: 700 }}>মাস/বছর:</td><td>{printRow.month}</td></tr>
                </tbody>
              </table>

              <div style={{ fontWeight: 700, margin: "10px 0 4px" }}>Fixed Part</div>
              <table style={{ marginBottom: 10 }}>
                <thead><tr><th>বিবরণ</th><th>পরিমাণ (৳)</th><th>নোট</th></tr></thead>
                <tbody>
                  {(printRow.fixed_items || []).map((it, i) => <tr key={i}><td>{it.label}</td><td>{fmt(it.amount)}</td><td>{it.note}</td></tr>)}
                  <tr><td style={{ fontWeight: 700 }}>Fixed Part সাবটোটাল</td><td style={{ fontWeight: 700 }}>{fmt(subtotalOf(printRow.fixed_items))}</td><td></td></tr>
                </tbody>
              </table>

              <div style={{ fontWeight: 700, margin: "10px 0 4px" }}>KPI Part</div>
              <table style={{ marginBottom: 10 }}>
                <thead><tr><th>বিবরণ</th><th>পরিমাণ (৳)</th><th>নোট</th></tr></thead>
                <tbody>
                  {(printRow.kpi_items || []).length === 0 ? <tr><td colSpan={3}>—</td></tr> : (printRow.kpi_items || []).map((it, i) => <tr key={i}><td>{it.label}</td><td>{fmt(it.amount)}</td><td>{it.note}</td></tr>)}
                  <tr><td style={{ fontWeight: 700 }}>KPI Part সাবটোটাল</td><td style={{ fontWeight: 700 }}>{fmt(subtotalOf(printRow.kpi_items))}</td><td></td></tr>
                </tbody>
              </table>

              <div style={{ fontWeight: 700, margin: "10px 0 4px" }}>Penalty (ধারা ১২.১ অনুযায়ী)</div>
              <table style={{ marginBottom: 10 }}>
                <thead><tr><th>বিবরণ</th><th>পরিমাণ (৳)</th><th>নোট</th></tr></thead>
                <tbody>
                  {(printRow.penalty_items || []).length === 0 ? <tr><td colSpan={3}>—</td></tr> : (printRow.penalty_items || []).map((it, i) => <tr key={i}><td>{it.label}</td><td>{fmt(it.amount)}</td><td>{it.note}</td></tr>)}
                  <tr><td style={{ fontWeight: 700 }}>মোট পেনাল্টি কর্তন</td><td style={{ fontWeight: 700 }}>{fmt(subtotalOf(printRow.penalty_items))}</td><td></td></tr>
                </tbody>
              </table>

              <table>
                <tbody>
                  <tr><td style={{ fontWeight: 700, fontSize: "11pt" }}>মোট প্রদেয় (Net Payable)</td><td style={{ fontWeight: 700, fontSize: "11pt", color: "#2e7d32" }}>{fmt(netOf(printRow))}</td></tr>
                  <tr><td>বিতরণ চ্যানেল</td><td>{DISBURSEMENT_CHANNELS.find(c => c.id === (printRow.disbursement_channel || "Bank"))?.label}</td></tr>
                  <tr><td>স্ট্যাটাস</td><td>{printRow.status === "Paid" ? "পরিশোধিত" : "অপরিশোধিত"}</td></tr>
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
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
    alert("✅ " + count + "টি লেনদেন আপলোড হয়েছে!"); onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="💰" label="মোট আয়" value={fmt(totalIncome)} color="#F0FFF4" />
        <StatCard icon="💸" label="মোট ব্যয়" value={fmt(totalExpense)} color="#FFF5F5" />
        <StatCard icon="📊" label="নিট লাভ" value={fmt(totalIncome - totalExpense)} color={totalIncome >= totalExpense ? "#F0FFF4" : "#FFF5F5"} />
      </div>
      <SectionHeader title="আর্থিক ব্যবস্থাপনা" action="নতুন লেনদেন" onAction={openAdd} onExport={handleExport} onPrint={() => { printSection("আর্থিক রিপোর্ট", "finance-content"); }} onUpload={handleUpload} uploadRef={uploadRef} />
      <Card>
        <div id="finance-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "ধরন", "বিভাগ", "বিবরণ", "পরিমাণ", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>{h}</th>)}</tr></thead>
            <tbody>
              {[...data].reverse().map(t => (
                <tr key={t.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
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
// RECRUITMENT
// ============================================================
const CANDIDATE_STAGES = [
  { id: "Applied", label: "আবেদনকৃত", color: C.blue },
  { id: "Interview", label: "ইন্টারভিউ", color: "#E0A800" },
  { id: "Hired", label: "নিয়োগপ্রাপ্ত", color: C.green },
  { id: "Rejected", label: "বাতিল", color: C.red },
];

function Recruitment() {
  const [jobs, setJobs] = useState([]);
  const [selectedJob, setSelectedJob] = useState(null);
  const [candidates, setCandidates] = useState([]);
  const [showJobModal, setShowJobModal] = useState(false);
  const [showCandModal, setShowCandModal] = useState(false);
  const [jobForm, setJobForm] = useState({ title: "", department: "ডিজাইন", type: "Full Time", description: "", status: "Open" });
  const [candForm, setCandForm] = useState({ name: "", phone: "", email: "", resume_url: "", notes: "" });
  const candUploadRef = useRef();

  const loadJobs = async () => { const { data } = await supabase.from("job_postings").select("*").order("posted_at", { ascending: false }); setJobs(data || []); if (!selectedJob && data?.length) setSelectedJob(data[0]); };
  const loadCandidates = async (jobId) => { if (!jobId) return setCandidates([]); const { data } = await supabase.from("candidates").select("*").eq("job_id", jobId).order("applied_at", { ascending: false }); setCandidates(data || []); };

  useEffect(() => { loadJobs(); }, []);
  useEffect(() => { loadCandidates(selectedJob?.id); }, [selectedJob?.id]);

  const saveJob = async () => {
    if (!jobForm.title) return alert("পদের নাম আবশ্যক");
    const { data, error } = await supabase.from("job_postings").insert([{ ...jobForm, posted_at: new Date().toISOString() }]).select().single();
    if (error) return alert("Error: " + error.message);
    setShowJobModal(false); setJobForm({ title: "", department: "ডিজাইন", type: "Full Time", description: "", status: "Open" });
    await loadJobs(); setSelectedJob(data);
  };

  const toggleJobStatus = async (job) => {
    await supabase.from("job_postings").update({ status: job.status === "Open" ? "Closed" : "Open" }).eq("id", job.id);
    loadJobs();
  };

  const saveCandidate = async () => {
    if (!candForm.name) return alert("নাম আবশ্যক");
    await supabase.from("candidates").insert([{ ...candForm, job_id: selectedJob.id, stage: "Applied", applied_at: new Date().toISOString() }]);
    setShowCandModal(false); setCandForm({ name: "", phone: "", email: "", resume_url: "", notes: "" });
    loadCandidates(selectedJob.id);
  };

  const moveStage = async (cand, stage) => {
    await supabase.from("candidates").update({ stage }).eq("id", cand.id);
    loadCandidates(selectedJob.id);
  };

  const uploadResume = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const url = await uploadImage(file, "resumes");
    if (url) setCandForm({ ...candForm, resume_url: url });
  };

  return (
    <div>
      <SectionHeader title="🧑‍💼 নিয়োগ (Recruitment)" action="নতুন পদ" onAction={() => setShowJobModal(true)} />
      <div style={{ display: "grid", gridTemplateColumns: "260px 1fr", gap: 16 }}>
        <Card style={{ padding: 12, height: "fit-content" }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 10, padding: "0 6px" }}>পদসমূহ</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {jobs.map(j => (
              <div key={j.id} onClick={() => setSelectedJob(j)} style={{ padding: "10px 12px", borderRadius: 8, cursor: "pointer", background: selectedJob?.id === j.id ? C.primaryBg : "transparent", border: "1px solid " + (selectedJob?.id === j.id ? C.primaryLight : "transparent") }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.primaryDark }}>{j.title}</div>
                <div style={{ fontSize: 11, color: C.gray600 }}>{j.department} · {j.type}</div>
                <Badge label={j.status === "Open" ? "🟢 Open" : "🔴 Closed"} color={j.status === "Open" ? "green" : "red"} />
              </div>
            ))}
            {jobs.length === 0 && <div style={{ fontSize: 12, color: C.gray400, padding: 10 }}>কোনো পদ যোগ করা হয়নি</div>}
          </div>
        </Card>

        <div>
          {!selectedJob ? (
            <Card><div style={{ textAlign: "center", padding: 30, color: C.gray400, fontSize: 13 }}>একটা পদ বাছাই করুন</div></Card>
          ) : (
            <>
              <Card style={{ marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
                <div>
                  <div style={{ fontSize: 16, fontWeight: 700, color: C.primaryDark }}>{selectedJob.title}</div>
                  <div style={{ fontSize: 12, color: C.gray600 }}>{selectedJob.department} · {selectedJob.type} · {fmtNum(candidates.length)} জন আবেদন করেছে</div>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => toggleJobStatus(selectedJob)} style={btnEdit}>{selectedJob.status === "Open" ? "🔴 বন্ধ করুন" : "🟢 চালু করুন"}</button>
                  <button onClick={() => setShowCandModal(true)} style={{ ...btnPrimary, width: "auto", margin: 0 }}>➕ প্রার্থী যোগ করুন</button>
                </div>
              </Card>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12 }}>
                {CANDIDATE_STAGES.map(stage => (
                  <div key={stage.id}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: stage.color, marginBottom: 8, padding: "6px 10px", background: stage.color + "22", borderRadius: 8, textAlign: "center" }}>
                      {stage.label} ({candidates.filter(c => c.stage === stage.id).length})
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 8, minHeight: 60 }}>
                      {candidates.filter(c => c.stage === stage.id).map(c => (
                        <Card key={c.id} style={{ padding: 10 }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark }}>{c.name}</div>
                          <div style={{ fontSize: 11, color: C.gray600, marginBottom: 6 }}>{c.phone}</div>
                          {c.resume_url && <a href={c.resume_url} target="_blank" rel="noreferrer" style={{ fontSize: 11, color: C.blue, display: "block", marginBottom: 6 }}>📄 Resume</a>}
                          <select value={c.stage} onChange={e => moveStage(c, e.target.value)} style={{ ...inputStyle, padding: "4px 8px", fontSize: 11 }}>
                            {CANDIDATE_STAGES.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                          </select>
                        </Card>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </div>

      {showJobModal && (
        <Modal title="নতুন পদ পোস্ট করুন" onClose={() => setShowJobModal(false)}>
          <FormField label="পদের নাম *"><input style={inputStyle} value={jobForm.title} onChange={e => setJobForm({ ...jobForm, title: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="বিভাগ"><select style={inputStyle} value={jobForm.department} onChange={e => setJobForm({ ...jobForm, department: e.target.value })}>{["ডিজাইন", "নির্মাণ", "প্রশাসন", "বিপণন"].map(d => <option key={d}>{d}</option>)}</select></FormField>
            <FormField label="ধরন"><select style={inputStyle} value={jobForm.type} onChange={e => setJobForm({ ...jobForm, type: e.target.value })}>{["Full Time", "Part Time", "Contract", "Internship"].map(t => <option key={t}>{t}</option>)}</select></FormField>
          </div>
          <FormField label="বিবরণ"><textarea style={{ ...inputStyle, minHeight: 80 }} value={jobForm.description} onChange={e => setJobForm({ ...jobForm, description: e.target.value })} /></FormField>
          <button onClick={saveJob} style={btnPrimary}>✅ পোস্ট করুন</button>
        </Modal>
      )}

      {showCandModal && (
        <Modal title="নতুন প্রার্থী যোগ করুন" onClose={() => setShowCandModal(false)}>
          <FormField label="নাম *"><input style={inputStyle} value={candForm.name} onChange={e => setCandForm({ ...candForm, name: e.target.value })} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="ফোন"><input style={inputStyle} value={candForm.phone} onChange={e => setCandForm({ ...candForm, phone: e.target.value })} /></FormField>
            <FormField label="ইমেইল"><input style={inputStyle} value={candForm.email} onChange={e => setCandForm({ ...candForm, email: e.target.value })} /></FormField>
          </div>
          <FormField label="জীবনবৃত্তান্ত (Resume)">
            <input type="file" ref={candUploadRef} onChange={uploadResume} style={inputStyle} />
            {candForm.resume_url && <div style={{ fontSize: 11, color: C.green, marginTop: 4 }}>✅ আপলোড হয়েছে</div>}
          </FormField>
          <FormField label="নোট"><textarea style={{ ...inputStyle, minHeight: 60 }} value={candForm.notes} onChange={e => setCandForm({ ...candForm, notes: e.target.value })} /></FormField>
          <button onClick={saveCandidate} style={btnPrimary}>✅ যোগ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// HR REPORTS (All Employees Summary with date filters)
// ============================================================
function HRReports({ employees }) {
  const [range, setRange] = useState("today");
  const [customStart, setCustomStart] = useState(new Date().toISOString().split("T")[0]);
  const [customEnd, setCustomEnd] = useState(new Date().toISOString().split("T")[0]);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);

  const getRange = () => {
    const today = new Date();
    const todayStr = today.toISOString().split("T")[0];
    if (range === "today") return { start: todayStr, end: todayStr };
    if (range === "week") {
      const d = new Date(today); d.setDate(d.getDate() - 6);
      return { start: d.toISOString().split("T")[0], end: todayStr };
    }
    if (range === "month") {
      const d = new Date(today.getFullYear(), today.getMonth(), 1);
      return { start: d.toISOString().split("T")[0], end: todayStr };
    }
    return { start: customStart, end: customEnd };
  };

  const load = async () => {
    setLoading(true);
    const { start, end } = getRange();
    const { data } = await supabase.from("attendance").select("*").gte("date", start).lte("date", end);
    const byEmp = {};
    employees.forEach(e => { byEmp[e.id] = { employee: e, present: 0, absent: 0, late: 0, leave: 0 }; });
    (data || []).forEach(a => {
      const row = byEmp[a.employee_id]; if (!row) return;
      if (a.status === "উপস্থিত") row.present++;
      else if (a.status === "অনুপস্থিত") row.absent++;
      else if (a.status === "অর্ধদিন") row.late++;
      else if (a.status === "ছুটি") row.leave++;
    });
    setRows(Object.values(byEmp));
    setLoading(false);
  };
  useEffect(() => { load(); }, [range, customStart, customEnd]);

  const { start, end } = getRange();
  const handleExport = () => exportToExcel(rows.map(r => ({ নাম: r.employee.name, বিভাগ: r.employee.dept, উপস্থিত: r.present, অনুপস্থিত: r.absent, অর্ধদিন: r.late, ছুটি: r.leave })), "HR Report", "All Employees Summary");

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 15 }}>📊 All Employees Summary</div>
        <button onClick={handleExport} style={btnEdit}>⬇️ Export</button>
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
        {[["today", "আজ"], ["week", "এই সপ্তাহ"], ["month", "এই মাস"], ["custom", "কাস্টম"]].map(([id, label]) => (
          <button key={id} onClick={() => setRange(id)} style={{ background: range === id ? C.primary : C.white, color: range === id ? C.white : C.gray800, border: "1px solid " + (range === id ? C.primary : C.gray200), borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
        ))}
        {range === "custom" && (
          <>
            <input type="date" value={customStart} onChange={e => setCustomStart(e.target.value)} style={{ ...inputStyle, width: "auto" }} />
            <span style={{ alignSelf: "center", color: C.gray600 }}>→</span>
            <input type="date" value={customEnd} onChange={e => setCustomEnd(e.target.value)} style={{ ...inputStyle, width: "auto" }} />
          </>
        )}
      </div>
      <div style={{ fontSize: 12, color: C.gray400, marginBottom: 12 }}>{start} → {end}</div>

      <Card>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "বিভাগ", "উপস্থিত", "অনুপস্থিত", "অর্ধদিন", "ছুটি"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {rows.map(r => (
                <tr key={r.employee.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{r.employee.name}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={r.employee.dept} color="primary" /></td>
                  <td style={{ padding: "10px 14px", color: C.green, fontWeight: 700 }}>{fmtNum(r.present)}</td>
                  <td style={{ padding: "10px 14px", color: C.red, fontWeight: 700 }}>{fmtNum(r.absent)}</td>
                  <td style={{ padding: "10px 14px", color: "#856404", fontWeight: 700 }}>{fmtNum(r.late)}</td>
                  <td style={{ padding: "10px 14px", color: C.blue, fontWeight: 700 }}>{fmtNum(r.leave)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

// ============================================================
// PERFORMANCE MANAGEMENT
// ============================================================
function PerformanceManagement({ employees }) {
  const [reviews, setReviews] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState({ employee_id: "", period: "", goals: "" });
  const [reviewForm, setReviewForm] = useState({ manager_review: "", rating: 3, status: "Completed" });

  const load = async () => {
    setLoading(true);
    const { data } = await supabase.from("performance_reviews").select("*").order("created_at", { ascending: false });
    setReviews(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, []);

  const empById = {}; employees.forEach(e => { empById[String(e.id)] = e; });

  const createCycle = async () => {
    if (!form.employee_id || !form.period) return alert("কর্মী ও মেয়াদ আবশ্যক");
    await supabase.from("performance_reviews").insert([{ ...form, status: "Draft" }]);
    setShowModal(false); setForm({ employee_id: "", period: "", goals: "" }); load();
  };

  const openReview = (r) => { setEditItem(r); setReviewForm({ manager_review: r.manager_review || "", rating: r.rating || 3, status: "Completed" }); };
  const saveReview = async () => {
    await supabase.from("performance_reviews").update({ manager_review: reviewForm.manager_review, rating: +reviewForm.rating, status: reviewForm.status, reviewed_at: new Date().toISOString() }).eq("id", editItem.id);
    setEditItem(null); load();
  };

  const statusColor = { Draft: "gray", Submitted: "yellow", Completed: "green" };

  return (
    <div>
      <SectionHeader title="🎯 পারফরম্যান্স ম্যানেজমেন্ট" action="নতুন Review Cycle" onAction={() => setShowModal(true)} />
      <Card>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : reviews.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো Review Cycle তৈরি হয়নি</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "মেয়াদ", "লক্ষ্য (Goals)", "Self Review", "রেটিং", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {reviews.map(r => (
                <tr key={r.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{empById[String(r.employee_id)]?.name || "—"}</td>
                  <td style={{ padding: "10px 14px" }}>{r.period}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.goals}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.self_review || "—"}</td>
                  <td style={{ padding: "10px 14px" }}>{r.rating ? "⭐".repeat(r.rating) : "—"}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={r.status} color={statusColor[r.status] || "gray"} /></td>
                  <td style={{ padding: "10px 14px" }}><button onClick={() => openReview(r)} style={btnEdit}>✏️ Review</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>

      {showModal && (
        <Modal title="নতুন Review Cycle তৈরি করুন" onClose={() => setShowModal(false)}>
          <FormField label="কর্মী *">
            <select style={inputStyle} value={form.employee_id} onChange={e => setForm({ ...form, employee_id: e.target.value })}>
              <option value="">— বাছাই করুন —</option>
              {employees.map(e => <option key={e.id} value={e.id}>{e.name}</option>)}
            </select>
          </FormField>
          <FormField label="মেয়াদ *"><input style={inputStyle} value={form.period} onChange={e => setForm({ ...form, period: e.target.value })} placeholder="যেমন: 2026-Q3" /></FormField>
          <FormField label="লক্ষ্য (Goals)"><textarea style={{ ...inputStyle, minHeight: 80 }} value={form.goals} onChange={e => setForm({ ...form, goals: e.target.value })} /></FormField>
          <button onClick={createCycle} style={btnPrimary}>✅ তৈরি করুন</button>
        </Modal>
      )}

      {editItem && (
        <Modal title={"Review — " + (empById[String(editItem.employee_id)]?.name || "")} onClose={() => setEditItem(null)}>
          <div style={{ fontSize: 12, color: C.gray600, marginBottom: 10 }}><strong>লক্ষ্য:</strong> {editItem.goals || "—"}</div>
          <div style={{ fontSize: 12, color: C.gray600, marginBottom: 14 }}><strong>Self Review:</strong> {editItem.self_review || "কর্মী এখনো জমা দেননি"}</div>
          <FormField label="Manager Review"><textarea style={{ ...inputStyle, minHeight: 80 }} value={reviewForm.manager_review} onChange={e => setReviewForm({ ...reviewForm, manager_review: e.target.value })} /></FormField>
          <FormField label="রেটিং (১-৫)">
            <select style={inputStyle} value={reviewForm.rating} onChange={e => setReviewForm({ ...reviewForm, rating: e.target.value })}>
              {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{"⭐".repeat(n)} ({n})</option>)}
            </select>
          </FormField>
          <button onClick={saveReview} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// TIME TRACKER
// ============================================================
function TimeTracker({ employees }) {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [range, setRange] = useState({ start: new Date(new Date().setDate(new Date().getDate() - 6)).toISOString().split("T")[0], end: new Date().toISOString().split("T")[0] });

  const load = async () => {
    setLoading(true);
    const { data } = await supabase.from("time_logs").select("*").gte("date", range.start).lte("date", range.end).order("date", { ascending: false });
    setLogs(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, [range]);

  const empById = {}; employees.forEach(e => { empById[String(e.id)] = e; });
  const hoursOf = (l) => { if (!l.start_time || !l.end_time) return 0; const [sh, sm] = l.start_time.split(":").map(Number), [eh, em] = l.end_time.split(":").map(Number); return Math.max(((eh * 60 + em) - (sh * 60 + sm)) / 60, 0); };

  const byEmployee = {};
  logs.forEach(l => { const k = String(l.employee_id); byEmployee[k] = (byEmployee[k] || 0) + hoursOf(l); });
  const totalHours = logs.reduce((s, l) => s + hoursOf(l), 0);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard icon="⏲️" label="মোট লগ হওয়া ঘন্টা" value={fmtNum(totalHours.toFixed(1))} color={C.primaryBg} />
        <StatCard icon="👷" label="সক্রিয় কর্মী" value={fmtNum(Object.keys(byEmployee).length)} color={C.greenLight} />
      </div>

      <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 16, flexWrap: "wrap" }}>
        <input type="date" value={range.start} onChange={e => setRange({ ...range, start: e.target.value })} style={{ ...inputStyle, width: "auto" }} />
        <span style={{ color: C.gray600 }}>→</span>
        <input type="date" value={range.end} onChange={e => setRange({ ...range, end: e.target.value })} style={{ ...inputStyle, width: "auto" }} />
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 12, fontSize: 14 }}>কর্মী-ভিত্তিক মোট ঘন্টা</div>
        {Object.keys(byEmployee).length === 0 ? <div style={{ color: C.gray400, textAlign: "center", padding: 16, fontSize: 13 }}>কোনো লগ নেই</div> : (
          Object.entries(byEmployee).map(([empId, hrs]) => (
            <div key={empId} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid " + C.gray100, fontSize: 13 }}>
              <span style={{ color: C.gray800 }}>{empById[empId]?.name || "—"}</span>
              <span style={{ fontWeight: 700, color: C.primaryDark }}>{hrs.toFixed(1)} ঘন্টা</span>
            </div>
          ))
        )}
      </Card>

      <Card>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 12, fontSize: 14 }}>সব লগ</div>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : logs.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো লগ নেই</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "তারিখ", "কাজ", "সময়", "ঘন্টা", "নোট"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {logs.map(l => (
                <tr key={l.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{empById[String(l.employee_id)]?.name || "—"}</td>
                  <td style={{ padding: "10px 14px" }}>{l.date}</td>
                  <td style={{ padding: "10px 14px" }}>{l.task}{l.project_name ? " (" + l.project_name + ")" : ""}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12 }}>{l.start_time} - {l.end_time}</td>
                  <td style={{ padding: "10px 14px", fontWeight: 700 }}>{hoursOf(l).toFixed(1)}</td>
                  <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600 }}>{l.note}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

// ============================================================
// HR HELP DESK
// ============================================================
const TICKET_CATEGORIES = ["পে-রোল", "ছুটি", "IT সমস্যা", "সাধারণ", "অভিযোগ"];

function HRHelpDesk({ employees }) {
  const [tickets, setTickets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("open");
  const [editItem, setEditItem] = useState(null);
  const [reply, setReply] = useState("");
  const [status, setStatus] = useState("Resolved");

  const load = async () => {
    setLoading(true);
    const { data } = await supabase.from("hr_tickets").select("*").order("created_at", { ascending: false });
    setTickets(data || []);
    setLoading(false);
  };
  useEffect(() => { load(); }, []);

  const empById = {}; employees.forEach(e => { empById[String(e.id)] = e; });

  const openTicket = (t) => { setEditItem(t); setReply(t.admin_reply || ""); setStatus(t.status === "Open" ? "In Progress" : t.status); };
  const saveReply = async () => {
    await supabase.from("hr_tickets").update({ admin_reply: reply, status, resolved_at: (status === "Resolved" || status === "Closed") ? new Date().toISOString() : null }).eq("id", editItem.id);
    setEditItem(null); load();
  };

  const open = tickets.filter(t => t.status === "Open" || t.status === "In Progress");
  const resolved = tickets.filter(t => t.status === "Resolved" || t.status === "Closed");
  const list = tab === "open" ? open : resolved;
  const statusColor = { Open: "yellow", "In Progress": "primary", Resolved: "green", Closed: "gray" };

  return (
    <div>
      <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 14, fontSize: 15 }}>🎫 HR Help Desk</div>
      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {[["open", "🟡 খোলা (" + open.length + ")"], ["resolved", "✅ সমাধান হয়েছে"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{ background: tab === id ? C.primary : C.white, color: tab === id ? C.white : C.gray800, border: "1px solid " + (tab === id ? C.primary : C.gray200), borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{label}</button>
        ))}
      </div>
      <Card>
        {loading ? <div style={{ textAlign: "center", padding: 20, color: C.gray400 }}>⏳</div> : list.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো টিকেট নেই</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["কর্মী", "বিষয়", "Category", "স্ট্যাটাস", "তারিখ", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {list.map(t => (
                <tr key={t.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{empById[String(t.employee_id)]?.name || "—"}</td>
                  <td style={{ padding: "10px 14px", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.subject}</td>
                  <td style={{ padding: "10px 14px" }}><Badge label={t.category} color="primary" /></td>
                  <td style={{ padding: "10px 14px" }}><Badge label={t.status} color={statusColor[t.status] || "gray"} /></td>
                  <td style={{ padding: "10px 14px", fontSize: 12 }}>{new Date(t.created_at).toLocaleDateString("bn-BD")}</td>
                  <td style={{ padding: "10px 14px" }}><button onClick={() => openTicket(t)} style={btnEdit}>💬 উত্তর দিন</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>

      {editItem && (
        <Modal title={"টিকেট — " + editItem.subject} onClose={() => setEditItem(null)}>
          <div style={{ fontSize: 12, color: C.gray600, marginBottom: 4 }}><strong>কর্মী:</strong> {empById[String(editItem.employee_id)]?.name}</div>
          <div style={{ fontSize: 12, color: C.gray600, marginBottom: 14 }}><strong>বিবরণ:</strong> {editItem.description}</div>
          <FormField label="Admin Reply"><textarea style={{ ...inputStyle, minHeight: 80 }} value={reply} onChange={e => setReply(e.target.value)} /></FormField>
          <FormField label="স্ট্যাটাস">
            <select style={inputStyle} value={status} onChange={e => setStatus(e.target.value)}>
              <option value="In Progress">🔵 In Progress</option>
              <option value="Resolved">✅ Resolved</option>
              <option value="Closed">⚪ Closed</option>
            </select>
          </FormField>
          <button onClick={saveReply} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ============================================================
// HR SYSTEM HUB (single consolidated menu with subcategories)
// ============================================================
function HRSystemHub({ data, onRefresh, lang, currentUser, isAdmin }) {
  const visibleTabs = isAdmin ? HR_SUBTABS : HR_SUBTABS.filter(([id]) => (currentUser?.permissions || []).includes(id));
  const [sub, setSub] = useState(visibleTabs[0]?.[0] || "hr_employees");

  useEffect(() => {
    if (!visibleTabs.some(([id]) => id === sub)) setSub(visibleTabs[0]?.[0] || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isAdmin, currentUser?.permissions]);

  if (visibleTabs.length === 0) {
    return <Card><div style={{ textAlign: "center", padding: 30, color: C.gray600, fontSize: 13 }}>⚠️ আপনার এই সিস্টেমের কোনো অংশে access নেই। Admin-কে জানান।</div></Card>;
  }

  return (
    <div>
      <div style={{ marginBottom: 18 }}>
        <div style={{ fontSize: 18, fontWeight: 700, color: C.primaryDark }}>👥 HR ও পে-রোল সিস্টেম</div>
        <div style={{ fontSize: 12, color: C.gray400 }}>কর্মী, ছুটি, উপস্থিতি, পে-রোল ও নিয়োগ — সব একজায়গায়</div>
      </div>

      <div style={{ display: "flex", gap: 6, marginBottom: 20, flexWrap: "wrap", borderBottom: "1px solid " + C.gray100, paddingBottom: 12 }}>
        {visibleTabs.map(([id, icon, label]) => (
          <button key={id} onClick={() => setSub(id)} style={{ background: sub === id ? C.primary : C.white, color: sub === id ? C.white : C.gray800, border: "1px solid " + (sub === id ? C.primary : C.gray200), borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>{icon} {label}</button>
        ))}
      </div>

      {sub === "hr_employees" && <Employees data={data.employees} onRefresh={onRefresh} />}
      {sub === "hr_leave" && <LeaveManagement employees={data.employees} />}
      {sub === "hr_attendance" && <Attendance employees={data.employees} />}
      {sub === "hr_smart_attendance" && <SmartAttendance employees={data.employees} lang={lang} />}
      {sub === "hr_payroll" && <Payroll employees={data.employees} />}
      {sub === "hr_recruitment" && <Recruitment />}
      {sub === "hr_reports" && <HRReports employees={data.employees} />}
      {sub === "hr_performance" && <PerformanceManagement employees={data.employees} />}
      {sub === "hr_timetracker" && <TimeTracker employees={data.employees} />}
      {sub === "hr_helpdesk" && <HRHelpDesk employees={data.employees} />}
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
    alert("✅ " + count + "টি সামগ্রী আপলোড হয়েছে!"); onRefresh(); e.target.value = "";
  };

  return (
    <div>
      <SectionHeader title="সামগ্রী ও ইনভেন্টরি" action="নতুন সামগ্রী" onAction={openAdd} onExport={handleExport} onPrint={() => { printSection("সামগ্রী রিপোর্ট", "materials-content"); }} onUpload={handleUpload} uploadRef={uploadRef} />
      <Card>
        <div id="materials-content" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["সামগ্রী", "একক", "স্টক", "ন্যূনতম", "মূল্য/একক", "মোট মূল্য", "সাপ্লায়ার", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>{h}</th>)}</tr></thead>
            <tbody>
              {data.map(m => (
                <tr key={m.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
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
  const uploadRef = useRef();
  const handleUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      if (!row["কাজ"] && !row["work"]) continue;
      await supabase.from("site_progress").insert([{ project: row["প্রজেক্ট"] || row["project"] || "", date: row["তারিখ"] || row["date"] || new Date().toISOString().split("T")[0], work: row["কাজ"] || row["work"] || "", phase: row["পর্যায়"] || row["phase"] || "ফাউন্ডেশন", workers: +row["শ্রমিক"] || 0, note: row["নোট"] || "", status: row["স্ট্যাটাস"] || "চলমান" }]);
      count++;
    }
    alert("✅ " + count + "টি আপলোড হয়েছে!"); onRefresh(); e.target.value = "";
  };

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
      <SectionHeader title="সাইট অগ্রগতি" action="নতুন আপডেট" onAction={openAdd} onExport={handleExport} onPrint={() => { printSection("সাইট অগ্রগতি রিপোর্ট", "site-content"); }} onUpload={handleUpload} uploadRef={uploadRef} />
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
  const [officeSalaryCost, setOfficeSalaryCost] = useState(0);

  useEffect(() => {
    (async () => {
      const month = new Date().toISOString().slice(0, 7);
      const { data } = await supabase.from("payroll_runs").select("fixed_items, kpi_items, penalty_items").eq("month", month);
      const subtotalOf = (items) => (items || []).reduce((s, it) => s + (+it.amount || 0), 0);
      const total = (data || []).reduce((s, r) => s + subtotalOf(r.fixed_items) + subtotalOf(r.kpi_items) - subtotalOf(r.penalty_items), 0);
      setOfficeSalaryCost(total);
    })();
  }, []);

  return (
    <div>
      <h2 style={{ color: C.primaryDark, fontSize: 18, fontWeight: 700, marginBottom: 18 }}>ড্যাশবোর্ড</h2>
      <div id="dashboard-content">
        <div style={{ marginBottom: 20, padding: "18px 22px", background: "linear-gradient(135deg, " + C.primaryDark + ", " + C.primary + ")", borderRadius: 14, color: C.white }}>
          <div style={{ fontSize: 13, opacity: 0.8 }}>স্বাগতম,</div>
          <div style={{ fontSize: 20, fontWeight: 700, marginTop: 4 }}>Noksha Interior & Construction</div>
          <div style={{ fontSize: 12, opacity: 0.7, marginTop: 4 }}>আজকের তারিখ: {new Date().toLocaleDateString("bn-BD")} | Supabase Connected ✅</div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 20 }}>
          <StatCard icon="🏗️" label="সক্রিয় প্রজেক্ট" value={fmtNum(activeProjects)} sub={"মোট " + projects.length + "টি"} color="#E8F5E9" />
          <StatCard icon="👥" label="মোট ক্লায়েন্ট" value={fmtNum(clients.length)} color="#F0FFF4" />
          <StatCard icon="👷" label="কর্মী সংখ্যা" value={fmtNum(employees.length)} color="#FFF8E1" />
          <StatCard icon="💰" label="মোট আয়" value={fmt(totalIncome)} color="#F0FFF4" />
          <StatCard icon="💸" label="মোট ব্যয়" value={fmt(totalExpense)} color="#FFF5F5" />
          <StatCard icon="🧾" label="Office Salary Staff (এই মাস)" value={fmt(officeSalaryCost)} color="#FFF8E1" />
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
              <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: "1px solid " + C.gray100 }}>
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
                <div key={m.id} style={{ background: C.redLight, border: "1px solid #F5C6CB", borderRadius: 10, padding: "10px 14px" }}>
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
const genSerialNo = () => {
  const d = new Date();
  const mon = d.toLocaleDateString("en-GB", { month: "short" }).toUpperCase();
  return "NIC" + String(d.getFullYear()).slice(-2) + mon + String(d.getDate()).padStart(2, "0");
};

const BOQ_DEFAULT_EXCLUSIONS = [
  "Loose Furniture & Seating: Living & Dining: Sofa sets, Center tables, Side tables (unless specified in BOQ), Dining table, and Dining chairs. Study & Bedroom: Study chairs, Relaxing chairs, Ottoman/Bench (at the foot of the bed), and any movable seating. Note: Only fixed woodwork (Wardrobes, TV Units, Cabinets) explicitly mentioned in the item list is included.",
  "Electronics & Appliances: Heavy Appliances: Air Conditioners (AC), Televisions (TV), Refrigerators, Washing Machines, and Ceiling Fans. Kitchen Gadgets: Built-in Microwave Oven, Baking Oven, Dishwasher, Rice Cooker, Coffee Machine, Toaster, etc. Note: We have provided the Kitchen Hood, Hob (Stove), Sink, and RO Filter only. All other electronics shown in the 3D design are for visualization purposes.",
  "Civil & Sanitary Exclusions (Cost Reduction Items): Bathtub: The bathtub shown in the master bathroom design is excluded (replaced with a standard glass partition/shower area). Imported Tiles: Quoted rates are based on Premium Local Brands (Akij/RAK/DBL). High-end Imported (China/Italian) tiles are excluded. Sanitary Ware: Premium imported sanitary fixtures are excluded. Rates are based on high-quality local brands (Sattar/RFL/Stella) or equivalent standard series.",
  "Soft Furnishing & Bedding: Bedding: Sleeping Mattresses (Toshok), Pillows, Bed Sheets, Blankets, and Cushions. Curtains & Carpets: We provide the Window Pelmet/Curtain Box only. The Curtain Fabrics, Rails/Rods, Blinds, and Floor Rugs/Carpets are to be purchased by the client.",
  "Decor & Miscellaneous: Accessories: Wall Paintings, Photo Frames, Table Lamps, Flower Vases, Books, and Decorative Plants. Window Blinds: Venetian or Roller blinds are excluded unless a separate line item is added.",
  "Paint & Finish Clarification: Paint Specification: All woodwork estimates are based on Premium Enamel / High-Gloss Spray Paint or Hand Polish. Automotive-grade \"Duco Paint\" is excluded to maintain the quoted budget.",
].join("\n");

const BOQ_DEFAULT_CLEARANCE = [
  "The quoted amount is only for the above-mentioned work. If any extra work is needed apart from the quoted work above, it will be adjusted/updated on the final bill.",
  "Civil, plumbing and sanitary, 3 marble/granite, exterior, sound system, LAN/WIFI, AC, any other Labour type work are not included in our inclusion. If required during the service period, changes will be added. We will follow only 3d rendering design.",
  "We are fully dependent on market suppliers. If the raw materials are not available/supply stockout on the market, complementary materials of similar quality and price shall be used to continue work.",
].join("\n");

const BOQ_DEFAULT_TERMS = [
  "Delivery/Handover Delay: Delivery date may change for any unavoidable circumstances.",
  "Fitting: For Knockdown Furnitures, the fitting team may go the same day or the next day.",
  "Mode of Payment: Payments are accepted through Bank Transfer/Cash on Delivery/Bkash/Nagad. NOTE: BKASH/NAGAD CHARGES ARE APPLICABLE.",
  "Labour and Materials Supply: Fully dependent on Project planning, work progress, and other related Normal technical Subjects.",
  "Return-Refund Policy: No Refund/Return Policy after order confirmation.",
  "Warranty: 2 years free service Warranty on product frame/loan/manufacturing fault. No warranty/guarantee for Wood, glass, fabric, PU Leather, and damage from rat use. 12 Years Hardware Parts Warranty-Hinge/Kobja, Chain, Channel, SS/MS Item, Handle, Or equivalent.",
  "Changing: No change is allowed during the installation period without charge.",
  "Delivery Time: 120 days.",
  "All taxes and Govt. duties are not applicable to this bill.",
].join("\n");

const BOQ_DEFAULT_PAYMENT_METHODS = [
  "You can pay by Cash/Cheque/Bank Transfer/BKash/Nagad.",
  "Bank Account Name: Noksha Interior & Construction",
  "Bank Account Number: 7862141003970224",
  "Bank Name: United Commercial Bank (UCB Bank)",
  "Branch Name: Faridpur.",
].join("\n");

const BOQ_DEFAULT_PAYMENT_TERMS = [
  "1st Payment = 50% advance payment is required before starting the project.",
  "2nd Payment = 40% payment before paint work.",
  "3rd Payment = Remaining 10% payment at handover day.",
].join("\n");

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
  const [showEditProjModal, setShowEditProjModal] = useState(false);
  const [showItemModal, setShowItemModal] = useState(false);
  const [showMultiItemModal, setShowMultiItemModal] = useState(false);
  const [editItem, setEditItem] = useState(null);

  useEffect(() => { loadProjects(); loadStdRates(); }, []);
  useEffect(() => { if (selProj) { loadBOQ(selProj); loadExpenses(selProj); loadSettings(selProj); } }, [selProj]);

  const loadProjects = async () => { const { data } = await supabase.from("project_settings").select("*").order("created_at", { ascending: false }); setProjects(data || []); };
  const loadSettings = async (pid) => { const { data } = await supabase.from("project_settings").select("*").eq("project_id", pid).single(); setSettings(data); };
  const loadBOQ = async (pid) => { setLoading(true); const { data } = await supabase.from("project_boq").select("*").eq("project_id", pid).order("id"); setBoqItems(data || []); setLoading(false); };
  const loadExpenses = async (pid) => { const { data } = await supabase.from("project_expenses").select("*").eq("project_id", pid).order("expense_date"); setExpenses(data || []); };
  const loadStdRates = async () => { const { data } = await supabase.from("standard_rates").select("*").order("category"); setStdRates(data || []); };

  const sortedBoqItems = [...boqItems].sort((a, b) => (a.sort_order ?? a.id ?? 0) - (b.sort_order ?? b.id ?? 0));
  const roomGroups = sortedBoqItems.reduce((acc, item) => { if (!acc[item.room_name]) acc[item.room_name] = []; acc[item.room_name].push(item); return acc; }, {});

  const moveItemInRoom = async (room, index, dir) => {
    const roomItems = roomGroups[room];
    const targetIndex = index + dir;
    if (targetIndex < 0 || targetIndex >= roomItems.length) return;
    const a = roomItems[index], b = roomItems[targetIndex];
    const aOrder = a.sort_order ?? a.id, bOrder = b.sort_order ?? b.id;
    await Promise.all([
      supabase.from("project_boq").update({ sort_order: bOrder }).eq("id", a.id),
      supabase.from("project_boq").update({ sort_order: aOrder }).eq("id", b.id),
    ]);
    await loadBOQ(selProj);
  };

  const moveRoom = async (room, dir) => {
    const roomNames = Object.keys(roomGroups);
    const idx = roomNames.indexOf(room);
    const targetIdx = idx + dir;
    if (targetIdx < 0 || targetIdx >= roomNames.length) return;
    const newOrder = [...roomNames];
    [newOrder[idx], newOrder[targetIdx]] = [newOrder[targetIdx], newOrder[idx]];
    let counter = 0;
    const updates = [];
    newOrder.forEach(r => {
      roomGroups[r].forEach(it => {
        updates.push(supabase.from("project_boq").update({ sort_order: counter }).eq("id", it.id));
        counter++;
      });
    });
    await Promise.all(updates);
    await loadBOQ(selProj);
  };
  const grandTotal = boqItems.reduce((s, i) => s + Number(i.amount || 0), 0);
  const deliveryCharge = Number(settings?.delivery_charge || 0);
  const subTotal = grandTotal + deliveryCharge;
  const totalExpenses = expenses.reduce((s, e) => s + Number(e.amount || 0), 0);
  const netProfit = subTotal - totalExpenses;

  const saveProject = async (form) => {
    const pid = genProjId();
    await supabase.from("project_settings").insert({
      project_id: pid, serial_no: genSerialNo(), project_name: form.project_name, client_name: form.client_name, client_address: form.client_address, client_phone: form.client_phone, delivery_charge: Number(form.delivery_charge) || 0,
      exclusions: form.exclusions ? form.exclusions.split("\n").filter(Boolean) : [],
      clearance: form.clearance ? form.clearance.split("\n").filter(Boolean) : [],
      terms_conditions: form.terms_conditions ? form.terms_conditions.split("\n").filter(Boolean) : [],
      payment_methods: form.payment_methods ? form.payment_methods.split("\n").filter(Boolean) : [],
      payment_terms_policy: form.payment_terms_policy ? form.payment_terms_policy.split("\n").filter(Boolean) : [],
    });
    await loadProjects(); setSelProj(pid); setShowProjModal(false);
  };

  const updateProjectSettings = async (form) => {
    const { error } = await supabase.from("project_settings").update({
      project_name: form.project_name, client_name: form.client_name, client_address: form.client_address, client_phone: form.client_phone, delivery_charge: Number(form.delivery_charge) || 0,
      exclusions: form.exclusions ? form.exclusions.split("\n").filter(Boolean) : [],
      clearance: form.clearance ? form.clearance.split("\n").filter(Boolean) : [],
      terms_conditions: form.terms_conditions ? form.terms_conditions.split("\n").filter(Boolean) : [],
      payment_methods: form.payment_methods ? form.payment_methods.split("\n").filter(Boolean) : [],
      payment_terms_policy: form.payment_terms_policy ? form.payment_terms_policy.split("\n").filter(Boolean) : [],
    }).eq("project_id", selProj);
    if (error) return alert("❌ সংরক্ষণ ব্যর্থ: " + error.message);
    await loadSettings(selProj); await loadProjects();
    setShowEditProjModal(false);
  };

  const saveBOQItem = async (form) => {
    if (!selProj) return alert("আগে একটি Project select করুন!");
    const qty = Number(form.qty) || 0;
    const rate = Number(form.rate) || 0;
    const amount = qty * rate;
    const maxSort = boqItems.reduce((m, it) => Math.max(m, it.sort_order ?? 0), 0);
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
    if (!editItem) payload.sort_order = maxSort + 1;
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

  const saveBOQItemsBulk = async (rows) => {
    if (!selProj) return alert("আগে একটি Project select করুন!");
    const valid = rows.filter(r => r.item_name && r.rate);
    if (valid.length === 0) return alert("কমপক্ষে একটি সারিতে Item Name ও Rate দিন!");
    const maxSort = boqItems.reduce((m, it) => Math.max(m, it.sort_order ?? 0), 0);
    const payloads = valid.map((r, i) => {
      const qty = Number(r.qty) || 0, rate = Number(r.rate) || 0;
      return {
        project_id: selProj, room_name: r.room_name || "Master Bedroom", code_no: r.code_no || "", item_no: Number(r.item_no) || 1,
        item_name: r.item_name || "", work_description: r.work_description || "", specification: r.specification || "",
        unit: r.unit || "sft", qty, rate, amount: qty * rate, is_rate_fixed: false, sort_order: maxSort + i + 1,
      };
    });
    const { error } = await supabase.from("project_boq").insert(payloads);
    if (error) return alert("❌ সংরক্ষণ ব্যর্থ: " + error.message);
    await loadBOQ(selProj);
    setShowMultiItemModal(false);
  };

  const saveExpense = async (form) => { if (!selProj) return alert("আগে Project select করুন!"); const qty2 = Number(form.qty) || 1; const rate2 = Number(form.rate) || 0; const r = await supabase.from("project_expenses").insert([{ project_id: selProj, expense_date: form.expense_date || new Date().toISOString().split("T")[0], item_name: form.item_name || "", description: form.description || "", qty: qty2, rate: rate2, amount: qty2 * rate2, category: form.category || "material" }]); if (r.error) { alert("Error: " + r.error.message); return; } await loadExpenses(selProj); };
  const deleteExpense = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("project_expenses").delete().eq("id", id); await loadExpenses(selProj); };

  const updateDelivery = async (val) => { await supabase.from("project_settings").update({ delivery_charge: Number(val) || 0 }).eq("project_id", selProj); await loadSettings(selProj); };

  const handlePrint = () => {
    printSection("BILL OF QUANTITIES (BOQ)", "boq-print-area", new Date());
  };

  const thS = { padding: "8px 10px", border: "1px solid #ddd", background: "#3F5F45", color: "#fff", fontSize: 12 };
  const tdS = { padding: "7px 10px", border: "1px solid #eee", fontSize: 13, textAlign: "center" };

  return (
    <div>
      {/* Header row */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, color: C.primaryDark, fontSize: 18, fontWeight: 700 }}>📋 Estimate Project</h2>
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
          {selProj && settings && <button onClick={() => setShowEditProjModal(true)} style={{ ...btnEdit, marginLeft: "auto" }}>✏️ Project Settings (নাম/Exclusions/Terms ইত্যাদি)</button>}
        </div>
      </Card>

      {/* Tabs - always visible, Item Library doesn't need a project */}
      <div style={{ display: "flex", borderBottom: "2px solid " + C.primary, marginBottom: 20 }}>
        {[{ k: "boq", l: "📋 BOQ" }, { k: "expenses", l: "💸 Daily Expenses" }, { k: "compare", l: "📊 Profit Analysis" }, { k: "items", l: "📚 Item Library" }].map(t => (
          <button key={t.k} onClick={() => setTab(t.k)} style={{ padding: "10px 20px", border: "none", borderBottom: tab === t.k ? "3px solid #C9A84C" : "3px solid transparent", background: "none", color: tab === t.k ? C.primaryDark : C.gray600, fontWeight: tab === t.k ? 700 : 400, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>{t.l}</button>
        ))}
      </div>

      {/* ITEM LIBRARY TAB — works without a project */}
      {tab === "items" && <BOQItemLibrary stdRates={stdRates} onRefresh={loadStdRates} />}

      {tab !== "items" && (!selProj ? (
        <Card style={{ textAlign: "center", padding: "60px 20px" }}>
          <div style={{ fontSize: 48, marginBottom: 12 }}>📋</div>
          <div style={{ color: C.primaryDark, fontWeight: 700, fontSize: 16, marginBottom: 8 }}>BOQ Management System</div>
          <div style={{ color: C.gray600, fontSize: 13 }}>{projects.length > 0 ? projects.length + "টি project আছে। উপর থেকে বেছে নিন।" : "শুরু করতে নতুন project তৈরি করুন।"}</div>
        </Card>
      ) : (
        <>
          {/* BOQ TAB */}
          {tab === "boq" && (
            <div>
              <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
                <button onClick={() => { setEditItem(null); setShowItemModal(true); }} style={{ background: C.primary, color: "#fff", border: "none", borderRadius: 8, padding: "8px 16px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}>+ Item যোগ করুন</button>
                <button onClick={() => setShowMultiItemModal(true)} style={{ background: "#2A5C8F", color: "#fff", border: "none", borderRadius: 8, padding: "8px 16px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}>+ একসাথে একাধিক Item</button>
                <button onClick={handlePrint} style={{ background: C.blue, color: "#fff", border: "none", borderRadius: 8, padding: "8px 16px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}>🖨️ Print / PDF</button>
              </div>

              {loading ? <div style={{ textAlign: "center", padding: 40, color: C.gray400 }}>লোড হচ্ছে...</div> : (
                <div id="boq-print-area">
                  <div style={{ marginBottom: 18 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                      <div style={{ fontStyle: "italic", fontWeight: 700, fontSize: 15 }}>PRELIMINARY QUOTATION FOR INTERIOR WORK (Flat)</div>
                      <div style={{ fontSize: 11, textAlign: "right" }}>
                        <div style={{ fontWeight: 700 }}>SL: {settings?.serial_no || genSerialNo()}</div>
                      </div>
                    </div>
                    <div style={{ fontSize: 12, lineHeight: 1.8 }}>
                      <div>To,</div>
                      <div>Client Name: {settings?.client_name || ""}</div>
                      <div>Project Location: {settings?.client_address || ""}</div>
                      <div style={{ marginTop: 8 }}>Dear Sir,</div>
                      <div>We are looking forward to receiving your response to discuss the design and specification.</div>
                    </div>
                  </div>
                  {Object.keys(roomGroups).length === 0 ? (
                    <Card style={{ textAlign: "center", padding: 50 }}><div style={{ fontSize: 36, marginBottom: 10 }}>📋</div><div style={{ color: C.gray400 }}>কোনো item নেই। "+ Item যোগ করুন" ক্লিক করুন।</div></Card>
                  ) : (
                    Object.entries(roomGroups).map(([room, items], roomIdx) => {
                      const roomTotal = items.reduce((s, i) => s + Number(i.amount || 0), 0);
                      const roomCount = Object.keys(roomGroups).length;
                      return (
                        <div key={room} style={{ marginBottom: 20 }}>
                          <div className="room-header" style={{ background: C.primary, color: "#fff", padding: "8px 16px", fontWeight: 700, borderRadius: "4px 4px 0 0", fontSize: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <span>{room}</span>
                            <span className="no-print" style={{ display: "flex", gap: 4 }}>
                              <button onClick={() => moveRoom(room, -1)} disabled={roomIdx === 0} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "#fff", cursor: roomIdx === 0 ? "default" : "pointer", opacity: roomIdx === 0 ? 0.4 : 1, borderRadius: 4, width: 22, height: 22, fontSize: 12 }}>▲</button>
                              <button onClick={() => moveRoom(room, 1)} disabled={roomIdx === roomCount - 1} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "#fff", cursor: roomIdx === roomCount - 1 ? "default" : "pointer", opacity: roomIdx === roomCount - 1 ? 0.4 : 1, borderRadius: 4, width: 22, height: 22, fontSize: 12 }}>▼</button>
                            </span>
                          </div>
                          <div style={{ overflowX: "auto", background: "#fff", border: "1px solid #eee", borderTop: "none" }}>
                            <table style={{ width: "100%", borderCollapse: "collapse" }}>
                              <thead>
                                <tr>
                                  {["", "Code No", "Item No", "Work Description & Specification", "Unit", "Qty", "Rate (৳)", "Amount (৳)", ""].map((h, i) => (
                                    <th key={i} className={(i === 0 || i === 8) ? "no-print" : undefined} style={{ ...thS, textAlign: i === 3 ? "left" : "center", minWidth: i === 3 ? 200 : 60 }}>{h}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {items.map((item, idx) => (
                                  <tr key={item.id} style={{ background: idx % 2 === 0 ? "#fff" : C.gray50 }}>
                                    <td style={{ ...tdS, width: 46 }} className="no-print">
                                      <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
                                        <button onClick={() => moveItemInRoom(room, idx, -1)} disabled={idx === 0} style={{ background: "none", border: "none", cursor: idx === 0 ? "default" : "pointer", color: idx === 0 ? C.gray200 : C.gray600, fontSize: 11, lineHeight: 1, padding: 2 }}>▲</button>
                                        <button onClick={() => moveItemInRoom(room, idx, 1)} disabled={idx === items.length - 1} style={{ background: "none", border: "none", cursor: idx === items.length - 1 ? "default" : "pointer", color: idx === items.length - 1 ? C.gray200 : C.gray600, fontSize: 11, lineHeight: 1, padding: 2 }}>▼</button>
                                      </div>
                                    </td>
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
                                  <td colSpan={7} style={{ ...tdS, textAlign: "right", fontWeight: 700 }}>Sub Total ({room}):</td>
                                  <td style={{ ...tdS, fontWeight: 700, color: C.primaryDark }}>৳ {fmtBOQ(roomTotal)}</td>
                                  <td className="no-print" />
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
                            <tr style={{ borderTop: "2px solid " + C.primary }}>
                              <td style={{ padding: "10px 0", fontWeight: 700, color: C.primaryDark, fontSize: 16 }}>Sub Total:</td>
                              <td style={{ textAlign: "right", fontWeight: 700, color: C.primaryDark, fontSize: 16 }}>৳ {fmtBOQ(subTotal)}</td>
                            </tr>
                            <tr><td colSpan={2} style={{ padding: "6px 0 0", fontSize: 11, fontStyle: "italic", color: C.gray600 }}>In Word: {numToWordsTaka(subTotal)}</td></tr>
                          </tbody>
                        </table>
                        {/* Payment Terms */}
                        <div style={{ marginTop: 14, paddingTop: 12, borderTop: "1px solid " + C.gray200 }}>
                          <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 8, fontSize: 13 }}>Payment Terms:</div>
                          {[[settings?.payment_term_1 || 50, "কাজ শুরুর আগে"], [settings?.payment_term_2 || 40, "কাজ চলাকালীন"], [settings?.payment_term_3 || 10, "কাজ সম্পন্নে"]].map(([pct, label], i) => (
                            <div key={i} style={{ fontSize: 13, marginBottom: 4 }}>• {i + 1}ম কিস্তি ({label}): <strong>{pct}%</strong> = ৳ {fmtBOQ(subTotal * pct / 100)}</div>
                          ))}
                        </div>
                      </Card>

                      {/* Exclusions */}
                      {settings?.exclusions?.length > 0 && (
                        <Card style={{ marginTop: 16, background: "#fff8e1", border: "1px solid #C9A84C" }}>
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

                      {/* Clearance */}
                      {settings?.clearance?.length > 0 && (
                        <Card style={{ marginTop: 12 }}>
                          <div style={{ fontWeight: 700, marginBottom: 8, color: C.primaryDark }}>CLEARANCE:</div>
                          {settings.clearance.map((c, i) => <div key={i} style={{ fontSize: 12, marginBottom: 3 }}>• {c}</div>)}
                        </Card>
                      )}

                      {/* Payment Methods */}
                      {settings?.payment_methods?.length > 0 && (
                        <Card style={{ marginTop: 12 }}>
                          <div style={{ fontWeight: 700, marginBottom: 8, color: C.primaryDark }}>PAYMENT METHODS:</div>
                          {settings.payment_methods.map((p, i) => <div key={i} style={{ fontSize: 12, marginBottom: 3 }}>{i === 0 ? p : "• " + p}</div>)}
                        </Card>
                      )}

                      {/* Payment Terms & Policy */}
                      {settings?.payment_terms_policy?.length > 0 && (
                        <Card style={{ marginTop: 12 }}>
                          <div style={{ fontWeight: 700, marginBottom: 8, color: C.primaryDark }}>Payment Terms & Policy:</div>
                          {settings.payment_terms_policy.map((p, i) => <div key={i} style={{ fontSize: 12, marginBottom: 3 }}>• {p}</div>)}
                        </Card>
                      )}

                      {/* Signatures */}
                      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 40, paddingTop: 14, borderTop: "1px solid " + C.gray200 }}>
                        <div style={{ textAlign: "center" }}><div style={{ borderTop: "1px solid " + C.gray800, paddingTop: 4, width: 150, fontSize: 12 }}>Client Signature</div></div>
                        <div style={{ textAlign: "center" }}><div style={{ borderTop: "1px solid " + C.gray800, paddingTop: 4, width: 150, fontSize: 12 }}>Authorized by NIC</div></div>
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
      ))}

      {/* Modals */}
      {showProjModal && <BOQProjectModal onSave={saveProject} onClose={() => setShowProjModal(false)} />}
      {showEditProjModal && settings && <BOQProjectModal existing={settings} onSave={updateProjectSettings} onClose={() => setShowEditProjModal(false)} />}
      {showItemModal && <BOQItemModal item={editItem} stdRates={stdRates} existingRooms={Object.keys(roomGroups)} onSave={saveBOQItem} onClose={() => { setShowItemModal(false); setEditItem(null); }} />}
      {showMultiItemModal && <BOQMultiItemModal stdRates={stdRates} existingRooms={Object.keys(roomGroups)} onSave={saveBOQItemsBulk} onClose={() => setShowMultiItemModal(false)} />}
    </div>
  );
}

function BOQDeliveryEdit({ value, onSave }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(value);
  useEffect(() => { setVal(value); }, [value]);
  if (editing) return (
    <span style={{ display: "flex", gap: 6, justifyContent: "flex-end", alignItems: "center" }}>
      <input type="number" value={val} onChange={e => setVal(e.target.value)} style={{ width: 90, padding: "3px 6px", border: "1px solid " + C.gray200, borderRadius: 4 }} />
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
            <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "Item", "বিবরণ", "Category", "Qty", "Rate", "Amount", ""].map(h => <th key={h} style={{ padding: "10px 12px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>{h}</th>)}</tr></thead>
            <tbody>
              {expenses.length === 0 ? <tr><td colSpan={8} style={{ textAlign: "center", padding: 40, color: C.gray400 }}>কোনো expense নেই</td></tr> : expenses.map((e, i) => (
                <tr key={e.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={ev => ev.currentTarget.style.background = C.primaryBg} onMouseLeave={ev => ev.currentTarget.style.background = "transparent"}>
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
        <StatCard icon="📋" label="BOQ Contract Value" value={"৳ " + fmtBOQ(subTotal)} color={C.primaryBg} />
        <StatCard icon="💸" label="মোট Actual Expense" value={"৳ " + fmtBOQ(totalExpenses)} color="#FFEBEE" />
        <StatCard icon={isProfit ? "📈" : "📉"} label={"Net " + (isProfit ? "Profit" : "Loss") + " (" + profitPct + "%)"} value={"৳ " + fmtBOQ(Math.abs(netProfit))} color={isProfit ? "#E8F5E9" : "#FFEBEE"} />
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 12 }}>Expense vs Revenue</div>
        <div style={{ background: C.gray100, borderRadius: 10, height: 32, overflow: "hidden", position: "relative" }}>
          <div style={{ width: expensePct + "%", height: "100%", background: C.red, borderRadius: 10, display: "flex", alignItems: "center", paddingLeft: 10, color: "#fff", fontSize: 12, fontWeight: 700, transition: "width 0.5s", minWidth: expensePct > 0 ? 80 : 0 }}>
            {expensePct > 5 && "Expense: " + expensePct.toFixed(1) + "%"}
          </div>
        </div>
        <div style={{ fontSize: 12, color: C.gray600, marginTop: 6 }}>Profit margin: <strong style={{ color: isProfit ? C.green : C.red }}>{profitPct}%</strong></div>
      </Card>

      <Card>
        <div style={{ fontWeight: 700, color: C.primaryDark, marginBottom: 12 }}>Room-wise BOQ Breakdown</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead><tr style={{ background: C.primaryBg }}>{["Room / Area", "BOQ Amount (৳)", "% of Total"].map(h => <th key={h} style={{ padding: "9px 12px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>{h}</th>)}</tr></thead>
          <tbody>
            {Object.entries(roomGroups).map(([room, items]) => {
              const rt = items.reduce((s, i) => s + Number(i.amount || 0), 0);
              const pct = grandTotal > 0 ? ((rt / grandTotal) * 100).toFixed(1) : 0;
              return (
                <tr key={room} style={{ borderBottom: "1px solid " + C.gray100 }}>
                  <td style={{ padding: "9px 12px", fontWeight: 600, color: C.primaryDark }}>{room}</td>
                  <td style={{ padding: "9px 12px", fontWeight: 700 }}>৳ {fmtBOQ(rt)}</td>
                  <td style={{ padding: "9px 12px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ flex: 1, background: C.gray100, borderRadius: 4, height: 8 }}><div style={{ width: pct + "%", height: "100%", background: C.primary, borderRadius: 4 }} /></div>
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

function BOQProjectModal({ onSave, onClose, existing }) {
  const arrToText = (a) => Array.isArray(a) ? a.join("\n") : (a || "");
  const [form, setForm] = useState(existing ? {
    project_name: existing.project_name || "", client_name: existing.client_name || "", client_address: existing.client_address || "", client_phone: existing.client_phone || "", delivery_charge: existing.delivery_charge || 0,
    exclusions: arrToText(existing.exclusions) || BOQ_DEFAULT_EXCLUSIONS,
    clearance: arrToText(existing.clearance) || BOQ_DEFAULT_CLEARANCE,
    terms_conditions: arrToText(existing.terms_conditions) || BOQ_DEFAULT_TERMS,
    payment_methods: arrToText(existing.payment_methods) || BOQ_DEFAULT_PAYMENT_METHODS,
    payment_terms_policy: arrToText(existing.payment_terms_policy) || BOQ_DEFAULT_PAYMENT_TERMS,
  } : {
    project_name: "", client_name: "", client_address: "", client_phone: "", delivery_charge: 0,
    exclusions: BOQ_DEFAULT_EXCLUSIONS,
    clearance: BOQ_DEFAULT_CLEARANCE,
    terms_conditions: BOQ_DEFAULT_TERMS,
    payment_methods: BOQ_DEFAULT_PAYMENT_METHODS,
    payment_terms_policy: BOQ_DEFAULT_PAYMENT_TERMS,
  });
  const serialPreview = existing?.serial_no || genSerialNo();
  return (
    <Modal title={existing ? "✏️ Project Settings সম্পাদনা" : "নতুন BOQ Project"} onClose={onClose} size={650}>
      <div style={{ background: C.primaryBg, borderRadius: 8, padding: "8px 12px", marginBottom: 14, fontSize: 12, color: C.primaryDark }}>
        📌 Serial No: <strong>{serialPreview}</strong>
      </div>
      <FormField label="Project Name * (rename করতে পারবেন)"><input style={inputStyle} value={form.project_name} onChange={e => setForm({ ...form, project_name: e.target.value })} placeholder="যেমন: Rahman Villa Interior" /></FormField>
      <FormField label="Client Name *"><input style={inputStyle} value={form.client_name} onChange={e => setForm({ ...form, client_name: e.target.value })} placeholder="Mr. Opu Sir (Managing Director of ...)" /></FormField>
      <FormField label="Project Location / Address"><input style={inputStyle} value={form.client_address} onChange={e => setForm({ ...form, client_address: e.target.value })} placeholder="Sunrise train school, rathkhola, Faridpur sadar, Faridpur." /></FormField>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <FormField label="Client Phone"><input style={inputStyle} value={form.client_phone} onChange={e => setForm({ ...form, client_phone: e.target.value })} /></FormField>
        <FormField label="Delivery Charge (৳)"><input style={inputStyle} type="number" value={form.delivery_charge} onChange={e => setForm({ ...form, delivery_charge: e.target.value })} /></FormField>
      </div>

      <div style={{ fontSize: 12, color: C.gray400, margin: "14px 0 10px", borderTop: "1px solid " + C.gray100, paddingTop: 10 }}>{existing ? "নিচের সব অংশ এই Project-এর জন্য edit করুন — Save করলেই print-এ আপডেট হয়ে যাবে।" : "নিচের সব অংশ default হিসেবে ভরা আছে (Noksha-এর standard বিবরণ) — চাইলে edit করুন, প্রতিটা নতুন Project-এ এটাই default থাকবে।"}</div>

      <FormField label="Project Exclusions (প্রতি লাইনে একটি ক্যাটাগরি)"><textarea style={{ ...inputStyle, height: 100, resize: "vertical", fontSize: 11 }} value={form.exclusions} onChange={e => setForm({ ...form, exclusions: e.target.value })} /></FormField>
      <FormField label="Clearance (প্রতি লাইনে একটি)"><textarea style={{ ...inputStyle, height: 80, resize: "vertical", fontSize: 11 }} value={form.clearance} onChange={e => setForm({ ...form, clearance: e.target.value })} /></FormField>
      <FormField label="Terms & Conditions (প্রতি লাইনে একটি)"><textarea style={{ ...inputStyle, height: 100, resize: "vertical", fontSize: 11 }} value={form.terms_conditions} onChange={e => setForm({ ...form, terms_conditions: e.target.value })} /></FormField>
      <FormField label="Payment Methods (প্রতি লাইনে একটি)"><textarea style={{ ...inputStyle, height: 80, resize: "vertical", fontSize: 11 }} value={form.payment_methods} onChange={e => setForm({ ...form, payment_methods: e.target.value })} /></FormField>
      <FormField label="Payment Terms & Policy (প্রতি লাইনে একটি)"><textarea style={{ ...inputStyle, height: 60, resize: "vertical", fontSize: 11 }} value={form.payment_terms_policy} onChange={e => setForm({ ...form, payment_terms_policy: e.target.value })} /></FormField>

      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
        <button onClick={onClose} style={{ ...btnPrimary, width: "auto", background: C.gray400, padding: "9px 16px" }}>বাতিল</button>
        <button onClick={() => { if (!form.project_name || !form.client_name) return alert("Project ও Client name দিন"); onSave(form); }} style={{ ...btnPrimary, width: "auto", padding: "9px 20px" }}>{existing ? "✅ সংরক্ষণ করুন" : "✅ তৈরি করুন"}</button>
      </div>
    </Modal>
  );
}

// ============================================================
// BOQ ITEM LIBRARY (standard items with description, for autocomplete)
// ============================================================
function BOQItemLibrary({ stdRates, onRefresh }) {
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const blankForm = { item_name: "", category: "", unit: "sft", rate: "", description: "", specification: "" };
  const [form, setForm] = useState(blankForm);
  const uploadRef = useRef();

  const save = async () => {
    if (!form.item_name) return alert("Item নাম আবশ্যক");
    const payload = { ...form, rate: +form.rate || 0 };
    const { error } = editItem ? await supabase.from("standard_rates").update(payload).eq("id", editItem.id) : await supabase.from("standard_rates").insert([payload]);
    if (error) return alert("❌ ব্যর্থ: " + error.message);
    setShowModal(false); setForm(blankForm); setEditItem(null); onRefresh();
  };

  const del = async (id) => { if (!confirm("এই Item মুছবেন?")) return; await supabase.from("standard_rates").delete().eq("id", id); onRefresh(); };

  const handleExport = () => exportToExcel(stdRates.map(sr => ({ "Item Name": sr.item_name, Category: sr.category, Unit: sr.unit, Rate: sr.rate, Description: sr.description || sr.work_description || "", Specification: sr.specification || "" })), "Item Library", "BOQ_Item_Library");

  const handleUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    const rows = await parseExcelFile(file);
    let count = 0;
    for (const row of rows) {
      const item_name = row["Item Name"] || row["item_name"] || row["Item"];
      if (!item_name) continue;
      await supabase.from("standard_rates").insert([{
        item_name, category: row["Category"] || row["category"] || "",
        unit: row["Unit"] || row["unit"] || "sft", rate: +row["Rate"] || +row["rate"] || 0,
        description: row["Description"] || row["description"] || "",
        specification: row["Specification"] || row["specification"] || "",
      }]);
      count++;
    }
    alert("✅ " + count + "টি Item আপলোড হয়েছে!"); e.target.value = ""; onRefresh();
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 15 }}>📚 Item Library ({fmtNum(stdRates.length)})</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={() => { setEditItem(null); setForm(blankForm); setShowModal(true); }} style={{ ...btnPrimary, width: "auto", margin: 0 }}>+ নতুন Item</button>
          <button onClick={handleExport} style={btnEdit}>⬇️ Excel Download</button>
          <button onClick={() => uploadRef.current?.click()} style={btnEdit}>⬆️ Excel Upload</button>
          <input type="file" ref={uploadRef} onChange={handleUpload} accept=".xlsx,.xls,.csv" style={{ display: "none" }} />
        </div>
      </div>
      <Card>
        {stdRates.length === 0 ? (
          <div style={{ textAlign: "center", padding: 20, color: C.gray400, fontSize: 13 }}>কোনো Item নেই। Excel Upload করুন অথবা ম্যানুয়ালি যোগ করুন।</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead><tr style={{ background: C.primaryBg }}>{["Item Name", "Category", "Unit", "Rate", "Description", "Specification", "Action"].map(h => <th key={h} style={{ padding: "8px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600 }}>{h}</th>)}</tr></thead>
              <tbody>
                {stdRates.map(sr => (
                  <tr key={sr.id} style={{ borderBottom: "1px solid " + C.gray100 }}>
                    <td style={{ padding: "8px 10px", fontWeight: 600, color: C.primaryDark }}>{sr.item_name}</td>
                    <td style={{ padding: "8px 10px" }}>{sr.category}</td>
                    <td style={{ padding: "8px 10px" }}>{sr.unit}</td>
                    <td style={{ padding: "8px 10px" }}>{fmt(sr.rate)}</td>
                    <td style={{ padding: "8px 10px", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {(sr.description || sr.work_description) ? <span style={{ color: C.gray600 }}>{sr.description || sr.work_description}</span> : <span style={{ color: C.red, fontStyle: "italic" }}>⚠️ খালি</span>}
                    </td>
                    <td style={{ padding: "8px 10px", color: C.gray600, maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{sr.specification || "—"}</td>
                    <td style={{ padding: "8px 10px" }}>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => { setEditItem(sr); setForm({ item_name: sr.item_name, category: sr.category || "", unit: sr.unit || "sft", rate: sr.rate || "", description: sr.description || sr.work_description || "", specification: sr.specification || "" }); setShowModal(true); }} style={btnEdit}>✏️</button>
                        <button onClick={() => del(sr.id)} style={btnDanger}>🗑️</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {showModal && (
        <Modal title={editItem ? "Item সম্পাদনা" : "নতুন Item"} onClose={() => setShowModal(false)}>
          <FormField label="Item Name *"><input style={inputStyle} value={form.item_name} onChange={e => setForm({ ...form, item_name: e.target.value })} placeholder="False Ceiling (Gypsum)" /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="Category"><input style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })} /></FormField>
            <FormField label="Unit"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{BOQ_UNITS.map(u => <option key={u}>{u}</option>)}</select></FormField>
          </div>
          <FormField label="Rate (৳)"><input type="number" style={inputStyle} value={form.rate} onChange={e => setForm({ ...form, rate: e.target.value })} /></FormField>
          <FormField label="Description"><textarea style={{ ...inputStyle, minHeight: 60 }} value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></FormField>
          <FormField label="Specification"><input style={inputStyle} value={form.specification} onChange={e => setForm({ ...form, specification: e.target.value })} placeholder="Brand, grade, model" /></FormField>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
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
        {stdRates.length > 0 && (
          <div style={{ gridColumn: "1/-1" }}>
            <FormField label="📚 Item Library থেকে বাছুন (auto-fill হবে)">
              <select
                style={{ ...inputStyle, fontWeight: 600, color: C.primaryDark }}
                value=""
                onChange={e => {
                  const sr = stdRates.find(s => String(s.id) === e.target.value);
                  if (!sr) return;
                  setForm(f => ({
                    ...f,
                    item_name: sr.item_name,
                    unit: sr.unit || f.unit,
                    rate: sr.rate,
                    work_description: sr.description || sr.work_description || "",
                    specification: sr.specification || "",
                    is_rate_fixed: true,
                  }));
                }}
              >
                <option value="">— একটা Item বাছুন —</option>
                {stdRates.map(sr => <option key={sr.id} value={sr.id}>{sr.item_name} ({fmt(sr.rate)}/{sr.unit})</option>)}
              </select>
            </FormField>
          </div>
        )}

        <div style={{ gridColumn: "1/-1" }}>
          <FormField label="Item Name * (উপর থেকে বাছার পর এখানে edit-ও করা যাবে)">
            <input
              style={inputStyle}
              list="std-item-names"
              value={form.item_name}
              onChange={e => {
                const val = e.target.value;
                const match = stdRates.find(sr => sr.item_name === val);
                if (match) {
                  setForm(f => ({ ...f, item_name: match.item_name, unit: match.unit, rate: match.rate, work_description: match.description || match.work_description || f.work_description, specification: match.specification || f.specification, is_rate_fixed: true }));
                } else {
                  setForm(f => ({ ...f, item_name: val }));
                }
              }}
              placeholder="যেমন: False Ceiling"
            />
            <datalist id="std-item-names">
              {stdRates.map(sr => <option key={sr.id} value={sr.item_name} />)}
            </datalist>
          </FormField>
        </div>
        <div style={{ gridColumn: "1/-1" }}><FormField label="Work Description"><input style={inputStyle} value={form.work_description} onChange={e => setForm({ ...form, work_description: e.target.value })} /></FormField></div>
        <div style={{ gridColumn: "1/-1" }}><FormField label="Specification"><input style={inputStyle} value={form.specification} onChange={e => setForm({ ...form, specification: e.target.value })} placeholder="Brand, grade, model" /></FormField></div>
        <FormField label="Unit"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{BOQ_UNITS.map(u => <option key={u}>{u}</option>)}</select></FormField>
        <FormField label="Qty *"><input style={inputStyle} type="number" value={form.qty} onChange={e => setForm({ ...form, qty: e.target.value })} /></FormField>
        <FormField label="Rate (৳) *"><input style={inputStyle} type="number" value={form.rate} onChange={e => setForm({ ...form, rate: e.target.value })} /></FormField>
        <div style={{ display: "flex", alignItems: "center", gap: 8, paddingTop: 18 }}><input type="checkbox" id="fixed" checked={form.is_rate_fixed} onChange={e => setForm({ ...form, is_rate_fixed: e.target.checked })} /><label htmlFor="fixed" style={{ cursor: "pointer", fontSize: 13 }}>🔒 Rate Fixed রাখুন</label></div>
        <div style={{ gridColumn: "1/-1", background: C.primaryBg, padding: "10px 14px", borderRadius: 8, textAlign: "center", fontWeight: 700, color: C.primaryDark }}>Amount: ৳ {amt}</div>
      </div>
      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
        <button onClick={onClose} style={{ ...btnPrimary, width: "auto", background: C.gray400, padding: "9px 16px" }}>বাতিল</button>
        <button onClick={() => { if (!form.item_name || !form.qty || !form.rate) return alert("Item, Qty ও Rate দিন"); onSave(form); }} style={{ ...btnPrimary, width: "auto", padding: "9px 20px" }}>{item ? "✅ Update করুন" : "✅ যোগ করুন"}</button>
      </div>
    </Modal>
  );
}

// ============================================================
// BOQ MULTI ITEM (bulk add, like Expenses bulk add)
// ============================================================
function BOQMultiItemModal({ stdRates, existingRooms, onSave, onClose }) {
  const allRooms = [...new Set([...existingRooms, ...BOQ_ROOMS])];
  const blankRow = () => ({ room_name: allRooms[0] || "Master Bedroom", code_no: "", item_no: 1, item_name: "", work_description: "", specification: "", unit: "sft", qty: "", rate: "" });
  const [rows, setRows] = useState([blankRow()]);

  const updateRow = (i, field, val) => {
    setRows(r => r.map((row, idx) => {
      if (idx !== i) return row;
      if (field === "item_name") {
        const match = stdRates.find(sr => sr.item_name === val);
        if (match) return { ...row, item_name: val, unit: match.unit || row.unit, rate: match.rate ?? row.rate, work_description: match.description || match.work_description || row.work_description, specification: match.specification || row.specification };
      }
      return { ...row, [field]: val };
    }));
  };
  const addRow = () => setRows(r => [...r, blankRow()]);
  const removeRow = (i) => setRows(r => r.filter((_, idx) => idx !== i));

  return (
    <Modal title="একসাথে একাধিক BOQ Item যোগ করুন" onClose={onClose} size={1050}>
      <datalist id="mi-std-item-names">{stdRates.map(sr => <option key={sr.id} value={sr.item_name} />)}</datalist>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ background: C.primaryBg }}>
              {["Room", "Code No.", "Item No.", "Item Name *", "Description", "Specification", "Unit", "Qty", "Rate *", ""].map(h => (
                <th key={h} style={{ padding: "8px 6px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => (
              <tr key={i} style={{ borderBottom: "1px solid " + C.gray100 }}>
                <td style={{ padding: 4 }}>
                  <select style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 110 }} value={row.room_name} onChange={e => updateRow(i, "room_name", e.target.value)}>
                    {allRooms.map(r => <option key={r}>{r}</option>)}
                  </select>
                </td>
                <td style={{ padding: 4 }}><input style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, width: 60 }} value={row.code_no} onChange={e => updateRow(i, "code_no", e.target.value)} /></td>
                <td style={{ padding: 4 }}><input type="number" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, width: 50 }} value={row.item_no} onChange={e => updateRow(i, "item_no", e.target.value)} /></td>
                <td style={{ padding: 4 }}><input list="mi-std-item-names" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 140 }} value={row.item_name} onChange={e => updateRow(i, "item_name", e.target.value)} placeholder="Item name" /></td>
                <td style={{ padding: 4 }}><input style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 160 }} value={row.work_description} onChange={e => updateRow(i, "work_description", e.target.value)} /></td>
                <td style={{ padding: 4 }}><input style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, minWidth: 120 }} value={row.specification} onChange={e => updateRow(i, "specification", e.target.value)} /></td>
                <td style={{ padding: 4 }}>
                  <select style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, width: 70 }} value={row.unit} onChange={e => updateRow(i, "unit", e.target.value)}>
                    {BOQ_UNITS.map(u => <option key={u}>{u}</option>)}
                  </select>
                </td>
                <td style={{ padding: 4 }}><input type="number" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, width: 60 }} value={row.qty} onChange={e => updateRow(i, "qty", e.target.value)} /></td>
                <td style={{ padding: 4 }}><input type="number" style={{ ...inputStyle, padding: "5px 6px", fontSize: 11, width: 70 }} value={row.rate} onChange={e => updateRow(i, "rate", e.target.value)} /></td>
                <td style={{ padding: 4 }}>{rows.length > 1 && <button onClick={() => removeRow(i)} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 14 }}>🗑️</button>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <button onClick={addRow} style={{ ...btnEdit, marginTop: 10 }}>➕ সারি যোগ করুন</button>
      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
        <button onClick={onClose} style={{ ...btnPrimary, width: "auto", background: C.gray400, padding: "9px 16px" }}>বাতিল</button>
        <button onClick={() => onSave(rows)} style={{ ...btnPrimary, width: "auto", padding: "9px 20px" }}>✅ সব Item সংরক্ষণ করুন ({rows.filter(r => r.item_name && r.rate).length})</button>
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
          <div style={{ color: C.gray600, fontSize: 13 }}>{projects.length > 0 ? projects.length + "টি প্রজেক্ট আছে। উপর থেকে বেছে নিন।" : "শুরু করতে নতুন প্রজেক্ট তৈরি করুন।"}</div>
        </Card>
      ) : (
        <>
          <div style={{ background: "linear-gradient(135deg, " + C.primaryDark + ", " + C.primary + ")", borderRadius: 12, padding: "14px 20px", marginBottom: 16, color: C.white, display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
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
          <div style={{ display: "flex", gap: 4, marginBottom: 20, borderBottom: "2px solid " + C.gray200, overflowX: "auto" }}>
            {CP_TABS.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)} style={{ padding: "10px 16px", border: "none", borderBottom: tab === t.id ? "3px solid " + C.primary : "3px solid transparent", background: "none", color: tab === t.id ? C.primaryDark : C.gray600, fontWeight: tab === t.id ? 700 : 400, cursor: "pointer", fontSize: 13, fontFamily: "inherit", whiteSpace: "nowrap" }}>
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
  const emptyForm = { update_date: new Date().toISOString().split("T")[0], work_done: "", category: "নির্মাণ কাজ", workers_count: "", weather: "স্বাভাবিক", progress_pct: "", issues: "", next_plan: "", reported_by: "", image_url: "" };
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
      <SectionHeader title="📅 Daily Work Updates" action="নতুন আপডেট" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => { printSection("Daily Work Updates", "daily-print"); }} onExport={() => exportToExcel(items.map(i => ({ তারিখ: i.update_date, কাজ: i.work_done, ক্যাটাগরি: i.category, শ্রমিক: i.workers_count, আবহাওয়া: i.weather, অগ্রগতি: i.progress_pct + "%", সমস্যা: i.issues, পরবর্তী_পরিকল্পনা: i.next_plan, রিপোর্টকারী: i.reported_by })), "Daily", "Daily_Updates")} />
      <div id="daily-print">
        {items.length === 0 ? <Card style={{ textAlign: "center", padding: 40, color: C.gray400 }}>কোনো আপডেট নেই!</Card> : items.map(item => (
          <Card key={item.id} style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8, flexWrap: "wrap" }}>
                  <span style={{ fontWeight: 700, color: C.primary, fontSize: 14 }}>📅 {item.update_date}</span>
                  <Badge label={item.category} color="primary" />
                  <Badge label={"👷 " + item.workers_count + " জন"} color="blue" />
                  <Badge label={"☁️ " + item.weather} color="gray" />
                  {item.progress_pct > 0 && <Badge label={"📈 " + item.progress_pct + "%"} color="green" />}
                </div>
                <div style={{ fontWeight: 600, color: C.primaryDark, fontSize: 14, marginBottom: 6 }}>{item.work_done}</div>
                {item.issues && <div style={{ fontSize: 12, color: C.red, background: C.redLight, padding: "6px 10px", borderRadius: 6, marginBottom: 6 }}>⚠️ {item.issues}</div>}
                {item.next_plan && <div style={{ fontSize: 12, color: C.gray600, background: C.gray50, padding: "6px 10px", borderRadius: 6 }}>📋 {item.next_plan}</div>}
                {item.reported_by && <div style={{ fontSize: 11, color: C.gray400, marginTop: 4 }}>রিপোর্টকারী: {item.reported_by}</div>}
                {item.image_url && <div style={{ marginTop: 8 }}><img src={item.image_url} alt="Site" style={{ maxWidth: "100%", maxHeight: 150, borderRadius: 8, border: "1px solid " + C.gray200, objectFit: "cover" }} /></div>}
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
          <ImageUploadField label="📷 সাইট ছবি (ঐচ্ছিক)" value={form.image_url} onChange={url => setForm(f => ({ ...f, image_url: url }))} folder="daily-updates" />
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

function CPExpenses({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [showMultiModal, setShowMultiModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { expense_date: new Date().toISOString().split("T")[0], category: "General Labour", item_name: "", description: "", unit: "পিস", quantity: 1, unit_price: "", amount: 0, supplier: "", payment_method: "নগদ", payment_status: "পরিশোধিত", received_by: "", note: "", image_url: "" };
  const [form, setForm] = useState(emptyForm);
  const categories = ["General Labour", "Construction Labour", "Electrical Labour", "Plumbing Labour", "Supervision from Head Office", "Painter", "Interior Materials", "Interior Labour", "নির্মাণ সামগ্রী", "শ্রমিক মজুরি", "ইলেকট্রিক সামগ্রী", "প্লাম্বিং সামগ্রী", "যন্ত্রপাতি ভাড়া", "পরিবহন", "টাইলস/পাথর", "রঙ সামগ্রী", "দরজা-জানালা", "হার্ডওয়্যার", "অন্যান্য"];
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("site_expenses").select("*").eq("project_id", projectId).order("expense_date", { ascending: false }); setItems(data || []); };
  const [rows, setRows] = useState([{ expense_date: new Date().toISOString().split("T")[0], category: "General Labour", item_name: "", description: "", unit: "পিস", quantity: 1, unit_price: "", payment_method: "নগদ", payment_status: "পরিশোধিত", supplier: "", received_by: "", note: "" }]);

  const addRow = () => setRows(r => [...r, { expense_date: r[r.length - 1]?.expense_date || new Date().toISOString().split("T")[0], category: "General Labour", item_name: "", description: "", unit: "পিস", quantity: 1, unit_price: "", payment_method: "নগদ", payment_status: "পরিশোধিত", supplier: "", received_by: "", note: "" }]);
  const removeRow = (idx) => setRows(r => r.filter((_, i) => i !== idx));
  const updateRow = (idx, field, val) => setRows(r => r.map((row, i) => i === idx ? { ...row, [field]: val, amount: field === "quantity" || field === "unit_price" ? ((field === "quantity" ? +val : +row.quantity || 1) * (field === "unit_price" ? +val : +row.unit_price || 0)) : row.amount } : row));

  const saveMultiRows = async () => {
    const valid = rows.filter(r => r.item_name && r.unit_price);
    if (valid.length === 0) return alert("কমপক্ষে একটি আইটেমের নাম ও মূল্য দিন!");
    const payloads = valid.map(r => ({ ...r, project_id: projectId, quantity: +r.quantity || 1, unit_price: +r.unit_price || 0, amount: (+r.quantity || 1) * (+r.unit_price || 0) }));
    await supabase.from("site_expenses").insert(payloads);
    await load(); setShowModal(false); setEditItem(null);
    setRows([{ expense_date: new Date().toISOString().split("T")[0], category: "General Labour", item_name: "", description: "", unit: "পিস", quantity: 1, unit_price: "", payment_method: "নগদ", payment_status: "পরিশোধিত", supplier: "", received_by: "", note: "" }]);
  };

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
      <div style={{ display: "flex", gap: 8, marginBottom: 4 }}>
        <button onClick={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} style={{ background: C.primary, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ একটি খরচ</button>
        <button onClick={() => { setEditItem(null); setShowMultiModal(true); }} style={{ background: "#2A5C8F", color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ একসাথে অনেক খরচ</button>
      </div>
      <SectionHeader title="💸 Project Expenses" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => { printSection("Project Expenses", "expenses-print"); }} onExport={() => exportToExcel(items.map(i => ({ তারিখ: i.expense_date, ক্যাটাগরি: i.category, আইটেম: i.item_name, বিবরণ: i.description, একক: i.unit, পরিমাণ: i.quantity, একক_মূল্য: i.unit_price, মোট: i.amount, সাপ্লায়ার: i.supplier, পেমেন্ট: i.payment_method, স্ট্যাটাস: i.payment_status })), "Expenses", "Site_Expenses")} />
      {Object.keys(catGroups).length > 0 && (
        <Card style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 13, marginBottom: 8 }}>ক্যাটাগরি অনুযায়ী:</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {Object.entries(catGroups).sort((a, b) => b[1] - a[1]).map(([cat, total]) => (
              <div key={cat} style={{ background: C.primaryBg, border: "1px solid " + C.primaryLight, borderRadius: 8, padding: "6px 12px" }}>
                <div style={{ fontSize: 11, color: C.gray600 }}>{cat}</div>
                <div style={{ fontSize: 13, fontWeight: 700, color: C.primaryDark }}>{fmt(total)}</div>
              </div>
            ))}
          </div>
        </Card>
      )}
      <Card>
        <div id="expenses-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
            <thead><tr style={{ background: C.primaryBg }}>
              {["তারিখ", "ক্যাটাগরি", "আইটেম", "বিবরণ", "পরিমাণ", "একক মূল্য", "মোট", "সাপ্লায়ার"].map(h => <th key={h} style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>)}
              <th className="no-print" style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>পেমেন্ট</th>
              <th className="no-print" style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>স্ট্যাটাস</th>
              <th className="no-print" style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>Action</th>
            </tr></thead>
            <tbody>
              {(() => {
                const dateGroups = items.reduce((acc, i) => { const d = i.expense_date || "—"; if (!acc[d]) acc[d] = []; acc[d].push(i); return acc; }, {});
                return Object.entries(dateGroups).map(([date, dateItems]) => {
                  const dateTotal = dateItems.reduce((s, i) => s + (i.amount || 0), 0);
                  return (
                    <Fragment key={date}>
                      {dateItems.map(item => (
                        <tr key={item.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                          <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{item.expense_date}</td>
                          <td style={{ padding: "8px 10px" }}><Badge label={item.category} color="primary" /></td>
                          <td style={{ padding: "8px 10px", fontWeight: 600, color: C.primaryDark }}>{item.item_name}</td>
                          <td style={{ padding: "8px 10px", fontSize: 11, color: C.gray600, maxWidth: 150 }}>{item.description || "—"}</td>
                          <td style={{ padding: "8px 10px" }}>{item.quantity} {item.unit}</td>
                          <td style={{ padding: "8px 10px" }}>{fmt(item.unit_price)}</td>
                          <td style={{ padding: "8px 10px", fontWeight: 700, color: C.red }}>{fmt(item.amount)}</td>
                          <td style={{ padding: "8px 10px", fontSize: 11 }}>{item.supplier || "—"}</td>
                          <td className="no-print" style={{ padding: "8px 10px", fontSize: 11 }}>{item.payment_method}</td>
                          <td className="no-print" style={{ padding: "8px 10px" }}><Badge label={item.payment_status} color={item.payment_status === "পরিশোধিত" ? "green" : "yellow"} /></td>
                          <td className="no-print" style={{ padding: "8px 10px" }}><div style={{ display: "flex", gap: 4 }}><button onClick={() => { setEditItem(item); setForm({ ...item }); setShowModal(true); }} style={btnEdit}>✏️</button><button onClick={() => del(item.id)} style={btnDanger}>🗑️</button></div></td>
                        </tr>
                      ))}
                      {dateItems.length > 1 && (
                        <tr style={{ background: C.gray50, fontWeight: 700 }}>
                          <td colSpan={6} style={{ padding: "8px 10px", textAlign: "right", color: C.primaryDark, fontSize: 12 }}>{date} — এই দিনের মোট:</td>
                          <td style={{ padding: "8px 10px", color: C.red, fontSize: 13 }}>{fmt(dateTotal)}</td>
                          <td className="no-print" colSpan={3}></td>
                        </tr>
                      )}
                    </Fragment>
                  );
                });
              })()}
              {items.length > 0 && <tr style={{ background: C.primaryBg, fontWeight: 700 }}><td colSpan={6} style={{ padding: "10px", textAlign: "right", color: C.primaryDark }}>সর্বমোট:</td><td style={{ padding: "10px", color: C.red, fontSize: 15 }}>{fmt(totalExpense)}</td><td></td><td className="no-print" colSpan={3}></td></tr>}
            </tbody>
          </table>
          {items.length === 0 && <div style={{ textAlign: "center", padding: 30, color: C.gray400 }}>কোনো খরচ নেই!</div>}
        </div>
      </Card>
      {showMultiModal && (
        <Modal title="একসাথে একাধিক খরচ যোগ করুন" onClose={() => setShowMultiModal(false)} size={900}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: C.primaryBg }}>
                  {["তারিখ", "ক্যাটাগরি", "আইটেম *", "বিবরণ", "একক", "পরিমাণ", "একক মূল্য *", "মোট", "পেমেন্ট", ""].map(h => (
                    <th key={h} style={{ padding: "8px 6px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, idx) => (
                  <tr key={idx} style={{ borderBottom: "1px solid " + C.gray100 }}>
                    <td style={{ padding: "4px" }}><input type="date" value={row.expense_date} onChange={e => updateRow(idx, "expense_date", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 110 }} /></td>
                    <td style={{ padding: "4px" }}>
                      <select value={row.category} onChange={e => updateRow(idx, "category", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 120 }}>
                        {["General Labour","Construction Labour","Electrical Labour","Plumbing Labour","Supervision from Head Office","Painter","Interior Materials","Interior Labour","নির্মাণ সামগ্রী","শ্রমিক মজুরি","ইলেকট্রিক","প্লাম্বিং","যন্ত্রপাতি","পরিবহন","টাইলস","রঙ","দরজা-জানালা","হার্ডওয়্যার","অন্যান্য"].map(c => <option key={c}>{c}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "4px" }}><input value={row.item_name} onChange={e => updateRow(idx, "item_name", e.target.value)} placeholder="আইটেম নাম *" style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 120 }} /></td>
                    <td style={{ padding: "4px" }}><input value={row.description} onChange={e => updateRow(idx, "description", e.target.value)} placeholder="বিবরণ" style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 100 }} /></td>
                    <td style={{ padding: "4px" }}>
                      <select value={row.unit} onChange={e => updateRow(idx, "unit", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 60 }}>
                        {["পিস","ব্যাগ","কেজি","টন","সিএফটি","বর্গফুট","রানিংফুট","লিটার","দিন","লট"].map(u => <option key={u}>{u}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "4px" }}><input type="number" value={row.quantity} onChange={e => updateRow(idx, "quantity", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, width: 60 }} /></td>
                    <td style={{ padding: "4px" }}><input type="number" value={row.unit_price} onChange={e => updateRow(idx, "unit_price", e.target.value)} placeholder="মূল্য *" style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, width: 80 }} /></td>
                    <td style={{ padding: "4px", fontWeight: 700, color: C.primaryDark, whiteSpace: "nowrap" }}>{fmt((+row.quantity || 1) * (+row.unit_price || 0))}</td>
                    <td style={{ padding: "4px" }}>
                      <select value={row.payment_status} onChange={e => updateRow(idx, "payment_status", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 80 }}>
                        {["পরিশোধিত","বকেয়া","আংশিক"].map(s => <option key={s}>{s}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "4px" }}>
                      {rows.length > 1 && <button onClick={() => removeRow(idx)} style={{ ...btnDanger, padding: "4px 8px" }}>×</button>}
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{ background: C.primaryBg }}>
                  <td colSpan={7} style={{ padding: "8px", textAlign: "right", fontWeight: 700, color: C.primaryDark }}>সর্বমোট:</td>
                  <td style={{ padding: "8px", fontWeight: 800, color: C.red, fontSize: 14 }}>{fmt(rows.reduce((s, r) => s + ((+r.quantity || 1) * (+r.unit_price || 0)), 0))}</td>
                  <td colSpan={2}></td>
                </tr>
              </tfoot>
            </table>
          </div>
          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <button onClick={addRow} style={{ background: C.primaryBg, color: C.primaryDark, border: "1px solid " + C.primary, borderRadius: 8, padding: "10px 20px", fontWeight: 600, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>+ নতুন সারি যোগ করুন</button>
            <button onClick={saveMultiRows} style={{ ...btnPrimary, width: "auto", padding: "10px 24px" }}>✅ সব সংরক্ষণ করুন ({rows.filter(r => r.item_name && r.unit_price).length}টি)</button>
          </div>
        </Modal>
      )}
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
            <div style={{ gridColumn: "1/-1" }}><ImageUploadField label="📷 রসিদ / ছবি (ঐচ্ছিক)" value={form.image_url} onChange={url => setForm(f => ({ ...f, image_url: url }))} folder="expenses" /></div>
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
      {lowStock.length > 0 && <Card style={{ background: C.redLight, border: "1px solid #F5C6CB", marginBottom: 14 }}><div style={{ fontWeight: 700, color: C.red, marginBottom: 8 }}>⚠️ কম স্টক:</div><div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>{lowStock.map(i => <span key={i.id} style={{ background: C.white, padding: "4px 10px", borderRadius: 6, fontSize: 12, color: C.red, fontWeight: 600 }}>{i.item_name}: {i.closing_stock} {i.unit}</span>)}</div></Card>}
      <SectionHeader title="🧱 Stock Register" action="নতুন আইটেম" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => { printSection("Stock Register", "stock-print"); }} onExport={() => exportToExcel(items.map(i => ({ আইটেম: i.item_name, ক্যাটাগরি: i.category, একক: i.unit, প্রারম্ভিক: i.opening_stock, প্রাপ্ত: i.received, ব্যবহৃত: i.used, অবশিষ্ট: i.closing_stock, একক_মূল্য: i.unit_price, মোট_মূল্য: i.total_value, সাপ্লায়ার: i.supplier })), "Stock", "Stock_Register")} />
      <Card>
        <div id="stock-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["আইটেম", "ক্যাটাগরি", "একক", "প্রারম্ভিক", "প্রাপ্ত", "ব্যবহৃত", "অবশিষ্ট", "একক মূল্য", "মোট মূল্য", "সাপ্লায়ার", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
            <tbody>
              {items.map(item => (
                <tr key={item.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
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
      <SectionHeader title="💰 Office থেকে প্রাপ্ত টাকা" action="নতুন প্রাপ্তি" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => { printSection("Office Payment Register", "payments-print"); }} onExport={() => exportToExcel(items.map(i => ({ তারিখ: i.receive_date, পরিমাণ: i.amount, কিস্তি: i.payment_type, পদ্ধতি: i.payment_method, প্রেরক: i.sender_name, গ্রহণকারী: i.received_by, রেফারেন্স: i.bank_ref || i.bkash_ref || "", নোট: i.note })), "Payments", "Office_Payments")} />
      <Card>
        <div id="payments-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "কিস্তির ধরন", "পরিমাণ", "পেমেন্ট পদ্ধতি", "প্রেরক", "গ্রহণকারী", "রেফারেন্স", "নোট", "Action"].map(h => <th key={h} style={{ padding: "9px 12px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
            <tbody>
              {items.map(item => (
                <tr key={item.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
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
      <SectionHeader title="📊 Project Summary" onPrint={() => { printSection("Project Summary", "summary-print"); }} />
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
  { id: "interior", icon: "🛋️", label: "Interior Projects", roles: ["admin", "site_engineer"] },
  { id: "boq", icon: "📋", label: "Estimate Project", roles: ["admin"] },
  { id: "clients", icon: "👥", label: "ক্লায়েন্ট", roles: ["admin"] },
  { id: "documents", icon: "🧾", label: "রশিদ ও চালান", roles: ["admin"] },
  { id: "hr_system", icon: "👥", label: "HR ও পে-রোল সিস্টেম", roles: ["admin"] },
  { id: "my_attendance", icon: "🙋", label: "আমার হাজিরা", roles: ["admin", "employee"] },
  { id: "finance", icon: "💰", label: "আর্থিক", roles: ["admin"] },
  { id: "site", icon: "📍", label: "সাইট প্রগ্রেস", roles: ["admin"] },
  { id: "materials", icon: "📦", label: "সামগ্রী", roles: ["admin"] },
  { id: "analytics", icon: "📊", label: "রিপোর্ট & Analytics", roles: ["admin"] },
  { id: "users", icon: "👤", label: "User Management", roles: ["admin"] },
  { id: "password", icon: "🔑", label: "পাসওয়ার্ড", roles: ["admin", "site_engineer"] },
];

// Menu items that actually render something for an "employee" role account.
// (Most screens above are admin-only dashboards and will show a blank page
// if granted to an employee, since they require isAdmin to render.)
// HR & Payroll System sub-tabs — each independently assignable as a permission
const HR_SUBTABS = [
  ["hr_employees", "👷", "কর্মী তালিকা"],
  ["hr_leave", "🌴", "ছুটি ব্যবস্থাপনা"],
  ["hr_attendance", "📋", "উপস্থিতি"],
  ["hr_smart_attendance", "⏱️", "স্মার্ট অ্যাটেন্ডেন্স"],
  ["hr_payroll", "🧾", "পে-রোল"],
  ["hr_recruitment", "🧑‍💼", "নিয়োগ"],
  ["hr_reports", "📊", "HR রিপোর্ট"],
  ["hr_performance", "🎯", "পারফরম্যান্স"],
  ["hr_timetracker", "⏲️", "টাইম ট্র্যাকার"],
  ["hr_helpdesk", "🎫", "HR Help Desk"],
];
const hasHRAccess = (currentUser, isAdmin) => isAdmin || (currentUser?.role === "employee" && HR_SUBTABS.some(([id]) => (currentUser.permissions || []).includes(id)));
const canAccessMenu = (currentUser, isAdmin, id) => isAdmin || (currentUser?.role === "employee" && (currentUser.permissions || []).includes(id));

const LEAVE_TYPES = [
  { id: "casual", label: "নৈমিত্তিক ছুটি", quota: 10 },
  { id: "sick", label: "অসুস্থতা ছুটি", quota: 14 },
  { id: "annual", label: "বার্ষিক ছুটি", quota: 15 },
  { id: "maternity", label: "মাতৃত্বকালীন ছুটি", quota: 112 },
  { id: "paternity", label: "পিতৃত্বকালীন ছুটি", quota: 7 },
];

const DISBURSEMENT_CHANNELS = [
  { id: "Bank", label: "ব্যাংক", color: "#5B4FCF" },
  { id: "Bkash", label: "বিকাশ", color: "#E0A800" },
  { id: "Nagad", label: "নগদ", color: "#FF5A5F" },
  { id: "Cash", label: "নগদ (Cash)", color: "#3498DB" },
];

const REQUEST_TYPES = {
  overtime: { label: "ওভারটাইম (Overtime)", icon: "⏱️", needsTime: true },
  regularization: { label: "সংশোধন (Regularization)", icon: "✏️", needsTime: false },
  on_duty: { label: "On Duty", icon: "🚗", needsTime: true },
  hourly_permission: { label: "Hourly Permission", icon: "⏳", needsTime: true },
  shift_change: { label: "Shift পরিবর্তনের অনুরোধ", icon: "🔁", needsTime: false, needsShift: true },
};




// ============================================================
// INTERIOR PROJECTS MODULE
// ============================================================
function IPNewProjectModal({ onSave, onClose }) {
  const [form, setForm] = useState({ name: "", client_name: "", site_address: "", project_type: "আবাসিক ইন্টেরিয়র", start_date: "", end_date: "", chief_designer: "", total_budget: "", deal_amount: "", status: "চলমান" });
  return (
    <Modal title="নতুন Interior প্রজেক্ট" onClose={onClose} size={580}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <div style={{ gridColumn: "1/-1" }}><FormField label="প্রজেক্টের নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="যেমন: রহিম সাহেবের ফ্ল্যাট ইন্টেরিয়র" /></FormField></div>
        <FormField label="ক্লায়েন্টের নাম *"><input style={inputStyle} value={form.client_name} onChange={e => setForm({ ...form, client_name: e.target.value })} /></FormField>
        <FormField label="প্রজেক্টের ধরন"><select style={inputStyle} value={form.project_type} onChange={e => setForm({ ...form, project_type: e.target.value })}>{["আবাসিক ইন্টেরিয়র", "কমার্শিয়াল ইন্টেরিয়র", "অফিস ইন্টেরিয়র", "রেস্তোরাঁ ইন্টেরিয়র", "মডুলার কিচেন", "থ্রিডি ডিজাইন", "রিনোভেশন", "অন্যান্য"].map(t => <option key={t}>{t}</option>)}</select></FormField>
        <div style={{ gridColumn: "1/-1" }}><FormField label="সাইটের ঠিকানা"><input style={inputStyle} value={form.site_address} onChange={e => setForm({ ...form, site_address: e.target.value })} /></FormField></div>
        <FormField label="প্রধান ডিজাইনার"><input style={inputStyle} value={form.chief_designer} onChange={e => setForm({ ...form, chief_designer: e.target.value })} /></FormField>
        <FormField label="স্ট্যাটাস"><select style={inputStyle} value={form.status} onChange={e => setForm({ ...form, status: e.target.value })}>{["চলমান", "বিরতি", "সম্পন্ন", "বাতিল"].map(s => <option key={s}>{s}</option>)}</select></FormField>
        <FormField label="শুরুর তারিখ"><input style={inputStyle} type="date" value={form.start_date} onChange={e => setForm({ ...form, start_date: e.target.value })} /></FormField>
        <FormField label="শেষের তারিখ"><input style={inputStyle} type="date" value={form.end_date} onChange={e => setForm({ ...form, end_date: e.target.value })} /></FormField>
        <FormField label="চুক্তি মূল্য (৳)"><input style={inputStyle} type="number" value={form.deal_amount} onChange={e => setForm({ ...form, deal_amount: e.target.value })} /></FormField>
        <FormField label="মোট বাজেট (৳)"><input style={inputStyle} type="number" value={form.total_budget} onChange={e => setForm({ ...form, total_budget: e.target.value })} /></FormField>
      </div>
      <button onClick={() => { if (!form.name || !form.client_name) return alert("নাম ও ক্লায়েন্ট আবশ্যক"); onSave(form); }} style={btnPrimary}>✅ প্রজেক্ট তৈরি করুন</button>
    </Modal>
  );
}

function InteriorProjects({ currentUser }) {
  const [projects, setProjects] = useState([]);
  const [selProject, setSelProject] = useState(null);
  const [tab, setTab] = useState("daily");
  const [showNewProject, setShowNewProject] = useState(false);
  const isAdmin = currentUser?.role === "admin";

  useEffect(() => { loadProjects(); }, []);

  const loadProjects = async () => {
    const { data } = await supabase.from("interior_projects").select("*").order("created_at", { ascending: false });
    let all = data || [];
    if (currentUser?.role === "site_engineer" && currentUser?.assigned_projects?.length > 0) {
      all = all.filter(p => currentUser.assigned_projects.includes(p.id));
    }
    setProjects(all);
  };

  const IP_TABS = [
    { id: "daily", icon: "📅", label: "Daily Updates" },
    { id: "expenses", icon: "💸", label: "Expenses" },
    { id: "stock", icon: "🪑", label: "Stock Register" },
    { id: "payments", icon: "💰", label: "Office থেকে টাকা" },
    { id: "summary", icon: "📊", label: "Project Summary" },
  ];

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, color: C.primaryDark, fontSize: 18, fontWeight: 700 }}>🛋️ Interior Projects</h2>
        {isAdmin && <button onClick={() => setShowNewProject(true)} style={{ background: C.primary, color: C.white, border: "none", borderRadius: 8, padding: "8px 16px", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>+ নতুন প্রজেক্ট</button>}
      </div>

      <Card style={{ marginBottom: 16, padding: "14px 18px" }}>
        <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
          <label style={{ fontWeight: 700, color: C.primaryDark, fontSize: 13, flexShrink: 0 }}>প্রজেক্ট বেছে নিন:</label>
          <select value={selProject?.id || ""} onChange={e => { const p = projects.find(x => x.id === e.target.value); setSelProject(p || null); setTab("daily"); }} style={{ ...inputStyle, maxWidth: 380, padding: "8px 12px" }}>
            <option value="">— প্রজেক্ট সিলেক্ট করুন —</option>
            {projects.map(p => <option key={p.id} value={p.id}>{p.name} — {p.client_name}</option>)}
          </select>
          {selProject && <span style={{ fontSize: 12, color: C.primaryLight, fontWeight: 600 }}>📍 {selProject.site_address}</span>}
        </div>
      </Card>

      {!selProject ? (
        <Card style={{ textAlign: "center", padding: "50px 20px" }}>
          <div style={{ fontSize: 48, marginBottom: 12 }}>🛋️</div>
          <div style={{ color: C.primaryDark, fontWeight: 700, fontSize: 16, marginBottom: 8 }}>Interior Project Management</div>
          <div style={{ color: C.gray600, fontSize: 13 }}>{projects.length > 0 ? projects.length + "টি প্রজেক্ট আছে। উপর থেকে বেছে নিন।" : "শুরু করতে নতুন প্রজেক্ট তৈরি করুন।"}</div>
        </Card>
      ) : (
        <>
          <div style={{ background: "linear-gradient(135deg, " + C.primaryDark + ", " + C.primary + ")", borderRadius: 12, padding: "14px 20px", marginBottom: 16, color: C.white, display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
            <div>
              <div style={{ fontSize: 16, fontWeight: 800 }}>{selProject.name}</div>
              <div style={{ fontSize: 12, opacity: 0.8, marginTop: 3 }}>ক্লায়েন্ট: {selProject.client_name} | {selProject.site_address}</div>
            </div>
            <div style={{ display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap" }}>
              <div style={{ textAlign: "center" }}><div style={{ fontSize: 11, opacity: 0.7 }}>ধরন</div><div style={{ fontSize: 12, fontWeight: 700 }}>{selProject.project_type}</div></div>
              <div style={{ textAlign: "center" }}><div style={{ fontSize: 11, opacity: 0.7 }}>চুক্তি</div><div style={{ fontSize: 13, fontWeight: 700 }}>{fmt(selProject.deal_amount)}</div></div>
              <span style={{ background: "rgba(255,255,255,0.2)", padding: "4px 12px", borderRadius: 20, fontSize: 12, fontWeight: 600 }}>{selProject.status}</span>
            </div>
          </div>

          <div style={{ display: "flex", gap: 4, marginBottom: 20, borderBottom: "2px solid " + C.gray200, overflowX: "auto" }}>
            {IP_TABS.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)} style={{ padding: "10px 16px", border: "none", borderBottom: tab === t.id ? "3px solid " + C.primary : "3px solid transparent", background: "none", color: tab === t.id ? C.primaryDark : C.gray600, fontWeight: tab === t.id ? 700 : 400, cursor: "pointer", fontSize: 13, fontFamily: "inherit", whiteSpace: "nowrap" }}>
                {t.icon} {t.label}
              </button>
            ))}
          </div>

          {tab === "daily" && <IPDailyUpdates projectId={selProject.id} />}
          {tab === "expenses" && <IPExpenses projectId={selProject.id} />}
          {tab === "stock" && <IPStock projectId={selProject.id} />}
          {tab === "payments" && <IPPayments projectId={selProject.id} />}
          {tab === "summary" && <IPSummary project={selProject} />}
        </>
      )}

      {showNewProject && <IPNewProjectModal onSave={async (form) => {
        const { error } = await supabase.from("interior_projects").insert([{ ...form, total_budget: +form.total_budget || 0, deal_amount: +form.deal_amount || 0, start_date: form.start_date || null, end_date: form.end_date || null }]);
        if (error) return alert("Error: " + error.message);
        await loadProjects(); setShowNewProject(false);
      }} onClose={() => setShowNewProject(false)} />}
    </div>
  );
}

// ---- INTERIOR DAILY UPDATES ----
function IPDailyUpdates({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { update_date: new Date().toISOString().split("T")[0], work_done: "", category: "ফার্নিচার কাজ", workers_count: "", progress_pct: "", issues: "", next_plan: "", reported_by: "", image_url: "" };
  const [form, setForm] = useState(emptyForm);
  const categories = ["ফার্নিচার কাজ", "সিলিং কাজ", "ওয়াল প্যানেলিং", "ফ্লোরিং", "পেইন্টিং", "ইলেকট্রিক", "প্লাম্বিং", "কিচেন ক্যাবিনেট", "ওয়ার্ডরোব", "গ্লাস ওয়ার্ক", "কার্টেন/ব্লাইন্ড", "লাইটিং", "পরিষ্কার", "থ্রিডি ভিজ্যুয়ালাইজেশন", "ক্লায়েন্ট মিটিং", "অন্যান্য"];
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("interior_daily_updates").select("*").eq("project_id", projectId).order("update_date", { ascending: false }); setItems(data || []); };
  const save = async () => {
    if (!form.work_done) return alert("কাজের বিবরণ আবশ্যক");
    const payload = { ...form, project_id: projectId, workers_count: +form.workers_count || 0, progress_pct: +form.progress_pct || 0 };
    if (editItem) { await supabase.from("interior_daily_updates").update(payload).eq("id", editItem.id); }
    else { await supabase.from("interior_daily_updates").insert([payload]); }
    await load(); setShowModal(false); setEditItem(null);
  };
  const del = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("interior_daily_updates").delete().eq("id", id); await load(); };
  return (
    <div>
      <SectionHeader title="📅 Daily Work Updates" action="নতুন আপডেট" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onPrint={() => { printSection("Interior Daily Updates", "ip-daily-print"); }} onExport={() => exportToExcel(items.map(i => ({ তারিখ: i.update_date, কাজ: i.work_done, ক্যাটাগরি: i.category, শ্রমিক: i.workers_count, অগ্রগতি: i.progress_pct + "%", সমস্যা: i.issues, পরিকল্পনা: i.next_plan, রিপোর্টকারী: i.reported_by })), "IPDaily", "Interior_Daily")} />
      <div id="ip-daily-print">
        {items.length === 0 ? <Card style={{ textAlign: "center", padding: 40, color: C.gray400 }}>কোনো আপডেট নেই!</Card> : items.map(item => (
          <Card key={item.id} style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8, flexWrap: "wrap" }}>
                  <span style={{ fontWeight: 700, color: C.primary, fontSize: 14 }}>📅 {item.update_date}</span>
                  <Badge label={item.category} color="primary" />
                  {item.workers_count > 0 && <Badge label={"👷 " + item.workers_count + " জন"} color="blue" />}
                  {item.progress_pct > 0 && <Badge label={"📈 " + item.progress_pct + "%"} color="green" />}
                </div>
                <div style={{ fontWeight: 600, color: C.primaryDark, fontSize: 14, marginBottom: 6 }}>{item.work_done}</div>
                {item.issues && <div style={{ fontSize: 12, color: C.red, background: C.redLight, padding: "6px 10px", borderRadius: 6, marginBottom: 6 }}>⚠️ {item.issues}</div>}
                {item.next_plan && <div style={{ fontSize: 12, color: C.gray600, background: C.gray50, padding: "6px 10px", borderRadius: 6 }}>📋 {item.next_plan}</div>}
                {item.reported_by && <div style={{ fontSize: 11, color: C.gray400, marginTop: 4 }}>রিপোর্টকারী: {item.reported_by}</div>}
                {item.image_url && <img src={item.image_url} alt="Work" style={{ maxWidth: "100%", maxHeight: 150, borderRadius: 8, marginTop: 8, objectFit: "cover" }} />}
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
            <FormField label="শ্রমিক/কারিগর সংখ্যা"><input style={inputStyle} type="number" value={form.workers_count} onChange={e => setForm({ ...form, workers_count: e.target.value })} /></FormField>
            <FormField label="অগ্রগতি (%)"><input style={inputStyle} type="number" min="0" max="100" value={form.progress_pct} onChange={e => setForm({ ...form, progress_pct: e.target.value })} /></FormField>
            <FormField label="রিপোর্টকারী"><input style={inputStyle} value={form.reported_by} onChange={e => setForm({ ...form, reported_by: e.target.value })} placeholder="ডিজাইনার/সুপারভাইজার" /></FormField>
            <div></div>
            <div style={{ gridColumn: "1/-1" }}><FormField label="সমস্যা / ইস্যু"><textarea style={{ ...inputStyle, height: 60, resize: "vertical" }} value={form.issues} onChange={e => setForm({ ...form, issues: e.target.value })} /></FormField></div>
            <div style={{ gridColumn: "1/-1" }}><FormField label="পরবর্তী পরিকল্পনা"><textarea style={{ ...inputStyle, height: 60, resize: "vertical" }} value={form.next_plan} onChange={e => setForm({ ...form, next_plan: e.target.value })} /></FormField></div>
            <div style={{ gridColumn: "1/-1" }}><ImageUploadField label="📷 কাজের ছবি (ঐচ্ছিক)" value={form.image_url} onChange={url => setForm(f => ({ ...f, image_url: url }))} folder="interior-daily" /></div>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ---- INTERIOR EXPENSES ----
function IPExpenses({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [showMultiModal, setShowMultiModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { expense_date: new Date().toISOString().split("T")[0], category: "Interior Materials", item_name: "", description: "", unit: "পিস", quantity: 1, unit_price: "", amount: 0, supplier: "", payment_method: "নগদ", payment_status: "পরিশোধিত", received_by: "", note: "", image_url: "" };
  const [form, setForm] = useState(emptyForm);
  const [rows, setRows] = useState([{ ...emptyForm }]);
  const categories = ["Interior Materials", "Interior Labour", "Furniture", "False Ceiling", "Wall Panelling", "Flooring", "Paint", "Electrical", "Plumbing", "Kitchen Cabinet", "Wardrobe", "Glass Work", "Curtain/Blind", "Lighting", "Hardware", "Transport", "অন্যান্য"];
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("interior_expenses").select("*").eq("project_id", projectId).order("expense_date", { ascending: false }); setItems(data || []); };
  const save = async () => {
    if (!form.item_name || !form.unit_price) return alert("আইটেম ও মূল্য আবশ্যক");
    const qty = +form.quantity || 1; const price = +form.unit_price || 0;
    const payload = { ...form, project_id: projectId, quantity: qty, unit_price: price, amount: qty * price };
    if (editItem) { await supabase.from("interior_expenses").update(payload).eq("id", editItem.id); }
    else { await supabase.from("interior_expenses").insert([payload]); }
    await load(); setShowModal(false); setEditItem(null);
  };
  const addRow = () => setRows(r => [...r, { ...emptyForm, expense_date: r[r.length - 1]?.expense_date || emptyForm.expense_date }]);
  const removeRow = (idx) => setRows(r => r.filter((_, i) => i !== idx));
  const updateRow = (idx, field, val) => setRows(r => r.map((row, i) => i === idx ? { ...row, [field]: val } : row));
  const saveMultiRows = async () => {
    const valid = rows.filter(r => r.item_name && r.unit_price);
    if (valid.length === 0) return alert("কমপক্ষে একটি আইটেমের নাম ও মূল্য দিন!");
    const payloads = valid.map(r => ({ ...r, project_id: projectId, quantity: +r.quantity || 1, unit_price: +r.unit_price || 0, amount: (+r.quantity || 1) * (+r.unit_price || 0) }));
    await supabase.from("interior_expenses").insert(payloads);
    await load(); setShowMultiModal(false);
    setRows([{ ...emptyForm }]);
  };
  const del = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("interior_expenses").delete().eq("id", id); await load(); };
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
      <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
        <button onClick={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} style={{ background: C.primary, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ একটি খরচ</button>
        <button onClick={() => { setShowMultiModal(true); }} style={{ background: "#2A5C8F", color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 13 }}>+ একসাথে অনেক খরচ</button>
        <button onClick={() => exportToExcel(items.map(i => ({ তারিখ: i.expense_date, ক্যাটাগরি: i.category, আইটেম: i.item_name, বিবরণ: i.description, একক: i.unit, পরিমাণ: i.quantity, একক_মূল্য: i.unit_price, মোট: i.amount, সাপ্লায়ার: i.supplier, স্ট্যাটাস: i.payment_status })), "IPExpenses", "Interior_Expenses")} style={{ background: C.green, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 12 }}>📊 Excel</button>
        <button onClick={() => { printSection("Interior Project Expenses", "ip-expenses-print"); }} style={{ background: C.blue, color: C.white, border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 600, cursor: "pointer", fontSize: 12 }}>🖨️ Print</button>
      </div>
      {Object.keys(catGroups).length > 0 && (
        <Card style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 13, marginBottom: 8 }}>ক্যাটাগরি অনুযায়ী:</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {Object.entries(catGroups).sort((a, b) => b[1] - a[1]).map(([cat, total]) => (
              <div key={cat} style={{ background: C.primaryBg, border: "1px solid " + C.primaryLight, borderRadius: 8, padding: "6px 12px" }}>
                <div style={{ fontSize: 11, color: C.gray600 }}>{cat}</div>
                <div style={{ fontSize: 13, fontWeight: 700, color: C.primaryDark }}>{fmt(total)}</div>
              </div>
            ))}
          </div>
        </Card>
      )}
      <Card>
        <div id="ip-expenses-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
            <thead><tr style={{ background: C.primaryBg }}>
              {["তারিখ", "ক্যাটাগরি", "আইটেম", "বিবরণ", "পরিমাণ", "একক মূল্য", "মোট", "সাপ্লায়ার"].map(h => <th key={h} style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>)}
              <th className="no-print" style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>স্ট্যাটাস</th>
              <th className="no-print" style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>Action</th>
            </tr></thead>
            <tbody>
              {(() => {
                const dateGroups = items.reduce((acc, i) => { const d = i.expense_date || "—"; if (!acc[d]) acc[d] = []; acc[d].push(i); return acc; }, {});
                return Object.entries(dateGroups).map(([date, dateItems]) => {
                  const dateTotal = dateItems.reduce((s, i) => s + (i.amount || 0), 0);
                  return (
                    <Fragment key={date}>
                      {dateItems.map(item => (
                        <tr key={item.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                          <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{item.expense_date}</td>
                          <td style={{ padding: "8px 10px" }}><Badge label={item.category} color="primary" /></td>
                          <td style={{ padding: "8px 10px", fontWeight: 600, color: C.primaryDark }}>{item.item_name}</td>
                          <td style={{ padding: "8px 10px", fontSize: 11, color: C.gray600 }}>{item.description || "—"}</td>
                          <td style={{ padding: "8px 10px" }}>{item.quantity} {item.unit}</td>
                          <td style={{ padding: "8px 10px" }}>{fmt(item.unit_price)}</td>
                          <td style={{ padding: "8px 10px", fontWeight: 700, color: C.red }}>{fmt(item.amount)}</td>
                          <td style={{ padding: "8px 10px", fontSize: 11 }}>{item.supplier || "—"}</td>
                          <td className="no-print" style={{ padding: "8px 10px" }}><Badge label={item.payment_status} color={item.payment_status === "পরিশোধিত" ? "green" : "yellow"} /></td>
                          <td className="no-print" style={{ padding: "8px 10px" }}><div style={{ display: "flex", gap: 4 }}><button onClick={() => { setEditItem(item); setForm({ ...item }); setShowModal(true); }} style={btnEdit}>✏️</button><button onClick={() => del(item.id)} style={btnDanger}>🗑️</button></div></td>
                        </tr>
                      ))}
                      {dateItems.length > 1 && (
                        <tr style={{ background: C.gray50, fontWeight: 700 }}>
                          <td colSpan={6} style={{ padding: "8px 10px", textAlign: "right", color: C.primaryDark, fontSize: 12 }}>{date} — এই দিনের মোট:</td>
                          <td style={{ padding: "8px 10px", color: C.red, fontSize: 13 }}>{fmt(dateTotal)}</td>
                          <td className="no-print" colSpan={2}></td>
                        </tr>
                      )}
                    </Fragment>
                  );
                });
              })()}
              {items.length > 0 && <tr style={{ background: C.primaryBg, fontWeight: 700 }}><td colSpan={6} style={{ padding: "10px", textAlign: "right", color: C.primaryDark }}>সর্বমোট:</td><td style={{ padding: "10px", color: C.red, fontSize: 15 }}>{fmt(totalExpense)}</td><td></td><td className="no-print" colSpan={2}></td></tr>}
            </tbody>
          </table>
          {items.length === 0 && <div style={{ textAlign: "center", padding: 30, color: C.gray400 }}>কোনো খরচ নেই!</div>}
        </div>
      </Card>
      {showMultiModal && (
        <Modal title="একসাথে একাধিক খরচ" onClose={() => setShowMultiModal(false)} size={900}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "ক্যাটাগরি", "আইটেম *", "বিবরণ", "একক", "পরিমাণ", "মূল্য *", "মোট", "স্ট্যাটাস", ""].map(h => <th key={h} style={{ padding: "8px 6px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
              <tbody>
                {rows.map((row, idx) => (
                  <tr key={idx} style={{ borderBottom: "1px solid " + C.gray100 }}>
                    <td style={{ padding: "4px" }}><input type="date" value={row.expense_date} onChange={e => updateRow(idx, "expense_date", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 110 }} /></td>
                    <td style={{ padding: "4px" }}><select value={row.category} onChange={e => updateRow(idx, "category", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 130 }}>{categories.map(c => <option key={c}>{c}</option>)}</select></td>
                    <td style={{ padding: "4px" }}><input value={row.item_name} onChange={e => updateRow(idx, "item_name", e.target.value)} placeholder="আইটেম *" style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 120 }} /></td>
                    <td style={{ padding: "4px" }}><input value={row.description} onChange={e => updateRow(idx, "description", e.target.value)} placeholder="বিবরণ" style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 100 }} /></td>
                    <td style={{ padding: "4px" }}><select value={row.unit} onChange={e => updateRow(idx, "unit", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 60 }}>{["পিস", "বর্গফুট", "রানিংফুট", "কেজি", "লিটার", "সেট", "লট", "দিন"].map(u => <option key={u}>{u}</option>)}</select></td>
                    <td style={{ padding: "4px" }}><input type="number" value={row.quantity} onChange={e => updateRow(idx, "quantity", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, width: 60 }} /></td>
                    <td style={{ padding: "4px" }}><input type="number" value={row.unit_price} onChange={e => updateRow(idx, "unit_price", e.target.value)} placeholder="মূল্য *" style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, width: 80 }} /></td>
                    <td style={{ padding: "4px", fontWeight: 700, color: C.primaryDark, whiteSpace: "nowrap" }}>{fmt((+row.quantity || 1) * (+row.unit_price || 0))}</td>
                    <td style={{ padding: "4px" }}><select value={row.payment_status} onChange={e => updateRow(idx, "payment_status", e.target.value)} style={{ ...inputStyle, padding: "4px 6px", fontSize: 11, minWidth: 80 }}>{["পরিশোধিত","বকেয়া","আংশিক"].map(s => <option key={s}>{s}</option>)}</select></td>
                    <td style={{ padding: "4px" }}>{rows.length > 1 && <button onClick={() => removeRow(idx)} style={{ ...btnDanger, padding: "4px 8px" }}>×</button>}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot><tr style={{ background: C.primaryBg }}><td colSpan={7} style={{ padding: "8px", textAlign: "right", fontWeight: 700, color: C.primaryDark }}>সর্বমোট:</td><td style={{ padding: "8px", fontWeight: 800, color: C.red, fontSize: 14 }}>{fmt(rows.reduce((s, r) => s + ((+r.quantity || 1) * (+r.unit_price || 0)), 0))}</td><td colSpan={2}></td></tr></tfoot>
            </table>
          </div>
          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <button onClick={addRow} style={{ background: C.primaryBg, color: C.primaryDark, border: "1px solid " + C.primary, borderRadius: 8, padding: "10px 20px", fontWeight: 600, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>+ নতুন সারি</button>
            <button onClick={saveMultiRows} style={{ ...btnPrimary, width: "auto", padding: "10px 24px" }}>✅ সব সংরক্ষণ ({rows.filter(r => r.item_name && r.unit_price).length}টি)</button>
          </div>
        </Modal>
      )}
      {showModal && (
        <Modal title={editItem ? "খরচ সম্পাদনা" : "নতুন খরচ"} onClose={() => { setShowModal(false); setEditItem(null); }} size={600}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="তারিখ *"><input style={inputStyle} type="date" value={form.expense_date} onChange={e => setForm({ ...form, expense_date: e.target.value })} /></FormField>
            <FormField label="ক্যাটাগরি *"><select style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{categories.map(c => <option key={c}>{c}</option>)}</select></FormField>
            <div style={{ gridColumn: "1/-1" }}><FormField label="আইটেমের নাম *"><input style={inputStyle} value={form.item_name} onChange={e => setForm({ ...form, item_name: e.target.value })} placeholder="যেমন: Board, Paint, Labour..." /></FormField></div>
            <div style={{ gridColumn: "1/-1" }}><FormField label="বিবরণ"><input style={inputStyle} value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} placeholder="ব্র্যান্ড, স্পেসিফিকেশন..." /></FormField></div>
            <FormField label="একক"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{["পিস", "বর্গফুট", "রানিংফুট", "কেজি", "লিটার", "সেট", "লট", "দিন"].map(u => <option key={u}>{u}</option>)}</select></FormField>
            <FormField label="পরিমাণ *"><input style={inputStyle} type="number" value={form.quantity} onChange={e => { const q = e.target.value; setForm(f => ({ ...f, quantity: q, amount: (q * (+f.unit_price || 0)).toFixed(2) })); }} /></FormField>
            <FormField label="একক মূল্য (৳) *"><input style={inputStyle} type="number" value={form.unit_price} onChange={e => { const p = e.target.value; setForm(f => ({ ...f, unit_price: p, amount: ((+f.quantity || 1) * p).toFixed(2) })); }} /></FormField>
            <div style={{ background: C.primaryBg, borderRadius: 8, padding: "10px", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column" }}>
              <div style={{ fontSize: 11, color: C.gray600 }}>মোট</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: C.primaryDark }}>{fmt(+form.amount || (+form.quantity || 1) * (+form.unit_price || 0))}</div>
            </div>
            <FormField label="সাপ্লায়ার"><input style={inputStyle} value={form.supplier} onChange={e => setForm({ ...form, supplier: e.target.value })} /></FormField>
            <FormField label="পেমেন্ট পদ্ধতি"><select style={inputStyle} value={form.payment_method} onChange={e => setForm({ ...form, payment_method: e.target.value })}>{["নগদ", "বিকাশ", "ব্যাংক", "চেক", "বকেয়া"].map(m => <option key={m}>{m}</option>)}</select></FormField>
            <FormField label="পেমেন্ট স্ট্যাটাস"><select style={inputStyle} value={form.payment_status} onChange={e => setForm({ ...form, payment_status: e.target.value })}>{["পরিশোধিত", "বকেয়া", "আংশিক"].map(s => <option key={s}>{s}</option>)}</select></FormField>
            <FormField label="গ্রহণকারী"><input style={inputStyle} value={form.received_by} onChange={e => setForm({ ...form, received_by: e.target.value })} /></FormField>
            <div style={{ gridColumn: "1/-1" }}><FormField label="নোট"><input style={inputStyle} value={form.note} onChange={e => setForm({ ...form, note: e.target.value })} /></FormField></div>
            <div style={{ gridColumn: "1/-1" }}><ImageUploadField label="📷 রসিদ / ছবি" value={form.image_url} onChange={url => setForm(f => ({ ...f, image_url: url }))} folder="interior-expenses" /></div>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ---- INTERIOR STOCK ----
function IPStock({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { item_name: "", category: "ইন্টেরিয়র সামগ্রী", unit: "পিস", opening_stock: 0, received: 0, used: 0, unit_price: "", supplier: "", last_updated: new Date().toISOString().split("T")[0], min_stock: 0 };
  const [form, setForm] = useState(emptyForm);
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("interior_stock").select("*").eq("project_id", projectId).order("category"); setItems(data || []); };
  const save = async () => {
    if (!form.item_name) return alert("নাম আবশ্যক");
    const received = +form.received || 0; const opening = +form.opening_stock || 0; const used = +form.used || 0;
    const closing = opening + received - used; const price = +form.unit_price || 0;
    const payload = { ...form, project_id: projectId, opening_stock: opening, received, used, closing_stock: closing, unit_price: price, total_value: closing * price, min_stock: +form.min_stock || 0 };
    if (editItem) { await supabase.from("interior_stock").update(payload).eq("id", editItem.id); }
    else { await supabase.from("interior_stock").insert([payload]); }
    await load(); setShowModal(false); setEditItem(null);
  };
  const del = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("interior_stock").delete().eq("id", id); await load(); };
  const totalValue = items.reduce((s, i) => s + (i.total_value || 0), 0);
  const cats = ["ইন্টেরিয়র সামগ্রী", "Board/Sheet", "Paint", "Hardware", "Lighting", "Glass", "Fabric", "অন্যান্য"];
  return (
    <div>
      <SectionHeader title="🪑 Stock Register" action="নতুন আইটেম" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onExport={() => exportToExcel(items.map(i => ({ আইটেম: i.item_name, ক্যাটাগরি: i.category, একক: i.unit, প্রারম্ভিক: i.opening_stock, প্রাপ্ত: i.received, ব্যবহৃত: i.used, অবশিষ্ট: i.closing_stock, মূল্য: i.unit_price, মোট: i.total_value })), "IPStock", "Interior_Stock")} onPrint={() => { printSection("Interior Stock Register", "ip-stock-print"); }} />
      <Card>
        <div id="ip-stock-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["আইটেম", "ক্যাটাগরি", "একক", "প্রারম্ভিক", "প্রাপ্ত", "ব্যবহৃত", "অবশিষ্ট", "মূল্য", "মোট মূল্য", "Action"].map(h => <th key={h} style={{ padding: "9px 10px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
            <tbody>
              {items.map(item => (
                <tr key={item.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "8px 10px", fontWeight: 600, color: C.primaryDark }}>{item.item_name}</td>
                  <td style={{ padding: "8px 10px" }}><Badge label={item.category} color="primary" /></td>
                  <td style={{ padding: "8px 10px" }}>{item.unit}</td>
                  <td style={{ padding: "8px 10px", textAlign: "center" }}>{item.opening_stock}</td>
                  <td style={{ padding: "8px 10px", textAlign: "center", color: C.green, fontWeight: 600 }}>+{item.received}</td>
                  <td style={{ padding: "8px 10px", textAlign: "center", color: C.red, fontWeight: 600 }}>-{item.used}</td>
                  <td style={{ padding: "8px 10px", textAlign: "center", fontWeight: 700 }}>{item.closing_stock}</td>
                  <td style={{ padding: "8px 10px" }}>{fmt(item.unit_price)}</td>
                  <td style={{ padding: "8px 10px", fontWeight: 700, color: C.primary }}>{fmt(item.total_value)}</td>
                  <td style={{ padding: "8px 10px" }}><div style={{ display: "flex", gap: 4 }}><button onClick={() => { setEditItem(item); setForm({ ...item }); setShowModal(true); }} style={btnEdit}>✏️</button><button onClick={() => del(item.id)} style={btnDanger}>🗑️</button></div></td>
                </tr>
              ))}
              {items.length > 0 && <tr style={{ background: C.primaryBg, fontWeight: 700 }}><td colSpan={8} style={{ padding: "10px", textAlign: "right", color: C.primaryDark }}>মোট:</td><td style={{ padding: "10px", color: C.primary, fontSize: 15 }}>{fmt(totalValue)}</td><td></td></tr>}
            </tbody>
          </table>
          {items.length === 0 && <div style={{ textAlign: "center", padding: 30, color: C.gray400 }}>কোনো স্টক নেই!</div>}
        </div>
      </Card>
      {showModal && (
        <Modal title={editItem ? "স্টক সম্পাদনা" : "নতুন স্টক আইটেম"} onClose={() => { setShowModal(false); setEditItem(null); }} size={540}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <div style={{ gridColumn: "1/-1" }}><FormField label="আইটেমের নাম *"><input style={inputStyle} value={form.item_name} onChange={e => setForm({ ...form, item_name: e.target.value })} /></FormField></div>
            <FormField label="ক্যাটাগরি"><select style={inputStyle} value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{cats.map(c => <option key={c}>{c}</option>)}</select></FormField>
            <FormField label="একক"><select style={inputStyle} value={form.unit} onChange={e => setForm({ ...form, unit: e.target.value })}>{["পিস", "বর্গফুট", "রানিংফুট", "কেজি", "লিটার", "সেট", "লট"].map(u => <option key={u}>{u}</option>)}</select></FormField>
            <FormField label="প্রারম্ভিক স্টক"><input style={inputStyle} type="number" value={form.opening_stock} onChange={e => setForm({ ...form, opening_stock: e.target.value })} /></FormField>
            <FormField label="নতুন প্রাপ্তি"><input style={inputStyle} type="number" value={form.received} onChange={e => setForm({ ...form, received: e.target.value })} /></FormField>
            <FormField label="ব্যবহৃত"><input style={inputStyle} type="number" value={form.used} onChange={e => setForm({ ...form, used: e.target.value })} /></FormField>
            <div style={{ background: C.primaryBg, borderRadius: 8, padding: "10px", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column" }}>
              <div style={{ fontSize: 11, color: C.gray600 }}>অবশিষ্ট</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: C.primaryDark }}>{(+form.opening_stock || 0) + (+form.received || 0) - (+form.used || 0)} {form.unit}</div>
            </div>
            <FormField label="একক মূল্য (৳)"><input style={inputStyle} type="number" value={form.unit_price} onChange={e => setForm({ ...form, unit_price: e.target.value })} /></FormField>
            <FormField label="সাপ্লায়ার"><input style={inputStyle} value={form.supplier} onChange={e => setForm({ ...form, supplier: e.target.value })} /></FormField>
          </div>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ---- INTERIOR PAYMENTS ----
function IPPayments({ projectId }) {
  const [items, setItems] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = { receive_date: new Date().toISOString().split("T")[0], amount: "", payment_type: "১ম কিস্তি", payment_method: "নগদ", received_by: "", sender_name: "", bkash_ref: "", note: "" };
  const [form, setForm] = useState(emptyForm);
  useEffect(() => { load(); }, [projectId]);
  const load = async () => { const { data } = await supabase.from("interior_received_payments").select("*").eq("project_id", projectId).order("receive_date", { ascending: false }); setItems(data || []); };
  const save = async () => {
    if (!form.amount) return alert("পরিমাণ আবশ্যক");
    const payload = { ...form, project_id: projectId, amount: +form.amount };
    if (editItem) { await supabase.from("interior_received_payments").update(payload).eq("id", editItem.id); }
    else { await supabase.from("interior_received_payments").insert([payload]); }
    await load(); setShowModal(false); setEditItem(null);
  };
  const del = async (id) => { if (!confirm("মুছবেন?")) return; await supabase.from("interior_received_payments").delete().eq("id", id); await load(); };
  const totalReceived = items.reduce((s, i) => s + (i.amount || 0), 0);
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 12, marginBottom: 16 }}>
        <StatCard icon="💰" label="মোট প্রাপ্ত" value={fmt(totalReceived)} color="#F0FFF4" />
        <StatCard icon="📋" label="মোট কিস্তি" value={fmtNum(items.length)} color={C.primaryBg} />
      </div>
      <SectionHeader title="💰 Office থেকে প্রাপ্ত টাকা" action="নতুন প্রাপ্তি" onAction={() => { setEditItem(null); setForm(emptyForm); setShowModal(true); }} onExport={() => exportToExcel(items.map(i => ({ তারিখ: i.receive_date, পরিমাণ: i.amount, কিস্তি: i.payment_type, পদ্ধতি: i.payment_method, প্রেরক: i.sender_name, গ্রহণকারী: i.received_by, রেফারেন্স: i.bkash_ref || "", নোট: i.note })), "IPPayments", "Interior_Payments")} onPrint={() => { printSection("Interior Payment Register", "ip-payments-print"); }} />
      <Card>
        <div id="ip-payments-print" style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: C.primaryBg }}>{["তারিখ", "কিস্তির ধরন", "পরিমাণ", "পেমেন্ট", "প্রেরক", "গ্রহণকারী", "রেফারেন্স", "নোট", "Action"].map(h => <th key={h} style={{ padding: "9px 12px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
            <tbody>
              {items.map(item => (
                <tr key={item.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "9px 12px" }}>{item.receive_date}</td>
                  <td style={{ padding: "9px 12px" }}><Badge label={item.payment_type} color="primary" /></td>
                  <td style={{ padding: "9px 12px", fontWeight: 700, color: C.green }}>+{fmt(item.amount)}</td>
                  <td style={{ padding: "9px 12px" }}>{item.payment_method}</td>
                  <td style={{ padding: "9px 12px" }}>{item.sender_name || "—"}</td>
                  <td style={{ padding: "9px 12px" }}>{item.received_by || "—"}</td>
                  <td style={{ padding: "9px 12px", fontSize: 11, color: C.gray400 }}>{item.bkash_ref || "—"}</td>
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
          <FormField label="কিস্তির ধরন"><select style={inputStyle} value={form.payment_type} onChange={e => setForm({ ...form, payment_type: e.target.value })}>{["১ম কিস্তি", "২য় কিস্তি", "৩য় কিস্তি", "৪র্থ কিস্তি", "চূড়ান্ত কিস্তি", "জরুরি অগ্রিম", "অন্যান্য"].map(t => <option key={t}>{t}</option>)}</select></FormField>
          <FormField label="পরিমাণ (৳) *"><input style={inputStyle} type="number" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} /></FormField>
          <FormField label="পেমেন্ট পদ্ধতি"><select style={inputStyle} value={form.payment_method} onChange={e => setForm({ ...form, payment_method: e.target.value })}>{["নগদ", "বিকাশ", "নগদ (মোবাইল)", "ব্যাংক ট্রান্সফার", "চেক"].map(m => <option key={m}>{m}</option>)}</select></FormField>
          <FormField label="প্রেরকের নাম"><input style={inputStyle} value={form.sender_name} onChange={e => setForm({ ...form, sender_name: e.target.value })} /></FormField>
          <FormField label="গ্রহণকারীর নাম"><input style={inputStyle} value={form.received_by} onChange={e => setForm({ ...form, received_by: e.target.value })} /></FormField>
          <FormField label="বিকাশ/ব্যাংক রেফারেন্স"><input style={inputStyle} value={form.bkash_ref} onChange={e => setForm({ ...form, bkash_ref: e.target.value })} /></FormField>
          <FormField label="নোট"><input style={inputStyle} value={form.note} onChange={e => setForm({ ...form, note: e.target.value })} /></FormField>
          <button onClick={save} style={btnPrimary}>✅ সংরক্ষণ করুন</button>
        </Modal>
      )}
    </div>
  );
}

// ---- INTERIOR SUMMARY ----
function IPSummary({ project }) {
  const [expenses, setExpenses] = useState([]);
  const [payments, setPayments] = useState([]);
  const [stock, setStock] = useState([]);
  const [updates, setUpdates] = useState([]);
  useEffect(() => {
    const load = async () => {
      const [e, p, s, d] = await Promise.all([
        supabase.from("interior_expenses").select("*").eq("project_id", project.id),
        supabase.from("interior_received_payments").select("*").eq("project_id", project.id),
        supabase.from("interior_stock").select("*").eq("project_id", project.id),
        supabase.from("interior_daily_updates").select("*").eq("project_id", project.id),
      ]);
      setExpenses(e.data || []); setPayments(p.data || []); setStock(s.data || []); setUpdates(d.data || []);
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
  const lastProgress = updates.length > 0 ? updates.sort((a, b) => new Date(b.update_date) - new Date(a.update_date))[0]?.progress_pct || 0 : 0;
  return (
    <div>
      <SectionHeader title="📊 Project Summary" onPrint={() => { printSection("Interior Project Summary", "ip-summary-print"); }} />
      <div id="ip-summary-print">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 20 }}>
          <StatCard icon="💰" label="চুক্তি মূল্য" value={fmt(dealAmount)} color={C.primaryBg} />
          <StatCard icon="📥" label="Office থেকে প্রাপ্ত" value={fmt(totalReceived)} color="#F0FFF4" />
          <StatCard icon="💸" label="মোট খরচ" value={fmt(totalExpense)} color="#FFF5F5" />
          <StatCard icon="🏦" label="হাতে ব্যালেন্স" value={fmt(balance)} color={balance >= 0 ? "#F0FFF4" : "#FFF5F5"} />
          <StatCard icon="📦" label="স্টক মূল্য" value={fmt(totalStockValue)} color="#FFF8E1" />
          <StatCard icon="📈" label="প্রজেক্টেড লাভ" value={fmt(projectedProfit)} color={projectedProfit >= 0 ? "#F0FFF4" : "#FFF5F5"} />
        </div>
        <Card style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14, marginBottom: 10 }}>সার্বিক অগ্রগতি</div>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
            <span style={{ fontSize: 12 }}>সর্বশেষ আপডেট অনুযায়ী</span>
            <span style={{ fontSize: 12, fontWeight: 700, color: C.primary }}>{lastProgress}%</span>
          </div>
          <ProgressBar value={lastProgress} />
        </Card>
        <Card style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14, marginBottom: 10 }}>বাজেট ব্যবহার</div>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
            <span style={{ fontSize: 12 }}>মোট খরচ: {fmt(totalExpense)}</span>
            <span style={{ fontSize: 12, fontWeight: 700 }}>{dealAmount > 0 ? Math.round((totalExpense / dealAmount) * 100) : 0}%</span>
          </div>
          <ProgressBar value={dealAmount > 0 ? Math.min(Math.round((totalExpense / dealAmount) * 100), 100) : 0} color={totalExpense > dealAmount ? C.red : C.primary} />
        </Card>
        <Card style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14, marginBottom: 10 }}>ক্যাটাগরি অনুযায়ী খরচ</div>
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
          <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 14, marginBottom: 10 }}>প্রজেক্টের তথ্য</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, fontSize: 13 }}>
            {[["প্রজেক্টের নাম", project.name], ["ক্লায়েন্ট", project.client_name], ["ঠিকানা", project.site_address], ["ধরন", project.project_type], ["প্রধান ডিজাইনার", project.chief_designer], ["শুরুর তারিখ", project.start_date], ["শেষের তারিখ", project.end_date], ["স্ট্যাটাস", project.status], ["Daily Updates", updates.length + "টি"]].map(([label, val]) => (
              <div key={label}><div style={{ fontSize: 11, color: C.gray400 }}>{label}</div><div style={{ fontWeight: 600, color: C.primaryDark }}>{val || "—"}</div></div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

// ============================================================
// PROFILE DROPDOWN
// ============================================================
function ProfileDropdown({ currentUser, onUpdate, onClose }) {
  const [tab, setTab] = useState("name");
  const [name, setName] = useState(currentUser?.name || "");
  const [oldPass, setOldPass] = useState("");
  const [newPass, setNewPass] = useState("");
  const [confirmPass, setConfirmPass] = useState("");
  const [msg, setMsg] = useState("");
  const [error, setError] = useState("");
  const [imgPreview, setImgPreview] = useState(localStorage.getItem("nic_profile_pic") || null);

  const saveName = async () => {
    if (!name.trim()) return setError("নাম দিন!");
    const { error: dbErr } = await supabase.from("app_users").update({ name }).eq("email", currentUser.email);
    if (dbErr) return setError("Error: " + dbErr.message);
    const updated = { ...currentUser, name };
    localStorage.setItem("nic_user", JSON.stringify(updated));
    setMsg("✅ নাম পরিবর্তন হয়েছে!"); setError("");
    onUpdate(updated);
  };

  const savePassword = async () => {
    setMsg(""); setError("");
    const currentPass = localStorage.getItem("nic_password") || "noksha2024";
    if (oldPass !== currentPass && oldPass !== currentUser?.password_hash) return setError("❌ বর্তমান পাসওয়ার্ড ভুল!");
    if (newPass.length < 6) return setError("❌ কমপক্ষে ৬ অক্ষর দিন!");
    if (newPass !== confirmPass) return setError("❌ পাসওয়ার্ড মিলছে না!");
    await supabase.from("app_users").update({ password_hash: newPass }).eq("email", currentUser.email);
    localStorage.setItem("nic_password", newPass);
    setMsg("✅ পাসওয়ার্ড পরিবর্তন হয়েছে!"); setError("");
    setOldPass(""); setNewPass(""); setConfirmPass("");
  };

  const handleImageUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (file.size > 5 * 1024 * 1024) return setError("❌ ছবি সর্বোচ্চ 5MB হতে হবে!");
    const reader = new FileReader();
    reader.onload = (ev) => {
      const b64 = ev.target.result;
      localStorage.setItem("nic_profile_pic", b64);
      setImgPreview(b64);
      setMsg("✅ Profile ছবি আপডেট হয়েছে!");
      onUpdate({ ...currentUser, profile_pic: b64 });
    };
    reader.readAsDataURL(file);
  };

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 2000, display: "flex", alignItems: "flex-start", justifyContent: "flex-end", padding: "64px 24px 0" }} onClick={onClose}>
      <div style={{ background: C.white, borderRadius: 16, width: 340, boxShadow: "0 8px 40px rgba(0,0,0,0.18)", border: "1px solid " + C.gray200, overflow: "hidden" }} onClick={e => e.stopPropagation()}>
        {/* Profile Header */}
        <div style={{ background: "linear-gradient(135deg, " + C.primaryDark + ", " + C.primary + ")", padding: "20px", textAlign: "center" }}>
          <div style={{ width: 64, height: 64, borderRadius: "50%", margin: "0 auto 10px", overflow: "hidden", border: "3px solid rgba(255,255,255,0.4)", background: "rgba(255,255,255,0.2)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            {imgPreview ? <img src={imgPreview} alt="Profile" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <span style={{ fontSize: 28, color: C.white, fontWeight: 700 }}>{currentUser?.name?.[0] || "R"}</span>}
          </div>
          <div style={{ color: C.white, fontWeight: 700, fontSize: 15 }}>{currentUser?.name}</div>
          <div style={{ color: "rgba(255,255,255,0.7)", fontSize: 12 }}>{currentUser?.email}</div>
          <div style={{ color: "rgba(255,255,255,0.6)", fontSize: 11, marginTop: 3 }}>{currentUser?.role === "admin" ? "🔑 Admin" : "👷 Site Engineer"}</div>
        </div>
        {/* Tabs */}
        <div style={{ display: "flex", borderBottom: "1px solid " + C.gray200 }}>
          {[["name", "👤 নাম"], ["password", "🔐 পাসওয়ার্ড"], ["photo", "🖼️ ছবি"]].map(([id, label]) => (
            <button key={id} onClick={() => { setTab(id); setMsg(""); setError(""); }} style={{ flex: 1, padding: "10px 4px", border: "none", borderBottom: tab === id ? "2px solid " + C.primary : "2px solid transparent", background: "none", color: tab === id ? C.primaryDark : C.gray600, fontWeight: tab === id ? 700 : 400, cursor: "pointer", fontSize: 12, fontFamily: "inherit" }}>{label}</button>
          ))}
        </div>
        <div style={{ padding: 20 }}>
          {error && <div style={{ background: C.redLight, color: C.red, padding: "8px 12px", borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{error}</div>}
          {msg && <div style={{ background: C.greenLight, color: C.green, padding: "8px 12px", borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{msg}</div>}
          {tab === "name" && (
            <div>
              <FormField label="নতুন নাম"><input style={inputStyle} value={name} onChange={e => setName(e.target.value)} placeholder="আপনার নাম" /></FormField>
              <button onClick={saveName} style={{ ...btnPrimary, marginTop: 4 }}>✅ নাম পরিবর্তন করুন</button>
            </div>
          )}
          {tab === "password" && (
            <div>
              <FormField label="বর্তমান পাসওয়ার্ড"><input style={inputStyle} type="password" value={oldPass} onChange={e => setOldPass(e.target.value)} /></FormField>
              <FormField label="নতুন পাসওয়ার্ড"><input style={inputStyle} type="password" value={newPass} onChange={e => setNewPass(e.target.value)} placeholder="কমপক্ষে ৬ অক্ষর" /></FormField>
              <FormField label="নিশ্চিত করুন"><input style={inputStyle} type="password" value={confirmPass} onChange={e => setConfirmPass(e.target.value)} /></FormField>
              <button onClick={savePassword} style={{ ...btnPrimary, marginTop: 4 }}>🔑 পাসওয়ার্ড পরিবর্তন করুন</button>
            </div>
          )}
          {tab === "photo" && (
            <div style={{ textAlign: "center" }}>
              <div style={{ width: 80, height: 80, borderRadius: "50%", margin: "0 auto 14px", overflow: "hidden", border: "3px solid " + C.primary, background: C.primaryBg, display: "flex", alignItems: "center", justifyContent: "center" }}>
                {imgPreview ? <img src={imgPreview} alt="Profile" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <span style={{ fontSize: 36 }}>👤</span>}
              </div>
              <label style={{ display: "inline-block", background: C.primary, color: C.white, padding: "10px 20px", borderRadius: 8, cursor: "pointer", fontWeight: 600, fontSize: 13 }}>
                📷 ছবি বেছে নিন
                <input type="file" accept="image/*" onChange={handleImageUpload} style={{ display: "none" }} />
              </label>
              <div style={{ fontSize: 11, color: C.gray400, marginTop: 8 }}>সর্বোচ্চ 5MB</div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ============================================================
// USER MANAGEMENT (Admin only)
// ============================================================
function UserManagement({ employees, lang }) {
  const T = TXT[lang || "bn"];
  const [users, setUsers] = useState([]);
  const [projects, setProjects] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const blankForm = { name: "", email: "", password_hash: "", role: "site_engineer", assigned_projects: [], employee_id: "", permissions: [], is_active: true };
  const [form, setForm] = useState(blankForm);

  useEffect(() => { loadUsers(); loadProjects(); }, []);
  const loadUsers = async () => { const { data } = await supabase.from("app_users").select("*").order("created_at"); setUsers(data || []); };
  const loadProjects = async () => { const { data } = await supabase.from("construction_projects").select("id, name, client_name").order("created_at"); setProjects(data || []); };

  const save = async () => {
    if (!form.name || !form.email || !form.password_hash) return alert("নাম, ইমেইল ও পাসওয়ার্ড আবশ্যক");
    const payload = { ...form, employee_id: form.employee_id || null };
    if (editItem) {
      await supabase.from("app_users").update(payload).eq("id", editItem.id);
    } else {
      const { error } = await supabase.from("app_users").insert([payload]);
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

  const togglePermission = (menuId) => {
    setForm(f => {
      const arr = f.permissions || [];
      return { ...f, permissions: arr.includes(menuId) ? arr.filter(x => x !== menuId) : [...arr, menuId] };
    });
  };

  const roleLabel = { admin: "🔑 Admin", site_engineer: "👷 Site Engineer", employee: T.role_employee };
  const roleColor = { admin: "green", site_engineer: "blue", employee: "primary" };

  return (
    <div>
      <SectionHeader title="👤 User Management" action="নতুন User" onAction={() => { setEditItem(null); setForm(blankForm); setShowModal(true); }} />
      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead><tr style={{ background: C.primaryBg }}>
            {["নাম", "ইমেইল", "Role", "বিস্তারিত", "স্ট্যাটাস", "Action"].map(h => <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: C.primaryDark, fontWeight: 600, borderBottom: "2px solid " + C.primary }}>{h}</th>)}
          </tr></thead>
          <tbody>
            {users.map(u => (
              <tr key={u.id} style={{ borderBottom: "1px solid " + C.gray100 }} onMouseEnter={e => e.currentTarget.style.background = C.primaryBg} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                <td style={{ padding: "10px 14px", fontWeight: 600, color: C.primaryDark }}>{u.name}</td>
                <td style={{ padding: "10px 14px" }}>{u.email}</td>
                <td style={{ padding: "10px 14px" }}><Badge label={roleLabel[u.role] || u.role} color={roleColor[u.role] || "gray"} /></td>
                <td style={{ padding: "10px 14px", fontSize: 12, color: C.gray600 }}>
                  {u.role === "admin" ? "সব প্রজেক্ট" :
                   u.role === "site_engineer" ? (u.assigned_projects?.length > 0 ? projects.filter(p => u.assigned_projects.includes(p.id)).map(p => p.name).join(", ") : "কোনো প্রজেক্ট নেই") :
                   (u.permissions?.length > 0 ? u.permissions.map(id => TXT[lang || "bn"][id] || HR_SUBTABS.find(h => h[0] === id)?.[2] || id).join(", ") : "কোনো Access নেই")}
                </td>
                <td style={{ padding: "10px 14px" }}><Badge label={u.is_active ? "✅ সক্রিয়" : "❌ নিষ্ক্রিয়"} color={u.is_active ? "green" : "red"} /></td>
                <td style={{ padding: "10px 14px" }}>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button onClick={() => { setEditItem(u); setForm({ ...blankForm, ...u, password_hash: "", employee_id: u.employee_id || "", permissions: u.permissions || [] }); setShowModal(true); }} style={btnEdit}>✏️</button>
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
          <FormField label="পূর্ণ নাম *"><input style={inputStyle} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="নাম" /></FormField>
          <FormField label="ইমেইল *"><input style={inputStyle} type="email" value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} placeholder="user@noksha.com" /></FormField>
          <FormField label={editItem ? "নতুন পাসওয়ার্ড (পরিবর্তন করতে চাইলে)" : "পাসওয়ার্ড *"}><input style={inputStyle} type="text" value={form.password_hash} onChange={e => setForm({ ...form, password_hash: e.target.value })} placeholder="password123" /></FormField>
          <FormField label="Role">
            <select style={inputStyle} value={form.role} onChange={e => setForm({ ...form, role: e.target.value })}>
              <option value="site_engineer">👷 Site Engineer</option>
              <option value="admin">🔑 Admin</option>
              <option value="employee">{T.role_employee}</option>
            </select>
          </FormField>

          {form.role === "site_engineer" && projects.length > 0 && (
            <FormField label="Assigned Projects (যে প্রজেক্টে access পাবে)">
              <div style={{ border: "1px solid " + C.gray200, borderRadius: 8, padding: 12, maxHeight: 200, overflowY: "auto" }}>
                {projects.map(p => (
                  <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, cursor: "pointer" }} onClick={() => toggleProject(p.id)}>
                    <input type="checkbox" checked={(form.assigned_projects || []).includes(p.id)} onChange={() => toggleProject(p.id)} style={{ accentColor: C.primary, width: 16, height: 16 }} />
                    <span style={{ fontSize: 13 }}>{p.name} — {p.client_name}</span>
                  </div>
                ))}
              </div>
            </FormField>
          )}

          {form.role === "employee" && (
            <>
              <FormField label={T.link_employee}>
                <select style={inputStyle} value={form.employee_id} onChange={e => setForm({ ...form, employee_id: e.target.value })}>
                  <option value="">— {T.select_employee} —</option>
                  {(employees || []).map(emp => <option key={emp.id} value={emp.id}>{emp.name} — {emp.dept}</option>)}
                </select>
              </FormField>
              <FormField label={T.permissions_label}>
                <div style={{ border: "1px solid " + C.gray200, borderRadius: 8, padding: 12, maxHeight: 320, overflowY: "auto" }}>
                  {ALL_MENU.filter(m => m.id !== "hr_system").map(m => (
                    <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, cursor: "pointer" }} onClick={() => togglePermission(m.id)}>
                      <input type="checkbox" checked={(form.permissions || []).includes(m.id)} onChange={() => togglePermission(m.id)} style={{ accentColor: C.primary, width: 16, height: 16 }} />
                      <span style={{ fontSize: 13 }}>{m.icon} {TXT[lang || "bn"][m.id] || m.label}</span>
                    </div>
                  ))}
                  <div style={{ borderTop: "1px solid " + C.gray100, marginTop: 10, paddingTop: 10 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 8 }}>👥 HR ও পে-রোল সিস্টেম (sub-menu)</div>
                    {HR_SUBTABS.map(([id, icon, label]) => (
                      <div key={id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, marginLeft: 12, cursor: "pointer" }} onClick={() => togglePermission(id)}>
                        <input type="checkbox" checked={(form.permissions || []).includes(id)} onChange={() => togglePermission(id)} style={{ accentColor: C.primary, width: 16, height: 16 }} />
                        <span style={{ fontSize: 13 }}>{icon} {label}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </FormField>
            </>
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
  const [lang, setLang] = useState(localStorage.getItem("nic_lang") || "bn");
  const [loggedIn, setLoggedIn] = useState(localStorage.getItem("nic_logged_in") === "true");
  const [currentUser, setCurrentUser] = useState(() => {
    try { return JSON.parse(localStorage.getItem("nic_user") || "null"); } catch { return null; }
  });
  const [active, setActive] = useState("dashboard");
  const [sideOpen, setSideOpen] = useState(true);
  const [showProfile, setShowProfile] = useState(false);
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
      supabase.from("employees").select("*").order("sort_order", { ascending: true, nullsFirst: false }).order("created_at", { ascending: false }),
      supabase.from("transactions").select("*").order("created_at", { ascending: false }),
      supabase.from("site_progress").select("*").order("created_at", { ascending: false }),
      supabase.from("materials").select("*").order("created_at", { ascending: false }),
    ]);
    setData({ projects: p.data || [], clients: c.data || [], employees: e.data || [], transactions: t.data || [], siteProgress: s.data || [], materials: m.data || [] });
    setLoading(false);
  };

  const logout = () => { localStorage.removeItem("nic_logged_in"); localStorage.removeItem("nic_user"); setLoggedIn(false); setCurrentUser(null); };

  if (!loggedIn) return <LoginPage onLogin={(user) => { setCurrentUser(user); setLoggedIn(true); if (user.role === "site_engineer") setActive("construction"); if (user.role === "employee") setActive("my_attendance"); }} />;

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.gray50, fontFamily: "'Hind Siliguri', Arial, sans-serif" }}>
      {/* SIDEBAR */}
      <div style={{ width: sideOpen ? 230 : 60, background: C.primaryDark, display: "flex", flexDirection: "column", transition: "width 0.3s", flexShrink: 0, position: "sticky", top: 0, height: "100vh", overflow: "hidden" }}>
        <div style={{ padding: sideOpen ? "18px 18px 14px" : "18px 10px 14px", borderBottom: "1px solid rgba(255,255,255,0.1)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <img src="https://jijxnycopycsysugppnw.supabase.co/storage/v1/object/public/Upload%20images/icon.png" alt="N" style={{ width: 36, height: 36, objectFit: "contain", flexShrink: 0 }} />
            {sideOpen && <div><div style={{ color: C.white, fontWeight: 900, fontSize: 15, letterSpacing: 1 }}>NOKSHA</div><div style={{ color: "rgba(255,255,255,0.5)", fontSize: 9, letterSpacing: 2 }}>INTERIOR & CONSTRUCTION</div></div>}
          </div>
        </div>
        <nav style={{ flex: 1, padding: "8px 0", overflowY: "auto" }}>
          {ALL_MENU.filter(m => currentUser?.role !== "employee" ? m.roles.includes(currentUser?.role || "admin") : m.id === "hr_system" ? hasHRAccess(currentUser, false) : (currentUser.permissions || []).includes(m.id)).map(m => (
            <button key={m.id} onClick={() => setActive(m.id)} style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", padding: sideOpen ? "10px 18px" : "10px 0", justifyContent: sideOpen ? "flex-start" : "center", background: active === m.id ? "rgba(255,255,255,0.12)" : "none", border: "none", borderLeft: active === m.id ? "3px solid " + C.primaryLight : "3px solid transparent", color: active === m.id ? C.white : "rgba(255,255,255,0.6)", cursor: "pointer", fontSize: 12, fontWeight: active === m.id ? 700 : 400, fontFamily: "inherit", transition: "all 0.15s" }}>
              <span style={{ fontSize: 16, flexShrink: 0 }}>{m.icon}</span>
              {sideOpen && <span style={{ whiteSpace: "nowrap", fontSize: 13 }}>{TXT[lang][m.id] || m.label}</span>}
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
        <div style={{ background: C.white, borderBottom: "1px solid " + C.gray200, padding: "12px 24px", display: "flex", justifyContent: "space-between", alignItems: "center", position: "sticky", top: 0, zIndex: 100, borderTop: "3px solid " + C.primary }}>
          <div>
            <div style={{ fontWeight: 700, color: C.primaryDark, fontSize: 16 }}>{ALL_MENU.find(m => m.id === active)?.label}</div>
            <div style={{ fontSize: 11, color: C.gray400 }}>Noksha Interior & Construction · ফরিদপুর & ঢাকা</div>
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            {isAdmin && <button onClick={loadAll} style={{ background: C.primaryBg, border: "1px solid " + C.primaryLight, borderRadius: 8, padding: "6px 12px", cursor: "pointer", fontSize: 12, color: C.primaryDark, fontFamily: "inherit" }}>🔄 রিফ্রেশ</button>}
            <button onClick={() => { const nl = lang === "bn" ? "en" : "bn"; setLang(nl); localStorage.setItem("nic_lang", nl); }} style={{ background: lang === "en" ? C.primaryDark : C.primaryBg, color: lang === "en" ? C.white : C.primaryDark, border: "1px solid " + C.primary, borderRadius: 8, padding: "6px 12px", cursor: "pointer", fontSize: 12, fontWeight: 700, fontFamily: "inherit" }}>{lang === "bn" ? "EN" : "বাং"}</button>
            <div onClick={() => setShowProfile(p => !p)} style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer", padding: "6px 10px", borderRadius: 10, background: showProfile ? C.primaryBg : "transparent", transition: "background 0.2s" }}>
              <div style={{ width: 36, height: 36, borderRadius: "50%", overflow: "hidden", border: "2px solid " + C.primary, background: C.primaryBg, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                {localStorage.getItem("nic_profile_pic") ? <img src={localStorage.getItem("nic_profile_pic")} alt="P" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <span style={{ fontWeight: 700, color: C.primary, fontSize: 14 }}>{currentUser?.name?.[0] || "R"}</span>}
              </div>
              <div style={{ fontSize: 12, color: C.gray600 }}>
                <div style={{ fontWeight: 600, color: C.primaryDark }}>{currentUser?.name || "মোঃ রানা"}</div>
                <div>{currentUser?.role === "admin" ? "Admin" : currentUser?.role === "employee" ? TXT[lang].role_employee : "Site Engineer"}</div>
              </div>
              <span style={{ fontSize: 10, color: C.gray400 }}>▼</span>
            </div>
            {showProfile && <ProfileDropdown currentUser={currentUser} onUpdate={(u) => { setCurrentUser(u); localStorage.setItem("nic_user", JSON.stringify(u)); }} onClose={() => setShowProfile(false)} />}
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
              {active === "dashboard" && canAccessMenu(currentUser, isAdmin, "dashboard") && <Dashboard projects={data.projects} clients={data.clients} employees={data.employees} transactions={data.transactions} materials={data.materials} />}
              {active === "projects" && canAccessMenu(currentUser, isAdmin, "projects") && <Projects data={data.projects} onRefresh={loadAll} />}
              {active === "construction" && <ConstructionProjects currentUser={currentUser} />}
              {active === "interior" && <InteriorProjects currentUser={currentUser} />}
              {active === "boq" && canAccessMenu(currentUser, isAdmin, "boq") && <BOQSystem />}
              {active === "clients" && canAccessMenu(currentUser, isAdmin, "clients") && <Clients data={data.clients} onRefresh={loadAll} />}
              {active === "documents" && canAccessMenu(currentUser, isAdmin, "documents") && <DocumentsHub clients={data.clients} />}
              {active === "hr_system" && hasHRAccess(currentUser, isAdmin) && <HRSystemHub data={data} onRefresh={loadAll} lang={lang} currentUser={currentUser} isAdmin={isAdmin} />}
              {active === "my_attendance" && (isAdmin || currentUser?.role === "employee") && <MyAttendance currentUser={currentUser} lang={lang} />}
              {active === "finance" && canAccessMenu(currentUser, isAdmin, "finance") && <Finance data={data.transactions} onRefresh={loadAll} />}
              {active === "site" && canAccessMenu(currentUser, isAdmin, "site") && <SiteProgress data={data.siteProgress} projects={data.projects} onRefresh={loadAll} />}
              {active === "materials" && canAccessMenu(currentUser, isAdmin, "materials") && <Materials data={data.materials} onRefresh={loadAll} />}
              {active === "analytics" && canAccessMenu(currentUser, isAdmin, "analytics") && <Analytics transactions={data.transactions} projects={data.projects} employees={data.employees} />}
              {active === "users" && canAccessMenu(currentUser, isAdmin, "users") && <UserManagement employees={data.employees} lang={lang} />}
              {active === "password" && <PasswordChange />}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
