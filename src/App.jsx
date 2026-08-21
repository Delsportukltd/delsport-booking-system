import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import JSZip from "jszip";
import {
  LayoutDashboard, CalendarDays, Building2, BarChart3, Plus, X, Menu, Link2,
  MapPin, Phone, Clock, PoundSterling, Check, Ban, Trash2, Pencil,
  ChevronLeft, ChevronRight, ChevronDown, AlertCircle, CircleDot, Users, Mail, Repeat,
  Download, Upload, ShieldCheck, Settings, Lock, Trophy, TrendingUp, TrendingDown, FileDown, Receipt
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid
} from "recharts";
import { supabase } from "./lib/supabaseClient.js";

// ---------- design tokens ----------
const C = {
  navy: "#0B2545",
  navyAlt: "#123262",
  cyan: "#00B4D8",
  cyanSoft: "#E4F7FB",
  pitch: "#2E7D5B",
  pitchSoft: "#E7F3ED",
  amber: "#E0A315",
  amberSoft: "#FBF0DA",
  coral: "#D64550",
  coralSoft: "#FBE7E8",
  bg: "#F5F7FA",
  card: "#FFFFFF",
  ink: "#1E293B",
  mute: "#64748B",
  line: "#E2E8F0",
};

const FACILITY_TYPES = ["Sports Hall", "3G/4G Pitch", "Grass Pitch", "Cricket Pitch", "Rugby Pitch", "MUGA / Outdoor Courts", "Gymnasium", "Dance Studio", "Fitness Suite", "Classroom", "Conference Room", "Lecture Theatre", "Theatre", "Main Hall", "Chapel", "Event Room & Bar", "Swimming Pool", "Tennis / Netball Court", "Other"];
const TIME_SLOTS = Array.from({ length: 35 }, (_, i) => {
  const totalMin = 6 * 60 + i * 30; // 06:00 -> 23:30
  const h = String(Math.floor(totalMin / 60)).padStart(2, "0");
  const m = String(totalMin % 60).padStart(2, "0");
  return `${h}:${m}`;
});
const STATUS_META = {
  confirmed: { label: "Confirmed", fg: C.pitch, bg: C.pitchSoft },
  cancelled: { label: "Cancelled", fg: C.coral, bg: C.coralSoft },
};
// legacy statuses from before the workflow simplified — kept only so old records still render sensibly
const LEGACY_STATUS_META = {
  pending: { label: "Pending", fg: C.amber, bg: C.amberSoft },
  declined: { label: "Declined", fg: C.mute, bg: "#EEF1F4" },
};
function statusMeta(status) { return STATUS_META[status] || LEGACY_STATUS_META[status] || STATUS_META.confirmed; }

const uid = () => crypto.randomUUID();
const money = (n) => `£${(Number(n) || 0).toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const todayISO = () => new Date().toISOString().slice(0, 10);
const dayLabel = (iso) => new Date(iso + "T00:00:00").toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short" });
const monthKey = (iso) => iso.slice(0, 7);
const monthLabel = (mk) => new Date(mk + "-01T00:00:00").toLocaleDateString("en-GB", { month: "short", year: "2-digit" });

// All date math below works in UTC-based epoch arithmetic rather than mixing local-time parsing
// with UTC formatting — that mismatch was the cause of recurring bookings drifting by a day each week.
function isoToUTCms(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return Date.UTC(y, m - 1, d);
}
function utcMsToISO(ms) {
  return new Date(ms).toISOString().slice(0, 10);
}
function addDays(iso, n) {
  return utcMsToISO(isoToUTCms(iso) + n * 86400000);
}
function startOfWeek(iso) {
  const ms = isoToUTCms(iso);
  const day = new Date(ms).getUTCDay();
  const diff = day === 0 ? -6 : 1 - day;
  return utcMsToISO(ms + diff * 86400000);
}
function timeToMin(t) {
  const [h, m] = t.split(":").map(Number);
  return h * 60 + m;
}

function buildMailto({ booking, member, facilityName, siteName, kind }) {
  const to = member?.email || "";
  const isCancel = kind === "cancelled";
  const subject = isCancel
    ? `Booking cancelled — ${facilityName}, ${dayLabel(booking.date)}`
    : `Booking confirmed — ${facilityName}, ${dayLabel(booking.date)}`;
  const lines = [
    `Hi ${member?.name || booking.hirerName || ""},`,
    "",
    isCancel
      ? "This is to confirm the following booking has been cancelled:"
      : "This is to confirm your booking:",
    "",
    `Facility: ${facilityName}${siteName ? ` (${siteName})` : ""}`,
    `Date: ${dayLabel(booking.date)}`,
    `Time: ${booking.startTime}–${booking.endTime}`,
    booking.purpose ? `Purpose: ${booking.purpose}` : "",
    !isCancel ? `Price: ${money(booking.price)}` : "",
    "",
    isCancel ? "Please get in touch if you'd like to rebook." : "If any of these details are incorrect, please let us know.",
    "",
    "Thanks,",
    "Delsport UK",
  ].filter(Boolean).join("\n");
  return `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(lines)}`;
}

function timesOverlap(aStart, aEnd, bStart, bEnd) {
  return timeToMin(aStart) < timeToMin(bEnd) && timeToMin(bStart) < timeToMin(aEnd);
}

// Returns the bookings (excluding excludeId, excluding declined/cancelled) that clash with the given
// facility/date/time window, used to check remaining capacity before saving a booking.
function findClashes(bookings, { facilityId, date, startTime, endTime, excludeId }) {
  return bookings.filter((b) =>
    b.id !== excludeId &&
    b.facilityId === facilityId &&
    b.date === date &&
    b.status !== "declined" && b.status !== "cancelled" &&
    timesOverlap(b.startTime, b.endTime, startTime, endTime)
  );
}

function spacesUsed(clashes) {
  return clashes.reduce((s, b) => s + (Number(b.spaces) || 1), 0);
}

// Bookings sharing a groupId are one multi-facility event under the hood —
// this returns every row in that event (or just the booking itself if it's
// not grouped), used anywhere an action needs to apply to the whole thing
// rather than just the one facility's leg.
function bookingsInGroup(booking, allBookings) {
  if (!booking?.groupId) return [booking].filter(Boolean);
  return allBookings.filter((b) => b.groupId === booking.groupId);
}

// Collapses bookings that share a groupId into a single invoice/confirmation
// row — combined facility names, summed price — instead of listing each
// facility separately (only the first leg actually carries the price, so an
// un-collapsed list would show confusing £0 lines for the others).
function buildDocumentRows(bookings, facilityById, siteName) {
  const seen = new Set();
  const rows = [];
  bookings.forEach((b) => {
    if (seen.has(b.id)) return;
    const linked = b.groupId ? bookings.filter((x) => x.groupId === b.groupId) : [b];
    linked.forEach((x) => seen.add(x.id));
    rows.push({
      dateLabel: dayLabel(b.date),
      facilityLabel: linked.map((x) => facilityById[x.facilityId]?.name || "Unknown").join(" + "),
      siteName: siteName || "",
      timeLabel: `${b.startTime}–${b.endTime}`,
      durationLabel: `${hoursBetween(b.startTime, b.endTime).toFixed(1)} hrs`,
      price: linked.reduce((s, x) => s + (Number(x.price) || 0), 0),
    });
  });
  return rows;
}

function hoursBetween(start, end) {
  return Math.max(0, (timeToMin(end) - timeToMin(start)) / 60);
}

// ---------- booking rule checks (notice period / advance limit) ----------
function noticeViolation(facility, date, startTime) {
  const hrs = Number(facility?.minNoticeHours) || 0;
  if (!hrs) return false;
  const bookingDT = new Date(`${date}T${startTime}:00`);
  const diffHours = (bookingDT.getTime() - Date.now()) / 3600000;
  return diffHours < hrs;
}
function advanceViolation(facility, date) {
  const days = Number(facility?.maxAdvanceDays) || 0;
  if (!days) return false;
  return date > addDays(todayISO(), days);
}

// ---------- iCal export ----------
function icsEscape(s) {
  return String(s || "").replace(/[\\;,]/g, (c) => "\\" + c).replace(/\n/g, "\\n");
}
function toICSDate(date, time) {
  return `${date.replace(/-/g, "")}T${time.replace(":", "")}00`;
}
function bookingToICSEvent(b, facilityName, siteName) {
  const desc = [b.purpose, b.company, `Price: ${money(b.price)}`].filter(Boolean).join(" — ");
  return [
    "BEGIN:VEVENT",
    `UID:${b.id}@delsportuk`,
    `DTSTART:${toICSDate(b.date, b.startTime)}`,
    `DTEND:${toICSDate(b.date, b.endTime)}`,
    `SUMMARY:${icsEscape(`${facilityName} — ${b.hirerName}`)}`,
    `LOCATION:${icsEscape([facilityName, siteName].filter(Boolean).join(", "))}`,
    `DESCRIPTION:${icsEscape(desc)}`,
    "END:VEVENT",
  ].join("\r\n");
}
function buildICS(bookings, facilityById, siteById) {
  const events = bookings.map((b) => bookingToICSEvent(b, facilityById[b.facilityId]?.name, siteById[facilityById[b.facilityId]?.siteId]?.name)).join("\r\n");
  return `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Delsport UK//Bookings//EN\r\n${events}\r\nEND:VCALENDAR`;
}
function downloadICS(filename, content) {
  const blob = new Blob([content], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function downloadJSON(filename, obj) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function csvEscape(v) {
  const s = String(v ?? "");
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
function downloadCSV(filename, headers, rows) {
  const csv = [headers, ...rows].map((r) => r.map(csvEscape).join(",")).join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// ---------- CSV import (members & bookings, e.g. transferring from an old system) ----------
// A small hand-rolled parser rather than a library, since the only quirk we
// need to handle is quoted fields that contain commas or newlines.
function parseCSVText(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  const pushField = () => { row.push(field); field = ""; };
  const pushRow = () => { pushField(); rows.push(row); row = []; };
  const clean = text.replace(/^﻿/, ""); // strip a UTF-8 BOM if Excel added one
  for (let i = 0; i < clean.length; i++) {
    const c = clean[i];
    if (inQuotes) {
      if (c === '"') {
        if (clean[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ",") pushField();
    else if (c === "\r") { /* ignore, \n below ends the row */ }
    else if (c === "\n") pushRow();
    else field += c;
  }
  if (field !== "" || row.length > 0) pushRow();
  return rows.filter((r) => r.length > 1 || (r[0] || "").trim() !== "");
}
function csvRowsToObjects(rows) {
  if (rows.length === 0) return [];
  const headers = rows[0].map((h) => h.trim());
  return rows.slice(1).map((r) => Object.fromEntries(headers.map((h, i) => [h, (r[i] ?? "").trim()])));
}
function normKey(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }
// Looks a value up by several possible header spellings, so a CSV exported
// from another system ("Full Name", "Client", "Hirer") still matches.
function pickField(rowObj, aliases) {
  const entries = Object.entries(rowObj).map(([k, v]) => [normKey(k), v]);
  for (const alias of aliases) {
    const hit = entries.find(([k]) => k === alias);
    if (hit && hit[1]) return hit[1];
  }
  return "";
}
// Accepts "2026-08-05" or UK-style "05/08/2026" / "05-08-2026".
function parseDateFlexible(s) {
  s = String(s || "").trim();
  let y, mo, d;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    [y, mo, d] = s.split("-").map(Number);
  } else {
    const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (!m) return null;
    d = Number(m[1]); mo = Number(m[2]); y = Number(m[3]);
  }
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const iso = `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  // round-trip through Date to reject impossible combos like 31/04 or 30/02
  const check = new Date(iso + "T00:00:00Z");
  if (check.getUTCFullYear() !== y || check.getUTCMonth() + 1 !== mo || check.getUTCDate() !== d) return null;
  return iso;
}
// Accepts "17:00", "17:00:00", "5:00" — always returns "HH:MM" or null.
function parseTimeFlexible(s) {
  s = String(s || "").trim();
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return `${m[1].padStart(2, "0")}:${m[2]}`;
}
function parseStatusFlexible(s) {
  const v = normKey(s);
  if (v === "cancelled" || v === "canceled") return "cancelled";
  return "confirmed";
}
// A facility name alone can be ambiguous across sites (e.g. "Sports Hall"
// exists at four of them), so match on site name too whenever it's given.
function resolveFacilityByName(facilities, sites, siteName, facilityName) {
  const fname = normKey(facilityName);
  let candidates = facilities.filter((f) => normKey(f.name) === fname);
  if (siteName) {
    const site = sites.find((s) => normKey(s.name) === normKey(siteName));
    if (!site) return { error: `no site named "${siteName}"` };
    candidates = candidates.filter((f) => f.siteId === site.id);
  }
  if (candidates.length === 0) return { error: `no facility named "${facilityName}"${siteName ? ` at "${siteName}"` : ""}` };
  if (candidates.length > 1) return { error: `"${facilityName}" matches facilities at more than one site — add a Site column to say which one` };
  return { facility: candidates[0] };
}
// Splits a "Site A, Site B" cell into site ids, e.g. for a member's allowed
// sites on import. Returns an error if any name doesn't match a real site.
function resolveSiteIdsByNames(sites, namesStr) {
  const names = String(namesStr || "").split(/[,;]/).map((s) => s.trim()).filter(Boolean);
  if (names.length === 0) return { siteIds: [] };
  const ids = [];
  for (const name of names) {
    const site = sites.find((s) => normKey(s.name) === normKey(name));
    if (!site) return { error: `no site named "${name}"` };
    ids.push(site.id);
  }
  return { siteIds: ids };
}

// ---------- blackout / closure rules ----------
const DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
function ruleAppliesToFacility(rule, facilityId, siteId) {
  if (rule.scope === "all") return true;
  if (rule.scope === "site") return rule.siteId === siteId; // legacy single-site rules
  if (rule.scope === "sites") return (rule.siteIds || []).includes(siteId);
  if (rule.scope === "facility") return rule.facilityId === facilityId;
  return false;
}
function findBlackout(blackouts, { facilityId, siteId, date, startTime, endTime }) {
  return blackouts.find((r) => {
    if (!ruleAppliesToFacility(r, facilityId, siteId)) return false;
    if (date < r.startDate || date > r.endDate) return false;
    if (r.days && r.days.length > 0) {
      const dow = new Date(date + "T00:00:00").getDay();
      if (!r.days.includes(dow)) return false;
    }
    if (r.allDay) return true;
    return timesOverlap(r.startTime, r.endTime, startTime, endTime);
  }) || null;
}

// ---------- starter data, pulled from delsportuk.com's site & facilities listings ----------
const STARTER_SITES = [
  { name: "Stourport High School", address: "Stourport High School and VIth Form College, Minster Road, Stourport-on-Severn, UK", logoUrl: "https://static.wixstatic.com/media/3cf024_3f7825d256b94204bc3d2ff392f14d8f~mv2.png" },
  { name: "St Augustine's Catholic High", address: "Stonepits Ln, Hunt End, Redditch B97 5LX, UK", logoUrl: "https://static.wixstatic.com/media/3cf024_8c20675a210743d1bae430d3fd5b3b2b~mv2.png" },
  { name: "Baxter College", address: "Baxter College, Habberley Road, Kidderminster, UK", logoUrl: "https://static.wixstatic.com/media/3cf024_e20c1aa2d181424f9fe010de00cdde86~mv2.png" },
  { name: "St Benedict's Catholic High", address: "St Benedict's Catholic High School, Kinwarton Road, Alcester, UK", logoUrl: "https://static.wixstatic.com/media/3cf024_486f4ac4b166486d94b376eba724c49f~mv2.png" },
  { name: "Stourport Cricket Club", address: "Walshes Meadow, Harold Davies Dr, Stourport-on-Severn DY13 0AA, UK", logoUrl: "https://static.wixstatic.com/media/3cf024_41a6ecdce7df48f6bdc7edf09ceb8c20~mv2.png" },
  { name: "Stourport Rugby Club", address: "Walshes Meadow, Harold Davies Dr, Stourport-on-Severn DY13 0AA, UK", logoUrl: "https://static.wixstatic.com/media/3cf024_b287aa83fea94e53b1eba738af7d5347~mv2.png" },
];
const STARTER_FACILITIES = [
  { site: "Stourport High School", name: "3G/4G Pitch", type: "3G/4G Pitch" },
  { site: "Stourport High School", name: "Sports Hall", type: "Sports Hall" },
  { site: "Stourport High School", name: "Main Hall (with stage)", type: "Main Hall" },
  { site: "Stourport High School", name: "Classroom", type: "Classroom" },
  { site: "Stourport High School", name: "Conference Room", type: "Conference Room" },
  { site: "Stourport High School", name: "Lecture Theatre", type: "Lecture Theatre" },
  { site: "St Augustine's Catholic High", name: "Sports Hall", type: "Sports Hall" },
  { site: "St Augustine's Catholic High", name: "Fitness Suite", type: "Fitness Suite" },
  { site: "St Augustine's Catholic High", name: "Gymnasium", type: "Gymnasium" },
  { site: "St Augustine's Catholic High", name: "Theatre", type: "Theatre" },
  { site: "St Augustine's Catholic High", name: "Outdoor Courts", type: "MUGA / Outdoor Courts" },
  { site: "St Augustine's Catholic High", name: "Sports Field", type: "Grass Pitch" },
  { site: "Baxter College", name: "Classroom", type: "Classroom" },
  { site: "Baxter College", name: "Conference Room", type: "Conference Room" },
  { site: "Baxter College", name: "Grass Sports Pitches", type: "Grass Pitch" },
  { site: "Baxter College", name: "3G/4G Pitch", type: "3G/4G Pitch" },
  { site: "Baxter College", name: "Sports Hall", type: "Sports Hall" },
  { site: "Baxter College", name: "Dance Studio", type: "Dance Studio" },
  { site: "Baxter College", name: "Fitness Suite", type: "Fitness Suite" },
  { site: "Baxter College", name: "Gymnasium", type: "Gymnasium" },
  { site: "Baxter College", name: "Theatre", type: "Theatre" },
  { site: "St Benedict's Catholic High", name: "Chapel", type: "Chapel" },
  { site: "St Benedict's Catholic High", name: "Gymnasium", type: "Gymnasium" },
  { site: "St Benedict's Catholic High", name: "Lecture Theatre", type: "Lecture Theatre" },
  { site: "St Benedict's Catholic High", name: "Outdoor Sports Courts", type: "MUGA / Outdoor Courts" },
  { site: "St Benedict's Catholic High", name: "Classroom", type: "Classroom" },
  { site: "St Benedict's Catholic High", name: "Conference Room", type: "Conference Room" },
  { site: "St Benedict's Catholic High", name: "Dance Studio", type: "Dance Studio" },
  { site: "St Benedict's Catholic High", name: "Grass Sports Pitches", type: "Grass Pitch" },
  { site: "St Benedict's Catholic High", name: "Sports Hall", type: "Sports Hall" },
  { site: "Stourport Cricket Club", name: "Cricket Pitches", type: "Cricket Pitch" },
  { site: "Stourport Cricket Club", name: "Event Room & Bar", type: "Event Room & Bar" },
  { site: "Stourport Rugby Club", name: "Rugby Pitches", type: "Rugby Pitch" },
];

function loadStarterData(sites, setSites, facilities, setFacilities) {
  const existingSiteNames = new Set(sites.map((s) => s.name.toLowerCase()));
  const newSites = STARTER_SITES.filter((s) => !existingSiteNames.has(s.name.toLowerCase())).map((s) => ({ ...s, id: uid(), contact: "" }));
  const allSites = [...sites, ...newSites];
  const siteIdByName = Object.fromEntries(allSites.map((s) => [s.name.toLowerCase(), s.id]));

  const existingFacKeys = new Set(facilities.map((f) => `${f.siteId}::${f.name.toLowerCase()}`));
  const newFacilities = STARTER_FACILITIES
    .map((f) => ({ siteId: siteIdByName[f.site.toLowerCase()], name: f.name, type: f.type }))
    .filter((f) => f.siteId && !existingFacKeys.has(`${f.siteId}::${f.name.toLowerCase()}`))
    .map((f) => ({ ...f, id: uid(), rate: "", capacity: 1, minNoticeHours: 0, maxAdvanceDays: 0, customFields: [] }));

  setSites(allSites);
  setFacilities((fs) => [...fs, ...newFacilities]);
  return { addedSites: newSites.length, addedFacilities: newFacilities.length };
}

// ---------- generic UI bits ----------
// Every delete action in the app — single-item or bulk — routes through this
// so nothing gets removed without an explicit confirm click.
function ConfirmDeleteModal({ title = "Delete this?", message, confirmLabel = "Yes, delete", onConfirm, onCancel }) {
  return (
    <Modal title={title} onClose={onCancel}>
      <div style={{ fontSize: 13.5, color: C.ink, marginBottom: 22, lineHeight: 1.5 }}>{message}</div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn variant="ghost" onClick={onCancel}>Cancel</Btn>
        <Btn variant="danger" onClick={onConfirm}>{confirmLabel}</Btn>
      </div>
    </Modal>
  );
}
function Modal({ title, onClose, children, wide }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16 }} onClick={onClose}>
      <div
        style={{ background: C.card, borderRadius: 14, width: "100%", maxWidth: wide ? 640 : 460, maxHeight: "88vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(11,37,69,0.35)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "18px 22px", borderBottom: `1px solid ${C.line}` }}>
          <h3 style={{ margin: 0, fontFamily: "'Space Grotesk', sans-serif", fontSize: 17, color: C.navy, fontWeight: 600 }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: C.mute, padding: 4 }}>
            <X size={20} />
          </button>
        </div>
        <div style={{ padding: 22 }}>{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "block", marginBottom: 14 }}>
      <span style={{ display: "block", fontSize: 12.5, fontWeight: 600, color: C.mute, marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.4 }}>{label}</span>
      {children}
    </label>
  );
}

const inputStyle = {
  width: "100%", boxSizing: "border-box", padding: "9px 11px", borderRadius: 8,
  border: `1px solid ${C.line}`, fontSize: 14.5, color: C.ink, fontFamily: "'Inter', sans-serif", outline: "none",
};

function Btn({ children, onClick, variant = "primary", small, icon: Icon, type = "button", disabled }) {
  const styles = {
    primary: { background: C.navy, color: "#fff", border: "none" },
    accent: { background: C.cyan, color: C.navy, border: "none" },
    ghost: { background: "transparent", color: C.navy, border: `1px solid ${C.line}` },
    danger: { background: C.coralSoft, color: C.coral, border: "none" },
    success: { background: C.pitchSoft, color: C.pitch, border: "none" },
  };
  return (
    <button
      type={type}
      disabled={disabled}
      onClick={onClick}
      style={{
        ...styles[variant], display: "inline-flex", alignItems: "center", gap: 6,
        padding: small ? "6px 11px" : "9px 16px", borderRadius: 8, fontSize: small ? 13 : 14,
        fontWeight: 600, cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.5 : 1,
        fontFamily: "'Inter', sans-serif", whiteSpace: "nowrap",
      }}
    >
      {Icon && <Icon size={small ? 14 : 15} />}
      {children}
    </button>
  );
}

function StatusPill({ status }) {
  const m = statusMeta(status);
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 5, background: m.bg, color: m.fg, fontSize: 12, fontWeight: 700, padding: "3px 9px", borderRadius: 999 }}>
      <CircleDot size={10} /> {m.label}
    </span>
  );
}

function EmptyState({ icon: Icon, title, sub, action }) {
  return (
    <div style={{ textAlign: "center", padding: "56px 20px", color: C.mute }}>
      <div style={{ width: 52, height: 52, borderRadius: 14, background: C.cyanSoft, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
        <Icon size={24} color={C.cyan} />
      </div>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 16, color: C.navy, fontWeight: 600, marginBottom: 4 }}>{title}</div>
      <div style={{ fontSize: 13.5, marginBottom: 16 }}>{sub}</div>
      {action}
    </div>
  );
}

// ---------- Supabase sync helpers ----------
// Converts JS camelCase keys <-> Postgres snake_case columns. Every field
// name in this app (siteId -> site_id, hirerName -> hirer_name, etc.)
// follows plain camelCase-to-snake_case, so one generic pair of converters
// covers sites, facilities, members, bookings and blackouts.
function camelToSnake(obj) {
  const out = {};
  for (const k in obj) out[k.replace(/[A-Z]/g, (m) => "_" + m.toLowerCase())] = obj[k];
  return out;
}
function snakeToCamel(obj) {
  const out = {};
  for (const k in obj) out[k.replace(/_([a-z])/g, (_, c) => c.toUpperCase())] = obj[k];
  return out;
}

// Postgres `time` columns round-trip as "HH:MM:SS" — the rest of the app
// (time-slot dropdowns, "17:00–18:00" labels, .ics export) works in "HH:MM",
// so trim the seconds back off right after loading.
function hhmm(t) {
  return typeof t === "string" ? t.slice(0, 5) : t;
}

// Only these fields are real columns on each table — local booking/blackout
// state also carries transient UI-only fields (e.g. a booking form's
// repeatMode/repeatUntil/multiDates while it's being created) that must be
// dropped before a row is written to Postgres.
const TABLE_COLUMNS = {
  sites: ["id", "name", "address", "contact", "logoUrl", "bankAccountName", "bankSortCode", "bankAccountNumber", "vatNumber"],
  facilities: ["id", "siteId", "name", "type", "rate", "capacity", "minNoticeHours", "maxAdvanceDays", "customFields"],
  members: ["id", "name", "company", "email", "phone", "siteIds"],
  bookings: ["id", "facilityId", "memberId", "date", "startTime", "endTime", "purpose", "price", "status", "notes", "spaces", "recurringId", "groupId", "hirerName", "hirerContact", "company", "customValues"],
  blackouts: ["id", "label", "scope", "siteIds", "facilityId", "startDate", "endDate", "allDay", "days", "startTime", "endTime"],
  invoices: ["id", "invoiceNumber", "memberId", "siteId", "periodStart", "periodEnd", "total", "bookingIds", "pdfPath"],
};

function toRow(table, item) {
  const picked = {};
  // Skip undefined fields entirely (not just falsy) — supabase-js builds its
  // upsert column list from Object.keys(), so a key merely *present* with an
  // undefined value (e.g. a CSV-imported booking that never set customValues)
  // tells PostgREST to write that column as NULL instead of using its default.
  for (const key of TABLE_COLUMNS[table]) if (item[key] !== undefined) picked[key] = item[key];
  if (table === "facilities") {
    picked.rate = Number(picked.rate) || 0;
    picked.capacity = Number(picked.capacity) || 1;
    picked.minNoticeHours = Number(picked.minNoticeHours) || 0;
    picked.maxAdvanceDays = Number(picked.maxAdvanceDays) || 0;
  }
  if (table === "bookings") {
    picked.price = Number(picked.price) || 0;
    picked.spaces = Number(picked.spaces) || 1;
    if (picked.customValues === undefined) picked.customValues = {};
  }
  if (table === "members" && picked.siteIds === undefined) picked.siteIds = [];
  if (table === "invoices") picked.total = Number(picked.total) || 0;
  return camelToSnake(picked);
}

// Diffs `items` against `previous` (both keyed by id) and pushes only what
// changed to Supabase — the equivalent of the old "save the whole blob"
// effect, but for a normalised table instead of a JSON blob.
async function syncTable(table, items, previous) {
  const prevById = Object.fromEntries(previous.map((x) => [x.id, x]));
  const nextIds = new Set(items.map((x) => x.id));
  const changed = items.filter((x) => JSON.stringify(x) !== JSON.stringify(prevById[x.id]));
  const removedIds = previous.map((x) => x.id).filter((id) => !nextIds.has(id));

  if (changed.length > 0) {
    const { error } = await supabase.from(table).upsert(changed.map((x) => toRow(table, x)));
    if (error) throw error;
  }
  if (removedIds.length > 0) {
    const { error } = await supabase.from(table).delete().in("id", removedIds);
    if (error) throw error;
  }
}

// Sits at the very top of the sidebar, above the page nav — picks which
// site(s) the whole app is currently scoped to. Empty selection = all sites.
function SiteScopeSelector({ sites, activeSiteIds, setActiveSiteIds }) {
  const [open, setOpen] = useState(false);
  const label = activeSiteIds.length === 0
    ? "All sites"
    : activeSiteIds.length === 1
      ? sites.find((s) => s.id === activeSiteIds[0])?.name || "1 site"
      : `${activeSiteIds.length} sites`;

  function toggleSite(id) {
    setActiveSiteIds((ids) => (ids.includes(id) ? ids.filter((x) => x !== id) : [...ids, id]));
  }

  return (
    <div style={{ position: "relative", marginBottom: 16 }}>
      <button onClick={() => setOpen((o) => !o)} style={{
        width: "100%", boxSizing: "border-box", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8,
        padding: "9px 12px", borderRadius: 9, background: "rgba(255,255,255,0.07)", border: "none",
        color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600, fontFamily: "'Inter', sans-serif",
      }}>
        <span style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
          <MapPin size={14} color={C.cyan} style={{ flexShrink: 0 }} />
          <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{label}</span>
        </span>
        <ChevronDown size={14} color="#8FA9C9" style={{ flexShrink: 0, transform: open ? "rotate(180deg)" : "none" }} />
      </button>

      {open && (
        <>
          <div style={{ position: "fixed", inset: 0, zIndex: 19 }} onClick={() => setOpen(false)} />
          <div style={{
            position: "absolute", top: "calc(100% + 6px)", left: 0, right: 0, zIndex: 20,
            background: C.card, borderRadius: 10, boxShadow: "0 16px 40px rgba(0,0,0,0.35)",
            padding: 6, maxHeight: 280, overflowY: "auto",
          }}>
            <button
              onClick={() => { setActiveSiteIds([]); setOpen(false); }}
              style={{
                width: "100%", textAlign: "left", padding: "7px 10px", borderRadius: 7, border: "none", cursor: "pointer",
                fontSize: 13, fontWeight: 700, fontFamily: "'Inter', sans-serif",
                background: activeSiteIds.length === 0 ? C.cyanSoft : "transparent", color: C.navy, marginBottom: 4,
              }}
            >
              All sites
            </button>
            {sites.map((s) => (
              <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 10px", borderRadius: 7, cursor: "pointer", fontSize: 13, color: C.ink }}>
                <input type="checkbox" checked={activeSiteIds.includes(s.id)} onChange={() => toggleSite(s.id)} />
                {s.name}
              </label>
            ))}
            {sites.length === 0 && <div style={{ padding: "7px 10px", fontSize: 12.5, color: C.mute }}>No sites yet.</div>}
          </div>
        </>
      )}
    </div>
  );
}

// ---------- main app ----------
function BookingApp({ currentUser, onLogout }) {
  const [sites, setSites] = useState([]);
  const [facilities, setFacilities] = useState([]);
  const [bookings, setBookings] = useState([]);
  const [members, setMembers] = useState([]);
  const [blackouts, setBlackouts] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [saveErr, setSaveErr] = useState(false);
  const [page, setPage] = useState("dashboard");
  const [backupOpen, setBackupOpen] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const prevRef = useRef({ sites: [], facilities: [], members: [], bookings: [], blackouts: [], invoices: [] });

  // Which site(s) the whole app is currently scoped to — empty means "all
  // sites". Kept in localStorage (per browser, not synced) so a site's own
  // staff can leave it parked on their site between visits.
  const [activeSiteIds, setActiveSiteIds] = useState(() => {
    try { return JSON.parse(localStorage.getItem("delsport-active-sites")) || []; } catch { return []; }
  });
  useEffect(() => {
    localStorage.setItem("delsport-active-sites", JSON.stringify(activeSiteIds));
  }, [activeSiteIds]);
  useEffect(() => {
    // drop any site that no longer exists (deleted, or from an old browser session)
    setActiveSiteIds((ids) => ids.filter((id) => sites.some((s) => s.id === id)));
  }, [sites]);

  useEffect(() => {
    (async () => {
      const [sitesRes, facilitiesRes, membersRes, bookingsRes, blackoutsRes, invoicesRes] = await Promise.all([
        supabase.from("sites").select("*"),
        supabase.from("facilities").select("*"),
        supabase.from("members").select("*"),
        supabase.from("bookings").select("*"),
        supabase.from("blackouts").select("*"),
        supabase.from("invoices").select("*"),
      ]);
      const loadedSites = (sitesRes.data || []).map(snakeToCamel);
      const loadedFacilities = (facilitiesRes.data || []).map(snakeToCamel);
      const loadedMembers = (membersRes.data || []).map(snakeToCamel);
      const loadedBookings = (bookingsRes.data || []).map(snakeToCamel).map((b) => ({ ...b, startTime: hhmm(b.startTime), endTime: hhmm(b.endTime) }));
      const loadedBlackouts = (blackoutsRes.data || []).map(snakeToCamel).map((r) => ({ ...r, startTime: hhmm(r.startTime), endTime: hhmm(r.endTime) }));
      const loadedInvoices = (invoicesRes.data || []).map(snakeToCamel);
      setSites(loadedSites);
      setFacilities(loadedFacilities);
      setMembers(loadedMembers);
      setBookings(loadedBookings);
      setBlackouts(loadedBlackouts);
      setInvoices(loadedInvoices);
      // seed the diff baseline with what we just loaded, so the sync effect
      // below doesn't immediately try to re-upsert every row it just read
      prevRef.current = { sites: loadedSites, facilities: loadedFacilities, members: loadedMembers, bookings: loadedBookings, blackouts: loadedBlackouts, invoices: loadedInvoices };
      setLoaded(true);
    })();
  }, []);

  // Pushes only what changed to Supabase, in FK-safe order (sites before
  // facilities/blackouts, facilities+members before bookings) — the
  // equivalent of the old "save the whole blob on every change" effects,
  // but as targeted upserts/deletes against real tables.
  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try {
        await syncTable("sites", sites, prevRef.current.sites);
        await syncTable("members", members, prevRef.current.members);
        await syncTable("facilities", facilities, prevRef.current.facilities);
        await syncTable("bookings", bookings, prevRef.current.bookings);
        await syncTable("blackouts", blackouts, prevRef.current.blackouts);
        await syncTable("invoices", invoices, prevRef.current.invoices);
        prevRef.current = { sites, facilities, members, bookings, blackouts, invoices };
        setSaveErr(false);
      } catch (e) {
        setSaveErr(true);
      }
    })();
  }, [sites, facilities, members, bookings, blackouts, invoices, loaded]);

  const facilityById = useMemo(() => Object.fromEntries(facilities.map((f) => [f.id, f])), [facilities]);
  const siteById = useMemo(() => Object.fromEntries(sites.map((s) => [s.id, s])), [sites]);
  const memberById = useMemo(() => Object.fromEntries(members.map((m) => [m.id, m])), [members]);

  // Everything below narrows to the active site scope — empty activeSiteIds
  // means unscoped (show everything), matching the same "no restriction"
  // convention used for a member's own siteIds.
  const scoped = activeSiteIds.length > 0;
  const visibleSites = scoped ? sites.filter((s) => activeSiteIds.includes(s.id)) : sites;
  const visibleFacilities = scoped ? facilities.filter((f) => activeSiteIds.includes(f.siteId)) : facilities;
  const visibleFacilityIds = useMemo(() => new Set(visibleFacilities.map((f) => f.id)), [visibleFacilities]);
  const visibleBookings = scoped ? bookings.filter((b) => visibleFacilityIds.has(b.facilityId)) : bookings;
  const visibleMembers = scoped ? members.filter((m) => !m.siteIds?.length || m.siteIds.some((id) => activeSiteIds.includes(id))) : members;
  const visibleBlackouts = scoped ? blackouts.filter((r) => {
    if (r.scope === "all") return true;
    if (r.scope === "sites") return (r.siteIds || []).some((id) => activeSiteIds.includes(id));
    if (r.scope === "facility") return visibleFacilityIds.has(r.facilityId);
    return true;
  }) : blackouts;
  const visibleInvoices = scoped ? invoices.filter((inv) => activeSiteIds.includes(inv.siteId)) : invoices;

  const nav = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "bookings", label: "Bookings", icon: CalendarDays },
    { id: "members", label: "Members", icon: Users },
    { id: "facilities", label: "Sites & Facilities", icon: Building2 },
    { id: "reports", label: "Reports", icon: BarChart3 },
    { id: "invoices", label: "Invoices", icon: Receipt },
    { id: "admin", label: "Admin", icon: Settings },
  ];

  const bottomNavItems = [
    { id: "dashboard", label: "Home", icon: LayoutDashboard },
    { id: "bookings", label: "Bookings", icon: CalendarDays },
    { id: "members", label: "Members", icon: Users },
    { id: "invoices", label: "Invoices", icon: Receipt },
  ];

  const sidebarInner = (
    <>
      <div style={{ padding: "0 8px 22px", display: "flex", alignItems: "center", gap: 10 }}>
        <img src="https://static.wixstatic.com/media/3cf024_f16abd3550ca4fce86e2d135f9d0f27c~mv2.png" alt="Delsport UK logo" style={{ width: 38, height: 38, objectFit: "contain", flexShrink: 0 }} />
        <div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 17, letterSpacing: 0.2, lineHeight: 1.15 }}>DELSPORT UK</div>
          <div style={{ fontSize: 10.5, color: "#8FA9C9", letterSpacing: 0.4, fontStyle: "italic" }}>Get more from sport</div>
        </div>
      </div>
      <SiteScopeSelector sites={sites} activeSiteIds={activeSiteIds} setActiveSiteIds={setActiveSiteIds} />
      <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
        {nav.map((n) => {
          const active = page === n.id;
          return (
            <button key={n.id} onClick={() => { setPage(n.id); setMobileMenuOpen(false); }}
              style={{
                display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 9,
                background: active ? "rgba(0,180,216,0.16)" : "transparent",
                border: "none", color: active ? C.cyan : "#C7D5E8", cursor: "pointer",
                fontSize: 14, fontWeight: active ? 700 : 500, textAlign: "left", fontFamily: "'Inter', sans-serif",
              }}>
              <n.icon size={17} />
              {n.label}
            </button>
          );
        })}
      </div>
      <div style={{ marginTop: "auto", paddingTop: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 12px", borderRadius: 9, background: "rgba(255,255,255,0.04)", marginBottom: 8 }}>
          <div style={{ width: 26, height: 26, borderRadius: 999, background: C.cyan, color: C.navy, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12, flexShrink: 0 }}>
            {(currentUser?.displayName || "?").slice(0, 1).toUpperCase()}
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 12.5, fontWeight: 600, color: "#fff", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{currentUser?.displayName}</div>
            <button onClick={onLogout} style={{ background: "none", border: "none", color: "#8FA9C9", fontSize: 11, cursor: "pointer", padding: 0 }}>Log out</button>
          </div>
        </div>
        <button onClick={() => { setBackupOpen(true); setMobileMenuOpen(false); }} style={{
          display: "flex", alignItems: "center", gap: 9, width: "100%", padding: "9px 12px", borderRadius: 9,
          background: "rgba(255,255,255,0.06)", border: "none", color: "#C7D5E8", cursor: "pointer", fontSize: 13, fontWeight: 600, marginBottom: 12,
        }}>
          <ShieldCheck size={16} /> Backup & restore
        </button>
        <div style={{ fontSize: 11, color: "#5F7EA3", lineHeight: 1.5 }}>
          {saveErr && <div style={{ color: "#F4B7BC", marginBottom: 8 }}>⚠ Couldn't save last change — check connection.</div>}
          Data is shared across your team.
        </div>
      </div>
    </>
  );

  return (
    <div className="dp-app-shell" style={{ display: "flex", minHeight: 640, background: C.bg, fontFamily: "'Inter', sans-serif", borderRadius: 16, overflow: "hidden", border: `1px solid ${C.line}` }}>
      <link rel="preconnect" href="https://fonts.googleapis.com" />
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        .dp-mobile-topbar, .dp-mobile-bottomnav, .dp-drawer-backdrop { display: none; }
        @media (max-width: 767px) {
          .dp-app-shell { border-radius: 0 !important; border: none !important; min-height: 100vh !important; overflow: visible !important; }
          .dp-sidebar { display: none !important; }
          .dp-main { padding: 68px 14px calc(76px + env(safe-area-inset-bottom, 0px)) !important; }
          .dp-mobile-topbar { display: flex; }
          .dp-mobile-bottomnav { display: flex; }
          .dp-drawer-backdrop.open { display: block; }
          .dp-stat-grid { grid-template-columns: repeat(2, 1fr) !important; }
          .dp-dash-grid, .dp-form-grid, .dp-invoice-filters { grid-template-columns: 1fr !important; }
          *:has(> table) { overflow-x: auto !important; -webkit-overflow-scrolling: touch; }
          table td, table th { white-space: nowrap; }
          h1 { font-size: 21px !important; }
        }
      `}</style>

      {/* mobile top bar */}
      <div className="dp-mobile-topbar" style={{ position: "fixed", top: 0, left: 0, right: 0, height: 56, background: C.navy, alignItems: "center", justifyContent: "space-between", padding: "0 8px 0 4px", zIndex: 40 }}>
        <button onClick={() => setMobileMenuOpen(true)} style={{ background: "none", border: "none", color: "#fff", padding: 10, display: "flex" }}>
          <Menu size={22} />
        </button>
        <div style={{ color: "#fff", fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif", fontSize: 15 }}>
          {nav.find((n) => n.id === page)?.label || "Delsport UK"}
        </div>
        <div style={{ width: 42 }} />
      </div>

      {/* mobile slide-in menu */}
      <div className={`dp-drawer-backdrop${mobileMenuOpen ? " open" : ""}`} style={{ position: "fixed", inset: 0, background: "rgba(11,37,69,0.5)", zIndex: 60 }} onClick={() => setMobileMenuOpen(false)}>
        <div style={{ width: 250, maxWidth: "80vw", height: "100%", background: C.navy, color: "#fff", padding: "14px 14px 22px", display: "flex", flexDirection: "column", overflowY: "auto" }} onClick={(e) => e.stopPropagation()}>
          <button onClick={() => setMobileMenuOpen(false)} style={{ alignSelf: "flex-end", background: "none", border: "none", color: "#C7D5E8", padding: 6, marginBottom: 6 }}>
            <X size={20} />
          </button>
          {sidebarInner}
        </div>
      </div>

      {/* sidebar (desktop only) */}
      <div className="dp-sidebar" style={{ width: 208, background: C.navy, color: "#fff", padding: "22px 14px", flexShrink: 0, display: "flex", flexDirection: "column" }}>
        {sidebarInner}
      </div>

      {/* mobile bottom tab bar */}
      <div className="dp-mobile-bottomnav" style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: C.navy, alignItems: "stretch", justifyContent: "space-around", padding: "4px 2px calc(4px + env(safe-area-inset-bottom, 0px))", zIndex: 40, boxShadow: "0 -2px 10px rgba(11,37,69,0.25)" }}>
        {bottomNavItems.map((n) => {
          const active = page === n.id;
          return (
            <button key={n.id} onClick={() => setPage(n.id)} style={{
              display: "flex", flexDirection: "column", alignItems: "center", gap: 2, flex: 1, padding: "6px 2px",
              background: "none", border: "none", color: active ? C.cyan : "#8FA9C9", cursor: "pointer",
            }}>
              <n.icon size={20} />
              <span style={{ fontSize: 10.5, fontWeight: active ? 700 : 500 }}>{n.label}</span>
            </button>
          );
        })}
        <button onClick={() => setMobileMenuOpen(true)} style={{
          display: "flex", flexDirection: "column", alignItems: "center", gap: 2, flex: 1, padding: "6px 2px",
          background: "none", border: "none", color: "#8FA9C9", cursor: "pointer",
        }}>
          <Menu size={20} />
          <span style={{ fontSize: 10.5, fontWeight: 500 }}>Menu</span>
        </button>
      </div>

      {/* main */}
      <div className="dp-main" style={{ flex: 1, padding: "26px 30px", overflowY: "auto" }}>
        {page === "dashboard" && <Dashboard sites={visibleSites} facilities={visibleFacilities} bookings={visibleBookings} siteById={siteById} facilityById={facilityById} goBookings={() => setPage("bookings")} scoped={scoped} />}
        {page === "bookings" && (
          <Bookings
            sites={visibleSites} facilities={visibleFacilities} bookings={visibleBookings} setBookings={setBookings}
            facilityById={facilityById} siteById={siteById} goFacilities={() => setPage("facilities")}
            members={visibleMembers} setMembers={setMembers} memberById={memberById} goMembers={() => setPage("members")}
            blackouts={visibleBlackouts}
          />
        )}
        {page === "members" && (
          <Members
            members={visibleMembers} setMembers={setMembers} bookings={visibleBookings} setBookings={setBookings}
            facilityById={facilityById} siteById={siteById} allSites={sites}
          />
        )}
        {page === "facilities" && <SitesFacilities sites={visibleSites} setSites={setSites} facilities={visibleFacilities} setFacilities={setFacilities} bookings={visibleBookings} />}
        {page === "reports" && <Reports bookings={visibleBookings} sites={visibleSites} facilities={visibleFacilities} facilityById={facilityById} siteById={siteById} members={visibleMembers} memberById={memberById} />}
        {page === "invoices" && (
          <Invoices
            sites={visibleSites} facilities={facilities} bookings={bookings} members={members}
            facilityById={facilityById} siteById={siteById} memberById={memberById}
            invoices={visibleInvoices} allInvoices={invoices} setInvoices={setInvoices}
          />
        )}
        {page === "admin" && <Admin blackouts={visibleBlackouts} setBlackouts={setBlackouts} sites={visibleSites} facilities={visibleFacilities} siteById={siteById} facilityById={facilityById} currentUser={currentUser} />}
      </div>

      {backupOpen && (
        <BackupModal
          data={{ sites, facilities, bookings, members, blackouts, invoices }}
          onClose={() => setBackupOpen(false)}
          onRestore={(data) => {
            setSites(data.sites || []);
            setFacilities(data.facilities || []);
            setBookings(data.bookings || []);
            setMembers(data.members || []);
            setBlackouts(data.blackouts || []);
            setInvoices(data.invoices || []);
            setBackupOpen(false);
          }}
        />
      )}
    </div>
  );
}

// ---------- login gate ----------
// Supabase checks the email/password server-side (auth.signInWithPassword);
// this screen just collects them and shows whatever error Supabase returns.
function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [checking, setChecking] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    setChecking(true);
    setError("");
    const { error: signInError } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
    setChecking(false);
    if (signInError) setError(signInError.message);
  }

  return (
    <div style={{ minHeight: 640, display: "flex", alignItems: "center", justifyContent: "center", background: C.navy, borderRadius: 16, fontFamily: "'Inter', sans-serif" }}>
      <link rel="preconnect" href="https://fonts.googleapis.com" />
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@600;700&family=Inter:wght@400;500;600;700&display=swap');`}</style>
      <form onSubmit={handleSubmit} style={{ background: C.card, borderRadius: 16, padding: "36px 34px", width: 340, boxShadow: "0 20px 60px rgba(0,0,0,0.35)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 24 }}>
          <img src="https://static.wixstatic.com/media/3cf024_f16abd3550ca4fce86e2d135f9d0f27c~mv2.png" alt="Delsport UK logo" style={{ width: 36, height: 36, objectFit: "contain" }} />
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16, color: C.navy }}>DELSPORT UK</div>
            <div style={{ fontSize: 10.5, color: C.mute, fontStyle: "italic" }}>Get more from sport</div>
          </div>
        </div>
        <Field label="Email">
          <input type="email" style={inputStyle} value={email} onChange={(e) => setEmail(e.target.value)} autoFocus autoCapitalize="none" />
        </Field>
        <Field label="Password">
          <input type="password" style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} />
        </Field>
        {error && <div style={{ fontSize: 12.5, color: C.coral, marginBottom: 12 }}>{error}</div>}
        <Btn type="submit" variant="accent" disabled={!email || !password || checking}>
          {checking ? "Checking…" : "Log in"}
        </Btn>
      </form>
    </div>
  );
}

export default function Root() {
  const [session, setSession] = useState(undefined); // undefined = still loading
  const [currentUser, setCurrentUser] = useState(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => setSession(newSession));
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) {
      setCurrentUser(null);
      return;
    }
    (async () => {
      const { data } = await supabase.from("profiles").select("display_name").eq("id", session.user.id).single();
      setCurrentUser({ id: session.user.id, displayName: data?.display_name || session.user.email });
    })();
  }, [session]);

  // auto-logout after 20 minutes of no mouse/keyboard activity, so an unattended open tab doesn't stay signed in
  useEffect(() => {
    if (!session) return;
    const TIMEOUT_MS = 20 * 60 * 1000;
    let timer = setTimeout(() => supabase.auth.signOut(), TIMEOUT_MS);
    const reset = () => { clearTimeout(timer); timer = setTimeout(() => supabase.auth.signOut(), TIMEOUT_MS); };
    window.addEventListener("mousemove", reset);
    window.addEventListener("keydown", reset);
    window.addEventListener("click", reset);
    return () => {
      clearTimeout(timer);
      window.removeEventListener("mousemove", reset);
      window.removeEventListener("keydown", reset);
      window.removeEventListener("click", reset);
    };
  }, [session]);

  if (session === undefined) {
    return <div style={{ minHeight: 640, display: "flex", alignItems: "center", justifyContent: "center", background: C.bg, borderRadius: 16, color: C.mute, fontFamily: "'Inter', sans-serif" }}>Loading…</div>;
  }

  if (!session || !currentUser) {
    return <LoginScreen />;
  }

  return (
    <BookingApp
      currentUser={currentUser}
      onLogout={() => supabase.auth.signOut()}
    />
  );
}

function BackupModal({ data, onClose, onRestore }) {
  const [confirmFile, setConfirmFile] = useState(null);
  const [error, setError] = useState("");
  const counts = { sites: data.sites.length, facilities: data.facilities.length, bookings: data.bookings.length, members: data.members.length, invoices: data.invoices?.length || 0 };

  function handleDownload() {
    downloadJSON(`delsport-backup-${todayISO()}.json`, { ...data, exportedAt: new Date().toISOString() });
  }

  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!parsed || typeof parsed !== "object") throw new Error("bad file");
        setConfirmFile(parsed);
      } catch (err) {
        setError("That doesn't look like a valid backup file.");
      }
    };
    reader.readAsText(file);
  }

  return (
    <Modal title="Backup & restore" onClose={onClose} wide>
      <div style={{ marginBottom: 22 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, fontSize: 14.5, marginBottom: 6 }}>Download a backup</div>
        <div style={{ fontSize: 13, color: C.mute, marginBottom: 12 }}>
          Saves a single file with everything currently in the system — {counts.sites} site{counts.sites === 1 ? "" : "s"}, {counts.facilities} facilit{counts.facilities === 1 ? "y" : "ies"}, {counts.members} member{counts.members === 1 ? "" : "s"}, {counts.bookings} booking{counts.bookings === 1 ? "" : "s"}, and {counts.invoices} invoice{counts.invoices === 1 ? "" : "s"}. Worth doing weekly, or before making big changes.
          <div style={{ marginTop: 6, fontSize: 12, color: C.mute }}>Note: this saves invoice records (numbers, totals, who was billed) but not the raised PDF files themselves — those stay in cloud storage.</div>
        </div>
        <Btn variant="accent" icon={Download} onClick={handleDownload}>Download backup (.json)</Btn>
      </div>

      <div style={{ borderTop: `1px solid ${C.line}`, paddingTop: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, fontSize: 14.5, marginBottom: 6 }}>Restore from a backup</div>
        <div style={{ fontSize: 13, color: C.mute, marginBottom: 12 }}>
          This replaces everything currently in the system with what's in the file. Use this to recover from a mistake or move data from another export — not something to do casually.
        </div>
        <input type="file" accept="application/json" onChange={handleFile} style={{ fontSize: 13 }} />
        {error && <div style={{ fontSize: 12.5, color: C.coral, marginTop: 8 }}>{error}</div>}

        {confirmFile && (
          <div style={{ marginTop: 14, border: `1px solid ${C.coral}`, background: C.coralSoft, borderRadius: 10, padding: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 7, fontWeight: 700, fontSize: 13.5, color: C.coral, marginBottom: 6 }}>
              <AlertCircle size={15} /> This will replace all current data
            </div>
            <div style={{ fontSize: 12.5, color: C.ink, marginBottom: 10 }}>
              File contains {confirmFile.sites?.length || 0} sites, {confirmFile.facilities?.length || 0} facilities, {confirmFile.members?.length || 0} members, {confirmFile.bookings?.length || 0} bookings, and {confirmFile.invoices?.length || 0} invoices
              {confirmFile.exportedAt ? ` — exported ${new Date(confirmFile.exportedAt).toLocaleString("en-GB")}` : ""}.
              Everything currently in the system will be overwritten and can't be undone.
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn small variant="ghost" onClick={() => setConfirmFile(null)}>Cancel</Btn>
              <Btn small variant="danger" onClick={() => onRestore(confirmFile)}>Yes, replace all data</Btn>
            </div>
          </div>
        )}
      </div>
    </Modal>
  );
}

// ---------- admin: blackout / closure rules ----------
function summariseRule(rule, siteById, facilityById) {
  const scopeText = rule.scope === "all" ? "All facilities"
    : rule.scope === "site" ? `${siteById[rule.siteId]?.name || "Unknown site"} (all facilities)`
    : rule.scope === "sites" ? (
        (rule.siteIds || []).length === 0 ? "No sites selected"
        : (rule.siteIds || []).length <= 2
          ? `${rule.siteIds.map((id) => siteById[id]?.name || "Unknown").join(", ")} (all facilities)`
          : `${rule.siteIds.length} sites (all facilities)`
      )
    : facilityById[rule.facilityId]?.name || "Unknown facility";
  const dateText = rule.startDate === rule.endDate ? dayLabel(rule.startDate) : `${dayLabel(rule.startDate)} – ${dayLabel(rule.endDate)}`;
  const timeText = rule.allDay
    ? "Whole day(s)"
    : `${(rule.days || []).map((d) => DAY_NAMES[d]).join(", ") || "No days selected"}, ${rule.startTime}–${rule.endTime}`;
  return { scopeText, dateText, timeText };
}

function Admin({ blackouts, setBlackouts, sites, facilities, siteById, facilityById, currentUser }) {
  const [tab, setTab] = useState("blackouts");
  const [modalOpen, setModalOpen] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDeleteRule, setConfirmDeleteRule] = useState(null);

  function save(rule) {
    if (rule.id) setBlackouts((bs) => bs.map((r) => (r.id === rule.id ? rule : r)));
    else setBlackouts((bs) => [...bs, { ...rule, id: uid() }]);
    setModalOpen(false);
    setEditing(null);
  }
  function remove(id) {
    setBlackouts((bs) => bs.filter((r) => r.id !== id));
  }

  const sorted = [...blackouts].sort((a, b) => a.startDate.localeCompare(b.startDate));

  return (
    <div>
      <PageHeader
        title="Admin"
        sub="Blackout rules and login accounts for this system."
        action={tab === "blackouts" ? <Btn variant="accent" icon={Plus} onClick={() => { setEditing(null); setModalOpen(true); }}>New blackout rule</Btn> : undefined}
      />

      <div style={{ display: "flex", background: "#EAEEF3", borderRadius: 9, padding: 3, width: "fit-content", marginBottom: 20 }}>
        {[["blackouts", "Blackout rules"], ["accounts", "Login accounts"]].map(([v, label]) => (
          <button key={v} onClick={() => setTab(v)} style={{
            padding: "6px 14px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
            background: tab === v ? C.card : "transparent", color: tab === v ? C.navy : C.mute,
          }}>{label}</button>
        ))}
      </div>

      {tab === "blackouts" && (
        sites.length === 0 ? (
          <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
            <EmptyState icon={Building2} title="Add a site first" sub="Blackout rules apply to sites and facilities, so add one before creating a rule." />
          </div>
        ) : sorted.length === 0 ? (
          <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
            <EmptyState icon={Lock} title="No blackout rules yet" sub="Use these to block whole holiday periods in one go, or recurring hours like the school day, without touching individual bookings." action={<Btn variant="accent" onClick={() => setModalOpen(true)}>New blackout rule</Btn>} />
          </div>
        ) : (
          <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                  {["Rule", "Applies to", "Dates", "Pattern", ""].map((h) => (
                    <th key={h} style={{ padding: "10px 16px", fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.5, color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sorted.map((rule) => {
                  const { scopeText, dateText, timeText } = summariseRule(rule, siteById, facilityById);
                  return (
                    <tr key={rule.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                      <td style={{ padding: "11px 16px", fontWeight: 600, fontSize: 13.5, color: C.ink, display: "flex", alignItems: "center", gap: 7 }}><Lock size={13} color={C.mute} /> {rule.label}</td>
                      <td style={{ padding: "11px 16px", fontSize: 13 }}>{scopeText}</td>
                      <td style={{ padding: "11px 16px", fontSize: 13 }}>{dateText}</td>
                      <td style={{ padding: "11px 16px", fontSize: 12.5, color: C.mute }}>{timeText}</td>
                      <td style={{ padding: "11px 16px" }}>
                        <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                          <Btn small variant="ghost" icon={Pencil} onClick={() => { setEditing(rule); setModalOpen(true); }} />
                          <Btn small variant="ghost" icon={Trash2} onClick={() => setConfirmDeleteRule(rule)} />
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )
      )}

      {tab === "accounts" && <TeamPanel currentUser={currentUser} />}

      {modalOpen && (
        <BlackoutModal initial={editing} sites={sites} facilities={facilities} onClose={() => { setModalOpen(false); setEditing(null); }} onSave={save} />
      )}
      {confirmDeleteRule && (
        <ConfirmDeleteModal
          title="Delete this blackout rule?"
          message={`"${confirmDeleteRule.label}" will be removed and will no longer block bookings. This can't be undone.`}
          onCancel={() => setConfirmDeleteRule(null)}
          onConfirm={() => { remove(confirmDeleteRule.id); setConfirmDeleteRule(null); }}
        />
      )}
    </div>
  );
}

// Login accounts now live in Supabase Auth (server-side), so this panel is
// read-only — it just shows who has a profile. Adding, removing, or
// resetting a password for an account is done from the Supabase dashboard
// (Authentication -> Users, plus a matching row in the profiles table).
function TeamPanel({ currentUser }) {
  const [profiles, setProfiles] = useState(null);

  useEffect(() => {
    supabase.from("profiles").select("id, display_name").then(({ data }) => setProfiles(data || []));
  }, []);

  return (
    <div>
      <div style={{ background: C.amberSoft, border: `1px solid ${C.amber}`, borderRadius: 10, padding: "12px 14px", marginBottom: 16, fontSize: 12.5, color: C.ink, lineHeight: 1.6 }}>
        <div style={{ fontWeight: 700, color: C.amber, marginBottom: 4, display: "flex", alignItems: "center", gap: 6 }}><Lock size={13} /> Managed in Supabase</div>
        Accounts are created, removed, or reset from your Supabase project's dashboard (Authentication → Users), not from here — that's what makes the login check happen on Supabase's servers instead of in the browser. You're automatically logged out after 20 minutes of inactivity.
      </div>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
              {["Name", ""].map((h) => (
                <th key={h} style={{ padding: "10px 16px", fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.5, color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {(profiles || []).map((p) => (
              <tr key={p.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                <td style={{ padding: "11px 16px", fontWeight: 600, fontSize: 13.5, color: C.ink }}>{p.display_name}</td>
                <td style={{ padding: "11px 16px" }}>{p.id === currentUser?.id && <span style={{ fontSize: 11, fontWeight: 700, color: C.pitch, background: C.pitchSoft, padding: "2px 8px", borderRadius: 999 }}>You</span>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

const WEEKDAY_ORDER = [1, 2, 3, 4, 5, 6, 0]; // Mon..Sun for display

function BlackoutModal({ initial, sites, facilities, onClose, onSave }) {
  const [form, setForm] = useState(() => {
    if (initial) {
      return initial.scope === "site" ? { ...initial, scope: "sites", siteIds: [initial.siteId] } : { ...initial, siteIds: initial.siteIds || [] };
    }
    return {
      label: "", scope: "all", siteIds: [], facilityId: facilities[0]?.id || "",
      startDate: todayISO(), endDate: todayISO(), allDay: true, days: [1, 2, 3, 4, 5], startTime: "07:00", endTime: "17:00",
    };
  });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  function useHolidayTemplate() {
    setForm((f) => ({ ...f, label: f.label || "Holiday period", scope: f.scope === "facility" ? f.scope : "all", allDay: true }));
  }
  function useTermTimeTemplate() {
    setForm((f) => ({
      ...f, label: f.label || "Term-time school hours",
      scope: f.scope === "facility" ? f.scope : "sites",
      siteIds: f.siteIds && f.siteIds.length ? f.siteIds : sites.map((s) => s.id),
      allDay: false, days: [1, 2, 3, 4, 5], startTime: "07:00", endTime: "17:00",
    }));
  }
  function toggleDay(d) {
    setForm((f) => ({ ...f, days: f.days.includes(d) ? f.days.filter((x) => x !== d) : [...f.days, d] }));
  }
  function toggleSite(id) {
    setForm((f) => ({ ...f, siteIds: f.siteIds.includes(id) ? f.siteIds.filter((x) => x !== id) : [...f.siteIds, id] }));
  }

  const valid = form.label.trim() && form.startDate && form.endDate && form.endDate >= form.startDate
    && (form.scope !== "sites" || form.siteIds.length > 0) && (form.scope !== "facility" || form.facilityId)
    && (form.allDay || (form.days.length > 0 && timeToMin(form.endTime) > timeToMin(form.startTime)));

  return (
    <Modal title={initial ? "Edit blackout rule" : "New blackout rule"} onClose={onClose} wide>
      {!initial && (
        <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
          <Btn small variant="ghost" onClick={useHolidayTemplate}>Use holiday-period template (all sites)</Btn>
          <Btn small variant="ghost" onClick={useTermTimeTemplate}>Use term-time hours template (pick sites)</Btn>
        </div>
      )}

      <Field label="Reason / label">
        <input style={inputStyle} value={form.label} onChange={set("label")} placeholder="e.g. Summer Holidays 2026, Term-time core hours" />
      </Field>

      <Field label="Applies to">
        <select style={inputStyle} value={form.scope} onChange={(e) => setForm((f) => ({ ...f, scope: e.target.value }))}>
          <option value="all">All sites & facilities</option>
          <option value="sites">Specific sites (all their facilities)</option>
          <option value="facility">A single facility</option>
        </select>
      </Field>
      {form.scope === "sites" && (
        <Field label="Sites">
          <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
            <button type="button" onClick={() => setForm((f) => ({ ...f, siteIds: sites.map((s) => s.id) }))} style={{ background: "none", border: "none", color: C.cyan, fontWeight: 700, fontSize: 12, cursor: "pointer", padding: 0 }}>Select all</button>
            <button type="button" onClick={() => setForm((f) => ({ ...f, siteIds: [] }))} style={{ background: "none", border: "none", color: C.mute, fontWeight: 700, fontSize: 12, cursor: "pointer", padding: 0 }}>Clear</button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 160, overflowY: "auto", border: `1px solid ${C.line}`, borderRadius: 8, padding: 10 }}>
            {sites.map((s) => (
              <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, cursor: "pointer" }}>
                <input type="checkbox" checked={form.siteIds.includes(s.id)} onChange={() => toggleSite(s.id)} />
                {s.logoUrl && <img src={s.logoUrl} alt="" style={{ width: 16, height: 16, objectFit: "contain" }} onError={(e) => { e.target.style.display = "none"; }} />}
                {s.name}
              </label>
            ))}
          </div>
        </Field>
      )}
      {form.scope === "facility" && (
        <Field label="Facility">
          <select style={inputStyle} value={form.facilityId} onChange={set("facilityId")}>
            {facilities.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
          </select>
        </Field>
      )}

      <div className="dp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
        <Field label="Start date">
          <input type="date" style={inputStyle} value={form.startDate} onChange={set("startDate")} />
        </Field>
        <Field label="End date">
          <input type="date" style={inputStyle} min={form.startDate} value={form.endDate} onChange={set("endDate")} />
        </Field>
      </div>

      <div style={{ background: "#F8FAFC", border: `1px solid ${C.line}`, borderRadius: 10, padding: 12, marginBottom: 14 }}>
        <div style={{ display: "flex", gap: 16, marginBottom: 10 }}>
          <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, fontWeight: 600, color: C.ink, cursor: "pointer" }}>
            <input type="radio" checked={form.allDay} onChange={() => setForm((f) => ({ ...f, allDay: true }))} /> Block whole day(s)
          </label>
          <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, fontWeight: 600, color: C.ink, cursor: "pointer" }}>
            <input type="radio" checked={!form.allDay} onChange={() => setForm((f) => ({ ...f, allDay: false }))} /> Specific hours, on chosen weekdays
          </label>
        </div>

        {form.allDay ? (
          <div style={{ fontSize: 12.5, color: C.mute }}>Every day from {dayLabel(form.startDate)} to {dayLabel(form.endDate)} will be fully blocked — good for holiday periods or maintenance closures.</div>
        ) : (
          <div>
            <div style={{ display: "flex", gap: 6, marginBottom: 10, flexWrap: "wrap" }}>
              {WEEKDAY_ORDER.map((d) => (
                <button key={d} type="button" onClick={() => toggleDay(d)} style={{
                  padding: "5px 10px", borderRadius: 7, fontSize: 12.5, fontWeight: 700, cursor: "pointer",
                  border: `1px solid ${form.days.includes(d) ? C.cyan : C.line}`,
                  background: form.days.includes(d) ? C.cyanSoft : "#fff", color: form.days.includes(d) ? C.navy : C.mute,
                }}>{DAY_NAMES[d]}</button>
              ))}
            </div>
            <div className="dp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
              <Field label="Blocked from">
                <select style={inputStyle} value={form.startTime} onChange={set("startTime")}>
                  {TIME_SLOTS.map((t) => <option key={t} value={t}>{t}</option>)}
                </select>
              </Field>
              <Field label="Blocked until">
                <select style={inputStyle} value={form.endTime} onChange={set("endTime")}>
                  {TIME_SLOTS.map((t) => <option key={t} value={t}>{t}</option>)}
                </select>
              </Field>
            </div>
            <div style={{ fontSize: 12.5, color: C.mute, marginTop: -8 }}>On the days ticked above, between {form.startTime} and {form.endTime}, from {dayLabel(form.startDate)} to {dayLabel(form.endDate)} — good for recurring school-day hours across a whole term.</div>
          </div>
        )}
      </div>

      {!valid && <div style={{ fontSize: 12.5, color: C.coral, marginBottom: 10 }}>Fill in a label, valid dates, and — if using specific hours — at least one weekday with an end time after the start time.</div>}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 6 }}>
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="accent" disabled={!valid} onClick={() => onSave({ ...form, id: initial?.id })}>{initial ? "Save changes" : "Create rule"}</Btn>
      </div>
    </Modal>
  );
}

// ---------- dashboard ----------
function Dashboard({ sites, facilities, bookings, siteById, facilityById, goBookings, scoped }) {
  const next7End = addDays(todayISO(), 7);
  const upcomingLettings = bookings
    .filter((b) => b.status === "confirmed" && b.date >= todayISO() && b.date <= next7End)
    .sort((a, b) => a.date.localeCompare(b.date) || a.startTime.localeCompare(b.startTime));
  const upcoming = bookings.filter((b) => b.status === "confirmed" && b.date >= todayISO()).sort((a, b) => a.date.localeCompare(b.date) || a.startTime.localeCompare(b.startTime)).slice(0, 6);
  const recentCancellations = bookings.filter((b) => b.status === "cancelled").sort((a, b) => b.date.localeCompare(a.date)).slice(0, 6);
  const thisMonth = monthKey(todayISO());
  const monthRevenue = bookings.filter((b) => b.status === "confirmed" && monthKey(b.date) === thisMonth).reduce((s, b) => s + (Number(b.price) || 0), 0);
  const totalRevenue = bookings.filter((b) => b.status === "confirmed").reduce((s, b) => s + (Number(b.price) || 0), 0);

  const stats = [
    { label: "Sites", value: sites.length, icon: MapPin, color: C.navy },
    { label: "Facilities", value: facilities.length, icon: Building2, color: C.navy },
    { label: "Cancelled bookings", value: bookings.filter((b) => b.status === "cancelled").length, icon: Ban, color: C.coral },
    { label: "This month's revenue", value: money(monthRevenue), icon: PoundSterling, color: C.pitch },
  ];

  return (
    <div>
      <PageHeader title="Dashboard" sub={`Total confirmed revenue to date: ${money(totalRevenue)}`} />

      {scoped && (
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, marginBottom: 26, overflow: "hidden" }}>
          <div style={{ padding: "14px 18px", background: C.cyan }}>
            <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, color: C.navy }}>Upcoming lettings (next 7 days)</span>
          </div>
          {upcomingLettings.length === 0 ? (
            <div style={{ padding: 24, color: C.mute, fontSize: 13.5 }}>No confirmed lettings in the next 7 days.</div>
          ) : (
            <div style={{ maxHeight: 337, overflowY: "auto", direction: "rtl" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", direction: "ltr" }}>
                <thead>
                  <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                    {["Date", "Time", "Facility", "Hirer"].map((h) => (
                      <th key={h} style={{ padding: "9px 18px", fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.5, color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}`, position: "sticky", top: 0, background: "#F8FAFC" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {upcomingLettings.map((b) => (
                    <tr key={b.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                      <td style={{ padding: "10px 18px", fontSize: 13, background: C.card }}>{dayLabel(b.date)}</td>
                      <td style={{ padding: "10px 18px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", background: C.card }}>{b.startTime}–{b.endTime}</td>
                      <td style={{ padding: "10px 18px", fontSize: 13, background: C.card }}>{facilityById[b.facilityId]?.name || "Unknown"}</td>
                      <td style={{ padding: "10px 18px", fontSize: 13, fontWeight: 600, color: C.ink, background: C.card }}>{b.hirerName}{b.company ? ` — ${b.company}` : ""}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      <div className="dp-stat-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 26 }}>
        {stats.map((s) => (
          <div key={s.label} style={{ background: C.card, borderRadius: 12, padding: "16px 18px", border: `1px solid ${C.line}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, color: C.mute, fontSize: 12.5, fontWeight: 600, marginBottom: 8 }}>
              <s.icon size={14} color={s.color} /> {s.label}
            </div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, color: C.navy }}>{s.value}</div>
          </div>
        ))}
      </div>

      <div className="dp-dash-grid" style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 18 }}>
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
          <div style={{ padding: "14px 18px", background: C.cyan, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, color: C.navy }}>Upcoming confirmed bookings</span>
            <Btn variant="ghost" small onClick={goBookings}>View all</Btn>
          </div>
          {upcoming.length === 0 ? (
            <div style={{ padding: 24, color: C.mute, fontSize: 13.5 }}>Nothing confirmed and upcoming yet.</div>
          ) : (
            <div>
              {upcoming.map((b) => (
                <div key={b.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 18px", borderBottom: `1px solid ${C.line}` }}>
                  <div>
                    <div style={{ fontWeight: 600, color: C.ink, fontSize: 14 }}>{b.hirerName} — {facilityById[b.facilityId]?.name || "Unknown facility"}</div>
                    <div style={{ fontSize: 12.5, color: C.mute, marginTop: 2 }}>{siteById[facilityById[b.facilityId]?.siteId]?.name || ""} · {dayLabel(b.date)}, {b.startTime}–{b.endTime}</div>
                  </div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: C.pitch, fontSize: 13.5 }}>{money(b.price)}</div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
          <div style={{ padding: "14px 18px", background: C.cyan, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, color: C.navy }}>Recently cancelled</span>
            <Btn variant="ghost" small onClick={goBookings}>View all</Btn>
          </div>
          {recentCancellations.length === 0 ? (
            <div style={{ padding: 24, color: C.mute, fontSize: 13.5 }}>No cancellations on record.</div>
          ) : (
            <div>
              {recentCancellations.map((b) => (
                <div key={b.id} style={{ padding: "12px 18px", borderBottom: `1px solid ${C.line}` }}>
                  <div style={{ fontWeight: 600, color: C.ink, fontSize: 13.5 }}>{b.hirerName}</div>
                  <div style={{ fontSize: 12, color: C.mute, marginTop: 2 }}>{facilityById[b.facilityId]?.name} · {dayLabel(b.date)}, {b.startTime}–{b.endTime}</div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function PageHeader({ title, sub, action }) {
  return (
    <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "flex-end", gap: 12, marginBottom: 20 }}>
      <div>
        <h1 style={{ margin: 0, fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, color: C.navy, fontWeight: 700 }}>{title}</h1>
        {sub && <div style={{ color: C.mute, fontSize: 13.5, marginTop: 4 }}>{sub}</div>}
      </div>
      {action}
    </div>
  );
}

// ---------- bookings ----------
function Bookings({ sites, facilities, bookings, setBookings, facilityById, siteById, goFacilities, members, setMembers, memberById, goMembers, blackouts }) {
  const [view, setView] = useState("list");
  const [modalOpen, setModalOpen] = useState(false);
  const [editing, setEditing] = useState(null);
  const [prefill, setPrefill] = useState(null);
  const [filterSite, setFilterSite] = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterMember, setFilterMember] = useState("");
  const [selectedIds, setSelectedIds] = useState(() => new Set());
  const [calSite, setCalSite] = useState("all");
  const [calFacility, setCalFacility] = useState("all");
  const [gridSite, setGridSite] = useState(sites[0]?.id || "");
  const [gridDate, setGridDate] = useState(todayISO());
  const [importOpen, setImportOpen] = useState(false);
  const [skippedMsg, setSkippedMsg] = useState("");
  const [confirmBulkDelete, setConfirmBulkDelete] = useState(false);

  useEffect(() => {
    if (!gridSite && sites.length > 0) setGridSite(sites[0].id);
  }, [sites, gridSite]);

  const memberQuery = filterMember.trim().toLowerCase();
  const filtered = bookings
    .filter((b) => filterSite === "all" || facilityById[b.facilityId]?.siteId === filterSite)
    .filter((b) => filterStatus === "all" || b.status === filterStatus)
    .filter((b) => {
      if (!memberQuery) return true;
      const member = memberById?.[b.memberId];
      const haystack = [b.hirerName, b.company, member?.name, member?.company].filter(Boolean).join(" ").toLowerCase();
      return haystack.includes(memberQuery);
    })
    .sort((a, b) => b.date.localeCompare(a.date) || a.startTime.localeCompare(b.startTime));

  function toggleSelect(id) {
    setSelectedIds((prev) => { const next = new Set(prev); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  }
  function toggleSelectAll() {
    setSelectedIds((prev) => (filtered.length > 0 && filtered.every((b) => prev.has(b.id)) ? new Set() : new Set(filtered.map((b) => b.id))));
  }
  function removeSelected() {
    // A selected booking that's part of a multi-facility group takes the
    // rest of its group with it, even if those legs weren't individually
    // checked — it's really one booking split across facilities.
    const idsToRemove = new Set(selectedIds);
    bookings.forEach((b) => {
      if (selectedIds.has(b.id) && b.groupId) {
        bookings.forEach((x) => { if (x.groupId === b.groupId) idsToRemove.add(x.id); });
      }
    });
    setBookings((bs) => bs.filter((b) => !idsToRemove.has(b.id)));
    setSelectedIds(new Set());
  }

  const calFacilityOptions = facilities.filter((f) => calSite === "all" || f.siteId === calSite);
  function onCalSiteChange(v) {
    setCalSite(v);
    if (v !== "all" && calFacility !== "all" && facilityById[calFacility]?.siteId !== v) setCalFacility("all");
  }
  const calBookings = bookings
    .filter((b) => calSite === "all" || facilityById[b.facilityId]?.siteId === calSite)
    .filter((b) => calFacility === "all" || b.facilityId === calFacility);

  const gridFacilities = facilities.filter((f) => f.siteId === gridSite);

  function openNew(pre) {
    setEditing(null);
    setPrefill(pre || null);
    setModalOpen(true);
  }

  function setStatus(id, status) {
    const target = bookings.find((b) => b.id === id);
    const groupIds = new Set(bookingsInGroup(target, bookings).map((b) => b.id));
    setBookings((bs) => bs.map((b) => (groupIds.has(b.id) ? { ...b, status } : b)));
  }
  function remove(id) {
    const target = bookings.find((b) => b.id === id);
    const groupIds = new Set(bookingsInGroup(target, bookings).map((b) => b.id));
    setBookings((bs) => bs.filter((b) => !groupIds.has(b.id)));
  }
  function save(data) {
    if (data.id) {
      setBookings((bs) => bs.map((b) => (b.id === data.id ? data : b)));
      setModalOpen(false);
      setEditing(null);
      setPrefill(null);
      return;
    }
    // new booking(s) — handle weekly series, an explicit set of picked dates,
    // and/or multiple facilities booked together (combinable with picked dates,
    // one group of facilities per date — see BookingModal's extraFacilityOptions)
    const isBlackedOutFor = (fid, d) => {
      const siteId = facilityById[fid]?.siteId;
      return !!findBlackout(blackouts, { facilityId: fid, siteId, date: d, startTime: data.startTime, endTime: data.endTime });
    };

    const newBookings = [];
    const skippedDates = [];
    const multiFacilityIds = data.facilityIds && data.facilityIds.length > 1 ? [...new Set(data.facilityIds)] : null;
    if (multiFacilityIds && data.repeatMode === "dates" && data.multiDates?.length > 0) {
      // Multiple facilities, each booked on every picked date — one row per
      // facility per date, grouped per date (so cancelling/deleting one date
      // only affects that date's facilities, not the whole series) and linked
      // by a shared recurringId across the series for the "repeat" indicator.
      const recurringId = uid();
      [...data.multiDates].sort().forEach((d) => {
        if (multiFacilityIds.some((fid) => isBlackedOutFor(fid, d))) { skippedDates.push(d); return; }
        const groupId = uid();
        multiFacilityIds.forEach((fid, i) => {
          newBookings.push({ ...data, id: uid(), facilityId: fid, date: d, price: i === 0 ? data.price : 0, groupId, recurringId });
        });
      });
    } else if (multiFacilityIds) {
      // One-off multi-facility booking — one real booking row per facility (so
      // each shows up on its own calendar/grid), linked by groupId so they act
      // as one event — only the first leg carries the price, so totals
      // elsewhere don't multiply-count the same money across facilities.
      const groupId = uid();
      multiFacilityIds.forEach((fid, i) => {
        newBookings.push({ ...data, id: uid(), facilityId: fid, price: i === 0 ? data.price : 0, groupId });
      });
    } else if (data.repeatMode === "weekly" && data.repeatUntil && data.repeatUntil >= data.date) {
      const recurringId = uid();
      let d = data.date;
      while (d <= data.repeatUntil) {
        if (isBlackedOutFor(data.facilityId, d)) skippedDates.push(d);
        else newBookings.push({ ...data, id: uid(), date: d, recurringId });
        d = addDays(d, 7);
      }
    } else if (data.repeatMode === "dates" && data.multiDates?.length > 0) {
      const recurringId = uid();
      [...data.multiDates].sort().forEach((d) => {
        if (isBlackedOutFor(data.facilityId, d)) skippedDates.push(d);
        else newBookings.push({ ...data, id: uid(), date: d, recurringId });
      });
    } else {
      newBookings.push({ ...data, id: uid() });
    }
    setBookings((bs) => [...bs, ...newBookings]);
    setModalOpen(false);
    setEditing(null);
    setPrefill(null);
    if (skippedDates.length > 0) {
      setSkippedMsg(
        `Created ${newBookings.length} booking${newBookings.length === 1 ? "" : "s"} — skipped ${skippedDates.length} date${skippedDates.length === 1 ? "" : "s"} blocked by a blackout rule: ${skippedDates.map(dayLabel).join(", ")}.`
      );
      setTimeout(() => setSkippedMsg(""), 10000);
    }
    if (data.emailConfirmation && newBookings.length > 0) {
      const member = memberById[data.memberId];
      const site = siteById[siteId];
      const rows = buildDocumentRows(newBookings, facilityById, site?.name);
      const total = rows.reduce((s, r) => s + r.price, 0);
      const dates = newBookings.map((b) => b.date).sort();
      const periodLabel = dates[0] === dates[dates.length - 1] ? dayLabel(dates[0]) : `${dayLabel(dates[0])} – ${dayLabel(dates[dates.length - 1])}`;
      buildConfirmationPDF({
        reference: newConfirmationReference(), issuedDate: dayLabel(todayISO()), member, site, periodLabel, rows, total,
      }).then((doc) => {
        const safeName = (member?.name || "confirmation").replace(/[^a-z0-9]+/gi, "-").toLowerCase();
        downloadBlob(`delsport-confirmation-${safeName}-${todayISO()}.pdf`, doc.output("blob"));
      }).catch((e) => console.error("Failed to build confirmation PDF", e));
    }
  }

  if (facilities.length === 0) {
    return (
      <div>
        <PageHeader title="Bookings" />
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
          <EmptyState icon={Building2} title="Add a facility first" sub="You'll need at least one site and facility before you can create bookings." action={<Btn variant="accent" onClick={goFacilities}>Go to Sites & Facilities</Btn>} />
        </div>
      </div>
    );
  }

  return (
    <div>
      <PageHeader
        title="Bookings"
        sub={`${bookings.length} total · ${bookings.filter((b) => b.status === "confirmed").length} confirmed · ${bookings.filter((b) => b.status === "cancelled").length} cancelled`}
        action={
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="ghost" icon={Upload} onClick={() => setImportOpen(true)}>Import from CSV</Btn>
            <Btn variant="accent" icon={Plus} onClick={() => openNew()}>New booking</Btn>
          </div>
        }
      />
      {skippedMsg && (
        <div style={{ background: C.amberSoft, color: C.amber, fontSize: 13, fontWeight: 600, padding: "10px 14px", borderRadius: 9, marginBottom: 16 }}>{skippedMsg}</div>
      )}
      {importOpen && (
        <ImportBookingsModal
          sites={sites}
          facilities={facilities}
          members={members}
          onClose={() => setImportOpen(false)}
          onImport={({ newMembers, newBookings }) => {
            if (newMembers.length > 0) setMembers((ms) => [...ms, ...newMembers]);
            setBookings((bs) => [...bs, ...newBookings]);
            setImportOpen(false);
          }}
        />
      )}

      <div style={{ display: "flex", gap: 10, marginBottom: 16, alignItems: "center", flexWrap: "wrap" }}>
        <div style={{ display: "flex", background: "#EAEEF3", borderRadius: 9, padding: 3 }}>
          {[["list", "List"], ["calendar", "Week"], ["grid", "Grid"]].map(([v, label]) => (
            <button key={v} onClick={() => setView(v)} style={{
              padding: "6px 14px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
              background: view === v ? C.card : "transparent", color: view === v ? C.navy : C.mute,
            }}>{label}</button>
          ))}
        </div>
        {view === "list" && (
          <>
            <select value={filterSite} onChange={(e) => setFilterSite(e.target.value)} style={{ ...inputStyle, width: "auto" }}>
              <option value="all">All sites</option>
              {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} style={{ ...inputStyle, width: "auto" }}>
              <option value="all">All statuses</option>
              {Object.entries(STATUS_META).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
            <input style={{ ...inputStyle, width: 220 }} placeholder="Search by member…" value={filterMember} onChange={(e) => setFilterMember(e.target.value)} />
            <Btn small variant="ghost" icon={CalendarDays} onClick={() => downloadICS("delsport-bookings.ics", buildICS(filtered.filter((b) => b.status !== "declined" && b.status !== "cancelled"), facilityById, siteById))}>Export to calendar (.ics)</Btn>
            {selectedIds.size > 0 && (
              <Btn small variant="danger" icon={Trash2} onClick={() => setConfirmBulkDelete(true)}>Delete {selectedIds.size} selected</Btn>
            )}
          </>
        )}
        {view === "calendar" && (
          <>
            <select value={calSite} onChange={(e) => onCalSiteChange(e.target.value)} style={{ ...inputStyle, width: "auto" }}>
              <option value="all">All sites</option>
              {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <select value={calFacility} onChange={(e) => setCalFacility(e.target.value)} style={{ ...inputStyle, width: "auto" }}>
              <option value="all">All facilities</option>
              {calFacilityOptions.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
            </select>
          </>
        )}
        {view === "grid" && (
          <select value={gridSite} onChange={(e) => setGridSite(e.target.value)} style={{ ...inputStyle, width: "auto" }}>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        )}
      </div>

      {view === "list" && (
        <BookingsList bookings={filtered} allBookings={bookings} facilityById={facilityById} siteById={siteById} memberById={memberById} setStatus={setStatus} remove={remove} onEdit={(b) => { setEditing(b); setPrefill(null); setModalOpen(true); }} selectedIds={selectedIds} toggleSelect={toggleSelect} toggleSelectAll={toggleSelectAll} />
      )}
      {view === "calendar" && (
        <BookingsCalendar bookings={calBookings} facilityById={facilityById} showFacilityName={calFacility === "all"} selectedFacility={calFacility !== "all" ? facilityById[calFacility] : null} setStatus={setStatus} />
      )}
      {view === "grid" && (
        gridFacilities.length === 0 ? (
          <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
            <EmptyState icon={Building2} title="No facilities at this site" sub="Add a facility to this site to see it on the grid." />
          </div>
        ) : (
          <PlannerGrid
            date={gridDate} setDate={setGridDate}
            facilities={gridFacilities} bookings={bookings}
            blackouts={blackouts} siteId={gridSite}
            onCellClick={(facilityId, startTime, endTime) => openNew({ facilityId, date: gridDate, startTime, endTime })}
            onBlockClick={(b) => { setEditing(b); setPrefill(null); setModalOpen(true); }}
          />
        )
      )}

      {modalOpen && (
        <BookingModal
          initial={editing}
          prefill={prefill}
          facilities={facilities}
          siteById={siteById}
          members={members}
          setMembers={setMembers}
          goMembers={goMembers}
          allBookings={bookings}
          blackouts={blackouts}
          onClose={() => { setModalOpen(false); setEditing(null); setPrefill(null); }}
          onSave={save}
        />
      )}
      {confirmBulkDelete && (
        <ConfirmDeleteModal
          title="Delete selected bookings?"
          message={`${selectedIds.size} booking${selectedIds.size === 1 ? "" : "s"} will be permanently removed. This can't be undone.`}
          onCancel={() => setConfirmBulkDelete(false)}
          onConfirm={() => { removeSelected(); setConfirmBulkDelete(false); }}
        />
      )}
    </div>
  );
}

// ---------- planner grid (SchoolBooking-style: facilities as columns, time slots as rows) ----------
const GRID_SLOTS = TIME_SLOTS.filter((t) => t >= "07:00" && t < "21:30");

function PlannerGrid({ date, setDate, facilities, bookings, blackouts, siteId, onCellClick, onBlockClick }) {
  const dayBookings = bookings.filter((b) => b.date === date);
  const hasClash = facilities.some((f) => {
    const active = dayBookings.filter((b) => b.facilityId === f.id && b.status !== "declined" && b.status !== "cancelled");
    return GRID_SLOTS.some((s) => {
      const e = TIME_SLOTS[TIME_SLOTS.indexOf(s) + 1] || "22:00";
      const inSlot = active.filter((b) => timesOverlap(b.startTime, b.endTime, s, e));
      return spacesUsed(inSlot) > (f.capacity || 1);
    });
  });

  return (
    <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", borderBottom: `1px solid ${C.line}`, background: C.navy, gap: 10, flexWrap: "wrap" }}>
        <Btn small variant="ghost" icon={ChevronLeft} onClick={() => setDate(addDays(date, -1))} />
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ color: "#fff", fontWeight: 600, fontFamily: "'Space Grotesk', sans-serif", fontSize: 14 }}>{dayLabel(date)}</span>
          <input type="date" value={date} onChange={(e) => e.target.value && setDate(e.target.value)}
            style={{ padding: "5px 8px", borderRadius: 7, border: "none", fontSize: 12.5, fontFamily: "'Inter', sans-serif", color: C.navy }} />
          {date !== todayISO() && <Btn small variant="accent" onClick={() => setDate(todayISO())}>Today</Btn>}
        </div>
        <Btn small variant="ghost" icon={ChevronRight} onClick={() => setDate(addDays(date, 1))} />
      </div>
      <div style={{
        padding: "9px 16px", fontSize: 12.5, fontWeight: 700, display: "flex", alignItems: "center", gap: 7,
        background: hasClash ? C.coralSoft : C.pitchSoft, color: hasClash ? C.coral : C.pitch, borderBottom: `1px solid ${C.line}`,
      }}>
        <AlertCircle size={14} /> {hasClash ? "Clash detected — a facility is over capacity for part of this day" : "No clashes detected — all bookings validated"}
      </div>
      <div style={{ overflowX: "auto" }}>
        <div style={{ display: "grid", gridTemplateColumns: `72px repeat(${facilities.length}, minmax(140px, 1fr))`, minWidth: 72 + facilities.length * 140 }}>
          <div style={{ background: "#F8FAFC", borderBottom: `1px solid ${C.line}`, borderRight: `1px solid ${C.line}` }} />
          {facilities.map((f) => (
            <div key={f.id} style={{ background: "#F8FAFC", borderBottom: `1px solid ${C.line}`, borderRight: `1px solid ${C.line}`, padding: "8px 10px" }}>
              <div style={{ fontWeight: 700, fontSize: 12.5, color: C.navy }}>{f.name}</div>
              <div style={{ fontSize: 10.5, color: C.mute }}>{f.capacity && f.capacity > 1 ? `${f.capacity} spaces` : "1 space"}</div>
            </div>
          ))}
          {GRID_SLOTS.map((s, si) => {
            const e = TIME_SLOTS[TIME_SLOTS.indexOf(s) + 1] || "22:00";
            const onHour = s.endsWith(":00");
            return (
              <FragmentRow key={s} time={s} end={e} onHour={onHour} facilities={facilities} dayBookings={dayBookings} date={date}
                blackouts={blackouts} siteId={siteId} onCellClick={onCellClick} onBlockClick={onBlockClick} />
            );
          })}
        </div>
      </div>
      <div style={{ padding: "8px 16px", fontSize: 11.5, color: C.mute, borderTop: `1px solid ${C.line}` }}>Click an empty cell to start a booking in that slot · click an existing booking to edit it. Hatched cells are blocked by an admin rule.</div>
    </div>
  );
}

function FragmentRow({ time, end, onHour, facilities, dayBookings, date, blackouts, siteId, onCellClick, onBlockClick }) {
  return (
    <>
      <div style={{ borderBottom: `1px solid ${C.line}`, borderRight: `1px solid ${C.line}`, padding: "4px 8px", fontSize: 10.5, color: C.mute, fontFamily: "'IBM Plex Mono', monospace", background: onHour ? "#FBFCFD" : "#fff" }}>
        {onHour ? time : ""}
      </div>
      {facilities.map((f) => {
        const active = dayBookings.filter((b) => b.facilityId === f.id && b.status !== "declined" && b.status !== "cancelled" && timesOverlap(b.startTime, b.endTime, time, end));
        const used = spacesUsed(active);
        const capacity = f.capacity || 1;
        const blocked = active.length === 0 && blackouts && findBlackout(blackouts, { facilityId: f.id, siteId, date, startTime: time, endTime: end });
        return (
          <div key={f.id}
            title={blocked ? blocked.label : undefined}
            onClick={() => { if (active.length > 0) onBlockClick(active[0]); else onCellClick(f.id, time, end); }}
            style={{
              borderBottom: `1px solid ${C.line}`, borderRight: `1px solid ${C.line}`, minHeight: 30, padding: "2px 4px",
              background: active.length > 0 ? (statusMeta(active[0].status)?.bg || C.pitchSoft)
                : blocked ? "repeating-linear-gradient(135deg, #EEF1F4, #EEF1F4 5px, #E2E7EC 5px, #E2E7EC 10px)"
                : (onHour ? "#FBFCFD" : "#fff"),
              cursor: "pointer", display: "flex", flexDirection: "column", gap: 2, justifyContent: "center",
            }}>
            {active.map((b) => (
              <div key={b.id} style={{ fontSize: 10.5, color: statusMeta(b.status)?.fg, fontWeight: 700, lineHeight: 1.2, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                {b.hirerName}{(Number(b.spaces) || 1) > 1 ? ` ×${b.spaces}` : ""}
              </div>
            ))}
            {active.length === 0 && blocked && <div style={{ display: "flex", alignItems: "center", justifyContent: "center", color: C.mute }}><Lock size={10} /></div>}
            {active.length === 0 && !blocked && <div style={{ fontSize: 13, color: "#CBD5E1", textAlign: "center" }}>+</div>}
          </div>
        );
      })}
    </>
  );
}

function ImportBookingsModal({ sites, facilities, members, onClose, onImport }) {
  const [rows, setRows] = useState(null);
  const [error, setError] = useState("");

  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    const reader = new FileReader();
    reader.onload = () => {
      const parsed = csvRowsToObjects(parseCSVText(String(reader.result)));
      if (parsed.length === 0) { setError("Couldn't find any rows in that file."); setRows(null); return; }
      setRows(parsed);
    };
    reader.readAsText(file);
  }

  const preview = useMemo(() => {
    if (!rows) return null;
    const newMembers = [];
    const memberByKey = new Map(); // "email:..."/"name:..." -> existing or newly-created member
    members.forEach((m) => {
      if (m.email) memberByKey.set(`email:${normKey(m.email)}`, m);
      memberByKey.set(`name:${normKey(m.name)}`, m);
    });
    function resolveMember(name, email, phone, company) {
      const emailKey = email ? `email:${normKey(email)}` : null;
      const nameKey = `name:${normKey(name)}`;
      const existing = (emailKey && memberByKey.get(emailKey)) || memberByKey.get(nameKey);
      if (existing) return existing;
      const created = { id: uid(), name, email, phone, company, siteIds: [] };
      newMembers.push(created);
      if (emailKey) memberByKey.set(emailKey, created);
      memberByKey.set(nameKey, created);
      return created;
    }

    const ready = [];
    const errors = [];
    rows.forEach((row, i) => {
      const rowNum = i + 2; // +1 for header row, +1 for 1-indexing
      const siteName = pickField(row, ["site", "sitename", "location", "venue"]);
      const facilityName = pickField(row, ["facility", "facilityname", "space", "room", "pitch"]);
      const memberName = pickField(row, ["membername", "name", "hirername", "clientname", "hirer"]);
      const email = pickField(row, ["email", "emailaddress"]);
      const phone = pickField(row, ["phone", "phonenumber", "mobile", "tel", "telephone"]);
      const company = pickField(row, ["company", "business", "organisation", "organization"]);
      const dateRaw = pickField(row, ["date"]);
      const startRaw = pickField(row, ["starttime", "start", "from"]);
      const endRaw = pickField(row, ["endtime", "end", "to"]);
      const priceRaw = pickField(row, ["price", "pricecharged", "cost", "fee"]);
      const statusRaw = pickField(row, ["status"]);
      const spacesRaw = pickField(row, ["spaces", "courts", "qty", "quantity"]);

      if (!facilityName) { errors.push({ row: rowNum, reason: "no facility given" }); return; }
      const resolved = resolveFacilityByName(facilities, sites, siteName, facilityName);
      if (resolved.error) { errors.push({ row: rowNum, reason: resolved.error }); return; }
      if (!memberName) { errors.push({ row: rowNum, reason: "no member/hirer name given" }); return; }
      const date = parseDateFlexible(dateRaw);
      if (!date) { errors.push({ row: rowNum, reason: `unrecognised date "${dateRaw}"` }); return; }
      const startTime = parseTimeFlexible(startRaw);
      const endTime = parseTimeFlexible(endRaw);
      if (!startTime || !endTime) { errors.push({ row: rowNum, reason: "unrecognised start/end time" }); return; }
      if (timeToMin(endTime) <= timeToMin(startTime)) { errors.push({ row: rowNum, reason: "end time must be after start time" }); return; }

      const member = resolveMember(memberName, email, phone, company);
      if (member.siteIds?.length && !member.siteIds.includes(resolved.facility.siteId)) {
        const allowedNames = member.siteIds.map((id) => sites.find((s) => s.id === id)?.name).filter(Boolean).join(", ");
        errors.push({ row: rowNum, reason: `"${member.name}" is restricted to ${allowedNames}, not this facility's site` });
        return;
      }
      ready.push({
        id: uid(),
        facilityId: resolved.facility.id,
        memberId: member.id,
        date, startTime, endTime,
        purpose: pickField(row, ["purpose", "activity", "description", "reason"]),
        price: Number(String(priceRaw).replace(/[^0-9.\-]/g, "")) || 0,
        status: statusRaw ? parseStatusFlexible(statusRaw) : "confirmed",
        notes: pickField(row, ["notes", "internalnotes", "comments"]),
        spaces: Number(spacesRaw) || 1,
        hirerName: member.name,
        hirerContact: member.email || member.phone || "",
        company: member.company || "",
      });
    });

    return { ready, errors, newMembers };
  }, [rows, facilities, sites, members]);

  function downloadTemplate() {
    downloadCSV(
      "delsport-bookings-template.csv",
      ["Site", "Facility", "Member Name", "Company", "Email", "Phone", "Date", "Start Time", "End Time", "Purpose", "Price", "Status", "Notes", "Spaces"],
      [["Stourport High School", "3G/4G Pitch", "Jamie Carter", "Riverside FC", "jamie@riversidefc.co.uk", "07700 900123", "05/08/2026", "17:00", "18:00", "5-a-side league", "45.00", "Confirmed", "", "1"]]
    );
  }

  return (
    <Modal title="Import bookings from CSV" onClose={onClose} wide>
      <div style={{ fontSize: 13, color: C.mute, marginBottom: 14, lineHeight: 1.6 }}>
        One row per booking. Needs a Site + Facility (facility names alone can be ambiguous across sites), a member/hirer name, a date, and a start/end time — everything else is optional. A member is matched by email or name if they're already on file, otherwise a new one is created automatically. If an existing member is restricted to specific sites, a row booking them elsewhere is rejected. Capacity and blackout rules aren't checked during import, since this is meant for bringing across bookings you've already agreed elsewhere.{" "}
        <button type="button" onClick={downloadTemplate} style={{ background: "none", border: "none", color: C.cyan, fontWeight: 700, cursor: "pointer", padding: 0, fontSize: 13 }}>Download a template</button>.
      </div>
      <input type="file" accept=".csv,text/csv" onChange={handleFile} style={{ fontSize: 13, marginBottom: 12 }} />
      {error && <div style={{ fontSize: 12.5, color: C.coral, marginBottom: 10 }}>{error}</div>}

      {preview && (
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 13.5, fontWeight: 600, color: C.navy, marginBottom: 8 }}>
            {preview.ready.length} booking{preview.ready.length === 1 ? "" : "s"} ready to import
            {preview.newMembers.length > 0 ? ` (${preview.newMembers.length} new member${preview.newMembers.length === 1 ? "" : "s"} will be added)` : ""}
            {preview.errors.length > 0 ? `, ${preview.errors.length} skipped` : ""}
          </div>
          {preview.ready.length > 0 && (
            <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "hidden", maxHeight: 240, overflowY: "auto", marginBottom: preview.errors.length ? 10 : 0 }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                    {["Hirer", "Facility", "Date", "Time", "Price", "Status"].map((h) => (
                      <th key={h} style={{ padding: "7px 12px", fontSize: 11, textTransform: "uppercase", color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.ready.map((b) => (
                    <tr key={b.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                      <td style={{ padding: "6px 12px", fontSize: 12.5 }}>{b.hirerName}</td>
                      <td style={{ padding: "6px 12px", fontSize: 12.5 }}>{facilities.find((f) => f.id === b.facilityId)?.name}</td>
                      <td style={{ padding: "6px 12px", fontSize: 12.5 }}>{dayLabel(b.date)}</td>
                      <td style={{ padding: "6px 12px", fontSize: 12.5, fontFamily: "'IBM Plex Mono', monospace" }}>{b.startTime}–{b.endTime}</td>
                      <td style={{ padding: "6px 12px", fontSize: 12.5 }}>{money(b.price)}</td>
                      <td style={{ padding: "6px 12px" }}><StatusPill status={b.status} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          {preview.errors.length > 0 && (
            <div style={{ fontSize: 12, color: C.mute }}>Skipped: {preview.errors.map((e) => `row ${e.row} (${e.reason})`).join("; ")}</div>
          )}
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="accent" disabled={!preview || preview.ready.length === 0} onClick={() => onImport({ newMembers: preview.newMembers, newBookings: preview.ready })}>
          Import {preview?.ready.length || ""} booking{preview?.ready.length === 1 ? "" : "s"}
        </Btn>
      </div>
    </Modal>
  );
}

function BookingsList({ bookings, allBookings, facilityById, siteById, memberById, setStatus, remove, onEdit, selectedIds, toggleSelect, toggleSelectAll }) {
  const [confirmDeleteBooking, setConfirmDeleteBooking] = useState(null);
  if (bookings.length === 0) {
    return (
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
        <EmptyState icon={CalendarDays} title="No bookings match" sub="Try clearing filters, or create a new booking." />
      </div>
    );
  }
  const allSelected = bookings.length > 0 && bookings.every((b) => selectedIds.has(b.id));
  return (
    <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
            <th style={{ padding: "10px 16px", width: 1, borderBottom: `1px solid ${C.line}` }}>
              <input type="checkbox" checked={allSelected} onChange={toggleSelectAll} />
            </th>
            {["Hirer", "Facility", "Date", "Time", "Price", "Status", ""].map((h) => (
              <th key={h} style={{ padding: "10px 16px", fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.5, color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {bookings.map((b) => {
            const f = facilityById[b.facilityId];
            const member = memberById?.[b.memberId];
            const canEmail = !!member?.email;
            const emailKind = b.status === "cancelled" || b.status === "declined" ? "cancelled" : "confirmed";
            return (
              <tr key={b.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                <td style={{ padding: "11px 16px" }}>
                  <input type="checkbox" checked={selectedIds.has(b.id)} onChange={() => toggleSelect(b.id)} />
                </td>
                <td style={{ padding: "11px 16px" }}>
                  <div style={{ fontWeight: 600, fontSize: 13.5, color: C.ink }}>{b.hirerName}{(Number(b.spaces) || 1) > 1 && <span style={{ marginLeft: 6, fontSize: 11, fontWeight: 700, color: C.navy, background: C.cyanSoft, padding: "1px 6px", borderRadius: 999 }}>×{b.spaces}</span>}{b.recurringId && <Repeat size={11} style={{ marginLeft: 5, verticalAlign: -1, color: C.mute }} />}{b.groupId && <Link2 size={11} title="Part of a multi-facility booking" style={{ marginLeft: 5, verticalAlign: -1, color: C.cyan }} />}</div>
                  <div style={{ fontSize: 12, color: C.mute }}>{[b.company, b.hirerContact].filter(Boolean).join(" · ")}</div>
                </td>
                <td style={{ padding: "11px 16px", fontSize: 13 }}>
                  <div>{f?.name || "—"}</div>
                  <div style={{ fontSize: 11.5, color: C.mute, display: "flex", alignItems: "center", gap: 4, marginTop: 2 }}>
                    {siteById[f?.siteId]?.logoUrl && <img src={siteById[f.siteId].logoUrl} alt="" style={{ width: 13, height: 13, objectFit: "contain" }} onError={(e) => { e.target.style.display = "none"; }} />}
                    {siteById[f?.siteId]?.name}
                  </div>
                </td>
                <td style={{ padding: "11px 16px", fontSize: 13 }}>{dayLabel(b.date)}</td>
                <td style={{ padding: "11px 16px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace" }}>{b.startTime}–{b.endTime}</td>
                <td style={{ padding: "11px 16px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{money(b.price)}</td>
                <td style={{ padding: "11px 16px" }}><StatusPill status={b.status} /></td>
                <td style={{ padding: "11px 16px" }}>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end", flexWrap: "wrap" }}>
                    {b.status === "confirmed" && (
                      <Btn small variant="danger" icon={Ban} onClick={() => setStatus(b.id, "cancelled")}>Cancel</Btn>
                    )}
                    {b.status === "cancelled" && (
                      <Btn small variant="success" icon={Check} onClick={() => setStatus(b.id, "confirmed")}>Reinstate</Btn>
                    )}
                    {canEmail && (
                      <a href={buildMailto({ booking: b, member, facilityName: f?.name, siteName: siteById[f?.siteId]?.name, kind: emailKind })} title={`Email ${emailKind}`}
                        style={{ display: "inline-flex", alignItems: "center", padding: "6px 8px", borderRadius: 8, background: C.cyanSoft, color: C.navy }}>
                        <Mail size={14} />
                      </a>
                    )}
                    <button title="Add to calendar" onClick={() => downloadICS(`${b.hirerName.replace(/\s+/g, "-") || "booking"}-${b.date}.ics`, buildICS([b], facilityById, siteById))}
                      style={{ display: "inline-flex", alignItems: "center", padding: "6px 8px", borderRadius: 8, background: "#EEF1F4", color: C.navy, border: "none", cursor: "pointer" }}>
                      <CalendarDays size={14} />
                    </button>
                    <Btn small variant="ghost" icon={Pencil} onClick={() => onEdit(b)} />
                    <Btn small variant="ghost" icon={Trash2} onClick={() => setConfirmDeleteBooking(b)} />
                  </div>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
      {confirmDeleteBooking && (
        <ConfirmDeleteModal
          title="Delete this booking?"
          message={(() => {
            const linked = bookingsInGroup(confirmDeleteBooking, allBookings || bookings).filter((b) => b.id !== confirmDeleteBooking.id);
            const base = `${confirmDeleteBooking.hirerName}'s booking on ${dayLabel(confirmDeleteBooking.date)} will be permanently removed.`;
            const groupNote = linked.length > 0
              ? ` It's linked to ${linked.length} other facilit${linked.length === 1 ? "y" : "ies"} (${linked.map((b) => facilityById[b.facilityId]?.name).filter(Boolean).join(", ")}) as one booking — those will be removed too.`
              : "";
            return `${base}${groupNote} This can't be undone.`;
          })()}
          onCancel={() => setConfirmDeleteBooking(null)}
          onConfirm={() => { remove(confirmDeleteBooking.id); setConfirmDeleteBooking(null); }}
        />
      )}
    </div>
  );
}

function BookingsCalendar({ bookings, facilityById, showFacilityName, selectedFacility, setStatus }) {
  const [weekStart, setWeekStart] = useState(startOfWeek(todayISO()));
  const days = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));
  const DAY_START = 7 * 60, DAY_END = 22 * 60;
  const totalMin = DAY_END - DAY_START;

  // assign each day's bookings to non-overlapping columns so concurrent bookings sit side by side
  function layoutDay(dayBookings) {
    const sorted = [...dayBookings].sort((a, b) => timeToMin(a.startTime) - timeToMin(b.startTime));
    const colEnds = []; // end time (min) currently occupying each column
    const placed = sorted.map((b) => {
      const start = timeToMin(b.startTime), end = timeToMin(b.endTime);
      let col = colEnds.findIndex((e) => e <= start);
      if (col === -1) { col = colEnds.length; colEnds.push(end); } else { colEnds[col] = end; }
      return { b, col };
    });
    const cols = Math.max(1, colEnds.length);
    return { placed, cols };
  }

  return (
    <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
      {selectedFacility && (
        <div style={{ padding: "10px 16px", background: C.cyanSoft, borderBottom: `1px solid ${C.line}`, fontSize: 12.5, color: C.navy, fontWeight: 600 }}>
          {selectedFacility.name} · {selectedFacility.capacity && selectedFacility.capacity > 1 ? `${selectedFacility.capacity} spaces per slot` : "1 space per slot"}
        </div>
      )}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", borderBottom: `1px solid ${C.line}`, background: C.navy }}>
        <Btn small variant="ghost" icon={ChevronLeft} onClick={() => setWeekStart(addDays(weekStart, -7))} />
        <span style={{ color: "#fff", fontWeight: 600, fontFamily: "'Space Grotesk', sans-serif", fontSize: 14 }}>
          {dayLabel(weekStart)} – {dayLabel(addDays(weekStart, 6))}
        </span>
        <Btn small variant="ghost" icon={ChevronRight} onClick={() => setWeekStart(addDays(weekStart, 7))} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", borderBottom: `2px solid ${C.cyan}` }}>
        {days.map((d) => (
          <div key={d} style={{ padding: "8px 10px", fontSize: 12, fontWeight: 700, color: d === todayISO() ? C.cyan : C.navy, borderRight: `1px solid ${C.line}`, background: "#F8FAFC" }}>
            {dayLabel(d)}
          </div>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", minHeight: 320 }}>
        {days.map((d) => {
          const dayBookings = bookings.filter((b) => b.date === d && b.status !== "cancelled" && b.status !== "declined");
          const { placed, cols } = layoutDay(dayBookings);
          return (
            <div key={d} style={{ position: "relative", borderRight: `1px solid ${C.line}`, background: "repeating-linear-gradient(180deg, transparent, transparent 39px, #F1F4F8 40px)", minHeight: 320 }}>
              {dayBookings.length === 0 && <div style={{ padding: 10, fontSize: 11, color: "#B8C2CF" }}>—</div>}
              {placed.map(({ b, col }) => {
                const top = Math.max(0, ((timeToMin(b.startTime) - DAY_START) / totalMin) * 320);
                const height = Math.max(22, ((timeToMin(b.endTime) - timeToMin(b.startTime)) / totalMin) * 320);
                const m = statusMeta(b.status);
                const widthPct = 100 / cols;
                const spaces = Number(b.spaces) || 1;
                return (
                  <div key={b.id} title={`${b.hirerName} · ${facilityById[b.facilityId]?.name}${spaces > 1 ? ` · ${spaces} spaces` : ""}`}
                    style={{
                      position: "absolute", top, height, left: `calc(${col * widthPct}% + 2px)`, width: `calc(${widthPct}% - 4px)`,
                      background: m.bg, borderLeft: `3px solid ${m.fg}`,
                      borderRadius: 5, padding: "3px 5px", fontSize: 10, overflow: "hidden",
                    }}>
                    <div style={{ fontWeight: 700, color: m.fg }}>{b.startTime}{spaces > 1 ? ` ×${spaces}` : ""}</div>
                    {showFacilityName && <div style={{ color: C.ink, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{facilityById[b.facilityId]?.name}</div>}
                    <div style={{ color: C.mute, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{b.hirerName}</div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
      <div style={{ padding: "8px 16px", fontSize: 11.5, color: C.mute, borderTop: `1px solid ${C.line}` }}>Cancelled bookings are hidden here so the freed-up slot is easy to rebook. Times shown 07:00–22:00. Side-by-side blocks are concurrent bookings sharing the facility's spaces.</div>
    </div>
  );
}

// ---------- multi-date picker (click days across any number of months to build an arbitrary date set) ----------
const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
function ymd(year, month, day) { return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`; }
function daysInMonth(year, month) { return new Date(Date.UTC(year, month, 0)).getUTCDate(); } // month is 1-indexed
function firstWeekdayOfMonth(year, month) { // Monday=0..Sunday=6
  const wd = new Date(Date.UTC(year, month - 1, 1)).getUTCDay();
  return wd === 0 ? 6 : wd - 1;
}

function MultiDatePicker({ selected, onChange }) {
  const today = todayISO();
  const [viewYear, setViewYear] = useState(Number(today.slice(0, 4)));
  const [viewMonth, setViewMonth] = useState(Number(today.slice(5, 7))); // 1-indexed

  function changeMonth(delta) {
    let m = viewMonth + delta, y = viewYear;
    if (m > 12) { m = 1; y += 1; }
    if (m < 1) { m = 12; y -= 1; }
    setViewMonth(m); setViewYear(y);
  }
  function toggleDate(iso) {
    onChange(selected.includes(iso) ? selected.filter((d) => d !== iso) : [...selected, iso].sort());
  }

  const offset = firstWeekdayOfMonth(viewYear, viewMonth);
  const total = daysInMonth(viewYear, viewMonth);
  const cells = [...Array(offset).fill(null), ...Array.from({ length: total }, (_, i) => i + 1)];

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
        <Btn small variant="ghost" icon={ChevronLeft} onClick={() => changeMonth(-1)} />
        <span style={{ fontWeight: 700, fontSize: 13.5, color: C.navy, fontFamily: "'Space Grotesk', sans-serif" }}>{MONTH_NAMES[viewMonth - 1]} {viewYear}</span>
        <Btn small variant="ghost" icon={ChevronRight} onClick={() => changeMonth(1)} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3, marginBottom: 4 }}>
        {["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"].map((d) => (
          <div key={d} style={{ textAlign: "center", fontSize: 10.5, fontWeight: 700, color: C.mute, padding: "2px 0" }}>{d}</div>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3 }}>
        {cells.map((day, i) => {
          if (day === null) return <div key={`b${i}`} />;
          const iso = ymd(viewYear, viewMonth, day);
          const isSel = selected.includes(iso);
          const isPast = iso < today;
          return (
            <button key={iso} type="button" disabled={isPast} onClick={() => toggleDate(iso)} style={{
              padding: "7px 0", borderRadius: 7, border: "none", cursor: isPast ? "default" : "pointer",
              fontSize: 12.5, fontWeight: isSel ? 700 : 500,
              background: isSel ? C.cyan : isPast ? "transparent" : "#F1F4F8",
              color: isSel ? C.navy : isPast ? "#CBD5E1" : C.ink,
            }}>{day}</button>
          );
        })}
      </div>
      {selected.length > 0 && (
        <div style={{ marginTop: 12, display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
          {selected.map((iso) => (
            <span key={iso} style={{ display: "inline-flex", alignItems: "center", gap: 5, background: C.cyanSoft, color: C.navy, fontSize: 11.5, fontWeight: 600, padding: "3px 8px", borderRadius: 999 }}>
              {dayLabel(iso)}
              <button type="button" onClick={() => toggleDate(iso)} style={{ background: "none", border: "none", cursor: "pointer", color: C.navy, padding: 0, display: "flex" }}><X size={11} /></button>
            </span>
          ))}
          <button type="button" onClick={() => onChange([])} style={{ background: "none", border: "none", color: C.mute, fontSize: 11.5, fontWeight: 600, cursor: "pointer", padding: 0 }}>Clear all</button>
        </div>
      )}
    </div>
  );
}

function BookingModal({ initial, prefill, facilities, siteById, members, setMembers, goMembers, allBookings, blackouts, onClose, onSave }) {
  const [form, setForm] = useState(initial || {
    facilityId: facilities[0]?.id || "", memberId: "", date: todayISO(),
    startTime: "17:00", endTime: "18:00", purpose: "", price: "", status: "confirmed", notes: "",
    repeatMode: "none", repeatUntil: "", multiDates: [], override: false, spaces: 1, customValues: {},
    emailConfirmation: true, extraFacilityIds: [],
    ...(prefill || {}),
  });
  const [quickAddOpen, setQuickAddOpen] = useState(false);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const setBool = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.checked }));
  const setCustom = (fieldId) => (e) => setForm((f) => ({ ...f, customValues: { ...f.customValues, [fieldId]: e.target.value } }));

  const selectedMember = members.find((m) => m.id === form.memberId);
  // A member with assigned sites can only book facilities at those sites —
  // narrow the picker rather than let the wrong one be selected at all.
  const allowedFacilities = selectedMember?.siteIds?.length
    ? facilities.filter((f) => selectedMember.siteIds.includes(f.siteId))
    : facilities;
  const selectedFacility = facilities.find((f) => f.id === form.facilityId);
  const capacity = selectedFacility?.capacity || 1;
  const requestedSpaces = Math.min(Number(form.spaces) || 1, capacity);
  // Only facilities at the same site as the primary one. Combinable with a one-off
  // booking or a set of picked dates (one group per date); weekly-repeat is excluded
  // — a multi-facility series running for months gets complicated fast.
  const extraFacilityOptions = !initial && (form.repeatMode === "none" || form.repeatMode === "dates") && selectedFacility
    ? allowedFacilities.filter((f) => f.siteId === selectedFacility.siteId && f.id !== form.facilityId)
    : [];

  function onFacilityChange(e) {
    const fac = facilities.find((f) => f.id === e.target.value);
    setForm((f) => ({
      ...f, facilityId: e.target.value, spaces: Math.min(Number(f.spaces) || 1, fac?.capacity || 1),
      // extras only make sense at the same site as the (new) primary facility
      extraFacilityIds: f.extraFacilityIds.filter((id) => facilities.find((x) => x.id === id)?.siteId === fac?.siteId),
    }));
  }
  function toggleExtraFacility(id) {
    setForm((f) => ({
      ...f,
      extraFacilityIds: f.extraFacilityIds.includes(id) ? f.extraFacilityIds.filter((x) => x !== id) : [...f.extraFacilityIds, id],
    }));
  }
  function onMemberChange(e) {
    const member = members.find((m) => m.id === e.target.value);
    setForm((f) => {
      const stillAllowed = !member?.siteIds?.length || facilities.find((fac) => fac.id === f.facilityId && member.siteIds.includes(fac.siteId));
      if (stillAllowed) return { ...f, memberId: e.target.value };
      const nextFacility = facilities.find((fac) => member.siteIds.includes(fac.siteId));
      return { ...f, memberId: e.target.value, facilityId: nextFacility?.id || "" };
    });
  }

  const anchorDate = form.repeatMode === "dates" ? [...form.multiDates].sort()[0] : form.date;

  const clashes = form.facilityId && anchorDate && form.startTime && form.endTime
    ? findClashes(allBookings, { facilityId: form.facilityId, date: anchorDate, startTime: form.startTime, endTime: form.endTime, excludeId: initial?.id })
    : [];
  const used = spacesUsed(clashes);
  const spacesLeft = capacity - used;
  const capacityIssue = requestedSpaces > spacesLeft;
  const noticeIssue = !initial && noticeViolation(selectedFacility, anchorDate, form.startTime);
  const advanceIssue = !initial && advanceViolation(selectedFacility, anchorDate);
  const blackoutRule = form.facilityId && anchorDate && form.startTime && form.endTime && blackouts
    ? findBlackout(blackouts, { facilityId: form.facilityId, siteId: selectedFacility?.siteId, date: anchorDate, startTime: form.startTime, endTime: form.endTime })
    : null;

  // Same checks, run again for each additional facility being booked
  // alongside the primary one — a clash on any of them should block saving
  // just as much as a clash on the primary facility would.
  const extraIssues = form.extraFacilityIds.map((fid) => {
    const fac = facilities.find((f) => f.id === fid);
    const facClashes = anchorDate && form.startTime && form.endTime
      ? findClashes(allBookings, { facilityId: fid, date: anchorDate, startTime: form.startTime, endTime: form.endTime, excludeId: initial?.id })
      : [];
    const facCapacity = fac?.capacity || 1;
    const facCapacityIssue = spacesUsed(facClashes) + 1 > facCapacity;
    const facBlackout = anchorDate && form.startTime && form.endTime && blackouts
      ? findBlackout(blackouts, { facilityId: fid, siteId: fac?.siteId, date: anchorDate, startTime: form.startTime, endTime: form.endTime })
      : null;
    const facNoticeIssue = noticeViolation(fac, anchorDate, form.startTime);
    const facAdvanceIssue = advanceViolation(fac, anchorDate);
    return { facility: fac, clashes: facClashes, capacityIssue: facCapacityIssue, blackoutRule: facBlackout, noticeIssue: facNoticeIssue, advanceIssue: facAdvanceIssue };
  }).filter((x) => x.clashes.length > 0 || x.capacityIssue || x.blackoutRule || x.noticeIssue || x.advanceIssue);

  const hasIssue = capacityIssue || noticeIssue || advanceIssue || !!blackoutRule || extraIssues.length > 0;

  const datesValid = form.repeatMode === "dates" ? form.multiDates.length > 0
    : form.repeatMode === "weekly" ? !!(form.date && form.repeatUntil && form.repeatUntil >= form.date)
    : !!form.date;

  const valid = form.facilityId && form.memberId && datesValid
    && form.startTime && form.endTime && timeToMin(form.endTime) > timeToMin(form.startTime)
    && (!hasIssue || form.override);

  function handleQuickAdd(newMember) {
    const withId = { ...newMember, id: uid(), siteIds: [] };
    setMembers((ms) => [...ms, withId]);
    setForm((f) => ({ ...f, memberId: withId.id }));
    setQuickAddOpen(false);
  }

  function handleSave() {
    onSave({
      ...form,
      id: initial?.id,
      price: form.price === "" ? 0 : form.price,
      spaces: requestedSpaces,
      hirerName: selectedMember?.name || "",
      hirerContact: selectedMember?.email || selectedMember?.phone || "",
      company: selectedMember?.company || "",
      emailConfirmation: !initial && form.emailConfirmation && !!selectedMember?.email,
      // extra facilities apply to a one-off booking or a picked-dates series, but
      // never to weekly-repeat — don't let a stale selection from before switching
      // repeat mode leak into a weekly save
      facilityIds: form.repeatMode === "weekly" ? [form.facilityId] : [form.facilityId, ...form.extraFacilityIds],
    });
  }

  return (
    <Modal title={initial ? "Edit booking" : "New booking"} onClose={onClose} wide>
      <Field label="Member / client">
        {members.length === 0 ? (
          <div style={{ fontSize: 13, color: C.mute, marginBottom: 8 }}>
            No members yet. <button type="button" onClick={() => setQuickAddOpen(true)} style={{ color: C.cyan, background: "none", border: "none", cursor: "pointer", fontWeight: 700, padding: 0 }}>Add one now</button>
          </div>
        ) : (
          <div style={{ display: "flex", gap: 8 }}>
            <select style={inputStyle} value={form.memberId} onChange={onMemberChange}>
              <option value="">Select a member…</option>
              {members.map((m) => <option key={m.id} value={m.id}>{m.name}{m.company ? ` — ${m.company}` : ""}</option>)}
            </select>
            <Btn small variant="ghost" icon={Plus} onClick={() => setQuickAddOpen(true)}>New</Btn>
          </div>
        )}
        {selectedMember && (
          <div style={{ fontSize: 12, color: C.mute, marginTop: 6 }}>
            {[selectedMember.email, selectedMember.phone].filter(Boolean).join(" · ") || "No email or phone on file"}
            {selectedMember.siteIds?.length > 0 && (
              <span> · restricted to {selectedMember.siteIds.map((id) => siteById[id]?.name).filter(Boolean).join(", ")}</span>
            )}
          </div>
        )}
      </Field>

      {quickAddOpen && (
        <div style={{ border: `1px dashed ${C.cyan}`, borderRadius: 10, padding: 14, marginBottom: 14, background: C.cyanSoft }}>
          <QuickAddMember onCancel={() => setQuickAddOpen(false)} onAdd={handleQuickAdd} />
        </div>
      )}

      <div className="dp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
        <Field label="Facility">
          <select style={inputStyle} value={form.facilityId} onChange={onFacilityChange} disabled={allowedFacilities.length === 0}>
            {allowedFacilities.length === 0 && <option value="">No facilities available for this member</option>}
            {allowedFacilities.map((f) => <option key={f.id} value={f.id}>{f.name} — {siteById[f.siteId]?.name}{f.capacity > 1 ? ` (${f.capacity} spaces)` : ""}</option>)}
          </select>
        </Field>
        <Field label="Status">
          <select style={inputStyle} value={form.status} onChange={set("status")}>
            {Object.entries(STATUS_META).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
          </select>
        </Field>
        {form.repeatMode !== "dates" && (
          <Field label="Date">
            <input type="date" style={inputStyle} value={form.date} onChange={set("date")} />
          </Field>
        )}
        <Field label="Purpose">
          <input style={inputStyle} value={form.purpose} onChange={set("purpose")} placeholder="5-a-side league, badminton club…" />
        </Field>
        <Field label="Start time">
          <select style={inputStyle} value={form.startTime} onChange={set("startTime")}>
            {TIME_SLOTS.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="End time">
          <select style={inputStyle} value={form.endTime} onChange={set("endTime")}>
            {TIME_SLOTS.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="Price charged (£)">
          <input type="number" min="0" step="0.01" style={inputStyle} value={form.price} onChange={set("price")} placeholder="0.00" />
        </Field>
        {capacity > 1 && (
          <Field label="Spaces needed">
            <select style={inputStyle} value={requestedSpaces} onChange={set("spaces")}>
              {Array.from({ length: capacity }, (_, i) => i + 1).map((n) => <option key={n} value={n}>{n} of {capacity}{n === capacity ? " (whole facility)" : ""}</option>)}
            </select>
          </Field>
        )}
      </div>

      {extraFacilityOptions.length > 0 && (
        <div style={{ background: "#F8FAFC", border: `1px solid ${C.line}`, borderRadius: 10, padding: 12, marginBottom: 14 }}>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: C.mute, textTransform: "uppercase", marginBottom: 8 }}>Also book at the same time (same member &amp; time{form.repeatMode === "dates" ? ", on each date picked below" : ""})</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {extraFacilityOptions.map((f) => (
              <label key={f.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: C.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={form.extraFacilityIds.includes(f.id)} onChange={() => toggleExtraFacility(f.id)} />
                {f.name}
              </label>
            ))}
          </div>
          {form.extraFacilityIds.length > 0 && (
            <div style={{ fontSize: 12, color: C.mute, marginTop: 8 }}>
              The price above covers all {form.extraFacilityIds.length + 1} facilities together{form.repeatMode === "dates" ? ", per date" : ""} — it'll show as one line on the invoice ({[selectedFacility?.name, ...form.extraFacilityIds.map((id) => facilities.find((f) => f.id === id)?.name)].filter(Boolean).join(" + ")}), and each facility still shows the booking on its own calendar.
            </div>
          )}
        </div>
      )}

      {selectedFacility?.customFields?.length > 0 && (
        <div style={{ background: "#F8FAFC", border: `1px solid ${C.line}`, borderRadius: 10, padding: 12, marginBottom: 14 }}>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: C.mute, textTransform: "uppercase", marginBottom: 8 }}>Additional info for {selectedFacility.name}</div>
          <div className="dp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
            {selectedFacility.customFields.map((cf) => (
              <Field key={cf.id} label={cf.label}>
                {cf.type === "dropdown" ? (
                  <select style={inputStyle} value={form.customValues?.[cf.id] || ""} onChange={setCustom(cf.id)}>
                    <option value="">Select…</option>
                    {cf.options.split(",").map((o) => o.trim()).filter(Boolean).map((o) => <option key={o} value={o}>{o}</option>)}
                  </select>
                ) : (
                  <input style={inputStyle} value={form.customValues?.[cf.id] || ""} onChange={setCustom(cf.id)} />
                )}
              </Field>
            ))}
          </div>
        </div>
      )}

      {(clashes.length > 0 || noticeIssue || advanceIssue || blackoutRule || extraIssues.length > 0) && (
        <div style={{
          border: `1px solid ${hasIssue ? C.coral : C.amber}`, background: hasIssue ? C.coralSoft : C.amberSoft,
          borderRadius: 10, padding: 12, marginBottom: 14,
        }}>
          {clashes.length > 0 && (
            <div style={{ marginBottom: (noticeIssue || advanceIssue || blackoutRule || extraIssues.length > 0) ? 8 : 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 7, fontWeight: 700, fontSize: 13.5, color: capacityIssue ? C.coral : C.amber }}>
                <AlertCircle size={15} />
                {capacityIssue ? `Not enough space — ${used}/${capacity} taken, you need ${requestedSpaces}` : `${spacesLeft} of ${capacity} space${capacity === 1 ? "" : "s"} free for this slot`}
              </div>
              <div style={{ fontSize: 12.5, color: C.ink, marginTop: 4 }}>
                Already booked: {clashes.map((c) => `${c.hirerName}${(Number(c.spaces) || 1) > 1 ? ` (${c.spaces})` : ""}`).join(", ")}
              </div>
            </div>
          )}
          {blackoutRule && (
            <div style={{ display: "flex", alignItems: "center", gap: 7, fontWeight: 700, fontSize: 13.5, color: C.coral, marginBottom: (noticeIssue || advanceIssue || extraIssues.length > 0) ? 6 : 0 }}>
              <Lock size={15} /> Blocked by admin rule: "{blackoutRule.label}"
            </div>
          )}
          {noticeIssue && (
            <div style={{ display: "flex", alignItems: "center", gap: 7, fontWeight: 700, fontSize: 13.5, color: C.coral, marginBottom: (advanceIssue || extraIssues.length > 0) ? 6 : 0 }}>
              <AlertCircle size={15} /> Less than {selectedFacility.minNoticeHours} hours' notice for this facility
            </div>
          )}
          {advanceIssue && (
            <div style={{ display: "flex", alignItems: "center", gap: 7, fontWeight: 700, fontSize: 13.5, color: C.coral, marginBottom: extraIssues.length > 0 ? 6 : 0 }}>
              <AlertCircle size={15} /> More than {selectedFacility.maxAdvanceDays} days ahead — beyond this facility's booking window
            </div>
          )}
          {extraIssues.map((issue, i) => (
            <div key={issue.facility?.id || i} style={{ fontSize: 12.5, color: C.coral, marginBottom: i < extraIssues.length - 1 ? 4 : 0 }}>
              <b>{issue.facility?.name}:</b>{" "}
              {issue.blackoutRule ? `blocked by "${issue.blackoutRule.label}"` :
                issue.capacityIssue ? "no space free for this slot" :
                issue.noticeIssue ? "less than the required notice" :
                issue.advanceIssue ? "beyond the booking window" :
                "already booked in this slot"}
            </div>
          ))}
          {hasIssue && (
            <label style={{ display: "flex", alignItems: "center", gap: 7, marginTop: 8, fontSize: 12.5, fontWeight: 600, color: C.ink, cursor: "pointer" }}>
              <input type="checkbox" checked={form.override} onChange={setBool("override")} />
              Proceed anyway, overriding this rule
            </label>
          )}
        </div>
      )}

      {!initial && (
        <div style={{ background: "#F8FAFC", border: `1px solid ${C.line}`, borderRadius: 10, padding: 12, marginBottom: 14 }}>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: C.mute, textTransform: "uppercase", marginBottom: 8 }}>Repeat this booking</div>
          <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
            {[
              ["none", "Just this one"],
              ["weekly", "Same day & time, weekly"],
              ["dates", "Pick specific dates"],
            ].map(([v, label]) => {
              const lockedOut = v === "weekly" && form.extraFacilityIds.length > 0;
              return (
                <button key={v} type="button" disabled={lockedOut} onClick={() => setForm((f) => ({ ...f, repeatMode: v }))} style={{
                  padding: "7px 12px", borderRadius: 8, fontSize: 12.5, fontWeight: 700, cursor: lockedOut ? "not-allowed" : "pointer",
                  border: `1px solid ${form.repeatMode === v ? C.cyan : C.line}`,
                  background: form.repeatMode === v ? C.cyanSoft : "#fff", color: form.repeatMode === v ? C.navy : C.mute,
                  opacity: lockedOut ? 0.45 : 1,
                }}>{v === "weekly" && <Repeat size={12} style={{ marginRight: 5, verticalAlign: -2 }} />}{label}</button>
              );
            })}
          </div>
          {form.extraFacilityIds.length > 0 && (
            <div style={{ fontSize: 12, color: C.mute, marginTop: -8, marginBottom: 12 }}>
              Weekly repeat isn't available for a multi-facility booking — remove the extra facilities above first. Picking specific dates is fine.
            </div>
          )}

          {form.repeatMode === "weekly" && (
            <div>
              <Field label="Repeat until (inclusive)">
                <input type="date" style={inputStyle} min={form.date} value={form.repeatUntil} onChange={set("repeatUntil")} />
              </Field>
              <div style={{ fontSize: 12, color: C.mute, marginTop: -8 }}>
                Creates one booking every {form.date ? dayLabel(form.date).split(",")[0] : "week"} from {form.date ? dayLabel(form.date) : "the start date"} through the date above, always at {form.startTime}–{form.endTime}. Availability above is only checked for the first date — worth a quick scan of the calendar for the later weeks.
              </div>
            </div>
          )}

          {form.repeatMode === "dates" && (
            <div>
              <MultiDatePicker selected={form.multiDates} onChange={(dates) => setForm((f) => ({ ...f, multiDates: dates }))} />
              <div style={{ fontSize: 12, color: C.mute, marginTop: 10 }}>
                Creates one booking on each date selected above, all at {form.startTime}–{form.endTime}. Availability above is only checked against the earliest date picked.
              </div>
            </div>
          )}
        </div>
      )}

      {!initial && (
        <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, fontSize: 13, fontWeight: 600, color: selectedMember?.email ? C.ink : C.mute, cursor: selectedMember?.email ? "pointer" : "not-allowed" }}>
          <input type="checkbox" checked={form.emailConfirmation && !!selectedMember?.email} disabled={!selectedMember?.email} onChange={setBool("emailConfirmation")} />
          <Mail size={14} />
          {selectedMember?.email ? "Download a confirmation PDF for this member after saving" : "Confirmation PDF (no email on file for this member)"}
        </label>
      )}

      <Field label="Notes (internal)">
        <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={form.notes} onChange={set("notes")} placeholder="Anything staff should know…" />
      </Field>
      {!valid && <div style={{ fontSize: 12.5, color: C.coral, marginBottom: 10 }}>
        {hasIssue && !form.override ? "There's a rule conflict above — tick \"Proceed anyway\" to override, or adjust the booking." : form.repeatMode === "dates" ? "Pick at least one date, and check a member, facility, and time are set." : `Check a member is selected, facility is set, end time is after start time${form.repeatMode === "weekly" ? ", and a repeat-until date is chosen" : ""}.`}
      </div>}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 6 }}>
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="accent" disabled={!valid} onClick={handleSave}>
          {initial ? "Save changes" : form.repeatMode === "none" ? "Create booking" : "Create recurring bookings"}
        </Btn>
      </div>
    </Modal>
  );
}

function QuickAddMember({ onCancel, onAdd }) {
  const [m, setM] = useState({ name: "", email: "", phone: "", company: "" });
  const set = (k) => (e) => setM((f) => ({ ...f, [k]: e.target.value }));
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 10px" }}>
        <Field label="Name"><input style={inputStyle} value={m.name} onChange={set("name")} /></Field>
        <Field label="Company"><input style={inputStyle} value={m.company} onChange={set("company")} /></Field>
        <Field label="Email"><input type="email" style={inputStyle} value={m.email} onChange={set("email")} /></Field>
        <Field label="Phone"><input style={inputStyle} value={m.phone} onChange={set("phone")} /></Field>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn small variant="ghost" onClick={onCancel}>Cancel</Btn>
        <Btn small variant="accent" disabled={!m.name} onClick={() => onAdd(m)}>Add member</Btn>
      </div>
    </div>
  );
}

// ---------- members ----------
function exportMembersXLSX(members, bookings) {
  const rows = members.map((m) => {
    const mine = bookings.filter((b) => b.memberId === m.id && b.status !== "declined" && b.status !== "cancelled");
    const totalSpend = mine.reduce((s, b) => s + (Number(b.price) || 0), 0);
    const totalHours = mine.reduce((s, b) => s + hoursBetween(b.startTime, b.endTime) * (Number(b.spaces) || 1), 0);
    return {
      Name: m.name,
      Company: m.company || "",
      Email: m.email || "",
      Phone: m.phone || "",
      "Bookings": mine.length,
      "Total hours": Math.round(totalHours * 10) / 10,
      "Total spend (£)": Math.round(totalSpend * 100) / 100,
    };
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{ wch: 24 }, { wch: 24 }, { wch: 28 }, { wch: 16 }, { wch: 10 }, { wch: 12 }, { wch: 15 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Members");
  XLSX.writeFile(wb, `delsport-members-${todayISO()}.xlsx`);
}

function Members({ members, setMembers, bookings, setBookings, facilityById, siteById, allSites }) {
  const [modalOpen, setModalOpen] = useState(false);
  const [editing, setEditing] = useState(null);
  const [detailMember, setDetailMember] = useState(null);
  const [importOpen, setImportOpen] = useState(false);
  const [q, setQ] = useState("");
  const [confirmDeleteMember, setConfirmDeleteMember] = useState(null);

  function save(data) {
    if (data.id) setMembers((ms) => ms.map((m) => (m.id === data.id ? data : m)));
    else setMembers((ms) => [...ms, { ...data, id: uid() }]);
    setModalOpen(false);
    setEditing(null);
  }
  function remove(id) {
    setMembers((ms) => ms.filter((m) => m.id !== id));
  }

  const filtered = members.filter((m) =>
    !q || [m.name, m.email, m.company, m.phone].some((v) => (v || "").toLowerCase().includes(q.toLowerCase()))
  );

  return (
    <div>
      <PageHeader
        title="Members"
        sub={`${members.length} member${members.length === 1 ? "" : "s"} on file`}
        action={
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="ghost" icon={Upload} onClick={() => setImportOpen(true)}>Import from CSV</Btn>
            <Btn variant="ghost" icon={FileDown} onClick={() => exportMembersXLSX(members, bookings)} disabled={members.length === 0}>Export to Excel</Btn>
            <Btn variant="accent" icon={Plus} onClick={() => { setEditing(null); setModalOpen(true); }}>New member</Btn>
          </div>
        }
      />
      {importOpen && (
        <ImportMembersModal
          existingMembers={members}
          sites={allSites}
          onClose={() => setImportOpen(false)}
          onImport={(newMembers) => { setMembers((ms) => [...ms, ...newMembers]); setImportOpen(false); }}
        />
      )}

      <input style={{ ...inputStyle, maxWidth: 320, marginBottom: 16 }} placeholder="Search by name, email, company…" value={q} onChange={(e) => setQ(e.target.value)} />

      {filtered.length === 0 ? (
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
          <EmptyState icon={Users} title={members.length === 0 ? "No members yet" : "No matches"} sub={members.length === 0 ? "Add clients here so you can attach them to bookings in one click." : "Try a different search."} action={members.length === 0 ? <Btn variant="accent" onClick={() => setModalOpen(true)}>New member</Btn> : null} />
        </div>
      ) : (
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                {["Name", "Company", "Email", "Phone", "Allowed sites", "Bookings", ""].map((h) => (
                  <th key={h} style={{ padding: "10px 16px", fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.5, color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((m) => {
                const count = bookings.filter((b) => b.memberId === m.id).length;
                const siteNames = (m.siteIds || []).map((id) => siteById[id]?.name).filter(Boolean);
                return (
                  <tr key={m.id} onClick={() => setDetailMember(m)} style={{ borderBottom: `1px solid ${C.line}`, cursor: "pointer" }}>
                    <td style={{ padding: "11px 16px", fontWeight: 600, fontSize: 13.5, color: C.navy, textDecoration: "underline", textDecorationColor: C.line }}>{m.name}</td>
                    <td style={{ padding: "11px 16px", fontSize: 13, color: C.mute }}>{m.company || "—"}</td>
                    <td style={{ padding: "11px 16px", fontSize: 13 }}>{m.email || "—"}</td>
                    <td style={{ padding: "11px 16px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace" }}>{m.phone || "—"}</td>
                    <td style={{ padding: "11px 16px", fontSize: 12.5, color: siteNames.length ? C.ink : C.mute }}>{siteNames.length ? siteNames.join(", ") : "All sites"}</td>
                    <td style={{ padding: "11px 16px", fontSize: 13 }}>{count}</td>
                    <td style={{ padding: "11px 16px" }} onClick={(e) => e.stopPropagation()}>
                      <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                        <Btn small variant="ghost" icon={Pencil} onClick={() => { setEditing(m); setModalOpen(true); }} />
                        <Btn small variant="ghost" icon={Trash2} onClick={() => setConfirmDeleteMember(m)} />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {modalOpen && <MemberModal initial={editing} sites={allSites} onClose={() => { setModalOpen(false); setEditing(null); }} onSave={save} />}
      {detailMember && <MemberDetailModal member={detailMember} bookings={bookings} setBookings={setBookings} facilityById={facilityById} siteById={siteById} onClose={() => setDetailMember(null)} />}
      {confirmDeleteMember && (
        <ConfirmDeleteModal
          title="Delete this member?"
          message={(() => {
            const count = bookings.filter((b) => b.memberId === confirmDeleteMember.id).length;
            return `${confirmDeleteMember.name} will be permanently removed${count > 0 ? `. They have ${count} booking${count === 1 ? "" : "s"} on record, which will keep showing but no longer be linked to a member` : ""}. This can't be undone.`;
          })()}
          onCancel={() => setConfirmDeleteMember(null)}
          onConfirm={() => { remove(confirmDeleteMember.id); setConfirmDeleteMember(null); }}
        />
      )}
    </div>
  );
}

function MemberDetailModal({ member, bookings, setBookings, facilityById, siteById, onClose }) {
  const [selected, setSelected] = useState(() => new Set());
  const mine = bookings.filter((b) => b.memberId === member.id).sort((a, b) => b.date.localeCompare(a.date));
  const active = mine.filter((b) => b.status !== "declined" && b.status !== "cancelled");
  const cancellable = mine.filter((b) => b.status === "confirmed");

  function toggleOne(id) {
    setSelected((s) => { const next = new Set(s); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  }
  function toggleAll() {
    setSelected((s) => (s.size === cancellable.length ? new Set() : new Set(cancellable.map((b) => b.id))));
  }
  function cancelSelected() {
    setBookings((bs) => bs.map((b) => (selected.has(b.id) ? { ...b, status: "cancelled" } : b)));
    setSelected(new Set());
  }
  const totalRevenue = active.reduce((s, b) => s + (Number(b.price) || 0), 0);
  const totalHours = active.reduce((s, b) => s + hoursBetween(b.startTime, b.endTime), 0);

  const byFacility = {};
  active.forEach((b) => {
    const name = facilityById[b.facilityId]?.name || "Unknown";
    if (!byFacility[name]) byFacility[name] = { count: 0, hours: 0, revenue: 0 };
    byFacility[name].count += 1;
    byFacility[name].hours += hoursBetween(b.startTime, b.endTime);
    byFacility[name].revenue += Number(b.price) || 0;
  });

  return (
    <Modal title={member.name} onClose={onClose} wide>
      <div style={{ fontSize: 13, color: C.mute, marginBottom: 16, display: "flex", gap: 14, flexWrap: "wrap" }}>
        {member.company && <span>{member.company}</span>}
        {member.email && <span>{member.email}</span>}
        {member.phone && <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{member.phone}</span>}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginBottom: 18 }}>
        <div style={{ background: C.cyanSoft, borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 11.5, color: C.mute, fontWeight: 700, textTransform: "uppercase" }}>Total spend</div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 20, fontWeight: 700, color: C.navy }}>{money(totalRevenue)}</div>
        </div>
        <div style={{ background: C.pitchSoft, borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 11.5, color: C.mute, fontWeight: 700, textTransform: "uppercase" }}>Total hours</div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 20, fontWeight: 700, color: C.navy }}>{totalHours.toFixed(1)}</div>
        </div>
        <div style={{ background: "#F0EEFB", borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 11.5, color: C.mute, fontWeight: 700, textTransform: "uppercase" }}>Active bookings</div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 20, fontWeight: 700, color: C.navy }}>{active.length}</div>
        </div>
      </div>

      {Object.keys(byFacility).length > 0 && (
        <div style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, fontSize: 14, marginBottom: 8 }}>By facility</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {Object.entries(byFacility).map(([name, d]) => (
              <div key={name} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "7px 0", borderBottom: `1px solid ${C.line}` }}>
                <span style={{ fontWeight: 600, color: C.ink }}>{name}</span>
                <span style={{ color: C.mute }}>{d.count} booking{d.count === 1 ? "" : "s"} · {d.hours.toFixed(1)}h · <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.pitch, fontWeight: 600 }}>{money(d.revenue)}</span></span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, fontSize: 14 }}>All bookings</div>
        {selected.size > 0 && (
          <Btn small variant="danger" icon={Ban} onClick={cancelSelected}>Cancel {selected.size} selected</Btn>
        )}
      </div>
      {mine.length === 0 ? (
        <div style={{ fontSize: 13, color: C.mute }}>No bookings for this member yet.</div>
      ) : (
        <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                <th style={{ padding: "8px 12px", width: 1 }}>
                  {cancellable.length > 0 && (
                    <input type="checkbox" checked={selected.size > 0 && selected.size === cancellable.length} onChange={toggleAll} />
                  )}
                </th>
                {["Date", "Facility", "Time", "Hours", "Price", "Status"].map((h) => (
                  <th key={h} style={{ padding: "8px 12px", fontSize: 11, textTransform: "uppercase", color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {mine.map((b) => (
                <tr key={b.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                  <td style={{ padding: "8px 12px" }}>
                    {b.status === "confirmed" && (
                      <input type="checkbox" checked={selected.has(b.id)} onChange={() => toggleOne(b.id)} />
                    )}
                  </td>
                  <td style={{ padding: "8px 12px", fontSize: 12.5 }}>{dayLabel(b.date)}</td>
                  <td style={{ padding: "8px 12px", fontSize: 12.5 }}>{facilityById[b.facilityId]?.name}<div style={{ fontSize: 11, color: C.mute }}>{siteById[facilityById[b.facilityId]?.siteId]?.name}</div></td>
                  <td style={{ padding: "8px 12px", fontSize: 12.5, fontFamily: "'IBM Plex Mono', monospace" }}>{b.startTime}–{b.endTime}</td>
                  <td style={{ padding: "8px 12px", fontSize: 12.5 }}>{hoursBetween(b.startTime, b.endTime).toFixed(1)}</td>
                  <td style={{ padding: "8px 12px", fontSize: 12.5, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{money(b.price)}</td>
                  <td style={{ padding: "8px 12px" }}><StatusPill status={b.status} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Modal>
  );
}

function MemberModal({ initial, sites, onClose, onSave }) {
  const [form, setForm] = useState(initial || { name: "", email: "", phone: "", company: "", siteIds: [] });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  function toggleSite(id) {
    setForm((f) => ({ ...f, siteIds: (f.siteIds || []).includes(id) ? f.siteIds.filter((x) => x !== id) : [...(f.siteIds || []), id] }));
  }
  return (
    <Modal title={initial ? "Edit member" : "New member"} onClose={onClose}>
      <Field label="Full name"><input style={inputStyle} value={form.name} onChange={set("name")} placeholder="e.g. Jamie Carter" /></Field>
      <Field label="Business / company name"><input style={inputStyle} value={form.company} onChange={set("company")} placeholder="e.g. Riverside FC" /></Field>
      <Field label="Email address"><input type="email" style={inputStyle} value={form.email} onChange={set("email")} placeholder="name@email.com" /></Field>
      <Field label="Phone number"><input style={inputStyle} value={form.phone} onChange={set("phone")} placeholder="07…" /></Field>
      <Field label="Allowed sites (leave blank for no restriction)">
        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
          <button type="button" onClick={() => setForm((f) => ({ ...f, siteIds: [] }))} style={{ background: "none", border: "none", color: C.mute, fontWeight: 700, fontSize: 12, cursor: "pointer", padding: 0 }}>Clear (no restriction)</button>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 160, overflowY: "auto", border: `1px solid ${C.line}`, borderRadius: 8, padding: 10 }}>
          {sites.length === 0 && <div style={{ fontSize: 12.5, color: C.mute }}>No sites yet.</div>}
          {sites.map((s) => (
            <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, cursor: "pointer" }}>
              <input type="checkbox" checked={(form.siteIds || []).includes(s.id)} onChange={() => toggleSite(s.id)} />
              {s.name}
            </label>
          ))}
        </div>
      </Field>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="accent" disabled={!form.name} onClick={() => onSave({ ...form, id: initial?.id })}>{initial ? "Save changes" : "Add member"}</Btn>
      </div>
    </Modal>
  );
}

function ImportMembersModal({ existingMembers, sites, onClose, onImport }) {
  const [rows, setRows] = useState(null);
  const [error, setError] = useState("");

  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    const reader = new FileReader();
    reader.onload = () => {
      const parsed = csvRowsToObjects(parseCSVText(String(reader.result)));
      if (parsed.length === 0) { setError("Couldn't find any rows in that file."); setRows(null); return; }
      setRows(parsed);
    };
    reader.readAsText(file);
  }

  const preview = useMemo(() => {
    if (!rows) return null;
    const existingKey = (name, email) => `${normKey(name)}::${normKey(email)}`;
    const existingSet = new Set(existingMembers.map((m) => existingKey(m.name, m.email)));
    const toImport = [];
    const skipped = [];
    rows.forEach((row, i) => {
      const rowNum = i + 2;
      const name = pickField(row, ["name", "fullname", "membername", "clientname", "hirername"]);
      if (!name) { skipped.push({ row: rowNum, reason: "no name" }); return; }
      const email = pickField(row, ["email", "emailaddress"]);
      const key = existingKey(name, email);
      if (existingSet.has(key)) { skipped.push({ row: rowNum, reason: `"${name}" already on file` }); return; }
      const sitesRaw = pickField(row, ["sites", "allowedsites", "site"]);
      const resolvedSites = resolveSiteIdsByNames(sites, sitesRaw);
      if (resolvedSites.error) { skipped.push({ row: rowNum, reason: resolvedSites.error }); return; }
      existingSet.add(key); // also catch duplicate rows within the file itself
      toImport.push({
        id: uid(),
        name,
        email,
        company: pickField(row, ["company", "business", "organisation", "organization"]),
        phone: pickField(row, ["phone", "phonenumber", "mobile", "tel", "telephone"]),
        siteIds: resolvedSites.siteIds,
      });
    });
    return { toImport, skipped };
  }, [rows, existingMembers, sites]);

  function downloadTemplate() {
    downloadCSV(
      "delsport-members-template.csv",
      ["Name", "Company", "Email", "Phone", "Sites"],
      [["Jamie Carter", "Riverside FC", "jamie@riversidefc.co.uk", "07700 900123", ""]]
    );
  }

  return (
    <Modal title="Import members from CSV" onClose={onClose} wide>
      <div style={{ fontSize: 13, color: C.mute, marginBottom: 14, lineHeight: 1.6 }}>
        Needs columns for name, company, email and phone — reasonable header variations are recognised automatically ("Full Name", "Client", etc). An optional Sites column (site names separated by commas) restricts a member to only booking at those sites — leave it blank for no restriction.{" "}
        <button type="button" onClick={downloadTemplate} style={{ background: "none", border: "none", color: C.cyan, fontWeight: 700, cursor: "pointer", padding: 0, fontSize: 13 }}>Download a template</button>.
      </div>
      <input type="file" accept=".csv,text/csv" onChange={handleFile} style={{ fontSize: 13, marginBottom: 12 }} />
      {error && <div style={{ fontSize: 12.5, color: C.coral, marginBottom: 10 }}>{error}</div>}

      {preview && (
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 13.5, fontWeight: 600, color: C.navy, marginBottom: 8 }}>
            {preview.toImport.length} member{preview.toImport.length === 1 ? "" : "s"} ready to import
            {preview.skipped.length > 0 ? `, ${preview.skipped.length} skipped` : ""}
          </div>
          {preview.toImport.length > 0 && (
            <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "hidden", maxHeight: 220, overflowY: "auto", marginBottom: preview.skipped.length ? 10 : 0 }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                    {["Name", "Company", "Email", "Phone", "Allowed sites"].map((h) => (
                      <th key={h} style={{ padding: "7px 12px", fontSize: 11, textTransform: "uppercase", color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.toImport.map((m) => (
                    <tr key={m.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                      <td style={{ padding: "6px 12px", fontSize: 12.5 }}>{m.name}</td>
                      <td style={{ padding: "6px 12px", fontSize: 12.5, color: C.mute }}>{m.company || "—"}</td>
                      <td style={{ padding: "6px 12px", fontSize: 12.5 }}>{m.email || "—"}</td>
                      <td style={{ padding: "6px 12px", fontSize: 12.5 }}>{m.phone || "—"}</td>
                      <td style={{ padding: "6px 12px", fontSize: 12.5 }}>{m.siteIds.length ? m.siteIds.map((id) => sites.find((s) => s.id === id)?.name).join(", ") : "All sites"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          {preview.skipped.length > 0 && (
            <div style={{ fontSize: 12, color: C.mute }}>Skipped: {preview.skipped.map((s) => `row ${s.row} (${s.reason})`).join(", ")}</div>
          )}
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="accent" disabled={!preview || preview.toImport.length === 0} onClick={() => onImport(preview.toImport)}>
          Import {preview?.toImport.length || ""} member{preview?.toImport.length === 1 ? "" : "s"}
        </Btn>
      </div>
    </Modal>
  );
}

// ---------- sites & facilities ----------
function SitesFacilities({ sites, setSites, facilities, setFacilities, bookings }) {
  const [siteModal, setSiteModal] = useState(null); // null closed, {} new, obj edit
  const [facModal, setFacModal] = useState(null);
  const [facSiteFor, setFacSiteFor] = useState(null);
  const [loadedMsg, setLoadedMsg] = useState("");
  const [confirmDeleteSite, setConfirmDeleteSite] = useState(null);
  const [confirmDeleteFac, setConfirmDeleteFac] = useState(null);

  function handleLoadStarter() {
    const { addedSites, addedFacilities } = loadStarterData(sites, setSites, facilities, setFacilities);
    setLoadedMsg(addedSites || addedFacilities
      ? `Added ${addedSites} site${addedSites === 1 ? "" : "s"} and ${addedFacilities} facilit${addedFacilities === 1 ? "y" : "ies"} from delsportuk.com.`
      : "Your sites and facilities already match the website — nothing new to add.");
    setTimeout(() => setLoadedMsg(""), 6000);
  }

  function saveSite(data) {
    if (data.id) setSites((s) => s.map((x) => (x.id === data.id ? data : x)));
    else setSites((s) => [...s, { ...data, id: uid() }]);
    setSiteModal(null);
  }
  function removeSite(id) {
    setSites((s) => s.filter((x) => x.id !== id));
    setFacilities((f) => f.filter((x) => x.siteId !== id));
  }
  function saveFac(data) {
    if (data.id) setFacilities((f) => f.map((x) => (x.id === data.id ? data : x)));
    else setFacilities((f) => [...f, { ...data, id: uid() }]);
    setFacModal(null);
  }
  function removeFac(id) {
    setFacilities((f) => f.filter((x) => x.id !== id));
  }

  return (
    <div>
      <PageHeader
        title="Sites & Facilities"
        sub={`${sites.length} sites · ${facilities.length} facilities`}
        action={
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="ghost" icon={Download} onClick={handleLoadStarter}>Load from delsportuk.com</Btn>
            <Btn variant="accent" icon={Plus} onClick={() => setSiteModal({})}>New site</Btn>
          </div>
        }
      />
      {loadedMsg && (
        <div style={{ background: C.pitchSoft, color: C.pitch, fontSize: 13, fontWeight: 600, padding: "10px 14px", borderRadius: 9, marginBottom: 16 }}>{loadedMsg}</div>
      )}

      {sites.length === 0 ? (
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
          <EmptyState icon={MapPin} title="Add your first school site" sub="Sites hold the facilities you let out — start here, or load your sites straight from delsportuk.com." action={
            <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
              <Btn variant="ghost" icon={Download} onClick={handleLoadStarter}>Load from delsportuk.com</Btn>
              <Btn variant="accent" onClick={() => setSiteModal({})}>New site</Btn>
            </div>
          } />
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {sites.map((site) => {
            const facs = facilities.filter((f) => f.siteId === site.id);
            return (
              <div key={site.id} style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
                <div style={{ padding: "14px 18px", display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `1px solid ${C.line}`, background: "#F8FAFC" }}>
                  <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                    {site.logoUrl && <img src={site.logoUrl} alt="" style={{ width: 40, height: 40, objectFit: "contain", borderRadius: 8, background: "#fff", border: `1px solid ${C.line}`, flexShrink: 0 }} onError={(e) => { e.target.style.display = "none"; }} />}
                    <div>
                      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, color: C.navy, fontSize: 16 }}>{site.name}</div>
                      <div style={{ fontSize: 12.5, color: C.mute, marginTop: 3, display: "flex", gap: 14 }}>
                        {site.address && <span><MapPin size={12} style={{ verticalAlign: -1 }} /> {site.address}</span>}
                        {site.contact && <span><Phone size={12} style={{ verticalAlign: -1 }} /> {site.contact}</span>}
                      </div>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <Btn small variant="ghost" icon={Plus} onClick={() => { setFacSiteFor(site.id); setFacModal({}); }}>Facility</Btn>
                    <Btn small variant="ghost" icon={Pencil} onClick={() => setSiteModal(site)} />
                    <Btn small variant="ghost" icon={Trash2} onClick={() => setConfirmDeleteSite(site)} />
                  </div>
                </div>
                {facs.length === 0 ? (
                  <div style={{ padding: 16, fontSize: 13, color: C.mute }}>No facilities added yet.</div>
                ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 12, padding: 16 }}>
                    {facs.map((f) => (
                      <div key={f.id} style={{ border: `1px solid ${C.line}`, borderRadius: 10, padding: "11px 13px" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                          <div>
                            <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink }}>{f.name}</div>
                            <div style={{ fontSize: 11.5, color: C.mute, marginTop: 2 }}>{f.type}</div>
                          </div>
                          <div style={{ display: "flex", gap: 4 }}>
                            <Btn small variant="ghost" icon={Pencil} onClick={() => { setFacSiteFor(site.id); setFacModal(f); }} />
                            <Btn small variant="ghost" icon={Trash2} onClick={() => setConfirmDeleteFac(f)} />
                          </div>
                        </div>
                        <div style={{ marginTop: 8, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: C.pitch, fontWeight: 600 }}>{money(f.rate)}/hr</span>
                          <span style={{ fontSize: 11, fontWeight: 700, color: C.mute, background: "#EEF1F4", padding: "2px 7px", borderRadius: 999 }}>{f.capacity && f.capacity > 1 ? `${f.capacity} spaces` : "1 space"}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {siteModal !== null && <SiteModal initial={siteModal.id ? siteModal : null} onClose={() => setSiteModal(null)} onSave={saveSite} />}
      {facModal !== null && <FacilityModal initial={facModal.id ? facModal : null} onClose={() => setFacModal(null)} onSave={(d) => saveFac({ ...d, siteId: facModal.id ? d.siteId : facSiteFor })} />}
      {confirmDeleteSite && (
        <ConfirmDeleteModal
          title="Delete this site?"
          message={(() => {
            const facCount = facilities.filter((f) => f.siteId === confirmDeleteSite.id).length;
            return `${confirmDeleteSite.name} will be permanently removed${facCount > 0 ? `, along with its ${facCount} facilit${facCount === 1 ? "y" : "ies"}` : ""}. This can't be undone.`;
          })()}
          onCancel={() => setConfirmDeleteSite(null)}
          onConfirm={() => { removeSite(confirmDeleteSite.id); setConfirmDeleteSite(null); }}
        />
      )}
      {confirmDeleteFac && (
        <ConfirmDeleteModal
          title="Delete this facility?"
          message={`${confirmDeleteFac.name} will be permanently removed. This can't be undone.`}
          onCancel={() => setConfirmDeleteFac(null)}
          onConfirm={() => { removeFac(confirmDeleteFac.id); setConfirmDeleteFac(null); }}
        />
      )}
    </div>
  );
}

function SiteModal({ initial, onClose, onSave }) {
  const [form, setForm] = useState(initial || {
    name: "", address: "", contact: "", logoUrl: "",
    bankAccountName: "", bankSortCode: "", bankAccountNumber: "", vatNumber: "",
  });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  return (
    <Modal title={initial ? "Edit site" : "New site"} onClose={onClose}>
      <Field label="School / site name"><input style={inputStyle} value={form.name} onChange={set("name")} placeholder="e.g. Oakfield Community School" /></Field>
      <Field label="Address"><input style={inputStyle} value={form.address} onChange={set("address")} placeholder="Street, town, postcode" /></Field>
      <Field label="Site contact"><input style={inputStyle} value={form.contact} onChange={set("contact")} placeholder="Name / phone / email" /></Field>
      <Field label="Logo URL (optional)">
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          {form.logoUrl && <img src={form.logoUrl} alt="" style={{ width: 36, height: 36, objectFit: "contain", borderRadius: 6, border: `1px solid ${C.line}`, flexShrink: 0 }} onError={(e) => { e.target.style.visibility = "hidden"; }} />}
          <input style={inputStyle} value={form.logoUrl} onChange={set("logoUrl")} placeholder="https://…" />
        </div>
      </Field>

      <div style={{ borderTop: `1px solid ${C.line}`, paddingTop: 14, marginTop: 4, marginBottom: 14 }}>
        <div style={{ fontSize: 11.5, fontWeight: 700, color: C.mute, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 10 }}>
          Payment details (for invoices — bookings here are paid directly to this site)
        </div>
        <Field label="Bank account name"><input style={inputStyle} value={form.bankAccountName} onChange={set("bankAccountName")} placeholder="e.g. Oakfield Community School" /></Field>
        <div className="dp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
          <Field label="Sort code"><input style={inputStyle} value={form.bankSortCode} onChange={set("bankSortCode")} placeholder="00-00-00" /></Field>
          <Field label="Account number"><input style={inputStyle} value={form.bankAccountNumber} onChange={set("bankAccountNumber")} placeholder="12345678" /></Field>
        </div>
        <Field label="VAT number (optional)"><input style={inputStyle} value={form.vatNumber} onChange={set("vatNumber")} placeholder="e.g. GB123456789" /></Field>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="accent" disabled={!form.name} onClick={() => onSave({ ...form, id: initial?.id })}>{initial ? "Save changes" : "Add site"}</Btn>
      </div>
    </Modal>
  );
}

function FacilityModal({ initial, onClose, onSave }) {
  const [form, setForm] = useState(initial || { name: "", type: FACILITY_TYPES[0], rate: "", capacity: 1, minNoticeHours: "", maxAdvanceDays: "" });
  const [customFields, setCustomFields] = useState(initial?.customFields || []);
  const [confirmDeleteField, setConfirmDeleteField] = useState(null);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  function addField() {
    setCustomFields((cf) => [...cf, { id: uid(), label: "", type: "text", options: "" }]);
  }
  function updateField(id, patch) {
    setCustomFields((cf) => cf.map((f) => (f.id === id ? { ...f, ...patch } : f)));
  }
  function removeField(id) {
    setCustomFields((cf) => cf.filter((f) => f.id !== id));
  }

  return (
    <Modal title={initial ? "Edit facility" : "New facility"} onClose={onClose} wide>
      <Field label="Facility name"><input style={inputStyle} value={form.name} onChange={set("name")} placeholder="e.g. Main Sports Hall" /></Field>
      <div className="dp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
        <Field label="Type">
          <select style={inputStyle} value={form.type} onChange={set("type")}>
            {FACILITY_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="Default rate (£/hour)"><input type="number" min="0" step="0.01" style={inputStyle} value={form.rate} onChange={set("rate")} placeholder="0.00" /></Field>
        <Field label="Available spaces per time slot">
          <select style={inputStyle} value={form.capacity ?? 1} onChange={set("capacity")}>
            {[1, 2, 3, 4].map((n) => <option key={n} value={n}>{n} {n === 1 ? "(single booking only)" : "(concurrent bookings allowed)"}</option>)}
          </select>
        </Field>
        <div />
        <Field label="Minimum notice (hours)">
          <input type="number" min="0" style={inputStyle} value={form.minNoticeHours} onChange={set("minNoticeHours")} placeholder="0 = no limit" />
        </Field>
        <Field label="Max advance booking (days)">
          <input type="number" min="0" style={inputStyle} value={form.maxAdvanceDays} onChange={set("maxAdvanceDays")} placeholder="0 = no limit" />
        </Field>
      </div>
      <div style={{ fontSize: 12, color: C.mute, marginTop: -8, marginBottom: 14 }}>Leave either blank or at 0 for no restriction. E.g. 24 hours notice, or bookable up to 90 days ahead.</div>

      <div style={{ borderTop: `1px solid ${C.line}`, paddingTop: 14, marginBottom: 10 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <span style={{ fontSize: 12.5, fontWeight: 700, color: C.mute, textTransform: "uppercase", letterSpacing: 0.4 }}>Extra info to capture at booking</span>
          <Btn small variant="ghost" icon={Plus} onClick={addField}>Add field</Btn>
        </div>
        {customFields.length === 0 ? (
          <div style={{ fontSize: 12.5, color: C.mute }}>None — bookings for this facility only ask the standard questions.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {customFields.map((cf) => (
              <div key={cf.id} style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <input style={{ ...inputStyle, flex: 2 }} placeholder="Field label, e.g. Insurance reference" value={cf.label} onChange={(e) => updateField(cf.id, { label: e.target.value })} />
                <select style={{ ...inputStyle, flex: 1, width: "auto" }} value={cf.type} onChange={(e) => updateField(cf.id, { type: e.target.value })}>
                  <option value="text">Text</option>
                  <option value="dropdown">Dropdown</option>
                </select>
                {cf.type === "dropdown" && (
                  <input style={{ ...inputStyle, flex: 2 }} placeholder="Options, comma separated" value={cf.options} onChange={(e) => updateField(cf.id, { options: e.target.value })} />
                )}
                <Btn small variant="ghost" icon={Trash2} onClick={() => setConfirmDeleteField(cf)} />
              </div>
            ))}
          </div>
        )}
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="accent" disabled={!form.name} onClick={() => onSave({
          ...form, id: initial?.id, siteId: initial?.siteId, capacity: Number(form.capacity) || 1,
          minNoticeHours: Number(form.minNoticeHours) || 0, maxAdvanceDays: Number(form.maxAdvanceDays) || 0,
          customFields: customFields.filter((cf) => cf.label.trim()),
        })}>{initial ? "Save changes" : "Add facility"}</Btn>
      </div>
      {confirmDeleteField && (
        <ConfirmDeleteModal
          title="Remove this field?"
          message={`"${confirmDeleteField.label || "This field"}" will be removed from the form.`}
          onCancel={() => setConfirmDeleteField(null)}
          onConfirm={() => { removeField(confirmDeleteField.id); setConfirmDeleteField(null); }}
        />
      )}
    </Modal>
  );
}

// ---------- reports ----------
const OPEN_HOURS_PER_DAY = 15; // 07:00–22:00, used as the availability baseline for utilisation

function startOfMonthISO(iso) { return iso.slice(0, 7) + "-01"; }
function endOfMonthISO(iso) {
  const [y, m] = iso.split("-").map(Number);
  return utcMsToISO(Date.UTC(y, m, 0)); // day 0 of next month = last day of this month
}
function startOfYearISO(iso) { return iso.slice(0, 4) + "-01-01"; }
function endOfYearISO(iso) { return iso.slice(0, 4) + "-12-31"; }
function daysBetween(a, b) { return Math.round((new Date(b + "T00:00:00") - new Date(a + "T00:00:00")) / 86400000) + 1; }

const RANGE_PRESETS = [
  { id: "this_month", label: "This month" },
  { id: "last_30", label: "Last 30 days" },
  { id: "this_year", label: "This year" },
  { id: "all_time", label: "All time" },
  { id: "custom", label: "Custom range" },
];
function getRangeDates(preset, customStart, customEnd) {
  const today = todayISO();
  switch (preset) {
    case "this_month": return [startOfMonthISO(today), endOfMonthISO(today)];
    case "last_30": return [addDays(today, -29), today];
    case "this_year": return [startOfYearISO(today), endOfYearISO(today)];
    case "all_time": return ["2000-01-01", "2100-12-31"];
    case "custom": return [customStart || today, customEnd || today];
    default: return [addDays(today, -29), today];
  }
}

function Reports({ bookings, sites, facilities, facilityById, siteById, members, memberById }) {
  const [preset, setPreset] = useState("last_30");
  const [customStart, setCustomStart] = useState(addDays(todayISO(), -29));
  const [customEnd, setCustomEnd] = useState(todayISO());
  const [utilSiteId, setUtilSiteId] = useState("all");

  // "All time" used a fixed 2000–2100 window, which made the utilisation % (and the header's
  // date range) meaningless — a facility with hundreds of confirmed hours still rounded to 0%
  // against a 100-year availability baseline. Bound it to the actual span of booking data instead.
  const allTimeBounds = useMemo(() => {
    if (!bookings.length) { const t = todayISO(); return [t, t]; }
    let min = bookings[0].date, max = bookings[0].date;
    bookings.forEach((b) => { if (b.date < min) min = b.date; if (b.date > max) max = b.date; });
    const today = todayISO();
    return [min, max > today ? max : today];
  }, [bookings]);
  const [rangeStart, rangeEnd] = preset === "all_time" ? allTimeBounds : getRangeDates(preset, customStart, customEnd);

  useEffect(() => {
    if (utilSiteId !== "all" && !sites.some((s) => s.id === utilSiteId)) setUtilSiteId("all");
  }, [sites, utilSiteId]);

  const inRange = bookings.filter((b) => b.date >= rangeStart && b.date <= rangeEnd);
  const confirmedInRange = inRange.filter((b) => b.status === "confirmed");
  const totalRevenue = confirmedInRange.reduce((s, b) => s + (Number(b.price) || 0), 0);
  const avgBookingValue = confirmedInRange.length ? totalRevenue / confirmedInRange.length : 0;

  const declinedOrCancelled = inRange.filter((b) => b.status === "declined" || b.status === "cancelled").length;
  const cancellationRate = inRange.length ? (declinedOrCancelled / inRange.length) * 100 : 0;

  // month-on-month style comparison: previous period of equal length immediately before this one
  const periodLen = daysBetween(rangeStart, rangeEnd);
  const prevEnd = addDays(rangeStart, -1);
  const prevStart = addDays(rangeStart, -periodLen);
  const prevRevenue = bookings.filter((b) => b.status === "confirmed" && b.date >= prevStart && b.date <= prevEnd).reduce((s, b) => s + (Number(b.price) || 0), 0);
  const revenueDelta = prevRevenue ? ((totalRevenue - prevRevenue) / prevRevenue) * 100 : null;

  // 6-month trend — always shows the trailing 6 months regardless of the range picker, for a stable long view
  const confirmedAll = bookings.filter((b) => b.status === "confirmed");
  const byMonth = {};
  confirmedAll.forEach((b) => { const k = monthKey(b.date); byMonth[k] = (byMonth[k] || 0) + Number(b.price || 0); });
  const monthData = Object.entries(byMonth).sort(([a], [b]) => a.localeCompare(b)).slice(-6).map(([k, v]) => ({ month: monthLabel(k), revenue: Math.round(v * 100) / 100 }));

  const bySite = {};
  confirmedInRange.forEach((b) => {
    const siteName = siteById[facilityById[b.facilityId]?.siteId]?.name || "Unknown";
    bySite[siteName] = (bySite[siteName] || 0) + Number(b.price || 0);
  });
  const siteData = Object.entries(bySite).sort(([, a], [, b]) => b - a).map(([name, revenue]) => ({ name, revenue: Math.round(revenue * 100) / 100 }));

  const byFacility = {};
  confirmedInRange.forEach((b) => {
    const name = facilityById[b.facilityId]?.name || "Unknown";
    byFacility[name] = (byFacility[name] || 0) + Number(b.price || 0);
  });
  const facilityData = Object.entries(byFacility).sort(([, a], [, b]) => b - a).slice(0, 8).map(([name, revenue]) => ({ name, revenue: Math.round(revenue * 100) / 100 }));

  const byMember = {};
  confirmedInRange.forEach((b) => {
    if (!b.memberId) return;
    byMember[b.memberId] = (byMember[b.memberId] || 0) + Number(b.price || 0);
  });
  const topMembers = Object.entries(byMember).sort(([, a], [, b]) => b - a).slice(0, 5)
    .map(([id, revenue]) => ({ member: memberById?.[id], revenue: Math.round(revenue * 100) / 100 }));

  const statusCounts = Object.keys(STATUS_META).map((k) => ({ key: k, count: inRange.filter((b) => b.status === k).length }));

  // utilisation: confirmed hours booked in the selected range per facility, vs. available hours in that window
  const utilisation = facilities
    .filter((f) => utilSiteId === "all" || f.siteId === utilSiteId)
    .map((f) => {
      const hrs = confirmedInRange.filter((b) => b.facilityId === f.id).reduce((s, b) => s + hoursBetween(b.startTime, b.endTime) * (Number(b.spaces) || 1), 0);
      const available = OPEN_HOURS_PER_DAY * periodLen * (f.capacity || 1);
      return { id: f.id, name: f.name, siteId: f.siteId, siteName: siteById[f.siteId]?.name, pct: available ? Math.min(100, (hrs / available) * 100) : 0, hrs };
    }).sort((a, b) => b.pct - a.pct);

  // busiest start-time slots within the selected range
  const slotCounts = {};
  confirmedInRange.forEach((b) => { slotCounts[b.startTime] = (slotCounts[b.startTime] || 0) + 1; });
  const busiestSlots = Object.entries(slotCounts).sort(([, a], [, b]) => b - a).slice(0, 5);

  function exportCSV() {
    const rows = inRange.map((b) => {
      const f = facilityById[b.facilityId];
      return [b.date, f?.name || "", siteById[f?.siteId]?.name || "", b.hirerName, b.company || "", b.startTime, b.endTime, hoursBetween(b.startTime, b.endTime).toFixed(2), b.price || 0, b.status];
    });
    downloadCSV(`delsport-bookings-${rangeStart}-to-${rangeEnd}.csv`,
      ["Date", "Facility", "Site", "Hirer", "Company", "Start", "End", "Hours", "Price", "Status"], rows);
  }

  if (bookings.length === 0) {
    return (
      <div>
        <PageHeader title="Reports" />
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
          <EmptyState icon={BarChart3} title="No data yet" sub="Reports will populate once you have bookings recorded." />
        </div>
      </div>
    );
  }

  return (
    <div>
      <PageHeader
        title="Reports"
        sub={`Showing ${dayLabel(rangeStart)} – ${dayLabel(rangeEnd)}`}
        action={<Btn variant="ghost" icon={FileDown} onClick={exportCSV}>Export bookings (.csv)</Btn>}
      />

      <div style={{ display: "flex", gap: 10, marginBottom: 20, alignItems: "center", flexWrap: "wrap" }}>
        <select value={preset} onChange={(e) => setPreset(e.target.value)} style={{ ...inputStyle, width: "auto" }}>
          {RANGE_PRESETS.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
        </select>
        {preset === "custom" && (
          <>
            <input type="date" style={{ ...inputStyle, width: "auto" }} value={customStart} onChange={(e) => setCustomStart(e.target.value)} />
            <span style={{ color: C.mute, fontSize: 13 }}>to</span>
            <input type="date" style={{ ...inputStyle, width: "auto" }} min={customStart} value={customEnd} onChange={(e) => setCustomEnd(e.target.value)} />
          </>
        )}
      </div>

      <div className="dp-stat-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginBottom: 20 }}>
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "14px 16px" }}>
          <div style={{ fontSize: 11.5, color: C.mute, fontWeight: 700, textTransform: "uppercase" }}>Confirmed revenue</div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginTop: 6 }}>
            <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: C.navy }}>{money(totalRevenue)}</span>
            {revenueDelta !== null && (
              <span style={{ display: "inline-flex", alignItems: "center", gap: 2, fontSize: 12, fontWeight: 700, color: revenueDelta >= 0 ? C.pitch : C.coral }}>
                {revenueDelta >= 0 ? <TrendingUp size={12} /> : <TrendingDown size={12} />} {Math.abs(revenueDelta).toFixed(0)}%
              </span>
            )}
          </div>
        </div>
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "14px 16px" }}>
          <div style={{ fontSize: 11.5, color: C.mute, fontWeight: 700, textTransform: "uppercase" }}>Avg. booking value</div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: C.navy, marginTop: 6 }}>{money(avgBookingValue)}</div>
        </div>
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "14px 16px" }}>
          <div style={{ fontSize: 11.5, color: C.mute, fontWeight: 700, textTransform: "uppercase" }}>Cancellation rate</div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: cancellationRate > 20 ? C.coral : C.navy, marginTop: 6 }}>{cancellationRate.toFixed(0)}%</div>
        </div>
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "14px 16px" }}>
          <div style={{ fontSize: 11.5, color: C.mute, fontWeight: 700, textTransform: "uppercase" }}>Bookings in range</div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: C.navy, marginTop: 6 }}>{inRange.length}</div>
        </div>
      </div>

      <div className="dp-stat-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginBottom: 20 }}>
        {statusCounts.map((s) => {
          const m = STATUS_META[s.key];
          return (
            <div key={s.key} style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "14px 16px" }}>
              <StatusPill status={s.key} />
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: C.navy, marginTop: 8 }}>{s.count}</div>
            </div>
          );
        })}
      </div>

      <div className="dp-dash-grid" style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 18, marginBottom: 18 }}>
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "16px 18px" }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, marginBottom: 12 }}>Confirmed revenue — trailing 6 months</div>
          {monthData.length === 0 ? <div style={{ color: C.mute, fontSize: 13 }}>No confirmed bookings yet.</div> : (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={monthData}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.line} vertical={false} />
                <XAxis dataKey="month" tick={{ fontSize: 12, fill: C.mute }} axisLine={{ stroke: C.line }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: C.mute }} axisLine={false} tickLine={false} tickFormatter={(v) => `£${v}`} />
                <Tooltip formatter={(v) => money(v)} contentStyle={{ borderRadius: 8, border: `1px solid ${C.line}`, fontSize: 13 }} />
                <Bar dataKey="revenue" fill={C.cyan} radius={[5, 5, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>

        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "16px 18px" }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, marginBottom: 12 }}>Top members by spend</div>
          {topMembers.length === 0 ? <div style={{ color: C.mute, fontSize: 13 }}>No confirmed bookings with a member attached yet.</div> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {topMembers.map((t, i) => (
                <div key={t.member?.id || i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, padding: "6px 0", borderBottom: i < topMembers.length - 1 ? `1px solid ${C.line}` : "none" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                    <Trophy size={13} color={i === 0 ? C.amber : C.mute} />
                    <span style={{ fontWeight: 600, color: C.ink }}>{t.member?.name || "Unknown member"}</span>
                  </div>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: C.pitch }}>{money(t.revenue)}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="dp-dash-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 18 }}>
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "16px 18px" }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, marginBottom: 12 }}>Revenue by site</div>
          {siteData.length === 0 ? <div style={{ color: C.mute, fontSize: 13 }}>No confirmed bookings in this range.</div> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {siteData.map((s) => {
                const pct = totalRevenue ? (s.revenue / totalRevenue) * 100 : 0;
                return (
                  <div key={s.name}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 4 }}>
                      <span style={{ fontWeight: 600, color: C.ink }}>{s.name}</span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.mute }}>{money(s.revenue)}</span>
                    </div>
                    <div style={{ height: 7, background: "#EEF1F4", borderRadius: 99 }}>
                      <div style={{ height: "100%", width: `${pct}%`, background: C.pitch, borderRadius: 99 }} />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "16px 18px" }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, marginBottom: 12 }}>Revenue by facility</div>
          {facilityData.length === 0 ? <div style={{ color: C.mute, fontSize: 13 }}>No confirmed bookings in this range.</div> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {facilityData.map((s) => {
                const pct = totalRevenue ? (s.revenue / totalRevenue) * 100 : 0;
                return (
                  <div key={s.name}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 4 }}>
                      <span style={{ fontWeight: 600, color: C.ink }}>{s.name}</span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.mute }}>{money(s.revenue)}</span>
                    </div>
                    <div style={{ height: 7, background: "#EEF1F4", borderRadius: 99 }}>
                      <div style={{ height: "100%", width: `${pct}%`, background: C.cyan, borderRadius: 99 }} />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      <div className="dp-dash-grid" style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 18 }}>
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "16px 18px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginBottom: 4, flexWrap: "wrap" }}>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy }}>Facility utilisation</div>
            {sites.length > 1 && (
              <select value={utilSiteId} onChange={(e) => setUtilSiteId(e.target.value)} style={{ ...inputStyle, width: "auto", padding: "4px 8px", fontSize: 12.5 }}>
                <option value="all">All sites</option>
                {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            )}
          </div>
          <div style={{ fontSize: 11.5, color: C.mute, marginBottom: 12 }}>Confirmed hours booked vs. available hours (07:00–22:00, all spaces) in the selected range</div>
          {utilisation.length === 0 ? <div style={{ color: C.mute, fontSize: 13 }}>Add a facility to see utilisation.</div> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {utilisation.map((u) => (
                <div key={u.id}>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 4 }}>
                    <span style={{ fontWeight: 600, color: C.ink }}>{u.name} <span style={{ color: C.mute, fontWeight: 400 }}>· {u.siteName}</span></span>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: C.mute }}>{u.pct.toFixed(0)}% · {u.hrs.toFixed(0)}h</span>
                  </div>
                  <div style={{ height: 7, background: "#EEF1F4", borderRadius: 99 }}>
                    <div style={{ height: "100%", width: `${u.pct}%`, background: u.pct > 70 ? C.pitch : u.pct > 30 ? C.cyan : C.amber, borderRadius: 99 }} />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: "16px 18px" }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, color: C.navy, marginBottom: 12 }}>Busiest start times</div>
          {busiestSlots.length === 0 ? <div style={{ color: C.mute, fontSize: 13 }}>No confirmed bookings in this range.</div> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {busiestSlots.map(([time, count], i) => (
                <div key={time} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, padding: "6px 0", borderBottom: i < busiestSlots.length - 1 ? `1px solid ${C.line}` : "none" }}>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: C.navy }}>{time}</span>
                  <span style={{ color: C.mute }}>{count} booking{count === 1 ? "" : "s"}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------- invoices & booking confirmations (shared PDF document builder) ----------
const DELSPORT_LOGO_URL = "https://static.wixstatic.com/media/3cf024_f16abd3550ca4fce86e2d135f9d0f27c~mv2.png";
// Sites that belong to a multi-academy trust carry that trust's logo
// alongside the school's own on invoices — keyed by site name so each site
// can point at whichever trust it's actually under.
// Prefixed with BASE_URL (not a plain root-absolute path) because these are read at
// runtime by JS, not by Vite's HTML asset pipeline — a hosting sub-path (e.g. GitHub
// Pages project sites) would otherwise 404 on a hardcoded "/foo.jpeg".
const SITE_TRUST_LOGO_URLS = {
  "St Augustine's Catholic High": `${import.meta.env.BASE_URL}magnificat-trust-logo.jpeg`,
  "St Benedict's Catholic High": `${import.meta.env.BASE_URL}magnificat-trust-logo.jpeg`,
  "Stourport High School": `${import.meta.env.BASE_URL}saet-trust-logo.png`,
  "Baxter College": `${import.meta.env.BASE_URL}saet-trust-logo.png`,
};
// Delsport UK stays the letterhead on every document, but payment for an
// invoice goes straight to whichever site did the hiring — bank details and
// VAT number live on the site record itself (Sites & Facilities), not here.
const DELSPORT_BUSINESS = {
  name: "Delsport UK Ltd",
  website: "www.delsportuk.com",
  email: "delsportuk@outlook.com",
  phones: ["07970 933445", "07851 319201"],
  paymentTerms: "Payable within 14 days by bank transfer.",
};

const DOC_COLORS = {
  navy: [11, 37, 69],
  cyan: [0, 180, 216],
  cyanSoft: [228, 247, 251],
  pitch: [46, 125, 91],
  pitchSoft: [231, 243, 237],
  mute: [100, 116, 139],
  ink: [30, 41, 59],
  line: [226, 232, 240],
  taglineBlue: [143, 169, 201],
};

// One continuous sequence across all sites (not per-site), since that's what
// a business's invoice numbering normally means — looks at every invoice
// raised anywhere, not just the ones for the site currently being billed.
function nextInvoiceNumber(allInvoices, year) {
  const prefix = `DEL-${year}-`;
  const used = allInvoices
    .map((inv) => inv.invoiceNumber)
    .filter((n) => n && n.startsWith(prefix))
    .map((n) => Number(n.slice(prefix.length)) || 0);
  const next = (used.length > 0 ? Math.max(...used) : 0) + 1;
  return `${prefix}${String(next).padStart(4, "0")}`;
}
// Only the diocese-trust sites need a reconciliation-friendly payment
// reference on the invoice itself — everyone else just uses the invoice
// number. Format: <site code><member first initial><3-letter surname>
// <running count for that member at that site>, e.g. STAmwad001.
const SITE_PAYMENT_REF_CODES = {
  "St Augustine's Catholic High": "STA",
  "St Benedict's Catholic High": "SBE",
  "Stourport High School": "SHS",
  "Baxter College": "BAX",
};
function buildPaymentReference(site, member, allInvoices) {
  const code = SITE_PAYMENT_REF_CODES[site?.name];
  if (!code) return null;
  const tokens = (member?.name || "").trim().split(/\s+/).filter(Boolean);
  const initial = (tokens[0]?.[0] || "").toLowerCase();
  const surname = (tokens[tokens.length - 1] || "").slice(0, 3).toLowerCase();
  const priorCount = allInvoices.filter((inv) => inv.siteId === site.id && inv.memberId === member.id).length;
  return `${code}${initial}${surname}${String(priorCount + 1).padStart(3, "0")}`;
}
// Reconstructs the payment reference an already-raised invoice was given at
// the time — same logic as buildPaymentReference, but ranking the invoice
// within its own site+member history instead of counting what came before a
// not-yet-created one. Lets the history list show it without storing it.
function paymentReferenceForInvoice(inv, allInvoices, siteById, memberById) {
  const site = siteById[inv.siteId];
  const code = SITE_PAYMENT_REF_CODES[site?.name];
  if (!code) return null;
  const member = memberById[inv.memberId];
  const tokens = (member?.name || "").trim().split(/\s+/).filter(Boolean);
  const initial = (tokens[0]?.[0] || "").toLowerCase();
  const surname = (tokens[tokens.length - 1] || "").slice(0, 3).toLowerCase();
  const sameGroup = allInvoices
    .filter((other) => other.siteId === inv.siteId && other.memberId === inv.memberId)
    .sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));
  const rank = sameGroup.findIndex((other) => other.id === inv.id) + 1;
  return `${code}${initial}${surname}${String(rank || 1).padStart(3, "0")}`;
}
// Confirmations aren't tracked/billing documents the way invoices are, so
// this doesn't need database-backed sequencing — just something unique and
// readable on the document.
function newConfirmationReference() {
  const rand = Math.random().toString(36).slice(2, 7).toUpperCase();
  return `CONF-${new Date().getFullYear()}-${rand}`;
}

// Fetches an image and returns it as a data URL for jsPDF's addImage — resolves
// to null on any failure (missing logo, network hiccup, blocked fetch) so the
// document still renders fine without it rather than throwing.
async function loadImageDataUrl(url) {
  if (!url) return null;
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

// Fits an image inside a box without distorting its aspect ratio.
function fitImage(doc, dataUrl, boxX, boxY, boxW, boxH) {
  try {
    const props = doc.getImageProperties(dataUrl);
    const ratio = Math.min(boxW / props.width, boxH / props.height);
    const w = props.width * ratio;
    const h = props.height * ratio;
    return { w, h, x: boxX + (boxW - w) / 2, y: boxY + (boxH - h) / 2, format: props.fileType };
  } catch {
    return null;
  }
}

function drawStatusPill(doc, text, x, y, bg, fg) {
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  const w = doc.getTextWidth(text) + 12;
  doc.setFillColor(...bg);
  doc.roundedRect(x, y - 8, w, 12, 6, 6, "F");
  doc.setTextColor(...fg);
  doc.text(text, x + 6, y);
}

// Navy header band shared by both document types — Delsport's own logo,
// the site's logo (its "facility partner" badge), and a big colour-coded
// title with a reference number. Returns the y-coordinate body content
// should start below.
function drawDocHeader(doc, { kind, reference, site, delsportLogo, siteLogo, trustLogo }) {
  const pageW = doc.internal.pageSize.getWidth();
  const marginX = 48;
  const headerH = 150;
  const { navy, cyan, pitch, taglineBlue } = DOC_COLORS;

  doc.setFillColor(...navy);
  doc.rect(0, 0, pageW, headerH, "F");

  if (delsportLogo) {
    const fit = fitImage(doc, delsportLogo, marginX, 30, 40, 40);
    if (fit) doc.addImage(delsportLogo, fit.format || "PNG", fit.x, fit.y, fit.w, fit.h);
  }
  doc.setFont("helvetica", "bold");
  doc.setFontSize(17);
  doc.setTextColor(255, 255, 255);
  doc.text("DELSPORT UK", marginX + 50, 48);
  doc.setFont("helvetica", "italic");
  doc.setFontSize(9);
  doc.setTextColor(...taglineBlue);
  doc.text("Get more from sport", marginX + 50, 61);

  const title = kind === "confirmation" ? "CONFIRMATION" : "Lettings Invoice";
  doc.setFont("helvetica", "bold");
  doc.setFontSize(21);
  doc.setTextColor(...(kind === "confirmation" ? pitch : cyan));
  doc.text(title, pageW - marginX, 46, { align: "right" });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.setTextColor(215, 226, 238);
  doc.text(reference, pageW - marginX, 62, { align: "right" });

  const boxX = marginX, boxY = 88, boxW = 90, boxH = 44;
  let logoRight = boxX;
  if (siteLogo) {
    const fit = fitImage(doc, siteLogo, logoRight, boxY, boxW, boxH);
    if (fit) doc.addImage(siteLogo, fit.format || "PNG", fit.x, fit.y, fit.w, fit.h);
    logoRight = boxX + boxW;
  }
  if (trustLogo) {
    const trustW = 44, gap = 12;
    const fit = fitImage(doc, trustLogo, logoRight + gap, boxY, trustW, boxH);
    if (fit) doc.addImage(trustLogo, fit.format || "PNG", fit.x, fit.y, fit.w, fit.h);
    logoRight = logoRight + gap + trustW;
  }

  const nameX = logoRight + 14;
  const nameMaxW = Math.max(pageW - marginX - nameX, 60);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7);
  doc.setTextColor(...taglineBlue);
  doc.text("FACILITY PARTNER", nameX, boxY + 12);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12.5);
  doc.setTextColor(255, 255, 255);
  doc.splitTextToSize(site?.name || "—", nameMaxW).slice(0, 2).forEach((l, i) => doc.text(l, nameX, boxY + 28 + i * 15));

  return headerH;
}

function drawDocFooter(doc, thanksLine) {
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const marginX = 48;
  const { mute, line } = DOC_COLORS;
  const fy = pageH - 40;
  doc.setDrawColor(...line);
  doc.line(marginX, fy, pageW - marginX, fy);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(...mute);
  doc.text(`Delsport UK · ${DELSPORT_BUSINESS.website} · ${DELSPORT_BUSINESS.email} | ${thanksLine}`, pageW / 2, fy + 16, { align: "center" });
}

// Two-column "who/what" block used near the top of both document types —
// left is the client, right is a set of key:value facts about the document.
function drawInfoColumns(doc, y, { leftLabel, rightLabel, member, rightRows }) {
  const pageW = doc.internal.pageSize.getWidth();
  const marginX = 48;
  const midX = pageW / 2 + 10;
  const { mute, ink } = DOC_COLORS;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(8.5);
  doc.setTextColor(...mute);
  doc.text(leftLabel, marginX, y);
  doc.text(rightLabel, midX, y);
  y += 16;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(...ink);
  const heading = member.company || member.name || "—";
  doc.text(heading, marginX, y);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.setTextColor(...mute);
  let leftY = y + 15;
  if (member.company && member.name) { doc.text(`FAO: ${member.name}`, marginX, leftY); leftY += 13; }
  if (member.email) { doc.text(member.email, marginX, leftY); leftY += 13; }
  if (member.phone) { doc.text(member.phone, marginX, leftY); leftY += 13; }

  let rightY = y;
  doc.setFontSize(9.5);
  rightRows.forEach(([label, value]) => {
    doc.setFont("helvetica", "normal");
    doc.setTextColor(...mute);
    doc.text(label, midX, rightY);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...ink);
    doc.text(String(value), pageW - marginX, rightY, { align: "right" });
    rightY += 15;
  });

  return Math.max(leftY, rightY) + 16;
}

async function buildInvoicePDF({ invoiceNumber, issuedDate, dueDate, member, site, periodLabel, rows, total, paymentReference }) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const marginX = 48;
  const { navy, cyanSoft, mute, ink, line, pitch } = DOC_COLORS;

  const trustLogoUrl = SITE_TRUST_LOGO_URLS[site?.name];
  const [delsportLogo, siteLogo, trustLogo] = await Promise.all([
    loadImageDataUrl(DELSPORT_LOGO_URL), loadImageDataUrl(site?.logoUrl), trustLogoUrl ? loadImageDataUrl(trustLogoUrl) : Promise.resolve(null),
  ]);
  let y = drawDocHeader(doc, { kind: "invoice", reference: paymentReference || invoiceNumber, site, delsportLogo, siteLogo, trustLogo });
  y += 40;

  y = drawInfoColumns(doc, y, {
    leftLabel: "BILL TO", rightLabel: "INVOICE DETAILS", member,
    rightRows: [["Site:", site?.name || "—"], ["Period:", periodLabel], ["Issue date:", issuedDate], ["Due date:", dueDate]],
  });

  const colX = { date: marginX, facility: marginX + 60, time: marginX + 220, duration: marginX + 300 };
  doc.setFillColor(248, 250, 252);
  doc.rect(marginX, y - 12, pageW - marginX * 2, 20, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.setTextColor(...mute);
  doc.text("DATE", colX.date, y);
  doc.text("FACILITY", colX.facility, y);
  doc.text("TIME", colX.time, y);
  doc.text("DURATION", colX.duration, y);
  doc.text("FEE", pageW - marginX, y, { align: "right" });
  y += 16;
  doc.setDrawColor(...line);
  doc.line(marginX, y - 10, pageW - marginX, y - 10);

  const facilityColWidth = colX.time - colX.facility - 10;
  rows.forEach((r, i) => {
    if (y > pageH - 190) { doc.addPage(); y = 64; }
    // A combined multi-facility label ("Sports Hall + Classroom") can run
    // well past a single-facility one, so it wraps instead of overlapping
    // the Time column — the row just grows to fit it. Font must be set
    // before measuring, or the wrap width is calculated against whatever
    // font size was last active (wrong for bold vs normal).
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.5);
    const facilityLines = doc.splitTextToSize(r.facilityLabel, facilityColWidth);
    const rowH = Math.max(30, facilityLines.length * 11 + 21);
    if (i % 2 === 1) { doc.setFillColor(250, 251, 252); doc.rect(marginX, y - 12, pageW - marginX * 2, rowH, "F"); }
    doc.setTextColor(...ink);
    doc.text(r.dateLabel, colX.date, y);
    facilityLines.forEach((line, li) => doc.text(line, colX.facility, y + li * 11));
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(...mute);
    doc.text(r.siteName || "", colX.facility, y + facilityLines.length * 11);
    doc.setTextColor(...ink);
    doc.text(r.timeLabel, colX.time, y);
    doc.text(r.durationLabel, colX.duration, y);
    doc.text(money(r.price), pageW - marginX, y, { align: "right" });
    y += rowH;
  });

  y += 10;
  const summaryX = pageW - marginX - 170;
  doc.setDrawColor(...line);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.setTextColor(...ink);
  doc.text("Subtotal", summaryX, y);
  doc.text(money(total), pageW - marginX, y, { align: "right" });
  y += 15;
  doc.setTextColor(...mute);
  doc.text("VAT (0% — not registered)", summaryX, y);
  doc.setTextColor(...ink);
  doc.text(money(0), pageW - marginX, y, { align: "right" });
  y += 8;
  doc.line(summaryX, y, pageW - marginX, y);
  y += 18;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.setTextColor(...navy);
  doc.text("Total due", summaryX, y);
  doc.setTextColor(...pitch);
  doc.text(money(total), pageW - marginX, y, { align: "right" });

  // Payment goes straight to the site, so its bank details drive this box.
  const bankLines = [];
  if (site?.bankAccountName) bankLines.push(`Account name: ${site.bankAccountName}`);
  if (site?.bankSortCode) bankLines.push(`Sort code: ${site.bankSortCode}`);
  if (site?.bankAccountNumber) bankLines.push(`Account number: ${site.bankAccountNumber}`);
  const bankLine = bankLines.length > 0 ? bankLines.join("   ·   ") : `Contact ${site?.name || "the site"} for payment details.`;

  const detailLines = [DELSPORT_BUSINESS.paymentTerms, bankLine, `Reference: ${paymentReference || invoiceNumber}`];
  if (site?.vatNumber) detailLines.push(`VAT number: ${site.vatNumber}`);

  y += 30;
  const boxH = 32 + (detailLines.length - 1) * 13 + 14;
  doc.setFillColor(...cyanSoft);
  doc.roundedRect(marginX, y, pageW - marginX * 2, boxH, 6, 6, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.setTextColor(...navy);
  doc.text("PAYMENT DETAILS", marginX + 14, y + 18);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(...ink);
  detailLines.forEach((line, i) => doc.text(line, marginX + 14, y + 32 + i * 13));

  drawDocFooter(doc, "Thank you for partnering with Delsport UK");
  return doc;
}

async function buildConfirmationPDF({ reference, issuedDate, member, site, periodLabel, rows, total }) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const marginX = 48;
  const { navy, pitch, pitchSoft, cyanSoft, mute, ink, line } = DOC_COLORS;

  const [delsportLogo, siteLogo] = await Promise.all([loadImageDataUrl(DELSPORT_LOGO_URL), loadImageDataUrl(site?.logoUrl)]);
  let y = drawDocHeader(doc, { kind: "confirmation", reference, site, delsportLogo, siteLogo });
  y += 26;

  doc.setFillColor(...pitchSoft);
  doc.roundedRect(marginX, y, pageW - marginX * 2, 24, 6, 6, "F");
  doc.setDrawColor(...pitch);
  doc.setLineWidth(1.4);
  doc.line(marginX + 10, y + 13, marginX + 13, y + 16);
  doc.line(marginX + 13, y + 16, marginX + 19, y + 9);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9.5);
  doc.setTextColor(...pitch);
  doc.text("All bookings listed below are confirmed — no action needed unless a date changes.", marginX + 26, y + 16);
  y += 46;

  y = drawInfoColumns(doc, y, {
    leftLabel: "CONFIRMED FOR", rightLabel: "CONFIRMATION DETAILS", member,
    rightRows: [["Site:", site?.name || "—"], ["Period:", periodLabel], ["Confirmed on:", issuedDate], ["Reference:", reference]],
  });

  const colX = { date: marginX, facility: marginX + 60, time: marginX + 200, duration: marginX + 270, status: marginX + 330 };
  doc.setFillColor(248, 250, 252);
  doc.rect(marginX, y - 12, pageW - marginX * 2, 20, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.setTextColor(...mute);
  doc.text("DATE", colX.date, y);
  doc.text("FACILITY", colX.facility, y);
  doc.text("TIME", colX.time, y);
  doc.text("DURATION", colX.duration, y);
  doc.text("STATUS", colX.status, y);
  doc.text("PRICE", pageW - marginX, y, { align: "right" });
  y += 16;
  doc.setDrawColor(...line);
  doc.line(marginX, y - 10, pageW - marginX, y - 10);

  const facilityColWidth = colX.time - colX.facility - 10;
  rows.forEach((r, i) => {
    if (y > pageH - 190) { doc.addPage(); y = 64; }
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.5);
    const facilityLines = doc.splitTextToSize(r.facilityLabel, facilityColWidth);
    const rowH = Math.max(30, facilityLines.length * 11 + 21);
    if (i % 2 === 1) { doc.setFillColor(250, 251, 252); doc.rect(marginX, y - 12, pageW - marginX * 2, rowH, "F"); }
    doc.setTextColor(...ink);
    doc.text(r.dateLabel, colX.date, y);
    facilityLines.forEach((line, li) => doc.text(line, colX.facility, y + li * 11));
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(...mute);
    doc.text(r.siteName || "", colX.facility, y + facilityLines.length * 11);
    doc.setTextColor(...ink);
    doc.text(r.timeLabel, colX.time, y);
    doc.text(r.durationLabel, colX.duration, y);
    drawStatusPill(doc, "Confirmed", colX.status, y, pitchSoft, pitch);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9.5);
    doc.setTextColor(...ink);
    doc.text(money(r.price), pageW - marginX, y, { align: "right" });
    y += rowH;
  });

  y += 6;
  doc.setDrawColor(...line);
  doc.line(marginX, y, pageW - marginX, y);
  y += 18;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11.5);
  doc.setTextColor(...navy);
  doc.text("Total value of bookings", colX.duration, y);
  doc.text(money(total), pageW - marginX, y, { align: "right" });
  y += 13;
  doc.setFont("helvetica", "italic");
  doc.setFontSize(8);
  doc.setTextColor(...mute);
  doc.text("Shown for your records — this is a confirmation, not an invoice.", pageW - marginX, y, { align: "right" });

  y += 26;
  doc.setFillColor(...cyanSoft);
  doc.roundedRect(marginX, y, pageW - marginX * 2, 54, 6, 6, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.setTextColor(...navy);
  doc.text("NEED TO MAKE A CHANGE?", marginX + 14, y + 18);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(...ink);
  const changeText = `Get in touch with Delsport UK as early as possible if any of these dates or times need to move — reply to this email or call us on ${DELSPORT_BUSINESS.phones.join(" / ")}.`;
  doc.splitTextToSize(changeText, pageW - marginX * 2 - 28).forEach((l, i) => doc.text(l, marginX + 14, y + 32 + i * 12));

  drawDocFooter(doc, "Thank you for booking with Delsport UK");
  return doc;
}

function Invoices({ sites, facilities, bookings, members, facilityById, siteById, memberById, invoices, allInvoices, setInvoices }) {
  const [siteId, setSiteId] = useState(sites[0]?.id || "");
  const [periodStart, setPeriodStart] = useState(startOfMonthISO(addDays(todayISO(), -30)));
  const [periodEnd, setPeriodEnd] = useState(endOfMonthISO(addDays(todayISO(), -30)));
  const [generating, setGenerating] = useState(false);
  const [resultMsg, setResultMsg] = useState("");
  const [selectedPreview, setSelectedPreview] = useState(() => new Set());
  const [selectedHistory, setSelectedHistory] = useState(() => new Set());
  const [deleting, setDeleting] = useState(false);
  const [confirmDeleteInvoices, setConfirmDeleteInvoices] = useState(false);

  useEffect(() => {
    if ((!siteId || !sites.some((s) => s.id === siteId)) && sites.length > 0) setSiteId(sites[0].id);
  }, [sites, siteId]);

  useEffect(() => {
    setSelectedPreview(new Set()); // stale selections don't carry across a site/period change
  }, [siteId, periodStart, periodEnd]);

  const siteFacilityIds = useMemo(() => new Set(facilities.filter((f) => f.siteId === siteId).map((f) => f.id)), [facilities, siteId]);
  const alreadyInvoicedIds = useMemo(() => {
    const ids = new Set();
    invoices.filter((inv) => inv.siteId === siteId).forEach((inv) => (inv.bookingIds || []).forEach((id) => ids.add(id)));
    return ids;
  }, [invoices, siteId]);

  const eligible = useMemo(() => bookings.filter((b) =>
    siteFacilityIds.has(b.facilityId) &&
    b.status === "confirmed" &&
    // a grouped booking's £0 legs are still eligible — the group carries a
    // real price collectively even though only one leg holds it
    ((Number(b.price) || 0) > 0 || !!b.groupId) &&
    b.date >= periodStart && b.date <= periodEnd &&
    !alreadyInvoicedIds.has(b.id)
  ), [bookings, siteFacilityIds, periodStart, periodEnd, alreadyInvoicedIds]);

  const unmatched = eligible.filter((b) => !b.memberId || !memberById[b.memberId]);
  const byMember = useMemo(() => {
    const groups = new Map();
    eligible.forEach((b) => {
      if (!b.memberId || !memberById[b.memberId]) return;
      if (!groups.has(b.memberId)) groups.set(b.memberId, []);
      groups.get(b.memberId).push(b);
    });
    return Array.from(groups.entries()).map(([memberId, list]) => ({
      member: memberById[memberId],
      bookings: list.sort((a, b) => a.date.localeCompare(b.date)),
      total: list.reduce((s, b) => s + (Number(b.price) || 0), 0),
    })).sort((a, b) => a.member.name.localeCompare(b.member.name));
  }, [eligible, memberById]);

  const site = siteById[siteId];
  const periodLabel = `${dayLabel(periodStart)} – ${dayLabel(periodEnd)}`;

  function togglePreview(memberId) {
    setSelectedPreview((s) => { const next = new Set(s); if (next.has(memberId)) next.delete(memberId); else next.add(memberId); return next; });
  }
  function toggleAllPreview() {
    setSelectedPreview((s) => (s.size === byMember.length ? new Set() : new Set(byMember.map((g) => g.member.id))));
  }

  // Shared by "raise all" and "raise selected" — builds a PDF per member,
  // uploads each to Supabase Storage so it has a permanent link in the
  // history list below, and zips them together for one download if there's
  // more than one (a single invoice downloads directly, no zip needed).
  async function raiseGroups(groups) {
    if (groups.length === 0) return;
    setGenerating(true);
    setResultMsg("");
    try {
      const zip = groups.length > 1 ? new JSZip() : null;
      const year = new Date().getFullYear();
      const newInvoices = [];
      let soloBlob = null;
      let soloFilename = "";
      for (const group of groups) {
        const invoiceNumber = nextInvoiceNumber([...allInvoices, ...newInvoices], year);
        const paymentReference = buildPaymentReference(site, group.member, [...allInvoices, ...newInvoices]);
        const rows = buildDocumentRows(group.bookings, facilityById, site?.name);
        const doc = await buildInvoicePDF({
          invoiceNumber, issuedDate: dayLabel(todayISO()), dueDate: dayLabel(addDays(todayISO(), 14)),
          member: group.member, site, periodLabel, rows, total: group.total, paymentReference,
        });
        const blob = doc.output("blob");
        const safeName = group.member.name.replace(/[^a-z0-9]+/gi, "-").toLowerCase();
        const filename = `${invoiceNumber}-${safeName}.pdf`;
        const pdfPath = `${siteId}/${filename}`;

        const { error: uploadErr } = await supabase.storage.from("invoices").upload(pdfPath, blob, { contentType: "application/pdf", upsert: true });
        if (uploadErr) throw uploadErr;

        if (zip) zip.file(filename, blob);
        else { soloBlob = blob; soloFilename = filename; }

        newInvoices.push({
          id: uid(), invoiceNumber, memberId: group.member.id, siteId,
          periodStart, periodEnd, total: group.total,
          bookingIds: group.bookings.map((b) => b.id),
          pdfPath,
          createdAt: new Date().toISOString(), // display-only; the real value comes from the DB default on reload
        });
      }
      if (zip) {
        const zipBlob = await zip.generateAsync({ type: "blob" });
        const siteSlug = (site?.name || "site").replace(/[^a-z0-9]+/gi, "-").toLowerCase();
        downloadBlob(`delsport-invoices-${siteSlug}-${periodStart}.zip`, zipBlob);
      } else if (soloBlob) {
        downloadBlob(soloFilename, soloBlob);
      }
      setInvoices((prev) => [...prev, ...newInvoices]);
      setSelectedPreview(new Set());
      setResultMsg(`Raised ${newInvoices.length} invoice${newInvoices.length === 1 ? "" : "s"}, totalling ${money(newInvoices.reduce((s, i) => s + i.total, 0))}.`);
    } catch (e) {
      setResultMsg(`Couldn't raise invoices: ${e.message || e}`);
    } finally {
      setGenerating(false);
    }
  }

  const selectedGroups = byMember.filter((g) => selectedPreview.has(g.member.id));

  async function downloadInvoicePdf(inv) {
    if (!inv.pdfPath) return;
    const { data, error } = await supabase.storage.from("invoices").createSignedUrl(inv.pdfPath, 60);
    if (error) { setResultMsg(`Couldn't open that PDF: ${error.message}`); return; }
    window.open(data.signedUrl, "_blank");
  }

  const history = [...invoices].sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || ""));

  function toggleHistory(id) {
    setSelectedHistory((s) => { const next = new Set(s); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  }
  function toggleAllHistory() {
    setSelectedHistory((s) => (s.size === history.length ? new Set() : new Set(history.map((inv) => inv.id))));
  }
  async function deleteSelectedHistory() {
    const toDelete = history.filter((inv) => selectedHistory.has(inv.id));
    if (toDelete.length === 0) return;
    setDeleting(true);
    try {
      const paths = toDelete.map((inv) => inv.pdfPath).filter(Boolean);
      if (paths.length > 0) await supabase.storage.from("invoices").remove(paths);
      const idsToDelete = new Set(toDelete.map((inv) => inv.id));
      setInvoices((prev) => prev.filter((inv) => !idsToDelete.has(inv.id)));
      setSelectedHistory(new Set());
    } finally {
      setDeleting(false);
    }
  }

  return (
    <div>
      <PageHeader title="Invoices" sub="Generate monthly invoices as PDFs, one per member, for a single site." />

      {sites.length === 0 ? (
        <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}` }}>
          <EmptyState icon={Receipt} title="Add a site first" sub="You'll need at least one site with bookings before you can raise invoices." />
        </div>
      ) : (
        <>
          <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, padding: 18, marginBottom: 20 }}>
            <div className="dp-invoice-filters" style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr 1fr", gap: 16, marginBottom: 16 }}>
              <Field label="Site">
                <select style={inputStyle} value={siteId} onChange={(e) => setSiteId(e.target.value)}>
                  {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </Field>
              <Field label="Period start">
                <input type="date" style={inputStyle} value={periodStart} onChange={(e) => setPeriodStart(e.target.value)} />
              </Field>
              <Field label="Period end">
                <input type="date" style={inputStyle} value={periodEnd} onChange={(e) => setPeriodEnd(e.target.value)} min={periodStart} />
              </Field>
            </div>

            {byMember.length === 0 ? (
              <div style={{ fontSize: 13.5, color: C.mute, padding: "10px 0" }}>
                No un-invoiced confirmed bookings with a price at {site?.name || "this site"} in that period.
              </div>
            ) : (
              <div style={{ border: `1px solid ${C.line}`, borderRadius: 10, overflow: "hidden", marginBottom: 16 }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                      <th style={{ padding: "9px 14px", width: 1 }}>
                        <input type="checkbox" checked={selectedPreview.size > 0 && selectedPreview.size === byMember.length} onChange={toggleAllPreview} />
                      </th>
                      {["Member", "Sessions", "Total"].map((h) => (
                        <th key={h} style={{ padding: "9px 14px", fontSize: 11, textTransform: "uppercase", color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {byMember.map((g) => (
                      <tr key={g.member.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                        <td style={{ padding: "9px 14px" }}>
                          <input type="checkbox" checked={selectedPreview.has(g.member.id)} onChange={() => togglePreview(g.member.id)} />
                        </td>
                        <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 600, color: C.ink }}>{g.member.name}{g.member.company ? ` — ${g.member.company}` : ""}</td>
                        <td style={{ padding: "9px 14px", fontSize: 13 }}>{g.bookings.length}</td>
                        <td style={{ padding: "9px 14px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: C.pitch }}>{money(g.total)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {unmatched.length > 0 && (
              <div style={{ fontSize: 12.5, color: C.amber, marginBottom: 12 }}>
                {unmatched.length} booking{unmatched.length === 1 ? "" : "s"} in this period have no linked member and were skipped — link them to a member to include them.
              </div>
            )}

            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <Btn variant="accent" icon={Receipt} onClick={() => raiseGroups(byMember)} disabled={byMember.length === 0 || generating}>
                {generating ? "Generating…" : `Raise all ${byMember.length} for ${site?.name || "this site"}`}
              </Btn>
              <Btn variant="ghost" icon={Receipt} onClick={() => raiseGroups(selectedGroups)} disabled={selectedGroups.length === 0 || generating}>
                Raise {selectedGroups.length || ""} selected
              </Btn>
            </div>
            {resultMsg && <div style={{ fontSize: 12.5, color: resultMsg.startsWith("Couldn't") ? C.coral : C.pitch, marginTop: 10, fontWeight: 600 }}>{resultMsg}</div>}
          </div>

          <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.line}`, overflow: "hidden" }}>
            <div style={{ padding: "14px 18px", background: C.cyan, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, color: C.navy }}>Invoice history</span>
              {selectedHistory.size > 0 && (
                <Btn small variant="danger" icon={Trash2} onClick={() => setConfirmDeleteInvoices(true)} disabled={deleting}>
                  {deleting ? "Deleting…" : `Delete ${selectedHistory.size} selected`}
                </Btn>
              )}
            </div>
            {history.length === 0 ? (
              <EmptyState icon={Receipt} title="No invoices raised yet" sub="Once you raise invoices, they'll be listed here." />
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: "#F8FAFC", textAlign: "left" }}>
                    <th style={{ padding: "9px 14px", width: 1 }}>
                      <input type="checkbox" checked={selectedHistory.size > 0 && selectedHistory.size === history.length} onChange={toggleAllHistory} />
                    </th>
                    {["Invoice #", "Member", "Site", "Period", "Total", "Issued", ""].map((h) => (
                      <th key={h} style={{ padding: "9px 14px", fontSize: 11, textTransform: "uppercase", color: C.mute, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {history.map((inv) => (
                    <tr key={inv.id} style={{ borderBottom: `1px solid ${C.line}` }}>
                      <td style={{ padding: "9px 14px" }}>
                        <input type="checkbox" checked={selectedHistory.has(inv.id)} onChange={() => toggleHistory(inv.id)} />
                      </td>
                      <td style={{ padding: "9px 14px", fontSize: 12.5, fontFamily: "'IBM Plex Mono', monospace" }}>{paymentReferenceForInvoice(inv, history, siteById, memberById) || inv.invoiceNumber}</td>
                      <td style={{ padding: "9px 14px", fontSize: 13 }}>{memberById[inv.memberId]?.name || "—"}</td>
                      <td style={{ padding: "9px 14px", fontSize: 13 }}>{siteById[inv.siteId]?.name || "—"}</td>
                      <td style={{ padding: "9px 14px", fontSize: 12.5, color: C.mute }}>{dayLabel(inv.periodStart)} – {dayLabel(inv.periodEnd)}</td>
                      <td style={{ padding: "9px 14px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{money(inv.total)}</td>
                      <td style={{ padding: "9px 14px", fontSize: 12.5, color: C.mute }}>{inv.createdAt ? new Date(inv.createdAt).toLocaleDateString("en-GB") : "—"}</td>
                      <td style={{ padding: "9px 14px" }}>
                        {inv.pdfPath && <Btn small variant="ghost" icon={FileDown} onClick={() => downloadInvoicePdf(inv)}>PDF</Btn>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </>
      )}
      {confirmDeleteInvoices && (
        <ConfirmDeleteModal
          title="Delete selected invoices?"
          message={`${selectedHistory.size} invoice${selectedHistory.size === 1 ? "" : "s"} — and its PDF — will be permanently removed. This can't be undone.`}
          onCancel={() => setConfirmDeleteInvoices(false)}
          onConfirm={() => { deleteSelectedHistory(); setConfirmDeleteInvoices(false); }}
        />
      )}
    </div>
  );
}
