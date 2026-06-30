// Shared data layer for index.html (builder) and admin.html (live viewer).
//
// SOURCE OF TRUTH: pricing-data.xlsx, committed in this same folder in the
// Git repo. It has two tabs: "Offerings" and "Settings" (same column schema
// as before). To update pricing: open pricing-data.xlsx (Excel/LibreOffice/
// Google Sheets — any of them can open and re-save .xlsx), edit it, save,
// then commit + push. Changes go live for everyone on next page load after
// the new commit is deployed.
//
// NOTE: only one person should edit at a time — since this is a binary file
// in Git, two people editing and pushing around the same time will silently
// overwrite each other rather than merge. Coordinate before editing.
//
// If the file is missing/unreadable, the site silently falls back to the
// bundled DEFAULT_OFFERINGS / DEFAULT_SETTINGS below.
const DATA_FILE_URL = "pricing-data.xlsx";

// Requires SheetJS (loaded via <script> tag in index.html / admin.html before this file).
const XLSX_CDN_NOTE = "Loaded from https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";

// ---- Bundled fallback data (used until the Sheet is wired up, or if it's unreachable) ----
const DEFAULT_OFFERINGS = {
  "Digital & Social": {
    icon: "📱",
    items: [
      { id: "social-media-promo", name: "Social Media Promotions (Insta/FB)", value: 25000, meaning: "Your brand featured across Ponn Onam 2026's official Instagram & Facebook posts, seen by thousands of residents and followers." },
      { id: "whatsapp-announce", name: "WhatsApp Group Announcements", value: 10000, meaning: "Direct mentions in resident WhatsApp groups — high open-rate, zero ad-spend wastage." },
      { id: "website-logo", name: "Website Logo Placement", value: 8000, meaning: "Permanent logo placement on the official Ponn Onam 2026 sponsor website." },
      { id: "insta-reel", name: "Instagram Reel / Highlight Feature", value: 15000, meaning: "A dedicated reel or story highlight showcasing your brand to the community's social audience." },
    ]
  },
  "Expo & Stall Space": {
    icon: "🏪",
    items: [
      { id: "stall-property", name: "Stall Space — Property Expo (Jul 25 & 26)", value: 75000, meaning: "A premium stall at the Property Expo on Jul 25 & 26, 2026, face-to-face with qualified, high-intent residential buyers." },
      { id: "stall-auto", name: "Stall Space — Auto Expo", value: 75000, meaning: "Showcase vehicles directly to residents during the Auto Expo — test drives, walk-ins, and live demos." },
      { id: "stall-lifestyle", name: "Stall Space — Lifestyle Expo", value: 60000, meaning: "A dedicated lifestyle stall reaching families during peak festival footfall." },
      { id: "entrance-branding", name: "Venue Entrance Branding", value: 40000, meaning: "First-impression branding at the venue entrance — every attendee walks past your logo." },
    ]
  },
  "Main Event — Ponn Onam (Sep 4, 5 & 6)": {
    icon: "🎤",
    items: [
      { id: "stage-branding", name: "Stage Branding", value: 100000, meaning: "Your logo on the main stage backdrop, visible in every photo, video, and livestream across all 3 nights (Sep 4–6)." },
      { id: "mc-mentions", name: "MC Mentions / Announcements", value: 20000, meaning: "Live on-mic shoutouts from the event MC throughout the 3-day celebration." },
      { id: "naming-rights", name: "Main Event Naming Rights", value: 300000, meaning: "\"Ponn Onam 2026, presented by [Your Brand]\" — top-tier recall across every touchpoint." },
      { id: "photo-backdrop", name: "Photo-Op Backdrop Branding", value: 35000, meaning: "Your brand in the background of thousands of resident selfies and family photos, organically shared on social media." },
      { id: "stall-main-event", name: "Stall Space — Main Event (2 nights)", value: 75000, meaning: "A branded stall at the main festival for 2 nights, engaging directly with 11,000+ residents at peak footfall." },
    ]
  },
  "Visuals & Activation": {
    icon: "🎬",
    items: [
      { id: "led-screen", name: "LED Screen / Festival Visuals Branding", value: 90000, meaning: "Your brand film or ad looped on the large LED screen seen by the entire festival crowd each night." },
      { id: "activation-booth", name: "Custom Brand Activation Booth", value: 120000, meaning: "A fully customized, branded activation space designed around your product or service." },
      { id: "photo-video-feature", name: "Photo/Video Coverage Feature", value: 30000, meaning: "Your brand featured in the official event photo/video coverage, used in post-event recap content." },
      { id: "pookalam-branding", name: "Pookalam Branding", value: 50000, meaning: "Association with the iconic Pookalam — Onam's most photographed and shared moment." },
    ]
  },
  "Community Engagement": {
    icon: "🤝",
    items: [
      { id: "contest-giveaway", name: "Contest / Giveaway Sponsorship", value: 25000, meaning: "Run a branded contest or giveaway, driving direct interaction and goodwill with residents." },
      { id: "sadya-sponsorship", name: "Onam Sadya Sponsorship", value: 150000, meaning: "Sponsor the Grand Onam Sadya — Onam's most cherished tradition, with 2,000+ participants." },
      { id: "cultural-program", name: "Cultural Program Sponsorship", value: 80000, meaning: "Brand association with traditional cultural performances and heritage programs." },
    ]
  },
};

const DEFAULT_SETTINGS = {
  gstRatePercent: 18,
  eventName: "Ponn Onam 2026",
  mainEventDates: "Sep 4, 5 & 6, 2026",
  propertyExpoDates: "Jul 25 & 26, 2026",
};

// ---- XLSX loading (via SheetJS) ----
async function fetchWorkbook(url){
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error('xlsx fetch failed: ' + res.status);
  const buf = await res.arrayBuffer();
  if (typeof XLSX === 'undefined') throw new Error('SheetJS (XLSX) library not loaded');
  return XLSX.read(buf, { type: 'array' });
}

function sheetToRows(workbook, sheetName){
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Sheet tab "${sheetName}" not found in pricing-data.xlsx`);
  // defval ensures missing cells come through as '' instead of being dropped
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function rowsToOfferings(rows){
  const grouped = {};
  rows.forEach(r => {
    const cat = r.category;
    if (!cat) return;
    if (!grouped[cat]) grouped[cat] = { icon: r.icon || '⭐', items: [] };
    grouped[cat].items.push({
      id: r.id || (cat + '-' + r.name).toLowerCase().replace(/[^a-z0-9]+/g, '-'),
      name: r.name || 'Untitled Offering',
      value: parseFloat(r.value) || 0,
      meaning: r.meaning || ''
    });
  });
  return grouped;
}

function rowsToSettings(rows){
  const s = { ...DEFAULT_SETTINGS };
  rows.forEach(r => {
    if (!r.key) return;
    if (r.key === 'gstRatePercent') s.gstRatePercent = parseFloat(r.value) || s.gstRatePercent;
    else if (r.key in s) s[r.key] = r.value;
  });
  return s;
}

// Fetch pricing-data.xlsx once; both loaders share the same in-flight/cached request.
let _workbookPromise = null;
function getWorkbook(){
  if (!_workbookPromise) _workbookPromise = fetchWorkbook(DATA_FILE_URL);
  return _workbookPromise;
}

// ---- Public async loaders: { data, source: 'file' | 'default', error? } ----
async function loadOfferingsAsync(){
  try {
    const wb = await getWorkbook();
    const data = rowsToOfferings(sheetToRows(wb, 'Offerings'));
    if (Object.keys(data).length) return { data, source: 'file' };
    throw new Error('Offerings tab parsed to zero rows');
  } catch (e) {
    return { data: JSON.parse(JSON.stringify(DEFAULT_OFFERINGS)), source: 'default', error: e.message };
  }
}

async function loadSettingsAsync(){
  try {
    const wb = await getWorkbook();
    const data = rowsToSettings(sheetToRows(wb, 'Settings'));
    return { data, source: 'file' };
  } catch (e) {
    return { data: { ...DEFAULT_SETTINGS }, source: 'default', error: e.message };
  }
}
