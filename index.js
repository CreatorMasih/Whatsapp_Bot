import makeWASocket, {
  useMultiFileAuthState,
  DisconnectReason,
  Browsers,
  fetchLatestBaileysVersion
} from "@whiskeysockets/baileys";

import fs from "node:fs";
import http from "node:http";
import dns from "node:dns";
import qrcode from "qrcode-terminal";
import cron from "node-cron";
import { google } from "googleapis";
import dayjs from "dayjs";
import pino from "pino";
import customParseFormat from "dayjs/plugin/customParseFormat.js";

dayjs.extend(customParseFormat);

if (typeof dns.setDefaultResultOrder === "function") {
  dns.setDefaultResultOrder("ipv4first");
}

let cronTask = null;
let reconnectTimer = null;
let reconnectAttempts = 0;
const RECONNECT_BASE_DELAY_MS = 3000;
const RECONNECT_MAX_DELAY_MS = 30000;

// =============================
// GOOGLE SHEET CONFIG
// =============================

const SHEET_ID = process.env.SHEET_ID || "1e0LzlBvuXqDt9r3mu24Z1DCFTSrtFXR1IYB0R6VAbK8";
const SHEET_NAME = process.env.SHEET_NAME || "Sheet1";
const AUTH_DIR = process.env.WA_AUTH_DIR || "./auth";
const WA_PAIRING_NUMBER = (process.env.WA_PAIRING_NUMBER || "").replace(/\D/g, "");
const PORT = Number(process.env.PORT || 0);
const BOOT_SIGNATURE = "boot-2026-03-03-v1";
const GOOGLE_SERVICE_ACCOUNT_PATH =
  process.env.GOOGLE_SERVICE_ACCOUNT_PATH || "./service-account.json";
let resolvedSheetName = null;

function startHealthServer() {
  if (!PORT || Number.isNaN(PORT)) return;

  const server = http.createServer((req, res) => {
    if (req.url === "/health") {
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: true }));
      return;
    }

    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("wa-bot running");
  });

  server.listen(PORT, () => {
    console.log(`Health server listening on port ${PORT}`);
  });
}

const googleAuthOptions = {
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
};

if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
  try {
    googleAuthOptions.credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  } catch (err) {
    console.log("Invalid GOOGLE_SERVICE_ACCOUNT_JSON:", err?.message || err);
    process.exit(1);
  }
} else {
  googleAuthOptions.keyFile = GOOGLE_SERVICE_ACCOUNT_PATH;
}

const authGoogle = new google.auth.GoogleAuth(googleAuthOptions);

const sheets = google.sheets({
  version: "v4",
  auth: authGoogle,
});

function getGoogleErrorMessage(err) {
  const status = err?.response?.status;
  const apiMsg = err?.response?.data?.error?.message;
  const localMsg = err?.message;
  return `status=${status ?? "unknown"} message=${apiMsg || localMsg || "unknown error"}`;
}

function isValidScheduleHeader(headerRow) {
  const values = (headerRow || []).map((value) => String(value || "").trim().toUpperCase());
  const [colA, colB, colC, colD] = values;
  const hasDate = ["DATE", "DATE/RULE"].includes(colA);
  const hasTime = colB === "TIME";
  const hasMessage = colC === "MESSAGE";
  const hasGroup = colD?.startsWith("GROUP");
  return hasDate && hasTime && hasMessage && hasGroup;
}

async function resolveScheduleSheetName() {
  if (resolvedSheetName) return resolvedSheetName;

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SHEET_ID,
    fields: "sheets.properties.title",
  });

  const allTitles = (meta.data.sheets || [])
    .map((sheet) => sheet?.properties?.title)
    .filter(Boolean);

  const candidateNames = [...new Set([SHEET_NAME, ...allTitles])];

  for (const title of candidateNames) {
    try {
      const headerRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${title}!A1:E1`,
      });

      const headerRow = headerRes.data.values?.[0] || [];
      if (isValidScheduleHeader(headerRow)) {
        resolvedSheetName = title;
        console.log(`Using sheet tab: ${resolvedSheetName}`);
        return resolvedSheetName;
      }
    } catch (err) {
      console.log(`Sheet header check failed for ${title}:`, getGoogleErrorMessage(err));
    }
  }

  resolvedSheetName = SHEET_NAME;
  console.log(`Using fallback sheet tab: ${resolvedSheetName}`);
  return resolvedSheetName;
}

async function verifySheetsAccess() {
  try {
    const activeSheetName = await resolveScheduleSheetName();
    await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${activeSheetName}!A1:A1`,
    });
    console.log("Google Sheets access check: OK");
  } catch (err) {
    console.log("Google Sheets access check failed:", getGoogleErrorMessage(err));
  }
}

// =============================
// GET SHEET DATA
// =============================

async function getSheetData() {
  const activeSheetName = await resolveScheduleSheetName();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${activeSheetName}!A2:E`,
  });

  return res.data.values || [];
}

// =============================
// UPDATE STATUS
// =============================

async function updateStatus(rowNumber, status) {
  const activeSheetName = await resolveScheduleSheetName();
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${activeSheetName}!E${rowNumber}`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[status]],
    },
  });
}

function normalizeDate(value) {
  if (!value) return null;
  const raw = String(value).trim();
  const formats = ["DD-MM-YYYY", "D-M-YYYY", "DD/MM/YYYY", "D/M/YYYY", "YYYY-MM-DD"];

  for (const fmt of formats) {
    const parsed = dayjs(raw, fmt, true);
    if (parsed.isValid()) return parsed.format("DD-MM-YYYY");
  }

  return null;
}

function normalizeTime(value) {
  if (!value) return null;
  const raw = String(value).trim().toUpperCase();
  const formats = ["H:mm", "HH:mm", "h:mm A", "hh:mm A", "h:mmA", "hh:mmA"];

  for (const fmt of formats) {
    const parsed = dayjs(raw, fmt, true);
    if (parsed.isValid()) return parsed.format("HH:mm");
  }

  return null;
}

const DAY_TOKEN_TO_INDEX = {
  SUN: 0,
  SUNDAY: 0,
  MON: 1,
  MONDAY: 1,
  TUE: 2,
  TUESDAY: 2,
  WED: 3,
  WEDNESDAY: 3,
  THU: 4,
  THURSDAY: 4,
  FRI: 5,
  FRIDAY: 5,
  SAT: 6,
  SATURDAY: 6,
};

function normalizeDayToken(token) {
  if (!token) return null;
  const normalized = String(token).trim().toUpperCase().replace(/[^A-Z]/g, "");
  return DAY_TOKEN_TO_INDEX[normalized] ?? null;
}

function buildDayRange(startDay, endDay) {
  if (startDay === null || endDay === null) return null;

  const days = [];
  let current = startDay;

  for (let i = 0; i < 7; i++) {
    days.push(current);
    if (current === endDay) return days;
    current = (current + 1) % 7;
  }

  return null;
}

function parseScheduleDateOrRule(value) {
  if (!value) return null;

  const normalizedDate = normalizeDate(value);
  if (normalizedDate) {
    return { type: "once", date: normalizedDate };
  }

  const raw = String(value).trim().toUpperCase().replace(/\s+/g, "");

  if (["EVERYDAY", "DAILY", "ALL", "*"].includes(raw)) {
    return { type: "recurring", days: [0, 1, 2, 3, 4, 5, 6] };
  }

  if (["WEEKDAYS", "MON-FRI", "MONDAY-FRIDAY"].includes(raw)) {
    return { type: "recurring", days: [1, 2, 3, 4, 5] };
  }

  if (["WEEKENDS", "SAT-SUN", "SATURDAY-SUNDAY"].includes(raw)) {
    return { type: "recurring", days: [6, 0] };
  }

  if (raw.includes("-")) {
    const [startToken, endToken] = raw.split("-");
    const startDay = normalizeDayToken(startToken);
    const endDay = normalizeDayToken(endToken);
    const dayRange = buildDayRange(startDay, endDay);

    if (dayRange) {
      return { type: "recurring", days: dayRange };
    }
  }

  const dayTokens = raw.split(",");
  const parsedDays = dayTokens
    .map((token) => normalizeDayToken(token))
    .filter((day) => day !== null);

  if (parsedDays.length > 0) {
    return { type: "recurring", days: [...new Set(parsedDays)] };
  }

  return null;
}

function getRecurringLastSentDate(statusValue) {
  const raw = String(statusValue || "").trim();
  if (!raw) return null;

  const marker = "LAST_SENT:";
  if (!raw.toUpperCase().startsWith(marker)) return null;

  const datePart = raw.slice(marker.length).trim();
  return normalizeDate(datePart);
}

// =============================
// WHATSAPP BOT
// =============================

async function startBot() {
  console.log(
    `[BOOT] ${BOOT_SIGNATURE} commit=${process.env.RAILWAY_GIT_COMMIT_SHA || "unknown"} node=${process.version}`
  );
  console.log(
    `[BOOT] pairingNumberSet=${WA_PAIRING_NUMBER ? "yes" : "no"} authDir=${AUTH_DIR}`
  );
  console.log(`Pairing mode: ${WA_PAIRING_NUMBER ? "enabled" : "disabled"}`);
  if (!fs.existsSync(AUTH_DIR)) {
    fs.mkdirSync(AUTH_DIR, { recursive: true });
  }
  const { state, saveCreds } = await useMultiFileAuthState(AUTH_DIR);
  const { version, isLatest } = await fetchLatestBaileysVersion();
  console.log(`Using WA Web version ${version.join(".")} (isLatest=${isLatest})`);
  let groupNameToIdCache = new Map();
  let lastGroupCacheAt = 0;
  const GROUP_CACHE_TTL_MS = 5 * 60 * 1000;

  const sock = makeWASocket({
    auth: state,
    version,
    logger: pino({ level: "silent" }),
    browser: Browsers.windows("Desktop"),
    syncFullHistory: false,
    shouldSyncHistoryMessage: () => false,
    markOnlineOnConnect: false,
    connectTimeoutMs: 60_000,
    defaultQueryTimeoutMs: 0,
  });

  sock.ev.on("creds.update", saveCreds);
  let pairingCodeRequested = false;

  async function requestPairingCodeIfNeeded() {
    if (!WA_PAIRING_NUMBER) return;
    if (state.creds.registered) return;
    if (pairingCodeRequested) return;

    pairingCodeRequested = true;
    try {
      const code = await sock.requestPairingCode(WA_PAIRING_NUMBER);
      const formatted = String(code || "")
        .replace(/\s+/g, "")
        .match(/.{1,4}/g)
        ?.join("-") || code;
      console.log(`Pairing code (${WA_PAIRING_NUMBER}): ${formatted}`);
    } catch (err) {
      pairingCodeRequested = false;
      console.log("Pairing code error:", err?.message || err);
    }
  }

  requestPairingCodeIfNeeded().catch((err) => {
    console.log("Pairing bootstrap error:", err?.message || err);
  });

  async function refreshGroupCache(force = false) {
    const nowMs = Date.now();
    if (!force && groupNameToIdCache.size > 0 && nowMs - lastGroupCacheAt < GROUP_CACHE_TTL_MS) {
      return groupNameToIdCache;
    }

    const groups = await sock.groupFetchAllParticipating();
    const nextMap = new Map();

    for (const [jid, meta] of Object.entries(groups || {})) {
      const subject = String(meta?.subject || "").trim();
      if (!subject) continue;
      const key = subject.toLowerCase();
      if (!nextMap.has(key)) {
        nextMap.set(key, jid);
      }
    }

    groupNameToIdCache = nextMap;
    lastGroupCacheAt = nowMs;
    return groupNameToIdCache;
  }

  async function resolveGroupId(groupInput) {
    const raw = String(groupInput || "").trim();
    if (!raw) return null;
    if (raw.endsWith("@g.us")) return raw;

    const groupMap = await refreshGroupCache();
    const lookupKey = raw.toLowerCase();

    if (groupMap.has(lookupKey)) {
      return groupMap.get(lookupKey);
    }

    const partialMatches = [...groupMap.entries()].filter(([name]) => name.includes(lookupKey));
    if (partialMatches.length === 1) {
      return partialMatches[0][1];
    }

    if (partialMatches.length > 1) {
      throw new Error(`Multiple groups matched name "${raw}"`);
    }

    throw new Error(`Group name not found: "${raw}"`);
  }

  sock.ev.on("connection.update", async (update) => {
    const { connection, qr, lastDisconnect } = update;

    if (connection === "connecting") {
      requestPairingCodeIfNeeded().catch((err) => {
        console.log("Pairing request on connect error:", err?.message || err);
      });
    }

    if (qr) {
      console.log("\nScan QR:\n");
      if (!WA_PAIRING_NUMBER) {
        qrcode.generate(qr, { small: true });
      }
    }

    if (connection === "open") {
      console.log("WhatsApp Connected");
      reconnectAttempts = 0;

      if (reconnectTimer) {
        clearTimeout(reconnectTimer);
        reconnectTimer = null;
      }

      if (cronTask) {
        cronTask.stop();
      }

      cronTask = cron.schedule("* * * * *", async () => {
        try {
          const rows = await getSheetData();

          const today = dayjs().format("DD-MM-YYYY");
          const now = dayjs().format("HH:mm");
          const currentDay = dayjs().day();
          console.log(`Checking ${today} ${now}`);

          for (let i = 0; i < rows.length; i++) {
            const [dateOrRule, time, message, groups, status] = rows[i];
            const schedule = parseScheduleDateOrRule(dateOrRule);
            const normalizedTime = normalizeTime(time);
            const normalizedStatus = String(status || "").trim().toUpperCase();

            if (!groups || !message) continue;
            if (!schedule || !normalizedTime) continue;
            if (["PAUSED", "DISABLED", "OFF"].includes(normalizedStatus)) continue;
            if (normalizedTime !== now) continue;

            const isRecurring = schedule.type === "recurring";
            if (!isRecurring && normalizedStatus === "SENT") continue;

            let shouldSend = false;
            if (isRecurring) {
              const lastSentDate = getRecurringLastSentDate(status);
              const alreadySentToday = lastSentDate === today;
              shouldSend = schedule.days.includes(currentDay) && !alreadySentToday;
            } else {
              shouldSend = schedule.date === today;
            }

            if (shouldSend) {
              const groupList = String(groups).split(",");

              for (const groupValue of groupList) {
                try {
                  const groupId = await resolveGroupId(groupValue);
                  if (!groupId) continue;
                  await sock.sendMessage(groupId, { text: message });
                  console.log("Sent:", groupId);
                } catch (err) {
                  console.log("Failed:", String(groupValue || "").trim(), err?.message || err);
                }
              }

              await updateStatus(i + 2, isRecurring ? `LAST_SENT:${today}` : "SENT");
            }
          }
        } catch (err) {
          console.log("Sheet Error:", getGoogleErrorMessage(err));
        }
      });
    }

    if (connection === "close") {
      if (cronTask) {
        cronTask.stop();
        cronTask = null;
      }

      const statusCode = lastDisconnect?.error?.output?.statusCode;
      const isLoggedOut = statusCode === DisconnectReason.loggedOut;
      const isUnregistered = !state.creds.registered;
      const shouldReconnect = !isLoggedOut || isUnregistered;

      console.log(
        `Disconnected: shouldReconnect=${shouldReconnect} statusCode=${statusCode ?? "unknown"} registered=${state.creds.registered}`
      );

      if (shouldReconnect) {
        reconnectAttempts += 1;
        const delay = Math.min(
          RECONNECT_BASE_DELAY_MS * reconnectAttempts,
          RECONNECT_MAX_DELAY_MS
        );
        console.log(`Reconnecting in ${delay / 1000}s...`);
        reconnectTimer = setTimeout(() => {
          startBot().catch((err) => {
            console.log("Restart Error:", err.message);
          });
        }, delay);
      }
    }
  });
}

startBot().catch((err) => {
  console.log("Startup Error:", err.message);
});

verifySheetsAccess().catch((err) => {
  console.log("Sheets verification error:", getGoogleErrorMessage(err));
});

startHealthServer();
