import express from "express";
import path from "path";
// Vite dynamic import
import axios from "axios";
import dotenv from "dotenv";
import fs from "fs";

dotenv.config();

// Gracefully handle uncaught exceptions and unhandled rejections to prevent crashing from client disconnects/socket errors
process.on("uncaughtException", (err: any) => {
  if (err?.code === "EPIPE" || err?.code === "ECONNRESET") {
    // Safe to completely ignore. Browser closed the socket while we were writing.
    return;
  }
  console.error("[PROCESS] Uncaught Exception:", err);
});

process.on("unhandledRejection", (reason: any, promise: Promise<any>) => {
  console.error("[PROCESS] Unhandled Rejection at:", promise, "reason:", reason);
});

const app = express();
const PORT = 3000;

// The Google Apps Script Web App URL provided by the user
const GAS_URL = (
  process.env.GAS_DEPLOYMENT_URL || 
  process.env.GAS_URL || 
  process.env.DEPLOYMENT_URL || 
  "https://script.google.com/macros/s/AKfycbxA0nzMWabxevsaWBoZinNmq7xBJvHcp9JNyQfn4Qs1gVcqlpmSD5yzYQhDofu7xYAl7w/exec"
).trim();
const S_ID    = process.env.SPREADSHEET_ID || process.env.SPREADSHEET_I || process.env.SPREADSHEET;
const SH_NAME = process.env.SHEET_NAME || process.env.SHEET_NAM || process.env.SHEET;

console.log(`[INIT] GAS Proxy target: ${GAS_URL.slice(0, 40)}... (Hidden S_ID: ${!!S_ID})`);

app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ limit: "50mb", extended: true }));

// Safeguard middleware to capture and ignore EPIPE / ECONNRESET errors on response stream
app.use((req, res, next) => {
  res.on("error", (err: any) => {
    if (err.code === "EPIPE" || err.code === "ECONNRESET") {
      console.warn(`[WARNING] Client connection closed prematurely (${err.code}).`);
      return;
    }
    console.error("[ERROR] Response stream error:", err);
  });
  next();
});

/*************** CACHE LAYER ***************/
const CACHE = new Map<string, { data: any, expiry: number }>();
const TTL = {
  OPTIONS: 14400000, // 4 hours
  SYNC: 30000,       // 30 seconds (nearly live check for sheet changes)
  FILTER: 3600000,   // 1 hour
  LOOKUP: 43200000   // 12 hours
};

// In Vercel, only /tmp is writable.
const isVercel = !!process.env.VERCEL;
const OPTIONS_FILE_PATH = path.join(isVercel ? "/tmp" : process.cwd(), "options_cache.json");

// Helper to read options cache from filesystem
function getFileCache(): any | null {
  try {
    if (fs.existsSync(OPTIONS_FILE_PATH)) {
      const stats = fs.statSync(OPTIONS_FILE_PATH);
      const dataStr = fs.readFileSync(OPTIONS_FILE_PATH, "utf8");
      const cached = JSON.parse(dataStr);
      console.log(`[CACHE] Read options file cache successfully. Last updated at: ${stats.mtime}`);
      return { data: cached, mtime: stats.mtimeMs };
    }
  } catch (err: any) {
    console.error("[CACHE] Error reading options file cache:", err.message);
  }
  return null;
}

// Helper to write options cache to filesystem
function setFileCache(data: any) {
  try {
    if (data && data.success) {
      fs.writeFileSync(OPTIONS_FILE_PATH, JSON.stringify(data), "utf8");
      console.log("[CACHE] Options file cache updated on disk.");
    }
  } catch (err: any) {
    console.error("[CACHE] Error writing options file cache to disk:", err.message);
  }
}

/**
 * Prime Time Logic: 8:00 AM to 11:00 PM (Local Time)
 * During these hours, we prioritize cache hits for "instant" feel.
 */
function isPrimeTime() {
  const now = new Date();
  const hour = now.getHours();
  // Adjust to cover 8:00 AM to 11:00 PM (hour < 23 covers up to 22:59:59)
  return hour >= 8 && hour < 23;
}

function getCache(key: string) {
  const entry = CACHE.get(key);
  if (entry && entry.expiry > Date.now()) {
    if (isPrimeTime()) {
      console.log(`[CACHE][PRIME] Instant hit: ${key}`);
    } else {
      console.log(`[CACHE] Hit: ${key}`);
    }
    return entry.data;
  }
  return null;
}

function setCache(key: string, data: any, ttlMs: number) {
  if (!data || !data.success) return; // Don't cache failures
  console.log(`[CACHE] Set: ${key} (ttl: ${ttlMs}ms)`);
  CACHE.set(key, { data, expiry: Date.now() + ttlMs });
}

/**
 * Proxy helper to call the Google Apps Script Web App
 * Handles redirects automatically.
 */
async function callGAS(action: string, payload: any = {}, method: 'GET' | 'POST' = 'POST', timeoutMs: number = 300000) {
  try {
    // We trim the URL in case it has trailing spaces from .env
    const targetUrl = GAS_URL.trim();
    if (!targetUrl.startsWith("https://script.google.com")) {
       return { success: false, error: "Invalid GAS URL. It must start with https://script.google.com" };
    }

    const url = new URL(targetUrl);
    url.searchParams.set("action", action);
    
    // Inject "Hidden" IDs from environment
    if (S_ID) url.searchParams.set("ssId", S_ID);
    if (SH_NAME) url.searchParams.set("sheetName", SH_NAME);

    // For GET requests, we append each payload key as a query parameter
    if (method === 'GET' && payload && typeof payload === 'object') {
      Object.keys(payload).forEach(key => {
        if (payload[key] !== undefined && payload[key] !== null) {
          if (typeof payload[key] === 'object') {
            url.searchParams.set(key, JSON.stringify(payload[key]));
          } else {
            url.searchParams.set(key, String(payload[key]));
          }
        }
      });
    }

    const maskedUrl = `${url.origin}${url.pathname.slice(0, 15)}.../exec?action=${action}`;
    console.log(`[PROXY] Calling GAS [${method}] [${action}]... URL: ${maskedUrl}`);
    
    try {
      const response = await axios({
        method,
        url: url.toString(),
        data: method === 'POST' ? { action, ...payload } : undefined,
        headers: { 
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        maxRedirects: 5,
        validateStatus: () => true,
        timeout: timeoutMs
      });
      
      console.log(`[PROXY] Status: ${response.status} for action: ${action}`);

      // Check if we got HTML (sign of a login redirect or error page)
      if (typeof response.data === 'string' && (response.data.includes('<!DOCTYPE html>') || response.data.includes('<html'))) {
        console.error(`[PROXY] GAS [${action}] returned HTML instead of JSON.`);
        return { 
          success: false, 
          error: "GAS returned a login page. Apps Script MUST be deployed with 'Who has access: Anyone'.",
          advice: "1. Click 'Deploy' > 'New Deployment'. 2. Set 'Who has access' to 'Anyone'. 3. COPY THE NEW URL and update it in AI Studio Secrets."
        };
      }

      if (response.status >= 400) {
         return { 
           success: false, 
           error: `Google returned HTTP ${response.status}`,
           advice: "Check if the script is deleted or the URL is incorrect."
         };
      }

      return response.data;
    } catch (innerError: any) {
      throw innerError;
    }
  } catch (error: any) {
    const errorMsg = error.response 
      ? `Status ${error.response.status}: ${JSON.stringify(error.response.data).slice(0, 500)}` 
      : error.message;
    
    console.error(`[PROXY] GAS Connection Error [${action}]:`, errorMsg);
    
    let advice = "Please check your Google Apps Script deployment URL.";
    if (error.code === 'ECONNABORTED') advice = "The request timed out. The spreadsheet might be too large or GAS is slow.";
    if (error.response?.status === 404) advice = "The deployment URL returned a 404. Is it the correct Exec URL?";
    
    return { 
      success: false, 
      error: `Proxy Error: ${errorMsg}`,
      advice
    };
  }
}

// API Routes - Forwarding to GAS
app.get("/api/ping", async (req, res) => {
  try {
    const data = await callGAS("ping", {}, 'GET');
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/diagnose", async (req, res) => {
  const steps: any[] = [];
  try {
    // Step 1: Check environment variables
    steps.push({
      name: "Check Env Variables",
      status: "success",
      details: `GAS_DEPLOYMENT_URL: ${GAS_URL ? "Configured" : "Missing"}. SPREADSHEET_ID: ${S_ID ? "Configured" : "Using GAS default"}. SHEET_NAME: ${SH_NAME ? "Configured" : "Using GAS default"}.`
    });

    // Step 2: Test URL format
    if (!GAS_URL.startsWith("https://script.google.com")) {
      steps.push({
        name: "Validate Web App URL format",
        status: "error",
        details: "URL must start with https://script.google.com"
      });
      return res.json({ success: false, steps });
    }
    steps.push({
      name: "Validate Web App URL format",
      status: "success",
      details: "URL starts with https://script.google.com"
    });

    // Step 3: Call Ping
    try {
      console.log("[DIAGNOSE] Pinging GAS...");
      const pingRes = await callGAS("ping", {}, 'GET', 15000);
      if (pingRes && pingRes.success && pingRes.pong) {
        steps.push({
          name: "Reach Apps Script Web App",
          status: "success",
          details: `Successfully connected. GAS Version: ${pingRes.version || 'unknown'}`
        });
      } else {
        const errStr = pingRes?.error || JSON.stringify(pingRes);
        steps.push({
          name: "Reach Apps Script Web App",
          status: "error",
          details: `Connected but got unsuccessful response: ${errStr}`,
          advice: pingRes?.advice || "Is the Web App deployed correctly with Who has access: Anyone? Make sure to authorize the script by choosing a function in Apps Script editor and clicking 'Run'!"
        });
        return res.json({ success: false, steps });
      }
    } catch (err: any) {
      steps.push({
        name: "Reach Apps Script Web App",
        status: "error",
        details: `Connection timeout or error: ${err.message}`,
        advice: "The Web App URL might be incorrect or not published yet. Please click 'Deploy' > 'New Deployment', make sure 'Who has access' is 'Anyone' and 'Execute as' is 'Me'. Copy the NEW Web App URL and update it in AI Studio Secrets."
      });
      return res.json({ success: false, steps });
    }

    // Step 4: Test Sheet Access / Count
    try {
      console.log("[DIAGNOSE] Checking Sheet access...");
      const syncRes = await callGAS("sync", {}, 'GET', 15000);
      if (syncRes && syncRes.success) {
        steps.push({
          name: "Read Spreadsheet & Sheet Access",
          status: "success",
          details: `Successfully connected to spreadsheet! Row count: ${syncRes.rowCount}`
        });
      } else {
        steps.push({
          name: "Read Spreadsheet & Sheet Access",
          status: "error",
          details: `Error: ${syncRes.error || "Failed to access spreadsheet"}.`,
          advice: "Please check if your Spreadsheet ID is correct and the sheet Name is 'Examiner Information'. Also, make sure the Google Apps Script has been authorized by clicking Run on any function in the Apps Script editor!"
        });
        return res.json({ success: false, steps });
      }
    } catch (err: any) {
      steps.push({
        name: "Read Spreadsheet & Sheet Access",
        status: "error",
        details: `Failed to test sheet access: ${err.message}`,
        advice: "This usually means the Apps Script script does not have permission to access the Spreadsheet. Go to the Apps Script editor, select a function (e.g. 'testSheetAccess'), click 'Run' to authorize, then retry!"
      });
      return res.json({ success: false, steps });
    }

    // All steps successful!
    return res.json({ success: true, steps });
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message, steps });
  }
});

app.get("/api/options", async (req, res) => {
  try {
    const cacheKey = "options";
    
    // 1. Check memory cache first for ultra-fast instant hits
    const cachedMemory = getCache(cacheKey);
    if (cachedMemory) {
      console.log("[API] serving options from memory cache");
      return res.json(cachedMemory);
    }

    // 2. Check file cache next to persist across reboots/restarts
    const fileCache = getFileCache();
    if (fileCache) {
      const { data, mtime } = fileCache;
      const ageMs = Date.now() - mtime;
      // Stale threshold is 15 minutes (900,000ms) during Prime Time (8 AM to 11 PM), else 4 hours (TTL.OPTIONS)
      const maxAge = isPrimeTime() ? 900000 : TTL.OPTIONS;
      const isStale = ageMs > maxAge;

      console.log(`[API] serving options from file cache (age: ${Math.round(ageMs/1000/60)} minutes, isStale: ${isStale}, prime: ${isPrimeTime()})`);
      
      // Serve file cache instantly so the user NEVER has to wait
      res.json(data);

      // Warm memory cache for subsequent requests
      setCache(cacheKey, data, TTL.OPTIONS);

      // Revalidate in the background if stale
      if (isStale) {
        console.log("[API] Options cache is stale, revalidating in background...");
        (async () => {
          try {
            const freshData = await callGAS("options", {}, "GET");
            if (freshData && freshData.success) {
              setCache(cacheKey, freshData, TTL.OPTIONS);
              setFileCache(freshData);
              console.log("[API] Options cache revalidated successfully in background.");
            }
          } catch (bgErr: any) {
            console.error("[API] Background options revalidation failed:", bgErr.message);
          }
        })();
      }
      return;
    }

    // 3. Fallback: No cache found. Fetch synchronously.
    console.log("[API] No cache found. Fetching options synchronously from GAS...");
    const data = await callGAS("options", {}, 'GET');
    if (data && data.success) {
      setCache(cacheKey, data, TTL.OPTIONS);
      setFileCache(data);
    }
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post("/api/filter", async (req, res) => {
  try {
    const cacheKey = `filter_${JSON.stringify(req.body)}`;
    const cached = getCache(cacheKey);
    if (cached) return res.json(cached);

    const data = await callGAS("filter", req.body, 'POST');
    setCache(cacheKey, data, TTL.FILTER);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/sync", async (req, res) => {
  try {
    const cacheKey = "sync";
    const cached = getCache(cacheKey);
    if (cached) return res.json(cached);

    const data = await callGAS("sync", {}, 'GET');
    setCache(cacheKey, data, TTL.SYNC);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/all-data", async (req, res) => {
  try {
    const cacheKey = "all_data";
    const cached = getCache(cacheKey);
    if (cached) return res.json(cached);

    console.log("[API] Fetching ALL data from GAS...");
    const data = await callGAS("filter", { 
      filters: {
        institute: [],
        department: [],
        batch: [],
        trainingsSelected: [],
        trainingDatesSelected: [],
        campusesSelected: [],
        tpinsSelected: [],
        subjectsSelected: [],
        onlyAllowed: false,
        subjectLogic: "any",
        allowEnglish: 55,
        allowOthers: 48
      }, 
      page: 1, 
      pageSize: 200000 
    }, 'POST', 600000); // 10 min timeout for bulk export
    
    if (data && data.success) {
      setCache(cacheKey, data, TTL.FILTER);
    }
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/lookup", async (req, res) => {
  try {
    const query = req.query.query as string;
    if (!query) return res.json({ success: true, found: false });

    const cacheKey = `lookup_${query.toLowerCase()}`;
    const cached = getCache(cacheKey);
    if (cached) return res.json(cached);

    const data = await callGAS("lookup", { query }, 'GET');
    setCache(cacheKey, data, TTL.LOOKUP);
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/clearCache", async (req, res) => {
  try {
    console.log("[CACHE] Clearing all local cache");
    CACHE.clear();
    
    // Also delete filesystem cache
    if (fs.existsSync(OPTIONS_FILE_PATH)) {
      try {
        fs.unlinkSync(OPTIONS_FILE_PATH);
        console.log("[CACHE] Deleted options file cache");
      } catch (err: any) {
        console.error("[CACHE] Failed to delete options file cache:", err.message);
      }
    }

    const data = await callGAS("clearCache", {}, 'GET');
    res.json(data);
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// Vite Middleware for Development / Static serving for Production
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  const server = app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
    console.log(`GAS Proxy active to: ${GAS_URL}`);
  });

  // Handle client error socket events safely to prevent EPIPE/ECONNRESET crashes
  server.on("clientError", (err: any, socket: any) => {
    if (err.code === "EPIPE" || !socket.writable) {
      return;
    }
    socket.end("HTTP/1.1 400 Bad Request\r\n\r\n");
  });
}

if (!process.env.VERCEL) {
  startServer();
}

export default app;
