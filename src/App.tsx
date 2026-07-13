import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { 
  Users, 
  Search, 
  RefreshCcw, 
  Printer, 
  X, 
  AlertCircle,
  FileSpreadsheet,
  ChevronLeft,
  ChevronRight,
  Loader2,
  CheckCircle2,
  ArrowRight,
  Database,
  RefreshCw,
  ShieldAlert,
  FileDown,
  Circle
} from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import axios from "axios";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { useAuth, AuthProvider } from "./components/auth-context";
import { FilterDropdown } from "./components/FilterDropdown";
import { cn } from "./lib/utils";
import { dbStorage } from "./lib/db";

// Types
interface Subject {
  key: string;
  label: string;
}

interface FilterOptions {
  institutes: string[];
  departments: string[];
  batches: string[];
  trainings: string[];
  trainingDates: string[];
  campuses: string[];
  tpins: string[];
  subjects: Subject[];
  rowCount: number;
  rows?: string[][];
  header?: string[];
}

interface SearchFilters {
  institute: string[];
  department: string[];
  batch: string[];
  trainingsSelected: string[];
  trainingDatesSelected: string[];
  campusesSelected: string[];
  tpinsSelected: string[];
  subjectsSelected: string[];
  onlyAllowed: boolean;
  subjectLogic: "all" | "any";
  allowEnglish: number | null;
  allowOthers: number | null;
}

interface FilterResult {
  header: string[];
  rows: string[][];
  total: number;
  page: number;
  pageSize: number;
  totalPages: number;
  allow: { ENGLISH: number; OTHERS: number };
}

const DEFAULT_FILTERS: SearchFilters = {
  institute: [],
  department: [],
  batch: [],
  trainingsSelected: [],
  trainingDatesSelected: [],
  campusesSelected: [],
  tpinsSelected: [],
  subjectsSelected: [],
  onlyAllowed: true,
  subjectLogic: "any",
  allowEnglish: 55,
  allowOthers: 48,
};

const Dashboard: React.FC = () => {
  const { user, token, initialized, login, logout, isLoggingIn } = useAuth();
  const [options, setOptions] = useState<FilterOptions | null>(() => {
    try {
      const savedMeta = localStorage.getItem("ex_options_meta");
      if (savedMeta) {
        const meta = JSON.parse(savedMeta);
        console.log("[App] Instant synchronous load of options metadata from localStorage. Row count:", meta.rowCount);
        return {
          institutes: [],
          departments: [],
          batches: [],
          trainings: [],
          trainingDates: [],
          campuses: [],
          tpins: [],
          subjects: [],
          rowCount: meta.rowCount || 47202
        };
      }
    } catch (e) {
      console.warn("Error reading options from localStorage on load", e);
    }
    return null;
  });

  const filteredBatches = useMemo(() => {
    if (!options?.batches) return [];
    return options.batches.filter(b => {
      const s = String(b).trim();
      return s !== "" && /^\d+$/.test(s);
    });
  }, [options?.batches]);
  const [isInstantLoaded, setIsInstantLoaded] = useState(() => {
    try {
      const savedMeta = localStorage.getItem("ex_options_meta");
      if (savedMeta) {
        const meta = JSON.parse(savedMeta);
        return !!meta.hasData;
      }
    } catch {
      return false;
    }
    return false;
  });
  const [filters, setFilters] = useState<SearchFilters>(() => {
    try {
      const saved = localStorage.getItem("ex_filters");
      return saved ? { ...DEFAULT_FILTERS, ...JSON.parse(saved) } : DEFAULT_FILTERS;
    } catch { return DEFAULT_FILTERS; }
  });

  const [result, setResult] = useState<FilterResult | null>(null);
  const [fullResult, setFullResult] = useState<FilterResult | null>(null);
  
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isBgSyncing, setIsBgSyncing] = useState(false);
  const [latency, setLatency] = useState<string>("Cached (Disk)");
  const [isOnline, setIsOnline] = useState<boolean>(typeof navigator !== "undefined" ? navigator.onLine : true);
  const [searchLatency, setSearchLatency] = useState<string>("");
  const [showSyncToast, setShowSyncToast] = useState<boolean>(false);
  const [syncToastMessage, setSyncToastMessage] = useState<string>("");
  const fetchAbortControllerRef = useRef<AbortController | null>(null);

  // Connection online/offline handlers
  useEffect(() => {
    if (typeof window === "undefined") return;
    const handleOnline = () => {
      setIsOnline(true);
      setLatency("Reconnected (Live)");
      // Automatically pull latest on reconnection
      loadOptions(false, true, false);
    };
    const handleOffline = () => {
      setIsOnline(false);
      setLatency("Offline Mode");
    };
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, []);
  
  // Async IndexedDB initialization effect
  useEffect(() => {
    const initDbData = async () => {
      try {
        const startTime = performance.now();
        console.log("[App] Reading cached options and results from IndexedDB...");
        const cachedOptions = await dbStorage.getItem<FilterOptions>("ex_options");
        if (cachedOptions) {
          const duration = Math.round(performance.now() - startTime);
          console.log(`[App] Loaded cached options from IndexedDB in ${duration}ms, row count:`, cachedOptions.rowCount);
          setLatency(`Cached (${duration}ms)`);
          setOptions(cachedOptions);
          setIsInstantLoaded(true);
          setIsLoading(false);
        }
        
        const cachedResult = await dbStorage.getItem<FilterResult>("ex_last_res");
        if (cachedResult) {
          console.log("[App] Loaded cached search results from IndexedDB, total:", cachedResult.total);
          setResult(cachedResult);
        }
      } catch (err) {
        console.error("[App] Failed to load from IndexedDB:", err);
      }
    };
    initDbData();
  }, []);

  // Persistence effect
  useEffect(() => {
    localStorage.setItem("ex_filters", JSON.stringify(filters));
  }, [filters]);
  const [errorDetails, setErrorDetails] = useState<{ message: string; advice?: string } | null>(null);
  const [diagnostics, setDiagnostics] = useState<any[] | null>(null);
  const [isDiagnosing, setIsDiagnosing] = useState<boolean>(false);

  const runDiagnostics = async () => {
    setIsDiagnosing(true);
    setDiagnostics([]);
    try {
      const res = await axios.get("/api/diagnose");
      if (res.data && res.data.steps) {
        setDiagnostics(res.data.steps);
      } else {
        setDiagnostics([{ name: "Run Diagnostics Call", status: "error", details: "Unable to parse diagnostic response from the server." }]);
      }
    } catch (err: any) {
      setDiagnostics([{ name: "Run Diagnostics Call", status: "error", details: err.message || "Failed to call backend diagnostics." }]);
    } finally {
      setIsDiagnosing(false);
    }
  };

  const [searchQuery, setSearchQuery] = useState("");
  
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(200);
  const [autoRefresh, setAutoRefresh] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState<Date>(() => {
    try {
      const saved = localStorage.getItem("ex_last_sync_time");
      return saved ? new Date(saved) : new Date();
    } catch {
      return new Date();
    }
  });
  const [hasSheetUpdates, setHasSheetUpdates] = useState(false);

  // Initialize data
  useEffect(() => {
    if (initialized) {
      console.log("[App] Component initialized, checking sync status...");
      
      const lastSync = localStorage.getItem("ex_last_sync_time");
      let shouldSync = true;
      
      if (lastSync) {
        const lastSyncDate = new Date(lastSync);
        const now = new Date();
        const isToday = lastSyncDate.toDateString() === now.toDateString();
        const hour = now.getHours();
        const isWithinHours = hour >= 8 && hour < 23;
        
        // If it was synced today within active hours, avoid a heavy sync
        if (isToday && isWithinHours) {
            shouldSync = false; 
        }
      }
      
      // If we need to sync, OR if we don't have instant loaded data yet, we loadOptions.
      // If we don't need to sync and we ALREADY have data, we SKIP loadOptions.
      
      if (shouldSync || !isInstantLoaded) {
          loadOptions(false, !shouldSync, true);
      }
      
      const cleanup = startPolling();
      return cleanup;
    }
  }, [initialized, token, isInstantLoaded]);

  // Polling for sheet updates
  const startPolling = () => {
    const timer = setInterval(async () => {
      if (!token || !initialized) return;
      try {
        const res = await axios.get("/api/sync", {
          headers: token ? { Authorization: `Bearer ${token}` } : {}
        });
        if (res.data.success) {
          const now = new Date();
          setLastSyncTime(now);
          localStorage.setItem("ex_last_sync_time", now.toISOString());
          if (options && res.data.rowCount !== options.rowCount) {
            setHasSheetUpdates(true);
            console.log("[App] Sheet rowCount changed from", options.rowCount, "to", res.data.rowCount, ". Reloading options silently...");
            loadOptions(true, true, false); // Force refresh options silently in background
            if (autoRefresh) handleSearch();
          }
        }
      } catch (err) {
        console.warn("[App] Sync/Polling error:", err);
      }
    }, 60000); // Check every 60s
    return () => clearInterval(timer);
  };

  // Background interval auto-sync (every 5 minutes, between 8 AM and 11 PM)
  useEffect(() => {
    if (!initialized || !token) return;
    
    const interval = setInterval(() => {
      const now = new Date();
      const hour = now.getHours();
      // Only sync automatically between 8:00 AM and 11:00 PM (hour < 23 covers up to 22:59)
      if (hour >= 8 && hour < 23) {
        console.log("[App] Active hours auto-sync: fetching updated data in background...");
        loadOptions(false, true, false); // Silent background update
      }
    }, 300000); // 5 minutes
    
    return () => clearInterval(interval);
  }, [initialized, token, options]);

  const [isPinged, setIsPinged] = useState<boolean | null>(null);

  const checkPing = async () => {
    try {
      const res = await axios.get("/api/ping");
      const ok = !!(res.data && res.data.success);
      setIsPinged(ok);
      console.log("[App] Backend Ping:", ok ? "Online" : "Offline");
    } catch (err: any) {
      setIsPinged(false);
      console.error("[App] Ping failed:", err.message);
    }
  };

  const loadOptions = async (forceServerRefresh = false, isSilent = false, showToastOnSuccess = false) => {
    // If we have options, and the caller requested silent mode, or if we didn't specify and options exist, default to silent
    const finalSilent = isSilent || (!!options && !forceServerRefresh && !errorDetails);
    
    if (finalSilent) {
      if (isBgSyncing) return;
      setIsBgSyncing(true);
    } else {
      if (isSyncing) return;
      setIsSyncing(true);
    }
    
    setErrorDetails(null);
    console.log(`[App] Fetching filter options (force: ${forceServerRefresh}, isSilent: ${finalSilent})...`);
    
    // Non-blocking background checks
    checkPing();

    const fetchStartTime = performance.now();
    try {
      if (forceServerRefresh) {
        await axios.get("/api/clearCache", {
          headers: token ? { Authorization: `Bearer ${token}` } : {}
        });
      }

      const res = await axios.get("/api/options", {
        headers: token ? { Authorization: `Bearer ${token}` } : {},
        timeout: 150000 
      });
      
      if (res.data && res.data.success) {
        // Fetch all data for local caching so search is instant
        try {
          const allDataRes = await axios.get("/api/all-data", {
            headers: token ? { Authorization: `Bearer ${token}` } : {},
            timeout: 600000 
          });
          if (allDataRes.data && allDataRes.data.success) {
            res.data.rows = allDataRes.data.rows;
            res.data.header = allDataRes.data.header;
          }
        } catch (e) {
          console.warn("[App] Failed to fetch all data for caching:", e);
        }

        const fetchDuration = Math.round(performance.now() - fetchStartTime);
        setLatency(`Live (${fetchDuration}ms)`);
        
        // Show silent update toast to amaze the user with background capability!
        if (showToastOnSuccess) {
          if (finalSilent && options && options.rowCount !== res.data.rowCount) {
            setSyncToastMessage(`Database updated! Cached ${res.data.rowCount} records in background.`);
            setShowSyncToast(true);
            setTimeout(() => setShowSyncToast(false), 5000);
          } else if (finalSilent && (!options || options.rowCount === res.data.rowCount)) {
            // Just verified
            setSyncToastMessage(`Verified live sheet data. Server response: ${fetchDuration}ms.`);
            setShowSyncToast(true);
            setTimeout(() => setShowSyncToast(false), 3000);
          } else if (!finalSilent) {
            setSyncToastMessage(`Successfully synced database with Google Sheets. Cached ${res.data.rowCount} records.`);
            setShowSyncToast(true);
            setTimeout(() => setShowSyncToast(false), 4000);
          }
        }

        setOptions(res.data);
        setIsInstantLoaded(true);
        
        // Save to IndexedDB to completely avoid any localStorage 5MB size limit
        await dbStorage.setItem("ex_options", res.data);
        
        // Save minimal options metadata to localStorage for synchronous instant load
        // Do NOT save the full res.data to localStorage to avoid QuotaExceededError
        localStorage.setItem("ex_options_meta", JSON.stringify({
          rowCount: res.data.rowCount,
          hasData: true
        }));
        
        setHasSheetUpdates(false);
        if (res.data.rowCount !== undefined) {
          const now = new Date();
          setLastSyncTime(now);
          localStorage.setItem("ex_last_sync_time", now.toISOString());
        }
      } else {
        if (!finalSilent) {
          let msg = res.data?.error || "The backend returned an unsuccessful response.";
          if (typeof msg === "object") msg = JSON.stringify(msg);
          setErrorDetails({ 
            message: msg, 
            advice: res.data?.advice 
          });
        } else {
          console.warn("[App] Silent options fetch returned unsuccessful response:", res.data);
        }
      }
    } catch (err: any) {
      if (!finalSilent) {
        let msg = err.response?.data?.error || err.message || "Failed to reach the backend server.";
        if (typeof msg === "object") msg = JSON.stringify(msg);
        let adv = err.response?.data?.advice;
        if (typeof adv === "object") adv = JSON.stringify(adv);
        setErrorDetails({ 
          message: msg, 
          advice: adv 
        });
      } else {
        console.warn("[App] Silent options fetch failed:", err);
      }
    } finally {
      if (finalSilent) {
        setIsBgSyncing(false);
      } else {
        setIsSyncing(false);
      }
      setIsLoading(false);
    }
  };

  const handleLookup = useCallback(async (query: string) => {
    if (!query || !token) return;
    setIsLoading(true);
    setErrorDetails(null);
    const startTime = performance.now();
    try {
      const res = await axios.get(`/api/lookup?query=${encodeURIComponent(query)}`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (res.data.success && res.data.found) {
        const duration = Math.round(performance.now() - startTime);
        setSearchLatency(`${duration}ms`);
        setResult({
          header: res.data.header,
          rows: [res.data.row],
          total: 1,
          page: 1,
          pageSize: 1,
          totalPages: 1,
          allow: { ENGLISH: 55, OTHERS: 48 }
        });
      } else {
        alert("No examiner found with that ID or Mobile number.");
      }
    } catch (err: any) {
      console.error("Lookup failed", err);
      let msg = err.response?.data?.error || err.message || "Lookup failed";
      if (typeof msg === "object") msg = JSON.stringify(msg);
      setErrorDetails({ 
        message: msg, 
        advice: "Check your connection and try again." 
      });
    } finally {
      setIsLoading(false);
    }
  }, [token]);

  const handleSearch = useCallback(async (newPage = 1, isSilent = false) => {
    // Validation: Require at least one specific filter or search query
    const hasSpecificFilter = 
      filters.institute.length > 0 ||
      filters.department.length > 0 ||
      filters.batch.length > 0 ||
      filters.trainingsSelected.length > 0 ||
      filters.campusesSelected.length > 0 ||
      filters.tpinsSelected.length > 0 ||
      (filters.subjectsSelected && filters.subjectsSelected.length > 0) ||
      searchQuery.trim().length > 0;

    if (!hasSpecificFilter && !isSilent) {
      if (result) {
        setResult(null); // Clear previous results if everything is deselected
        setFullResult(null);
      }
      return;
    } else if (!hasSpecificFilter && isSilent) {
      setResult(null);
      setFullResult(null);
      return;
    }

    // Remove strict token requirement as the backend proxy doesn't strictly enforce it for now 
    // and we want to avoid silent failures on page refresh if token isn't restored.
    // We still log if it's missing.
    if (!token) {
      console.warn("[App] Searching without token (might be after refresh)");
    }
    
    if (searchQuery.trim() && newPage === 1) {
      handleLookup(searchQuery.trim());
      return;
    }

    if (!isSilent) setIsLoading(true);
    setHasSheetUpdates(false);
    setErrorDetails(null);

    // Cancel pending requests to prevent 502 overloading
    if (fetchAbortControllerRef.current) {
      fetchAbortControllerRef.current.abort();
    }
    fetchAbortControllerRef.current = new AbortController();

    const startTime = performance.now();
    try {
      console.log("[App] Executing search for page:", newPage);
      
      // If we have local rows cached, filter locally for INSTANT results!
      if (options?.rows && options.rows.length > 0) {
        console.log("[App] Filtering locally from", options.rows.length, "cached rows...");
        const allSubjects = ["english", "bangla", "physics", "chemistry", "math", "biology", "ict"];
        
        const COL_TPIN = options.header.indexOf('T-PIN');
        const COL_INST = options.header.indexOf('Inst.');
        const COL_DEPT = options.header.indexOf('Dept.');
        const COL_BATCH = options.header.indexOf('HSC Batch');
        const COL_MOB1 = options.header.indexOf('Mobile Number');
        const COL_ALT = options.header.indexOf('Alternate');
        const COL_TRAIN = options.header.indexOf('Training Report');
        const COL_TRAIN_DATE = options.header.indexOf('Training Date');
        const COL_CAMPUS = options.header.indexOf('Physical Campus');
        
        const subjCols: Record<string, number> = {
           english: options.header.indexOf('English(%)'),
           bangla: options.header.indexOf('Bangla(%)'),
           physics: options.header.indexOf('Physics(%)'),
           chemistry: options.header.indexOf('Chemistry(%)'),
           math: options.header.indexOf('Math(%)'),
           biology: options.header.indexOf('Biology(%)'),
           ict: options.header.indexOf('ICT(%)')
        };

        const allowEnglish = filters.allowEnglish !== null ? filters.allowEnglish : 55;
        const allowOthers = filters.allowOthers !== null ? filters.allowOthers : 48;
        const allowThreshold = (k: string) => k === 'english' ? allowEnglish : allowOthers;

        const subjectKeys = (filters.subjectsSelected && filters.subjectsSelected.length > 0) 
            ? filters.subjectsSelected 
            : allSubjects;
            
        const isAllowed = (row: string[]) => {
           if (filters.subjectLogic === "all") {
             for (const k of subjectKeys) {
               const val = parseFloat(row[subjCols[k]]);
               if (isNaN(val) || val < allowThreshold(k)) return false;
             }
             return true;
           } else {
             for (const k of subjectKeys) {
               const val = parseFloat(row[subjCols[k]]);
               if (!isNaN(val) && val >= allowThreshold(k)) return true;
             }
             return false;
           }
        };

        const normalizeStr = (s: string) => String(s || "").replace(/[\u200B-\u200D\uFEFF]/g, '').trim().replace(/\s+/g, ' ').toLowerCase();
        
        const instSet = new Set(filters.institute.map(normalizeStr));
        const deptSet = new Set(filters.department.map(normalizeStr));
        const batchSet = new Set(filters.batch.map(normalizeStr));
        const trainSet = new Set(filters.trainingsSelected.map(normalizeStr));
        const tDateSet = new Set(filters.trainingDatesSelected.map(normalizeStr));
        const campusSet = new Set(filters.campusesSelected.map(normalizeStr));
        const tpinSet = new Set(filters.tpinsSelected.map(normalizeStr));
        
        const q = normalizeStr(searchQuery);

        const filteredRows = options.rows.filter(r => {
           // TPIN, Mobile 1, or Alternate Mobile
           if (q.length > 0) {
             const tpinMatch = (COL_TPIN !== -1) && normalizeStr(r[COL_TPIN]).includes(q);
             const mob1Match = (COL_MOB1 !== -1) && normalizeStr(r[COL_MOB1]).includes(q);
             const altMatch = (COL_ALT !== -1) && normalizeStr(r[COL_ALT]).includes(q);
             if (!tpinMatch && !mob1Match && !altMatch) return false;
           }

           if (instSet.size > 0 && (COL_INST === -1 || !instSet.has(normalizeStr(r[COL_INST])))) return false;
           if (deptSet.size > 0 && (COL_DEPT === -1 || !deptSet.has(normalizeStr(r[COL_DEPT])))) return false;
           if (batchSet.size > 0 && (COL_BATCH === -1 || !batchSet.has(normalizeStr(r[COL_BATCH])))) return false;
           
           if (COL_TRAIN !== -1) {
             const trn = normalizeStr(r[COL_TRAIN]);
             const normalizedTrn = (trn === "" || trn === "(blank)") ? "__blank__" : trn;
             if (trainSet.size > 0 && !trainSet.has(normalizedTrn)) return false;
           }
           
           if (COL_TRAIN_DATE !== -1) {
             const tDate = normalizeStr(r[COL_TRAIN_DATE]);
             const normalizedTDate = (tDate === "" || tDate === "(blank)") ? "__blank__" : tDate;
             if (tDateSet.size > 0 && !tDateSet.has(normalizedTDate)) return false;
           }
           
           if (COL_CAMPUS !== -1) {
             const cam = normalizeStr(r[COL_CAMPUS]);
             const normalizedCam = (cam === "" || cam === "(blank)") ? "__blank__" : cam;
             if (campusSet.size > 0 && !campusSet.has(normalizedCam)) return false;
           }

           if (tpinSet.size > 0) {
              const tpin = normalizeStr(r[COL_TPIN]);
              const normalizedTpin = (tpin === "" || tpin === "(blank)") ? "__blank__" : tpin;
              if (COL_TPIN === -1 || !tpinSet.has(normalizedTpin)) return false;
           }
           
           if (filters.onlyAllowed && !isAllowed(r)) return false;
           
           return true;
        });

        // Debugging logs if needed for troubleshooting
        // console.log("[App] Total rows:", options.rows.length, "Filtered rows:", filteredRows.length);
        // if (filteredRows.length === 0 && options.rows.length > 0) {
        //    console.log("[App] Sample row to check:", options.rows[0]);
        //    console.log("[App] Current Filters:", filters);
        // }

        const startIndex = (newPage - 1) * pageSize;
        const paginatedRows = filteredRows.slice(startIndex, startIndex + pageSize);

        const duration = Math.round(performance.now() - startTime);
        setSearchLatency(`${duration}ms (Local Cache)`);
        
        const resData = {
           success: true,
           header: options.header || [],
           rows: paginatedRows,
           total: filteredRows.length,
           page: newPage,
           pageSize,
           totalPages: Math.ceil(filteredRows.length / pageSize),
           allow: { ENGLISH: allowEnglish, OTHERS: allowOthers }
        };
        
        setResult(resData);
        dbStorage.setItem("ex_last_res", resData);
        setPage(newPage);
        setFullResult({ ...resData, rows: filteredRows, page: 1, totalPages: 1 });

        if (!isSilent) {
          setTimeout(() => {
            const el = document.getElementById("results-section");
            if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }, 50);
        }
        setIsLoading(false);
        return;
      }

      const res = await axios.post("/api/filter", {
        filters,
        page: newPage,
        pageSize,
        returnAll: false
      }, {
        headers: token ? { Authorization: `Bearer ${token}` } : {},
        timeout: 120000,
        signal: fetchAbortControllerRef.current.signal
      });

      if (res.data && res.data.success) {
        const duration = Math.round(performance.now() - startTime);
        setSearchLatency(`${duration}ms`);
        setResult(res.data);
        dbStorage.setItem("ex_last_res", res.data);
        setPage(newPage);
        
        // Phase 2: Background fetch all for export is removed to prevent GAS concurrency locking and 30-40s delays.
        // Full export can be done on-demand or use currently loaded results.
        setFullResult(null);

        // Scroll to results if not silent
        if (!isSilent) {
          setTimeout(() => {
            const el = document.getElementById("results-section");
            if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }, 100);
        }
      } else {
        let msg = res.data?.error || "Search failed";
        if (typeof msg === "object") msg = JSON.stringify(msg);
        setErrorDetails({ 
          message: msg, 
          advice: res.data?.advice || "Please check your connection and script deployment."
        });
      }
    } catch (err: any) {
      if (axios.isCancel(err)) {
        console.log("[App] Search cancelled due to new request.");
        return;
      }
      console.error("[App] Search error:", err);
      
      let errMsg = err.response?.data?.error || err.message || "Search failed";
      if (typeof errMsg === "object") errMsg = JSON.stringify(errMsg);
      
      // Improve 502 / overload messaging so it doesn't look like a total crash
      setErrorDetails({ 
        message: typeof errMsg === "string" && errMsg.includes("502") ? "Backend Service Timeout (502)" : errMsg, 
        advice: (typeof errMsg === "string" && errMsg.includes("502")) || err.code === "ECONNABORTED" 
          ? "The server is overloaded from too many requests. Please wait 1-2 minutes and press 'Search' again." 
          : "The backend might be overloaded or the script is unavailable. Try again in a moment."
      });
    } finally {
      if (!fetchAbortControllerRef.current?.signal.aborted) {
        setIsLoading(false);
      }
    }
  }, [token, filters, pageSize, searchQuery, handleLookup]);

  const handleClear = () => {
    setFilters(DEFAULT_FILTERS);
    setSearchQuery("");
    setResult(null);
    setFullResult(null);
    setPage(1);
    setHasSheetUpdates(false);
    dbStorage.removeItem("ex_last_res");
    localStorage.removeItem("ex_filters");
  };

  // Clear results if all filters and search query are removed
  useEffect(() => {
    const hasSpecificFilter = 
      filters.institute.length > 0 ||
      filters.department.length > 0 ||
      filters.batch.length > 0 ||
      filters.trainingsSelected.length > 0 ||
      filters.campusesSelected.length > 0 ||
      filters.tpinsSelected.length > 0 ||
      (filters.subjectsSelected && filters.subjectsSelected.length > 0);
      
    if (!hasSpecificFilter && searchQuery.trim().length === 0) {
       setResult(null);
       setFullResult(null);
       setPage(1);
    }
  }, [filters, searchQuery]);

  const handleExportExcel = () => {
    const dataToExport = fullResult || result;
    if (!dataToExport) return;

    // Transform data for Excel (similar to PDF logic if needed, but usually Excel users want raw data)
    // However, if we want to match the "Comment" column request for Excel too:
    const excelHeader = Array.isArray(dataToExport.header) ? [...dataToExport.header] : [];
    const statusIdx = excelHeader.findIndex(h => h && h.toLowerCase().includes("allow status"));
    if (statusIdx !== -1) excelHeader[statusIdx] = "Comment";

    const excelRows = Array.isArray(dataToExport.rows) ? dataToExport.rows.map(row => {
      const newRow = Array.isArray(row) ? [...row] : [];
      if (statusIdx !== -1 && newRow.length > statusIdx) newRow[statusIdx] = ""; 
      return newRow;
    }) : [];

    const worksheet = XLSX.utils.aoa_to_sheet([excelHeader, ...excelRows]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Examiners");
    
    XLSX.writeFile(workbook, `Examiner_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleExportPDF = () => {
    const dataToExport = fullResult || result;
    if (!dataToExport) return;

    // Transform data for printing as per user request:
    // 1. Rename "Allow Status" to "Comment"
    // 2. Clear values in that column (make boxes empty)
    const printHeader = Array.isArray(dataToExport.header) ? [...dataToExport.header] : [];
    const statusIdx = printHeader.findIndex(h => h && h.toLowerCase().includes("allow status"));
    
    if (statusIdx !== -1) {
      printHeader[statusIdx] = "Comment";
    }

    const printRows = Array.isArray(dataToExport.rows) ? dataToExport.rows.map(row => {
      const newRow = Array.isArray(row) ? [...row] : [];
      if (statusIdx !== -1 && newRow.length > statusIdx) {
        newRow[statusIdx] = ""; // Clear content
      }
      return newRow;
    }) : [];

    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const pageW = doc.internal.pageSize.getWidth();
    
    // Title
    doc.setFontSize(14);
    doc.setTextColor(26, 86, 219);
    doc.text("Examiner Filter Report", pageW / 2, 10, { align: "center" });
    
    // Summary
    doc.setFontSize(9);
    doc.setTextColor(100);
    doc.text(`Found ${dataToExport.total} records | Generated: ${new Date().toLocaleString()}`, pageW / 2, 16, { align: "center" });
    
    autoTable(doc, {
      head: [printHeader],
      body: printRows,
      startY: 22,
      styles: { fontSize: 7, cellPadding: 1, halign: 'center' },
      headStyles: { fillColor: [26, 86, 219], textColor: 255 },
      margin: { left: 5, right: 5 }
    });

    window.open(doc.output("bloburl"), "_blank");
  };

  if (!initialized) return null;

  const formattedSyncTime = (() => {
    const hours = lastSyncTime.getHours();
    const minutes = lastSyncTime.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const displayHours = hours % 12 || 12;
    const displayMinutes = minutes < 10 ? '0' + minutes : minutes;
    return `${displayHours}:${displayMinutes} ${ampm}`;
  })();

  // Bypass login requirement since data comes from Google Sheet proxy
  // if (!user) {
  //   ...
  // }

  return (
    <div className="min-h-screen bg-[#F0F4F9] text-gray-900 font-sans pb-12">
      {/* Top Navigation */}
      <header className="sticky top-0 z-[100] bg-blue-600 shadow-md">
        <div className="w-full px-4 sm:px-6 h-16 flex flex-row items-center justify-between gap-3 md:gap-6">
          <div className="flex items-center gap-2.5 sm:gap-3 shrink-0">
            <div className="bg-white/10 p-1.5 sm:p-2 rounded-xl border border-white/20">
              <FileSpreadsheet className="w-4 h-4 sm:w-5 sm:h-5 text-white" />
            </div>
            <h1 className="text-sm sm:text-lg font-bold text-white tracking-tight mr-1 sm:mr-2">Examiner Pro</h1>
            
            {/* Pill: Auto Live Sync */}
            <div className="flex items-center gap-1 px-2 py-0.5 sm:px-3 sm:py-1 bg-white/10 rounded-full border border-white/10 whitespace-nowrap">
              <div className={cn(
                "w-1 sm:w-1.5 h-1 sm:h-1.5 rounded-full animate-pulse",
                !isOnline ? "bg-amber-400" : isSyncing ? "bg-blue-400 animate-bounce" : "bg-green-400"
              )} />
              <span className="text-[8px] sm:text-[10px] font-bold text-white uppercase tracking-wider">
                {!isOnline ? (
                  <>
                    <span className="hidden sm:inline">Auto Sync: Offline (Cached)</span>
                    <span className="sm:hidden">Offline (Cached)</span>
                  </>
                ) : isSyncing ? (
                  <>
                    <span className="hidden sm:inline">Auto Sync: Syncing...</span>
                    <span className="sm:hidden">Syncing...</span>
                  </>
                ) : (
                  <>
                    <span className="hidden sm:inline">Auto Sync: {formattedSyncTime}</span>
                    <span className="sm:hidden">Sync: {formattedSyncTime}</span>
                  </>
                )}
              </span>
            </div>

          </div>

          <div className="bg-white p-0.5 sm:p-1 pr-1.5 sm:pr-2 pl-2 sm:pl-3 rounded-full shadow-sm flex items-center gap-1 sm:gap-1.5 ml-auto flex-nowrap max-w-full">
            {/* Stats & Status Area */}
            <div className="flex items-center gap-1.5 sm:gap-2 shrink-0">
              {/* Total Records */}
              <div className="flex items-center gap-1 sm:gap-1.5 shrink-0">
                <div className="p-1 sm:p-1.5 bg-blue-50 rounded-lg">
                  <Users className="w-3 sm:w-3.5 h-3 sm:h-3.5 text-blue-600" />
                </div>
                <div className="flex flex-col">
                  <span className="text-[7px] sm:text-[9px] font-bold text-gray-400 capitalize whitespace-nowrap hidden lg:inline">Total Records</span>
                  <span className="text-[7px] sm:text-[9px] font-bold text-gray-400 capitalize whitespace-nowrap lg:hidden">Total</span>
                  <div className="flex items-center gap-0.5 sm:gap-1">
                    <span className="text-[10px] sm:text-xs font-black text-gray-800 leading-none">
                      {options?.rowCount !== undefined ? options.rowCount : "-"}
                    </span>
                    <button 
                      onClick={() => loadOptions(true, false, true)}
                      className={cn("p-0.5 hover:bg-gray-100 rounded transition-colors", isSyncing && "animate-spin")}
                      title="Force Refresh Data"
                    >
                      <RefreshCw className="w-2 sm:w-2.5 h-2 sm:h-2.5 text-gray-300" />
                    </button>
                  </div>
                </div>
              </div>
              
              <div className="w-px h-5 sm:h-6 bg-gray-100" />

              {/* Backend Status */}
              <div className="flex items-center gap-1 sm:gap-1.5 shrink-0">
                <div className="p-1 sm:p-1.5 bg-indigo-50 rounded-lg">
                  <div className={cn(
                    "w-3 sm:w-3.5 h-3 sm:h-3.5 rounded-full flex items-center justify-center transition-all",
                    isSyncing ? "bg-blue-500 animate-pulse" : (options ? "bg-emerald-500" : (isPinged ? "bg-amber-400" : "bg-red-500"))
                  )}>
                    <div className="w-1 sm:w-1.5 h-1 sm:h-1.5 bg-white rounded-full opacity-60" />
                  </div>
                </div>
                <div className="flex flex-col">
                  <span className="text-[7px] sm:text-[9px] font-bold text-gray-400 capitalize whitespace-nowrap hidden lg:inline">Backend Status</span>
                  <span className="text-[7px] sm:text-[9px] font-bold text-gray-400 capitalize whitespace-nowrap lg:hidden">Backend</span>
                  <div className="flex items-center gap-1 mt-0.5">
                    <span className={cn(
                      "text-[8px] sm:text-[10px] font-black leading-none uppercase tracking-wider",
                      isInstantLoaded ? "text-emerald-600" : (options ? "text-emerald-600" : (isPinged ? "text-amber-600" : "text-red-600"))
                    )}>
                      {isInstantLoaded ? (
                        <>
                          <span className="hidden lg:inline">Instant Mode Active</span>
                          <span className="lg:hidden">Instant</span>
                        </>
                      ) : (isSyncing ? "Syncing..." : (options ? "Connected" : (isPinged === false ? "Offline" : "Checking...")))}
                    </span>
                    {!options && !isSyncing && (
                      <button 
                        onClick={() => loadOptions()}
                        className="p-0.5 hover:bg-gray-100 rounded-md transition-colors"
                        title="Retry Connection"
                      >
                        <RefreshCw className="w-2 sm:w-2.5 h-2 sm:h-2.5 text-gray-400" />
                      </button>
                    )}
                  </div>
                </div>
              </div>

              {/* Dynamic Results Section */}
              {result && (
                <>
                  <div className="w-px h-5 sm:h-6 bg-gray-100" />
                  <div className="flex items-center gap-1 sm:gap-2 text-blue-600 shrink-0">
                    <div className="p-1 sm:p-1.5 bg-blue-50 rounded-lg">
                      <CheckCircle2 className="w-3 sm:w-3.5 h-3 sm:h-3.5" />
                    </div>
                    <div className="flex flex-col">
                      <span className="text-[7px] sm:text-[9px] font-bold opacity-60 capitalize whitespace-nowrap hidden lg:inline">Results</span>
                      <span className="text-[10px] sm:text-xs font-black leading-none">
                        {result.total} <span className="hidden lg:inline">matching</span><span className="lg:hidden">match</span>
                      </span>
                    </div>
                  </div>
                </>
              )}
            </div>

            {/* Divider between Status and Actions */}
            <div className="w-px h-5 sm:h-6 bg-gray-100" />

            {/* Actions Block */}
            <div className="flex items-center gap-1.5 sm:gap-3 shrink-0">
              <div className="flex items-center gap-1.5 sm:gap-3 pr-1.5 sm:pr-3 border-r border-gray-100 shrink-0">
                <button
                  onClick={() => handleSearch(1)}
                  disabled={isLoading}
                  className="flex items-center gap-1 sm:gap-2 bg-blue-600 hover:bg-blue-700 text-white px-2 sm:px-4 py-1 sm:py-1.5 rounded-lg sm:rounded-xl font-black text-[10px] sm:text-xs transition-all shadow-md shadow-blue-100 active:scale-95 disabled:opacity-50 whitespace-nowrap"
                >
                  {isLoading ? <Loader2 className="w-3 h-3 sm:w-4 sm:h-4 animate-spin" /> : <Search className="w-3 h-3 sm:w-4 sm:h-4 shrink-0" />}
                  <span className="hidden md:inline">Search</span>
                </button>
                <button
                  onClick={handleClear}
                  className="flex items-center gap-1 sm:gap-1.5 group text-gray-400 hover:text-gray-600 transition-colors shrink-0"
                >
                  <X className="w-3 h-3 sm:w-3.5 sm:h-3.5 group-hover:rotate-90 transition-transform duration-300 shrink-0" />
                  <span className="text-[10px] sm:text-xs font-bold hidden md:inline">Clear</span>
                </button>
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                <button
                  onClick={handleExportPDF}
                  disabled={!result}
                  className="p-2 bg-purple-50 text-purple-700 hover:bg-purple-100 rounded-lg font-bold transition-all disabled:opacity-40 shrink-0"
                  title="Print PDF Report"
                >
                  <Printer className="w-4 h-4 shrink-0" />
                </button>
                <button
                  onClick={handleExportExcel}
                  disabled={!result}
                  className="p-2 bg-green-50 text-green-700 hover:bg-green-100 rounded-lg font-bold transition-all disabled:opacity-40 shrink-0"
                  title="Export Excel (XLSX)"
                >
                  <FileDown className="w-4 h-4 shrink-0" />
                </button>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="w-full px-4 sm:px-6 py-3 flex flex-col gap-3">

        {/* Filter Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 xl:grid-cols-4 gap-x-4 gap-y-2 bg-white p-4 rounded-2xl shadow-sm border border-gray-100">
          <FilterDropdown 
            label="Institute Name" 
            placeholder="Select Institute" 
            options={options?.institutes || []} 
            selected={filters.institute}
            onChange={(v) => setFilters(prev => ({ ...prev, institute: v }))}
          />
          <FilterDropdown 
            label="Department Name" 
            placeholder="Select Department" 
            options={options?.departments || []} 
            selected={filters.department}
            onChange={(v) => setFilters(prev => ({ ...prev, department: v }))}
          />
          <FilterDropdown 
            label="HSC Batch" 
            placeholder="Select Batch" 
            options={filteredBatches} 
            selected={filters.batch}
            onChange={(v) => setFilters(prev => ({ ...prev, batch: v }))}
          />
          <FilterDropdown 
            label="Subject" 
            placeholder="Select Subjects" 
            options={Array.isArray(options?.subjects) ? options!.subjects.map(s => ({ label: s?.label || "", value: s?.key || "" })) : []} 
            selected={filters.subjectsSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, subjectsSelected: v }))}
            emptyMeansAll
          />
          <FilterDropdown 
            label="Training Report" 
            placeholder="Select Training" 
            options={options?.trainings || []} 
            selected={filters.trainingsSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, trainingsSelected: v }))}
          />
          <FilterDropdown 
            label="Training Date" 
            placeholder="Select Training Date" 
            options={options?.trainingDates || []} 
            selected={filters.trainingDatesSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, trainingDatesSelected: v }))}
          />
          <FilterDropdown 
            label="Physical Campus" 
            placeholder="Select Campus" 
            options={options?.campuses || []} 
            selected={filters.campusesSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, campusesSelected: v }))}
          />
          <FilterDropdown 
            label="T-PIN" 
            placeholder="Select T-PIN" 
            options={options?.tpins || []} 
            selected={filters.tpinsSelected}
            onChange={(v) => setFilters(prev => ({ ...prev, tpinsSelected: v }))}
          />

          <div className="flex flex-col gap-3 p-4 bg-gray-50/50 rounded-[24px] border border-gray-100">
            <div className="flex flex-col gap-4">
              {/* Settings Row */}
              <div className="flex items-center gap-4">
                <label className="flex items-center gap-2 cursor-pointer group flex-1">
                  <div className={cn(
                    "w-4 h-4 rounded border flex items-center justify-center transition-all",
                    filters.onlyAllowed ? "bg-blue-600 border-blue-600 text-white shadow-sm" : "border-gray-200 bg-white group-hover:border-blue-300"
                  )}>
                    <input 
                      type="checkbox" 
                      className="hidden" 
                      checked={filters.onlyAllowed} 
                      onChange={(e) => setFilters(prev => ({ ...prev, onlyAllowed: e.target.checked }))} 
                    />
                    {filters.onlyAllowed && <CheckCircle2 className="w-3 h-3" />}
                  </div>
                  <span className="text-[11px] font-bold text-gray-600 truncate">Only ALLOWED</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer group flex-1">
                  <div className={cn(
                    "w-4 h-4 rounded border flex items-center justify-center transition-all",
                    filters.subjectLogic === "all" ? "bg-blue-600 border-blue-600 text-white shadow-sm" : "border-gray-200 bg-white group-hover:border-blue-300"
                  )}>
                    <input 
                      type="checkbox" 
                      className="hidden" 
                      checked={filters.subjectLogic === "all"} 
                      onChange={(e) => setFilters(prev => ({ ...prev, subjectLogic: e.target.checked ? "all" : "any" }))} 
                    />
                    {filters.subjectLogic === "all" && <CheckCircle2 className="w-3 h-3" />}
                  </div>
                  <span className="text-[11px] font-bold text-gray-600 truncate">Require ALL Sub</span>
                </label>
              </div>

              {/* Threshold Row */}
              <div className="flex items-center gap-3">
                <div className="flex flex-col gap-1 flex-1">
                  <input 
                    type="number" 
                    value={filters.allowEnglish || ""} 
                    onChange={(e) => setFilters(prev => ({ ...prev, allowEnglish: e.target.value === "" ? null : Number(e.target.value) }))}
                    className="w-full bg-white border border-gray-100 rounded-xl px-2 py-1.5 text-center font-bold text-blue-600 text-xs focus:outline-none focus:ring-2 focus:ring-blue-50 focus:border-blue-400 transition-all shadow-sm"
                    placeholder="55"
                  />
                  <div className="text-[10px] font-bold text-gray-400 text-center tracking-wider mt-0.5">English</div>
                </div>
                <div className="flex flex-col gap-1 flex-1">
                  <input 
                    type="number" 
                    value={filters.allowOthers || ""} 
                    onChange={(e) => setFilters(prev => ({ ...prev, allowOthers: e.target.value === "" ? null : Number(e.target.value) }))}
                    className="w-full bg-white border border-gray-100 rounded-xl px-2 py-1.5 text-center font-bold text-blue-600 text-xs focus:outline-none focus:ring-2 focus:ring-blue-50 focus:border-blue-400 transition-all shadow-sm"
                    placeholder="48"
                  />
                  <div className="text-[10px] font-bold text-gray-400 text-center tracking-wider mt-0.5">Ban, P, C, M, B, I</div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Results Section */}
        <div 
          id="results-section"
          className={cn(
            "flex flex-col gap-0 bg-white rounded-[32px] overflow-hidden shadow-sm border border-gray-100 transition-all duration-500",
            result ? "min-h-[400px]" : "min-h-[100px]"
          )}
        >
          <AnimatePresence>
            {hasSheetUpdates && !autoRefresh && (
              <motion.button
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: "auto", opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                onClick={() => handleSearch(1)}
                className="w-full bg-blue-600 text-white font-bold py-3 text-sm flex items-center justify-center gap-2 hover:bg-blue-700 transition-colors"
              >
                <RefreshCcw className="w-4 h-4 animate-spin-slow" />
                Sheet updated! Click to reload with new data
              </motion.button>
            )}
          </AnimatePresence>


          {errorDetails && (
            <div className="p-12 flex flex-col items-center justify-center text-center">
              <div className="w-16 h-16 bg-red-50 text-red-500 rounded-2xl flex items-center justify-center mb-6 shadow-sm">
                <ShieldAlert className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-black text-gray-900 mb-3 tracking-tight">Connectivity Issue</h3>
              <div className="bg-red-50/50 border border-red-100 rounded-2xl p-6 max-w-lg mb-8 text-center text-balance">
                <p className="text-red-800 text-sm font-medium leading-relaxed mb-1">{errorDetails.message}</p>
                {errorDetails.advice && (
                  <p className="text-red-600 text-xs font-bold mt-2 uppercase tracking-wide">💡 {errorDetails.advice}</p>
                )}
                
                <div className="text-left space-y-2 mt-4 pt-4 border-t border-red-100">
                  <p className="text-xs text-red-600 font-bold uppercase tracking-wider">Troubleshooting:</p>
                  <ul className="text-xs text-red-700/70 list-disc list-inside space-y-1">
                    <li>Is GAS deployed as <b>"Anyone"</b> under "Who has access"?</li>
                    <li>Did you copy the <b>NEW Exec URL</b> from the deployment screen?</li>
                    <li>Ensure <b>Execute as: Me</b> is selected in the deployment settings.</li>
                    <li>Verify the <b>GAS_DEPLOYMENT_URL</b> in AI Studio Secrets matches exactly.</li>
                  </ul>
                </div>
              </div>
              <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                <button 
                  onClick={() => { setErrorDetails(null); setDiagnostics(null); loadOptions(false, false, true); }}
                  className="flex items-center gap-2 bg-gray-900 hover:bg-black text-white px-8 py-3 rounded-2xl font-bold text-sm transition-all active:scale-95 shadow-xl shadow-gray-200"
                >
                  <RefreshCw className={cn("w-4 h-4", isSyncing && "animate-spin")} />
                  Retry Connection
                </button>
                <button 
                  onClick={runDiagnostics}
                  disabled={isDiagnosing}
                  className="flex items-center gap-2 bg-blue-50 hover:bg-blue-100 text-blue-700 px-8 py-3 rounded-2xl font-bold text-sm transition-all active:scale-95 disabled:opacity-50"
                >
                  {isDiagnosing ? <Loader2 className="w-4 h-4 animate-spin" /> : <ShieldAlert className="w-4 h-4" />}
                  {isDiagnosing ? "Running Diagnostics..." : "Run Connection Diagnostics"}
                </button>
              </div>

              {diagnostics && diagnostics.length > 0 && (
                <div className="mt-8 bg-gray-50/50 border border-gray-100 rounded-[24px] p-6 max-w-xl w-full text-left">
                  <h4 className="text-sm font-bold text-gray-800 mb-4 flex items-center gap-2">
                    <CheckCircle2 className="w-4 h-4 text-blue-500" />
                    Step-by-Step Diagnostic Report
                  </h4>
                  <div className="space-y-4">
                    {diagnostics.map((step, idx) => (
                      <div key={idx} className="flex gap-3 items-start">
                        <div className={cn(
                          "w-5 h-5 rounded-full flex items-center justify-center shrink-0 mt-0.5",
                          step.status === "success" ? "bg-emerald-100 text-emerald-600" : "bg-red-100 text-red-600"
                        )}>
                          {step.status === "success" ? (
                            <CheckCircle2 className="w-3.5 h-3.5" />
                          ) : (
                            <X className="w-3.5 h-3.5" />
                          )}
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className="text-xs font-bold text-gray-800">{step.name}</p>
                          <p className="text-xs text-gray-500 mt-0.5">{step.details}</p>
                          {step.advice && (
                            <p className="text-xs text-amber-600 font-bold mt-1 bg-amber-50/50 border border-amber-100/30 px-2 py-1 rounded-lg">
                              💡 Advice: {step.advice}
                            </p>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}

          {!result && !errorDetails && (
            <div className="flex flex-col items-center justify-center py-16 px-4 text-center">
              {!options && !isLoading ? (
                <>
                  <div className="w-20 h-20 bg-amber-50 text-amber-500 rounded-3xl flex items-center justify-center mb-6 animate-pulse shadow-xl shadow-amber-100/50">
                    <Database className="w-10 h-10" />
                  </div>
                  <h3 className="text-2xl font-black text-gray-900 mb-3 tracking-tight">Database Connectivity</h3>
                  <p className="text-gray-500 text-sm max-w-sm leading-relaxed mb-8 font-medium">
                    The application hasn't connected to your Google Sheet yet. Please ensure the backend is deployed correctly.
                  </p>
                  <button 
                    onClick={() => { loadOptions(false, false, true); }}
                    className="flex items-center gap-2 bg-gray-900 hover:bg-black text-white px-8 py-3 rounded-2xl font-bold text-sm transition-all active:scale-95 shadow-xl shadow-gray-200"
                  >
                    <RefreshCcw className="w-4 h-4" />
                    Connect to Sheet
                  </button>
                </>
              ) : (
                <>
                  <motion.div 
                    className="w-20 h-20 bg-blue-50 text-blue-500 rounded-[24px] flex items-center justify-center mb-6 shadow-xl shadow-blue-100/50"
                    animate={isLoading ? {
                      scale: [1, 1.25, 1],
                    } : {
                      scale: 1
                    }}
                    transition={isLoading ? {
                      duration: 1.5,
                      repeat: Infinity,
                      ease: "easeInOut"
                    } : {}}
                  >
                    <motion.div
                      animate={isLoading ? {
                        rotate: 360
                      } : {
                        rotate: 0
                      }}
                      transition={isLoading ? {
                        duration: 1.8,
                        repeat: Infinity,
                        ease: "linear"
                      } : {}}
                    >
                      <Circle className="w-10 h-10" />
                    </motion.div>
                  </motion.div>
                  <h3 className="text-xl font-black text-gray-900 tracking-tight transition-all duration-300">
                    {isLoading ? "Searching......" : "Ready to Search"}
                  </h3>
                </>
              )}
            </div>
          )}

          {result && (
            <>
              <div className="px-6 py-4 bg-gray-50 border-b border-gray-100 flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="flex items-center gap-2">
                    <span className="px-3 py-1 bg-white border border-gray-200 rounded-lg text-[13px] font-bold text-gray-800">
                      {result.total} Records Found
                    </span>
                    {searchLatency && (
                      <span className="flex items-center gap-1 px-2.5 py-0.5 bg-green-50 text-green-700 text-[10px] font-bold rounded-full border border-green-100/50">
                        ⚡ Speed: {searchLatency}
                      </span>
                    )}
                    <span className="text-xs text-gray-400 font-medium italic">
                      Showing Page {page} of {result.totalPages}
                    </span>
                  </div>
                </div>

                <div className="flex items-center gap-4">
                  <div className="flex items-center gap-2 bg-white p-1 rounded-xl border border-gray-100 shadow-sm">
                    <button
                      onClick={() => handleSearch(page - 1)}
                      disabled={page <= 1 || isLoading}
                      className="p-1.5 hover:bg-gray-100 rounded-lg transition-colors disabled:opacity-30 disabled:cursor-not-allowed"
                    >
                      <ChevronLeft className="w-5 h-5" />
                    </button>
                    <span className="px-3 text-sm font-bold text-gray-700">{page}</span>
                    <button
                      onClick={() => handleSearch(page + 1)}
                      disabled={page >= result.totalPages || isLoading}
                      className="p-1.5 hover:bg-gray-100 rounded-lg transition-colors disabled:opacity-30 disabled:cursor-not-allowed"
                    >
                      <ChevronRight className="w-5 h-5" />
                    </button>
                  </div>
                  
                  <select 
                    value={pageSize}
                    onChange={(e) => {
                      setPageSize(Number(e.target.value));
                      // Search reset happens after state updates in effect or manual trigger
                    }}
                    className="bg-white border border-gray-100 rounded-xl px-3 py-2 text-xs font-bold shadow-sm focus:outline-none"
                  >
                    {[100, 200, 300, 500].map(size => (
                      <option key={size} value={size}>{size} per page</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="flex-1 overflow-auto relative max-h-[650px] custom-scrollbar">
                <table className="w-full border-collapse text-left text-sm table-auto min-w-[1200px]">
                  <thead className="sticky top-0 z-10">
                    <tr className="bg-blue-600 text-white font-bold capitalize text-[11px]">
                      {Array.isArray(result.header) && result.header.map((h, i) => (
                        <th key={i} className="px-4 py-4 whitespace-nowrap border-r border-blue-500/30 last:border-0 text-center">
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {Array.isArray(result.rows) && result.rows.map((row, i) => (
                      <tr key={i} className="hover:bg-blue-50/50 transition-colors group">
                        {Array.isArray(row) && row.map((cell, j) => {
                          const isStatus = result.header && result.header[j] && result.header[j].toLowerCase().includes("allow status");
                          return (
                            <td 
                              key={j} 
                              className={cn(
                                "px-4 py-3 whitespace-nowrap text-center text-gray-600 border-r border-gray-50 last:border-0",
                                isStatus && cell === "ALLOWED" && "text-green-600 font-extrabold",
                                isStatus && cell === "NOT ALLOWED" && "text-red-500 font-extrabold"
                              )}
                            >
                              {cell}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              
              <div className="p-6 bg-gray-50/60 border-t border-gray-100 flex flex-col gap-2">
                 <p className="text-[11px] text-gray-400 font-bold capitalize text-center"> End of Report </p>
                 <div className="flex items-center justify-between mt-2">
                    <span className="text-[10px] text-gray-400">Page sizing set to {pageSize} rows</span>
                    <span className="text-[10px] text-gray-400 italic">Fast rendering mode active</span>
                 </div>
              </div>
            </>
          )}
        </div>
      </main>
      
      {/* Footer Branding */}
      <footer className="mt-4 mb-12 text-center">
        <p className="text-[11px] text-gray-400 font-bold capitalize mb-1">Examiner Information Management</p>
        <p className="text-[9px] text-gray-300">Secure Environment / Data Protected by Google Workspace Policy</p>
      </footer>

      {/* Floating Success Toast / Background Sync Notification */}
      {showSyncToast && (
        <div className="fixed bottom-6 right-6 z-[999] max-w-sm bg-gray-900 text-white px-5 py-4 rounded-[20px] shadow-2xl flex items-center gap-4 border border-white/10 animate-in fade-in slide-in-from-bottom duration-300">
          <div className="bg-emerald-500/10 p-2 rounded-xl border border-emerald-500/30">
            <CheckCircle2 className="w-5 h-5 text-emerald-400 animate-bounce" />
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-xs font-black text-gray-100 uppercase tracking-wider">Sync Status</p>
            <p className="text-xs text-gray-300 font-medium mt-0.5 leading-relaxed">{syncToastMessage}</p>
          </div>
          <button 
            onClick={() => setShowSyncToast(false)}
            className="p-1 hover:bg-white/10 rounded-full transition-colors text-gray-400 hover:text-white"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}
    </div>
  );
};

export default function App() {
  return (
    <AuthProvider>
      <Dashboard />
    </AuthProvider>
  );
}
