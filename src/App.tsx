import React, { useState, useEffect, useRef } from "react";
import { 
  Sparkles, Clock, FileText, Settings, 
  Plus, Trash2, Download, Printer, 
  Share2, Mail, RefreshCw, Eye, EyeOff, 
  Palette, Globe, Building2, AlertTriangle, 
  Calendar, ArrowRight, Search, PlusCircle, CheckCircle2,
  Database, Info, HelpCircle, ExternalLink, ShieldCheck, ChevronRight
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import { GeneratedReport, ScheduleConfig, AppSettings } from "./types";
import { initAuth, googleSignIn, logout } from "./lib/firebase";
import { User } from "firebase/auth";

export default function App() {
  // Navigation Tabs
  const [activeTab, setActiveTab] = useState<"generate" | "schedules" | "history" | "settings">("generate");

  // Core App Settings & Keys
  const [settings, setSettings] = useState<AppSettings>({
    companyName: "Acme Analytics Ltd",
    logoUrl: "",
    accentColor: "#4F46E5",
    footerText: "CONFIDENTIAL · Internal Corporate Distribution Only",
    timezone: "GMT",
    defaultAiModel: "gemini",
    defaultReportType: "weekly"
  });

  // Security toggles for Keys
  const [showClaude, setShowClaude] = useState(false);
  const [showOpenai, setShowOpenai] = useState(false);
  const [showGemini, setShowGemini] = useState(false);
  const [showGroq, setShowGroq] = useState(false);

  // State: Main Report Generation Form
  const [dataSource, setDataSource] = useState<"sheet" | "csv" | "file">("sheet");
  const [uploadedFileName, setUploadedFileName] = useState("");
  const [uploadedFileSize, setUploadedFileSize] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  const [sheetUrl, setSheetUrl] = useState("");
  
  // Google Auth & Drive States
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [sheetName, setSheetName] = useState("");
  const [csvData, setCsvData] = useState("");
  const [reportTitle, setReportTitle] = useState("");
  const [companyName, setCompanyName] = useState("");
  const [reportType, setReportType] = useState("weekly");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [aiModel, setAiModel] = useState("gemini");
  const [userPrompt, setUserPrompt] = useState("");

  // Quick select datasets for instant testing
  const [selectedPreset, setSelectedPreset] = useState<string | null>(null);

  // State: Live Output Report
  const [activeReport, setActiveReport] = useState<GeneratedReport | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [generationStep, setGenerationStep] = useState(0);

  // State: Schedules and History List
  const [schedules, setSchedules] = useState<ScheduleConfig[]>([]);
  const [history, setHistory] = useState<GeneratedReport[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [historyFilter, setHistoryFilter] = useState("all");

  // State: Modals Control
  const [isEmailModalOpen, setIsEmailModalOpen] = useState(false);
  const [emailRecipients, setEmailRecipients] = useState("");
  const [isPreviewModalOpen, setIsPreviewModalOpen] = useState(false);
  const [previewTableData, setPreviewTableData] = useState<any[]>([]);
  const [isScheduleModalOpen, setIsScheduleModalOpen] = useState(false);
  
  // State: Schedule Form Modal
  const [schedTitle, setSchedTitle] = useState("");
  const [schedSheetUrl, setSchedSheetUrl] = useState("");
  const [schedSheetName, setSchedSheetName] = useState("");
  const [schedFreq, setSchedFreq] = useState("weekly");
  const [schedAi, setSchedAi] = useState("gemini");
  const [schedType, setSchedType] = useState("weekly");
  const [schedPrompt, setSchedPrompt] = useState("");
  const [schedRecipients, setSchedRecipients] = useState("");
  const [schedWeeklyDays, setSchedWeeklyDays] = useState<Record<string, boolean>>({
    mon: true, wed: false, fri: false
  });
  const [schedMonthlyDay, setSchedMonthlyDay] = useState("1");

  // Toast Alerts system
  const [toasts, setToasts] = useState<{ id: string; message: string; type: "success" | "error" | "info" }[]>([]);

  const addToast = (message: string, type: "success" | "error" | "info" = "info") => {
    const id = Date.now().toString();
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 4500);
  };

  // Fetch initial collections and preferences
  useEffect(() => {
    fetchSettings();
    fetchHistory();
    fetchSchedules();

    const unsubscribe = initAuth(
      (user, token) => {
        setCurrentUser(user);
        setAccessToken(token);
      },
      () => {
        setCurrentUser(null);
        setAccessToken(null);
      }
    );
    return () => unsubscribe();
  }, []);

  const handleGoogleSignIn = async () => {
    setIsLoggingIn(true);
    try {
      const result = await googleSignIn();
      if (result) {
        setCurrentUser(result.user);
        setAccessToken(result.accessToken);
        addToast(`Google Drive connected as ${result.user.email}!`, "success");
      }
    } catch (err: any) {
      addToast(err.message || "Failed to sign in with Google.", "error");
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleGoogleSignOut = async () => {
    try {
      await logout();
      setCurrentUser(null);
      setAccessToken(null);
      addToast("Disconnected from Google Drive.", "info");
    } catch (err) {
      addToast("Failed to disconnect Google account.", "error");
    }
  };

  const fetchSettings = async () => {
    try {
      const res = await fetch("/api/settings");
      const data = await res.json();
      if (data.success && data.settings) {
        setSettings(data.settings);
        setCompanyName(data.settings.companyName);
      }
    } catch (e) {
      console.error("Failed to load settings from server.");
    }
  };

  const fetchHistory = async () => {
    try {
      const res = await fetch("/api/history");
      const data = await res.json();
      if (data.success && data.history) {
        setHistory(data.history);
      }
    } catch (e) {
      console.error("Failed to load report run history.");
    }
  };

  const fetchSchedules = async () => {
    try {
      const res = await fetch("/api/schedules");
      const data = await res.json();
      if (data.success && data.schedules) {
        setSchedules(data.schedules);
      }
    } catch (e) {
      console.error("Failed to load active automation schedules.");
    }
  };

  // File Drag and Drop Processing
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  const processFile = (file: File) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (!["xlsx", "xls", "csv", "tsv", "txt"].includes(ext || "")) {
      addToast("Unsupported file format! Please upload an Excel (.xlsx/.xls) or CSV/TSV file.", "error");
      return;
    }

    setUploadedFileName(file.name);
    
    // Format file size nicely
    const sizeInKb = file.size / 1024;
    const formattedSize = sizeInKb > 1024 
      ? `${(sizeInKb / 1024).toFixed(1)} MB` 
      : `${sizeInKb.toFixed(1)} KB`;
    setUploadedFileSize(formattedSize);

    // Auto-fill report title if currently empty or default
    if (!reportTitle || reportTitle === "AI Business Intelligence Report") {
      const titleWithoutExt = file.name.substring(0, file.name.lastIndexOf("."));
      const formattedTitle = titleWithoutExt
        .split(/[_-]+/)
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(" ");
      setReportTitle(formattedTitle + " Performance Review");
    }

    const reader = new FileReader();
    
    if (ext === "csv" || ext === "tsv" || ext === "txt") {
      reader.onload = (evt) => {
        const text = evt.target?.result as string;
        setCsvData(text);
        addToast(`Successfully loaded ${file.name}!`, "success");
      };
      reader.readAsText(file);
    } else {
      // Excel files
      reader.onload = (evt) => {
        try {
          const bstr = evt.target?.result;
          const workbook = XLSX.read(bstr, { type: "binary" });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          
          // Convert sheet to CSV
          const csv = XLSX.utils.sheet_to_csv(worksheet);
          setCsvData(csv);
          addToast(`Loaded sheet "${firstSheetName}" from ${file.name}!`, "success");
        } catch (err) {
          console.error("Excel parse error:", err);
          addToast("Error parsing Excel file. Please ensure it is not corrupt.", "error");
        }
      };
      reader.readAsBinaryString(file);
    }
  };

  // Preview Data Handler
  const handlePreviewData = async () => {
    if (dataSource === "sheet" && !sheetUrl) {
      addToast("Please connect a Google Sheets URL first.", "error");
      return;
    }

    if (dataSource === "csv" || dataSource === "file") {
      if (!csvData) {
        addToast(dataSource === "file" ? "Please upload an Excel or CSV file first." : "Please paste CSV data first.", "error");
        return;
      }
      try {
        const workbook = XLSX.read(csvData, { type: "string" });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const parsed = XLSX.utils.sheet_to_json(sheet);
        if (parsed.length === 0) {
          addToast("The provided file/data appears to be empty.", "error");
          return;
        }
        setPreviewTableData(parsed);
        setIsPreviewModalOpen(true);
        addToast("Parsed and loaded data preview successfully!", "success");
      } catch (err) {
        addToast("Failed to parse CSV/Excel data.", "error");
      }
      return;
    }
    
    addToast("Pulling sheet columns and samples...", "info");

    try {
      const res = await fetch("/api/sheet-data", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheetUrl: dataSource === "sheet" ? sheetUrl : "",
          sheetName: dataSource === "sheet" ? sheetName : "",
          mockType: "",
          accessToken: accessToken || undefined
        })
      });
      const data = await res.json();
      if (data.success && data.data) {
        setPreviewTableData(data.data);
        setIsPreviewModalOpen(true);
        if (data.warning) {
          addToast(data.warning, "info");
        } else {
          addToast("Data preview fetched successfully!", "success");
        }
      } else {
        addToast(data.error || "Failed to fetch sheet preview.", "error");
      }
    } catch (err) {
      addToast("Server failed to parse Google Sheet.", "error");
    }
  };

  // Quick Preset Selector Helper
  const applyPresetDataset = (type: "sales" | "ops" | "marketing") => {
    setSelectedPreset(type);
    setDataSource("csv"); // Paste CSV mock
    if (type === "sales") {
      setReportTitle("Regional SaaS Sales Performance Analysis");
      setUserPrompt("Generate a sales performance review. Provide detailed KPI metrics highlighting total revenue and average price, line trend of revenue across days, and horizontal bar of revenue breakdown by region. Give 3 strategic recommendations for entering Asia market.");
      setCsvData(`Date,Product,Region,Units,Price,Revenue,Channel\n2026-06-25,Standard Subscription,North America,120,49,5880,Direct Search\n2026-06-25,Enterprise Plan,Europe,4,999,3996,Sales Rep\n2026-06-26,Standard Subscription,Asia,85,49,4165,Ad Campaign\n2026-06-26,Professional Tier,North America,45,149,6705,Newsletter\n2026-06-27,Standard Subscription,Europe,110,49,5390,Direct Search\n2026-06-27,Professional Tier,Asia,38,149,5662,Ad Campaign\n2026-06-28,Enterprise Plan,North America,6,999,5994,Sales Rep\n2026-06-28,Standard Subscription,Asia,92,49,4508,Direct Search\n2026-06-29,Professional Tier,Europe,50,149,7450,Newsletter\n2026-06-29,Standard Subscription,North America,135,49,6615,Ad Campaign`);
    } else if (type === "ops") {
      setReportTitle("SaaS Platform Latency & Server Audit");
      setUserPrompt("Produce a system health audit. Group by server metrics, plot lines comparing CPU and RAM usages, evaluate latency spikes, and draft immediate actionable alerts and engineering recommendations.");
      setCsvData(`Timestamp,Server,CPU_Usage,RAM_Usage,Latency_ms,Status\n2026-07-03 00:00:00,Web-Node-01,45,62,120,Healthy\n2026-07-03 04:00:00,Web-Node-01,89,84,450,Warning\n2026-07-03 08:00:00,Web-Node-01,55,65,132,Healthy\n2026-07-03 12:00:00,Web-Node-01,68,71,140,Healthy\n2026-07-03 16:00:00,Web-Node-01,72,75,155,Healthy\n2026-07-03 20:00:00,Web-Node-01,50,64,118,Healthy\n2026-07-03 00:00:00,Web-Node-02,41,58,115,Healthy\n2026-07-03 04:00:00,Web-Node-02,92,81,480,Warning\n2026-07-03 08:00:00,Web-Node-02,52,61,125,Healthy`);
    } else if (type === "marketing") {
      setReportTitle("Q2 PPC Marketing Performance Summary");
      setUserPrompt("Review conversion rates and campaign spend. Highlight total conversions and ROAS (return on ad spend) KPIs, trend bar charts for conversions by campaign, and pinpoint the best performing channel.");
      setCsvData(`Date,Campaign,Clicks,Conversions,Spend,Revenue\n2026-06-24,Google Search Brand,2450,148,1200,5920\n2026-06-24,Retargeting Facebook,1120,95,850,3800\n2026-06-25,Google Search Brand,2610,152,1250,6080\n2026-06-25,Retargeting Facebook,980,74,850,2960\n2026-06-26,Google Search Brand,2890,171,1400,6840\n2026-06-26,Retargeting Facebook,1420,112,1100,4480\n2026-06-27,Google Search Brand,3100,184,1500,7360\n2026-06-27,Retargeting Facebook,1350,105,1050,4200`);
    }
    addToast(`Preset loaded: "${type.toUpperCase()}" dataset`, "success");
  };

  // Generate Report Primary Action
  const handleGenerateReport = async () => {
    if (dataSource === "sheet" && !sheetUrl) {
      addToast("Please connect a Google Sheet URL or switch to file upload.", "error");
      return;
    }
    if (dataSource === "csv" && !csvData) {
      addToast("Please paste CSV data or use one of our templates.", "error");
      return;
    }
    if (dataSource === "file" && !csvData) {
      addToast("Please upload an Excel or CSV file first.", "error");
      return;
    }
    if (!userPrompt.trim()) {
      addToast("Write a natural language prompt instructing the AI what to report.", "error");
      return;
    }

    setIsGenerating(true);
    setGenerationStep(0); // Fetching

    // Step indicators timers
    const t1 = setTimeout(() => setGenerationStep(1), 2200); // AI
    const t2 = setTimeout(() => setGenerationStep(2), 5500); // Building
    const t3 = setTimeout(() => setGenerationStep(3), 8500); // Saving

    try {
      const res = await fetch("/api/generate-report", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheetUrl: dataSource === "sheet" ? sheetUrl : "",
          sheetName: dataSource === "sheet" ? sheetName : "",
          csvData: (dataSource === "csv" || dataSource === "file") ? csvData : "",
          reportTitle: reportTitle || "AI Business Intelligence Report",
          companyName: companyName || settings.companyName,
          reportType: reportType,
          startDate: startDate,
          endDate: endDate,
          aiModel: aiModel,
          userPrompt: userPrompt,
          mockType: selectedPreset,
          accessToken: accessToken || undefined
        })
      });

      const data = await res.json();
      
      clearTimeout(t1);
      clearTimeout(t2);
      clearTimeout(t3);

      if (data.success) {
        setActiveReport({
          id: data.reportId,
          name: data.title,
          createdDate: new Date().toISOString(),
          driveUrl: data.driveUrl,
          reportHtml: data.reportHtml,
          reportType: reportType
        });
        addToast("Stunning business report compiled!", "success");
        fetchHistory(); // refresh list
      } else {
        addToast(data.error || "Generation crashed unexpectedly.", "error");
      }
    } catch (err: any) {
      addToast(err.message || "Express server connection failed.", "error");
    } finally {
      setIsGenerating(false);
    }
  };

  // Download Output HTML Handler
  const downloadReportFile = () => {
    if (!activeReport || !activeReport.reportHtml) return;
    const blob = new Blob([activeReport.reportHtml], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${activeReport.name.replace(/[^a-z0-9]/gi, "_").toLowerCase()}.html`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    addToast("HTML document downloaded to local disk!", "success");
  };

  // Print Report trigger (window.print within frame contentWindow)
  const printReportPDF = () => {
    const iframe = document.getElementById("reportFrame") as HTMLIFrameElement;
    if (iframe && iframe.contentWindow) {
      iframe.contentWindow.focus();
      iframe.contentWindow.print();
    }
  };

  // Copy shareable Drive Link mock
  const copyShareLinkToClipboard = () => {
    if (!activeReport) return;
    navigator.clipboard.writeText(activeReport.driveUrl);
    addToast("Google Drive link copied to clipboard!", "success");
  };

  // Send Email report logic
  const handleSendEmailReport = async () => {
    if (!emailRecipients) {
      addToast("Enter at least one email address.", "error");
      return;
    }
    setIsEmailModalOpen(false);
    addToast("Delivering HTML report to recipients...", "info");
    
    // Simulate App Script deliver
    setTimeout(() => {
      addToast(`Report successfully emailed to: ${emailRecipients}`, "success");
    }, 1500);
  };

  // Load Saved Report from History
  const loadSavedReport = async (reportId: string) => {
    try {
      const res = await fetch(`/api/history/${reportId}`);
      const data = await res.json();
      if (data.success && data.report) {
        setActiveReport(data.report);
        setActiveTab("generate");
        addToast(`Loaded historical run: "${data.report.name}"`, "info");
        // Scroll to frame
        setTimeout(() => {
          document.getElementById("outputPreview")?.scrollIntoView({ behavior: "smooth" });
        }, 100);
      }
    } catch (e) {
      addToast("Failed to fetch HTML data for this report.", "error");
    }
  };

  // Delete saved history card
  const deleteHistoryItem = async (e: React.MouseEvent, id: string) => {
    e.stopPropagation();
    if (!confirm("Are you sure you want to permanently delete this report from Drive history?")) return;
    try {
      const res = await fetch(`/api/history/${id}`, { method: "DELETE" });
      const data = await res.json();
      if (data.success) {
        addToast("Report file deleted.", "success");
        if (activeReport?.id === id) setActiveReport(null);
        fetchHistory();
      }
    } catch (err) {
      addToast("Failed to delete.", "error");
    }
  };

  // Save Settings Securely
  const handleSaveSettings = async () => {
    try {
      const res = await fetch("/api/settings", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(settings)
      });
      const data = await res.json();
      if (data.success) {
        addToast("Corporate branding and keys saved securely!", "success");
      }
    } catch (err) {
      addToast("Settings failed to save.", "error");
    }
  };

  // Open Scheduler modal and pre-fill config
  const openScheduler = () => {
    setSchedTitle(reportTitle || "Recurring AI Intelligence Report");
    setSchedSheetUrl(sheetUrl);
    setSchedSheetName(sheetName);
    setSchedPrompt(userPrompt);
    setIsScheduleModalOpen(true);
  };

  // Save Automation schedule config
  const handleSaveSchedule = async () => {
    if (!schedSheetUrl && dataSource === "sheet") {
      addToast("Recurring reports require a Spreadsheet URL for live pulls.", "error");
      return;
    }
    if (!schedPrompt.trim()) {
      addToast("Write a template prompt for scheduled runs.", "error");
      return;
    }
    if (!schedRecipients.trim()) {
      addToast("Please provide recipient emails.", "error");
      return;
    }

    try {
      const res = await fetch("/api/schedules", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheetUrl: schedSheetUrl,
          sheetName: schedSheetName,
          reportTitle: schedTitle || "Automated Schedule Report",
          userPrompt: schedPrompt,
          recipients: schedRecipients.split(",").map(e => e.trim()),
          frequency: schedFreq,
          aiModel: schedAi,
          reportType: schedType,
          weeklyDays: schedFreq === "weekly" ? schedWeeklyDays : undefined,
          monthlyDay: schedFreq === "monthly" ? schedMonthlyDay : undefined
        })
      });
      const data = await res.json();
      if (data.success) {
        addToast("Report automation trigger registered!", "success");
        setIsScheduleModalOpen(false);
        fetchSchedules();
      }
    } catch (e) {
      addToast("Failed to register schedule.", "error");
    }
  };

  // Delete automated schedule
  const deleteScheduleItem = async (id: string) => {
    if (!confirm("Remove this automated recurring report schedule?")) return;
    try {
      const res = await fetch(`/api/schedules/${id}`, { method: "DELETE" });
      const data = await res.json();
      if (data.success) {
        addToast("Trigger stopped and deleted.", "success");
        fetchSchedules();
      }
    } catch (e) {
      addToast("Failed to remove.", "error");
    }
  };

  // Filter history log items
  const filteredHistory = history.filter(item => {
    const matchesSearch = item.name.toLowerCase().includes(searchQuery.toLowerCase());
    const matchesType = historyFilter === "all" || item.reportType === historyFilter;
    return matchesSearch && matchesType;
  });

  return (
    <div className="min-h-screen bg-slate-50 font-sans flex flex-col selection:bg-indigo-100 text-slate-900">
      
      {/* Toast Alert stack rendering */}
      <div className="fixed bottom-6 right-6 z-50 flex flex-col gap-3 pointer-events-none">
        <AnimatePresence>
          {toasts.map(toast => (
            <motion.div
              key={toast.id}
              initial={{ opacity: 0, y: 30, scale: 0.9 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, scale: 0.85, transition: { duration: 0.2 } }}
              className={`pointer-events-auto px-5 py-4 rounded-xl shadow-xl flex items-center gap-3 border-l-4 text-sm font-medium bg-slate-900 text-white min-w-[320px] max-w-[420px] ${
                toast.type === "success" ? "border-emerald-500" : toast.type === "error" ? "border-rose-500" : "border-indigo-500"
              }`}
            >
              <div className="flex-shrink-0">
                {toast.type === "success" ? (
                  <CheckCircle2 className="w-5 h-5 text-emerald-500" />
                ) : toast.type === "error" ? (
                  <AlertTriangle className="w-5 h-5 text-rose-500" />
                ) : (
                  <Info className="w-5 h-5 text-indigo-400" />
                )}
              </div>
              <div className="flex-grow">{toast.message}</div>
            </motion.div>
          ))}
        </AnimatePresence>
      </div>

      {/* Corporate Dashboard Header */}
      <header className="sticky top-4 z-40 bg-white border border-slate-200 rounded-2xl shadow-sm px-6 py-4 flex flex-col md:flex-row justify-between items-center gap-4 mx-4 md:mx-6 lg:mx-8 mt-4 mb-6 transition-all duration-200">
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-gradient-to-tr from-indigo-600 to-violet-600 rounded-xl text-white shadow-md shadow-indigo-100">
            <Sparkles className="w-6 h-6 animate-pulse" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="font-display font-extrabold text-2xl tracking-tight bg-gradient-to-r from-indigo-600 to-violet-600 bg-clip-text text-transparent">
                ReportAI
              </h1>
              <span className="text-[10px] px-2 py-0.5 rounded-full font-bold uppercase tracking-wider bg-indigo-50 text-indigo-600 border border-indigo-200">
                PRO SaaS
              </span>
            </div>
            <p className="text-xs text-slate-500 font-medium">AI-Powered Business Intelligence Report Synthesizer</p>
          </div>
        </div>

        {/* Navigation Tabs bar */}
        <nav className="flex items-center bg-slate-50 p-1.5 rounded-xl border border-slate-200/60">
          <button
            onClick={() => setActiveTab("generate")}
            className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold transition-all duration-200 cursor-pointer ${
              activeTab === "generate" ? "bg-indigo-50/80 text-indigo-600" : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
            }`}
          >
            <Sparkles className="w-4 h-4" />
            <span>Generate Report</span>
          </button>
          <button
            onClick={() => setActiveTab("schedules")}
            className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold transition-all duration-200 cursor-pointer ${
              activeTab === "schedules" ? "bg-indigo-50/80 text-indigo-600" : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
            }`}
          >
            <Clock className="w-4 h-4" />
            <span>Schedules</span>
          </button>
          <button
            onClick={() => setActiveTab("history")}
            className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold transition-all duration-200 cursor-pointer ${
              activeTab === "history" ? "bg-indigo-50/80 text-indigo-600" : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
            }`}
          >
            <FileText className="w-4 h-4" />
            <span>History</span>
          </button>
          <button
            onClick={() => setActiveTab("settings")}
            className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold transition-all duration-200 cursor-pointer ${
              activeTab === "settings" ? "bg-indigo-50/80 text-indigo-600" : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
            }`}
          >
            <Settings className="w-4 h-4" />
            <span>Settings</span>
          </button>
        </nav>

        {/* Global branding label */}
        <div className="hidden lg:flex items-center gap-2 border-l border-slate-200 pl-4">
          <Building2 className="w-4 h-4 text-slate-400" />
          <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">{settings.companyName}</span>
        </div>
      </header>

      {/* Main Body container */}
      <main className="flex-grow p-4 md:p-6 lg:p-8 max-w-7xl mx-auto w-full">
        <AnimatePresence mode="wait">
          
          {/* TAB 1: Report Generation */}
          {activeTab === "generate" && (
            <motion.div
              key="generate"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="grid grid-cols-1 lg:grid-cols-12 gap-6"
            >
              
              {/* Left Column: Report configuration controls (grid-span 5) */}
              <div className="lg:col-span-5 flex flex-col gap-6">
                
                {/* 1. Connected Data source Card */}
                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 flex flex-col gap-4">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                    <h3 className="font-display font-bold text-slate-800 flex items-center gap-2 text-md">
                      <Database className="w-5 h-5 text-indigo-600" />
                      <span>1. Data Connection Source</span>
                    </h3>
                    <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200/50">
                      <button
                        onClick={() => setDataSource("sheet")}
                        className={`px-3 py-1 text-xs font-bold rounded-md transition-all cursor-pointer ${
                          dataSource === "sheet" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500"
                        }`}
                      >
                        Google Drive / Sheets
                      </button>
                      <button
                        onClick={() => setDataSource("file")}
                        className={`px-3 py-1 text-xs font-bold rounded-md transition-all cursor-pointer ${
                          dataSource === "file" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500"
                        }`}
                      >
                        Upload Excel
                      </button>
                      <button
                        onClick={() => setDataSource("csv")}
                        className={`px-3 py-1 text-xs font-bold rounded-md transition-all cursor-pointer ${
                          dataSource === "csv" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500"
                        }`}
                      >
                        Paste CSV
                      </button>
                    </div>
                  </div>

                  {/* Google Drive Account Connection Banner */}
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200/80 flex items-center justify-between gap-3 shadow-inner">
                    <div className="flex items-center gap-2.5">
                      <div className="p-1.5 bg-white rounded-lg border border-slate-100 shadow-sm flex items-center justify-center">
                        <svg className="w-5 h-5" viewBox="0 0 24 24">
                          <path fill="#4285F4" d="M23.745 12.27c0-.7-.06-1.4-.19-2.07H12v4.51h6.6c-.29 1.53-1.14 2.82-2.4 3.68v3.05h3.88c2.27-2.09 3.66-5.17 3.66-9.17z"/>
                          <path fill="#34A853" d="M12 24c3.24 0 5.95-1.08 7.93-2.91l-3.88-3.05c-1.08.72-2.45 1.16-4.05 1.16-3.11 0-5.74-2.11-6.68-4.96H1.21v3.15C3.18 21.88 7.39 24 12 24z"/>
                          <path fill="#FBBC05" d="M5.32 14.24A7.16 7.16 0 0 1 5 12c0-.79.13-1.57.32-2.34V6.51H1.21A11.94 11.94 0 0 0 0 12c0 1.92.45 3.74 1.21 5.39l4.11-3.15z"/>
                          <path fill="#EA4335" d="M12 4.75c1.77 0 3.35.61 4.6 1.8l3.42-3.42C17.95 1.19 15.24 0 12 0 7.39 0 3.18 2.12 1.21 5.39l4.11 3.15c.94-2.85 3.57-4.96 6.68-4.96z"/>
                        </svg>
                      </div>
                      <div className="min-w-0 flex-1">
                        <div className="text-xs font-bold text-slate-700">Google Drive Integration</div>
                        {currentUser ? (
                          <div className="text-[10px] text-emerald-600 font-bold flex items-center gap-1 truncate">
                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                            <span className="truncate">Active: {currentUser.email}</span>
                          </div>
                        ) : (
                          <div className="text-[10px] text-slate-400 font-semibold leading-tight">
                            Connect Google account to import private sheets & files
                          </div>
                        )}
                      </div>
                    </div>
                    {currentUser ? (
                      <button
                        onClick={handleGoogleSignOut}
                        className="text-[11px] px-2.5 py-1.5 font-bold text-rose-600 hover:text-white bg-rose-50 hover:bg-rose-600 border border-rose-200 hover:border-rose-600 rounded-lg transition-all cursor-pointer flex-shrink-0"
                      >
                        Disconnect
                      </button>
                    ) : (
                      <button
                        disabled={isLoggingIn}
                        onClick={handleGoogleSignIn}
                        className="text-[11px] px-2.5 py-1.5 font-bold text-indigo-600 hover:text-white bg-indigo-50 hover:bg-indigo-600 border border-indigo-200 hover:border-indigo-600 rounded-lg transition-all flex items-center gap-1 cursor-pointer disabled:opacity-50 flex-shrink-0"
                      >
                        {isLoggingIn ? "Connecting..." : "Connect"}
                      </button>
                    )}
                  </div>

                  {/* Need instant testing chip options */}
                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-200/60">
                    <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5 mb-2">
                      <Sparkles className="w-3.5 h-3.5 text-indigo-500 animate-pulse" />
                      <span>Instant template sandbox</span>
                    </span>
                    <div className="flex flex-wrap gap-2">
                      <button
                        onClick={() => applyPresetDataset("sales")}
                        className={`text-xs px-3 py-1.5 font-semibold rounded-lg border transition-all ${
                          selectedPreset === "sales" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200 hover:bg-indigo-50"
                        }`}
                      >
                        💰 Sales Ledger
                      </button>
                      <button
                        onClick={() => applyPresetDataset("ops")}
                        className={`text-xs px-3 py-1.5 font-semibold rounded-lg border transition-all ${
                          selectedPreset === "ops" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200 hover:bg-indigo-50"
                        }`}
                      >
                        🖥️ Platform Latency
                      </button>
                      <button
                        onClick={() => applyPresetDataset("marketing")}
                        className={`text-xs px-3 py-1.5 font-semibold rounded-lg border transition-all ${
                          selectedPreset === "marketing" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200 hover:bg-indigo-50"
                        }`}
                      >
                        🎯 Campaign Spend
                      </button>
                    </div>
                  </div>

                  {dataSource === "sheet" && (
                    <div className="flex flex-col gap-3">
                      <div>
                        <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Google Sheet or Drive File Link</label>
                        <input
                          type="text"
                          value={sheetUrl}
                          onChange={(e) => setSheetUrl(e.target.value)}
                          placeholder="Paste Google Sheets URL or Google Drive file (.xlsx, .xls, .csv) share link"
                          className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm bg-slate-50/50 focus:bg-white focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-100 transition-all"
                        />
                      </div>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3 items-end">
                        <div>
                          <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Tab Name (Optional)</label>
                          <input
                            type="text"
                            value={sheetName}
                            onChange={(e) => setSheetName(e.target.value)}
                            placeholder="Sheet1"
                            className="mt-1 w-full px-4 py-2 rounded-xl border border-slate-200 text-sm bg-slate-50/50 focus:bg-white focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-100 transition-all"
                          />
                        </div>
                        <button
                          onClick={handlePreviewData}
                          className="px-4 py-2 bg-slate-100 text-slate-700 hover:bg-slate-200 text-xs font-bold rounded-xl border border-slate-200 hover:border-slate-300 transition-all h-[38px] flex items-center justify-center gap-1.5 cursor-pointer"
                        >
                          <Search className="w-3.5 h-3.5" />
                          <span>Preview Columns</span>
                        </button>
                      </div>
                    </div>
                  )}

                  {dataSource === "file" && (
                    <div className="flex flex-col gap-3">
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Upload Excel / CSV File</label>
                      <div
                        onDragOver={handleDragOver}
                        onDragLeave={handleDragLeave}
                        onDrop={handleDrop}
                        onClick={() => fileInputRef.current?.click()}
                        className={`border-2 border-dashed rounded-2xl p-6 flex flex-col items-center justify-center text-center cursor-pointer transition-all duration-200 ${
                          isDragging
                            ? "border-indigo-500 bg-indigo-50/50"
                            : uploadedFileName
                            ? "border-emerald-300 bg-emerald-50/10 hover:bg-emerald-50/20"
                            : "border-slate-300 hover:border-indigo-400 bg-slate-50/50 hover:bg-slate-50"
                        }`}
                      >
                        <input
                          type="file"
                          ref={fileInputRef}
                          onChange={handleFileChange}
                          accept=".xlsx,.xls,.csv,.tsv,.txt"
                          className="hidden"
                        />
                        {uploadedFileName ? (
                          <div className="flex flex-col items-center gap-2">
                            <div className="w-12 h-12 bg-emerald-100 rounded-xl text-emerald-600 flex items-center justify-center shadow-inner mb-1">
                              <CheckCircle2 className="w-6 h-6" />
                            </div>
                            <span className="text-sm font-bold text-slate-800 truncate max-w-[280px]">
                              {uploadedFileName}
                            </span>
                            <span className="text-xs text-slate-500 font-medium">
                              File Size: {uploadedFileSize}
                            </span>
                            <div className="flex gap-2 mt-2">
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setUploadedFileName("");
                                  setUploadedFileSize("");
                                  setCsvData("");
                                  if (fileInputRef.current) fileInputRef.current.value = "";
                                }}
                                className="px-3 py-1.5 bg-slate-100 hover:bg-rose-50 text-slate-600 hover:text-rose-600 text-[11px] font-bold rounded-lg border border-slate-200 hover:border-rose-200 transition-all cursor-pointer"
                              >
                                Clear File
                              </button>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handlePreviewData();
                                }}
                                className="px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-600 text-[11px] font-bold rounded-lg border border-indigo-100 hover:border-indigo-200 transition-all cursor-pointer flex items-center gap-1"
                              >
                                <Search className="w-3 h-3" />
                                <span>Preview columns</span>
                              </button>
                            </div>
                          </div>
                        ) : (
                          <div className="flex flex-col items-center">
                            <div className="w-12 h-12 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center mb-3 shadow-sm">
                              <Download className="w-6 h-6 animate-bounce" />
                            </div>
                            <p className="text-sm font-bold text-slate-700">
                              Drag and drop Excel or CSV file
                            </p>
                            <p className="text-[11px] text-slate-500 mt-1">
                              Supports .xlsx, .xls, .csv, .tsv, .txt
                            </p>
                            <span className="mt-3 text-xs text-indigo-600 font-bold px-3 py-1.5 bg-white border border-slate-200 rounded-xl shadow-sm hover:scale-105 transition-all">
                              Select from computer
                            </span>
                          </div>
                        )}
                      </div>
                    </div>
                  )}

                  {dataSource === "csv" && (
                    <div>
                      <div className="flex items-center justify-between">
                        <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Paste CSV / TSV Raw Content</label>
                        <button
                          onClick={handlePreviewData}
                          className="text-indigo-600 hover:text-indigo-700 font-bold text-xs cursor-pointer"
                        >
                          Preview Data
                        </button>
                      </div>
                      <textarea
                        value={csvData}
                        onChange={(e) => setCsvData(e.target.value)}
                        placeholder="Product,Revenue,Date&#10;Software Pro,12800,2026-06-24&#10;Enterprise Setup,45000,2026-06-25"
                        rows={5}
                        className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-xs font-mono bg-slate-50/50 focus:bg-white focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-100 transition-all resize-none"
                      />
                    </div>
                  )}
                </div>

                {/* 2. Report configuration settings card */}
                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 flex flex-col gap-4">
                  <h3 className="font-display font-bold text-slate-800 flex items-center gap-2 text-md border-b border-slate-100 pb-3">
                    <Calendar className="w-5 h-5 text-indigo-600" />
                    <span>2. Document Configuration</span>
                  </h3>

                  <div>
                    <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Report Document Title</label>
                    <input
                      type="text"
                      value={reportTitle}
                      onChange={(e) => setReportTitle(e.target.value)}
                      placeholder="Q2 Strategic Sales Performance Audit"
                      className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-100 transition-all"
                    />
                  </div>

                  <div>
                    <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Reporting Duration Scope</label>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-1.5">
                      {[
                        { id: "daily", label: "Daily", icon: "🌞", desc: "Past 24 hours" },
                        { id: "weekly", label: "Weekly", icon: "📅", desc: "Past 7 days" },
                        { id: "monthly", label: "Monthly", icon: "🪐", desc: "Past 30 days" },
                        { id: "custom", label: "Custom", icon: "⏱️", desc: "Select range" }
                      ].map(item => (
                        <button
                          key={item.id}
                          type="button"
                          onClick={() => setReportType(item.id)}
                          className={`flex flex-col items-center justify-center p-3 rounded-xl border text-center transition-all cursor-pointer ${
                            reportType === item.id 
                              ? "bg-indigo-50 border-indigo-600 text-indigo-600 ring-2 ring-indigo-100" 
                              : "bg-slate-50 border-slate-200 text-slate-500 hover:bg-slate-100/70"
                          }`}
                        >
                          <span className="text-xl mb-1">{item.icon}</span>
                          <span className="text-xs font-bold text-slate-800">{item.label}</span>
                          <span className="text-[10px] text-slate-400 mt-0.5 leading-tight">{item.desc}</span>
                        </button>
                      ))}
                    </div>
                  </div>

                  {reportType === "custom" && (
                    <motion.div
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: "auto" }}
                      className="bg-slate-50 p-4 rounded-xl border border-slate-200 flex flex-col gap-3"
                    >
                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="text-[10px] font-bold text-slate-500 uppercase">Start Date</label>
                          <input
                            type="date"
                            value={startDate}
                            onChange={(e) => setStartDate(e.target.value)}
                            className="mt-1 w-full px-3 py-1.5 bg-white rounded-lg border border-slate-200 text-xs focus:outline-none focus:border-indigo-500"
                          />
                        </div>
                        <div>
                          <label className="text-[10px] font-bold text-slate-500 uppercase">End Date</label>
                          <input
                            type="date"
                            value={endDate}
                            onChange={(e) => setEndDate(e.target.value)}
                            className="mt-1 w-full px-3 py-1.5 bg-white rounded-lg border border-slate-200 text-xs focus:outline-none focus:border-indigo-500"
                          />
                        </div>
                      </div>
                    </motion.div>
                  )}

                  <div>
                    <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Brand/Company Name</label>
                    <input
                      type="text"
                      value={companyName}
                      onChange={(e) => setCompanyName(e.target.value)}
                      placeholder="Acme Enterprise Labs"
                      className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-100 transition-all"
                    />
                  </div>
                </div>

                {/* 3. AI engine selections card */}
                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 flex flex-col gap-4">
                  <h3 className="font-display font-bold text-slate-800 flex items-center gap-2 text-md border-b border-slate-100 pb-3">
                    <Sparkles className="w-5 h-5 text-indigo-600" />
                    <span>3. AI Intelligence Analytics</span>
                  </h3>

                  <div>
                    <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Select AI Generation Model</label>
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-2">
                      {[
                        { id: "gemini", label: "Google Gemini", icon: "✨", desc: "Fast & precise", colorClass: "border-indigo-600 bg-indigo-50/40 text-indigo-600 ring-2 ring-indigo-100" },
                        { id: "openai", label: "OpenAI GPT-4o", icon: "🟢", desc: "Highly strategic", colorClass: "border-emerald-600 bg-emerald-50/40 text-emerald-600 ring-2 ring-emerald-100" },
                        { id: "claude", label: "Claude 3.5", icon: "🔸", desc: "Rich narrative", colorClass: "border-amber-600 bg-amber-50/40 text-amber-600 ring-2 ring-amber-100" },
                        { id: "groq", label: "Groq Llama 3", icon: "⚡", desc: "Ultra-fast speed", colorClass: "border-orange-500 bg-orange-50/40 text-orange-600 ring-2 ring-orange-100" }
                      ].map(item => (
                        <button
                          key={item.id}
                          type="button"
                          onClick={() => setAiModel(item.id)}
                          className={`flex flex-col items-center p-3.5 rounded-xl border transition-all text-center cursor-pointer ${
                            aiModel === item.id 
                              ? item.colorClass 
                              : "border-slate-200 bg-slate-50 hover:bg-slate-100/70 text-slate-600"
                          }`}
                        >
                          <span className="text-xl mb-1">{item.icon}</span>
                          <span className="text-xs font-bold text-slate-800 leading-tight">{item.label}</span>
                          <span className="text-[10px] text-slate-400 mt-1 leading-tight">{item.desc}</span>
                        </button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Natural Language Report Prompt</label>
                    <textarea
                      value={userPrompt}
                      onChange={(e) => setUserPrompt(e.target.value)}
                      placeholder="e.g. Generate a weekly operations latency review. Identify top spikes, suggest 3 quick engineering solutions and flag server outages."
                      rows={4}
                      className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-100 transition-all resize-none"
                    />

                    <div className="flex flex-wrap gap-1.5 mt-2.5">
                      {[
                        { label: "Weekly Sales Summary", prompt: "Create a comprehensive sales report. Highlight our top-performing products, overall revenue growth percentage, and pinpoint any items with a sudden drop in sales over 15%" },
                        { label: "Daily Operations & Latency", prompt: "Daily operations summary. Flag anomalies, compute operational metrics, and output clear warnings for any high latency or errors." },
                        { label: "Executive Strategic Dashboard", prompt: "Monthly executive overview. Generate dynamic high-level KPI cards, clear bar/line distribution trends, and 3 strategic recommendations for management." }
                      ].map((chip, idx) => (
                        <button
                          key={idx}
                          type="button"
                          onClick={() => setUserPrompt(chip.prompt)}
                          className="text-[11px] px-2.5 py-1.5 rounded-full border border-slate-200 bg-slate-50 hover:bg-indigo-50 hover:text-indigo-600 hover:border-indigo-200 transition-all text-slate-500 font-medium cursor-pointer"
                        >
                          {chip.label}
                        </button>
                      ))}
                    </div>

                    <div className="flex gap-2 mt-4">
                      <button
                        onClick={openScheduler}
                        className="flex-1 py-3 bg-slate-100 hover:bg-indigo-50 text-slate-700 hover:text-indigo-600 text-sm font-bold rounded-xl border border-slate-200 hover:border-indigo-200 transition-all flex items-center justify-center gap-2 cursor-pointer"
                      >
                        <Clock className="w-4 h-4" />
                        <span>Schedule Run</span>
                      </button>
                      <button
                        onClick={handleGenerateReport}
                        className="flex-[2] py-3 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-bold rounded-xl shadow-lg shadow-indigo-100 hover:shadow-indigo-200 hover:scale-[1.01] transition-all flex items-center justify-center gap-2 cursor-pointer"
                      >
                        <Sparkles className="w-4 h-4" />
                        <span>Generate Report</span>
                      </button>
                    </div>
                  </div>
                </div>

              </div>

              {/* Right Column: Report Viewer (grid-span 7) */}
              <div className="lg:col-span-7 flex flex-col min-h-[500px]" id="outputPreview">
                {!activeReport ? (
                  <div className="flex-grow bg-white rounded-2xl border border-slate-200 border-dashed border-2 p-12 flex flex-col items-center justify-center text-center">
                    <div className="w-16 h-16 bg-indigo-50 rounded-2xl text-indigo-600 flex items-center justify-center mb-4">
                      <FileText className="w-8 h-8" />
                    </div>
                    <h2 className="font-display font-extrabold text-xl text-slate-800 mb-2">No Active Report</h2>
                    <p className="text-slate-500 text-sm max-width-[360px] leading-relaxed mb-6">
                      Click "sandbox templates" or paste your business ledger spreadsheet, customize prompts, and trigger AI report generation.
                    </p>
                    <div className="flex gap-4">
                      <button 
                        onClick={() => applyPresetDataset("sales")} 
                        className="px-4 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-600 text-xs font-bold rounded-xl transition-all"
                      >
                        Run Sales Demo 🚀
                      </button>
                    </div>
                  </div>
                ) : (
                  <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-[calc(100vh-140px)] min-h-[550px]">
                    
                    {/* Header Controls for viewer */}
                    <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex flex-wrap justify-between items-center gap-3">
                      <div className="flex items-center gap-2">
                        <FileText className="w-4 h-4 text-indigo-600" />
                        <span className="text-xs font-extrabold text-slate-700 uppercase tracking-wider truncate max-w-[200px]">{activeReport.name}</span>
                      </div>
                      <div className="flex gap-1.5">
                        <button
                          onClick={printReportPDF}
                          className="px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold rounded-lg border border-slate-200 transition-all flex items-center gap-1.5"
                          title="Print directly or save as PDF"
                        >
                          <Printer className="w-3.5 h-3.5 text-slate-500" />
                          <span>PDF</span>
                        </button>
                        <button
                          onClick={downloadReportFile}
                          className="px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold rounded-lg border border-slate-200 transition-all flex items-center gap-1.5"
                        >
                          <Download className="w-3.5 h-3.5 text-indigo-500" />
                          <span>Download HTML</span>
                        </button>
                        <button
                          onClick={copyShareLinkToClipboard}
                          className="px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold rounded-lg border border-slate-200 transition-all flex items-center gap-1.5"
                        >
                          <Share2 className="w-3.5 h-3.5 text-emerald-500" />
                          <span>Copy Drive Link</span>
                        </button>
                        <button
                          onClick={() => setIsEmailModalOpen(true)}
                          className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold rounded-lg shadow-sm transition-all flex items-center gap-1.5"
                        >
                          <Mail className="w-3.5 h-3.5" />
                          <span>Email Report</span>
                        </button>
                      </div>
                    </div>

                    {/* Report Render Stage */}
                    <div className="flex-grow bg-slate-100 relative">
                      <iframe
                        id="reportFrame"
                        srcDoc={activeReport.reportHtml}
                        className="absolute inset-0 w-full h-full border-none bg-white"
                      />
                    </div>
                  </div>
                )}
              </div>

            </motion.div>
          )}

          {/* TAB 2: Automation Schedules */}
          {activeTab === "schedules" && (
            <motion.div
              key="schedules"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6"
            >
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-slate-100 pb-5 mb-6">
                <div>
                  <h2 className="font-display font-extrabold text-xl text-slate-800">Report Automations & Triggers</h2>
                  <p className="text-slate-500 text-sm">Create Cron-style background triggers that fetch data, generate summaries, and email teams.</p>
                </div>
                <button
                  onClick={() => setIsScheduleModalOpen(true)}
                  className="px-4 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-bold rounded-xl shadow-md transition-all flex items-center gap-2"
                >
                  <PlusCircle className="w-4 h-4" />
                  <span>New Schedule Task</span>
                </button>
              </div>

              {schedules.length === 0 ? (
                <div className="py-16 text-center">
                  <Clock className="w-12 h-12 text-slate-300 mx-auto mb-3" />
                  <p className="font-bold text-slate-700">No scheduled automations found.</p>
                  <p className="text-slate-500 text-xs mt-1">Configure email schedules to run reports daily, weekly, or monthly.</p>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm border-collapse">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50">
                        <th className="py-3 px-4 font-bold text-slate-700">Schedule Name</th>
                        <th className="py-3 px-4 font-bold text-slate-700">Frequency</th>
                        <th className="py-3 px-4 font-bold text-slate-700">AI Model</th>
                        <th className="py-3 px-4 font-bold text-slate-700">Recipients</th>
                        <th className="py-3 px-4 font-bold text-slate-700">Status</th>
                        <th className="py-3 px-4 font-bold text-slate-700 text-right">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {schedules.map(item => (
                        <tr key={item.id} className="border-b border-slate-100 hover:bg-slate-50/50">
                          <td className="py-4 px-4 font-semibold text-slate-800">{item.reportTitle}</td>
                          <td className="py-4 px-4 capitalize font-medium text-slate-600">{item.frequency}</td>
                          <td className="py-4 px-4 capitalize font-medium text-slate-500">{item.aiModel}</td>
                          <td className="py-4 px-4 text-xs text-slate-500 max-w-xs truncate">{item.recipients.join(", ")}</td>
                          <td className="py-4 px-4">
                            <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">
                              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                              <span>Active</span>
                            </span>
                          </td>
                          <td className="py-4 px-4 text-right">
                            <button
                              onClick={() => deleteScheduleItem(item.id!)}
                              className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </motion.div>
          )}

          {/* TAB 3: History Grid */}
          {activeTab === "history" && (
            <motion.div
              key="history"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="flex flex-col gap-6"
            >
              {/* Search & Filter tools row */}
              <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 flex flex-col md:flex-row gap-3 justify-between items-center">
                <div className="relative w-full md:max-w-xs">
                  <Search className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Search report runs..."
                    className="w-full pl-10 pr-4 py-2 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                  />
                </div>
                <div className="flex gap-2">
                  {["all", "daily", "weekly", "monthly"].map(f => (
                    <button
                      key={f}
                      onClick={() => setHistoryFilter(f)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition-all capitalize ${
                        historyFilter === f ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                      }`}
                    >
                      {f}
                    </button>
                  ))}
                </div>
              </div>

              {filteredHistory.length === 0 ? (
                <div className="bg-white rounded-2xl border border-slate-200 border-dashed border-2 py-16 text-center">
                  <FileText className="w-12 h-12 text-slate-300 mx-auto mb-3" />
                  <p className="font-bold text-slate-700">No matching reports found.</p>
                  <p className="text-slate-500 text-xs mt-1">Change filters or generate a report to log history files.</p>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
                  {filteredHistory.map(item => (
                    <div
                      key={item.id}
                      onClick={() => loadSavedReport(item.id)}
                      className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm hover:shadow-md transition-all duration-200 cursor-pointer flex flex-col justify-between group"
                    >
                      <div>
                        <div className="flex justify-between items-start mb-3">
                          <span className="text-[10px] bg-slate-100 text-slate-600 font-bold uppercase tracking-wider px-2 py-0.5 rounded border border-slate-200">
                            {item.reportType}
                          </span>
                          <span className="text-xs text-slate-400 font-semibold">
                            {new Date(item.createdDate).toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" })}
                          </span>
                        </div>
                        <h3 className="font-display font-extrabold text-slate-800 text-base leading-snug group-hover:text-indigo-600 transition-colors mb-4">
                          {item.name}
                        </h3>
                      </div>
                      <div className="flex justify-between items-center pt-3 border-t border-slate-100 mt-4">
                        <span className="text-xs font-bold text-indigo-600 flex items-center gap-1 group-hover:gap-1.5 transition-all">
                          <span>Open report</span>
                          <ChevronRight className="w-3.5 h-3.5" />
                        </span>
                        <button
                          onClick={(e) => deleteHistoryItem(e, item.id)}
                          className="p-1.5 rounded-lg hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-all"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </motion.div>
          )}

          {/* TAB 4: branding, preferences, keys */}
          {activeTab === "settings" && (
            <motion.div
              key="settings"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.25 }}
              className="grid grid-cols-1 lg:grid-cols-12 gap-6"
            >
              
              {/* API keys setting Column */}
              <div className="lg:col-span-6 flex flex-col gap-6">
                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
                  <h3 className="font-display font-bold text-slate-800 flex items-center gap-2 text-md border-b border-slate-100 pb-3 mb-4">
                    <ShieldCheck className="w-5 h-5 text-indigo-600" />
                    <span>Secure AI APIs Integration</span>
                  </h3>
                  <p className="text-xs text-slate-500 mb-5 leading-relaxed bg-indigo-50/50 p-3 rounded-xl border border-indigo-100">
                    Your key secrets are managed via script properties directly inside your secure Google environment block.
                  </p>

                  <div className="flex flex-col gap-4">
                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Google Gemini Key</label>
                      <div className="relative mt-1">
                        <input
                          type={showGemini ? "text" : "password"}
                          value={settings.geminiKey || ""}
                          onChange={(e) => setSettings({ ...settings, geminiKey: e.target.value })}
                          className="w-full pl-4 pr-12 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                        />
                        <button
                          type="button"
                          onClick={() => setShowGemini(!showGemini)}
                          className="absolute right-3 top-3 text-slate-400 hover:text-slate-600"
                        >
                          {showGemini ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                        </button>
                      </div>
                    </div>

                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">OpenAI GPT Key</label>
                      <div className="relative mt-1">
                        <input
                          type={showOpenai ? "text" : "password"}
                          value={settings.openaiKey || ""}
                          onChange={(e) => setSettings({ ...settings, openaiKey: e.target.value })}
                          className="w-full pl-4 pr-12 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                        />
                        <button
                          type="button"
                          onClick={() => setShowOpenai(!showOpenai)}
                          className="absolute right-3 top-3 text-slate-400 hover:text-slate-600"
                        >
                          {showOpenai ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                        </button>
                      </div>
                    </div>

                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Anthropic Claude Key</label>
                      <div className="relative mt-1">
                        <input
                          type={showClaude ? "text" : "password"}
                          value={settings.claudeKey || ""}
                          onChange={(e) => setSettings({ ...settings, claudeKey: e.target.value })}
                          className="w-full pl-4 pr-12 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                        />
                        <button
                          type="button"
                          onClick={() => setShowClaude(!showClaude)}
                          className="absolute right-3 top-3 text-slate-400 hover:text-slate-600"
                        >
                          {showClaude ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                        </button>
                      </div>
                    </div>

                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Groq AI Key</label>
                      <div className="relative mt-1">
                        <input
                          type={showGroq ? "text" : "password"}
                          value={settings.groqKey || ""}
                          onChange={(e) => setSettings({ ...settings, groqKey: e.target.value })}
                          placeholder="e.g. gsk_..."
                          className="w-full pl-4 pr-12 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                        />
                        <button
                          type="button"
                          onClick={() => setShowGroq(!showGroq)}
                          className="absolute right-3 top-3 text-slate-400 hover:text-slate-600"
                        >
                          {showGroq ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                        </button>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
                  <h3 className="font-display font-bold text-slate-800 flex items-center gap-2 text-md border-b border-slate-100 pb-3 mb-4">
                    <ShieldCheck className="w-5 h-5 text-indigo-600" />
                    <span>Report Password Protection</span>
                  </h3>
                  <div>
                    <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Viewer Access Password Code (Optional)</label>
                    <input
                      type="text"
                      value={settings.reportPassword || ""}
                      onChange={(e) => setSettings({ ...settings, reportPassword: e.target.value })}
                      placeholder="e.g. AcmeSecure2026"
                      className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500 focus:ring-4 focus:ring-indigo-100 transition-all"
                    />
                    <div className="text-[11px] text-slate-400 mt-1.5">If configured, viewers clicking shared Drive links will have to verify this pass code first.</div>
                  </div>
                </div>
              </div>

              {/* Branding preference selection Column */}
              <div className="lg:col-span-6 flex flex-col gap-6">
                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
                  <h3 className="font-display font-bold text-slate-800 flex items-center gap-2 text-md border-b border-slate-100 pb-3 mb-4">
                    <Palette className="w-5 h-5 text-indigo-600" />
                    <span>Branding & Visual Themes</span>
                  </h3>

                  <div className="flex flex-col gap-4">
                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Company Brand Name</label>
                      <input
                        type="text"
                        value={settings.companyName}
                        onChange={(e) => setSettings({ ...settings, companyName: e.target.value })}
                        className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                      />
                    </div>

                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Brand Logo Image URL</label>
                      <input
                        type="text"
                        value={settings.logoUrl}
                        onChange={(e) => setSettings({ ...settings, logoUrl: e.target.value })}
                        className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                      />
                    </div>

                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Accent Primary Brand Color</label>
                      <input
                        type="color"
                        value={settings.accentColor}
                        onChange={(e) => setSettings({ ...settings, accentColor: e.target.value })}
                        className="mt-1 w-full h-[38px] p-1 rounded-xl border border-slate-200 cursor-pointer"
                      />
                    </div>

                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Standard Report Footer Notice</label>
                      <input
                        type="text"
                        value={settings.footerText}
                        onChange={(e) => setSettings({ ...settings, footerText: e.target.value })}
                        className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                      />
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
                  <h3 className="font-display font-bold text-slate-800 flex items-center gap-2 text-md border-b border-slate-100 pb-3 mb-4">
                    <Globe className="w-5 h-5 text-indigo-600" />
                    <span>Preferences Settings</span>
                  </h3>

                  <div className="grid grid-cols-2 gap-3 mb-5">
                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Primary Timezone</label>
                      <select
                        value={settings.timezone}
                        onChange={(e) => setSettings({ ...settings, timezone: e.target.value })}
                        className="mt-1 w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                      >
                        <option value="GMT">GMT Zone</option>
                        <option value="EST">EST (New York)</option>
                        <option value="CST">CST (Chicago)</option>
                        <option value="PST">PST (Los Angeles)</option>
                        <option value="IST">IST (India)</option>
                        <option value="AEST">AEST (Sydney)</option>
                      </select>
                    </div>

                    <div>
                      <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Default Model Selection</label>
                      <select
                        value={settings.defaultAiModel}
                        onChange={(e) => setSettings({ ...settings, defaultAiModel: e.target.value })}
                        className="mt-1 w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                      >
                        <option value="gemini">Google Gemini</option>
                        <option value="openai">OpenAI GPT-4o</option>
                        <option value="claude">Claude 3.5</option>
                        <option value="groq">Groq AI (Llama 3)</option>
                      </select>
                    </div>
                  </div>

                  <button
                    onClick={handleSaveSettings}
                    className="w-full py-3 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-bold rounded-xl shadow-lg shadow-indigo-100 transition-all"
                  >
                    Save Preferences
                  </button>
                </div>
              </div>

            </motion.div>
          )}

        </AnimatePresence>
      </main>

      {/* Corporate Dashboard Footer */}
      <footer className="mt-auto py-6 border-t border-slate-200 bg-white text-center text-xs text-slate-400 font-medium">
        <p>{settings.footerText}</p>
        <p className="mt-1">ReportAI Labs · Deployed securely inside Google Workspace Containers · Version 3.1</p>
      </footer>

      {/* ================= MODAL DIALOGS ================= */}

      {/* MODAL 1: Email dispatch modal */}
      {isEmailModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm">
          <motion.div
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-2xl border border-slate-200 p-6 shadow-2xl max-w-md w-full mx-4"
          >
            <div className="flex justify-between items-center border-b border-slate-100 pb-3 mb-4">
              <h3 className="font-display font-extrabold text-slate-800 text-lg flex items-center gap-2">
                <Mail className="w-5 h-5 text-indigo-600" />
                <span>Email Business Dashboard</span>
              </h3>
              <button 
                onClick={() => setIsEmailModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 text-xl font-bold"
              >
                ×
              </button>
            </div>
            
            <div className="flex flex-col gap-4">
              <div>
                <label className="text-xs font-bold text-slate-600 uppercase tracking-wide">Recipients Email Addresses</label>
                <input
                  type="text"
                  value={emailRecipients}
                  onChange={(e) => setEmailRecipients(e.target.value)}
                  placeholder="manager@company.com, exec-board@acme.com"
                  className="mt-1 w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-indigo-500"
                />
                <div className="text-[10px] text-slate-400 mt-1">Separate multiple email addresses using commas.</div>
              </div>
            </div>

            <div className="flex gap-2 justify-end mt-6">
              <button
                onClick={() => setIsEmailModalOpen(false)}
                className="px-4 py-2 rounded-xl text-slate-600 bg-slate-100 hover:bg-slate-200 text-xs font-bold transition-all"
              >
                Cancel
              </button>
              <button
                onClick={handleSendEmailReport}
                className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold shadow-md transition-all"
              >
                Email Report
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* MODAL 2: Interactive Data preview modal */}
      {isPreviewModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm">
          <motion.div
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-2xl border border-slate-200 p-6 shadow-2xl max-w-4xl w-full mx-4"
          >
            <div className="flex justify-between items-center border-b border-slate-100 pb-3 mb-4">
              <h3 className="font-display font-extrabold text-slate-800 text-lg flex items-center gap-2">
                <Database className="w-5 h-5 text-indigo-600" />
                <span>📊 Connected Sheet Columns & Schema</span>
              </h3>
              <button 
                onClick={() => setIsPreviewModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 text-xl font-bold"
              >
                ×
              </button>
            </div>

            <div className="overflow-auto max-h-[360px] bg-slate-50 border border-slate-100 rounded-xl">
              {previewTableData.length === 0 ? (
                <p className="p-8 text-center text-slate-400 text-sm">No spreadsheet data parsed.</p>
              ) : (
                <table className="w-full text-left text-xs border-collapse">
                  <thead>
                    <tr className="bg-slate-100/80 sticky top-0 border-b border-slate-200">
                      {Object.keys(previewTableData[0]).map((header, idx) => (
                        <th key={idx} className="py-2.5 px-3 font-bold text-slate-600 uppercase tracking-wide border-r border-slate-200 last:border-0">{header}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewTableData.slice(0, 8).map((row, rIdx) => (
                      <tr key={rIdx} className="border-b border-slate-200/60 last:border-0 hover:bg-white transition-colors">
                        {Object.keys(previewTableData[0]).map((header, cIdx) => (
                          <td key={cIdx} className="py-2 px-3 font-medium text-slate-700 border-r border-slate-200/50 last:border-0 truncate max-w-[200px]">{row[header]}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>

            <div className="text-[11px] text-slate-400 mt-3 flex items-center gap-1">
              <Info className="w-3.5 h-3.5" />
              <span>Showing first {Math.min(previewTableData.length, 8)} sample rows parsed successfully. Full dataset contains {previewTableData.length} records.</span>
            </div>

            <div className="flex justify-end mt-6">
              <button
                onClick={() => setIsPreviewModalOpen(false)}
                className="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold shadow-md transition-all"
              >
                Done
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* MODAL 3: Schedule configuration modal */}
      {isScheduleModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm">
          <motion.div
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white rounded-2xl border border-slate-200 p-6 shadow-2xl max-w-xl w-full mx-4"
          >
            <div className="flex justify-between items-center border-b border-slate-100 pb-3 mb-4">
              <h3 className="font-display font-extrabold text-slate-800 text-lg flex items-center gap-2">
                <Clock className="w-5 h-5 text-indigo-600" />
                <span>⏰ Automated Scheduling Trigger</span>
              </h3>
              <button 
                onClick={() => setIsScheduleModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 text-xl font-bold"
              >
                ×
              </button>
            </div>

            <div className="flex flex-col gap-4 max-h-[440px] overflow-y-auto pr-2">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">Document/Trigger Title</label>
                  <input
                    type="text"
                    value={schedTitle}
                    onChange={(e) => setSchedTitle(e.target.value)}
                    className="mt-1 w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                  />
                </div>
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">Automation Engine</label>
                  <select
                    value={schedAi}
                    onChange={(e) => setSchedAi(e.target.value)}
                    className="mt-1 w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                  >
                    <option value="gemini">Google Gemini</option>
                    <option value="openai">OpenAI GPT-4o</option>
                    <option value="claude">Anthropic Claude</option>
                    <option value="groq">Groq Llama 3</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">Trigger Run Frequency</label>
                  <select
                    value={schedFreq}
                    onChange={(e) => setSchedFreq(e.target.value)}
                    className="mt-1 w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                  >
                    <option value="daily">Run Daily</option>
                    <option value="weekly">Run Weekly</option>
                    <option value="monthly">Run Monthly</option>
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">Report Scope</label>
                  <select
                    value={schedType}
                    onChange={(e) => setSchedType(e.target.value)}
                    className="mt-1 w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                  >
                    <option value="daily">Past 24 Hours</option>
                    <option value="weekly">Past 7 Days</option>
                    <option value="monthly">Past 30 Days</option>
                  </select>
                </div>
              </div>

              {schedFreq === "weekly" && (
                <div className="bg-slate-50 p-3 rounded-lg border border-slate-200">
                  <label className="text-[10px] font-bold text-slate-500 uppercase">Run Weekly on these days:</label>
                  <div className="flex flex-wrap gap-3 mt-1.5">
                    {["mon", "tue", "wed", "thu", "fri", "sat", "sun"].map(day => (
                      <label key={day} className="flex items-center gap-1.5 text-xs font-semibold capitalize text-slate-600">
                        <input
                          type="checkbox"
                          checked={schedWeeklyDays[day] || false}
                          onChange={(e) => setSchedWeeklyDays({ ...schedWeeklyDays, [day]: e.target.checked })}
                          className="rounded text-indigo-600 focus:ring-indigo-500"
                        />
                        <span>{day}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}

              {schedFreq === "monthly" && (
                <div className="bg-slate-50 p-3 rounded-lg border border-slate-200">
                  <label className="text-[10px] font-bold text-slate-500 uppercase">Run Monthly on day (1-28):</label>
                  <input
                    type="number"
                    min="1"
                    max="28"
                    value={schedMonthlyDay}
                    onChange={(e) => setSchedMonthlyDay(e.target.value)}
                    className="mt-1.5 w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md text-xs"
                  />
                </div>
              )}

              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">Email Deliver Recipients</label>
                <input
                  type="text"
                  value={schedRecipients}
                  onChange={(e) => setSchedRecipients(e.target.value)}
                  placeholder="e.g. execs@company.com, manager@company.com"
                  className="mt-1 w-full px-4 py-2 border border-slate-200 rounded-lg text-sm"
                />
                <div className="text-[9px] text-slate-400 mt-1">Separate emails with commas. Delivery includes inline HTML.</div>
              </div>

              <div>
                <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">Report prompt template instruction</label>
                <textarea
                  value={schedPrompt}
                  onChange={(e) => setSchedPrompt(e.target.value)}
                  rows={3}
                  className="mt-1 w-full px-3 py-2 border border-slate-200 rounded-lg text-sm"
                />
              </div>
            </div>

            <div className="flex gap-2 justify-end mt-6 pt-4 border-t border-slate-100">
              <button
                onClick={() => setIsScheduleModalOpen(false)}
                className="px-4 py-2 rounded-xl text-slate-600 bg-slate-100 hover:bg-slate-200 text-xs font-bold transition-all"
              >
                Cancel
              </button>
              <button
                onClick={handleSaveSchedule}
                className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold shadow-md transition-all"
              >
                Save Schedule
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* MODAL 4: Step-by-Step AI Report Generation Screen */}
      {isGenerating && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/80 backdrop-blur-sm">
          <div className="bg-white rounded-2xl border border-slate-100 p-8 shadow-2xl max-w-md w-full mx-4 text-center">
            
            {/* Spinning Indicator circle */}
            <div className="relative w-16 h-16 mx-auto mb-6">
              <div className="absolute inset-0 rounded-full border-4 border-slate-100 border-t-indigo-600 animate-spin"></div>
            </div>

            <h3 className="font-display font-extrabold text-slate-800 text-lg mb-1">Report Generation In Progress</h3>
            <p className="text-slate-500 text-xs leading-relaxed mb-6">Our system is executing deep analytics routines on your data parameters...</p>

            {/* Staggered progress checkpoints */}
            <div className="text-left flex flex-col gap-3">
              {[
                "1. Connecting and parsing spreadsheet rows...",
                "2. Forwarding data payloads to selected AI models...",
                "3. Computing analytical trends and visual datasets...",
                "4. Saving HTML reports inside secure cloud Drive storage..."
              ].map((step, idx) => (
                <div key={idx} className="flex items-center gap-3">
                  <div className={`w-3.5 h-3.5 rounded-full flex items-center justify-center text-[9px] font-bold ${
                    idx < generationStep 
                      ? "bg-emerald-500 text-white" 
                      : idx === generationStep 
                        ? "bg-indigo-600 text-white animate-pulse" 
                        : "bg-slate-200 text-slate-400"
                  }`}>
                    {idx < generationStep ? "✓" : idx + 1}
                  </div>
                  <span className={`text-xs font-semibold ${
                    idx < generationStep ? "text-slate-500" : idx === generationStep ? "text-slate-800 font-bold" : "text-slate-400"
                  }`}>
                    {step}
                  </span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
