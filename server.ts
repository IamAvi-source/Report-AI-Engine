import express from "express";
import path from "path";
import fs from "fs";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";
import * as XLSX from "xlsx";

dotenv.config();

const app = express();
const PORT = 3000;

// Setup JSON parsing limits for large spreadsheet datasets
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

// In-memory data store for the live server session (mocking Database/Drive)
interface GeneratedReport {
  id: string;
  name: string;
  createdDate: string;
  reportHtml: string;
  driveUrl: string;
  reportType: string;
}

interface ScheduleConfig {
  id: string;
  sheetUrl?: string;
  sheetName?: string;
  reportTitle: string;
  userPrompt: string;
  recipients: string[];
  frequency: string;
  aiModel: string;
  reportType: string;
  weeklyDays?: Record<string, boolean>;
  monthlyDay?: string;
  status: string;
  createdDate: string;
}

interface AppSettings {
  companyName: string;
  logoUrl: string;
  accentColor: string;
  footerText: string;
  timezone: string;
  defaultAiModel: string;
  defaultReportType: string;
  reportPassword?: string;
  claudeKey?: string;
  openaiKey?: string;
  geminiKey?: string;
  groqKey?: string;
}

let reportsHistory: GeneratedReport[] = [];
let schedulesList: ScheduleConfig[] = [];
let appSettings: AppSettings = {
  companyName: "ReportAI Labs",
  logoUrl: "",
  accentColor: "#4F46E5",
  footerText: "CONFIDENTIAL · Internal Use Only",
  timezone: "GMT",
  defaultAiModel: "gemini",
  defaultReportType: "weekly",
  groqKey: ""
};

// Seed some beautiful initial reports in history so the app has data out-of-the-box
reportsHistory = [
  {
    id: "rep-sales-01",
    name: "Acme Inc. Monthly Revenue Analysis",
    createdDate: new Date(Date.now() - 24 * 3600 * 1000 * 3).toISOString(), // 3 days ago
    reportType: "monthly",
    driveUrl: "#",
    reportHtml: `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Sales Performance</title>
        <style>
          body { font-family: sans-serif; padding: 30px; background: #F8F9FC; color: #1E293B; }
          .card { background: white; padding: 24px; border-radius: 12px; border: 1px solid #E2E8F0; }
          h1 { color: #4F46E5; }
        </style>
      </head>
      <body>
        <div class="card">
          <h1>📈 Acme Inc. Sales Performance Report</h1>
          <p>Generated on July 1, 2026</p>
          <hr style="margin: 20px 0; border: none; border-top: 1px solid #E2E8F0;" />
          <p><strong>Executive Summary:</strong> Q2 closed with record high margins driven by a 24% spike in software subscription sales. Hardware logistics incurred minor delayed delivery costs.</p>
          <ul>
            <li>Total Revenue: $248,500 (▲ 14.5%)</li>
            <li>Conversion Rate: 3.82% (▲ 0.4%)</li>
            <li>New Subscribers: 1,480 (▲ 22%)</li>
          </ul>
        </div>
      </body>
      </html>
    `
  },
  {
    id: "rep-ops-02",
    name: "System Reliability & Latency Audit",
    createdDate: new Date(Date.now() - 24 * 3600 * 1000).toISOString(), // Yesterday
    reportType: "daily",
    driveUrl: "#",
    reportHtml: `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Operations Audit</title>
        <style>
          body { font-family: sans-serif; padding: 30px; background: #F8F9FC; color: #1E293B; }
          .card { background: white; padding: 24px; border-radius: 12px; border: 1px solid #E2E8F0; }
          h1 { color: #10B981; }
        </style>
      </head>
      <body>
        <div class="card">
          <h1>🛡️ Server Reliability & Latency Audit</h1>
          <p>Generated on July 3, 2026</p>
          <hr style="margin: 20px 0; border: none; border-top: 1px solid #E2E8F0;" />
          <p><strong>Executive Summary:</strong> Average load time remained stable at 148ms. DB query thread locks spiked between 04:00 - 04:15 due to backup routines, causing an automatic warning status.</p>
          <ul>
            <li>Overall Uptime: 99.98% (Stable)</li>
            <li>Errors Logged: 12 (▼ 40%)</li>
            <li>Peak Traffic: 84,200 req/min (▲ 8%)</li>
          </ul>
        </div>
      </body>
      </html>
    `
  }
];

// Rich predefined template datasets so the user doesn't even need a Google Sheet to get started!
const MOCK_DATASETS: Record<string, any[]> = {
  sales: [
    { Date: "2026-06-25", Product: "Standard Subscription", Region: "North America", Units: 120, Price: 49, Revenue: 5880, Channel: "Direct Search" },
    { Date: "2026-06-25", Product: "Enterprise Plan", Region: "Europe", Units: 4, Price: 999, Revenue: 3996, Channel: "Sales Rep" },
    { Date: "2026-06-26", Product: "Standard Subscription", Region: "Asia", Units: 85, Price: 49, Revenue: 4165, Channel: "Ad Campaign" },
    { Date: "2026-06-26", Product: "Professional Tier", Region: "North America", Units: 45, Price: 149, Revenue: 6705, Channel: "Newsletter" },
    { Date: "2026-06-27", Product: "Standard Subscription", Region: "Europe", Units: 110, Price: 49, Revenue: 5390, Channel: "Direct Search" },
    { Date: "2026-06-27", Product: "Professional Tier", Region: "Asia", Units: 38, Price: 149, Revenue: 5662, Channel: "Ad Campaign" },
    { Date: "2026-06-28", Product: "Enterprise Plan", Region: "North America", Units: 6, Price: 999, Revenue: 5994, Channel: "Sales Rep" },
    { Date: "2026-06-28", Product: "Standard Subscription", Region: "Asia", Units: 92, Price: 49, Revenue: 4508, Channel: "Direct Search" },
    { Date: "2026-06-29", Product: "Professional Tier", Region: "Europe", Units: 50, Price: 149, Revenue: 7450, Channel: "Newsletter" },
    { Date: "2026-06-29", Product: "Standard Subscription", Region: "North America", Units: 135, Price: 49, Revenue: 6615, Channel: "Ad Campaign" },
    { Date: "2026-06-30", Product: "Professional Tier", Region: "North America", Units: 62, Price: 149, Revenue: 9238, Channel: "Direct Search" },
    { Date: "2026-06-30", Product: "Enterprise Plan", Region: "Asia", Units: 5, Price: 999, Revenue: 4995, Channel: "Sales Rep" }
  ],
  ops: [
    { Timestamp: "2026-07-03 00:00:00", Server: "Web-Node-01", CPU_Usage: 45, RAM_Usage: 62, Latency_ms: 120, Status: "Healthy" },
    { Timestamp: "2026-07-03 04:00:00", Server: "Web-Node-01", CPU_Usage: 89, RAM_Usage: 84, Latency_ms: 450, Status: "Warning" },
    { Timestamp: "2026-07-03 08:00:00", Server: "Web-Node-01", CPU_Usage: 55, RAM_Usage: 65, Latency_ms: 132, Status: "Healthy" },
    { Timestamp: "2026-07-03 12:00:00", Server: "Web-Node-01", CPU_Usage: 68, RAM_Usage: 71, Latency_ms: 140, Status: "Healthy" },
    { Timestamp: "2026-07-03 16:00:00", Server: "Web-Node-01", CPU_Usage: 72, RAM_Usage: 75, Latency_ms: 155, Status: "Healthy" },
    { Timestamp: "2026-07-03 20:00:00", Server: "Web-Node-01", CPU_Usage: 50, RAM_Usage: 64, Latency_ms: 118, Status: "Healthy" },
    { Timestamp: "2026-07-03 00:00:00", Server: "Web-Node-02", CPU_Usage: 41, RAM_Usage: 58, Latency_ms: 115, Status: "Healthy" },
    { Timestamp: "2026-07-03 04:00:00", Server: "Web-Node-02", CPU_Usage: 92, RAM_Usage: 81, Latency_ms: 480, Status: "Warning" },
    { Timestamp: "2026-07-03 08:00:00", Server: "Web-Node-02", CPU_Usage: 52, RAM_Usage: 61, Latency_ms: 125, Status: "Healthy" },
    { Timestamp: "2026-07-03 12:00:00", Server: "Web-Node-02", CPU_Usage: 61, RAM_Usage: 68, Latency_ms: 130, Status: "Healthy" },
    { Timestamp: "2026-07-03 16:00:00", Server: "Web-Node-02", CPU_Usage: 75, RAM_Usage: 74, Latency_ms: 162, Status: "Healthy" },
    { Timestamp: "2026-07-03 20:00:00", Server: "Web-Node-02", CPU_Usage: 48, RAM_Usage: 60, Latency_ms: 110, Status: "Healthy" }
  ],
  marketing: [
    { Date: "2026-06-24", Campaign: "Google Search Brand", Clicks: 2450, Conversions: 148, Spend: 1200, Revenue: 5920 },
    { Date: "2026-06-24", Campaign: "Retargeting Facebook", Clicks: 1120, Conversions: 95, Spend: 850, Revenue: 3800 },
    { Date: "2026-06-25", Campaign: "Google Search Brand", Clicks: 2610, Conversions: 152, Spend: 1250, Revenue: 6080 },
    { Date: "2026-06-25", Campaign: "Retargeting Facebook", Clicks: 980, Conversions: 74, Spend: 850, Revenue: 2960 },
    { Date: "2026-06-26", Campaign: "Google Search Brand", Clicks: 2890, Conversions: 171, Spend: 1400, Revenue: 6840 },
    { Date: "2026-06-26", Campaign: "Retargeting Facebook", Clicks: 1420, Conversions: 112, Spend: 1100, Revenue: 4480 },
    { Date: "2026-06-27", Campaign: "Google Search Brand", Clicks: 3100, Conversions: 184, Spend: 1500, Revenue: 7360 },
    { Date: "2026-06-27", Campaign: "Retargeting Facebook", Clicks: 1350, Conversions: 105, Spend: 1050, Revenue: 4200 },
    { Date: "2026-06-28", Campaign: "Google Search Brand", Clicks: 2950, Conversions: 168, Spend: 1450, Revenue: 6720 },
    { Date: "2026-06-28", Campaign: "Retargeting Facebook", Clicks: 1200, Conversions: 88, Spend: 1000, Revenue: 3520 }
  ]
};

// Helper: Fetch and Parse spreadsheet data from Google Sheets or Google Drive files (Excel, CSV)
async function fetchAndParseData(url: string, sheetName?: string, accessToken?: string): Promise<any[]> {
  let exportUrl = url.trim();
  let isGoogleSheets = false;
  let isGoogleDriveFile = false;
  let fileId = "";

  // 1. Resolve direct download / export URLs
  if (exportUrl.includes("docs.google.com/spreadsheets")) {
    isGoogleSheets = true;
    const match = exportUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      fileId = match[1];
      exportUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=csv`;
      const gidMatch = url.match(/[#&?]gid=([0-9]+)/);
      if (gidMatch && gidMatch[1]) {
        exportUrl += `&gid=${gidMatch[1]}`;
      } else if (sheetName) {
        exportUrl += `&sheet=${encodeURIComponent(sheetName)}`;
      }
    }
  } else if (exportUrl.includes("drive.google.com/file/d/") || exportUrl.includes("docs.google.com/file/d/")) {
    isGoogleDriveFile = true;
    const match = exportUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      fileId = match[1];
      if (accessToken) {
        exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
      } else {
        exportUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
      }
    }
  } else if (exportUrl.includes("drive.google.com/open?id=")) {
    isGoogleDriveFile = true;
    const match = exportUrl.match(/id=([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      fileId = match[1];
      if (accessToken) {
        exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
      } else {
        exportUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
      }
    }
  }

  // 2. Fetch the resource
  const headers: Record<string, string> = {};
  if (accessToken) {
    headers["Authorization"] = `Bearer ${accessToken}`;
  }

  const response = await fetch(exportUrl, { headers });
  if (!response.ok) {
    let extraHelp = "";
    if (response.status === 401 || response.status === 403) {
      if (!accessToken) {
        extraHelp = " This file or sheet is private. Click 'Sign in with Google' on the dashboard to access your files, or make the file public using Google Drive/Sheets share settings.";
      } else {
        extraHelp = " Access denied. Please ensure you have access to this file and your Google login is active.";
      }
    }
    throw new Error(`Failed to fetch sheet/file (Status: ${response.status}).${extraHelp}`);
  }

  // 3. Handle Native Google Sheets CSV Export format
  if (isGoogleSheets) {
    const csvText = await response.text();
    if (csvText.trim().startsWith("<!DOCTYPE") || csvText.trim().startsWith("<html")) {
      throw new Error("Failed to fetch Google Sheet. The response returned web HTML (login redirect), meaning it is private. Please click 'Sign in with Google' on the dashboard, or set access to 'Anyone with the link can view' and try again.");
    }
    const parsed = parseCSVString(csvText);
    if (!parsed || parsed.length === 0) {
      throw new Error("The fetched Google Sheet has no readable data rows.");
    }
    return parsed;
  }

  // 4. Handle other raw file formats (.xlsx, .xls, .csv, etc.) from Google Drive or general URLs
  const arrayBuffer = await response.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);

  // Check if it's HTML redirect (private link page)
  const previewString = buffer.toString("utf8", 0, 500).trim();
  if (previewString.startsWith("<!DOCTYPE") || previewString.startsWith("<html")) {
    throw new Error("The file on Google Drive returned an HTML page instead of raw file data. This indicates the file is private. Please click 'Sign in with Google' on the dashboard, or set permissions to 'Anyone with the link can view' and try again.");
  }

  // Parse using SheetJS (XLSX)
  try {
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const selectedSheet = sheetName && workbook.SheetNames.includes(sheetName)
      ? sheetName
      : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[selectedSheet];
    const parsedData = XLSX.utils.sheet_to_json(worksheet);
    if (parsedData && parsedData.length > 0) {
      return parsedData;
    }
  } catch (xlsxError: any) {
    // Fallback: try parsing as simple CSV text
    try {
      const text = buffer.toString("utf8");
      const parsed = parseCSVString(text);
      if (parsed && parsed.length > 0) {
        return parsed;
      }
    } catch (csvError) {}
    throw new Error(`Failed to parse file content: ${xlsxError.message}`);
  }

  throw new Error("The spreadsheet file contains no readable records or has an unsupported format.");
}

// ================= API ENDPOINTS =================

// 1. Fetch Google Sheet or return Mock Data
app.post("/api/sheet-data", async (req, res) => {
  const { sheetUrl, sheetName, mockType, accessToken } = req.body;
  
  // Return predefined mock dataset directly if sheetUrl is empty, is a mock/example keyword, or is a placeholder
  const isMockUrl = !sheetUrl || sheetUrl.includes("example") || sheetUrl === "mock";
  if (isMockUrl) {
    const selectedMock = mockType || "sales";
    const data = MOCK_DATASETS[selectedMock] || MOCK_DATASETS.sales;
    return res.json({ success: true, data });
  }

  try {
    const parsedData = await fetchAndParseData(sheetUrl, sheetName, accessToken);
    return res.json({ success: true, data: parsedData });
  } catch (err: any) {
    console.error("Sheet read error:", err);
    return res.status(400).json({
      success: false,
      error: err.message || "Error reading Google Sheet/File."
    });
  }
});

// 2. Generate Report Endpoint via Gemini AI
app.post("/api/generate-report", async (req, res) => {
  try {
    const config = req.body;
    let sheetData: any[] = [];
    let dataSourceName = "Sample Business Dataset";

    // Obtain dataset
    if (config.csvData) {
      sheetData = parseCSVString(config.csvData);
      dataSourceName = "Pasted Table/CSV Data";
    } else {
      const mock = config.sheetUrl ? null : (config.mockType || "sales");
      if (mock) {
        sheetData = MOCK_DATASETS[mock] || MOCK_DATASETS.sales;
        dataSourceName = `Sample ${mock.toUpperCase()} Data`;
      } else {
        try {
          sheetData = await fetchAndParseData(config.sheetUrl, config.sheetName, config.accessToken);
          dataSourceName = config.sheetName || "Connected Sheet Tab";
        } catch (fetchErr: any) {
          return res.status(400).json({
            success: false,
            error: `Failed to fetch/parse the connected sheet/file: ${fetchErr.message}`
          });
        }
      }
    }

    if (!sheetData || sheetData.length === 0) {
      sheetData = MOCK_DATASETS.sales;
    }

    // Limit dataset rows to fit token limits comfortably
    const filteredData = sheetData.slice(0, 500);

    const systemPrompt = `
      You are an expert business data analyst. You receive business data in JSON format 
      and a user's reporting instruction. Your ONLY output must be valid JSON — no markdown, 
      no explanation, no code fences, just raw JSON.

      Return a JSON object with EXACTLY this structure:
      {
        "title": "Report title string",
        "period": "e.g. Week of June 24–30, 2025",
        "executive_summary": "2-3 sentence plain English summary of the most important findings",
        "kpis": [
          {
            "label": "Metric name",
            "value": "Formatted value e.g. $12,400 or 847 units",
            "change": 12.5,
            "change_period": "vs last week",
            "direction": "up",
            "is_positive": true,
            "highlight": true
          }
        ],
        "charts": [
          {
            "id": "chart1",
            "type": "line",
            "title": "Chart title",
            "labels": ["Mon", "Tue", "Wed"],
            "datasets": [
              {
                "label": "Revenue",
                "data": [1200, 1400, 1100],
                "color": "#4F46E5"
              }
            ]
          },
          {
            "id": "chart2",
            "type": "bar",
            "title": "Second chart title",
            "labels": ["Product A", "Product B"],
            "datasets": [
              {
                "label": "Units Sold",
                "data": [340, 280],
                "color": "#10B981"
              }
            ]
          }
        ],
        "table": {
          "title": "Detailed breakdown",
          "headers": ["Column1", "Column2", "Column3"],
          "rows": [
            ["Row1Val1", "Row1Val2", "Row1Val3"]
          ],
          "sort_column": 0,
          "sort_direction": "desc"
        },
        "insights": [
          "Key finding 1 — specific, data-driven",
          "Key finding 2 — specific, data-driven", 
          "Key finding 3 — specific, data-driven"
        ],
        "recommendations": [
          "Actionable recommendation 1",
          "Actionable recommendation 2",
          "Actionable recommendation 3"
        ],
        "alerts": [
          {
            "severity": "warning",
            "title": "Alert title",
            "message": "Alert detail"
          }
        ]
      }

      Rules:
      - kpis: always include 4-6 KPIs
      - charts: always include exactly 2 charts
      - insights: always exactly 3
      - recommendations: always exactly 3
      - alerts: include only if there are genuine anomalies (can be empty array [])
      - All numbers must be real values derived from the provided data
      - direction must be "up", "down", or "neutral"
      - is_positive: true means the direction is good for business
    `;

    const userMessage = `
      DATA: ${JSON.stringify(filteredData)}

      REPORT INSTRUCTION: ${config.userPrompt}

      REPORT TYPE: ${config.reportType || "weekly"}
      DATE RANGE: ${config.startDate || "N/A"} to ${config.endDate || "N/A"}
      COMPANY: ${config.companyName || appSettings.companyName}

      Generate the report JSON now.
    `;

    let parsedAiData: any = null;

    if (config.aiModel === "groq") {
      const groqApiKey = appSettings.groqKey || process.env.GROQ_API_KEY;
      if (!groqApiKey) {
        throw new Error("GROQ_API_KEY is not configured. Please add it in Settings > Secure AI APIs Integration.");
      }

      const groqResponse = await fetch("https://api.groq.com/openai/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${groqApiKey}`
        },
        body: JSON.stringify({
          model: "llama-3.3-70b-versatile",
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: userMessage }
          ],
          response_format: { type: "json_object" },
          temperature: 0.1
        })
      });

      if (!groqResponse.ok) {
        const errorText = await groqResponse.text();
        throw new Error(`Groq API error: ${errorText}`);
      }

      const groqData = await groqResponse.json();
      const content = groqData.choices?.[0]?.message?.content || "";
      parsedAiData = cleanAndParseJSON(content);
    } else {
      // Call Gemini API using modern SDK
      const apiKey = appSettings.geminiKey || process.env.GEMINI_API_KEY;
      if (!apiKey) {
        throw new Error("GEMINI_API_KEY environment variable is not configured in the workspace secrets.");
      }

      const ai = new GoogleGenAI({
        apiKey: apiKey,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build",
          },
        },
      });

      // Query Gemini
      const response = await ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: userMessage,
        config: {
          systemInstruction: systemPrompt,
          responseMimeType: "application/json",
          temperature: 0.1,
        }
      });

      const responseText = response.text || "";
      parsedAiData = cleanAndParseJSON(responseText);
    }

    // Read the template HTML and replace placeholders
    const templatePath = path.join(process.cwd(), "Report.html");
    let reportHtml = "";
    if (fs.existsSync(templatePath)) {
      reportHtml = fs.readFileSync(templatePath, "utf-8");
    } else {
      // Fallback in-case template is missing
      reportHtml = "<html><body><h1>ReportAI Draft</h1></body></html>";
    }

    const generatedHtml = buildReportHTML(parsedAiData, config, reportHtml, dataSourceName, appSettings);

    // Save report in History
    const reportId = "rep-" + Date.now();
    const driveUrl = `https://drive.google.com/open?id=${reportId}`;
    const finalTitle = config.reportTitle || parsedAiData.title || "AI Business Report";
    
    reportsHistory.unshift({
      id: reportId,
      name: finalTitle,
      createdDate: new Date().toISOString(),
      reportHtml: generatedHtml,
      driveUrl: driveUrl,
      reportType: config.reportType || "weekly"
    });

    return res.json({
      success: true,
      reportId: reportId,
      reportHtml: generatedHtml,
      driveUrl: driveUrl,
      title: finalTitle
    });

  } catch (err: any) {
    console.error("Report generation error:", err);
    return res.status(500).json({ success: false, error: err.message || "Unknown error during generation." });
  }
});

// 3. Schedules endpoints
app.get("/api/schedules", (req, res) => {
  return res.json({ success: true, schedules: schedulesList });
});

app.post("/api/schedules", (req, res) => {
  const config = req.body;
  const newSchedule: ScheduleConfig = {
    ...config,
    id: "schedule-" + Date.now(),
    status: "active",
    createdDate: new Date().toISOString()
  };
  schedulesList.unshift(newSchedule);
  return res.json({ success: true, scheduleId: newSchedule.id });
});

app.delete("/api/schedules/:id", (req, res) => {
  const id = req.params.id;
  schedulesList = schedulesList.filter(s => s.id !== id);
  return res.json({ success: true });
});

// 4. Report History endpoints
app.get("/api/history", (req, res) => {
  // Omit the massive HTML payload from list responses for performance
  const list = reportsHistory.map(r => ({
    id: r.id,
    name: r.name,
    createdDate: r.createdDate,
    driveUrl: r.driveUrl,
    reportType: r.reportType
  }));
  return res.json({ success: true, history: list });
});

app.get("/api/history/:id", (req, res) => {
  const id = req.params.id;
  const found = reportsHistory.find(r => r.id === id);
  if (found) {
    return res.json({ success: true, report: found });
  }
  return res.status(404).json({ success: false, error: "Report not found." });
});

app.delete("/api/history/:id", (req, res) => {
  const id = req.params.id;
  reportsHistory = reportsHistory.filter(r => r.id !== id);
  return res.json({ success: true });
});

// 5. Settings endpoints
app.get("/api/settings", (req, res) => {
  return res.json({ success: true, settings: appSettings });
});

app.post("/api/settings", (req, res) => {
  appSettings = { ...appSettings, ...req.body };
  return res.json({ success: true, settings: appSettings });
});

// Helper: Custom CSV Parser
function parseCSVString(csvText: string): any[] {
  const lines: string[][] = [];
  let row = [""];
  let inQuotes = false;

  for (let i = 0; i < csvText.length; i++) {
    const c = csvText[i];
    const next = csvText[i + 1];

    if (c === '"') {
      if (inQuotes && next === '"') {
        row[row.length - 1] += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === "," && !inQuotes) {
      row.push("");
    } else if ((c === "\r" || c === "\n") && !inQuotes) {
      if (c === "\r" && next === "\n") {
        i++;
      }
      lines.push(row);
      row = [""];
    } else {
      row[row.length - 1] += c;
    }
  }

  if (row.length > 1 || row[0] !== "") {
    lines.push(row);
  }

  if (lines.length <= 1) return [];

  const headers = lines[0].map(h => h.trim());
  const data: any[] = [];

  for (let r = 1; r < lines.length; r++) {
    const vals = lines[r];
    const rowObj: any = {};
    let hasVals = false;

    for (let col = 0; col < headers.length; col++) {
      const header = headers[col] || `Column_${col}`;
      const val = vals[col] ? vals[col].trim() : "";
      rowObj[header] = val;
      if (val !== "") hasVals = true;
    }

    if (hasVals) {
      data.push(rowObj);
    }
  }

  return data;
}

// Helper: AI Text Cleaner
function cleanAndParseJSON(responseText: string): any {
  let cleaned = responseText.trim();
  if (cleaned.startsWith("```json")) {
    cleaned = cleaned.substring(7);
  } else if (cleaned.startsWith("```")) {
    cleaned = cleaned.substring(3);
  }

  if (cleaned.endsWith("```")) {
    cleaned = cleaned.substring(0, cleaned.length - 3);
  }

  cleaned = cleaned.trim();
  return JSON.parse(cleaned);
}

// Helper: Hex color opacity convertor
function hexToRgbA(hex: string, alpha: number): string {
  let c: any;
  if (/^#([A-Fa-f0-9]{3}){1,2}$/.test(hex)) {
    c = hex.substring(1).split("");
    if (c.length === 3) {
      c = [c[0], c[0], c[1], c[1], c[2], c[2]];
    }
    c = "0x" + c.join("");
    return "rgba(" + [(c >> 16) & 255, (c >> 8) & 255, c & 255].join(",") + "," + alpha + ")";
  }
  return hex;
}

// Helper: HTML Report builder
function buildReportHTML(aiData: any, config: any, templateHtml: string, dataSourceName: string, settings: AppSettings): string {
  const company = config.companyName || settings.companyName || "Acme Brand";
  const accentColor = settings.accentColor || "#4F46E5";

  // Process KPI cards
  let kpiCardsHtml = "";
  if (aiData.kpis && aiData.kpis.length > 0) {
    aiData.kpis.forEach((kpi: any) => {
      const dirClass = kpi.direction === "up" ? "up" : (kpi.direction === "down" ? "down" : "neutral");
      const isPosClass = kpi.is_positive ? "positive" : "negative";
      const highlightClass = kpi.highlight ? "highlight" : "";

      let changeHtml = "";
      if (kpi.change !== undefined && kpi.direction !== "neutral") {
        const sign = kpi.direction === "up" ? "▲" : "▼";
        changeHtml = `<div class="kpi-change ${dirClass} ${isPosClass}">
                       <span>${sign}</span> ${Math.abs(kpi.change)}%
                     </div>`;
      } else {
        changeHtml = '<div class="kpi-change neutral">Neutral</div>';
      }

      kpiCardsHtml += `
        <div class="kpi-card ${highlightClass}">
          <div class="kpi-label">${kpi.label}</div>
          <div class="kpi-value">${kpi.value}</div>
          <div class="kpi-change-row">
            ${changeHtml}
            <span class="kpi-period">${kpi.change_period || ""}</span>
          </div>
        </div>`;
    });
  }

  // Process Alerts
  let alertsHtml = "";
  if (aiData.alerts && aiData.alerts.length > 0) {
    aiData.alerts.forEach((alert: any) => {
      const sevClass = alert.severity === "critical" ? "critical" : (alert.severity === "warning" ? "warning" : "info");
      const icon = sevClass === "critical" ? "🚨" : (sevClass === "warning" ? "⚠️" : "ℹ️");

      alertsHtml += `
        <div class="alert ${sevClass}">
          <span class="alert-icon">${icon}</span>
          <div class="alert-content">
            <div class="alert-title">${alert.title}</div>
            <div>${alert.message}</div>
          </div>
        </div>`;
    });
  }

  // Process Table
  let tableHeadersHtml = "";
  let tableRowsHtml = "";
  if (aiData.table) {
    if (aiData.table.headers) {
      aiData.table.headers.forEach((header: string) => {
        tableHeadersHtml += `<th>${header}</th>`;
      });
    }
    if (aiData.table.rows) {
      aiData.table.rows.forEach((row: any[]) => {
        tableRowsHtml += "<tr>";
        row.forEach((cell: any) => {
          tableRowsHtml += `<td>${cell}</td>`;
        });
        tableRowsHtml += "</tr>";
      });
    }
  }

  // Process Insights List
  let insightsHtml = "";
  if (aiData.insights) {
    aiData.insights.forEach((insight: string, idx: number) => {
      insightsHtml += `
        <div class="insight-item">
          <div class="insight-num">${idx + 1}</div>
          <div>${insight}</div>
        </div>`;
    });
  }

  // Process Recommendations List
  let recommendationsHtml = "";
  if (aiData.recommendations) {
    aiData.recommendations.forEach((rec: string) => {
      recommendationsHtml += `
        <div class="rec-item">
          <span class="rec-arrow">➔</span>
          <div>${rec}</div>
        </div>`;
    });
  }

  // Get charts configuration
  let chart1Title = "Chart 1";
  let chart1Type = "line";
  let chart1Labels = "";
  let chart1Datasets = "";

  let chart2Title = "Chart 2";
  let chart2Type = "bar";
  let chart2Labels = "";
  let chart2Datasets = "";

  if (aiData.charts && aiData.charts.length > 0) {
    const c1 = aiData.charts[0];
    chart1Title = c1.title || "Metric Trend";
    chart1Type = c1.type || "line";
    chart1Labels = (c1.labels || []).map((l: string) => `"${l}"`).join(",");

    const datasets1: string[] = [];
    if (c1.datasets) {
      c1.datasets.forEach((ds: any) => {
        const col = ds.color || accentColor;
        datasets1.push(JSON.stringify({
          label: ds.label,
          data: ds.data,
          backgroundColor: hexToRgbA(col, 0.08),
          borderColor: col,
          borderWidth: 2.5,
          tension: 0.35,
          fill: chart1Type !== "bar"
        }));
      });
    }
    chart1Datasets = datasets1.join(",");
  }

  if (aiData.charts && aiData.charts.length > 1) {
    const c2 = aiData.charts[1];
    chart2Title = c2.title || "Metric Distribution";
    chart2Type = c2.type || "bar";
    chart2Labels = (c2.labels || []).map((l: string) => `"${l}"`).join(",");

    const datasets2: string[] = [];
    if (c2.datasets) {
      c2.datasets.forEach((ds: any) => {
        const col = ds.color || "#10B981";
        datasets2.push(JSON.stringify({
          label: ds.label,
          data: ds.data,
          backgroundColor: hexToRgbA(col, 0.15),
          borderColor: col,
          borderWidth: 2,
          tension: 0.35,
          fill: chart2Type !== "bar"
        }));
      });
    }
    chart2Datasets = datasets2.join(",");
  }

  // Replace placeholders in the HTML Report template
  return templateHtml
    .replace("[COMPANY_NAME]", company)
    .replace("[TITLE]", aiData.title || config.reportTitle || "AI Business Report")
    .replace("[PERIOD]", aiData.period || config.reportType || "Analysis Report")
    .replace("[GENERATED_DATE]", new Date().toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" }))
    .replace("[EXECUTIVE_SUMMARY]", aiData.executive_summary || "Report analysis ready.")
    .replace("[ALERTS_HTML]", alertsHtml)
    .replace("[KPI_CARDS_HTML]", kpiCardsHtml)
    .replace("[CHART1_TITLE]", chart1Title)
    .replace("[CHART1_TYPE]", chart1Type)
    .replace("[CHART1_LABELS]", chart1Labels)
    .replace("[CHART1_DATASETS]", chart1Datasets)
    .replace("[CHART2_TITLE]", chart2Title)
    .replace("[CHART2_TYPE]", chart2Type)
    .replace("[CHART2_LABELS]", chart2Labels)
    .replace("[CHART2_DATASETS]", chart2Datasets)
    .replace("[TABLE_TITLE]", (aiData.table && aiData.table.title) ? aiData.table.title : "Detailed Data Breakdown")
    .replace("[TABLE_HEADERS]", tableHeadersHtml)
    .replace("[TABLE_ROWS]", tableRowsHtml)
    .replace("[INSIGHTS_HTML]", insightsHtml)
    .replace("[RECOMMENDATIONS_HTML]", recommendationsHtml)
    .replace("[DATA_SOURCE_NAME]", dataSourceName)
    .replace("[REPORT_TYPE]", (config.reportType || "WEEKLY").toUpperCase())
    .replace(/#4F46E5/g, accentColor);
}

// ================= VITE DEV / PRODUCTION FLOW =================

async function startServer() {
  // Mount Vite middleware for development, or serve compiled files in production
  if (process.env.NODE_ENV !== "production") {
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

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
