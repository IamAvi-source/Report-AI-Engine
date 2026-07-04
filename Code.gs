/**
 * ReportAI — Google Apps Script Backend Code
 * Deploy as Web App. Make sure to enable Drive API and Sheets API under Services if needed,
 * though native Apps Script service bindings are usually automatic.
 */

// Global constant for Drive Folder
var FOLDER_NAME = "ReportAI Reports";

/**
 * Handle GET requests to serve the web app or individual reports
 */
function doGet(e) {
  var reportId = e && e.parameter ? e.parameter.report : null;
  
  if (reportId) {
    try {
      // Load and serve a specific generated report from Google Drive
      var file = DriveApp.getFileById(reportId);
      var content = file.getAs('text/html').getDataAsString();
      
      // Basic password protection check (if password param is provided in URL)
      var passwordParam = e.parameter.p;
      var savedSettings = JSON.parse(PropertiesService.getScriptProperties().getProperty("reportai_settings") || "{}");
      if (savedSettings.reportPassword && savedSettings.reportPassword !== passwordParam) {
        return HtmlService.createHtmlOutput("<h3>🔒 This report is password protected. Please provide the correct passcode in the URL parameter (?p=your_password).</h3>");
      }
      
      return HtmlService.createHtmlOutput(content)
        .setTitle(file.getName())
        .setXFrameOptionsMode(HtmlService.SandboxMode.ALLOWALL);
    } catch (err) {
      return HtmlService.createHtmlOutput("<h3>⚠️ Report not found or access denied. Error: " + err.message + "</h3>");
    }
  }
  
  // Otherwise, serve the main dashboard UI
  var template = HtmlService.createTemplateFromFile('UI');
  template.companyName = getCompanyName();
  
  return template.evaluate()
    .setTitle("ReportAI - AI Report Generator")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.SandboxMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper to include other files (CSS/JS) inside the HTML template
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get company name from settings or fallback
 */
function getCompanyName() {
  try {
    var settings = JSON.parse(PropertiesService.getScriptProperties().getProperty("reportai_settings") || "{}");
    return settings.companyName || "My Company";
  } catch(e) {
    return "My Company";
  }
}

/**
 * Fetch and parse data from a Google Sheet. Caches results to avoid frequent re-fetching.
 */
function getSheetData(sheetUrl, sheetName) {
  if (!sheetUrl) {
    throw new Error("No Sheet URL provided.");
  }
  
  var cache = CacheService.getScriptCache();
  var cacheKey = "sheet_" + MD5(sheetUrl + "_" + (sheetName || "default"));
  var cachedData = cache.get(cacheKey);
  
  if (cachedData) {
    return JSON.parse(cachedData);
  }
  
  try {
    var ss = SpreadsheetApp.openByUrl(sheetUrl);
    var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    
    if (!sheet) {
      throw new Error("Sheet not found: " + (sheetName || "First Sheet"));
    }
    
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      return [];
    }
    
    // Limit to 1000 rows max
    var rangeLimit = Math.min(lastRow, 1001);
    var range = sheet.getRange(1, 1, rangeLimit, lastColumn);
    var values = range.getValues();
    
    var headers = values[0];
    var data = [];
    
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var rowObj = {};
      var hasValues = false;
      
      for (var c = 0; c < headers.length; c++) {
        var header = headers[c] ? headers[c].toString().trim() : ("Column_" + c);
        var cellVal = row[c];
        
        // Format dates as ISO string for proper JS/AI handling
        if (cellVal instanceof Date) {
          cellVal = cellVal.toISOString();
        }
        
        rowObj[header] = cellVal;
        if (cellVal !== "" && cellVal !== null && cellVal !== undefined) {
          hasValues = true;
        }
      }
      
      if (hasValues) {
        data.push(rowObj);
      }
    }
    
    // Cache the data for 5 minutes (300 seconds)
    try {
      cache.put(cacheKey, JSON.stringify(data), 300);
    } catch (cacheErr) {
      // Ignore if payload exceeds cache limit (100KB)
    }
    
    return data;
  } catch (err) {
    throw new Error("Failed to read Google Sheet: " + err.message);
  }
}

/**
 * Generate a complete report using Sheet or CSV data and AI
 */
function generateReport(config) {
  try {
    var sheetData = [];
    var sourceName = "Pasted CSV Data";
    
    if (config.sheetUrl) {
      sheetData = getSheetData(config.sheetUrl, config.sheetName);
      sourceName = config.sheetName || "First Sheet of Google Sheet";
    } else if (config.csvData) {
      sheetData = parseCSV(config.csvData);
      sourceName = "Pasted Table/CSV";
    } else {
      return { success: false, error: "No data source specified. Connect a sheet or paste CSV data." };
    }
    
    if (sheetData.length === 0) {
      return { success: false, error: "The selected data source is empty." };
    }
    
    // Filter data by dates if applicable (assuming some column has date info, or auto-detect)
    var filteredData = filterDataByDates(sheetData, config.reportType, config.startDate, config.endDate);
    
    // Set company branding
    var settings = JSON.parse(PropertiesService.getScriptProperties().getProperty("reportai_settings") || "{}");
    config.companyName = config.companyName || settings.companyName || "My Company";
    
    // Prepare prompt
    var systemPrompt = getSystemPrompt();
    var userMessage = getJsonUserPrompt(filteredData, config);
    
    // Call correct AI API
    var aiResponseText = "";
    var selectedModel = config.aiModel || settings.defaultAiModel || "gemini";
    
    if (selectedModel === "gemini") {
      aiResponseText = callGeminiAPI(userMessage, systemPrompt);
    } else if (selectedModel === "openai") {
      aiResponseText = callOpenAIAPI(userMessage, systemPrompt);
    } else if (selectedModel === "claude") {
      aiResponseText = callClaudeAPI(userMessage, systemPrompt);
    } else {
      throw new Error("Unsupported AI model requested: " + selectedModel);
    }
    
    // Parse response
    var parsedAiData = cleanAndParseJSON(aiResponseText);
    
    // Build HTML output
    var reportHtml = buildReportHTML(parsedAiData, config, sourceName);
    
    // Save report to Google Drive
    var driveInfo = saveReportToDrive(reportHtml, config.reportTitle || parsedAiData.title || "ReportAI Run");
    
    return {
      success: true,
      reportId: driveInfo.id,
      reportHtml: reportHtml,
      driveUrl: driveInfo.url,
      title: config.reportTitle || parsedAiData.title
    };
    
  } catch (err) {
    Logger.log("Report generation error: " + err.toString());
    return { success: false, error: "Failed to generate report: " + err.message };
  }
}

/**
 * Call Gemini API via Google Apps Script UrlFetchApp
 */
function callGeminiAPI(prompt, systemPrompt) {
  var properties = PropertiesService.getScriptProperties();
  var apiKey = properties.getProperty("gemini_key");
  
  if (!apiKey) {
    throw new Error("Google Gemini API Key is not configured in Settings.");
  }
  
  // Model mapping to 2.5 flash or standard
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  
  var payload = {
    contents: [
      {
        parts: [
          { text: prompt }
        ]
      }
    ],
    systemInstruction: {
      parts: [
        { text: systemPrompt }
      ]
    },
    generationConfig: {
      responseMimeType: "application/json",
      temperature: 0.2
    }
  };
  
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error("Gemini API error (" + responseCode + "): " + responseText);
  }
  
  var resObj = JSON.parse(responseText);
  if (resObj.candidates && resObj.candidates[0] && resObj.candidates[0].content && resObj.candidates[0].content.parts && resObj.candidates[0].content.parts[0]) {
    return resObj.candidates[0].content.parts[0].text;
  }
  
  throw new Error("Invalid response format from Gemini API: " + responseText);
}

/**
 * Call OpenAI API GPT-4o
 */
function callOpenAIAPI(prompt, systemPrompt) {
  var properties = PropertiesService.getScriptProperties();
  var apiKey = properties.getProperty("openai_key");
  
  if (!apiKey) {
    throw new Error("OpenAI API Key is not configured in Settings.");
  }
  
  var url = "https://api.openai.com/v1/chat/completions";
  
  var payload = {
    model: "gpt-4o",
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: prompt }
    ],
    response_format: { type: "json_object" },
    temperature: 0.2
  };
  
  var options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + apiKey
    },
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error("OpenAI API error (" + responseCode + "): " + responseText);
  }
  
  var resObj = JSON.parse(responseText);
  if (resObj.choices && resObj.choices[0] && resObj.choices[0].message) {
    return resObj.choices[0].message.content;
  }
  
  throw new Error("Invalid response format from OpenAI API: " + responseText);
}

/**
 * Call Anthropic Claude API
 */
function callClaudeAPI(prompt, systemPrompt) {
  var properties = PropertiesService.getScriptProperties();
  var apiKey = properties.getProperty("claude_key");
  
  if (!apiKey) {
    throw new Error("Anthropic Claude API Key is not configured in Settings.");
  }
  
  var url = "https://api.anthropic.com/v1/messages";
  
  var payload = {
    model: "claude-3-5-sonnet-20241022",
    max_tokens: 4000,
    system: systemPrompt,
    messages: [
      { role: "user", content: prompt }
    ],
    temperature: 0.2
  };
  
  var options = {
    method: "post",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error("Claude API error (" + responseCode + "): " + responseText);
  }
  
  var resObj = JSON.parse(responseText);
  if (resObj.content && resObj.content[0] && resObj.content[0].type === "text") {
    return resObj.content[0].text;
  }
  
  throw new Error("Invalid response format from Claude API: " + responseText);
}

/**
 * Clean up text blocks and ensure proper parsing of JSON strings returned by the AIs
 */
function cleanAndParseJSON(responseText) {
  var cleanedText = responseText.trim();
  
  // Strip markdown code block wrappers if any
  if (cleanedText.indexOf("```json") === 0) {
    cleanedText = cleanedText.substring(7);
  } else if (cleanedText.indexOf("```") === 0) {
    cleanedText = cleanedText.substring(3);
  }
  
  if (cleanedText.lastIndexOf("```") === cleanedText.length - 3) {
    cleanedText = cleanedText.substring(0, cleanedText.length - 3);
  }
  
  cleanedText = cleanedText.trim();
  
  try {
    return JSON.parse(cleanedText);
  } catch (err) {
    Logger.log("JSON Parse Error on raw content: " + responseText);
    throw new Error("AI returned invalid JSON format: " + err.message);
  }
}

/**
 * Parse standard CSV string into a JSON array of objects
 */
function parseCSV(csvText) {
  var lines = [];
  var row = [""];
  var inQuotes = false;

  for (var i = 0; i < csvText.length; i++) {
    var c = csvText[i];
    var next = csvText[i+1];
    
    if (c === '"') {
      if (inQuotes && next === '"') {
        row[row.length - 1] += '"'; // Double quotes inside quote
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === ',' && !inQuotes) {
      row.push('');
    } else if ((c === '\r' || c === '\n') && !inQuotes) {
      if (c === '\r' && next === '\n') {
        i++;
      }
      lines.push(row);
      row = [''];
    } else {
      row[row.length - 1] += c;
    }
  }
  
  if (row.length > 1 || row[0] !== '') {
    lines.push(row);
  }
  
  if (lines.length <= 1) return [];
  
  var headers = lines[0];
  var data = [];
  
  for (var r = 1; r < lines.length; r++) {
    var vals = lines[r];
    var rowObj = {};
    var hasVals = false;
    
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col] ? headers[col].toString().trim() : "Column_" + col;
      var val = vals[col] ? vals[col].toString().trim() : "";
      rowObj[header] = val;
      if (val !== "") hasVals = true;
    }
    
    if (hasVals) {
      data.push(rowObj);
    }
  }
  
  return data;
}

/**
 * Filter data by dates based on chosen report type or custom picker
 */
function filterDataByDates(data, reportType, startDateStr, endDateStr) {
  // If custom date picker is not active and no specific limits, return full (or first 500 rows for sizing)
  if (!reportType || reportType === "custom" && (!startDateStr || !endDateStr)) {
    return data;
  }
  
  var now = new Date();
  var startLimit = new Date();
  var endLimit = new Date();
  
  if (reportType === "daily") {
    // Past 24 hours
    startLimit.setDate(now.getDate() - 1);
  } else if (reportType === "weekly") {
    // Past 7 days
    startLimit.setDate(now.getDate() - 7);
  } else if (reportType === "monthly") {
    // Past 30 days
    startLimit.setDate(now.getDate() - 30);
  } else if (reportType === "custom") {
    startLimit = new Date(startDateStr);
    endLimit = new Date(endDateStr);
    endLimit.setHours(23, 59, 59, 999);
  }
  
  // Try to find a date column in rows
  var firstRow = data[0];
  var dateKey = null;
  var keys = Object.keys(firstRow);
  
  for (var k = 0; k < keys.length; k++) {
    var keyLower = keys[k].toLowerCase();
    if (keyLower.indexOf("date") !== -1 || keyLower.indexOf("time") !== -1 || keyLower.indexOf("timestamp") !== -1) {
      dateKey = keys[k];
      break;
    }
  }
  
  if (!dateKey) {
    // Fallback: return data as is, if date column not detected
    return data;
  }
  
  return data.filter(function(row) {
    var rowDateVal = row[dateKey];
    if (!rowDateVal) return false;
    
    var rowDate = new Date(rowDateVal);
    if (isNaN(rowDate.getTime())) return true; // Keep if invalid date format, just to be safe
    
    if (reportType === "custom") {
      return rowDate >= startLimit && rowDate <= endLimit;
    } else {
      return rowDate >= startLimit;
    }
  });
}

/**
 * Save report HTML string to Google Drive
 */
function saveReportToDrive(htmlContent, fileName) {
  var folder;
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(FOLDER_NAME);
  }
  
  var finalName = fileName + " (" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") + ").html";
  var file = folder.createFile(finalName, htmlContent, "text/html");
  
  // Enable link sharing
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return {
    id: file.getId(),
    url: file.getUrl()
  };
}

/**
 * Retrieve saved report history
 */
function getReportHistory() {
  var list = [];
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  
  if (!folders.hasNext()) {
    return [];
  }
  
  var folder = folders.next();
  var files = folder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === "text/html") {
      list.push({
        id: file.getId(),
        name: file.getName(),
        createdDate: file.getDateCreated().toISOString(),
        driveUrl: file.getUrl()
      });
    }
  }
  
  // Sort by date descending
  list.sort(function(a, b) {
    return new Date(b.createdDate) - new Date(a.createdDate);
  });
  
  // Return last 20 reports
  return list.slice(0, 20);
}

/**
 * Construct HTML file by replacing tokens inside template
 */
function buildReportHTML(aiData, config, dataSourceName) {
  var templateHtml = HtmlService.createHtmlOutputFromFile("Report").getContent();
  
  var settings = JSON.parse(PropertiesService.getScriptProperties().getProperty("reportai_settings") || "{}");
  var company = config.companyName || settings.companyName || "My Company";
  var accentColor = settings.accentColor || "#4F46E5";
  
  // Process KPI cards HTML
  var kpiCardsHtml = "";
  if (aiData.kpis && aiData.kpis.length > 0) {
    aiData.kpis.forEach(function(kpi) {
      var dirClass = kpi.direction === "up" ? "up" : (kpi.direction === "down" ? "down" : "neutral");
      var isPosClass = kpi.is_positive ? "positive" : "negative";
      var highlightClass = kpi.highlight ? "highlight" : "";
      
      var changeHtml = "";
      if (kpi.change !== undefined && kpi.direction !== "neutral") {
        var sign = kpi.direction === "up" ? "▲" : "▼";
        changeHtml = '<div class="kpi-change ' + dirClass + ' ' + isPosClass + '">' +
                       '<span>' + sign + '</span> ' + Math.abs(kpi.change) + '%' +
                     '</div>';
      } else {
        changeHtml = '<div class="kpi-change neutral">Neutral</div>';
      }
      
      kpiCardsHtml += '<div class="kpi-card ' + highlightClass + '">' +
                        '<div class="kpi-label">' + kpi.label + '</div>' +
                        '<div class="kpi-value">' + kpi.value + '</div>' +
                        '<div class="kpi-change-row">' +
                          changeHtml +
                          '<span class="kpi-period">' + (kpi.change_period || "") + '</span>' +
                        '</div>' +
                      '</div>';
    });
  }
  
  // Process Alerts HTML
  var alertsHtml = "";
  if (aiData.alerts && aiData.alerts.length > 0) {
    aiData.alerts.forEach(function(alert) {
      var sevClass = alert.severity === "critical" ? "critical" : (alert.severity === "warning" ? "warning" : "info");
      var icon = sevClass === "critical" ? "🚨" : (sevClass === "warning" ? "⚠️" : "ℹ️");
      
      alertsHtml += '<div class="alert ' + sevClass + '">' +
                      '<span class="alert-icon">' + icon + '</span>' +
                      '<div class="alert-content">' +
                        '<div class="alert-title">' + alert.title + '</div>' +
                        '<div>' + alert.message + '</div>' +
                      '</div>' +
                    '</div>';
    });
  }
  
  // Process Table Headers and Rows
  var tableHeadersHtml = "";
  var tableRowsHtml = "";
  if (aiData.table) {
    if (aiData.table.headers) {
      aiData.table.headers.forEach(function(header) {
        tableHeadersHtml += '<th>' + header + '</th>';
      });
    }
    if (aiData.table.rows) {
      aiData.table.rows.forEach(function(row) {
        tableRowsHtml += '<tr>';
        row.forEach(function(cell) {
          tableRowsHtml += '<td>' + cell + '</td>';
        });
        tableRowsHtml += '</tr>';
      });
    }
  }
  
  // Process Insights List
  var insightsHtml = "";
  if (aiData.insights) {
    aiData.insights.forEach(function(insight, idx) {
      insightsHtml += '<div class="insight-item">' +
                        '<div class="insight-num">' + (idx + 1) + '</div>' +
                        '<div>' + insight + '</div>' +
                      '</div>';
    });
  }
  
  // Process Recommendations List
  var recommendationsHtml = "";
  if (aiData.recommendations) {
    aiData.recommendations.forEach(function(rec) {
      recommendationsHtml += '<div class="rec-item">' +
                               '<span class="rec-arrow">➔</span>' +
                               '<div>' + rec + '</div>' +
                             '</div>';
    });
  }
  
  // Get charts configuration
  var chart1Title = "Chart 1";
  var chart1Type = "line";
  var chart1Labels = "";
  var chart1Datasets = "";
  
  var chart2Title = "Chart 2";
  var chart2Type = "bar";
  var chart2Labels = "";
  var chart2Datasets = "";
  
  if (aiData.charts && aiData.charts.length > 0) {
    var c1 = aiData.charts[0];
    chart1Title = c1.title || "First Metric Trend";
    chart1Type = c1.type || "line";
    chart1Labels = (c1.labels || []).map(function(l) { return '"' + l + '"'; }).join(",");
    
    var datasets1 = [];
    if (c1.datasets) {
      c1.datasets.forEach(function(ds) {
        var col = ds.color || accentColor;
        datasets1.push(JSON.stringify({
          label: ds.label,
          data: ds.data,
          backgroundColor: hexToRgbA(col, 0.1),
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
    var c2 = aiData.charts[1];
    chart2Title = c2.title || "Second Metric Trend";
    chart2Type = c2.type || "bar";
    chart2Labels = (c2.labels || []).map(function(l) { return '"' + l + '"'; }).join(",");
    
    var datasets2 = [];
    if (c2.datasets) {
      c2.datasets.forEach(function(ds) {
        var col = ds.color || "#10B981";
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
  
  // Replace all placeholders
  var finalHtml = templateHtml
    .replace("[COMPANY_NAME]", company)
    .replace("[TITLE]", aiData.title || config.reportTitle || "AI Business Report")
    .replace("[PERIOD]", aiData.period || config.reportType || "Analysis Report")
    .replace("[GENERATED_DATE]", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy"))
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
    .replace("[REPORT_TYPE]", config.reportType ? config.reportType.toUpperCase() : "CUSTOM")
    
    // Apply Branding Colors if customized
    .replace(/#4F46E5/g, accentColor);
    
  return finalHtml;
}

/**
 * Handle Report Scheduling
 */
function scheduleReport(scheduleConfig) {
  try {
    var properties = PropertiesService.getScriptProperties();
    var scheduleId = "schedule_" + Utilities.getUuid();
    
    scheduleConfig.id = scheduleId;
    scheduleConfig.status = "active";
    scheduleConfig.createdDate = new Date().toISOString();
    
    properties.setProperty(scheduleId, JSON.stringify(scheduleConfig));
    
    // Set up daily background trigger if it doesn't exist
    var triggers = ScriptApp.getProjectTriggers();
    var hasTrigger = false;
    
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "runScheduledReports") {
        hasTrigger = true;
        break;
      }
    }
    
    if (!hasTrigger) {
      ScriptApp.newTrigger("runScheduledReports")
        .timeBased()
        .everyDays(1)
        .atHour(7) // Auto-run at 7:00 AM daily
        .create();
    }
    
    return { success: true, scheduleId: scheduleId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Retrieve active schedules
 */
function getSchedules() {
  var list = [];
  var properties = PropertiesService.getScriptProperties().getProperties();
  
  for (var key in properties) {
    if (key.indexOf("schedule_") === 0) {
      list.push(JSON.parse(properties[key]));
    }
  }
  
  // Sort by created date
  list.sort(function(a, b) {
    return new Date(b.createdDate) - new Date(a.createdDate);
  });
  
  return list;
}

/**
 * Delete a saved schedule
 */
function deleteSchedule(scheduleId) {
  try {
    PropertiesService.getScriptProperties().deleteProperty(scheduleId);
    
    // Clean up trigger if no schedules are left
    var schedules = getSchedules();
    if (schedules.length === 0) {
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "runScheduledReports") {
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
    }
    
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * This function is called by the time trigger.
 * It checks schedules and generates/emails reports.
 */
function runScheduledReports() {
  var schedules = getSchedules();
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0 is Sunday, 1 is Monday, etc.
  var dayOfMonth = today.getDate();
  
  schedules.forEach(function(schedule) {
    if (schedule.status !== "active") return;
    
    var shouldRun = false;
    
    if (schedule.frequency === "daily") {
      shouldRun = true;
    } else if (schedule.frequency === "weekly") {
      // Check if current day of week is checked in config
      // Mon=1, Tue=2... Sun=0. Match weeklyDays: {mon: true, ...}
      var weekdayKeys = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
      var currentDayKey = weekdayKeys[dayOfWeek];
      if (schedule.weeklyDays && schedule.weeklyDays[currentDayKey]) {
        shouldRun = true;
      }
    } else if (schedule.frequency === "monthly") {
      // Check if current day of month matches config
      if (Number(schedule.monthlyDay) === dayOfMonth) {
        shouldRun = true;
      }
    }
    
    if (shouldRun) {
      try {
        var reportConfig = {
          sheetUrl: schedule.sheetUrl,
          sheetName: schedule.sheetName,
          reportType: schedule.reportType,
          userPrompt: schedule.userPrompt,
          aiModel: schedule.aiModel,
          reportTitle: schedule.reportTitle || "Automated Schedule Run"
        };
        
        var result = generateReport(reportConfig);
        
        if (result.success) {
          emailReport(result.reportHtml, schedule.recipients, result.title || "Scheduled AI Business Report");
        }
      } catch (err) {
        Logger.log("Error running schedule " + schedule.id + ": " + err.toString());
      }
    }
  });
}

/**
 * Send Generated Report via Email
 */
function emailReport(htmlContent, recipients, subject) {
  if (!recipients || recipients.length === 0) return;
  
  var recipientString = Array.isArray(recipients) ? recipients.join(",") : recipients;
  
  // Format body text
  var bodyText = "Please find attached the latest automated business report generated by ReportAI.\n\nNote: If you cannot view HTML emails, you can view the attached report directly.";
  
  // Create HTML attachment blob
  var attachment = Utilities.newBlob(htmlContent, "text/html", subject.replace(/[^a-z0-9]/gi, '_').toLowerCase() + ".html");
  
  MailApp.sendEmail({
    to: recipientString,
    subject: "📊 ReportAI: " + subject,
    body: bodyText,
    htmlBody: htmlContent, // Inlines report nicely inside Gmail clients
    attachments: [attachment]
  });
}

/**
 * Settings Get/Set Storage
 */
function saveSettings(settings) {
  try {
    var properties = PropertiesService.getScriptProperties();
    
    // Ensure we don't overwrite existing secrets with masked values
    var existingSettings = JSON.parse(properties.getProperty("reportai_settings") || "{}");
    
    if (settings.claudeKey === "••••••••••••") settings.claudeKey = properties.getProperty("claude_key") || "";
    if (settings.openaiKey === "••••••••••••") settings.openaiKey = properties.getProperty("openai_key") || "";
    if (settings.geminiKey === "••••••••••••") settings.geminiKey = properties.getProperty("gemini_key") || "";
    
    // Save physical keys in distinct variables for security and to prevent reading them back
    if (settings.claudeKey) properties.setProperty("claude_key", settings.claudeKey);
    if (settings.openaiKey) properties.setProperty("openai_key", settings.openaiKey);
    if (settings.geminiKey) properties.setProperty("gemini_key", settings.geminiKey);
    
    // Save general public settings block
    var generalSettings = {
      companyName: settings.companyName,
      logoUrl: settings.logoUrl,
      accentColor: settings.accentColor,
      footerText: settings.footerText,
      timezone: settings.timezone,
      defaultAiModel: settings.defaultAiModel,
      defaultReportType: settings.defaultReportType,
      reportPassword: settings.reportPassword,
      hasClaude: !!settings.claudeKey,
      hasOpenai: !!settings.openaiKey,
      hasGemini: !!settings.geminiKey
    };
    
    properties.setProperty("reportai_settings", JSON.stringify(generalSettings));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getSettings() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var raw = properties.getProperty("reportai_settings");
    var settings = raw ? JSON.parse(raw) : {
      companyName: "My Company",
      accentColor: "#4F46E5",
      timezone: "GMT",
      defaultAiModel: "gemini",
      defaultReportType: "weekly"
    };
    
    // Mask real API keys
    settings.claudeKey = properties.getProperty("claude_key") ? "••••••••••••" : "";
    settings.openaiKey = properties.getProperty("openai_key") ? "••••••••••••" : "";
    settings.geminiKey = properties.getProperty("gemini_key") ? "••••••••••••" : "";
    
    return settings;
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * System and User prompts for report generation
 */
function getSystemPrompt() {
  return "You are an expert business data analyst. You receive business data in JSON format " +
         "and a user's reporting instruction. Your ONLY output must be valid JSON — no markdown, " +
         "no explanation, no code fences, just raw JSON.\n\n" +
         "Return a JSON object with EXACTLY this structure:\n" +
         "{\n" +
         "  \"title\": \"Report title string\",\n" +
         "  \"period\": \"e.g. Week of June 24–30, 2025\",\n" +
         "  \"executive_summary\": \"2-3 sentence plain English summary of the most important findings\",\n" +
         "  \"kpis\": [\n" +
         "    {\n" +
         "      \"label\": \"Metric name\",\n" +
         "      \"value\": \"Formatted value e.g. $12,400 or 847 units\",\n" +
         "      \"change\": 12.5,\n" +
         "      \"change_period\": \"vs last week\",\n" +
         "      \"direction\": \"up\",\n" +
         "      \"is_positive\": true,\n" +
         "      \"highlight\": true\n" +
         "    }\n" +
         "  ],\n" +
         "  \"charts\": [\n" +
         "    {\n" +
         "      \"id\": \"chart1\",\n" +
         "      \"type\": \"line\",\n" +
         "      \"title\": \"Chart title\",\n" +
         "      \"labels\": [\"Mon\", \"Tue\", \"Wed\"],\n" +
         "      \"datasets\": [\n" +
         "        {\n" +
         "          \"label\": \"Revenue\",\n" +
         "          \"data\": [1200, 1400, 1100],\n" +
         "          \"color\": \"#4F46E5\"\n" +
         "        }\n" +
         "      ]\n" +
         "    },\n" +
         "    {\n" +
         "      \"id\": \"chart2\",\n" +
         "      \"type\": \"bar\",\n" +
         "      \"title\": \"Second chart title\",\n" +
         "      \"labels\": [\"Product A\", \"Product B\"],\n" +
         "      \"datasets\": [\n" +
         "        {\n" +
         "          \"label\": \"Units Sold\",\n" +
         "          \"data\": [340, 280],\n" +
         "          \"color\": \"#10B981\"\n" +
         "        }\n" +
         "      ]\n" +
         "    }\n" +
         "  ],\n" +
         "  \"table\": {\n" +
         "    \"title\": \"Detailed breakdown\",\n" +
         "    \"headers\": [\"Column1\", \"Column2\", \"Column3\"],\n" +
         "    \"rows\": [\n" +
         "      [\"Row1Val1\", \"Row1Val2\", \"Row1Val3\"]\n" +
         "    ],\n" +
         "    \"sort_column\": 0,\n" +
         "    \"sort_direction\": \"desc\"\n" +
         "  },\n" +
         "  \"insights\": [\n" +
         "    \"Key finding 1 — specific, data-driven\",\n" +
         "    \"Key finding 2 — specific, data-driven\",\n" +
         "    \"Key finding 3 — specific, data-driven\"\n" +
         "  ],\n" +
         "  \"recommendations\": [\n" +
         "    \"Actionable recommendation 1\",\n" +
         "    \"Actionable recommendation 2\",\n" +
         "    \"Actionable recommendation 3\"\n" +
         "  ],\n" +
         "  \"alerts\": [\n" +
         "    {\n" +
         "      \"severity\": \"warning\",\n" +
         "      \"title\": \"Alert title\",\n" +
         "      \"message\": \"Alert detail\"\n" +
         "    }\n" +
         "  ]\n" +
         "}\n\n" +
         "Rules:\n" +
         "- kpis: always include 4-6 KPIs\n" +
         "- charts: always include exactly 2 charts\n" +
         "- insights: always exactly 3\n" +
         "- recommendations: always exactly 3\n" +
         "- alerts: include only if there are genuine anomalies (can be empty array [])\n" +
         "- All numbers must be real values derived from the provided data\n" +
         "- direction must be 'up', 'down', or 'neutral'\n" +
         "- is_positive: true means the direction is good for business\n";
}

function getJsonUserPrompt(sheetData, config) {
  return "DATA: " + JSON.stringify(sheetData) + "\n\n" +
         "REPORT INSTRUCTION: " + config.userPrompt + "\n\n" +
         "REPORT TYPE: " + config.reportType + "\n" +
         "DATE RANGE: " + config.startDate + " to " + config.endDate + "\n" +
         "COMPANY: " + config.companyName + "\n\n" +
         "Generate the report JSON now.";
}

/**
 * MD5 implementation for cache key generation
 */
function MD5(str) {
  var signature = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str, Utilities.Charset.UTF_8);
  var out = "";
  for (var i = 0; i < signature.length; i++) {
    var val = signature[i];
    if (val < 0) val += 256;
    var byteString = val.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    out += byteString;
  }
  return out;
}

/**
 * Hex to RGBA color helper
 */
function hexToRgbA(hex, alpha) {
  var c;
  if(/^#([A-Fa-f0-9]{3}){1,2}$/.test(hex)){
    c= hex.substring(1).split('');
    if(c.length== 3){
      c= [c[0], c[0], c[1], c[1], c[2], c[2]];
    }
    c= '0x' + c.join('');
    return 'rgba(' + [(c>>16)&255, (c>>8)&255, c&255].join(',') + ',' + alpha + ')';
  }
  return hex;
}
