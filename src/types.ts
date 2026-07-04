export interface GeneratedReport {
  id: string;
  name: string;
  createdDate: string;
  driveUrl: string;
  reportType: string;
  reportHtml?: string;
}

export interface ScheduleConfig {
  id?: string;
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
  status?: string;
  createdDate?: string;
}

export interface AppSettings {
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
