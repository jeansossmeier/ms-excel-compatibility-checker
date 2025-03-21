export interface CheckResult {
  status: 'success' | 'warning' | 'error';
  message: string;
  details?: string;
  location?: string;
}

export interface CompatibilityReport {
  fileName: string;
  fileSize: number;
  lastModified: Date;
  results: CheckResult[];
  totalIssues: number;
  isCompatible: boolean;
} 