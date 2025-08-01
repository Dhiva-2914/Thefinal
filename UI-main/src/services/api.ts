const API_BASE_URL = '';

export interface SearchRequest {
  space_key: string;
  page_titles: string[];
  query: string;
}

export interface CodeRequest {
  space_key: string;
  page_title: string;
  instruction: string;
  target_language?: string;
}

export interface ImpactRequest {
  space_key: string;
  old_page_title: string;
  new_page_title: string;
  question?: string;
  enable_stack_overflow_check?: boolean;
}

export interface DirectCodeImpactRequest {
  old_code: string;
  new_code: string;
  question?: string;
  enable_stack_overflow_check?: boolean;
}

export interface PushToJiraConfluenceSlackRequest {
  summary: string;
  video_title: string;
}

export interface PushToJiraConfluenceSlackResponse {
  success: boolean;
  tasks_created: number;
  jira_issues_created: number;
  confluence_updated: boolean;
  slack_notifications_sent: number;
}

export interface TestRequest {
  space_key: string;
  code_page_title: string;
  test_input_page_title?: string;
  question?: string;
}

export interface GitHubActionsRequest {
  space_key: string;
  code_page_title: string;
  test_input_page_title?: string;
  repository_name: string;
  branch_name?: string;
  github_token?: string;
  enable_parallel_testing?: boolean;
  test_frameworks?: string[];
  platforms?: string[];
  auto_push?: boolean;
}

export interface GitHubActionsResponse {
  workflow_content: string;
  test_files: Array<{filename: string, content: string}>;
  setup_instructions: string;
  integration_status: string;
  estimated_duration: string;
  coverage_estimate: string;
  language_info?: {
    language: string;
    framework: string;
    test_framework: string;
    package_manager: string;
    build_tool: string;
  };
  auto_push_result?: {
    success: boolean;
    files_pushed: string[];
    errors: string[];
    repository_url: string;
    user_info?: {
      username: string;
      repository: string;
    };
  };
}

export interface ExportRequest {
  content: string;
  format: string;
  filename: string;
}

export interface Space {
  name: string;
  key: string;
}

export interface SearchResponse {
  response: string;
  pages_analyzed: number;
  page_titles: string[];
  source?: string;
}

export interface CodeResponse {
  summary: string;
  original_code: string;
  detected_language: string;
  modified_code?: string;
  converted_code?: string;
  target_language?: string;
}

export interface ImpactResponse {
  lines_added: number;
  lines_removed: number;
  files_changed: number;
  percentage_change: number;
  impact_analysis: string;
  recommendations: string;
  risk_analysis: string;
  risk_level: string;
  risk_score: number;
  risk_factors: string[];
  answer?: string;
  diff: string;
  stack_overflow_risks?: StackOverflowRisk[];
}

export interface StackOverflowRisk {
  pattern: string;
  risk_level: 'low' | 'medium' | 'high';
  description: string;
  stack_overflow_links: string[];
  alternative_suggestions: string[];
  deprecation_warning?: string;
}

export interface TestResponse {
  test_strategy: string;
  cross_platform_testing: string;
  sensitivity_analysis?: string;
  ai_response?: string;
}

export interface ExportResponse {
  file: string;
  mime: string;
  filename: string;
}

export interface VideoRequest {
  video_url?: string;
  space_key: string;
  page_title: string;
  question?: string;
}

export interface VideoResponse {
  summary: string;
  quotes: string[];
  timestamps?: string[];
  qa: Array<{question: string, answer: string}>;
  page_title: string;
  answer?: string;
}

export interface ImageRequest {
  space_key: string;
  page_title: string;
  image_url: string;
}

export interface ImageSummaryRequest {
  space_key: string;
  page_title: string;
  image_url: string;
  summary: string;
  question: string;
}

export interface ChartRequest {
  space_key: string;
  page_title: string;
  image_url?: string;
  table_html?: string;
  excel_url?: string;
  chart_type: string;
  filename: string;
  format: string;
}

export interface ImageResponse {
  summary: string;
}

export interface ImageQAResponse {
  answer: string;
}

export interface ChartResponse {
  chart_data: string;
  mime_type: string;
  filename: string;
}

export interface SaveToConfluenceRequest {
  space_key: string;
  page_title: string;
  content: string;
  mode?: string;
}

export interface PreviewSaveToConfluenceRequest {
  space_key: string;
  page_title: string;
  content: string;
  mode: string;
}

export interface PreviewSaveToConfluenceResponse {
  preview_content: string;
  diff: string;
}

export interface SaveToConfluenceResponse {
  success: boolean;
  message?: string;
}

export interface InsightSourcesResponse {
  images: string[];
  tables: string[]; // HTML strings
  excels: string[]; // URLs
}

export interface TableSummaryRequest {
  space_key: string;
  page_title: string;
  table_html: string;
}
export interface ExcelSummaryRequest {
  space_key: string;
  page_title: string;
  excel_url: string;
}
export interface SummaryResponse {
  summary: string;
}

export interface PageWithType {
  id: string;
  title: string;
  content_type: string;
}

class ApiService {
  private getSelectedApiKey(): string | undefined {
    if (typeof window !== 'undefined' && localStorage.getItem('selectedApiKeyId')) {
      return localStorage.getItem('selectedApiKeyId') || undefined;
    }
    return undefined;
  }

  private async makeRequest<T>(endpoint: string, options?: RequestInit): Promise<T> {
    const apiKey = this.getSelectedApiKey();
    const baseHeaders: Record<string, string> = {
      'Content-Type': 'application/json',
    };
    const optHeaders = (options && options.headers && typeof options.headers === 'object' && !Array.isArray(options.headers)) ? options.headers : {};
    const headers: Record<string, string> = Object.assign({}, baseHeaders, optHeaders);
    if (apiKey) {
      headers['x-api-key'] = apiKey;
    }
    const response = await fetch(`${API_BASE_URL}${endpoint}`, {
      headers,
      ...options,
    });
    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.detail || 'API request failed');
    }
    return response.json();
  }

  async getSpaces(): Promise<{ spaces: Space[] }> {
    return this.makeRequest<{ spaces: Space[] }>('/spaces');
  }

  async getPages(spaceKey: string): Promise<{ pages: string[] }> {
    return this.makeRequest<{ pages: string[] }>(`/pages/${spaceKey}`);
  }

  async search(request: SearchRequest): Promise<SearchResponse> {
    return this.makeRequest<SearchResponse>('/search', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async videoSummarizer(request: VideoRequest): Promise<VideoResponse> {
    return this.makeRequest<VideoResponse>('/video-summarizer', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async codeAssistant(request: CodeRequest): Promise<CodeResponse> {
    return this.makeRequest<CodeResponse>('/code-assistant', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async impactAnalyzer(request: ImpactRequest): Promise<ImpactResponse> {
    return this.makeRequest<ImpactResponse>('/impact-analyzer', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async directCodeImpactAnalyzer(request: DirectCodeImpactRequest): Promise<ImpactResponse> {
    return this.makeRequest<ImpactResponse>('/direct-code-impact-analyzer', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async pushToJiraConfluenceSlack(request: PushToJiraConfluenceSlackRequest): Promise<PushToJiraConfluenceSlackResponse> {
    return this.makeRequest<PushToJiraConfluenceSlackResponse>('/push-to-jira-confluence-slack', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async testSupport(request: TestRequest): Promise<TestResponse> {
    return this.makeRequest<TestResponse>('/test-support', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async githubActionsIntegration(request: GitHubActionsRequest): Promise<GitHubActionsResponse> {
    return this.makeRequest<GitHubActionsResponse>('/github-actions-integration', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async getImages(spaceKey: string, pageTitle: string): Promise<InsightSourcesResponse> {
    return this.makeRequest<InsightSourcesResponse>(`/images/${spaceKey}/${encodeURIComponent(pageTitle)}`);
  }

  async imageSummary(request: ImageRequest): Promise<ImageResponse> {
    return this.makeRequest<ImageResponse>('/image-summary', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async imageQA(request: ImageSummaryRequest): Promise<ImageQAResponse> {
    return this.makeRequest<ImageQAResponse>('/image-qa', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async createChart(request: ChartRequest): Promise<ChartResponse> {
    return this.makeRequest<ChartResponse>('/create-chart', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async exportContent(request: ExportRequest): Promise<Blob> {
    const response = await fetch(`${API_BASE_URL}/export`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(request),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.detail || 'Export failed');
    }

    const result = await response.json();
    
    // Handle binary files (PDF, DOCX, PPTX) that are base64 encoded
    if (request.format === 'pdf' || request.format === 'docx' || request.format === 'pptx') {
      const binaryString = atob(result.file);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      return new Blob([bytes], { type: result.mime });
    } else {
      // Handle text files
      const encoder = new TextEncoder();
      const bytes = encoder.encode(result.file);
      return new Blob([bytes], { type: result.mime });
    }
  }

  async saveToConfluence(request: SaveToConfluenceRequest): Promise<SaveToConfluenceResponse> {
    return this.makeRequest<SaveToConfluenceResponse>('/save-to-confluence', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async previewSaveToConfluence(request: PreviewSaveToConfluenceRequest): Promise<PreviewSaveToConfluenceResponse> {
    return this.makeRequest<PreviewSaveToConfluenceResponse>('/preview-save-to-confluence', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async tableSummary(request: TableSummaryRequest): Promise<SummaryResponse> {
    return this.makeRequest<SummaryResponse>('/table-summary', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }
  async excelSummary(request: ExcelSummaryRequest): Promise<SummaryResponse> {
    return this.makeRequest<SummaryResponse>('/excel-summary', {
      method: 'POST',
      body: JSON.stringify(request),
    });
  }

  async sendToGoogleChat(summary: string): Promise<{ status: string; message: string }> {
    return this.makeRequest<{ status: string; message: string }>('/send-to-google-chat', {
      method: 'POST',
      body: JSON.stringify({ summary }),
    });
  }
}

export const apiService = new ApiService(); 

export async function analyzeGoal(goal: string, availablePages: string[]) {
  const res = await fetch(`${API_BASE_URL}/analyze-goal`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': localStorage.getItem('selectedApiKeyId') || '',
    },
    body: JSON.stringify({ goal, available_pages: availablePages }),
  });
  if (!res.ok) throw new Error('Failed to analyze goal');
  return res.json();
} 

export async function getPagesWithType(spaceKey: string): Promise<{ pages: PageWithType[] }> {
  const res = await fetch(`${API_BASE_URL}/pages-with-type/${spaceKey}`, {
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': localStorage.getItem('selectedApiKeyId') || '',
    },
  });
  if (!res.ok) throw new Error('Failed to fetch pages with type');
  return res.json();
}
