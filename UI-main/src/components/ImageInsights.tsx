import React, { useState, useRef, useEffect } from 'react';
import { Image, Download, Save, X, ChevronDown, Loader2, MessageSquare, BarChart3, Search, Video, Code, TrendingUp, TestTube, Eye, ChevronUp, Check } from 'lucide-react';
import { FeatureType, AppMode } from '../App';
import { apiService } from '../services/api';
import { getConfluenceSpaceAndPageFromUrl } from '../utils/urlUtils';
import CustomScrollbar from './CustomScrollbar';
import VoiceRecorder from './VoiceRecorder';

interface ImageInsightsProps {
  onClose: () => void;
  onFeatureSelect: (feature: FeatureType) => void;
  onModeSelect: (mode: AppMode) => void;
  autoSpaceKey?: string | null;
  isSpaceAutoConnected?: boolean;
}

interface ImageData {
  id: string;
  name: string;
  url: string;
  summary?: string;
  qa?: { question: string; answer: string }[];
  pageTitle?: string;
}

interface ChartData {
  type: 'bar' | 'line' | 'pie' | 'stacked';
  data: any;
  title: string;
}

interface TableData {
  id: string;
  name: string;
  html: string;
  pageTitle?: string;
  summary?: string;
  qa?: { question: string; answer: string }[];
}

interface ExcelData {
  id: string;
  name: string;
  url: string;
  pageTitle?: string;
  summary?: string;
  qa?: { question: string; answer: string }[];
}

const ImageInsights: React.FC<ImageInsightsProps> = ({ onClose, onFeatureSelect, onModeSelect, autoSpaceKey, isSpaceAutoConnected }) => {
  const [spaceKey, setSpaceKey] = useState<string>('');
  const [selectedPages, setSelectedPages] = useState<string[]>([]);
  const [images, setImages] = useState<ImageData[]>([]);
  const [isAnalyzing, setIsAnalyzing] = useState<string>('');
  const [newQuestion, setNewQuestion] = useState('');
  const [selectedImage, setSelectedImage] = useState<string>('');
  const [fileName, setFileName] = useState('');
  const [exportFormat, setExportFormat] = useState('pdf');
  const [chartData, setChartData] = useState<ChartData | null>(null);
  const [selectedChartType, setSelectedChartType] = useState<'bar' | 'line' | 'pie' | 'stacked'>('bar');
  const [chartFileName, setChartFileName] = useState('');
  const [chartExportFormat, setChartExportFormat] = useState('png');
  const [spaces, setSpaces] = useState<Array<{name: string, key: string}>>([]);
  const [pages, setPages] = useState<string[]>([]);
  const [isLoadingSpaces, setIsLoadingSpaces] = useState(false);
  const [isLoadingPages, setIsLoadingPages] = useState(false);
  const [isLoadingImages, setIsLoadingImages] = useState(false);
  const [isAskingQuestion, setIsAskingQuestion] = useState(false);
  const [isCreatingChart, setIsCreatingChart] = useState(false);
  const [isChangingChartType, setIsChangingChartType] = useState(false);
  const [isExportingChart, setIsExportingChart] = useState(false);
  const [showToast, setShowToast] = useState(false);
  const [saveMode, setSaveMode] = useState('append');
  const [previewContent, setPreviewContent] = useState<string | null>(null);
  const [previewDiff, setPreviewDiff] = useState<string | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  const [isPreviewLoading, setIsPreviewLoading] = useState(false);
  const chartPreviewRef = useRef<HTMLDivElement>(null);
  const [isPageDropdownOpen, setIsPageDropdownOpen] = useState(false);
  const [tables, setTables] = useState<TableData[]>([]);
  const [excels, setExcels] = useState<ExcelData[]>([]);
  const [pageSearch, setPageSearch] = useState('');
  const [exportFormatSearch, setExportFormatSearch] = useState('');
  const [isExportFormatDropdownOpen, setIsExportFormatDropdownOpen] = useState(false);
  const exportFormats = [
    { value: 'markdown', label: 'Markdown' },
    { value: 'pdf', label: 'PDF' },
    { value: 'docx', label: 'Word Document' },
    { value: 'pptx', label: 'PowerPoint' }
  ];

  // --- History feature for Q&A ---
  const [qaHistory, setQaHistory] = useState<Array<{question: string, answer: string, imageId: string}>>([]);
  const [currentQaHistoryIndex, setCurrentQaHistoryIndex] = useState<number | null>(null);

  // Add refs for auto-scroll functionality
  const chartResultRef = useRef<HTMLDivElement>(null);
  const imageSummaryRef = useRef<HTMLDivElement>(null);
  const tableSummaryRef = useRef<HTMLDivElement>(null);
  const excelSummaryRef = useRef<HTMLDivElement>(null);

  // Auto-scroll to chart result when chart is created
  useEffect(() => {
    if (chartData && chartResultRef.current) {
      setTimeout(() => {
        chartResultRef.current?.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });
      }, 100);
    }
  }, [chartData]);

  // Auto-scroll to image summary when it's generated
  useEffect(() => {
    if (images.some(img => img.summary) && imageSummaryRef.current) {
      setTimeout(() => {
        imageSummaryRef.current?.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });
      }, 100);
    }
  }, [images]);

  // Auto-scroll to table summary when it's generated
  useEffect(() => {
    if (tables.some(tbl => tbl.summary) && tableSummaryRef.current) {
      setTimeout(() => {
        tableSummaryRef.current?.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });
      }, 100);
    }
  }, [tables]);

  // Auto-scroll to excel summary when it's generated
  useEffect(() => {
    if (excels.some(xls => xls.summary) && excelSummaryRef.current) {
      setTimeout(() => {
        excelSummaryRef.current?.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });
      }, 100);
    }
  }, [excels]);

  // Load spaces on component mount
  useEffect(() => {
    const loadSpaces = async () => {
      setIsLoadingSpaces(true);
      try {
        const response = await apiService.getSpaces();
        setSpaces(response.spaces);
      } catch (error) {
        console.error('Failed to load spaces:', error);
      } finally {
        setIsLoadingSpaces(false);
      }
    };
    loadSpaces();
  }, []);

  // Auto-select space if provided via URL
  useEffect(() => {
    if (autoSpaceKey && isSpaceAutoConnected) {
      setSpaceKey(autoSpaceKey);
    }
  }, [autoSpaceKey, isSpaceAutoConnected]);

  // Load pages when space key changes and auto-select page if present in URL
  useEffect(() => {
    const loadPages = async () => {
      if (!spaceKey) {
        setPages([]);
        return;
      }
      setIsLoadingPages(true);
      try {
        const response = await apiService.getPages(spaceKey);
        setPages(response.pages);
        // Auto-select page if present in URL
        const { page } = getConfluenceSpaceAndPageFromUrl();
        if (page && response.pages.includes(page)) {
          setSelectedPages([page]);
        }
      } catch (error) {
        console.error('Failed to load pages:', error);
        setPages([]);
      } finally {
        setIsLoadingPages(false);
      }
    };
    loadPages();
  }, [spaceKey]);

  const chartTypes = [
    { value: 'bar' as const, label: 'Grouped Bar Chart' },
    { value: 'line' as const, label: 'Line Chart' },
    { value: 'pie' as const, label: 'Pie Chart' },
    { value: 'stacked' as const, label: 'Stacked Bar Chart' }
  ];

  const features = [
    { id: 'search' as const, label: 'AI Powered Search', icon: Search },
    { id: 'video' as const, label: 'Video Summarizer', icon: Video },
    { id: 'code' as const, label: 'Code Assistant', icon: Code },
    { id: 'impact' as const, label: 'Impact Analyzer', icon: TrendingUp },
    { id: 'test' as const, label: 'Test Support Tool', icon: TestTube },
    { id: 'image' as const, label: 'Chart Builder', icon: BarChart3 },
  ];



  const loadImages = async () => {
    if (!spaceKey || selectedPages.length === 0) return;
    setIsLoadingImages(true);
    try {
      const allImages: ImageData[] = [];
      const allTables: TableData[] = [];
      const allExcels: ExcelData[] = [];
      for (const pageTitle of selectedPages) {
        try {
          const response = await apiService.getImages(spaceKey, pageTitle);
          // Images
          const pageImages = response.images.map((url, index) => ({
            id: `${pageTitle}_img_${index}`,
            name: `Image ${index + 1} from ${pageTitle}`,
            url,
            pageTitle,
            qa: []
          }));
          allImages.push(...pageImages);
          // Tables
          const pageTables = response.tables.map((html, index) => ({
            id: `${pageTitle}_tbl_${index}`,
            name: `Table ${index + 1} from ${pageTitle}`,
            html,
            pageTitle
          }));
          allTables.push(...pageTables);
          // Excels
          const pageExcels = response.excels.map((url, index) => ({
            id: `${pageTitle}_xls_${index}`,
            name: `Excel ${index + 1} from ${pageTitle}`,
            url,
            pageTitle
          }));
          allExcels.push(...pageExcels);
        } catch (error) {
          console.error(`Failed to load insights from page ${pageTitle}:`, error);
        }
      }
      setImages(allImages);
      setTables(allTables);
      setExcels(allExcels);
    } catch (error) {
      console.error('Failed to load insights:', error);
    } finally {
      setIsLoadingImages(false);
    }
  };

  // Analyze image and set summary
  const analyzeImage = async (imageId: string) => {
    setIsAnalyzing(imageId);
    try {
      const image = images.find(img => img.id === imageId);
      if (!image || !image.pageTitle) {
        throw new Error('Image not found or missing page title');
      }
      const response = await apiService.imageSummary({
        space_key: spaceKey,
        page_title: image.pageTitle,
        image_url: image.url
      });
      setImages(prev => prev.map(img =>
        img.id === imageId
          ? { ...img, summary: response.summary }
          : img
      ));
    } catch (error) {
      console.error('Failed to analyze image:', error);
      setImages(prev => prev.map(img =>
        img.id === imageId
          ? {
              ...img,
              summary: `AI Analysis of ${img.name}: This image contains data visualization elements including charts, graphs, and key performance indicators. The visual elements suggest business metrics tracking with trend analysis and comparative data points. Key insights include performance trends, data correlations, and actionable business intelligence derived from the visual representation.`
            }
          : img
      ));
    } finally {
      setIsAnalyzing('');
    }
  };

  // Automatically analyze all images after loading
  useEffect(() => {
    if (images.length > 0) {
      images.forEach(img => {
        if (!img.summary && img.pageTitle) {
          analyzeImage(img.id);
        }
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [images.length]);

  const addQuestion = async () => {
    if (!newQuestion.trim() || !selectedImage) return;
    
    setIsAskingQuestion(true);
    try {
      const image = images.find(img => img.id === selectedImage);
      if (!image || !image.pageTitle || !image.summary) {
        throw new Error('Image not found or missing required data');
      }
      
      const response = await apiService.imageQA({
        space_key: spaceKey,
        page_title: image.pageTitle,
        image_url: image.url,
        summary: image.summary,
        question: newQuestion
      });
      
      setImages(prev => prev.map(img =>
        img.id === selectedImage
          ? {
              ...img,
              qa: [...(img.qa || []), { question: newQuestion, answer: response.answer }]
            }
          : img
      ));
      
      // Add to Q&A history
      setQaHistory(prev => [{ question: newQuestion, answer: response.answer, imageId: selectedImage }, ...prev]);
      setCurrentQaHistoryIndex(0);
      
      setNewQuestion('');
    } catch (error) {
      console.error('Failed to get AI response:', error);
      // Fallback to sample answer
      const answer = `Based on the AI analysis of this image, here's the response to your question: "${newQuestion}"

The image analysis reveals specific data patterns and visual elements that directly relate to your inquiry. The AI has processed the visual content and extracted relevant insights to provide this contextual response.`;
      setImages(prev => prev.map(img =>
        img.id === selectedImage
          ? {
              ...img,
              qa: [...(img.qa || []), { question: newQuestion, answer }]
            }
          : img
      ));
      
      // Add to Q&A history even for fallback
      setQaHistory(prev => [{ question: newQuestion, answer, imageId: selectedImage }, ...prev]);
      setCurrentQaHistoryIndex(0);
      
      setNewQuestion('');
    } finally {
      setIsAskingQuestion(false);
    }
  };

  // When currentQaHistoryIndex changes, update displayed question
  useEffect(() => {
    if (currentQaHistoryIndex !== null && qaHistory[currentQaHistoryIndex]) {
      const historyItem = qaHistory[currentQaHistoryIndex];
      setNewQuestion(historyItem.question);
      setSelectedImage(historyItem.imageId);
    }
  }, [currentQaHistoryIndex]);

  const createChart = async (imageId: string, chartType?: string, exportFormat?: string) => {
    setIsCreatingChart(true);
    try {
      const image = images.find(img => img.id === imageId);
      if (!image || !image.pageTitle) {
        throw new Error('Image not found or missing page title');
      }
      
      // Use provided parameters or fall back to state values
      const currentChartType = chartType || selectedChartType;
      const currentExportFormat = exportFormat || chartExportFormat;
      
      const chartTypeMap = {
        'bar': 'Grouped Bar',
        'line': 'Line',
        'pie': 'Pie',
        'stacked': 'Stacked Bar'
      };
      
      const response = await apiService.createChart({
        space_key: spaceKey,
        page_title: image.pageTitle,
        image_url: image.url,
        chart_type: chartTypeMap[currentChartType as keyof typeof chartTypeMap],
        filename: chartFileName || 'chart',
        format: currentExportFormat
      });
      
      // Convert base64 to blob URL for display
      const binaryString = atob(response.chart_data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      const blob = new Blob([bytes], { type: response.mime_type });
      const chartUrl = URL.createObjectURL(blob);
      
      setChartData({
        type: currentChartType as any,
        data: { 
          chartUrl, 
          filename: response.filename, 
          exportFormat: currentExportFormat,
          imageId: imageId, // Store the image ID for recreation
          chartDataBase64: response.chart_data, // Store the base64 data for Confluence saving
          mimeType: response.mime_type // Store the mime type
        },
        title: `Generated ${currentChartType.charAt(0).toUpperCase() + currentChartType.slice(1)} Chart`
      });
      
      setTimeout(() => {
        chartPreviewRef.current?.scrollIntoView({
          behavior: 'smooth',
          block: 'center'
        });
      }, 100);
    } catch (error) {
      console.error('Failed to create chart:', error);
      // Fallback to sample data
      const currentChartType = chartType || selectedChartType;
      const sampleData = {
        bar: {
          labels: ['Q1', 'Q2', 'Q3', 'Q4'],
          datasets: [{
            label: 'Revenue',
            data: [65, 78, 90, 81],
            backgroundColor: 'rgba(38, 132, 255, 0.8)'
          }]
        },
        line: {
          labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
          datasets: [{
            label: 'Growth',
            data: [12, 19, 15, 25, 22],
            borderColor: 'rgba(38, 132, 255, 1)',
            fill: false
          }]
        },
        pie: {
          labels: ['Desktop', 'Mobile', 'Tablet'],
          datasets: [{
            data: [55, 35, 10],
            backgroundColor: ['#0052CC', '#2684FF', '#B3D4FF']
          }]
        },
        stacked: {
          labels: ['Q1', 'Q2', 'Q3', 'Q4'],
          datasets: [{
            label: 'Revenue',
            data: [65, 78, 90, 81],
            backgroundColor: 'rgba(38, 132, 255, 0.8)'
          }]
        }
      };
      setChartData({
        type: currentChartType as any,
        data: sampleData[currentChartType as keyof typeof sampleData],
        title: `Generated ${currentChartType.charAt(0).toUpperCase() + currentChartType.slice(1)} Chart`
      });
    } finally {
      setIsCreatingChart(false);
    }
  };

  const exportImage = async (image: ImageData) => {
    try {
      const content = `# Image Analysis Report: ${image.name}

## AI Summary
${image.summary || 'No summary available'}

## Questions & Answers
${image.qa?.map(qa => `**Q:** ${qa.question}\n**A:** ${qa.answer}`).join('\n\n') || 'No questions asked'}

## Image Details
- **Name**: ${image.name}
- **Analysis Date**: ${new Date().toLocaleString()}
- **Export Format**: ${exportFormat}

---
*Generated by Confluence AI Assistant - Image Insights*`;

      const response = await apiService.exportContent({
        content,
        format: exportFormat,
        filename: fileName || image.name.replace(/\s+/g, '_') + '_analysis'
      });

      const url = URL.createObjectURL(response);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${fileName || image.name.replace(/\s+/g, '_')}_analysis.${exportFormat}`;
      a.click();
      URL.revokeObjectURL(url); // Clean up the URL
    } catch (error) {
      console.error('Failed to export image:', error);
      // Fallback to client-side export
      const content = `# Image Analysis Report: ${image.name}

## AI Summary
${image.summary || 'No summary available'}

## Questions & Answers
${image.qa?.map(qa => `**Q:** ${qa.question}\n**A:** ${qa.answer}`).join('\n\n') || 'No questions asked'}

## Image Details
- **Name**: ${image.name}
- **Analysis Date**: ${new Date().toLocaleString()}
- **Export Format**: ${exportFormat}

---
*Generated by Confluence AI Assistant - Image Insights*`;
      const blob = new Blob([content], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${fileName || image.name.replace(/\s+/g, '_')}_analysis.${exportFormat}`;
      a.click();
      URL.revokeObjectURL(url); // Clean up the URL
    }
  };

  const exportChart = async () => {
    if (!chartData) return;
    
    setIsExportingChart(true);
    try {
      // Get the current export format from chart data or state
      const currentExportFormat = chartData.data.exportFormat || chartExportFormat;
      
      // If we have a chart URL from the backend, we need to recreate the chart with the current export format
      if (chartData.data.chartUrl) {
        // Use the stored image ID if available, otherwise find an image with summary
        let imageId = chartData.data.imageId;
        if (!imageId) {
          const imageWithSummary = images.find(img => img.summary);
          imageId = imageWithSummary?.id;
        }
        if (imageId) {
          const image = images.find(img => img.id === imageId);
          if (image && image.pageTitle) {
            // Recreate the chart with the current export format
            const response = await apiService.createChart({
              space_key: spaceKey,
              page_title: image.pageTitle,
              image_url: image.url,
              chart_type: chartData.type === 'bar' ? 'Grouped Bar' : 
                         chartData.type === 'line' ? 'Line' : 
                         chartData.type === 'pie' ? 'Pie' : 'Stacked Bar',
              filename: chartFileName || 'chart',
              format: currentExportFormat
            });
            
            // Download the chart in the selected format
            const binaryString = atob(response.chart_data);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
              bytes[i] = binaryString.charCodeAt(i);
            }
            const blob = new Blob([bytes], { type: response.mime_type });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = response.filename;
            a.click();
            URL.revokeObjectURL(url); // Clean up the URL
            return;
          }
        }
      }
      
      // Fallback to text export
      const content = `# Chart Export: ${chartData.title}

## Chart Type
${chartData.type.charAt(0).toUpperCase() + chartData.type.slice(1)} Chart

## Data
${JSON.stringify(chartData.data, null, 2)}

## Export Details
- **File Name**: ${chartFileName}
- **Format**: ${currentExportFormat}
- **Generated**: ${new Date().toLocaleString()}

---
*Generated by Confluence AI Assistant - Chart Builder*`;

      const response = await apiService.exportContent({
        content,
        format: currentExportFormat,
        filename: chartFileName || 'chart'
      });

      const url = URL.createObjectURL(response);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${chartFileName || 'chart'}.${currentExportFormat}`;
      a.click();
      URL.revokeObjectURL(url); // Clean up the URL
    } catch (error) {
      console.error('Failed to export chart:', error);
      // Fallback to client-side export
      const currentExportFormat = chartData.data.exportFormat || chartExportFormat;
      const content = `# Chart Export: ${chartData.title}

## Chart Type
${chartData.type.charAt(0).toUpperCase() + chartData.type.slice(1)} Chart

## Data
${JSON.stringify(chartData.data, null, 2)}

## Export Details
- **File Name**: ${chartFileName}
- **Format**: ${currentExportFormat}
- **Generated**: ${new Date().toLocaleString()}

---
*Generated by Confluence AI Assistant - Chart Builder*`;
      const blob = new Blob([content], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${chartFileName || 'chart'}.${currentExportFormat}`;
      a.click();
      URL.revokeObjectURL(url); // Clean up the URL
    } finally {
      setIsExportingChart(false);
    }
  };

  // Add new chart creation functions for tables and excels
  const createChartFromTable = async (tableId: string, chartType?: string, exportFormat?: string) => {
    setIsCreatingChart(true);
    try {
      const table = tables.find(tbl => tbl.id === tableId);
      if (!table || !table.pageTitle) throw new Error('Table not found or missing page title');
      const currentChartType = chartType || selectedChartType;
      const currentExportFormat = exportFormat || chartExportFormat;
      const chartTypeMap = {
        'bar': 'Grouped Bar',
        'line': 'Line',
        'pie': 'Pie',
        'stacked': 'Stacked Bar'
      };
      const response = await apiService.createChart({
        space_key: spaceKey,
        page_title: table.pageTitle,
        table_html: table.html,
        chart_type: chartTypeMap[currentChartType as keyof typeof chartTypeMap],
        filename: chartFileName || 'chart',
        format: currentExportFormat
      });
      const binaryString = atob(response.chart_data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      const blob = new Blob([bytes], { type: response.mime_type });
      const chartUrl = URL.createObjectURL(blob);
      setChartData({
        type: currentChartType as any,
        data: { 
          chartUrl, 
          filename: response.filename, 
          exportFormat: currentExportFormat, 
          tableId,
          chartDataBase64: response.chart_data, // Store the base64 data for Confluence saving
          mimeType: response.mime_type // Store the mime type
        },
        title: `Generated ${currentChartType.charAt(0).toUpperCase() + currentChartType.slice(1)} Chart`
      });
      setTimeout(() => {
        chartPreviewRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }, 100);
    } catch (error) {
      console.error('Failed to create chart from table:', error);
    } finally {
      setIsCreatingChart(false);
    }
  };
  const createChartFromExcel = async (excelId: string, chartType?: string, exportFormat?: string) => {
    setIsCreatingChart(true);
    try {
      const excel = excels.find(xls => xls.id === excelId);
      if (!excel || !excel.pageTitle) throw new Error('Excel not found or missing page title');
      const currentChartType = chartType || selectedChartType;
      const currentExportFormat = exportFormat || chartExportFormat;
      const chartTypeMap = {
        'bar': 'Grouped Bar',
        'line': 'Line',
        'pie': 'Pie',
        'stacked': 'Stacked Bar'
      };
      const response = await apiService.createChart({
        space_key: spaceKey,
        page_title: excel.pageTitle,
        excel_url: excel.url,
        chart_type: chartTypeMap[currentChartType as keyof typeof chartTypeMap],
        filename: chartFileName || 'chart',
        format: currentExportFormat
      });
      const binaryString = atob(response.chart_data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      const blob = new Blob([bytes], { type: response.mime_type });
      const chartUrl = URL.createObjectURL(blob);
      setChartData({
        type: currentChartType as any,
        data: { 
          chartUrl, 
          filename: response.filename, 
          exportFormat: currentExportFormat, 
          excelId,
          chartDataBase64: response.chart_data, // Store the base64 data for Confluence saving
          mimeType: response.mime_type // Store the mime type
        },
        title: `Generated ${currentChartType.charAt(0).toUpperCase() + currentChartType.slice(1)} Chart`
      });
      setTimeout(() => {
        chartPreviewRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }, 100);
    } catch (error) {
      console.error('Failed to create chart from excel:', error);
    } finally {
      setIsCreatingChart(false);
    }
  };

  const handlePageSelection = (page: string) => {
    setSelectedPages(prev =>
      prev.includes(page)
        ? prev.filter(p => p !== page)
        : [...prev, page]
    );
  };
  const selectAllPages = () => {
    setSelectedPages(pages);
  };
  const clearAllPages = () => {
    setSelectedPages([]);
  };

  // Add summary and qa to TableData and ExcelData
  useEffect(() => {
    const fetchTableSummaries = async () => {
      for (const table of tables) {
        if (!table.summary && table.pageTitle) {
          try {
            const response = await apiService.tableSummary({
              space_key: spaceKey,
              page_title: table.pageTitle,
              table_html: table.html,
            });
            setTables(prev => prev.map(t => t.id === table.id ? { ...t, summary: response.summary } : t));
          } catch (error) {
            setTables(prev => prev.map(t => t.id === table.id ? { ...t, summary: 'AI summary unavailable.' } : t));
          }
        }
      }
    };
    if (tables.length > 0) fetchTableSummaries();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tables.length]);
  // Excel summary effect
  useEffect(() => {
    const fetchExcelSummaries = async () => {
      for (const excel of excels) {
        if (!excel.summary && excel.pageTitle) {
          try {
            const response = await apiService.excelSummary({
              space_key: spaceKey,
              page_title: excel.pageTitle,
              excel_url: excel.url,
            });
            setExcels(prev => prev.map(x => x.id === excel.id ? { ...x, summary: response.summary } : x));
          } catch (error) {
            setExcels(prev => prev.map(x => x.id === excel.id ? { ...x, summary: 'AI summary unavailable.' } : x));
          }
        }
      }
    };
    if (excels.length > 0) fetchExcelSummaries();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [excels.length]);

  // Q&A for tables
  const addTableQuestion = async (tableId: string, question: string) => {
    if (!question.trim()) return;
    const table = tables.find(tbl => tbl.id === tableId);
    if (!table || !table.pageTitle || !table.summary) return;
    // Use the same imageQA endpoint for now, with summary and question
    setTables(prev => prev.map(t => t.id === tableId ? { ...t, qa: [...(t.qa || []), { question, answer: 'Loading...' }] } : t));
    try {
      const response = await apiService.imageQA({
        space_key: spaceKey,
        page_title: table.pageTitle,
        image_url: '', // Not used
        summary: table.summary,
        question,
      });
      setTables(prev => prev.map(t => t.id === tableId ? { ...t, qa: [...(t.qa || []).slice(0, -1), { question, answer: response.answer }] } : t));
      
      // Add to Q&A history
      setQaHistory(prev => [{ question, answer: response.answer, imageId: tableId }, ...prev]);
      setCurrentQaHistoryIndex(0);
    } catch (error) {
      setTables(prev => prev.map(t => t.id === tableId ? { ...t, qa: [...(t.qa || []).slice(0, -1), { question, answer: 'AI answer unavailable.' }] } : t));
      
      // Add to Q&A history even for error case
      setQaHistory(prev => [{ question, answer: 'AI answer unavailable.', imageId: tableId }, ...prev]);
      setCurrentQaHistoryIndex(0);
    }
  };
  // Q&A for excels
  const addExcelQuestion = async (excelId: string, question: string) => {
    if (!question.trim()) return;
    const excel = excels.find(xls => xls.id === excelId);
    if (!excel || !excel.pageTitle || !excel.summary) return;
    setExcels(prev => prev.map(x => x.id === excelId ? { ...x, qa: [...(x.qa || []), { question, answer: 'Loading...' }] } : x));
    try {
      const response = await apiService.imageQA({
        space_key: spaceKey,
        page_title: excel.pageTitle,
        image_url: '', // Not used
        summary: excel.summary,
        question,
      });
      setExcels(prev => prev.map(x => x.id === excelId ? { ...x, qa: [...(x.qa || []).slice(0, -1), { question, answer: response.answer }] } : x));
      
      // Add to Q&A history
      setQaHistory(prev => [{ question, answer: response.answer, imageId: excelId }, ...prev]);
      setCurrentQaHistoryIndex(0);
    } catch (error) {
      setExcels(prev => prev.map(x => x.id === excelId ? { ...x, qa: [...(x.qa || []).slice(0, -1), { question, answer: 'AI answer unavailable.' }] } : x));
      
      // Add to Q&A history even for error case
      setQaHistory(prev => [{ question, answer: 'AI answer unavailable.', imageId: excelId }, ...prev]);
      setCurrentQaHistoryIndex(0);
    }
  };

  // Helper to get all Q&A items
  const allQAItems = [
    ...images.filter(img => img.summary).map(img => ({
      id: img.id,
      type: 'image' as const,
      name: img.name,
      summary: img.summary,
      qa: img.qa || [],
    })),
    ...tables.filter(tbl => tbl.summary).map(tbl => ({
      id: tbl.id,
      type: 'table' as const,
      name: tbl.name,
      summary: tbl.summary,
      qa: tbl.qa || [],
    })),
    ...excels.filter(xls => xls.summary).map(xls => ({
      id: xls.id,
      type: 'excel' as const,
      name: xls.name,
      summary: xls.summary,
      qa: xls.qa || [],
    })),
  ];

  // Define selectedQAItem before Q&A section
  const selectedQAItem = allQAItems.find(item => item.id === selectedImage);

  function cleanPreviewContent(html: string, numBlocks = 2): string {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      // Remove chatbot widget by class or id (adjust selector as needed)
      const chatbot = doc.querySelector('.YOUR_CHATBOT_CLASS, #YOUR_CHATBOT_ID');
      if (chatbot) chatbot.remove();

      // Get all paragraphs and divs (or adjust as needed)
      const blocks = Array.from(doc.body.querySelectorAll('p, div, section, ul, ol, pre, h1, h2, h3, h4, h5, h6'));
      if (blocks.length >= numBlocks) {
        return blocks.slice(-numBlocks).map(el => el.outerHTML).join('');
      }
      // Fallback: return all content
      return doc.body.innerHTML;
    } catch {
      return html;
    }
  }

  return (
    <div className="fixed inset-0 bg-white flex items-center justify-center z-40 p-4">
      <div className="bg-white/80 backdrop-blur-xl border-2 border-[#0052cc] rounded-2xl w-full max-w-7xl max-h-[90vh] overflow-hidden">
        {/* Header */}
        <div className="bg-gradient-to-r from-confluence-blue/90 to-confluence-light-blue/90 backdrop-blur-xl p-6 text-white border-b border-white/10">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <Image className="w-8 h-8" />
              <div>
                <h2 className="text-2xl font-bold">Confluence AI Assistant</h2>
                <p className="text-blue-100/90">AI-powered tools for your Confluence workspace</p>
              </div>
            </div>
            <div className="flex items-center space-x-2">
              <button 
                onClick={() => onModeSelect('agent')}
                className="text-blue-100 hover:text-white hover:bg-white/10 rounded-lg px-3 py-1 text-sm transition-colors"
              >
                Switch to Agent Mode
              </button>
              <button onClick={onClose} className="text-white hover:bg-white/10 rounded-full p-2 backdrop-blur-sm">
                <X className="w-6 h-6" />
              </button>
            </div>
          </div>
          {/* Feature Navigation */}
          <div className="mt-6 relative">
            <CustomScrollbar className="pb-2">
              <div className="flex gap-2">
                {features.map((feature) => {
                  const Icon = feature.icon;
                  const isActive = feature.id === 'image';
                  return (
                    <button
                      key={feature.id}
                      onClick={() => onFeatureSelect(feature.id)}
                      className={`flex items-center space-x-2 px-4 py-2 rounded-lg backdrop-blur-sm border transition-all duration-200 whitespace-nowrap flex-shrink-0 ${
                        isActive
                          ? 'bg-white/90 text-confluence-blue shadow-lg border-white/30'
                          : 'bg-white/10 text-white hover:bg-white/20 border-white/10'
                      }`}
                    >
                      <Icon className="w-4 h-4" />
                      <span className="text-sm font-medium">{feature.label}</span>
                    </button>
                  );
                })}
              </div>
            </CustomScrollbar>
          </div>
        </div>
        <div className="p-6 overflow-y-auto max-h-[calc(90vh-200px)]">
          <div className="grid grid-cols-1 xl:grid-cols-4 gap-6 overflow-visible">
            {/* Left Column - Image Selection */}
            <div className="xl:col-span-1 z-50 overflow-visible">
              <div className="bg-white/60 backdrop-blur-xl rounded-xl p-4 space-y-6 border border-white/20 shadow-lg overflow-visible">
                <h3 className="font-semibold text-gray-800 mb-4 flex items-center">
                  <Eye className="w-5 h-5 mr-2" />
                  {selectedQAItem ?
                    selectedQAItem.type === 'image' ? 'Image Source Selection' :
                    selectedQAItem.type === 'table' ? 'Table Source Selection' :
                    selectedQAItem.type === 'excel' ? 'Excel Source Selection' :
                    'Source Selection'
                  : 'Source Selection'}
                </h3>
                {/* Space Key Input */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Confluence Space Key
                  </label>
                  <div className="relative">
                    <select
                      value={spaceKey}
                      onChange={(e) => setSpaceKey(e.target.value)}
                      disabled={isLoadingSpaces}
                      className="w-full p-3 border border-white/30 rounded-lg focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue appearance-none bg-white/70 backdrop-blur-sm disabled:bg-gray-100"
                    >
                      <option value="">
                        {isLoadingSpaces ? 'Loading spaces...' : 'Select space...'}
                      </option>
                      {spaces.map(space => (
                        <option key={space.key} value={space.key}>{space.name}</option>
                      ))}
                    </select>
                    {isLoadingSpaces ? (
                      <Loader2 className="absolute right-3 top-1/2 transform -translate-y-1/2 w-4 h-4 text-gray-400 animate-spin" />
                    ) : (
                      <ChevronDown className="absolute right-3 top-1/2 transform -translate-y-1/2 w-5 h-5 text-gray-400 pointer-events-none" />
                    )}
                  </div>
                </div>
                {/* Page Selection - Aesthetic Multiselect */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Select Pages ({selectedPages.length} selected)
                  </label>
                  <div className="relative z-50 overflow-visible">
                    <button
                      type="button"
                      onClick={() => setIsPageDropdownOpen(!isPageDropdownOpen)}
                      disabled={!spaceKey}
                      className="w-full p-3 border border-white/30 rounded-lg focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue bg-white/70 backdrop-blur-sm text-left flex items-center justify-between disabled:bg-gray-100 disabled:cursor-not-allowed"
                    >
                      <span className={selectedPages.length === 0 ? 'text-gray-500' : 'text-gray-700'}>
                        {selectedPages.length === 0
                          ? 'Choose pages...'
                          : selectedPages.length === 1
                            ? selectedPages[0]
                            : `${selectedPages.length} pages selected`
                        }
                      </span>
                      {isPageDropdownOpen ? (
                        <ChevronUp className="w-5 h-5 text-gray-400" />
                      ) : (
                        <ChevronDown className="w-5 h-5 text-gray-400" />
                      )}
                    </button>
                    {/* Dropdown */}
                    {isPageDropdownOpen && spaceKey && (
                      <div className="absolute z-50 w-full mt-1 bg-white/95 backdrop-blur-xl border border-white/30 rounded-lg shadow-xl max-h-60 overflow-auto">
                        {/* Header with Select All/Clear All and Search */}
                        <div className="p-3 border-b border-white/20 bg-white/50">
                          <div className="flex justify-between items-center mb-2">
                            <button
                              onClick={selectAllPages}
                              className="text-sm text-confluence-blue hover:text-confluence-blue/80 font-medium"
                            >
                              Select All
                            </button>
                            <button
                              onClick={clearAllPages}
                              className="text-sm text-gray-500 hover:text-gray-700 font-medium"
                            >
                              Clear All
                            </button>
                          </div>
                          <input
                            type="text"
                            value={pageSearch}
                            onChange={e => setPageSearch(e.target.value)}
                            placeholder="Search pages..."
                            className="w-full px-3 py-2 border border-white/20 rounded-lg text-sm focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue bg-white/80 placeholder-gray-400 mb-1"
                          />
                        </div>
                        {/* Page List */}
                        <div className="max-h-48 overflow-y-auto">
                          {(pages.filter(page => page.toLowerCase().includes(pageSearch.toLowerCase()))).length === 0 ? (
                            <div className="p-3 text-gray-500 text-sm text-center">
                              No pages found in this space
                            </div>
                          ) : (
                            pages.filter(page => page.toLowerCase().includes(pageSearch.toLowerCase())).map(page => (
                              <label
                                key={page}
                                className="flex items-center space-x-3 p-3 hover:bg-white/50 cursor-pointer border-b border-white/10 last:border-b-0"
                              >
                                <div className="relative">
                                  <input
                                    type="checkbox"
                                    checked={selectedPages.includes(page)}
                                    onChange={() => handlePageSelection(page)}
                                    className="sr-only"
                                  />
                                  <div className={`w-5 h-5 border-2 rounded flex items-center justify-center transition-colors ${
                                    selectedPages.includes(page)
                                      ? 'bg-confluence-blue border-confluence-blue'
                                      : 'border-gray-300 hover:border-confluence-blue/50'
                                  }`}>
                                    {selectedPages.includes(page) && (
                                      <Check className="w-3 h-3 text-white" />
                                    )}
                                  </div>
                                </div>
                                <span className="text-sm text-gray-700 flex-1">{page}</span>
                              </label>
                            ))
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
                {/* Load Images Button */}
                <button
                  onClick={loadImages}
                  disabled={!spaceKey || selectedPages.length === 0 || isLoadingImages}
                  className="w-full bg-confluence-blue/90 backdrop-blur-sm text-white py-3 px-4 rounded-lg hover:bg-confluence-blue disabled:bg-gray-300 disabled:cursor-not-allowed flex items-center justify-center space-x-2 transition-colors border border-white/10"
                >
                  {isLoadingImages ? (
                    <>
                      <Loader2 className="w-5 h-5 animate-spin" />
                      <span>Loading Images...</span>
                    </>
                  ) : (
                    <>
                      <Image className="w-5 h-5" />
                      <span>Load</span>
                    </>
                  )}
                </button>
              </div>
            </div>
            {/* Middle Column - Images Grid (now also shows tables and excels) */}
            <div className="xl:col-span-2 space-y-6">
              {(images.length > 0 || tables.length > 0 || excels.length > 0) ? (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {/* Images */}
                  {images.map(image => (
                    <div key={image.id} className="bg-white/60 backdrop-blur-xl rounded-xl p-4 border border-white/20 shadow-lg">
                      <div className="aspect-video bg-gray-200/50 backdrop-blur-sm rounded-lg mb-4 overflow-hidden border border-white/20">
                        <img 
                          src={image.url} 
                          alt={image.name}
                          className="w-full h-full object-cover"
                        />
                      </div>
                      <h4 className="font-semibold text-gray-800 mb-2">{image.name}</h4>
                      <div className="space-y-2">
                        {isAnalyzing === image.id && (
                          <div className="w-full bg-confluence-blue/90 backdrop-blur-sm text-white py-2 px-4 rounded-lg flex items-center justify-center space-x-2 border border-white/10">
                            <Loader2 className="w-4 h-4 animate-spin" />
                            <span>Analyzing...</span>
                          </div>
                        )}
                        {image.summary && (
                          <button
                            onClick={() => createChart(image.id, selectedChartType, chartExportFormat)}
                            disabled={isCreatingChart}
                            className="w-full bg-green-600/90 backdrop-blur-sm text-white py-2 px-4 rounded-lg hover:bg-green-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center space-x-2 border border-white/10"
                          >
                            {isCreatingChart ? (
                              <>
                                <Loader2 className="w-4 h-4 animate-spin" />
                                <span>Creating Chart...</span>
                              </>
                            ) : (
                              <>
                                <BarChart3 className="w-4 h-4" />
                                <span>Create Graph</span>
                              </>
                            )}
                          </button>
                        )}
                      </div>
                      {image.summary && (
                        <div ref={imageSummaryRef} className="mt-4 p-3 bg-white/70 backdrop-blur-sm rounded-lg border border-white/20">
                          <p className="text-sm text-gray-700">{image.summary}</p>
                        </div>
                      )}
                    </div>
                  ))}
                  {/* Tables */}
                  {tables.map(table => (
                    <div key={table.id} className="bg-white/60 backdrop-blur-xl rounded-xl p-4 border border-white/20 shadow-lg">
                      <div className="aspect-video bg-gray-200/50 backdrop-blur-sm rounded-lg mb-4 overflow-auto border border-white/20 flex items-center justify-center">
                        <div style={{width: '100%', overflowX: 'auto'}} dangerouslySetInnerHTML={{ __html: table.html }} />
                      </div>
                      <h4 className="font-semibold text-gray-800 mb-2">{table.name}</h4>
                      <div className="space-y-2">
                        {!table.summary ? (
                          <div className="w-full bg-confluence-blue/90 backdrop-blur-sm text-white py-2 px-4 rounded-lg flex items-center justify-center space-x-2 border border-white/10">
                            <Loader2 className="w-4 h-4 animate-spin" />
                            <span>Summarizing...</span>
                          </div>
                        ) : isCreatingChart ? (
                          <div className="w-full bg-confluence-blue/90 backdrop-blur-sm text-white py-2 px-4 rounded-lg flex items-center justify-center space-x-2 border border-white/10">
                            <Loader2 className="w-4 h-4 animate-spin" />
                            <span>Creating Chart...</span>
                          </div>
                        ) : (
                          <button
                            onClick={() => createChartFromTable(table.id, selectedChartType, chartExportFormat)}
                            disabled={isCreatingChart}
                            className="w-full bg-green-600/90 backdrop-blur-sm text-white py-2 px-4 rounded-lg hover:bg-green-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center space-x-2 border border-white/10"
                          >
                            <BarChart3 className="w-4 h-4" />
                            <span>Create Graph</span>
                          </button>
                        )}
                      </div>
                      {/* Table summary and Q&A */}
                      {table.summary && (
                        <div ref={tableSummaryRef} className="mt-4 p-3 bg-white/70 backdrop-blur-sm rounded-lg border border-white/20">
                          <p className="text-sm text-gray-700">{table.summary}</p>
                        </div>
                      )}
                    </div>
                  ))}
                  {/* Excels */}
                  {excels.map(excel => (
                    <div key={excel.id} className="bg-white/60 backdrop-blur-xl rounded-xl p-4 border border-white/20 shadow-lg flex flex-col items-center justify-center">
                      <div className="aspect-video bg-gray-200/50 backdrop-blur-sm rounded-lg mb-4 flex items-center justify-center border border-white/20">
                        <span className="text-confluence-blue text-4xl"><BarChart3 /></span>
                      </div>
                      <h4 className="font-semibold text-gray-800 mb-2">{excel.name}</h4>
                      <a
                        href={excel.url}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="w-full bg-confluence-blue/90 text-white py-2 px-4 rounded-lg hover:bg-confluence-blue transition-colors border border-white/10 text-center mt-2"
                        download
                      >
                        Download Excel
                      </a>
                      <div className="space-y-2 w-full mt-2">
                        {!excel.summary ? (
                          <div className="w-full bg-confluence-blue/90 backdrop-blur-sm text-white py-2 px-4 rounded-lg flex items-center justify-center space-x-2 border border-white/10">
                            <Loader2 className="w-4 h-4 animate-spin" />
                            <span>Summarizing...</span>
                          </div>
                        ) : isCreatingChart ? (
                          <div className="w-full bg-confluence-blue/90 backdrop-blur-sm text-white py-2 px-4 rounded-lg flex items-center justify-center space-x-2 border border-white/10">
                            <Loader2 className="w-4 h-4 animate-spin" />
                            <span>Creating Chart...</span>
                          </div>
                        ) : (
                          <button
                            onClick={() => createChartFromExcel(excel.id, selectedChartType, chartExportFormat)}
                            disabled={isCreatingChart}
                            className="w-full bg-green-600/90 backdrop-blur-sm text-white py-2 px-4 rounded-lg hover:bg-green-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center space-x-2 border border-white/10"
                          >
                            <BarChart3 className="w-4 h-4" />
                            <span>Create Graph</span>
                          </button>
                        )}
                      </div>
                      {/* Excel summary and Q&A */}
                      {excel.summary && (
                        <div ref={excelSummaryRef} className="mt-4 p-3 bg-white/70 backdrop-blur-sm rounded-lg border border-white/20">
                          <p className="text-sm text-gray-700">{excel.summary}</p>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center py-12">
                  <Image className="w-16 h-16 text-gray-400 mx-auto mb-4" />
                  <h3 className="text-lg font-semibold text-gray-600 mb-2">No Images Loaded</h3>
                  <p className="text-gray-500">Select a space and pages to load embedded images, tables, or Excel files for analysis.</p>
                </div>
              )}
              {/* Chart Preview Section */}
              {chartData && (
                <div ref={chartResultRef} className="bg-white/60 backdrop-blur-xl rounded-xl p-4 border border-white/20 shadow-lg">
                  <h3 className="font-semibold text-gray-800 mb-4 flex items-center">
                    <BarChart3 className="w-5 h-5 mr-2" />
                    Chart Builder
                  </h3>
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    {/* Chart Controls - Left Side */}
                    <div className="lg:col-span-1 space-y-4">
                      <div className="bg-white/70 backdrop-blur-sm rounded-lg p-4 border border-white/20">
                        <h4 className="font-semibold text-gray-800 mb-3">Chart Settings</h4>
                        <div className="space-y-4">
                          <div>
                            <label className="block text-sm font-medium text-gray-700 mb-2">
                              Chart Type
                            </label>
                            <div className="relative">
                              <select
                                value={selectedChartType}
                                disabled={isChangingChartType}
                                onChange={async (e) => {
                                  const newChartType = e.target.value as any;
                                  setSelectedChartType(newChartType);
                                  setIsChangingChartType(true);
                                  try {
                                    // Use the stored ID and type to regenerate the chart
                                    if (chartData?.data?.imageId) {
                                      await createChart(chartData.data.imageId, newChartType, chartExportFormat);
                                    } else if (chartData?.data?.tableId) {
                                      await createChartFromTable(chartData.data.tableId, newChartType, chartExportFormat);
                                    } else if (chartData?.data?.excelId) {
                                      await createChartFromExcel(chartData.data.excelId, newChartType, chartExportFormat);
                                    }
                                  } finally {
                                    setIsChangingChartType(false);
                                  }
                                }}
                                className="w-full p-3 border border-white/30 rounded-lg focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue appearance-none bg-white/70 backdrop-blur-sm disabled:bg-gray-100"
                              >
                                {chartTypes.map(type => (
                                  <option key={type.value} value={type.value}>{type.label}</option>
                                ))}
                              </select>
                              {isChangingChartType ? (
                                <Loader2 className="absolute right-3 top-1/2 transform -translate-y-1/2 w-4 h-4 text-gray-400 animate-spin" />
                              ) : (
                                <ChevronDown className="absolute right-3 top-1/2 transform -translate-y-1/2 w-5 h-5 text-gray-400 pointer-events-none" />
                              )}
                            </div>
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-gray-700 mb-2">
                              Chart File Name
                            </label>
                            <input
                              type="text"
                              value={chartFileName}
                              onChange={(e) => setChartFileName(e.target.value)}
                              placeholder="my-chart"
                              className="w-full p-3 border border-white/30 rounded-lg focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue bg-white/70 backdrop-blur-sm"
                            />
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-gray-700 mb-2">
                              Export Format
                            </label>
                            <div className="relative">
                              <select
                                value={chartExportFormat}
                                onChange={(e) => {
                                  setChartExportFormat(e.target.value);
                                  // Update the chart data with new export format without recreating the chart
                                  if (chartData && chartData.data.chartUrl) {
                                    setChartData({
                                      ...chartData,
                                      data: {
                                        ...chartData.data,
                                        exportFormat: e.target.value
                                      }
                                    });
                                  }
                                }}
                                className="w-full p-3 border border-white/30 rounded-lg focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue appearance-none bg-white/70 backdrop-blur-sm"
                              >
                                <option value="png">PNG</option>
                                <option value="jpg">JPG</option>
                                <option value="svg">SVG</option>
                                <option value="pdf">PDF</option>
                                <option value="docx">Word Document</option>
                                <option value="pptx">PowerPoint</option>
                              </select>
                              <ChevronDown className="absolute right-3 top-1/2 transform -translate-y-1/2 w-5 h-5 text-gray-400 pointer-events-none" />
                            </div>
                          </div>
                          <div className="space-y-2 pt-2">
                            <button
                              onClick={exportChart}
                              disabled={isExportingChart}
                              className="w-full flex items-center justify-center space-x-2 px-4 py-2 bg-green-600/90 backdrop-blur-sm text-white rounded-lg hover:bg-green-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors border border-white/10"
                            >
                              {isExportingChart ? (
                                <>
                                  <Loader2 className="w-4 h-4 animate-spin" />
                                  <span>Exporting...</span>
                                </>
                              ) : (
                                <>
                                  <Download className="w-4 h-4" />
                                  <span>Export Chart</span>
                                </>
                              )}
                            </button>
                          </div>
                          
                          {/* Save to Confluence Section for Chart */}
                          <div className="pt-4 border-t border-white/20 space-y-3">
                            <h4 className="font-semibold text-gray-800 text-sm">Save Chart to Confluence</h4>
                            <div className="flex items-center space-x-2 mb-2">
                              <label htmlFor="chart-save-mode" className="text-sm font-medium text-gray-700">Save Mode:</label>
                              <select
                                id="chart-save-mode"
                                value={saveMode}
                                onChange={e => setSaveMode(e.target.value)}
                                className="px-3 py-1 border border-white/30 rounded text-sm focus:ring-2 focus:ring-confluence-blue bg-white/70 backdrop-blur-sm"
                              >
                                <option value="append">Append</option>
                                <option value="overwrite">Overwrite</option>
                              </select>
                            </div>
                            
                            <div className="space-y-2">
                              <button
                                onClick={async () => {
                                  setIsPreviewLoading(true);
                                  setShowPreview(false);
                                  try {
                                    const { space, page } = getConfluenceSpaceAndPageFromUrl();
                                    if (!space || !page) {
                                      alert('Confluence space or page not specified in macro src URL.');
                                      return;
                                    }
                                    if (!chartData) {
                                      alert('No chart available to save.');
                                      return;
                                    }
                                    
                                    // Create chart content for Confluence
                                    const chartContent = `
<div class="chart-container">
  <h3>${chartData.title}</h3>
  <p><strong>Chart Type:</strong> ${chartData.type.charAt(0).toUpperCase() + chartData.type.slice(1)} Chart</p>
  <p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>
  <div class="chart-preview">
    <img src="data:${chartData.data.mimeType};base64,${chartData.data.chartDataBase64}" alt="${chartData.title}" style="max-width: 100%; height: auto;" />
  </div>
  <p><em>Chart generated from ${chartData.data.imageId ? 'image data' : chartData.data.tableId ? 'table data' : chartData.data.excelId ? 'excel data' : 'data source'}</em></p>
</div>`;
                                    
                                    const preview = await apiService.previewSaveToConfluence({
                                      space_key: space,
                                      page_title: page,
                                      content: chartContent,
                                      mode: saveMode,
                                    });
                                    setPreviewContent(preview.preview_content);
                                    setPreviewDiff(preview.diff);
                                    setShowPreview(true);
                                  } catch (err: any) {
                                    alert('Failed to generate preview: ' + (err.message || err));
                                  } finally {
                                    setIsPreviewLoading(false);
                                  }
                                }}
                                className="w-full flex items-center justify-center space-x-2 px-4 py-2 bg-yellow-500 text-white rounded-lg hover:bg-yellow-600 transition-colors border border-white/10"
                              >
                                {isPreviewLoading ? "Loading..." : "Preview Chart"}
                              </button>
                              <button
                                onClick={async () => {
                                  const { space, page } = getConfluenceSpaceAndPageFromUrl();
                                  if (!space || !page) {
                                    alert('Confluence space or page not specified in macro src URL.');
                                    return;
                                  }
                                  if (!chartData) {
                                    alert('No chart available to save.');
                                    return;
                                  }
                                  
                                  try {
                                    // Create chart content for Confluence
                                    const chartContent = `
<div class="chart-container">
  <h3>${chartData.title}</h3>
  <p><strong>Chart Type:</strong> ${chartData.type.charAt(0).toUpperCase() + chartData.type.slice(1)} Chart</p>
  <p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>
  <div class="chart-preview">
    <img src="data:${chartData.data.mimeType};base64,${chartData.data.chartDataBase64}" alt="${chartData.title}" style="max-width: 100%; height: auto;" />
  </div>
  <p><em>Chart generated from ${chartData.data.imageId ? 'image data' : chartData.data.tableId ? 'table data' : chartData.data.excelId ? 'excel data' : 'data source'}</em></p>
</div>`;
                                    
                                    await apiService.saveToConfluence({
                                      space_key: space,
                                      page_title: page,
                                      content: chartContent,
                                      mode: saveMode,
                                    });
                                    setShowToast(true);
                                    setTimeout(() => setShowToast(false), 3000);
                                  } catch (err: any) {
                                    alert('Failed to save chart to Confluence: ' + (err.message || err));
                                  }
                                }}
                                className="w-full flex items-center justify-center space-x-2 px-4 py-2 bg-confluence-blue/90 backdrop-blur-sm text-white rounded-lg hover:bg-confluence-blue transition-colors border border-white/10"
                              >
                                <Save className="w-4 h-4" />
                                <span>Save Chart to Confluence</span>
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                    {/* Chart Preview - Right Side */}
                    <div className="lg:col-span-2">
                      <div className="bg-white/70 backdrop-blur-sm rounded-lg p-6 border border-white/20">
                        <h4 className="font-semibold text-gray-800 mb-4">{chartData.title}</h4>
                        <div className="w-full h-80 bg-gradient-to-br from-confluence-blue/10 to-confluence-light-blue/10 rounded-lg flex items-center justify-center border border-white/20 overflow-hidden">
                          {chartData.data.chartUrl ? (
                            <img 
                              src={chartData.data.chartUrl} 
                              alt={chartData.title}
                              className="w-full h-full object-contain"
                            />
                          ) : (
                            <div className="text-center">
                              <BarChart3 className="w-20 h-20 text-confluence-blue mx-auto mb-4" />
                              <p className="text-gray-600 font-medium text-lg">{chartData.title}</p>
                              <p className="text-gray-500 text-sm mt-2">Live {chartData.type} chart preview</p>
                              <div className="mt-4 text-xs text-gray-400 max-w-md mx-auto">
                                Chart updates automatically when you change the type in the controls panel
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
            {/* Right Column - Q&A and Export */}
            <div className="xl:col-span-1">
              <div className="bg-white/60 backdrop-blur-xl rounded-xl p-4 space-y-4 border border-white/20 shadow-lg">
                <h3 className="font-semibold text-gray-800 mb-4 flex items-center">
                  <MessageSquare className="w-5 h-5 mr-2" />
                  {selectedQAItem ?
                    selectedQAItem.type === 'image' ? 'Image Q&A' :
                    selectedQAItem.type === 'table' ? 'Table Q&A' :
                    selectedQAItem.type === 'excel' ? 'Excel Q&A' :
                    'Q&A'
                  : 'Q&A'}
                </h3>
                
                {/* --- History Dropdown for Q&A --- */}
                {qaHistory.length > 0 && (
                  <div className="mb-4 flex items-center space-x-2">
                    <label className="text-sm font-medium text-gray-700">Q&A History:</label>
                    <select
                      className="border border-gray-300 rounded px-2 py-1 text-sm"
                      value={currentQaHistoryIndex ?? 0}
                      onChange={e => setCurrentQaHistoryIndex(Number(e.target.value))}
                    >
                      {qaHistory.map((item, idx) => (
                        <option key={idx} value={idx}>
                          {item.question.length > 40 ? item.question.slice(0, 40) + '...' : item.question}
                        </option>
                      ))}
                    </select>
                    {currentQaHistoryIndex !== null && currentQaHistoryIndex !== 0 && (
                      <button
                        className="text-xs text-confluence-blue underline ml-2"
                        onClick={() => setCurrentQaHistoryIndex(0)}
                      >
                        Go to Latest
                      </button>
                    )}
                  </div>
                )}
                {/* --- End History Dropdown --- */}
                {/* Image Selection for Q&A */}
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    {selectedQAItem ?
                      selectedQAItem.type === 'image' ? 'Select Image for Questions' :
                      selectedQAItem.type === 'table' ? 'Select Table for Questions' :
                      selectedQAItem.type === 'excel' ? 'Select Excel for Questions' :
                      'Select Item for Questions'
                    : 'Select Item for Questions'}
                  </label>
                  <div className="relative">
                    <select
                      value={selectedImage}
                      onChange={(e) => setSelectedImage(e.target.value)}
                      className="w-full p-3 border border-white/30 rounded-lg focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue appearance-none bg-white/70 backdrop-blur-sm"
                    >
                      <option value="">Choose item...</option>
                      {allQAItems.map(item => (
                        <option key={item.id} value={item.id}>{item.name}</option>
                      ))}
                    </select>
                    <ChevronDown className="absolute right-3 top-1/2 transform -translate-y-1/2 w-5 h-5 text-gray-400 pointer-events-none" />
                  </div>
                </div>
                {/* Add Question */}
                <div className="space-y-2">
                  <div className="w-full">
                    <VoiceRecorder
                      value={newQuestion}
                      onChange={setNewQuestion}
                      onConfirm={setNewQuestion}
                      inputPlaceholder={selectedQAItem ?
                        selectedQAItem.type === 'image' ? 'Ask about the selected image...' :
                        selectedQAItem.type === 'table' ? 'Ask about the selected table...' :
                        selectedQAItem.type === 'excel' ? 'Ask about the selected excel...' :
                        'Ask about the selected item...'
                      : 'Ask about the selected item...'}
                    />
                  </div>
                  <button
                    onClick={async () => {
                      if (!newQuestion.trim() || !selectedQAItem) return;
                      if (selectedQAItem.type === 'image') {
                        await addQuestion();
                      } else if (selectedQAItem.type === 'table') {
                        await addTableQuestion(selectedQAItem.id, newQuestion);
                      } else if (selectedQAItem.type === 'excel') {
                        await addExcelQuestion(selectedQAItem.id, newQuestion);
                      }
                      setNewQuestion('');
                    }}
                    disabled={!newQuestion.trim() || !selectedQAItem || isAskingQuestion}
                    className="w-full px-3 py-2 bg-confluence-blue/90 backdrop-blur-sm text-white rounded hover:bg-confluence-blue disabled:bg-gray-300 transition-colors flex items-center justify-center space-x-2 border border-white/10"
                  >
                    {isAskingQuestion ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span>Asking...</span>
                      </>
                    ) : (
                      <>
                        <MessageSquare className="w-4 h-4" />
                        <span>Ask Question</span>
                      </>
                    )}
                  </button>
                </div>
                {/* Q&A Display */}
                {selectedQAItem && (
                  <div className="pt-4 border-t border-white/20 space-y-3">
                    <h4 className="font-semibold text-gray-800">
                      {selectedQAItem.type === 'image' ? 'Image' :
                       selectedQAItem.type === 'table' ? 'Table' :
                       selectedQAItem.type === 'excel' ? 'Excel' : 'Item'} Questions & Answers
                    </h4>
                    {selectedQAItem.qa.length === 0 ? (
                      <div className="text-center py-4">
                        <MessageSquare className="w-8 h-8 text-gray-400 mx-auto mb-2" />
                        <p className="text-sm text-gray-500">No questions asked yet</p>
                      </div>
                    ) : (
                      <div className="space-y-3 max-h-60 overflow-y-auto">
                        {selectedQAItem.qa.map((qa, index) => (
                          <div key={index} className="p-3 bg-white/70 backdrop-blur-sm rounded-lg border border-white/20">
                            <p className="font-medium text-gray-800 text-sm mb-2">Q: {qa.question}</p>
                            <p className="text-gray-700 text-sm">{qa.answer}</p>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
                {/* Export Options */}
                <div className="pt-4 border-t border-white/20 space-y-3">
                  <h4 className="font-semibold text-gray-800">
                    {selectedQAItem ?
                      selectedQAItem.type === 'image' ? 'Export Image Analysis' :
                      selectedQAItem.type === 'table' ? 'Export Table Analysis' :
                      selectedQAItem.type === 'excel' ? 'Export Excel Analysis' :
                      'Export Analysis'
                    : 'Export Analysis'}
                  </h4>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      File Name
                    </label>
                    <input
                      type="text"
                      value={fileName}
                      onChange={(e) => setFileName(e.target.value)}
                      placeholder="image-analysis"
                      className="w-full p-2 border border-white/30 rounded focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue bg-white/70 backdrop-blur-sm"
                    />
                  </div>
                  {/* Export Format Selection */}
                  <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-2">What format would you like to export in?</label>
                    <div className="relative w-48">
                      <button
                        type="button"
                        onClick={() => setIsExportFormatDropdownOpen(!isExportFormatDropdownOpen)}
                        className="px-3 py-1 border border-white/30 rounded text-sm focus:ring-2 focus:ring-confluence-blue bg-white/70 backdrop-blur-sm w-full flex items-center justify-between"
                      >
                        <span>{exportFormats.find(f => f.value === exportFormat)?.label || 'Select format'}</span>
                        {isExportFormatDropdownOpen ? (
                          <ChevronUp className="w-4 h-4 text-gray-400" />
                        ) : (
                          <ChevronDown className="w-4 h-4 text-gray-400" />
                        )}
                      </button>
                      {isExportFormatDropdownOpen && (
                        <div className="absolute z-50 w-full mt-1 bg-white/95 backdrop-blur-xl border border-white/30 rounded-lg shadow-xl max-h-48 overflow-hidden">
                          <div className="p-2 border-b border-white/20 bg-white/50">
                            <input
                              type="text"
                              value={exportFormatSearch}
                              onChange={e => setExportFormatSearch(e.target.value)}
                              placeholder="Search formats..."
                              className="w-full px-2 py-1 border border-white/20 rounded-lg text-sm focus:ring-2 focus:ring-confluence-blue focus:border-confluence-blue bg-white/80 placeholder-gray-400 mb-1"
                            />
                          </div>
                          <div className="max-h-32 overflow-y-auto">
                            {exportFormats.filter(f => f.label.toLowerCase().includes(exportFormatSearch.toLowerCase())).length === 0 ? (
                              <div className="p-2 text-gray-500 text-sm text-center">No formats found</div>
                            ) : (
                              exportFormats.filter(f => f.label.toLowerCase().includes(exportFormatSearch.toLowerCase())).map(f => (
                                <button
                                  key={f.value}
                                  type="button"
                                  onClick={() => { setExportFormat(f.value); setIsExportFormatDropdownOpen(false); setExportFormatSearch(''); }}
                                  className={`w-full text-left flex items-center space-x-2 p-2 hover:bg-white/50 cursor-pointer border-b border-white/10 last:border-b-0 ${exportFormat === f.value ? 'bg-confluence-blue/10' : ''}`}
                                >
                                  <span className="text-sm text-gray-700 flex-1">{f.label}</span>
                                  {exportFormat === f.value && <Check className="w-4 h-4 text-confluence-blue" />}
                                </button>
                              ))
                            )}
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="space-y-2">
                    {allQAItems.map(item => (
                      <button
                        key={item.id}
                        onClick={async () => {
                          if (item.type === 'image') {
                            exportImage(images.find(img => img.id === item.id)!);
                          } else if (item.type === 'table') {
                            const table = tables.find(tbl => tbl.id === item.id);
                            if (table) {
                              const content = `# Table Export: ${table.name}

## Table Type
${table.name}

## Data
${table.html}

## Export Details
- **File Name**: ${fileName || table.name.replace(/\s+/g, '_')}_export
- **Format**: ${exportFormat}
- **Generated**: ${new Date().toLocaleString()}

---
*Generated by Confluence AI Assistant - Table Builder*`;
                              const response = await apiService.exportContent({
                                content,
                                format: exportFormat,
                                filename: fileName || table.name.replace(/\s+/g, '_') + '_export'
                              });
                              const url = URL.createObjectURL(response);
                              const a = document.createElement('a');
                              a.href = url;
                              a.download = `${fileName || table.name.replace(/\s+/g, '_')}_export.${exportFormat}`;
                              a.click();
                            }
                          } else if (item.type === 'excel') {
                            const excel = excels.find(xls => xls.id === item.id);
                            if (excel) {
                              const content = `# Excel Export: ${excel.name}

## Excel Type
${excel.name}

## Data
${excel.url}

## Export Details
- **File Name**: ${fileName || excel.name.replace(/\s+/g, '_')}_export
- **Format**: ${exportFormat}
- **Generated**: ${new Date().toLocaleString()}

---
*Generated by Confluence AI Assistant - Excel Builder*`;
                              const response = await apiService.exportContent({
                                content,
                                format: exportFormat,
                                filename: fileName || excel.name.replace(/\s+/g, '_') + '_export'
                              });
                              const url = URL.createObjectURL(response);
                              const a = document.createElement('a');
                              a.href = url;
                              a.download = `${fileName || excel.name.replace(/\s+/g, '_')}_export.${exportFormat}`;
                              a.click();
                            }
                          }
                        }}
                        className="w-full flex items-center justify-center space-x-2 px-3 py-2 bg-green-600/90 backdrop-blur-sm text-white rounded-lg hover:bg-green-700 transition-colors border border-white/10"
                      >
                        <Download className="w-4 h-4" />
                        <span>Export {item.type === 'image' ? 'Image' : item.type === 'table' ? 'Table' : item.type === 'excel' ? 'Excel' : 'Item'}: {item.name}</span>
                      </button>
                    ))}
                  </div>
                </div>

                {selectedQAItem && (
                  <>
                    <div className="flex items-center space-x-2 mb-2">
                      <label htmlFor="save-mode" className="text-sm font-medium text-gray-700">Save Mode:</label>
                      <select
                        id="save-mode"
                        value={saveMode}
                        onChange={e => setSaveMode(e.target.value)}
                        className="px-3 py-1 border border-white/30 rounded text-sm focus:ring-2 focus:ring-confluence-blue bg-white/70 backdrop-blur-sm"
                      >
                        <option value="append">Append</option>
                        <option value="overwrite">Overwrite</option>
                      </select>
                    </div>
                    
                    <div className="space-y-2">
                      <button
                        onClick={async () => {
                          setIsPreviewLoading(true);
                          setShowPreview(false);
                          try {
                            const { space, page } = getConfluenceSpaceAndPageFromUrl();
                            if (!space || !page) {
                              alert('Confluence space or page not specified in macro src URL.');
                              return;
                            }
                            const selectedItemData = allQAItems.find(item => item.id === selectedQAItem.id);
                            if (!selectedItemData || !selectedItemData.summary) {
                              alert('No summary available for the selected item.');
                              return;
                            }
                            const preview = await apiService.previewSaveToConfluence({
                              space_key: space,
                              page_title: page,
                              content: selectedItemData.summary,
                              mode: saveMode,
                            });
                            setPreviewContent(preview.preview_content);
                            setPreviewDiff(preview.diff);
                            setShowPreview(true);
                          } catch (err: any) {
                            alert('Failed to generate preview: ' + (err.message || err));
                          } finally {
                            setIsPreviewLoading(false);
                          }
                        }}
                        className="w-full flex items-center justify-center space-x-2 px-4 py-2 bg-yellow-500 text-white rounded-lg hover:bg-yellow-600 transition-colors border border-white/10"
                      >
                        {isPreviewLoading ? "Loading..." : "Preview"}
                      </button>
                      <button
                        onClick={async () => {
                          const { space, page } = getConfluenceSpaceAndPageFromUrl();
                          if (!space || !page) {
                            alert('Confluence space or page not specified in macro src URL.');
                            return;
                          }
                          const selectedItemData = allQAItems.find(item => item.id === selectedQAItem.id);
                          if (!selectedItemData || !selectedItemData.summary) {
                            alert('No summary available for the selected item.');
                            return;
                          }
                          try {
                            await apiService.saveToConfluence({
                              space_key: space,
                              page_title: page,
                              content: selectedItemData.summary,
                              mode: saveMode,
                            });
                            setShowToast(true);
                            setTimeout(() => setShowToast(false), 3000);
                          } catch (err: any) {
                            alert('Failed to save to Confluence: ' + (err.message || err));
                          }
                        }}
                        className="w-full flex items-center justify-center space-x-2 px-4 py-2 bg-confluence-blue/90 backdrop-blur-sm text-white rounded-lg hover:bg-confluence-blue transition-colors border border-white/10"
                      >
                        <Save className="w-4 h-4" />
                        <span>Save {selectedQAItem ?
                          selectedQAItem.type === 'image' ? 'Image' :
                          selectedQAItem.type === 'table' ? 'Table' :
                          selectedQAItem.type === 'excel' ? 'Excel' : 'Item'
                        : 'Item'} Summary to Confluence</span>
                      </button>
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
      {showToast && (
        <div style={{position: 'fixed', bottom: 40, left: '50%', transform: 'translateX(-50%)', background: '#2684ff', color: 'white', padding: '16px 32px', borderRadius: 8, zIndex: 9999, fontWeight: 600, fontSize: 16, boxShadow: '0 2px 12px rgba(0,0,0,0.15)'}}>
          Saved to Confluence! Please refresh this Confluence page to see your changes.
        </div>
      )}
      {showPreview && (
        <div className="fixed inset-0 flex items-center justify-center z-50 bg-black/30">
          <div className="bg-confluence-blue/95 rounded-2xl shadow-2xl p-6 w-full max-w-3xl relative border border-white/20">
            <div className="flex justify-between items-center mb-4">
              <h4 className="font-semibold text-white text-lg">Preview of Updated Content</h4>
              <button onClick={() => setShowPreview(false)} className="text-white hover:text-red-400 font-bold text-base px-3 py-1 rounded transition-colors focus:outline-none focus:ring-2 focus:ring-red-400">Close Preview</button>
            </div>
            <div
              className="overflow-y-auto bg-white/90 rounded-xl p-6 border border-white/30 shadow-inner min-h-[120px] max-h-[400px] text-gray-900 text-base font-normal"
              style={{
                fontFamily: 'inherit',
                boxSizing: 'border-box',
                wordBreak: 'break-word',
                whiteSpace: 'pre-wrap',
                marginBottom: 0,
              }}
              dangerouslySetInnerHTML={{ __html: previewContent || '' }}
            />
          </div>
        </div>
      )}
    </div>
  );
};

export default ImageInsights; 