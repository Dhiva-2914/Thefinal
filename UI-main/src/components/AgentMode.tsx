import React, { useState, useEffect } from 'react';
import { Zap, X, Send, Download, RotateCcw, FileText, Brain, CheckCircle, Loader2, MessageSquare, Plus } from 'lucide-react';
import type { AppMode } from '../App';
import { apiService, analyzeGoal } from '../services/api';

interface AgentModeProps {
  onClose: () => void;
  onModeSelect: (mode: AppMode) => void;
}

interface PlanStep {
  id: number;
  title: string;
  status: 'pending' | 'running' | 'completed';
  details?: string;
}

interface OutputTab {
  id: string;
  label: string;
  icon: React.ComponentType<any>;
  content: string;
}

const AgentMode: React.FC<AgentModeProps> = ({ onClose, onModeSelect }) => {
  const [goal, setGoal] = useState('');
  const [isPlanning, setIsPlanning] = useState(false);
  const [isExecuting, setIsExecuting] = useState(false);
  const [planSteps, setPlanSteps] = useState<PlanStep[]>([]);
  const [currentStep, setCurrentStep] = useState(0);
  const [activeTab, setActiveTab] = useState('answer');
  const [followUpQuestion, setFollowUpQuestion] = useState('');
  const [showFollowUp, setShowFollowUp] = useState(false);
  const [outputTabs, setOutputTabs] = useState<OutputTab[]>([]);

  // Add new state for space/page selection and API results
  const [spaces, setSpaces] = useState<{ name: string; key: string }[]>([]);
  const [pages, setPages] = useState<string[]>([]);
  const [selectedSpace, setSelectedSpace] = useState('');
  const [selectedPages, setSelectedPages] = useState<string[]>([]);
  const [error, setError] = useState('');

  // Load spaces on mount
  useEffect(() => {
    const loadSpaces = async () => {
      try {
        const result = await apiService.getSpaces();
        setSpaces(result.spaces);
      } catch (err) {
        setError('Failed to load spaces.');
      }
    };
    loadSpaces();
  }, []);

  // Load pages when space is selected
  useEffect(() => {
    if (selectedSpace) {
      const loadPages = async () => {
        try {
          const result = await apiService.getPages(selectedSpace);
          setPages(result.pages);
        } catch (err) {
          setError('Failed to load pages.');
        }
      };
      loadPages();
    }
  }, [selectedSpace]);

  const handleGoalSubmit = async () => {
    if (!goal.trim() || !selectedSpace || selectedPages.length === 0) {
      setError('Please enter a goal, select a space, and at least one page.');
      return;
    }
    setIsPlanning(true);
    setError('');
    setPlanSteps([
      { id: 1, title: 'Analyzing Goal', status: 'pending' },
      { id: 2, title: 'Executing', status: 'pending' },
    ]);
    setOutputTabs([]);
    setCurrentStep(0);
    setActiveTab('final-answer');
    let toolsToUse: string[] = [];
    let orchestrationReasoning = '';
    try {
      // Step 1: Use Gemini to analyze the goal and get tools
      setPlanSteps((steps) => steps.map((s) => s.id === 1 ? { ...s, status: 'running' } : s));
      setCurrentStep(0);
      const analysis = await analyzeGoal(goal, selectedPages); // Only pass user-selected pages
      toolsToUse = analysis.tools || [];
      let selectedPagesFromAI = analysis.pages || [];
      // Only allow pages that the user selected
      selectedPagesFromAI = selectedPagesFromAI.filter((p: string) => selectedPages.includes(p));
      orchestrationReasoning = analysis.reasoning || '';
      setPlanSteps((steps) => steps.map((s) => s.id === 1 ? { ...s, status: 'completed' } : s));
      setCurrentStep(1);
      // Step 2: Call the selected tools on the selected pages from AI
      setPlanSteps((steps) => steps.map((s) => s.id === 2 ? { ...s, status: 'running' } : s));
      const toolResults: Record<string, any> = {};
      // AI Powered Search
      if (toolsToUse.includes('ai_powered_search')) {
        const res = await apiService.search({
          space_key: selectedSpace,
          page_titles: selectedPagesFromAI,
          query: goal,
        });
        toolResults['AI Powered Search'] = res;
      }
      // Impact Analyzer (requires at least 2 pages)
      if (toolsToUse.includes('impact_analyzer') && selectedPagesFromAI.length >= 2) {
        const res = await apiService.impactAnalyzer({
          space_key: selectedSpace,
          old_page_title: selectedPagesFromAI[0],
          new_page_title: selectedPagesFromAI[1],
          question: goal,
        });
        toolResults['Impact Analyzer'] = res;
      }
      // Code Assistant
      if (toolsToUse.includes('code_assistant') && selectedPagesFromAI.length > 0) {
        const res = await apiService.codeAssistant({
          space_key: selectedSpace,
          page_title: selectedPagesFromAI[0],
          instruction: goal,
        });
        toolResults['Code Assistant'] = res;
      }
      // Video Summarizer
      if (toolsToUse.includes('video_summarizer') && selectedPagesFromAI.length > 0) {
        const res = await apiService.videoSummarizer({
          space_key: selectedSpace,
          page_title: selectedPagesFromAI[0],
        });
        toolResults['Video Summarizer'] = res;
      }
      // Test Support
      if (toolsToUse.includes('test_support') && selectedPagesFromAI.length > 0) {
        const res = await apiService.testSupport({
          space_key: selectedSpace,
          code_page_title: selectedPagesFromAI[0],
        });
        toolResults['Test Support'] = res;
      }
      // Image Insights
      if (toolsToUse.includes('image_insights') && selectedPagesFromAI.length > 0) {
        const images = await apiService.getImages(selectedSpace, selectedPagesFromAI[0]);
        if (images && images.images && images.images.length > 0) {
          const summaries = await Promise.all(images.images.map((imgUrl: string) => apiService.imageSummary({
            space_key: selectedSpace,
            page_title: selectedPagesFromAI[0],
            image_url: imgUrl,
          })));
          toolResults['Image Insights'] = summaries;
        }
      }
      // Chart Builder
      if (toolsToUse.includes('chart_builder') && selectedPagesFromAI.length > 0) {
        const images = await apiService.getImages(selectedSpace, selectedPagesFromAI[0]);
        if (images && images.images && images.images.length > 0) {
          const charts = await Promise.all(images.images.map((imgUrl: string) => apiService.createChart({
            space_key: selectedSpace,
            page_title: selectedPagesFromAI[0],
            image_url: imgUrl,
            chart_type: 'bar',
            filename: 'chart',
            format: 'png',
          })));
          toolResults['Chart Builder'] = charts;
        }
      }
      setPlanSteps((steps) => steps.map((s) => s.id === 2 ? { ...s, status: 'completed' } : s));
      setCurrentStep(planSteps.length - 1); // Ensure progress is 100% when done
      // Prepare output tabs
      const getRelevantOutput = (result: any) => {
        if (!result) return '';
        if (typeof result === 'string') return result;
        if (result.summary) return result.summary;
        if (result.impact_analysis) return result.impact_analysis;
        if (result.modified_code) return result.modified_code;
        if (result.converted_code) return result.converted_code;
        if (result.original_code) return result.original_code;
        if (result.response) return result.response;
        if (result.test_strategy) return result.test_strategy;
        if (result.chart_data) return '[Chart Image]';
        if (Array.isArray(result) && result.length > 0) return getRelevantOutput(result[0]);
        return '';
      };
      const usedToolsContent = Object.entries(toolResults).map(([tool, result]) => {
        const output = getRelevantOutput(result);
        return `## ${tool}\n${output}`;
      }).join('\n\n');
      const finalAnswer = Object.values(toolResults).map(getRelevantOutput).filter(Boolean).join('\n\n');
      const tabs = [
        {
          id: 'final-answer',
          label: 'Final Answer',
          icon: FileText,
          content: finalAnswer,
        },
        {
          id: 'reasoning',
          label: 'Reasoning',
          icon: Brain,
          content: orchestrationReasoning,
        },
        {
          id: 'selected-pages',
          label: 'Selected Pages',
          icon: FileText,
          content: selectedPagesFromAI.join(', '),
        },
        {
          id: 'used-tools',
          label: 'Used Tools',
          icon: Zap,
          content: usedToolsContent,
        },
      ];
      setOutputTabs(tabs);
      setActiveTab('final-answer');
    } catch (err: any) {
      setError(err.message || 'An error occurred during orchestration.');
    } finally {
      setIsPlanning(false);
    }
  };

  const executeSteps = async (steps: PlanStep[]) => {
    setIsExecuting(true);
    
    for (let i = 0; i < steps.length; i++) {
      setCurrentStep(i);
      
      // Update step to running
      setPlanSteps(prev => prev.map(step => 
        step.id === i + 1 
          ? { ...step, status: 'running', details: getStepDetails(i) }
          : step
      ));
      
      // Simulate step execution
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      // Update step to completed
      setPlanSteps(prev => prev.map(step => 
        step.id === i + 1 
          ? { ...step, status: 'completed', details: getCompletedDetails(i) }
          : step
      ));
    }
    
    // Generate output tabs
    const tabs: OutputTab[] = [
      {
        id: 'answer',
        label: 'Final Answer',
        icon: FileText,
        content: generateFinalAnswer()
      },
      {
        id: 'reasoning',
        label: 'Reasoning Steps',
        icon: Brain,
        content: generateReasoningSteps()
      },
      {
        id: 'tools',
        label: 'Used Tools',
        icon: Zap,
        content: generateUsedTools()
      },
      {
        id: 'qa',
        label: 'Follow-Up Q&A',
        icon: MessageSquare,
        content: 'Ask follow-up questions to refine or expand on this analysis.'
      }
    ];
    
    setOutputTabs(tabs);
    setIsExecuting(false);
    setShowFollowUp(true);
  };

  const getStepDetails = (stepIndex: number) => {
    const details = [
      '🔍 Searching Confluence...',
      '📊 Analyzing content...',
      '💡 Generating recommendations...'
    ];
    return details[stepIndex];
  };

  const getCompletedDetails = (stepIndex: number) => {
    const details = [
      '✅ Found 3 relevant pages',
      '✅ Content summarized',
      '✅ Recommendations generated'
    ];
    return details[stepIndex];
  };

  const generateFinalAnswer = () => {
    return `Based on your goal: "${goal}"

## Analysis Summary
I've analyzed the relevant Confluence content and identified key areas for improvement. The system has processed multiple pages and extracted actionable insights.

## Key Recommendations
1. **Immediate Actions**: Update documentation structure for better navigation
2. **Process Improvements**: Implement automated content review workflows  
3. **Long-term Strategy**: Establish content governance guidelines

## Next Steps
- Review the detailed reasoning in the "Reasoning Steps" tab
- Check which tools were used in the "Used Tools" tab
- Ask follow-up questions for clarification or refinement

*Analysis completed at ${new Date().toLocaleString()}*`;
  };

  const generateReasoningSteps = () => {
    return `## Step-by-Step Reasoning

### 1. Context Retrieval
- Searched across Engineering, Product, and Documentation spaces
- Identified 3 relevant pages containing goal-related information
- Extracted key themes and patterns from content

### 2. Content Analysis
- Summarized main points from each source
- Identified gaps and inconsistencies
- Analyzed current state vs desired outcomes

### 3. Recommendation Generation
- Applied best practices from similar scenarios
- Considered organizational constraints and capabilities
- Prioritized recommendations by impact and feasibility

### Decision Factors
- **Relevance**: How closely content matched the stated goal
- **Completeness**: Coverage of all aspects mentioned in the goal
- **Actionability**: Practical steps that can be implemented`;
  };

  const generateUsedTools = () => {
    return `## Tools Utilized in This Analysis

### 🔍 AI Powered Search
- **Purpose**: Retrieved relevant content from Confluence spaces
- **Scope**: Searched across 3 spaces, analyzed 5 pages
- **Results**: Found key documentation and process information

### 📊 Content Analyzer
- **Purpose**: Processed and summarized retrieved content
- **Method**: Natural language processing and pattern recognition
- **Output**: Structured insights and key themes

### 💡 Recommendation Engine
- **Purpose**: Generated actionable recommendations
- **Approach**: Best practice matching and gap analysis
- **Deliverable**: Prioritized action items with implementation guidance

### Integration Points
All tools worked together seamlessly to provide a comprehensive analysis of your goal.`;
  };

  const handleFollowUp = async () => {
    if (!followUpQuestion.trim() || !selectedSpace || selectedPages.length === 0) return;
    try {
      const searchResult = await apiService.search({
        space_key: selectedSpace,
        page_titles: selectedPages,
        query: followUpQuestion,
      });
      const qaContent = outputTabs.find(tab => tab.id === 'qa')?.content || '';
      const updatedQA = `${qaContent}\n\n**Q: ${followUpQuestion}**\n\nA: ${searchResult.response}`;
      setOutputTabs(prev => prev.map(tab =>
        tab.id === 'qa' ? { ...tab, content: updatedQA } : tab
      ));
      setFollowUpQuestion('');
    } catch (err) {
      setError('Failed to get follow-up answer.');
    }
  };

  const exportPlan = () => {
    const content = `# AI Agent Analysis Report

## Goal
${goal}

## Execution Plan
${planSteps.map(step => `${step.id}. ${step.title} - ${step.status}`).join('\n')}

## Final Answer
${outputTabs.find(tab => tab.id === 'answer')?.content || ''}

## Reasoning Steps
${outputTabs.find(tab => tab.id === 'reasoning')?.content || ''}

## Tools Used
${outputTabs.find(tab => tab.id === 'used-tools')?.content || ''}

---
*Generated by Confluence AI Assistant - Agent Mode*
*Date: ${new Date().toLocaleString()}*`;

    const blob = new Blob([content], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'ai-agent-analysis.md';
    a.click();
  };

  const replaySteps = () => {
    setPlanSteps([]);
    setCurrentStep(0);
    setOutputTabs([]);
    setShowFollowUp(false);
    setActiveTab('answer');
    handleGoalSubmit();
  };

  // Calculate progress percentage, clamp to 100% if complete
  const progressPercent = planSteps.length === 0 ? 0 :
    (currentStep + 1 >= planSteps.length ? 100 : Math.round(((currentStep + 1) / planSteps.length) * 100));

  return (
    <div className="fixed inset-0 bg-black/30 backdrop-blur-sm flex items-center justify-center z-40 p-4">
      <div className="bg-white/80 backdrop-blur-xl border border-white/20 rounded-2xl shadow-2xl w-full max-w-6xl max-h-[90vh] overflow-hidden">
        {/* Header */}
        <div className="bg-gradient-to-r from-orange-500/90 to-orange-600/90 backdrop-blur-xl p-6 text-white border-b border-orange-300/20">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <Zap className="w-8 h-8" />
              <div>
                <h2 className="text-2xl font-bold">Agent Mode</h2>
                <p className="text-orange-100/90">Goal-based AI assistance with planning and execution</p>
              </div>
            </div>
            <div className="flex items-center space-x-2">
              <button 
                onClick={() => onModeSelect('tool')}
                className="text-orange-100 hover:text-white hover:bg-white/10 rounded-lg px-3 py-1 text-sm transition-colors"
              >
                Switch to Tool Mode
              </button>
              <button onClick={onClose} className="text-white hover:bg-white/10 rounded-full p-2 backdrop-blur-sm">
                <X className="w-6 h-6" />
              </button>
            </div>
          </div>
        </div>

        <div className="p-6 overflow-y-auto max-h-[calc(90vh-200px)]">
          {/* Space and Page Selectors */}
          {!planSteps.length && !isPlanning && (
            <div className="max-w-4xl mx-auto mb-6">
              <div className="bg-white/60 backdrop-blur-xl rounded-xl p-6 border border-white/20 shadow-lg text-center">
                <h3 className="text-xl font-bold text-gray-800 mb-4">Select Space and Pages</h3>
                <div className="flex flex-col md:flex-row md:space-x-4 items-center justify-center mb-4">
                  <div className="mb-4 md:mb-0 w-full md:w-1/2">
                    <label className="block text-gray-700 mb-2 text-left">Space</label>
                    <select
                      className="w-full p-3 border border-orange-200/50 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500 bg-white/70 backdrop-blur-sm"
                      value={selectedSpace}
                      onChange={e => {
                        setSelectedSpace(e.target.value);
                        setSelectedPages([]);
                      }}
                    >
                      <option value="">Select a space...</option>
                      {spaces.map(space => (
                        <option key={space.key} value={space.key}>{space.name} ({space.key})</option>
                      ))}
                    </select>
                  </div>
                  <div className="w-full md:w-1/2">
                    <label className="block text-gray-700 mb-2 text-left">Pages</label>
                    <select
                      className="w-full p-3 border border-orange-200/50 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500 bg-white/70 backdrop-blur-sm"
                      multiple
                      size={Math.min(6, pages.length)}
                      value={selectedPages}
                      onChange={e => {
                        const options = Array.from(e.target.selectedOptions).map(opt => opt.value);
                        setSelectedPages(options);
                      }}
                      disabled={!selectedSpace}
                    >
                      {pages.map(page => (
                        <option key={page} value={page}>{page}</option>
                      ))}
                    </select>
                    <div className="text-xs text-gray-500 mt-1 text-left">Hold Ctrl (Windows) or Cmd (Mac) to select multiple pages.</div>
                  </div>
                </div>
                {error && <p className="text-red-500 text-sm mt-2">{error}</p>}
              </div>
            </div>
          )}

          {/* Goal Input Section */}
          {!planSteps.length && !isPlanning && (
            <div className="max-w-4xl mx-auto">
              <div className="bg-white/60 backdrop-blur-xl rounded-xl p-8 border border-white/20 shadow-lg text-center">
                <h3 className="text-2xl font-bold text-gray-800 mb-6">What do you want the assistant to help you achieve?</h3>
                <div className="relative">
                  <textarea
                    value={goal}
                    onChange={(e) => setGoal(e.target.value)}
                    placeholder="Describe your goal in detail... (e.g., 'Help me analyze our documentation structure and recommend improvements for better user experience')"
                    className="w-full p-4 border-2 border-orange-200/50 rounded-xl focus:ring-2 focus:ring-orange-500 focus:border-orange-500 resize-none bg-white/70 backdrop-blur-sm text-lg"
                    rows={4}
                  />
                  <button
                    onClick={handleGoalSubmit}
                    disabled={!goal.trim() || !selectedSpace || selectedPages.length === 0}
                    className="absolute bottom-4 right-4 bg-orange-500/90 backdrop-blur-sm text-white p-3 rounded-lg hover:bg-orange-600 disabled:bg-gray-300 disabled:cursor-not-allowed flex items-center space-x-2 transition-colors border border-white/10"
                  >
                    <Send className="w-5 h-5" />
                  </button>
                </div>
                {error && <p className="text-red-500 text-sm mt-4">{error}</p>}
              </div>
            </div>
          )}

          {/* Planning Phase */}
          {isPlanning && (
            <div className="max-w-4xl mx-auto">
              <div className="bg-white/60 backdrop-blur-xl rounded-xl p-8 border border-white/20 shadow-lg text-center">
                <div className="flex items-center justify-center space-x-3 mb-4">
                  <Brain className="w-8 h-8 text-orange-500 animate-pulse" />
                  <h3 className="text-xl font-bold text-gray-800">Planning steps...</h3>
                </div>
                <div className="flex items-center justify-center space-x-4 text-gray-600">
                  <span>1. Retrieve context</span>
                  <span>→</span>
                  <span>2. Summarize</span>
                  <span>→</span>
                  <span>3. Recommend changes</span>
                </div>
              </div>
            </div>
          )}

          {/* Execution Phase */}
          {planSteps.length > 0 && (
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              {/* Left Column - Progress Timeline */}
              <div className="lg:col-span-1">
                <div className="bg-white/60 backdrop-blur-xl rounded-xl p-4 border border-white/20 shadow-lg">
                  <h3 className="font-semibold text-gray-800 mb-4">Live Progress Log</h3>
                  <div className="space-y-4">
                    {planSteps.map((step, index) => (
                      <div key={step.id} className="flex items-start space-x-3">
                        <div className="flex-shrink-0 mt-1">
                          {step.status === 'completed' ? (
                            <CheckCircle className="w-5 h-5 text-green-500" />
                          ) : step.status === 'running' ? (
                            <Loader2 className="w-5 h-5 text-orange-500 animate-spin" />
                          ) : (
                            <div className="w-5 h-5 border-2 border-gray-300 rounded-full" />
                          )}
                        </div>
                        <div className="flex-1">
                          <div className="font-medium text-gray-800">{step.title}</div>
                          {step.details && (
                            <div className="text-sm text-gray-600 mt-1">{step.details}</div>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                  
                  {/* Progress Bar */}
                  <div className="mt-6">
                    <div className="flex items-center justify-between text-sm text-gray-600 mb-2">
                      <span>Progress</span>
                      <span>{progressPercent}%</span>
                    </div>
                    <div className="w-full bg-gray-200 rounded-full h-2">
                      <div 
                        className="bg-orange-500 h-2 rounded-full transition-all duration-500"
                        style={{ width: `${progressPercent}%` }}
                      />
                    </div>
                  </div>
                </div>
              </div>

              {/* Right Columns - Output Tabs */}
              <div className="lg:col-span-2">
                {outputTabs.length > 0 && (
                  <div className="bg-white/60 backdrop-blur-xl rounded-xl border border-white/20 shadow-lg overflow-hidden">
                    {/* Tab Headers */}
                    <div className="border-b border-white/20 bg-white/40 backdrop-blur-sm">
                      <div className="flex overflow-x-auto">
                        {outputTabs.map(tab => {
                          const Icon = tab.icon;
                          return (
                            <button
                              key={tab.id}
                              onClick={() => setActiveTab(tab.id)}
                              className={`flex items-center space-x-2 px-4 py-3 border-b-2 transition-colors whitespace-nowrap ${
                                activeTab === tab.id
                                  ? 'border-orange-500 text-orange-600 bg-white/50'
                                  : 'border-transparent text-gray-600 hover:text-gray-800 hover:bg-white/30'
                              }`}
                            >
                              <Icon className="w-4 h-4" />
                              <span className="text-sm font-medium">{tab.label}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Tab Content */}
                    <div className="p-6">
                      {outputTabs.find(tab => tab.id === activeTab) && (
                        <div className="prose prose-sm max-w-none">
                          {activeTab === 'qa' ? (
                            <div>
                              <div className="whitespace-pre-wrap text-gray-700 mb-4">
                                {outputTabs.find(tab => tab.id === activeTab)?.content}
                              </div>
                              {showFollowUp && (
                                <div className="border-t border-white/20 pt-4">
                                  <div className="flex space-x-2">
                                    <input
                                      type="text"
                                      value={followUpQuestion}
                                      onChange={(e) => setFollowUpQuestion(e.target.value)}
                                      placeholder="Ask a follow-up question..."
                                      className="flex-1 p-3 border border-white/30 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500 bg-white/70 backdrop-blur-sm"
                                      onKeyPress={(e) => e.key === 'Enter' && handleFollowUp()}
                                    />
                                    <button
                                      onClick={handleFollowUp}
                                      disabled={!followUpQuestion.trim() || !selectedSpace || selectedPages.length === 0}
                                      className="px-4 py-3 bg-orange-500/90 backdrop-blur-sm text-white rounded-lg hover:bg-orange-600 disabled:bg-gray-300 transition-colors flex items-center border border-white/10"
                                    >
                                      <Plus className="w-4 h-4" />
                                    </button>
                                  </div>
                                </div>
                              )}
                            </div>
                          ) : (
                            <div className="whitespace-pre-wrap text-gray-700">
                              {outputTabs.find(tab => tab.id === activeTab)?.content.split('\n').map((line, index) => {
                                if (line.startsWith('### ')) {
                                  return <h3 key={index} className="text-lg font-bold text-gray-800 mt-4 mb-2">{line.substring(4)}</h3>;
                                } else if (line.startsWith('## ')) {
                                  return <h2 key={index} className="text-xl font-bold text-gray-800 mt-6 mb-3">{line.substring(3)}</h2>;
                                } else if (line.startsWith('# ')) {
                                  return <h1 key={index} className="text-2xl font-bold text-gray-800 mt-8 mb-4">{line.substring(2)}</h1>;
                                } else if (line.startsWith('- **')) {
                                  const match = line.match(/- \*\*(.*?)\*\*: (.*)/);
                                  if (match) {
                                    return <p key={index} className="mb-2"><strong>{match[1]}:</strong> {match[2]}</p>;
                                  }
                                } else if (line.startsWith('- ')) {
                                  return <p key={index} className="mb-1 ml-4"> 2 {line.substring(2)}</p>;
                                } else if (line.trim()) {
                                  return <p key={index} className="mb-2 text-gray-700">{line}</p>;
                                }
                                return <br key={index} />;
                              })}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Actions */}
          {planSteps.length > 0 && !isPlanning && !isExecuting && (
            <div className="flex justify-end mt-8 space-x-4">
              <button
                onClick={exportPlan}
                className="px-6 py-3 bg-orange-500/90 text-white rounded-lg hover:bg-orange-600 transition-colors font-semibold shadow-md border border-white/10"
              >
                <Download className="w-5 h-5 inline-block mr-2" />
                Export Plan
              </button>
              <button
                onClick={replaySteps}
                className="px-6 py-3 bg-white/80 text-orange-600 rounded-lg hover:bg-orange-100 transition-colors font-semibold shadow-md border border-orange-200/50"
              >
                <RotateCcw className="w-5 h-5 inline-block mr-2" />
                Replay Steps
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default AgentMode; 