import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import {
  FileText,
  ZoomIn,
  ZoomOut,
  Maximize2,
  Minimize2,
  RotateCcw,
  Download,
  Printer,
  Copy,
  Check,
  Info,
  List,
  Search,
  ChevronUp,
  ChevronDown,
  X,
  Sun,
  Moon,
  Coffee,
  BookOpen,
  AlignLeft,
  AlertCircle,
  Loader2,
  ExternalLink,
  Clock,
  Type,
  Layers,
  FileCheck,
} from './icons';
import {
  parseWordDocument,
  type WordParsedDocument,
  type WordDocElement,
  type WordParagraphElement,
  type WordRun,
} from './word/docxParser';
import './styles.css';

export type WordViewerDictionary = typeof dictionaryEn;

export interface WordViewerProps {
  url: string;
  title?: string;
  className?: string;
  onClose?: () => void;
  isRtl?: boolean;
  locale?: 'en' | 'ar';
  dictionary?: Partial<WordViewerDictionary>;
  transformUrl?: (url: string) => string;
}

export type CustomWordViewerProps = WordViewerProps;

type ViewMode = 'continuous' | 'paged';
type ThemeMode = 'light' | 'dark' | 'sepia';

const dictionaryAr = {
  loading: 'جاري تحميل وقراءة مستند Word...',
  errorTitle: 'تعذر تحميل المستند',
  errorDesc: 'حدث خطأ أثناء قراءة ملف Word. يمكنك تحميل الملف لفتحه محلياً.',
  retry: 'إعادة المحاولة',
  download: 'تحميل الملف',
  print: 'طباعة المستند',
  copyText: 'نسخ النص',
  copied: 'تم النسخ!',
  docInfo: 'معلومات المستند',
  outline: 'فهرس المحتويات',
  noHeadings: 'لا توجد عناوين في هذا المستند',
  searchPlaceholder: 'بحث في المستند...',
  noSearchResults: 'لا توجد نتائج',
  matchCount: 'من',
  continuousView: 'عرض متواصل',
  pagedView: 'عرض الصفحات',
  lightTheme: 'الوضع الفاتح',
  darkTheme: 'الوضع الداكن',
  sepiaTheme: 'وضع القراءة الدافئ',
  zoomIn: 'تكبير',
  zoomOut: 'تصغير',
  zoomReset: 'إعادة ضبط الحجم',
  fullscreen: 'ملء الشاشة',
  exitFullscreen: 'إلغاء ملء الشاشة',
  page: 'صفحة',
  of: 'من',
  words: 'كلمة',
  characters: 'حرف',
  paragraphs: 'فقرة',
  pagesCount: 'صفحات',
  readingTime: 'وقت القراءة التقديري',
  minutes: 'دقيقة',
  fileSize: 'حجم الملف',
  author: 'المؤلف',
  created: 'تاريخ الإنشاء',
  modified: 'آخر تعديل',
  legacyDocNotice: 'مستند بتنسيق Word 97-2003 (قديم) - تم استخراج النصوص بنجاح',
  close: 'إغلاق',
};

const dictionaryEn = {
  loading: 'Loading and parsing Word document...',
  errorTitle: 'Unable to load document',
  errorDesc: 'An error occurred while reading the Word file. You can download it to view locally.',
  retry: 'Retry',
  download: 'Download File',
  print: 'Print Document',
  copyText: 'Copy Text',
  copied: 'Copied!',
  docInfo: 'Document Information',
  outline: 'Table of Contents',
  noHeadings: 'No headings found in this document',
  searchPlaceholder: 'Search document...',
  noSearchResults: 'No results found',
  matchCount: 'of',
  continuousView: 'Continuous View',
  pagedView: 'Paged View',
  lightTheme: 'Light Mode',
  darkTheme: 'Dark Mode',
  sepiaTheme: 'Sepia Reading Mode',
  zoomIn: 'Zoom In',
  zoomOut: 'Zoom Out',
  zoomReset: 'Reset Zoom',
  fullscreen: 'Fullscreen',
  exitFullscreen: 'Exit Fullscreen',
  page: 'Page',
  of: 'of',
  words: 'words',
  characters: 'characters',
  paragraphs: 'paragraphs',
  pagesCount: 'pages',
  readingTime: 'Estimated reading time',
  minutes: 'min',
  fileSize: 'File Size',
  author: 'Author',
  created: 'Created Date',
  modified: 'Last Modified',
  legacyDocNotice: 'Legacy Word 97-2003 format (.doc) - Content extracted successfully',
  close: 'Close',
};

export const WordViewer: React.FC<WordViewerProps> = ({
  url,
  title = 'Word Document',
  className = '',
  onClose,
  isRtl: isRtlProp,
  locale = 'en',
  dictionary,
  transformUrl,
}) => {
  const isRtl =
    isRtlProp !== undefined
      ? isRtlProp
      : locale === 'ar' || (typeof document !== 'undefined' && document.documentElement.dir === 'rtl');
  const baseDict = isRtl ? dictionaryAr : dictionaryEn;
  const d = useMemo(() => ({ ...baseDict, ...dictionary }), [baseDict, dictionary]);

  const safeUrl = transformUrl ? transformUrl(url) : url;

  // States
  const [doc, setDoc] = useState<WordParsedDocument | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  // Viewport & Layout
  const [scale, setScale] = useState<number>(1);
  const [viewMode, setViewMode] = useState<ViewMode>('continuous');
  const [themeMode, setThemeMode] = useState<ThemeMode>('light');
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);

  // Drawers & Modals
  const [showOutline, setShowOutline] = useState<boolean>(false);
  const [showInfo, setShowInfo] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);

  // Search
  const [showSearch, setShowSearch] = useState<boolean>(false);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [currentMatchIndex, setCurrentMatchIndex] = useState<number>(0);

  // Refs
  const containerRef = useRef<HTMLDivElement>(null);
  const contentScrollRef = useRef<HTMLDivElement>(null);
  const searchInputRef = useRef<HTMLInputElement>(null);

  // Load and parse document
  const loadDocument = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const response = await fetch(safeUrl);
      if (!response.ok) {
        throw new Error(`Failed to fetch document: HTTP ${response.status}`);
      }
      const buffer = await response.arrayBuffer();
      const parsed = await parseWordDocument(buffer);
      setDoc(parsed);
    } catch (err: unknown) {
      console.error('Word viewer parsing error:', err);
      setError(err instanceof Error ? err.message : 'Unknown error parsing Word document');
    } finally {
      setIsLoading(false);
    }
  }, [safeUrl]);

  useEffect(() => {
    loadDocument();
  }, [loadDocument]);

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        setShowSearch(true);
        setTimeout(() => searchInputRef.current?.focus(), 50);
      } else if (e.key === 'Escape') {
        if (showSearch) setShowSearch(false);
        if (showInfo) setShowInfo(false);
        if (showOutline) setShowOutline(false);
      } else if ((e.ctrlKey || e.metaKey) && (e.key === '+' || e.key === '=')) {
        e.preventDefault();
        handleZoomIn();
      } else if ((e.ctrlKey || e.metaKey) && e.key === '-') {
        e.preventDefault();
        handleZoomOut();
      } else if ((e.ctrlKey || e.metaKey) && e.key === '0') {
        e.preventDefault();
        setScale(1);
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
        e.preventDefault();
        handlePrint();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [showSearch, showInfo, showOutline]);

  // Zoom controls
  const handleZoomIn = () => setScale((s) => Math.min(2.5, Math.round((s + 0.15) * 100) / 100));
  const handleZoomOut = () => setScale((s) => Math.max(0.4, Math.round((s - 0.15) * 100) / 100));
  const handleResetZoom = () => setScale(1);

  // Fullscreen
  const toggleFullscreen = () => {
    if (!containerRef.current) return;
    if (!document.fullscreenElement) {
      containerRef.current.requestFullscreen?.().then(() => setIsFullscreen(true)).catch(() => {});
    } else {
      document.exitFullscreen?.().then(() => setIsFullscreen(false)).catch(() => {});
    }
  };

  useEffect(() => {
    const onFsChange = () => setIsFullscreen(Boolean(document.fullscreenElement));
    document.addEventListener('fullscreenchange', onFsChange);
    return () => document.removeEventListener('fullscreenchange', onFsChange);
  }, []);

  // Copy full text
  const handleCopyText = async () => {
    if (!doc?.plainText) return;
    try {
      await navigator.clipboard.writeText(doc.plainText);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // ignore
    }
  };

  // Print
  const handlePrint = () => {
    window.print();
  };

  // Download
  const handleDownload = () => {
    const a = document.createElement('a');
    a.href = safeUrl;
    a.download = title.endsWith('.docx') || title.endsWith('.doc') ? title : `${title}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // Jump to heading in TOC
  const scrollToHeading = (id: string) => {
    setShowOutline(false);
    const targetEl = document.getElementById(id);
    if (targetEl) {
      targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
      targetEl.classList.add('bg-blue-500/20');
      setTimeout(() => targetEl.classList.remove('bg-blue-500/20'), 1500);
    }
  };

  // Search matches calculation
  const searchMatches = useMemo(() => {
    if (!searchQuery.trim() || !doc?.plainText) return [];
    const query = searchQuery.toLowerCase();
    const matches: { text: string; index: number }[] = [];
    let startIdx = 0;
    const lowerDoc = doc.plainText.toLowerCase();

    while (startIdx < lowerDoc.length) {
      const idx = lowerDoc.indexOf(query, startIdx);
      if (idx === -1) break;
      matches.push({ text: query, index: idx });
      startIdx = idx + query.length;
    }
    return matches;
  }, [searchQuery, doc]);

  const handleNextSearch = () => {
    if (searchMatches.length === 0) return;
    setCurrentMatchIndex((prev) => (prev + 1) % searchMatches.length);
    highlightActiveSearchMatch((currentMatchIndex + 1) % searchMatches.length);
  };

  const handlePrevSearch = () => {
    if (searchMatches.length === 0) return;
    setCurrentMatchIndex((prev) => (prev - 1 + searchMatches.length) % searchMatches.length);
    highlightActiveSearchMatch((currentMatchIndex - 1 + searchMatches.length) % searchMatches.length);
  };

  const highlightActiveSearchMatch = (idx: number) => {
    const highlights = contentScrollRef.current?.querySelectorAll('mark[data-word-search="true"]');
    if (highlights && highlights[idx]) {
      highlights[idx].scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };

  // Reading time estimate (200 words per minute)
  const readingTimeMinutes = useMemo(() => {
    if (!doc?.metadata.words) return 1;
    return Math.max(1, Math.ceil(doc.metadata.words / 200));
  }, [doc]);

  // Format file size
  const formattedFileSize = useMemo(() => {
    const bytes = doc?.metadata.fileSizeBytes;
    if (!bytes) return null;
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  }, [doc]);

  // Theme paper classes
  const paperThemeClasses = useMemo(() => {
    switch (themeMode) {
      case 'dark':
        return 'bg-[#182334] text-[#f1f5f9] border-[#293c54] shadow-2xl';
      case 'sepia':
        return 'bg-[#fcf8f0] text-[#332211] border-[#e2d5c3] shadow-xl';
      case 'light':
      default:
        return 'bg-white text-[#111827] border-[#e2e8f0] shadow-xl';
    }
  }, [themeMode]);

  return (
    <div
      ref={containerRef}
      dir={isRtl ? 'rtl' : 'ltr'}
      className={`relative flex flex-col w-full h-full select-text bg-[#0b121c] text-white overflow-hidden ${className}`}
    >
      {/* Printable styles */}
      <style>{`
        @media print {
          body * {
            visibility: hidden !important;
          }
          #custom-word-printable, #custom-word-printable * {
            visibility: visible !important;
          }
          #custom-word-printable {
            position: absolute !important;
            left: 0 !important;
            top: 0 !important;
            width: 100% !important;
            background: white !important;
            color: black !important;
            padding: 0 !important;
            margin: 0 !important;
            box-shadow: none !important;
          }
        }
      `}</style>

      {/* TOP TOOLBAR */}
      <div className="flex-none flex items-center justify-between px-3 py-2 bg-[#121b29] border-b border-[#233549] gap-2 z-20">
        {/* Left: Document Info & Badge */}
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="flex items-center justify-center w-8 h-8 rounded-lg bg-gradient-to-br from-[#185abd] to-[#103f82] text-white shadow-md flex-shrink-0">
            <FileText className="w-4 h-4" />
          </div>
          <div className="min-w-0 flex flex-col">
            <div className="flex items-center gap-2">
              <span className="text-sm font-semibold text-white truncate max-w-[200px] sm:max-w-[320px]" title={title}>
                {title}
              </span>
              <span className="text-[10px] uppercase font-bold tracking-wider px-1.5 py-0.5 rounded bg-[#185abd]/30 text-[#60a5fa] border border-[#185abd]/50">
                {doc?.metadata.isLegacyDoc ? 'DOC' : 'DOCX'}
              </span>
            </div>
            {doc?.metadata.pages && (
              <span className="text-[11px] text-[#8fa0b5]">
                {d.pagesCount}: {doc.metadata.pages} • {doc.metadata.words || 0} {d.words}
              </span>
            )}
          </div>
        </div>

        {/* Center: Search Bar or Zoom Controls */}
        <div className="flex items-center gap-1 sm:gap-2">
          {/* Search Toggle / Box */}
          {showSearch ? (
            <div className="flex items-center bg-[#1a2636] border border-[#304560] rounded-lg px-2 py-1 gap-1.5 animate-fadeIn">
              <Search className="w-3.5 h-3.5 text-[#8fa0b5]" />
              <input
                ref={searchInputRef}
                type="text"
                placeholder={d.searchPlaceholder}
                value={searchQuery}
                onChange={(e) => {
                  setSearchQuery(e.target.value);
                  setCurrentMatchIndex(0);
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    if (e.shiftKey) handlePrevSearch();
                    else handleNextSearch();
                  }
                }}
                className="bg-transparent border-none outline-none text-xs text-white placeholder-[#8fa0b5] w-28 sm:w-44"
              />
              {searchMatches.length > 0 && (
                <span className="text-[11px] text-[#8fa0b5] whitespace-nowrap">
                  {currentMatchIndex + 1} {d.matchCount} {searchMatches.length}
                </span>
              )}
              {searchQuery && searchMatches.length === 0 && (
                <span className="text-[10px] text-amber-400 whitespace-nowrap">{d.noSearchResults}</span>
              )}
              <button
                onClick={handlePrevSearch}
                disabled={searchMatches.length === 0}
                className="p-0.5 text-[#8fa0b5] hover:text-white disabled:opacity-30"
                title="Previous match (Shift+Enter)"
              >
                <ChevronUp className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={handleNextSearch}
                disabled={searchMatches.length === 0}
                className="p-0.5 text-[#8fa0b5] hover:text-white disabled:opacity-30"
                title="Next match (Enter)"
              >
                <ChevronDown className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => {
                  setShowSearch(false);
                  setSearchQuery('');
                }}
                className="p-0.5 text-[#8fa0b5] hover:text-white"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          ) : (
            <button
              onClick={() => {
                setShowSearch(true);
                setTimeout(() => searchInputRef.current?.focus(), 50);
              }}
              className="p-1.5 text-[#8fa0b5] hover:text-white hover:bg-[#1a2636] rounded-md transition"
              title={`${d.searchPlaceholder} (Ctrl+F)`}
            >
              <Search className="w-4 h-4" />
            </button>
          )}

          <div className="h-4 w-[1px] bg-[#233549] mx-1 hidden sm:block" />

          {/* Zoom controls */}
          <div className="flex items-center gap-1 bg-[#162131] border border-[#233549] rounded-lg p-0.5">
            <button
              onClick={handleZoomOut}
              className="p-1 text-[#8fa0b5] hover:text-white hover:bg-[#203046] rounded transition"
              title={d.zoomOut}
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span
              onClick={handleResetZoom}
              className="text-[11px] font-mono px-1.5 py-0.5 text-[#9ab0c8] cursor-pointer hover:text-white min-w-[42px] text-center"
              title={d.zoomReset}
            >
              {Math.round(scale * 100)}%
            </span>
            <button
              onClick={handleZoomIn}
              className="p-1 text-[#8fa0b5] hover:text-white hover:bg-[#203046] rounded transition"
              title={d.zoomIn}
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={handleResetZoom}
              className="p-1 text-[#8fa0b5] hover:text-white hover:bg-[#203046] rounded transition"
              title={d.zoomReset}
            >
              <RotateCcw className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="h-4 w-[1px] bg-[#233549] mx-1 hidden md:block" />

          {/* View Mode: Continuous vs Paged */}
          <div className="hidden md:flex items-center bg-[#162131] border border-[#233549] rounded-lg p-0.5">
            <button
              onClick={() => setViewMode('continuous')}
              className={`p-1 rounded transition ${
                viewMode === 'continuous' ? 'bg-[#185abd] text-white shadow' : 'text-[#8fa0b5] hover:text-white'
              }`}
              title={d.continuousView}
            >
              <AlignLeft className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => setViewMode('paged')}
              className={`p-1 rounded transition ${
                viewMode === 'paged' ? 'bg-[#185abd] text-white shadow' : 'text-[#8fa0b5] hover:text-white'
              }`}
              title={d.pagedView}
            >
              <BookOpen className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Theme selector (single toggle button) */}
          <button
            onClick={() =>
              setThemeMode((prev) => (prev === 'light' ? 'sepia' : prev === 'sepia' ? 'dark' : 'light'))
            }
            className="p-1.5 text-[#8fa0b5] hover:text-white bg-[#162131] border border-[#233549] hover:bg-[#203046] rounded-lg transition hidden sm:inline-flex items-center justify-center"
            title={
              themeMode === 'light'
                ? d.lightTheme
                : themeMode === 'sepia'
                ? d.sepiaTheme
                : d.darkTheme
            }
          >
            {themeMode === 'light' ? (
              <Sun className="w-4 h-4 text-amber-400" />
            ) : themeMode === 'sepia' ? (
              <Coffee className="w-4 h-4 text-[#d4a373]" />
            ) : (
              <Moon className="w-4 h-4 text-blue-300" />
            )}
          </button>
        </div>

        {/* Right Action Buttons */}
        <div className="flex items-center gap-1">
          {/* Outline / Headings */}
          <button
            onClick={() => setShowOutline(!showOutline)}
            className={`p-1.5 rounded-lg border transition relative ${
              showOutline
                ? 'bg-[#185abd] text-white border-[#185abd]'
                : 'text-[#8fa0b5] hover:text-white bg-[#162131] border-[#233549] hover:bg-[#203046]'
            }`}
            title={d.outline}
          >
            <List className="w-4 h-4" />
            {doc?.headings && doc.headings.length > 0 && (
              <span className="absolute -top-1 -right-1 flex h-3 w-3">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-blue-400 opacity-75" />
                <span className="relative inline-flex rounded-full h-3 w-3 bg-[#185abd]" />
              </span>
            )}
          </button>

          {/* Copy plain text */}
          <button
            onClick={handleCopyText}
            className="p-1.5 text-[#8fa0b5] hover:text-white bg-[#162131] border border-[#233549] hover:bg-[#203046] rounded-lg transition"
            title={copied ? d.copied : d.copyText}
          >
            {copied ? <Check className="w-4 h-4 text-emerald-400" /> : <Copy className="w-4 h-4" />}
          </button>

          {/* Print */}
          <button
            onClick={handlePrint}
            className="hidden sm:inline-flex p-1.5 text-[#8fa0b5] hover:text-white bg-[#162131] border border-[#233549] hover:bg-[#203046] rounded-lg transition"
            title={d.print}
          >
            <Printer className="w-4 h-4" />
          </button>

          {/* Info Modal Toggle */}
          <button
            onClick={() => setShowInfo(!showInfo)}
            className={`p-1.5 rounded-lg border transition ${
              showInfo
                ? 'bg-[#185abd] text-white border-[#185abd]'
                : 'text-[#8fa0b5] hover:text-white bg-[#162131] border-[#233549] hover:bg-[#203046]'
            }`}
            title={d.docInfo}
          >
            <Info className="w-4 h-4" />
          </button>

          {/* Download */}
          <button
            onClick={handleDownload}
            className="p-1.5 text-[#8fa0b5] hover:text-white bg-[#162131] border border-[#233549] hover:bg-[#203046] rounded-lg transition"
            title={d.download}
          >
            <Download className="w-4 h-4" />
          </button>

          {/* Fullscreen */}
          <button
            onClick={toggleFullscreen}
            className="hidden md:inline-flex p-1.5 text-[#8fa0b5] hover:text-white bg-[#162131] border border-[#233549] hover:bg-[#203046] rounded-lg transition"
            title={isFullscreen ? d.exitFullscreen : d.fullscreen}
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>

          {/* Close dialog if supplied */}
          {onClose && (
            <button
              onClick={onClose}
              className="p-1.5 text-[#8fa0b5] hover:text-white hover:bg-rose-500/20 hover:text-rose-400 rounded-lg transition ml-1"
              title={d.close}
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* Legacy .doc format notice banner */}
      {doc?.metadata.isLegacyDoc && (
        <div className="bg-amber-500/10 border-b border-amber-500/30 px-4 py-2 flex items-center justify-between text-amber-200 text-xs">
          <div className="flex items-center gap-2">
            <AlertCircle className="w-4 h-4 text-amber-400 flex-shrink-0" />
            <span>{d.legacyDocNotice}</span>
          </div>
          <button
            onClick={handleDownload}
            className="underline hover:text-white text-xs font-semibold ml-2"
          >
            {d.download}
          </button>
        </div>
      )}

      {/* MAIN BODY AREA WITH SIDEBAR AND DOCUMENT */}
      <div className="relative flex-1 flex overflow-hidden">
        {/* OUTLINE / TABLE OF CONTENTS SIDEBAR */}
        {showOutline && (
          <div
            className={`w-72 bg-[#121c2a] border-${
              isRtl ? 'l' : 'r'
            } border-[#233549] flex flex-col z-10 transition-all shadow-2xl flex-shrink-0`}
          >
            <div className="p-3 border-b border-[#233549] flex items-center justify-between">
              <div className="flex items-center gap-2 text-sm font-semibold text-white">
                <List className="w-4 h-4 text-[#60a5fa]" />
                <span>{d.outline}</span>
              </div>
              <button
                onClick={() => setShowOutline(false)}
                className="text-[#8fa0b5] hover:text-white p-1 rounded hover:bg-[#1f2e43]"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-2 space-y-1">
              {!doc?.headings || doc.headings.length === 0 ? (
                <div className="text-center text-xs text-[#6b7c93] py-8">{d.noHeadings}</div>
              ) : (
                doc.headings.map((h, i) => (
                  <button
                    key={`${h.id}-${i}`}
                    onClick={() => scrollToHeading(h.id)}
                    className="w-full text-start px-2 py-1.5 rounded text-xs transition hover:bg-[#1a293e] text-[#c2d3e7] hover:text-white truncate block"
                    style={{
                      paddingInlineStart: `${Math.max(8, (h.level - 1) * 16 + 8)}px`,
                      fontWeight: h.level === 1 ? '700' : h.level === 2 ? '600' : '400',
                    }}
                    title={h.text}
                  >
                    • {h.text}
                  </button>
                ))
              )}
            </div>

            {/* Document summary info badge */}
            <div className="p-3 border-t border-[#233549] bg-[#0e1622] text-[11px] text-[#8fa0b5] space-y-1">
              <div className="flex justify-between">
                <span>{d.words}:</span>
                <span className="text-white font-medium">{doc?.metadata.words?.toLocaleString() || 0}</span>
              </div>
              <div className="flex justify-between">
                <span>{d.readingTime}:</span>
                <span className="text-white font-medium">
                  ~{readingTimeMinutes} {d.minutes}
                </span>
              </div>
            </div>
          </div>
        )}

        {/* DOCUMENT VIEWPORT SCROLLER */}
        <div
          ref={contentScrollRef}
          className="flex-1 overflow-auto p-4 sm:p-8 flex flex-col items-center relative bg-[#090f17]"
        >
          {/* Loading state */}
          {isLoading && (
            <div className="flex flex-col items-center justify-center my-auto p-8 text-center">
              <div className="relative mb-4">
                <div className="w-16 h-16 rounded-2xl bg-gradient-to-tr from-[#185abd] to-[#2563eb] flex items-center justify-center shadow-lg shadow-blue-500/20">
                  <FileText className="w-8 h-8 text-white animate-pulse" />
                </div>
                <Loader2 className="w-6 h-6 text-white absolute -bottom-1 -right-1 animate-spin" />
              </div>
              <h3 className="text-base font-semibold text-white mb-1">{d.loading}</h3>
              <p className="text-xs text-[#8fa0b5] max-w-sm">{title}</p>
            </div>
          )}

          {/* Error state */}
          {!isLoading && error && (
            <div className="flex flex-col items-center justify-center my-auto p-8 max-w-md text-center bg-[#131d2b] border border-[#2c3e55] rounded-2xl shadow-xl">
              <div className="w-14 h-14 rounded-full bg-rose-500/10 text-rose-400 flex items-center justify-center mb-4">
                <AlertCircle className="w-7 h-7" />
              </div>
              <h3 className="text-lg font-bold text-white mb-2">{d.errorTitle}</h3>
              <p className="text-xs text-[#8fa0b5] mb-6">{error || d.errorDesc}</p>
              <div className="flex items-center gap-3">
                <button
                  onClick={loadDocument}
                  className="px-4 py-2 bg-[#185abd] hover:bg-[#2069d6] text-white text-xs font-semibold rounded-lg shadow transition"
                >
                  {d.retry}
                </button>
                <button
                  onClick={handleDownload}
                  className="px-4 py-2 bg-[#1e2c3e] hover:bg-[#283b52] text-white text-xs font-semibold rounded-lg border border-[#314660] transition"
                >
                  {d.download}
                </button>
              </div>
            </div>
          )}

          {/* Document Content */}
          {!isLoading && !error && doc && (
            <div
              id="custom-word-printable"
              style={{
                transform: `scale(${scale})`,
                transformOrigin: 'top center',
                transition: 'transform 0.15s ease-out',
              }}
              className="flex flex-col items-center w-full max-w-[850px]"
            >
              {viewMode === 'paged' ? (
                // Paged Mode
                doc.pages.map((page) => (
                  <div
                    key={`page-${page.pageNumber}`}
                    className={`w-full min-h-[1100px] mb-8 p-12 sm:p-16 rounded-xl border relative flex flex-col justify-between ${paperThemeClasses}`}
                  >
                    {/* Header */}
                    <div className="flex items-center justify-between pb-6 mb-6 border-b border-black/10 dark:border-white/10 text-[11px] opacity-60">
                      <span className="truncate max-w-[300px]">{title}</span>
                      <span>
                        {d.page} {page.pageNumber} {d.of} {doc.pages.length}
                      </span>
                    </div>

                    {/* Page Content */}
                    <div className="flex-1 space-y-3">
                      {page.elements.map((el) => (
                        <DocumentElementRenderer
                          key={el.id}
                          element={el}
                          searchQuery={searchQuery}
                          currentMatchIndex={currentMatchIndex}
                        />
                      ))}
                    </div>

                    {/* Footer */}
                    <div className="pt-6 mt-6 border-t border-black/10 dark:border-white/10 text-center text-[11px] opacity-50">
                      — {page.pageNumber} —
                    </div>
                  </div>
                ))
              ) : (
                // Continuous View
                <div
                  className={`w-full min-h-[1100px] p-8 sm:p-14 rounded-2xl border space-y-3 ${paperThemeClasses}`}
                >
                  {doc.allElements.map((el) => (
                    <DocumentElementRenderer
                      key={el.id}
                      element={el}
                      searchQuery={searchQuery}
                      currentMatchIndex={currentMatchIndex}
                    />
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* METADATA / DOCUMENT INFO MODAL */}
      {showInfo && doc && (
        <div className="absolute inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50 animate-fadeIn">
          <div className="bg-[#121c29] border border-[#2b415a] rounded-2xl w-full max-w-md p-6 shadow-2xl space-y-4">
            <div className="flex items-center justify-between border-b border-[#233549] pb-3">
              <div className="flex items-center gap-2">
                <FileCheck className="w-5 h-5 text-[#60a5fa]" />
                <h3 className="text-base font-bold text-white">{d.docInfo}</h3>
              </div>
              <button
                onClick={() => setShowInfo(false)}
                className="text-[#8fa0b5] hover:text-white p-1 rounded hover:bg-[#1e2d42]"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="grid grid-cols-2 gap-3 text-xs">
              <div className="bg-[#182537] p-3 rounded-xl border border-[#283d56]">
                <div className="flex items-center gap-1.5 text-[#8fa0b5] mb-1">
                  <Type className="w-3.5 h-3.5 text-blue-400" />
                  <span>{d.words}</span>
                </div>
                <div className="text-base font-bold text-white">
                  {doc.metadata.words?.toLocaleString() || 0}
                </div>
              </div>

              <div className="bg-[#182537] p-3 rounded-xl border border-[#283d56]">
                <div className="flex items-center gap-1.5 text-[#8fa0b5] mb-1">
                  <Layers className="w-3.5 h-3.5 text-emerald-400" />
                  <span>{d.characters}</span>
                </div>
                <div className="text-base font-bold text-white">
                  {doc.metadata.characters?.toLocaleString() || 0}
                </div>
              </div>

              <div className="bg-[#182537] p-3 rounded-xl border border-[#283d56]">
                <div className="flex items-center gap-1.5 text-[#8fa0b5] mb-1">
                  <Clock className="w-3.5 h-3.5 text-amber-400" />
                  <span>{d.readingTime}</span>
                </div>
                <div className="text-base font-bold text-white">
                  ~{readingTimeMinutes} {d.minutes}
                </div>
              </div>

              <div className="bg-[#182537] p-3 rounded-xl border border-[#283d56]">
                <div className="flex items-center gap-1.5 text-[#8fa0b5] mb-1">
                  <BookOpen className="w-3.5 h-3.5 text-purple-400" />
                  <span>{d.pagesCount}</span>
                </div>
                <div className="text-base font-bold text-white">{doc.metadata.pages || 1}</div>
              </div>
            </div>

            <div className="space-y-2 text-xs text-[#8fa0b5] pt-2 border-t border-[#233549]">
              {doc.metadata.title && (
                <div className="flex justify-between">
                  <span>العنوان:</span>
                  <span className="text-white font-medium text-end">{doc.metadata.title}</span>
                </div>
              )}
              {doc.metadata.creator && (
                <div className="flex justify-between">
                  <span>{d.author}:</span>
                  <span className="text-white font-medium">{doc.metadata.creator}</span>
                </div>
              )}
              {formattedFileSize && (
                <div className="flex justify-between">
                  <span>{d.fileSize}:</span>
                  <span className="text-white font-medium">{formattedFileSize}</span>
                </div>
              )}
              {doc.metadata.created && (
                <div className="flex justify-between">
                  <span>{d.created}:</span>
                  <span className="text-white font-medium">
                    {new Date(doc.metadata.created).toLocaleDateString()}
                  </span>
                </div>
              )}
              {doc.metadata.modified && (
                <div className="flex justify-between">
                  <span>{d.modified}:</span>
                  <span className="text-white font-medium">
                    {new Date(doc.metadata.modified).toLocaleDateString()}
                  </span>
                </div>
              )}
            </div>

            <div className="pt-2 flex justify-end">
              <button
                onClick={() => setShowInfo(false)}
                className="px-4 py-2 bg-[#185abd] hover:bg-[#2069d6] text-white text-xs font-semibold rounded-lg transition"
              >
                {d.close}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

/**
 * Component to render individual document elements (Paragraphs, Tables, Images)
 */
const DocumentElementRenderer: React.FC<{
  element: WordDocElement;
  searchQuery: string;
  currentMatchIndex: number;
}> = ({ element, searchQuery }) => {
  if (element.type === 'pageBreak') {
    return <hr className="my-6 border-t-2 border-dashed border-gray-300 dark:border-gray-700" />;
  }

  if (element.type === 'image') {
    const alignClass =
      element.alignment === 'center'
        ? 'mx-auto'
        : element.alignment === 'right'
        ? 'ml-auto'
        : 'mr-auto';
    return (
      <div className={`my-4 flex ${element.alignment === 'center' ? 'justify-center' : ''}`}>
        <img
          src={element.src}
          alt={element.alt || 'Document image'}
          style={{
            maxWidth: element.widthPx ? `${element.widthPx}px` : '100%',
            height: element.heightPx ? `${element.heightPx}px` : 'auto',
          }}
          className={`rounded-lg shadow-md max-w-full ${alignClass}`}
          loading="lazy"
        />
      </div>
    );
  }

  if (element.type === 'table') {
    return (
      <div className="my-4 overflow-x-auto">
        <table className="w-full border-collapse border border-gray-300 dark:border-gray-600 text-xs sm:text-sm">
          <tbody>
            {element.rows.map((row) => (
              <tr
                key={row.id}
                className={row.isHeader ? 'bg-gray-100 dark:bg-gray-800 font-semibold' : ''}
              >
                {row.cells.map((cell) => (
                  <td
                    key={cell.id}
                    colSpan={cell.colspan || 1}
                    rowSpan={cell.rowspan || 1}
                    style={{
                      backgroundColor: cell.backgroundColor || undefined,
                      borderColor: cell.borderColor || undefined,
                    }}
                    className="border border-gray-300 dark:border-gray-600 p-2 sm:p-2.5 align-top"
                  >
                    {cell.elements.map((subEl) => (
                      <DocumentElementRenderer
                        key={subEl.id}
                        element={subEl}
                        searchQuery={searchQuery}
                        currentMatchIndex={0}
                      />
                    ))}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  if (element.type === 'paragraph') {
    return <ParagraphRenderer p={element} searchQuery={searchQuery} />;
  }

  return null;
};

/**
 * Paragraph & Runs Renderer with Typography and Search Highlights
 */
const ParagraphRenderer: React.FC<{
  p: WordParagraphElement;
  searchQuery: string;
}> = ({ p, searchQuery }) => {
  const alignClass =
    p.alignment === 'center'
      ? 'text-center'
      : p.alignment === 'right'
      ? 'text-right'
      : p.alignment === 'justify'
      ? 'text-justify'
      : 'text-left';

  // Heading styles
  let headingClass = '';
  if (p.headingLevel === 1) {
    headingClass = 'text-2xl sm:text-3xl font-bold mt-6 mb-3 text-blue-600 dark:text-blue-400 border-b pb-2';
  } else if (p.headingLevel === 2) {
    headingClass = 'text-xl sm:text-2xl font-bold mt-5 mb-2.5 text-blue-500 dark:text-blue-300';
  } else if (p.headingLevel === 3) {
    headingClass = 'text-lg sm:text-xl font-semibold mt-4 mb-2';
  } else if (p.headingLevel && p.headingLevel >= 4) {
    headingClass = 'text-base font-semibold mt-3 mb-1.5';
  } else if (p.isTitle) {
    headingClass = 'text-3xl sm:text-4xl font-extrabold text-center my-6 text-blue-700 dark:text-blue-300';
  } else if (p.isSubtitle) {
    headingClass = 'text-lg sm:text-xl text-center text-gray-500 mb-6 italic';
  }

  // List bullet or numbering
  const listIndent = p.isList ? `${Math.max(1, (p.listLevel || 0) + 1) * 1.25}rem` : undefined;

  const style: React.CSSProperties = {
    paddingInlineStart: listIndent,
    marginTop: p.spacingBeforePt ? `${p.spacingBeforePt}pt` : undefined,
    marginBottom: p.spacingAfterPt ? `${p.spacingAfterPt}pt` : undefined,
    lineHeight: p.lineSpacing ? p.lineSpacing : 1.6,
  };

  return (
    <div
      id={p.id}
      dir={p.isRtl ? 'rtl' : 'ltr'}
      style={style}
      className={`relative leading-relaxed ${alignClass} ${headingClass} transition-colors duration-300`}
    >
      {p.isList && (
        <span className="inline-block mx-2 font-bold select-none text-blue-500">
          {p.listType === 'number' ? '•' : '▪'}
        </span>
      )}

      {p.runs.map((run, idx) => (
        <RunRenderer key={`${p.id}-r-${idx}`} run={run} searchQuery={searchQuery} />
      ))}
    </div>
  );
};

/**
 * Text Run Renderer
 */
const RunRenderer: React.FC<{
  run: WordRun;
  searchQuery: string;
}> = ({ run, searchQuery }) => {
  const style: React.CSSProperties = {
    fontWeight: run.bold ? 'bold' : 'normal',
    fontStyle: run.italic ? 'italic' : 'normal',
    textDecoration: [run.underline ? 'underline' : '', run.strike ? 'line-through' : '']
      .filter(Boolean)
      .join(' ') || undefined,
    color: run.color || undefined,
    backgroundColor: run.highlight ? getHighlightColor(run.highlight) : undefined,
    fontSize: run.fontSizePt ? `${run.fontSizePt}pt` : undefined,
    fontFamily: run.fontFamily || undefined,
    verticalAlign: run.superscript ? 'super' : run.subscript ? 'sub' : undefined,
  };

  let content: React.ReactNode = run.text;

  // Search highlighting
  if (searchQuery.trim().length > 0 && run.text.toLowerCase().includes(searchQuery.toLowerCase())) {
    const parts = run.text.split(new RegExp(`(${escapeRegExp(searchQuery)})`, 'gi'));
    content = parts.map((part, i) =>
      part.toLowerCase() === searchQuery.toLowerCase() ? (
        <mark
          key={i}
          data-word-search="true"
          className="bg-amber-300 text-black px-0.5 rounded shadow-sm font-semibold"
        >
          {part}
        </mark>
      ) : (
        part
      )
    );
  }

  if (run.href) {
    return (
      <a
        href={run.href}
        target="_blank"
        rel="noopener noreferrer"
        style={style}
        className="text-blue-500 hover:text-blue-700 underline inline-flex items-center gap-0.5 cursor-pointer"
      >
        <span>{content}</span>
        <ExternalLink className="w-3 h-3 inline-block opacity-70" />
      </a>
    );
  }

  return (
    <span style={style} className="whitespace-pre-wrap">
      {content}
    </span>
  );
};

function getHighlightColor(hl: string): string {
  const map: Record<string, string> = {
    yellow: '#ffeb3b',
    green: '#a5d6a7',
    cyan: '#80deea',
    magenta: '#f48fb1',
    blue: '#90caf9',
    red: '#ef9a9a',
    darkBlue: '#1565c0',
    darkCyan: '#00838f',
    darkGreen: '#2e7d32',
    darkMagenta: '#6a1b9a',
    darkRed: '#c62828',
    darkYellow: '#fbc02d',
    lightGray: '#e0e0e0',
  };
  return map[hl.toLowerCase()] || hl;
}

function escapeRegExp(string: string): string {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export const CustomWordViewer = WordViewer;
export default WordViewer;
