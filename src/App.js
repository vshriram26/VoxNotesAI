import { useCallback, useEffect, useRef, useState } from 'react';
import { ChevronLeft, ChevronRight, Download, Eraser, FileText, FolderOpen, Home, Mic, MicOff, NotebookPen, Plus, Redo2, Save, Search, Star, Trash2, Undo2 } from 'lucide-react';
import { PDFDocument, rgb } from 'pdf-lib';
import { GlobalWorkerOptions, getDocument, version as pdfjsVersion } from 'pdfjs-dist/legacy/build/pdf';
import { getStoredAnnotations, listStoredAnnotations, saveStoredAnnotations } from './annotationStorage';
import './App.css';

GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsVersion}/legacy/build/pdf.worker.min.mjs`;

const NOTEBOOK_DOCUMENT_ID = 'notebook:quick-notes';
const NOTEBOOK_DOCUMENT_NAME = 'Quick Notes';
const LIBRARY_TILES = [
  { key: 'all-notes', title: 'All Notes', icon: NotebookPen, accentClass: 'warm', badge: '3' },
  { key: 'starred', title: 'Starred', icon: Star, accentClass: 'sand' },
  { key: 'unfiled', title: 'Unfiled', icon: FolderOpen, accentClass: 'mint' },
  { key: 'trash', title: 'Trash', icon: Trash2, accentClass: 'sage' },
];
const NOTEBOOK_PAGE_WIDTH = 860;
const NOTEBOOK_PAGE_HEIGHT = 1180;
const PAGE_RENDER_AHEAD = 2;
const DEFAULT_COLOR = '#1f2937';
const DEFAULT_TOOL = 'pen';
const TOOL_PRESETS = {
  pen: { width: 3, opacity: 1 },
  highlighter: { width: 14, opacity: 0.28 },
};
const COLOR_SWATCHS = ['#dc2626', '#f59e0b', '#eab308', '#14b8a6', '#3b82f6', '#1d4ed8', '#9333ea', '#1f2937'];
const QUICK_TOOL_COLORS = ['#dc2626', '#1f2937', '#eab308', '#3b82f6'];
const PEN_VARIANTS = [
  { id: 'fountain', label: 'Fountain', previewClass: 'fountain', widthScale: 1.45, opacity: 1 },
  { id: 'ballpoint', label: 'Ballpoint', previewClass: 'ballpoint', widthScale: 1, opacity: 1 },
  { id: 'fine', label: 'Fine', previewClass: 'fine', widthScale: 0.72, opacity: 0.96 },
  { id: 'brush', label: 'Brush', previewClass: 'brush', widthScale: 1.9, opacity: 0.92 },
];
const HIGHLIGHTER_VARIANTS = [
  { id: 'neon', label: 'Neon', previewClass: 'neon', widthScale: 1.15, opacity: 0.3 },
  { id: 'soft', label: 'Soft', previewClass: 'soft', widthScale: 0.95, opacity: 0.22 },
  { id: 'mint', label: 'Mint', previewClass: 'mint', widthScale: 1.05, opacity: 0.26 },
  { id: 'chisel', label: 'Chisel', previewClass: 'chisel', widthScale: 1.3, opacity: 0.32 },
];
const ERASER_SIZES = [12, 20, 32, 48];
const FAVORITE_SLOT_COUNT = 3;
const FAVORITES_STORAGE_KEYS = {
  pen: 'voxnotes:pen-favorites',
  highlighter: 'voxnotes:highlighter-favorites',
};
const TOOLBAR_POSITION_STORAGE_KEY = 'voxnotes:toolbar-position';
const TOOLBAR_COLLAPSED_STORAGE_KEY = 'voxnotes:toolbar-collapsed';
const TOOLBAR_EDGE_PADDING = 8;
const TOOLBAR_TOP_PADDING = 10;
const TOOLBAR_LEFT_OVERHANG = 10;
const DEFAULT_TOOLBAR_POSITION = { x: -10, y: 18 };
const VOICE_REMINDER_WORD_LIMIT = 6;

const clampNumber = (value, minimum, maximum) => {
  return Math.min(Math.max(value, minimum), maximum);
};

const getToolbarDragBounds = (frameBounds, toolbarBounds) => {
  if (!frameBounds || !toolbarBounds) {
    return {
      minimumX: DEFAULT_TOOLBAR_POSITION.x,
      maximumX: DEFAULT_TOOLBAR_POSITION.x,
      minimumY: DEFAULT_TOOLBAR_POSITION.y,
      maximumY: DEFAULT_TOOLBAR_POSITION.y,
    };
  }

  const maximumX = Math.max(
    -TOOLBAR_LEFT_OVERHANG,
    frameBounds.width - toolbarBounds.width - TOOLBAR_EDGE_PADDING,
  );
  const maximumY = Math.max(
    TOOLBAR_TOP_PADDING,
    frameBounds.height - toolbarBounds.height - TOOLBAR_EDGE_PADDING,
  );

  return {
    minimumX: -TOOLBAR_LEFT_OVERHANG,
    maximumX,
    minimumY: TOOLBAR_TOP_PADDING,
    maximumY,
  };
};

const clampToolbarPositionToBounds = (position, frameBounds, toolbarBounds) => {
  const bounds = getToolbarDragBounds(frameBounds, toolbarBounds);

  return {
    x: clampNumber(position.x, bounds.minimumX, bounds.maximumX),
    y: clampNumber(position.y, bounds.minimumY, bounds.maximumY),
  };
};

const getSpeechRecognitionConstructor = () => {
  if (typeof window === 'undefined') {
    return null;
  }

  return window.SpeechRecognition || window.webkitSpeechRecognition || null;
};

const sanitizeStoredReminders = (storedReminders) => {
  if (!Array.isArray(storedReminders)) {
    return [];
  }

  return storedReminders.reduce((nextReminders, reminder, index) => {
    if (!reminder || typeof reminder !== 'object' || typeof reminder.text !== 'string') {
      return nextReminders;
    }

    const text = reminder.text.trim();

    if (!text) {
      return nextReminders;
    }

    nextReminders.push({
      id: typeof reminder.id === 'string' ? reminder.id : `reminder-${index}`,
      text,
      createdAt: typeof reminder.createdAt === 'string' ? reminder.createdAt : new Date().toISOString(),
      pageNumber: Number.isFinite(reminder.pageNumber) ? reminder.pageNumber : 1,
    });

    return nextReminders;
  }, []);
};

const extractReminderText = (transcript) => {
  if (!transcript) {
    return '';
  }

  const tokens = transcript.split(/\s+/).map((token) => token.trim()).filter(Boolean);
  const reminderIndex = tokens.findIndex((token) => token.toLowerCase() === 'reminder');

  if (reminderIndex === -1) {
    return '';
  }

  const capturedWords = [];

  for (let index = reminderIndex + 1; index < tokens.length && capturedWords.length < VOICE_REMINDER_WORD_LIMIT; index += 1) {
    const cleanedToken = tokens[index].replace(/^[^\w\u0900-\u097F]+|[^\w\u0900-\u097F]+$/g, '');

    if (!cleanedToken) {
      continue;
    }

    const lowerToken = cleanedToken.toLowerCase();

    if (lowerToken === 'next' || lowerToken === 'highlight' || cleanedToken === 'पुढचा') {
      break;
    }

    capturedWords.push(cleanedToken);
  }

  return capturedWords.join(' ');
};

const sanitizeStoredVoiceNotes = (storedVoiceNotes) => {
  if (typeof storedVoiceNotes !== 'string') {
    return '';
  }

  return storedVoiceNotes.trim();
};

const appendDictatedText = (currentText, nextText) => {
  const trimmedNextText = nextText.trim();

  if (!trimmedNextText) {
    return currentText;
  }

  if (!currentText.trim()) {
    return trimmedNextText;
  }

  const separator = /[\s\n]$/.test(currentText) ? '' : ' ';
  return `${currentText}${separator}${trimmedNextText}`;
};

const createFavoriteSlots = () => Array.from({ length: FAVORITE_SLOT_COUNT }, () => null);

const readStoredFavorites = (toolName) => {
  if (typeof window === 'undefined') {
    return createFavoriteSlots();
  }

  try {
    const rawValue = window.localStorage.getItem(FAVORITES_STORAGE_KEYS[toolName]);

    if (!rawValue) {
      return createFavoriteSlots();
    }

    const parsedValue = JSON.parse(rawValue);

    if (!Array.isArray(parsedValue)) {
      return createFavoriteSlots();
    }

    return createFavoriteSlots().map((_, index) => {
      const item = parsedValue[index];

      if (!item || typeof item !== 'object') {
        return null;
      }

      return {
        variant: typeof item.variant === 'string' ? item.variant : '',
        color: typeof item.color === 'string' ? item.color : DEFAULT_COLOR,
        size: Number.isFinite(item.size) ? item.size : TOOL_PRESETS[toolName].width,
      };
    });
  } catch (error) {
    return createFavoriteSlots();
  }
};

const writeStoredFavorites = (toolName, favorites) => {
  if (typeof window === 'undefined') {
    return;
  }

  try {
    window.localStorage.setItem(FAVORITES_STORAGE_KEYS[toolName], JSON.stringify(favorites));
  } catch (error) {
    // Ignore favorite persistence failures and keep the session usable.
  }
};

const readStoredToolbarPosition = () => {
  if (typeof window === 'undefined') {
    return DEFAULT_TOOLBAR_POSITION;
  }

  try {
    const rawValue = window.localStorage.getItem(TOOLBAR_POSITION_STORAGE_KEY);

    if (!rawValue) {
      return DEFAULT_TOOLBAR_POSITION;
    }

    const parsedValue = JSON.parse(rawValue);

    if (!parsedValue || !Number.isFinite(parsedValue.x) || !Number.isFinite(parsedValue.y)) {
      return DEFAULT_TOOLBAR_POSITION;
    }

    return parsedValue;
  } catch (error) {
    return DEFAULT_TOOLBAR_POSITION;
  }
};

const readStoredToolbarCollapsed = () => {
  if (typeof window === 'undefined') {
    return false;
  }

  try {
    return window.localStorage.getItem(TOOLBAR_COLLAPSED_STORAGE_KEY) === 'true';
  } catch (error) {
    return false;
  }
};

const getPointPressureMultiplier = (stroke, point) => {
  if (!stroke?.pressureEnabled) {
    return 1;
  }

  const pressure = Number.isFinite(point?.pressure) ? point.pressure : 1;
  const minimum = stroke.tool === 'highlighter' ? 0.82 : 0.42;
  const maximum = stroke.tool === 'highlighter' ? 1.22 : 1.58;

  return clampNumber(0.38 + pressure * 0.92, minimum, maximum);
};

const ToolTipIllustration = ({ tool, variant, color }) => {
  const strokeColor = color || DEFAULT_COLOR;
  const accentColor = tool === 'highlighter' ? 'rgba(255, 255, 255, 0.55)' : 'rgba(255, 255, 255, 0.22)';

  if (tool === 'highlighter') {
    const highlighterMarkup = {
      neon: (
        <>
          <rect x="13" y="7" width="14" height="10" rx="4" fill={accentColor} />
          <rect x="10" y="16" width="20" height="15" rx="5" fill={strokeColor} />
          <path d="M14 31H26L24 37H16L14 31Z" fill={strokeColor} opacity="0.8" />
        </>
      ),
      soft: (
        <>
          <rect x="12" y="8" width="16" height="9" rx="4" fill={accentColor} />
          <rect x="11" y="16" width="18" height="14" rx="5" fill={strokeColor} opacity="0.92" />
          <path d="M14 30H26L23 36H17L14 30Z" fill={strokeColor} opacity="0.72" />
        </>
      ),
      mint: (
        <>
          <rect x="12" y="7" width="16" height="10" rx="4" fill={accentColor} />
          <rect x="9" y="16" width="22" height="14" rx="6" fill={strokeColor} />
          <path d="M13 31H27L25 37H15L13 31Z" fill={strokeColor} opacity="0.78" />
        </>
      ),
      chisel: (
        <>
          <rect x="11" y="7" width="18" height="9" rx="4" fill={accentColor} />
          <path d="M10 16H30V28C30 30.5 27.5 33 24.5 33H15.5C12.5 33 10 30.5 10 28V16Z" fill={strokeColor} />
          <path d="M15 33H27L21 38L15 33Z" fill={strokeColor} opacity="0.82" />
        </>
      ),
    };

    return (
      <svg viewBox="0 0 40 40" aria-hidden="true" focusable="false">
        {highlighterMarkup[variant] ?? highlighterMarkup.neon}
      </svg>
    );
  }

  const penMarkup = {
    fountain: (
      <>
        <path d="M20 5L28 13L23 35H17L12 13L20 5Z" fill={strokeColor} />
        <circle cx="20" cy="20" r="2.2" fill={accentColor} />
      </>
    ),
    ballpoint: (
      <>
        <rect x="16" y="6" width="8" height="22" rx="4" fill={strokeColor} />
        <path d="M17 28H23L20 36L17 28Z" fill={strokeColor} opacity="0.84" />
      </>
    ),
    fine: (
      <>
        <rect x="17.5" y="5" width="5" height="24" rx="2.5" fill={strokeColor} />
        <path d="M18 29H22L20 37L18 29Z" fill={strokeColor} opacity="0.82" />
      </>
    ),
    brush: (
      <>
        <path d="M13 10C13 7.8 14.8 6 17 6H23C25.2 6 27 7.8 27 10V18C27 22 24.2 25 20 25C15.8 25 13 22 13 18V10Z" fill={strokeColor} />
        <path d="M16 24H24L20 37L16 24Z" fill={strokeColor} opacity="0.88" />
      </>
    ),
  };

  return (
    <svg viewBox="0 0 40 40" aria-hidden="true" focusable="false">
      {penMarkup[variant] ?? penMarkup.ballpoint}
    </svg>
  );
};

const RailGlyph = ({ kind }) => {
  return <span className={`rail-glyph ${kind}`} aria-hidden="true" />;
};

const createNotebookName = () => {
  return `Notebook ${new Date().toLocaleDateString([], { month: 'short', day: 'numeric' })}`;
};

const createEmptyPageAnnotations = () => ({ strokes: [], redoStack: [] });

const cloneStroke = (stroke) => ({
  points: Array.isArray(stroke?.points)
    ? stroke.points.map((point) => ({
      x: point?.x ?? 0,
      y: point?.y ?? 0,
      pressure: Number.isFinite(point?.pressure) ? clampNumber(point.pressure, 0.05, 1) : 1,
    }))
    : [],
  tool: stroke?.tool ?? DEFAULT_TOOL,
  variant: stroke?.variant ?? 'ballpoint',
  color: stroke?.color ?? DEFAULT_COLOR,
  width: Number.isFinite(stroke?.width) ? stroke.width : TOOL_PRESETS.pen.width,
  opacity: Number.isFinite(stroke?.opacity) ? stroke.opacity : TOOL_PRESETS.pen.opacity,
  pressureEnabled: Boolean(stroke?.pressureEnabled),
});

const createNotebookPages = (count = 1) => {
  return Array.from({ length: count }, (_, index) => ({
    pageNumber: index + 1,
    width: NOTEBOOK_PAGE_WIDTH,
    height: NOTEBOOK_PAGE_HEIGHT,
    kind: 'blank',
  }));
};

const buildDocumentId = (name, byteArray) => {
  const signature = Array.from(byteArray.slice(0, 16))
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('');

  return `${name}:${byteArray.byteLength}:${signature}`;
};

const formatSavedTime = (value) => {
  if (!value) {
    return '';
  }

  const parsedValue = value instanceof Date ? value : new Date(value);

  if (Number.isNaN(parsedValue.getTime())) {
    return '';
  }

  return parsedValue.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
};

const buildExportName = (name, sourceType) => {
  if (sourceType === 'notebook') {
    return 'quick-notes.pdf';
  }

  if (!name || name === 'No document loaded') {
    return 'voxnotes-export.pdf';
  }

  if (name.toLowerCase().endsWith('.pdf')) {
    return `${name.slice(0, -4)}-annotated.pdf`;
  }

  return `${name}.pdf`;
};

const sanitizeStrokeList = (strokes) => {
  if (!Array.isArray(strokes)) {
    return [];
  }

  return strokes
    .map((stroke) => cloneStroke({
      ...stroke,
      points: Array.isArray(stroke?.points)
        ? stroke.points.filter((point) => Number.isFinite(point?.x) && Number.isFinite(point?.y))
        : [],
    }))
    .filter((stroke) => stroke.points.length > 0);
};

const sanitizeStoredAnnotations = (storedAnnotations) => {
  if (!storedAnnotations || typeof storedAnnotations !== 'object' || Array.isArray(storedAnnotations)) {
    return {};
  }

  return Object.entries(storedAnnotations).reduce((nextAnnotations, [pageNumber, pageAnnotations]) => {
    nextAnnotations[pageNumber] = {
      strokes: sanitizeStrokeList(pageAnnotations?.strokes),
      redoStack: sanitizeStrokeList(pageAnnotations?.redoStack),
    };

    return nextAnnotations;
  }, {});
};

const sanitizeStoredPages = (storedPages, fallbackPages) => {
  if (!Array.isArray(storedPages) || !storedPages.length) {
    return fallbackPages;
  }

  return storedPages.map((page, index) => ({
    pageNumber: index + 1,
    width: Number.isFinite(page?.width) ? page.width : NOTEBOOK_PAGE_WIDTH,
    height: Number.isFinite(page?.height) ? page.height : NOTEBOOK_PAGE_HEIGHT,
    kind: page?.kind === 'pdf' ? 'pdf' : 'blank',
  }));
};

const distanceToSegment = (point, startPoint, endPoint) => {
  const deltaX = endPoint.x - startPoint.x;
  const deltaY = endPoint.y - startPoint.y;

  if (deltaX === 0 && deltaY === 0) {
    return Math.hypot(point.x - startPoint.x, point.y - startPoint.y);
  }

  const projection = ((point.x - startPoint.x) * deltaX + (point.y - startPoint.y) * deltaY) / (deltaX ** 2 + deltaY ** 2);
  const clampedProjection = Math.max(0, Math.min(1, projection));
  const projectedX = startPoint.x + clampedProjection * deltaX;
  const projectedY = startPoint.y + clampedProjection * deltaY;

  return Math.hypot(point.x - projectedX, point.y - projectedY);
};

function App() {
  const viewerShellRef = useRef(null);
  const overlayFrameRef = useRef(null);
  const toolbarFloatRef = useRef(null);
  const fileInputRef = useRef(null);
  const pdfDocumentRef = useRef(null);
  const pdfCanvasRefs = useRef({});
  const inkCanvasRefs = useRef({});
  const pageStageRefs = useRef({});
  const pageLayoutMapRef = useRef({});
  const annotationsRef = useRef({});
  const remindersRef = useRef([]);
  const dictatedNotesRef = useRef('');
  const renderedPdfPagesRef = useRef({});
  const drawingStateRef = useRef({ isDrawing: false, pageNumber: null, stroke: null });
  const originalPdfBytesRef = useRef(null);
  const documentIdRef = useRef('');
  const previousInkToolRef = useRef(DEFAULT_TOOL);
  const toolbarDragStateRef = useRef({ active: false, startX: 0, startY: 0, originX: 0, originY: 0 });
  const speechRecognitionRef = useRef(null);
  const shouldKeepListeningRef = useRef(false);

  const [pages, setPages] = useState(createNotebookPages(1));
  const [currentView, setCurrentView] = useState('home');
  const [libraryItems, setLibraryItems] = useState([]);
  const [isLibraryLoading, setIsLibraryLoading] = useState(true);
  const [sourceType, setSourceType] = useState('notebook');
  const [isListening, setIsListening] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [documentName, setDocumentName] = useState(NOTEBOOK_DOCUMENT_NAME);
  const [statusMessage, setStatusMessage] = useState('Notebook ready. Start writing or upload a PDF.');
  const [activePageNumber, setActivePageNumber] = useState(1);
  const [pageJumpValue, setPageJumpValue] = useState('1');
  const [lastSavedAt, setLastSavedAt] = useState('');
  const [storageLabel, setStorageLabel] = useState('IndexedDB primary storage');
  const [renderWindowPageNumbers, setRenderWindowPageNumbers] = useState([1]);
  const [activeTool, setActiveTool] = useState(DEFAULT_TOOL);
  const [inkColor, setInkColor] = useState(DEFAULT_COLOR);
  const [toolPanel, setToolPanel] = useState(null);
  const [penVariant, setPenVariant] = useState('ballpoint');
  const [highlighterVariant, setHighlighterVariant] = useState('neon');
  const [penSize, setPenSize] = useState(TOOL_PRESETS.pen.width);
  const [highlighterSize, setHighlighterSize] = useState(TOOL_PRESETS.highlighter.width);
  const [eraserSize, setEraserSize] = useState(ERASER_SIZES[1]);
  const [penFavorites, setPenFavorites] = useState(() => readStoredFavorites('pen'));
  const [highlighterFavorites, setHighlighterFavorites] = useState(() => readStoredFavorites('highlighter'));
  const [reminders, setReminders] = useState([]);
  const [dictatedNotes, setDictatedNotes] = useState('');
  const [liveTranscript, setLiveTranscript] = useState('');
  const [toolbarPosition, setToolbarPosition] = useState(() => readStoredToolbarPosition());
  const [isToolbarCollapsed, setIsToolbarCollapsed] = useState(() => readStoredToolbarCollapsed());
  const [autoSelectPreviousTool, setAutoSelectPreviousTool] = useState(false);
  const [eraseEntireStroke, setEraseEntireStroke] = useState(true);
  const [eraseHighlighterOnly, setEraseHighlighterOnly] = useState(false);
  const [erasePenOnly, setErasePenOnly] = useState(false);
  const [annotationVersion, setAnnotationVersion] = useState(0);

  const pageCount = pages.length;
  const hasPdf = sourceType === 'pdf';
  const latestReminder = reminders[0] ?? null;
  const selectedPenPreset = PEN_VARIANTS.find((preset) => preset.id === penVariant) ?? PEN_VARIANTS[1];
  const selectedHighlighterPreset = HIGHLIGHTER_VARIANTS.find((preset) => preset.id === highlighterVariant) ?? HIGHLIGHTER_VARIANTS[0];

  const handleToolSelect = useCallback((toolName) => {
    if (toolName !== 'eraser') {
      previousInkToolRef.current = toolName;
    }

    setActiveTool(toolName);
    setToolPanel((currentPanel) => currentPanel === toolName ? null : toolName);
  }, []);

  const buildStrokeSegments = useCallback((stroke, shouldErasePoint) => {
    const nextSegments = [];
    let currentPoints = [];

    stroke.points.forEach((point) => {
      if (shouldErasePoint(point)) {
        if (currentPoints.length) {
          nextSegments.push(cloneStroke({ ...stroke, points: currentPoints }));
          currentPoints = [];
        }

        return;
      }

      currentPoints.push(point);
    });

    if (currentPoints.length) {
      nextSegments.push(cloneStroke({ ...stroke, points: currentPoints }));
    }

    return nextSegments;
  }, []);

  const applyFavorite = useCallback((toolName, favorite) => {
    if (!favorite) {
      return;
    }

    setInkColor(favorite.color);
    if (toolName !== 'eraser') {
      previousInkToolRef.current = toolName;
    }

    setActiveTool(toolName);
    setToolPanel(toolName);

    if (toolName === 'pen') {
      setPenVariant(favorite.variant);
      setPenSize(favorite.size);
      setStatusMessage(`Applied pen favorite ${favorite.variant}.`);
      return;
    }

    setHighlighterVariant(favorite.variant);
    setHighlighterSize(favorite.size);
    setStatusMessage(`Applied highlighter favorite ${favorite.variant}.`);
  }, []);

  const saveFavorite = useCallback((toolName, slotIndex) => {
    if (toolName === 'pen') {
      const nextFavorites = penFavorites.map((favorite, index) => {
        if (index !== slotIndex) {
          return favorite;
        }

        return { variant: penVariant, color: inkColor, size: penSize };
      });

      setPenFavorites(nextFavorites);
      writeStoredFavorites('pen', nextFavorites);
      setStatusMessage(`Saved pen favorite ${slotIndex + 1}.`);
      return;
    }

    const nextFavorites = highlighterFavorites.map((favorite, index) => {
      if (index !== slotIndex) {
        return favorite;
      }

      return { variant: highlighterVariant, color: inkColor, size: highlighterSize };
    });

    setHighlighterFavorites(nextFavorites);
    writeStoredFavorites('highlighter', nextFavorites);
    setStatusMessage(`Saved highlighter favorite ${slotIndex + 1}.`);
  }, [highlighterFavorites, highlighterSize, highlighterVariant, inkColor, penFavorites, penSize, penVariant]);

  const updateRenderWindow = useCallback((centerPageNumber) => {
    const startPage = Math.max(1, centerPageNumber - PAGE_RENDER_AHEAD);
    const endPage = Math.min(pageCount, centerPageNumber + PAGE_RENDER_AHEAD);
    const nextWindow = [];

    for (let pageNumber = startPage; pageNumber <= endPage; pageNumber += 1) {
      nextWindow.push(pageNumber);
    }

    setRenderWindowPageNumbers((currentWindow) => {
      if (currentWindow.length === nextWindow.length && currentWindow.every((value, index) => value === nextWindow[index])) {
        return currentWindow;
      }

      return nextWindow;
    });
  }, [pageCount]);

  const refreshLibrary = useCallback(async () => {
    setIsLibraryLoading(true);

    try {
      const records = await listStoredAnnotations();
      setLibraryItems(records.map((record, index) => ({
        documentId: record.documentId,
        documentName: record.documentName,
        sourceType: record.sourceType ?? 'notebook',
        updatedAt: record.updatedAt,
        pageCount: Array.isArray(record.pages) ? record.pages.length : 0,
        strokeCount: Object.values(record.annotations ?? {}).reduce((count, pageAnnotations) => {
          return count + (pageAnnotations?.strokes?.length ?? 0);
        }, 0),
        accentClass: ['violet', 'paper', 'linen', 'mint'][index % 4],
        sourceBytes: record.sourceBytes ?? null,
      })));
    } catch (error) {
      setLibraryItems([]);
    } finally {
      setIsLibraryLoading(false);
    }
  }, []);

  const persistSession = useCallback(async (
    nextAnnotations,
    nextPages = pages,
    nextSourceType = sourceType,
    nextDocumentName = documentName,
    nextDocumentId = documentIdRef.current,
    nextSourceBytes = originalPdfBytesRef.current,
    nextReminders = remindersRef.current,
    nextDictatedNotes = dictatedNotesRef.current,
  ) => {
    if (!nextDocumentId) {
      return;
    }

    const updatedAt = new Date().toISOString();

    try {
      await saveStoredAnnotations({
        documentId: nextDocumentId,
        documentName: nextDocumentName,
        sourceType: nextSourceType,
        updatedAt,
        pages: nextPages,
        annotations: nextAnnotations,
        sourceBytes: nextSourceType === 'pdf' && nextSourceBytes ? nextSourceBytes : null,
        reminders: nextReminders,
        voiceNotes: nextDictatedNotes,
      });
      setStorageLabel('IndexedDB primary storage');
      setLastSavedAt(formatSavedTime(updatedAt));
      refreshLibrary();
    } catch (error) {
      setStorageLabel('Unable to persist session');
      setStatusMessage('IndexedDB is unavailable in this browser session.');
    }
  }, [documentName, pages, refreshLibrary, sourceType]);

  const ensurePageAnnotations = useCallback((pageNumber) => {
    if (!annotationsRef.current[pageNumber]) {
      annotationsRef.current[pageNumber] = createEmptyPageAnnotations();
    }

    return annotationsRef.current[pageNumber];
  }, []);

  const getPageSize = useCallback((pageNumber) => {
    const pageLayout = pageLayoutMapRef.current[pageNumber];

    return {
      width: pageLayout?.width ?? NOTEBOOK_PAGE_WIDTH,
      height: pageLayout?.height ?? NOTEBOOK_PAGE_HEIGHT,
    };
  }, []);

  const clearCanvas = useCallback((canvas) => {
    if (!canvas) {
      return;
    }

    const context = canvas.getContext('2d');

    if (!context) {
      return;
    }

    context.save();
    context.setTransform(1, 0, 0, 1, 0, 0);
    context.clearRect(0, 0, canvas.width, canvas.height);
    context.restore();
  }, []);

  const configureBaseContext = useCallback((context, ratio) => {
    context.setTransform(ratio, 0, 0, ratio, 0, 0);
    context.lineCap = 'round';
    context.lineJoin = 'round';
  }, []);

  const applyStrokeStyle = useCallback((context, stroke) => {
    context.strokeStyle = stroke.color;
    context.fillStyle = stroke.color;
    context.globalAlpha = stroke.opacity;
  }, []);

  const drawStroke = useCallback((context, pageNumber, stroke) => {
    if (!context || !stroke?.points?.length) {
      return;
    }

    const { width, height } = getPageSize(pageNumber);

    applyStrokeStyle(context, stroke);

    if (stroke.points.length === 1) {
      const point = stroke.points[0];
      const pointWidth = stroke.width * getPointPressureMultiplier(stroke, point);
      context.beginPath();
      context.arc(point.x * width, point.y * height, pointWidth / 2, 0, Math.PI * 2);
      context.fill();
      context.closePath();
      context.globalAlpha = 1;
      return;
    }

    for (let index = 1; index < stroke.points.length; index += 1) {
      const startPoint = stroke.points[index - 1];
      const endPoint = stroke.points[index];
      const startX = startPoint.x * width;
      const startY = startPoint.y * height;
      const endX = endPoint.x * width;
      const endY = endPoint.y * height;
      const segmentDistance = Math.hypot(endX - startX, endY - startY);
      const startWidth = stroke.width * getPointPressureMultiplier(stroke, startPoint);
      const endWidth = stroke.width * getPointPressureMultiplier(stroke, endPoint);
      const segmentSteps = Math.max(1, Math.ceil(segmentDistance / 2));

      let previousX = startX;
      let previousY = startY;

      for (let step = 1; step <= segmentSteps; step += 1) {
        const progress = step / segmentSteps;
        const nextX = startX + (endX - startX) * progress;
        const nextY = startY + (endY - startY) * progress;

        context.lineWidth = startWidth + (endWidth - startWidth) * progress;
        context.beginPath();
        context.moveTo(previousX, previousY);
        context.lineTo(nextX, nextY);
        context.stroke();
        context.closePath();

        previousX = nextX;
        previousY = nextY;
      }
    }

    context.globalAlpha = 1;
  }, [applyStrokeStyle, getPageSize]);

  const redrawInkPage = useCallback((pageNumber) => {
    const canvas = inkCanvasRefs.current[pageNumber];
    const context = canvas?.getContext('2d');

    if (!canvas || !context) {
      return;
    }

    clearCanvas(canvas);

    const pageAnnotations = annotationsRef.current[pageNumber]?.strokes ?? [];

    pageAnnotations.forEach((stroke) => {
      drawStroke(context, pageNumber, stroke);
    });
  }, [clearCanvas, drawStroke]);

  const syncInkCanvas = useCallback((pageNumber, width, height) => {
    const canvas = inkCanvasRefs.current[pageNumber];

    if (!canvas) {
      return;
    }

    const ratio = window.devicePixelRatio || 1;
    const context = canvas.getContext('2d');

    if (!context) {
      return;
    }

    canvas.width = Math.floor(width * ratio);
    canvas.height = Math.floor(height * ratio);
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;

    configureBaseContext(context, ratio);
    redrawInkPage(pageNumber);
  }, [configureBaseContext, redrawInkPage]);

  const releaseCanvasResources = useCallback((canvas, width, height) => {
    if (!canvas) {
      return;
    }

    canvas.width = 1;
    canvas.height = 1;
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;
  }, []);

  const renderPdfPage = useCallback(async ({ pageNumber, width, height }) => {
    const pdf = pdfDocumentRef.current;
    const canvas = pdfCanvasRefs.current[pageNumber];

    if (!pdf || !canvas) {
      return;
    }

    const signature = `${width}x${height}`;

    if (renderedPdfPagesRef.current[pageNumber] === signature) {
      return;
    }

    const page = await pdf.getPage(pageNumber);
    const baseViewport = page.getViewport({ scale: 1 });
    const viewport = page.getViewport({ scale: width / baseViewport.width });
    const ratio = window.devicePixelRatio || 1;
    const context = canvas.getContext('2d');

    if (!context) {
      return;
    }

    canvas.width = Math.floor(viewport.width * ratio);
    canvas.height = Math.floor(viewport.height * ratio);
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;

    context.setTransform(ratio, 0, 0, ratio, 0, 0);
    context.clearRect(0, 0, viewport.width, viewport.height);

    await page.render({ canvasContext: context, viewport }).promise;
    renderedPdfPagesRef.current[pageNumber] = signature;
  }, []);

  const loadNotebookSession = useCallback(async ({ documentId = NOTEBOOK_DOCUMENT_ID, notebookName = NOTEBOOK_DOCUMENT_NAME, fresh = false } = {}) => {
    const fallbackPages = createNotebookPages(1);
    let record = null;

    if (!fresh) {
      try {
        record = await getStoredAnnotations(documentId);
      } catch (error) {
        record = null;
      }
    }

    const nextPages = sanitizeStoredPages(record?.pages, fallbackPages).map((page) => ({ ...page, kind: 'blank' }));

    pdfDocumentRef.current = null;
    originalPdfBytesRef.current = null;
    documentIdRef.current = documentId;
    annotationsRef.current = sanitizeStoredAnnotations(record?.annotations ?? {});
    remindersRef.current = sanitizeStoredReminders(record?.reminders);
    dictatedNotesRef.current = sanitizeStoredVoiceNotes(record?.voiceNotes);
    renderedPdfPagesRef.current = {};
    setPages(nextPages);
    setSourceType('notebook');
    setDocumentName(notebookName);
    setReminders(remindersRef.current);
    setDictatedNotes(dictatedNotesRef.current);
    setLiveTranscript('');
    setActivePageNumber(1);
    setPageJumpValue('1');
    setLastSavedAt(formatSavedTime(record?.updatedAt));
    setStorageLabel('IndexedDB primary storage');
    setStatusMessage(`Opened ${notebookName}. Start writing or upload a PDF.`);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setCurrentView('editor');
    updateRenderWindow(1);
    if (fresh) {
      persistSession({}, nextPages, 'notebook', notebookName, documentId, null, [], '');
    }
  }, [persistSession, updateRenderWindow]);

  const buildPdfPages = useCallback(async (pdf) => {
    const availableWidth = Math.max(Math.min((viewerShellRef.current?.clientWidth ?? 980) - 56, 980), 360);
    const layouts = [];

    for (let index = 1; index <= pdf.numPages; index += 1) {
      const page = await pdf.getPage(index);
      const viewport = page.getViewport({ scale: 1 });
      const scale = availableWidth / viewport.width;

      layouts.push({
        pageNumber: index,
        width: viewport.width * scale,
        height: viewport.height * scale,
        kind: 'pdf',
      });
    }

    return layouts;
  }, []);

  const loadPdfSession = useCallback(async (fileName, pdfBytes, documentId, existingRecord = null) => {
    setStatusMessage('Loading document...');

    try {
      const pdf = await getDocument({ data: pdfBytes }).promise;
      const nextPages = await buildPdfPages(pdf);
      let record = existingRecord;

      if (!record) {
        try {
          record = await getStoredAnnotations(documentId);
        } catch (error) {
          record = null;
        }
      }

      pdfDocumentRef.current = pdf;
      originalPdfBytesRef.current = pdfBytes;
      documentIdRef.current = documentId;
      annotationsRef.current = sanitizeStoredAnnotations(record?.annotations ?? {});
      remindersRef.current = sanitizeStoredReminders(record?.reminders);
      dictatedNotesRef.current = sanitizeStoredVoiceNotes(record?.voiceNotes);
      renderedPdfPagesRef.current = {};
      setPages(nextPages);
      setSourceType('pdf');
      setDocumentName(fileName);
      setReminders(remindersRef.current);
      setDictatedNotes(dictatedNotesRef.current);
      setLiveTranscript('');
      setActivePageNumber(1);
      setPageJumpValue('1');
      setLastSavedAt(formatSavedTime(record?.updatedAt));
      setStorageLabel('IndexedDB primary storage');
      setStatusMessage('PDF ready. Scroll, jump, annotate, or export.');
      setAnnotationVersion((currentValue) => currentValue + 1);
      setCurrentView('editor');
      updateRenderWindow(1);
      persistSession(record?.annotations ?? {}, nextPages, 'pdf', fileName, documentId, pdfBytes, remindersRef.current, dictatedNotesRef.current);
    } catch (error) {
      setStatusMessage('Unable to render that PDF. Try another file.');
    }
  }, [buildPdfPages, persistSession, updateRenderWindow]);

  const storeDictatedNotes = useCallback((nextVoiceNotes) => {
    dictatedNotesRef.current = nextVoiceNotes;
    setDictatedNotes(nextVoiceNotes);
    persistSession(
      annotationsRef.current,
      pages,
      sourceType,
      documentName,
      documentIdRef.current,
      originalPdfBytesRef.current,
      remindersRef.current,
      nextVoiceNotes,
    );
  }, [documentName, pages, persistSession, sourceType]);

  const handleDictatedNotesChange = useCallback((event) => {
    storeDictatedNotes(event.target.value);
  }, [storeDictatedNotes]);

  const saveReminder = useCallback((reminderText) => {
    const trimmedText = reminderText.trim();

    if (!trimmedText) {
      return false;
    }

    const reminder = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      text: trimmedText,
      createdAt: new Date().toISOString(),
      pageNumber: activePageNumber,
    };
    const nextReminders = [reminder, ...remindersRef.current].slice(0, 12);

    remindersRef.current = nextReminders;
    setReminders(nextReminders);
    persistSession(annotationsRef.current, pages, sourceType, documentName, documentIdRef.current, originalPdfBytesRef.current, nextReminders);
    return true;
  }, [activePageNumber, documentName, pages, persistSession, sourceType]);

  useEffect(() => {
    refreshLibrary();
  }, [refreshLibrary]);

  useEffect(() => {
    pageLayoutMapRef.current = pages.reduce((nextPages, page) => {
      nextPages[page.pageNumber] = page;
      return nextPages;
    }, {});
  }, [pages]);

  useEffect(() => {
    updateRenderWindow(activePageNumber);
  }, [activePageNumber, updateRenderWindow]);

  useEffect(() => {
    setPageJumpValue(String(activePageNumber));
  }, [activePageNumber]);

  useEffect(() => {
    pages.forEach(({ pageNumber, width, height }) => {
      if (renderWindowPageNumbers.includes(pageNumber)) {
        syncInkCanvas(pageNumber, width, height);
        return;
      }

      if (drawingStateRef.current.isDrawing && drawingStateRef.current.pageNumber === pageNumber) {
        return;
      }

      releaseCanvasResources(inkCanvasRefs.current[pageNumber], width, height);
      releaseCanvasResources(pdfCanvasRefs.current[pageNumber], width, height);
      delete renderedPdfPagesRef.current[pageNumber];
    });
  }, [pages, releaseCanvasResources, renderWindowPageNumbers, syncInkCanvas]);

  useEffect(() => {
    try {
      window.localStorage.setItem(TOOLBAR_POSITION_STORAGE_KEY, JSON.stringify(toolbarPosition));
    } catch (error) {
      // Ignore toolbar position persistence failures.
    }
  }, [toolbarPosition]);

  useEffect(() => {
    try {
      window.localStorage.setItem(TOOLBAR_COLLAPSED_STORAGE_KEY, String(isToolbarCollapsed));
    } catch (error) {
      // Ignore toolbar collapse persistence failures.
    }
  }, [isToolbarCollapsed]);

  useEffect(() => {
    const handlePointerMove = (event) => {
      if (!toolbarDragStateRef.current.active) {
        return;
      }

      const frameBounds = overlayFrameRef.current?.getBoundingClientRect();
      const toolbarBounds = toolbarFloatRef.current?.getBoundingClientRect();
      const rawX = toolbarDragStateRef.current.originX + event.clientX - toolbarDragStateRef.current.startX;
      const rawY = toolbarDragStateRef.current.originY + event.clientY - toolbarDragStateRef.current.startY;

      if (!frameBounds || !toolbarBounds) {
        setToolbarPosition({ x: rawX, y: rawY });
        return;
      }

      setToolbarPosition(clampToolbarPositionToBounds({ x: rawX, y: rawY }, frameBounds, toolbarBounds));
    };

    const stopDragging = () => {
      toolbarDragStateRef.current.active = false;
    };

    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', stopDragging);
    window.addEventListener('pointercancel', stopDragging);

    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', stopDragging);
      window.removeEventListener('pointercancel', stopDragging);
    };
  }, []);

  useEffect(() => {
    const normalizeToolbarPosition = () => {
      const frameBounds = overlayFrameRef.current?.getBoundingClientRect();
      const toolbarBounds = toolbarFloatRef.current?.getBoundingClientRect();

      if (!frameBounds || !toolbarBounds) {
        return;
      }

      setToolbarPosition((currentPosition) => {
        const nextPosition = clampToolbarPositionToBounds(currentPosition, frameBounds, toolbarBounds);

        if (nextPosition.x === currentPosition.x && nextPosition.y === currentPosition.y) {
          return currentPosition;
        }

        return nextPosition;
      });
    };

    normalizeToolbarPosition();
    window.addEventListener('resize', normalizeToolbarPosition);

    return () => {
      window.removeEventListener('resize', normalizeToolbarPosition);
    };
  }, [isToolbarCollapsed, toolPanel, currentView]);

  useEffect(() => {
    if (!hasPdf) {
      return;
    }

    let cancelled = false;

    const renderVisiblePages = async () => {
      const pagesToRender = pages.filter((page) => renderWindowPageNumbers.includes(page.pageNumber) && page.kind === 'pdf');

      for (const page of pagesToRender) {
        if (cancelled) {
          return;
        }

        await renderPdfPage(page);
      }
    };

    renderVisiblePages();

    return () => {
      cancelled = true;
    };
  }, [hasPdf, pages, renderPdfPage, renderWindowPageNumbers]);

  const getPoint = useCallback((pageNumber, event) => {
    const bounds = inkCanvasRefs.current[pageNumber]?.getBoundingClientRect();

    if (!bounds) {
      return { x: 0, y: 0 };
    }

    return {
      x: event.clientX - bounds.left,
      y: event.clientY - bounds.top,
    };
  }, []);

  const getNormalizedPoint = useCallback((pageLayout, point, event) => {
    const pressure = event.pointerType === 'pen' && Number.isFinite(event.pressure) && event.pressure > 0
      ? clampNumber(event.pressure, 0.05, 1)
      : 1;

    return {
      x: point.x / pageLayout.width,
      y: point.y / pageLayout.height,
      pressure,
    };
  }, []);

  const eraseStrokeAtPoint = useCallback((pageNumber, point) => {
    const pageAnnotations = ensurePageAnnotations(pageNumber);
    const pageSize = getPageSize(pageNumber);
    const threshold = eraserSize;
    const pointInPixels = {
      x: point.x * pageSize.width,
      y: point.y * pageSize.height,
    };

    let didErase = false;
    const nextStrokes = pageAnnotations.strokes.flatMap((stroke) => {
      if (eraseHighlighterOnly && stroke.tool !== 'highlighter') {
        return [stroke];
      }

      if (erasePenOnly && stroke.tool === 'highlighter') {
        return [stroke];
      }

      const strokePoints = stroke.points.map((strokePoint) => ({
        x: strokePoint.x * pageSize.width,
        y: strokePoint.y * pageSize.height,
      }));

      const touchesStroke = strokePoints.some((strokePoint, index) => {
        if (Math.hypot(pointInPixels.x - strokePoint.x, pointInPixels.y - strokePoint.y) <= threshold) {
          return true;
        }

        if (index === 0) {
          return false;
        }

        return distanceToSegment(pointInPixels, strokePoints[index - 1], strokePoint) <= threshold;
      });

      if (!touchesStroke) {
        return [stroke];
      }

      didErase = true;

      if (eraseEntireStroke) {
        return [];
      }

      return buildStrokeSegments(stroke, (strokePoint) => {
        return Math.hypot(pointInPixels.x - strokePoint.x * pageSize.width, pointInPixels.y - strokePoint.y * pageSize.height) <= threshold;
      });
    });

    if (!didErase) {
      return;
    }

    const nextAnnotations = {
      ...annotationsRef.current,
      [pageNumber]: {
        strokes: nextStrokes,
        redoStack: [],
      },
    };

    annotationsRef.current = nextAnnotations;
    redrawInkPage(pageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Erased ink on page ${pageNumber}.`);
    if (autoSelectPreviousTool) {
      setActiveTool(previousInkToolRef.current);
      setToolPanel(null);
    }
  }, [autoSelectPreviousTool, buildStrokeSegments, ensurePageAnnotations, eraseEntireStroke, eraseHighlighterOnly, erasePenOnly, eraserSize, getPageSize, persistSession, redrawInkPage]);

  const handlePointerDown = useCallback((pageNumber) => (event) => {
    if (event.pointerType === 'mouse' && event.button !== 0) {
      return;
    }

    const pageLayout = pageLayoutMapRef.current[pageNumber];
    const canvas = inkCanvasRefs.current[pageNumber];

    if (!pageLayout || !canvas) {
      return;
    }

    const point = getPoint(pageNumber, event);
    const normalizedPoint = getNormalizedPoint(pageLayout, point, event);

    if (activeTool === 'eraser') {
      eraseStrokeAtPoint(pageNumber, normalizedPoint);
      return;
    }

    if (!renderWindowPageNumbers.includes(pageNumber)) {
      updateRenderWindow(pageNumber);
      syncInkCanvas(pageNumber, pageLayout.width, pageLayout.height);
    }

    const context = canvas.getContext('2d');

    if (!context) {
      return;
    }

    const stroke = {
      points: [normalizedPoint],
      tool: activeTool,
      variant: activeTool === 'highlighter' ? selectedHighlighterPreset.id : selectedPenPreset.id,
      color: inkColor,
      width: activeTool === 'highlighter'
        ? Math.max(highlighterSize * selectedHighlighterPreset.widthScale, TOOL_PRESETS.highlighter.width * 0.8)
        : Math.max(penSize * selectedPenPreset.widthScale, 1.5),
      opacity: activeTool === 'highlighter' ? selectedHighlighterPreset.opacity : selectedPenPreset.opacity,
      pressureEnabled: event.pointerType === 'pen' && Number.isFinite(event.pressure) && event.pressure > 0,
    };

    drawingStateRef.current = { isDrawing: true, pageNumber, stroke };
    canvas.setPointerCapture?.(event.pointerId);
    redrawInkPage(pageNumber);
    drawStroke(context, pageNumber, stroke);
    setActivePageNumber(pageNumber);
  }, [activeTool, drawStroke, eraseStrokeAtPoint, getNormalizedPoint, getPoint, highlighterSize, inkColor, penSize, redrawInkPage, renderWindowPageNumbers, selectedHighlighterPreset, selectedPenPreset, syncInkCanvas, updateRenderWindow]);

  const handlePointerMove = useCallback((pageNumber) => (event) => {
    if (!drawingStateRef.current.isDrawing || drawingStateRef.current.pageNumber !== pageNumber) {
      return;
    }

    const context = inkCanvasRefs.current[pageNumber]?.getContext('2d');
    const pageLayout = pageLayoutMapRef.current[pageNumber];

    if (!context || !pageLayout) {
      return;
    }

    const point = getPoint(pageNumber, event);

    drawingStateRef.current.stroke.points.push(getNormalizedPoint(pageLayout, point, event));
    drawingStateRef.current.stroke.pressureEnabled = drawingStateRef.current.stroke.pressureEnabled
      || (event.pointerType === 'pen' && Number.isFinite(event.pressure) && event.pressure > 0);

    redrawInkPage(pageNumber);
    drawStroke(context, pageNumber, drawingStateRef.current.stroke);
  }, [drawStroke, getNormalizedPoint, getPoint, redrawInkPage]);

  const handlePointerUp = useCallback((pageNumber) => (event) => {
    if (!drawingStateRef.current.isDrawing || drawingStateRef.current.pageNumber !== pageNumber) {
      return;
    }

    const completedStroke = cloneStroke(drawingStateRef.current.stroke);

    drawingStateRef.current = { isDrawing: false, pageNumber: null, stroke: null };

    const canvas = inkCanvasRefs.current[pageNumber];

    canvas?.releasePointerCapture?.(event.pointerId);

    if (!completedStroke.points.length) {
      return;
    }

    const pageAnnotations = ensurePageAnnotations(pageNumber);
    const nextAnnotations = {
      ...annotationsRef.current,
      [pageNumber]: {
        strokes: [...pageAnnotations.strokes, completedStroke],
        redoStack: [],
      },
    };

    annotationsRef.current = nextAnnotations;
    redrawInkPage(pageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Stroke saved on page ${pageNumber}.`);
  }, [ensurePageAnnotations, persistSession, redrawInkPage]);

  const handleUndo = useCallback(() => {
    const pageAnnotations = ensurePageAnnotations(activePageNumber);

    if (!pageAnnotations.strokes.length) {
      return;
    }

    const lastStroke = pageAnnotations.strokes[pageAnnotations.strokes.length - 1];
    const nextAnnotations = {
      ...annotationsRef.current,
      [activePageNumber]: {
        strokes: pageAnnotations.strokes.slice(0, -1),
        redoStack: [cloneStroke(lastStroke), ...pageAnnotations.redoStack],
      },
    };

    annotationsRef.current = nextAnnotations;
    redrawInkPage(activePageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Undid the last stroke on page ${activePageNumber}.`);
  }, [activePageNumber, ensurePageAnnotations, persistSession, redrawInkPage]);

  const handleRedo = useCallback(() => {
    const pageAnnotations = ensurePageAnnotations(activePageNumber);

    if (!pageAnnotations.redoStack.length) {
      return;
    }

    const [restoredStroke, ...remainingRedo] = pageAnnotations.redoStack;
    const nextAnnotations = {
      ...annotationsRef.current,
      [activePageNumber]: {
        strokes: [...pageAnnotations.strokes, cloneStroke(restoredStroke)],
        redoStack: remainingRedo,
      },
    };

    annotationsRef.current = nextAnnotations;
    redrawInkPage(activePageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Restored a stroke on page ${activePageNumber}.`);
  }, [activePageNumber, ensurePageAnnotations, persistSession, redrawInkPage]);

  const handleClearPage = useCallback(() => {
    const nextAnnotations = {
      ...annotationsRef.current,
      [activePageNumber]: createEmptyPageAnnotations(),
    };

    annotationsRef.current = nextAnnotations;
    redrawInkPage(activePageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Cleared page ${activePageNumber}.`);
  }, [activePageNumber, persistSession, redrawInkPage]);

  const scrollToPage = useCallback((pageNumber) => {
    const target = pageStageRefs.current[pageNumber];

    if (!target) {
      return;
    }

    updateRenderWindow(pageNumber);
    target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    setActivePageNumber(pageNumber);
  }, [updateRenderWindow]);

  const processVoiceTranscript = useCallback((transcript) => {
    const normalizedTranscript = transcript.trim();

    if (!normalizedTranscript) {
      return;
    }

    const normalizedCommandTranscript = normalizedTranscript.replace(/[^\w\u0900-\u097F\s]/g, '');
    const compactCommandTranscript = normalizedCommandTranscript.trim().toLowerCase();
    const feedback = [];
    let handledCommand = false;

    if (compactCommandTranscript === 'highlight') {
      previousInkToolRef.current = 'highlighter';
      setActiveTool('highlighter');
      setToolPanel('highlighter');
      feedback.push('highlighter ready');
      handledCommand = true;
    }

    if (compactCommandTranscript === 'पुढचा' || compactCommandTranscript === 'next') {
      const nextPageNumber = Math.min(pageCount, activePageNumber + 1);

      if (nextPageNumber !== activePageNumber) {
        scrollToPage(nextPageNumber);
        feedback.push(`moved to page ${nextPageNumber}`);
      } else {
        feedback.push('already on the last page');
      }

      handledCommand = true;
    }

    const reminderText = extractReminderText(normalizedTranscript);

    if (reminderText) {
      if (saveReminder(reminderText)) {
        feedback.push(`saved reminder: ${reminderText}`);
      }

      handledCommand = true;
    }

    if (!handledCommand) {
      const nextVoiceNotes = appendDictatedText(dictatedNotesRef.current, normalizedTranscript);

      storeDictatedNotes(nextVoiceNotes);
      feedback.push('dictation added');
    }

    if (feedback.length) {
      setStatusMessage(`Voice command: ${feedback.join(' • ')}.`);
    }
  }, [activePageNumber, pageCount, saveReminder, scrollToPage, storeDictatedNotes]);

  useEffect(() => {
    shouldKeepListeningRef.current = isListening;

    if (!isListening) {
      const recognition = speechRecognitionRef.current;

      if (recognition) {
        recognition.onresult = null;
        recognition.onerror = null;
        recognition.onend = null;

        try {
          recognition.stop();
        } catch (error) {
          // Ignore redundant stop calls.
        }
      }

      speechRecognitionRef.current = null;
      return undefined;
    }

    const SpeechRecognitionConstructor = getSpeechRecognitionConstructor();

    if (!SpeechRecognitionConstructor) {
      setIsListening(false);
      setStatusMessage('Speech recognition is not supported in this browser.');
      return undefined;
    }

    const recognition = new SpeechRecognitionConstructor();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.lang = 'en-IN';
    recognition.onresult = (event) => {
      let nextLiveTranscript = '';

      for (let index = event.resultIndex; index < event.results.length; index += 1) {
        const result = event.results[index];

        if (!result.isFinal) {
          nextLiveTranscript += `${result[0]?.transcript ?? ''} `;
          continue;
        }

        processVoiceTranscript(result[0]?.transcript ?? '');
      }

      setLiveTranscript(nextLiveTranscript.trim());
    };
    recognition.onerror = (event) => {
      if (event.error === 'not-allowed' || event.error === 'service-not-allowed') {
        shouldKeepListeningRef.current = false;
        setIsListening(false);
        setStatusMessage('Microphone access was blocked, so voice commands are unavailable.');
        return;
      }

      if (event.error === 'no-speech') {
        return;
      }

      setStatusMessage(`Voice commands paused: ${event.error}.`);
    };
    recognition.onend = () => {
      if (!shouldKeepListeningRef.current) {
        return;
      }

      try {
        recognition.start();
      } catch (error) {
        setIsListening(false);
        setStatusMessage('Voice commands could not restart. Tap the mic to try again.');
      }
    };

    speechRecognitionRef.current = recognition;

    try {
      recognition.start();
      setStatusMessage('Voice dictation is live. Speak normally to type, or say Next, Highlight, or Reminder.');
    } catch (error) {
      speechRecognitionRef.current = null;
      shouldKeepListeningRef.current = false;
      setIsListening(false);
      setStatusMessage('Voice commands could not start. Check microphone access and try again.');
    }

    return () => {
      shouldKeepListeningRef.current = false;
      setLiveTranscript('');
      recognition.onresult = null;
      recognition.onerror = null;
      recognition.onend = null;

      try {
        recognition.stop();
      } catch (error) {
        // Ignore redundant stop calls.
      }

      if (speechRecognitionRef.current === recognition) {
        speechRecognitionRef.current = null;
      }
    };
  }, [isListening, processVoiceTranscript]);

  const handleJumpToPage = useCallback((event) => {
    event.preventDefault();

    const parsedPage = Number.parseInt(pageJumpValue, 10);

    if (!Number.isFinite(parsedPage)) {
      setPageJumpValue(String(activePageNumber));
      setStatusMessage('Enter a valid page number before jumping.');
      return;
    }

    const nextPageNumber = Math.max(1, Math.min(pageCount, parsedPage));
    scrollToPage(nextPageNumber);
    setStatusMessage(`Jumped to page ${nextPageNumber}.`);
  }, [activePageNumber, pageCount, pageJumpValue, scrollToPage]);

  const handleViewerScroll = useCallback(() => {
    const shellTop = viewerShellRef.current?.getBoundingClientRect().top ?? 0;
    let closestPage = activePageNumber;
    let closestDistance = Number.POSITIVE_INFINITY;

    pages.forEach((page) => {
      const rect = pageStageRefs.current[page.pageNumber]?.getBoundingClientRect();

      if (!rect) {
        return;
      }

      const distance = Math.abs(rect.top - shellTop - 88);

      if (distance < closestDistance) {
        closestDistance = distance;
        closestPage = page.pageNumber;
      }
    });

    if (closestPage !== activePageNumber) {
      updateRenderWindow(closestPage);
      setActivePageNumber(closestPage);
    }
  }, [activePageNumber, pages, updateRenderWindow]);

  const handleAddPage = useCallback(() => {
    const nextPages = [...pages, {
      pageNumber: pages.length + 1,
      width: NOTEBOOK_PAGE_WIDTH,
      height: NOTEBOOK_PAGE_HEIGHT,
      kind: 'blank',
    }];

    setPages(nextPages);
    persistSession(annotationsRef.current, nextPages, 'notebook', documentName);
    setStatusMessage(`Added blank page ${nextPages.length}.`);
    window.setTimeout(() => {
      scrollToPage(nextPages.length);
    }, 0);
  }, [documentName, pages, persistSession, scrollToPage]);

  const handleFileChange = useCallback(async (event) => {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    const pdfBytes = new Uint8Array(await file.arrayBuffer());
    await loadPdfSession(file.name, pdfBytes, buildDocumentId(file.name, pdfBytes));
    event.target.value = '';
  }, [loadPdfSession]);

  const handleOpenHome = useCallback(() => {
    setCurrentView('home');
    setStatusMessage('Choose a note workspace to open.');
    refreshLibrary();
  }, [refreshLibrary]);

  const handleCreateFreshNotebook = useCallback(() => {
    const notebookName = createNotebookName();
    const notebookId = `notebook:${Date.now()}`;

    loadNotebookSession({ documentId: notebookId, notebookName, fresh: true });
  }, [loadNotebookSession]);

  const handleOpenStoredItem = useCallback(async (libraryItem) => {
    if (!libraryItem) {
      return;
    }

    let fullRecord = null;

    try {
      fullRecord = await getStoredAnnotations(libraryItem.documentId);
    } catch (error) {
      fullRecord = null;
    }

    if (libraryItem.sourceType === 'pdf') {
      const sourceBytes = fullRecord?.sourceBytes ?? libraryItem.sourceBytes;

      if (!sourceBytes) {
        setStatusMessage('This PDF needs to be imported again because its source file is unavailable.');
        return;
      }

      await loadPdfSession(
        libraryItem.documentName,
        new Uint8Array(sourceBytes),
        libraryItem.documentId,
        fullRecord,
      );
      return;
    }

    await loadNotebookSession({
      documentId: libraryItem.documentId,
      notebookName: libraryItem.documentName,
      fresh: false,
    });
  }, [loadNotebookSession, loadPdfSession]);

  const handleToolbarDragStart = useCallback((event) => {
    toolbarDragStateRef.current = {
      active: true,
      startX: event.clientX,
      startY: event.clientY,
      originX: toolbarPosition.x,
      originY: toolbarPosition.y,
    };
  }, [toolbarPosition.x, toolbarPosition.y]);

  const drawExportStroke = useCallback((page, stroke) => {
    const { width, height } = page.getSize();
    const rgbColor = rgb(
      Number.parseInt(stroke.color.slice(1, 3), 16) / 255,
      Number.parseInt(stroke.color.slice(3, 5), 16) / 255,
      Number.parseInt(stroke.color.slice(5, 7), 16) / 255,
    );

    if (stroke.points.length === 1) {
      const point = stroke.points[0];
      page.drawCircle({
        x: point.x * width,
        y: height - point.y * height,
        size: (stroke.width * getPointPressureMultiplier(stroke, point)) / 2,
        color: rgbColor,
        opacity: stroke.opacity,
      });
      return;
    }

    for (let index = 1; index < stroke.points.length; index += 1) {
      const startPoint = stroke.points[index - 1];
      const endPoint = stroke.points[index];

      page.drawLine({
        start: { x: startPoint.x * width, y: height - startPoint.y * height },
        end: { x: endPoint.x * width, y: height - endPoint.y * height },
        thickness: stroke.width * ((getPointPressureMultiplier(stroke, startPoint) + getPointPressureMultiplier(stroke, endPoint)) / 2),
        color: rgbColor,
        opacity: stroke.opacity,
      });
    }
  }, []);

  const drawNotebookRuling = useCallback((page) => {
    const { width, height } = page.getSize();

    for (let y = 52; y < height; y += 42) {
      page.drawLine({
        start: { x: 36, y },
        end: { x: width - 24, y },
        thickness: 0.8,
        color: rgb(0.82, 0.84, 0.88),
      });
    }

    page.drawLine({
      start: { x: 110, y: 24 },
      end: { x: 110, y: height - 24 },
      thickness: 1.1,
      color: rgb(0.93, 0.55, 0.5),
    });
  }, []);

  const handleExportPdf = useCallback(async () => {
    setIsExporting(true);
    setStatusMessage('Preparing export...');

    try {
      const exportDocument = hasPdf && originalPdfBytesRef.current
        ? await PDFDocument.load(originalPdfBytesRef.current.slice())
        : await PDFDocument.create();

      if (!hasPdf) {
        pages.forEach(() => {
          exportDocument.addPage([NOTEBOOK_PAGE_WIDTH, NOTEBOOK_PAGE_HEIGHT]);
        });
      }

      exportDocument.getPages().forEach((page, index) => {
        if (!hasPdf) {
          drawNotebookRuling(page);
        }

        const pageAnnotations = annotationsRef.current[index + 1]?.strokes ?? [];
        pageAnnotations.forEach((stroke) => drawExportStroke(page, stroke));
      });

      const exportBytes = await exportDocument.save();
      const exportBlob = new Blob([exportBytes], { type: 'application/pdf' });
      const exportUrl = window.URL.createObjectURL(exportBlob);
      const link = document.createElement('a');

      link.href = exportUrl;
      link.download = buildExportName(documentName, sourceType);
      link.click();
      window.setTimeout(() => {
        window.URL.revokeObjectURL(exportUrl);
      }, 0);
      setStatusMessage(`Exported ${buildExportName(documentName, sourceType)}.`);
    } catch (error) {
      setStatusMessage('Unable to export the current document.');
    } finally {
      setIsExporting(false);
    }
  }, [documentName, drawExportStroke, drawNotebookRuling, hasPdf, pages, sourceType]);

  const activePageAnnotations = annotationsRef.current[activePageNumber] ?? createEmptyPageAnnotations();
  const totalStrokeCount = Object.values(annotationsRef.current).reduce((strokeCount, pageAnnotations) => {
    return strokeCount + (pageAnnotations?.strokes?.length ?? 0);
  }, 0);
  const _annotationVersion = annotationVersion;
  void _annotationVersion;

  return (
    <div className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Tablet-style notes</p>
          <h1>VoxNotes AI</h1>
          <p className="header-copy">
            Start from a notes library, then open a ruled notebook or imported PDF. Write with pen or highlighter, erase strokes, add blank pages, jump through long documents, and export everything back to PDF.
          </p>
        </div>

        <button
          type="button"
          className={`listen-chip ${isListening ? 'active' : ''}`}
          onClick={() => setIsListening((currentValue) => !currentValue)}
        >
          {isListening ? <Mic size={18} /> : <MicOff size={18} />}
          <span>{isListening ? 'Listening...' : 'Mic idle'}</span>
        </button>
      </header>

      <main className="workspace-card">
        {currentView === 'home' ? (
          <div className="library-shell">
            <aside className="library-sidebar">
              <h2 className="library-brand">VoxNotes</h2>

              <div className="library-tile-grid">
                {LIBRARY_TILES.map((tile) => {
                  const TileIcon = tile.icon;

                  return (
                    <button key={tile.key} type="button" className={`library-tile ${tile.accentClass}`}>
                      <span className="library-tile-icon"><TileIcon size={16} /></span>
                      <span>{tile.title}</span>
                      {tile.badge ? <strong>{tile.badge}</strong> : null}
                    </button>
                  );
                })}
              </div>

              <div className="library-folder-card">
                <div className="library-section-head">
                  <h3>Folders</h3>
                </div>
                <button type="button" className="folder-row">
                  <FolderOpen size={16} />
                  <span>My Notes</span>
                </button>
                <button type="button" className="folder-row" onClick={handleCreateFreshNotebook}>
                  <Plus size={16} />
                  <span>New Folder</span>
                </button>
              </div>

              <div className="library-upgrade-card">
                <span className="upgrade-badge">Workspace</span>
                <strong>Open the right note fast</strong>
                <p>Start from a quick note, create a fresh notebook, or reopen a saved session.</p>
              </div>
            </aside>

            <section className="library-main">
              <div className="library-main-head">
                <div>
                  <p className="eyebrow library-eyebrow">Library</p>
                  <h2>All Notes</h2>
                </div>
              </div>

              <div className="library-actions-row">
                <button type="button" className="library-action-card" onClick={() => loadNotebookSession()}>
                  <span className="library-action-icon plus">+</span>
                  <strong>Quick Note</strong>
                </button>
                <button type="button" className="library-action-card" onClick={handleCreateFreshNotebook}>
                  <span className="library-action-icon notebook"><NotebookPen size={18} /></span>
                  <strong>New Notebook</strong>
                </button>
                <button type="button" className="library-action-card" onClick={() => fileInputRef.current?.click()}>
                  <span className="library-action-icon import"><Download size={18} /></span>
                  <strong>Import File</strong>
                </button>
              </div>

              <div className="library-section-head">
                <h3>Recent</h3>
                <span>{isLibraryLoading ? 'Loading...' : `${libraryItems.length} saved`}</span>
              </div>

              <div className="library-note-grid">
                {libraryItems.length ? libraryItems.map((item) => (
                  <button key={item.documentId} type="button" className="note-card" onClick={() => handleOpenStoredItem(item)}>
                    <div className={`note-card-cover ${item.accentClass} ${item.sourceType === 'pdf' ? 'pdf' : 'notebook'}`}>
                      {item.sourceType === 'pdf' ? <FileText size={22} /> : <NotebookPen size={22} />}
                    </div>
                    <strong>{item.documentName}</strong>
                    <span>{item.sourceType === 'pdf' ? 'Imported PDF' : 'Notebook'}</span>
                    <span>{item.pageCount} pages</span>
                    <span>{item.updatedAt ? new Date(item.updatedAt).toLocaleString([], { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : 'Never saved'}</span>
                  </button>
                )) : (
                  <div className="library-empty-state">
                    <NotebookPen size={28} />
                    <strong>No saved notes yet</strong>
                    <p>Start with a quick note, create a notebook, or import a PDF to build your library.</p>
                  </div>
                )}
              </div>
            </section>
          </div>
        ) : (
          <>
            <div className="viewer-panel">
              <div className="viewer-overlay" aria-label="Editor controls">
                <div className="overlay-frame" ref={overlayFrameRef}>
                  <div ref={toolbarFloatRef} className="viewer-tool-float" style={{ transform: `translate(${toolbarPosition.x}px, ${toolbarPosition.y}px)` }}>
                    <div className="tool-rail" role="toolbar" aria-label="Writing tools">
                      <button type="button" className="tool-rail-handle" onPointerDown={handleToolbarDragStart} aria-label="Move toolbar">
                        <span />
                        <span />
                        <span />
                      </button>
                      <button type="button" className="tool-rail-toggle" onClick={() => setIsToolbarCollapsed((currentValue) => !currentValue)} aria-label={isToolbarCollapsed ? 'Show top toolbar' : 'Hide top toolbar'}>
                        <span>{isToolbarCollapsed ? '+' : '-'}</span>
                      </button>
                      <button type="button" className={`tool-rail-button ${activeTool === 'pen' ? 'selected' : ''}`} onClick={() => handleToolSelect('pen')} aria-label="Pen">
                        <span className="tool-rail-icon simple">
                          <RailGlyph kind="pen" />
                        </span>
                      </button>
                      <button type="button" className={`tool-rail-button ${activeTool === 'highlighter' ? 'selected' : ''}`} onClick={() => handleToolSelect('highlighter')} aria-label="Highlighter">
                        <span className="tool-rail-icon simple">
                          <RailGlyph kind="highlighter" />
                        </span>
                      </button>
                      <button type="button" className={`tool-rail-button ${activeTool === 'eraser' ? 'selected' : ''}`} onClick={() => handleToolSelect('eraser')} aria-label="Eraser">
                        <span className="tool-rail-icon simple">
                          <RailGlyph kind="eraser" />
                        </span>
                      </button>
                      <button type="button" className="tool-rail-button action" onClick={handleUndo} disabled={!activePageAnnotations.strokes.length} aria-label="Undo">
                        <Undo2 size={13} />
                      </button>
                      <button type="button" className="tool-rail-button action" onClick={handleRedo} disabled={!activePageAnnotations.redoStack.length} aria-label="Redo">
                        <Redo2 size={13} />
                      </button>

                      <div className="tool-rail-colors" aria-label="Quick colors">
                        {QUICK_TOOL_COLORS.map((color) => (
                          <button
                            key={color}
                            type="button"
                            className={`tool-rail-color-dot ${inkColor === color ? 'selected' : ''}`}
                            style={{ backgroundColor: color }}
                            onClick={() => setInkColor(color)}
                            aria-label={`Quick color ${color}`}
                          />
                        ))}
                      </div>
                    </div>

                    {toolPanel ? (
                      <div className="tool-popover tool-popover-floating" role="dialog" aria-label={`${toolPanel} settings`}>
                        {toolPanel === 'pen' ? (
                          <>
                            <div className="tool-popover-head">
                              <strong>Pen tips</strong>
                              <span className="tool-hint">Pressure reacts automatically on supported pens.</span>
                            </div>
                            <div className="tool-preset-row">
                              {PEN_VARIANTS.map((preset) => (
                                <button key={preset.id} type="button" className={`tool-preset ${penVariant === preset.id ? 'selected' : ''}`} onClick={() => setPenVariant(preset.id)}>
                                  <span className="tool-tip-preview" style={{ color: inkColor }}>
                                    <ToolTipIllustration tool="pen" variant={preset.id} color={inkColor} />
                                  </span>
                                  <span>{preset.label}</span>
                                </button>
                              ))}
                            </div>

                            <div className="tool-favorites-row" aria-label="Pen favorites">
                              {penFavorites.map((favorite, index) => (
                                <div key={`pen-favorite-${index}`} className="favorite-slot">
                                  <button
                                    type="button"
                                    className={`favorite-chip ${favorite ? 'filled' : 'empty'}`}
                                    onClick={() => applyFavorite('pen', favorite)}
                                    disabled={!favorite}
                                    aria-label={favorite ? `Apply pen favorite ${index + 1}` : `Empty pen favorite ${index + 1}`}
                                  >
                                    {favorite ? (
                                      <>
                                        <span className="tool-tip-preview favorite-preview" style={{ color: favorite.color }}>
                                          <ToolTipIllustration tool="pen" variant={favorite.variant} color={favorite.color} />
                                        </span>
                                        <span>P{index + 1}</span>
                                      </>
                                    ) : <span>P{index + 1}</span>}
                                  </button>
                                  <button type="button" className="favorite-save-button" onClick={() => saveFavorite('pen', index)} aria-label={`Save pen favorite ${index + 1}`}>
                                    <Save size={14} />
                                  </button>
                                </div>
                              ))}
                            </div>

                            <div className="tool-slider-row">
                              <button type="button" className="mini-icon-button" onClick={() => setPenSize((currentValue) => Math.max(2, currentValue - 1))}>-</button>
                              <div className="tool-slider-pill">{penSize}</div>
                              <input type="range" min="2" max="14" value={penSize} onChange={(event) => setPenSize(Number.parseInt(event.target.value, 10))} />
                              <button type="button" className="mini-icon-button" onClick={() => setPenSize((currentValue) => Math.min(14, currentValue + 1))}>+</button>
                            </div>

                            <div className="color-group tool-popover-colors" aria-label="Pen colors">
                              {COLOR_SWATCHS.map((color) => (
                                <button
                                  key={color}
                                  type="button"
                                  className={`color-chip ${inkColor === color ? 'selected' : ''}`}
                                  style={{ backgroundColor: color }}
                                  onClick={() => setInkColor(color)}
                                  aria-label={`Color ${color}`}
                                />
                              ))}
                            </div>

                            <label className="custom-color-row">
                              <span>Custom color</span>
                              <input type="color" value={inkColor} onChange={(event) => setInkColor(event.target.value)} aria-label="Custom pen color" />
                              <strong>{inkColor.toUpperCase()}</strong>
                            </label>
                          </>
                        ) : null}

                        {toolPanel === 'highlighter' ? (
                          <>
                            <div className="tool-popover-head">
                              <strong>Highlighter tips</strong>
                              <span className="tool-hint">Save a few favorite marker setups for fast switching.</span>
                            </div>
                            <div className="tool-preset-row">
                              {HIGHLIGHTER_VARIANTS.map((preset) => (
                                <button key={preset.id} type="button" className={`tool-preset ${highlighterVariant === preset.id ? 'selected' : ''}`} onClick={() => setHighlighterVariant(preset.id)}>
                                  <span className="tool-tip-preview" style={{ color: inkColor }}>
                                    <ToolTipIllustration tool="highlighter" variant={preset.id} color={inkColor} />
                                  </span>
                                  <span>{preset.label}</span>
                                </button>
                              ))}
                            </div>

                            <div className="tool-favorites-row" aria-label="Highlighter favorites">
                              {highlighterFavorites.map((favorite, index) => (
                                <div key={`highlighter-favorite-${index}`} className="favorite-slot">
                                  <button
                                    type="button"
                                    className={`favorite-chip ${favorite ? 'filled' : 'empty'}`}
                                    onClick={() => applyFavorite('highlighter', favorite)}
                                    disabled={!favorite}
                                    aria-label={favorite ? `Apply highlighter favorite ${index + 1}` : `Empty highlighter favorite ${index + 1}`}
                                  >
                                    {favorite ? (
                                      <>
                                        <span className="tool-tip-preview favorite-preview" style={{ color: favorite.color }}>
                                          <ToolTipIllustration tool="highlighter" variant={favorite.variant} color={favorite.color} />
                                        </span>
                                        <span>H{index + 1}</span>
                                      </>
                                    ) : <span>H{index + 1}</span>}
                                  </button>
                                  <button type="button" className="favorite-save-button" onClick={() => saveFavorite('highlighter', index)} aria-label={`Save highlighter favorite ${index + 1}`}>
                                    <Save size={14} />
                                  </button>
                                </div>
                              ))}
                            </div>

                            <div className="tool-slider-row">
                              <button type="button" className="mini-icon-button" onClick={() => setHighlighterSize((currentValue) => Math.max(8, currentValue - 2))}>-</button>
                              <div className="tool-slider-pill">{highlighterSize}</div>
                              <input type="range" min="8" max="32" value={highlighterSize} onChange={(event) => setHighlighterSize(Number.parseInt(event.target.value, 10))} />
                              <button type="button" className="mini-icon-button" onClick={() => setHighlighterSize((currentValue) => Math.min(32, currentValue + 2))}>+</button>
                            </div>

                            <div className="color-group tool-popover-colors" aria-label="Highlighter colors">
                              {COLOR_SWATCHS.map((color) => (
                                <button
                                  key={color}
                                  type="button"
                                  className={`color-chip ${inkColor === color ? 'selected' : ''}`}
                                  style={{ backgroundColor: color }}
                                  onClick={() => setInkColor(color)}
                                  aria-label={`Color ${color}`}
                                />
                              ))}
                            </div>

                            <label className="custom-color-row">
                              <span>Custom color</span>
                              <input type="color" value={inkColor} onChange={(event) => setInkColor(event.target.value)} aria-label="Custom highlighter color" />
                              <strong>{inkColor.toUpperCase()}</strong>
                            </label>
                          </>
                        ) : null}

                        {toolPanel === 'eraser' ? (
                          <>
                            <div className="eraser-size-row">
                              {ERASER_SIZES.map((size) => (
                                <button key={size} type="button" className={`eraser-size-chip ${eraserSize === size ? 'selected' : ''}`} onClick={() => setEraserSize(size)}>
                                  <span style={{ width: `${Math.max(10, size / 1.6)}px`, height: `${Math.max(10, size / 1.6)}px` }} />
                                </button>
                              ))}
                            </div>

                            <label className="eraser-toggle-row">
                              <span>Auto Select Previous Tool</span>
                              <input type="checkbox" checked={autoSelectPreviousTool} onChange={(event) => setAutoSelectPreviousTool(event.target.checked)} />
                            </label>
                            <label className="eraser-toggle-row">
                              <span>Erase Entire Stroke</span>
                              <input type="checkbox" checked={eraseEntireStroke} onChange={(event) => setEraseEntireStroke(event.target.checked)} />
                            </label>
                            <label className="eraser-toggle-row">
                              <span>Erase Highlighter Only</span>
                              <input
                                type="checkbox"
                                checked={eraseHighlighterOnly}
                                onChange={(event) => {
                                  setEraseHighlighterOnly(event.target.checked);
                                  if (event.target.checked) {
                                    setErasePenOnly(false);
                                  }
                                }}
                              />
                            </label>
                            <label className="eraser-toggle-row">
                              <span>Erase Pencil Only</span>
                              <input
                                type="checkbox"
                                checked={erasePenOnly}
                                onChange={(event) => {
                                  setErasePenOnly(event.target.checked);
                                  if (event.target.checked) {
                                    setEraseHighlighterOnly(false);
                                  }
                                }}
                              />
                            </label>
                            <button type="button" className="text-action-button" onClick={handleClearPage}>Clear Page</button>
                          </>
                        ) : null}
                      </div>
                    ) : null}
                  </div>

                  {!isToolbarCollapsed ? (
                    <div className="overlay-stack">
                      <div className="overlay-toolbar">
                        <div className="mode-group">
                          <button type="button" className="overlay-button" onClick={handleOpenHome}>
                            <Home size={16} />
                            <span>Library</span>
                          </button>
                          <button type="button" className={`overlay-button ${sourceType === 'notebook' ? 'selected' : ''}`} onClick={() => loadNotebookSession()}>
                            <span>Quick Note</span>
                          </button>
                          <button type="button" className="overlay-button primary" onClick={() => fileInputRef.current?.click()}>
                            <FileText size={16} />
                            <span>Upload PDF</span>
                          </button>
                          <button type="button" className="overlay-button" onClick={handleAddPage} disabled={sourceType !== 'notebook'}>
                            <Plus size={16} />
                            <span>Add page</span>
                          </button>
                        </div>

                        <div className="page-controls">
                          <button type="button" className="overlay-button" onClick={() => scrollToPage(Math.max(1, activePageNumber - 1))} disabled={activePageNumber === 1}>
                            <ChevronLeft size={16} />
                            <span>Prev</span>
                          </button>
                          <div className="page-pill">
                            <span>Page</span>
                            <strong>{`${activePageNumber} / ${pageCount}`}</strong>
                          </div>
                          <button type="button" className="overlay-button" onClick={() => scrollToPage(Math.min(pageCount, activePageNumber + 1))} disabled={activePageNumber === pageCount}>
                            <span>Next</span>
                            <ChevronRight size={16} />
                          </button>
                        </div>

                        <form className="page-jump-form" onSubmit={handleJumpToPage}>
                          <label className="page-jump-label" htmlFor="page-jump-input">Jump</label>
                          <input
                            id="page-jump-input"
                            className="page-jump-input"
                            type="number"
                            min="1"
                            max={pageCount}
                            value={pageJumpValue}
                            onChange={(event) => setPageJumpValue(event.target.value)}
                          />
                          <button type="submit" className="overlay-button compact">
                            <Search size={14} />
                            <span>Go</span>
                          </button>
                        </form>

                        <button type="button" className="overlay-button compact" onClick={handleClearPage}>
                          <Eraser size={15} />
                          <span>Clear</span>
                        </button>
                        <button type="button" className="overlay-button compact" onClick={handleExportPdf} disabled={isExporting}>
                          <Download size={15} />
                          <span>{isExporting ? 'Exporting...' : 'Export'}</span>
                        </button>

                        <div className="save-pill compact-pill">
                          <Save size={14} />
                          <span>{lastSavedAt ? `Saved ${lastSavedAt}` : 'Auto-save ready'}</span>
                        </div>

                        <div className={`status-pill compact-pill ${isListening ? 'active' : ''}`}>
                          <span className="status-dot" />
                          <span>{isListening ? 'Listening...' : statusMessage}</span>
                        </div>
                      </div>
                    </div>
                  ) : null}
                </div>
              </div>

              <div className="viewer-shell" ref={viewerShellRef} onScroll={handleViewerScroll}>
                <div className="page-stack">
                  {pages.map((page) => {
                    const isInRenderWindow = renderWindowPageNumbers.includes(page.pageNumber);

                    return (
                      <section
                        key={page.pageNumber}
                        ref={(node) => {
                          pageStageRefs.current[page.pageNumber] = node;
                        }}
                        className={`page-stage ${activePageNumber === page.pageNumber ? 'active' : ''}`}
                        onPointerEnter={() => setActivePageNumber(page.pageNumber)}
                        onClick={() => setActivePageNumber(page.pageNumber)}
                      >
                        <div className="page-badge">Page {page.pageNumber}</div>
                        <div className="stack-stage" style={{ width: `${page.width}px`, height: `${page.height}px` }}>
                          {page.kind === 'pdf' ? (
                            <div className={`stack-layer pdf-layer ${isInRenderWindow ? 'active' : 'deferred'}`}>
                              {!isInRenderWindow ? <div className="deferred-page-hint">Rendering resumes when this page enters the active window.</div> : null}
                              <canvas
                                className="pdf-canvas"
                                ref={(node) => {
                                  pdfCanvasRefs.current[page.pageNumber] = node;
                                }}
                                aria-label={`PDF page ${page.pageNumber}`}
                              />
                            </div>
                          ) : (
                            <div className="stack-layer notebook-layer" aria-label={`Blank ruled page ${page.pageNumber}`} />
                          )}

                          <canvas
                            ref={(node) => {
                              inkCanvasRefs.current[page.pageNumber] = node;
                            }}
                            className={`stack-layer ink-layer ${activeTool === 'eraser' ? 'eraser-cursor' : ''}`}
                            aria-label={`Handwriting canvas page ${page.pageNumber}`}
                            onPointerDown={handlePointerDown(page.pageNumber)}
                            onPointerMove={handlePointerMove(page.pageNumber)}
                            onPointerUp={handlePointerUp(page.pageNumber)}
                            onPointerLeave={handlePointerUp(page.pageNumber)}
                          />
                        </div>
                      </section>
                    );
                  })}
                </div>
              </div>
            </div>

            <div className="workspace-meta">
              <div>
                <span className="meta-label">Current workspace</span>
                <strong>{documentName}</strong>
              </div>
              <div>
                <span className="meta-label">Mode</span>
                <strong>{hasPdf ? 'PDF annotation' : 'Blank notebook writing'}</strong>
              </div>
              <div>
                <span className="meta-label">Tooling</span>
                <strong>{activeTool === 'highlighter' ? 'Highlighter active' : activeTool === 'eraser' ? 'Stroke eraser active' : 'Pen active'}</strong>
              </div>
              <div>
                <span className="meta-label">Persistence</span>
                <strong>{storageLabel}</strong>
              </div>
              <div>
                <span className="meta-label">Ink summary</span>
                <strong>{totalStrokeCount} saved strokes</strong>
              </div>
              <div>
                <span className="meta-label">Voice reminders</span>
                <strong>{latestReminder ? latestReminder.text : 'No reminders yet'}</strong>
                <span className="meta-subcopy">{reminders.length ? `${reminders.length} saved • page ${latestReminder.pageNumber}` : 'Say “Reminder …” while the mic is live'}</span>
              </div>
            </div>

            <section className="voice-notes-panel" aria-label="Voice notes">
              <div className="voice-notes-head">
                <div>
                  <span className="meta-label">Voice Notes</span>
                  <strong>Dictation pad</strong>
                </div>
                <span className={`voice-notes-badge ${isListening ? 'active' : ''}`}>{isListening ? 'Typing from mic' : 'Manual + mic'}</span>
              </div>
              <textarea
                className="voice-notes-input"
                value={dictatedNotes}
                onChange={handleDictatedNotesChange}
                placeholder="Speak while the mic is on, or type here directly. Normal speech is added here like dictated notes."
              />
              <div className="voice-notes-footer">
                <span className="meta-subcopy">{liveTranscript ? `Hearing: ${liveTranscript}` : 'Commands stay active only for exact phrases like “Next”, “Highlight”, and “Reminder …”.'}</span>
              </div>
            </section>
          </>
        )}

        <input ref={fileInputRef} type="file" accept="application/pdf" hidden onChange={handleFileChange} />
      </main>
    </div>
  );
}

export default App;