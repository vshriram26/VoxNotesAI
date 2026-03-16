import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Bookmark, ChevronLeft, ChevronRight, Crown, Download, EllipsisVertical, Eraser, FileText, FolderOpen, Home, Image, Keyboard, Mic, NotebookPen, Plus, Redo2, Save, Search, Settings2, Star, Trash2, Undo2 } from 'lucide-react';
import { PDFDocument, rgb } from 'pdf-lib';
import { GlobalWorkerOptions, getDocument } from 'pdfjs-dist/legacy/build/pdf';
import { getStoredAnnotations, listStoredAnnotations, saveStoredAnnotations, saveReminder } from './annotationStorage';
import './App.css';

GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/legacy/build/pdf.worker.min.mjs', import.meta.url).toString();

const NOTEBOOK_DOCUMENT_ID = 'notebook:quick-notes';
const NOTEBOOK_DOCUMENT_NAME = 'Quick Notes';
const LIBRARY_TILES = [
  { key: 'all-notes', title: 'All Notes', icon: Home, accentClass: 'warm', badge: '3', active: true },
  { key: 'starred', title: 'Starred', icon: Star, accentClass: 'sand' },
  { key: 'unfiled', title: 'Unfiled', icon: FolderOpen, accentClass: 'mint' },
  { key: 'trash', title: 'Trash', icon: Trash2, accentClass: 'sage' },
  { key: 'templates', title: 'Templates', icon: Crown, accentClass: 'olive', badge: 'New' },
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
const SHAPE_VARIANTS = [
  { id: 'line', label: 'Line' },
  { id: 'arrow', label: 'Arrow' },
  { id: 'rectangle', label: 'Rectangle' },
  { id: 'circle', label: 'Circle' },
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
const DOUBLE_TAP_WINDOW_MS = 280;
const DOUBLE_TAP_MAX_DISTANCE = 24;
const MINI_TOOLBAR_HIDE_DELAY_MS = 1400;
const LASSO_HANDLE_HIT_RADIUS_PX = 14;
const LASSO_HANDLE_RENDER_SIZE_PX = 10;
const SHORTCUTS_STORAGE_KEY = 'voxnotes:keyboard-shortcuts';
const ONBOARDING_STORAGE_KEY = 'voxnotes:onboarding-complete';
const DEFAULT_THEME_PRESET = 'forest-calm';
const DEFAULT_PAPER_TEMPLATE = 'ruled';
const DEFAULT_NOTES_LAYOUT = 'tablet';
const DEFAULT_SPLIT_VIEW = false;
const DEFAULT_ROTATION_DEGREES = 0;
const DEFAULT_VIEW_SCALE = 1;
const DEFAULT_SCROLL_DIRECTION = 'vertical';
const DEFAULT_PAGE_BACKGROUND = '#fffdf8';
const DEFAULT_COVER_VARIANT = 'illustration-1';
const DAILY_CARD_STORAGE_KEY = 'voxnotes:daily-card-note';
const UI_PREFS_STORAGE_KEY = 'voxnotes:ui-preferences';
const PAGE_TEMPLATE_OPTIONS = [
  { id: 'blank', label: 'Blank' },
  { id: 'ruled-fine', label: 'Fine Ruled' },
  { id: 'ruled', label: 'Ruled' },
  { id: 'ruled-wide', label: 'Wide Ruled' },
  { id: 'grid', label: 'Grid' },
  { id: 'grid-bold', label: 'Bold Grid' },
  { id: 'dot', label: 'Dot' },
  { id: 'dot-wide', label: 'Wide Dot' },
  { id: 'checklist', label: 'Checklist' },
  { id: 'cornell', label: 'Cornell' },
  { id: 'planner', label: 'Planner' },
  { id: 'music', label: 'Music' },
];
const PAGE_BACKGROUND_SWATCHES = ['#fffdf8', '#fff7ed', '#f5f3ff', '#ecfeff', '#ecfccb'];
const COVER_VARIANTS = [
  { id: 'illustration-1', label: 'Pattern', className: 'cover-pattern' },
  { id: 'illustration-2', label: 'Illustration', className: 'cover-illustration' },
  { id: 'illustration-3', label: 'Text', className: 'cover-text' },
  { id: 'illustration-4', label: 'Coast', className: 'cover-coast' },
  { id: 'illustration-5', label: 'Floral', className: 'cover-floral' },
  { id: 'illustration-6', label: 'Sunset', className: 'cover-sunset' },
  { id: 'illustration-7', label: 'Meadow', className: 'cover-meadow' },
  { id: 'illustration-8', label: 'Forest', className: 'cover-forest' },
  { id: 'illustration-9', label: 'Sea', className: 'cover-sea' },
];
const DEFAULT_SHORTCUT_CONFIG = {
  pen: 'p',
  highlighter: 'h',
  eraser: 'e',
  undo: 'z',
  redo: 'y',
};
const VOICE_SCROLL_STEP_PX = 420;
const VOICE_LANGUAGE_OPTIONS = [
  { id: 'en-US', label: 'English' },
  { id: 'mr-IN', label: 'Marathi' },
];
const VOICE_COMMAND_KEYWORD_MAPS = {
  'en-US': {
    pen: ['pen', 'write', 'lekhani', 'pencil'],
    highlighter: ['highlight', 'highlighter', 'mark', 'marker'],
    eraser: ['eraser', 'erase', 'rubber'],
    lasso: ['lasso', 'select'],
    shape: ['shape', 'draw shape'],
    scrollDown: ['scroll down', 'go down'],
    scrollUp: ['scroll up', 'go up'],
    nextPage: ['next page', 'page next'],
    previousPage: ['previous page', 'prev page', 'back page'],
    jumpToPage: ['go to page', 'jump to page', 'page'],
    remind: ['remind', 'remind me'],
  },
  'mr-IN': {
    pen: ['पेन', 'लिहा', 'लेखनी'],
    highlighter: ['हायलाइट', 'मार्क'],
    eraser: ['इरेजर', 'खोडा', 'पुसा'],
    lasso: ['निवडा', 'लासो'],
    shape: ['आकार'],
    scrollDown: ['खाली स्क्रोल', 'खाली जा'],
    scrollUp: ['वर स्क्रोल', 'वर जा'],
    nextPage: ['पुढचे पान', 'पुढील पान'],
    previousPage: ['मागचे पान', 'मागील पान'],
    jumpToPage: ['पान', 'पानावर जा'],
    remind: ['स्मरण', 'आठवण'],
  },
};
const VOICE_COLOR_MAP = {
  red: '#dc2626',
  orange: '#f59e0b',
  yellow: '#eab308',
  green: '#14b8a6',
  blue: '#3b82f6',
  navy: '#1d4ed8',
  purple: '#9333ea',
  black: '#1f2937',
};

const clampNumber = (value, minimum, maximum) => {
  return Math.min(Math.max(value, minimum), maximum);
};

const transcriptIncludesAnyKeyword = (transcript, keywords) => {
  return keywords.some((keyword) => transcript.includes(keyword));
};

const parseReminderDueAt = (transcript) => {
  const normalized = transcript.toLowerCase();
  const now = new Date();

  if (normalized.includes('tomorrow')) {
    const dueAt = new Date(now);
    dueAt.setDate(dueAt.getDate() + 1);
    dueAt.setHours(9, 0, 0, 0);
    return dueAt;
  }

  if (normalized.includes('उद्या')) {
    const dueAt = new Date(now);
    dueAt.setDate(dueAt.getDate() + 1);
    dueAt.setHours(9, 0, 0, 0);
    return dueAt;
  }

  const relativeMatch = normalized.match(/\bin\s+(\d{1,3})\s*(minute|minutes|hour|hours|day|days)\b/i);
  if (relativeMatch) {
    const amount = Number.parseInt(relativeMatch[1], 10);
    if (!Number.isFinite(amount)) {
      return null;
    }

    const dueAt = new Date(now);
    const unit = relativeMatch[2].toLowerCase();

    if (unit.startsWith('minute')) {
      dueAt.setMinutes(dueAt.getMinutes() + amount);
      return dueAt;
    }

    if (unit.startsWith('hour')) {
      dueAt.setHours(dueAt.getHours() + amount);
      return dueAt;
    }

    dueAt.setDate(dueAt.getDate() + amount);
    return dueAt;
  }

  const marathiMinuteMatch = normalized.match(/(\d{1,3})\s*मिनिट/);
  if (marathiMinuteMatch) {
    const amount = Number.parseInt(marathiMinuteMatch[1], 10);
    if (Number.isFinite(amount)) {
      const dueAt = new Date(now);
      dueAt.setMinutes(dueAt.getMinutes() + amount);
      return dueAt;
    }
  }

  return null;
};

const parseReminderTask = (transcript) => {
  const normalized = transcript.trim();
  const remindMatch = normalized.match(/(?:remind(?:\s+me)?|स्मरण|आठवण)\s+(.+)/i);

  if (!remindMatch?.[1]) {
    return null;
  }

  const taskDescription = remindMatch[1].trim();

  if (!taskDescription) {
    return null;
  }

  return {
    taskDescription,
    dueAt: parseReminderDueAt(transcript),
  };
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

const normalizeShortcutKey = (value) => {
  if (typeof value !== 'string') {
    return '';
  }

  return value.trim().toLowerCase().slice(0, 1);
};

const readStoredShortcutConfig = () => {
  if (typeof window === 'undefined') {
    return DEFAULT_SHORTCUT_CONFIG;
  }

  try {
    const rawValue = window.localStorage.getItem(SHORTCUTS_STORAGE_KEY);

    if (!rawValue) {
      return DEFAULT_SHORTCUT_CONFIG;
    }

    const parsedValue = JSON.parse(rawValue);

    if (!parsedValue || typeof parsedValue !== 'object') {
      return DEFAULT_SHORTCUT_CONFIG;
    }

    return {
      pen: normalizeShortcutKey(parsedValue.pen) || DEFAULT_SHORTCUT_CONFIG.pen,
      highlighter: normalizeShortcutKey(parsedValue.highlighter) || DEFAULT_SHORTCUT_CONFIG.highlighter,
      eraser: normalizeShortcutKey(parsedValue.eraser) || DEFAULT_SHORTCUT_CONFIG.eraser,
      undo: normalizeShortcutKey(parsedValue.undo) || DEFAULT_SHORTCUT_CONFIG.undo,
      redo: normalizeShortcutKey(parsedValue.redo) || DEFAULT_SHORTCUT_CONFIG.redo,
    };
  } catch (error) {
    return DEFAULT_SHORTCUT_CONFIG;
  }
};

const writeStoredShortcutConfig = (shortcutConfig) => {
  if (typeof window === 'undefined') {
    return;
  }

  try {
    window.localStorage.setItem(SHORTCUTS_STORAGE_KEY, JSON.stringify(shortcutConfig));
  } catch (error) {
    // Ignore shortcut persistence failures and keep editing smooth.
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
  shapeType: stroke?.shapeType ?? 'line',
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

const sanitizeStoredBookmarks = (storedBookmarks, pageCount = Number.POSITIVE_INFINITY) => {
  if (!Array.isArray(storedBookmarks)) {
    return [];
  }

  const deduped = new Set();

  storedBookmarks.forEach((pageNumber) => {
    if (!Number.isFinite(pageNumber)) {
      return;
    }

    const normalizedPage = Math.trunc(pageNumber);

    if (normalizedPage < 1 || normalizedPage > pageCount) {
      return;
    }

    deduped.add(normalizedPage);
  });

  return Array.from(deduped).sort((leftPage, rightPage) => leftPage - rightPage);
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

const isPointInsidePolygon = (point, polygonPoints) => {
  if (!Array.isArray(polygonPoints) || polygonPoints.length < 3) {
    return false;
  }

  let inside = false;

  for (let index = 0, previousIndex = polygonPoints.length - 1; index < polygonPoints.length; previousIndex = index, index += 1) {
    const pointA = polygonPoints[index];
    const pointB = polygonPoints[previousIndex];

    const intersects = ((pointA.y > point.y) !== (pointB.y > point.y))
      && (point.x < ((pointB.x - pointA.x) * (point.y - pointA.y)) / ((pointB.y - pointA.y) || Number.EPSILON) + pointA.x);

    if (intersects) {
      inside = !inside;
    }
  }

  return inside;
};

const getBoundsFromPoints = (points) => {
  if (!Array.isArray(points) || !points.length) {
    return null;
  }

  let minimumX = Number.POSITIVE_INFINITY;
  let minimumY = Number.POSITIVE_INFINITY;
  let maximumX = Number.NEGATIVE_INFINITY;
  let maximumY = Number.NEGATIVE_INFINITY;

  points.forEach((point) => {
    minimumX = Math.min(minimumX, point.x);
    minimumY = Math.min(minimumY, point.y);
    maximumX = Math.max(maximumX, point.x);
    maximumY = Math.max(maximumY, point.y);
  });

  return {
    minX: minimumX,
    minY: minimumY,
    maxX: maximumX,
    maxY: maximumY,
    width: maximumX - minimumX,
    height: maximumY - minimumY,
  };
};

const isPointWithinBounds = (point, bounds, padding = 0) => {
  if (!bounds) {
    return false;
  }

  return point.x >= bounds.minX - padding
    && point.x <= bounds.maxX + padding
    && point.y >= bounds.minY - padding
    && point.y <= bounds.maxY + padding;
};

const translateStrokePoints = (stroke, deltaX, deltaY) => {
  return cloneStroke({
    ...stroke,
    points: stroke.points.map((point) => ({
      ...point,
      x: clampNumber(point.x + deltaX, 0, 1),
      y: clampNumber(point.y + deltaY, 0, 1),
    })),
  });
};

const scaleStrokePoints = (stroke, centerPoint, scaleFactor) => {
  return cloneStroke({
    ...stroke,
    points: stroke.points.map((point) => ({
      ...point,
      x: clampNumber(centerPoint.x + (point.x - centerPoint.x) * scaleFactor, 0, 1),
      y: clampNumber(centerPoint.y + (point.y - centerPoint.y) * scaleFactor, 0, 1),
    })),
  });
};

const scaleStrokePointsXY = (stroke, anchorPoint, scaleX, scaleY) => {
  return cloneStroke({
    ...stroke,
    points: stroke.points.map((point) => ({
      ...point,
      x: clampNumber(anchorPoint.x + (point.x - anchorPoint.x) * scaleX, 0, 1),
      y: clampNumber(anchorPoint.y + (point.y - anchorPoint.y) * scaleY, 0, 1),
    })),
  });
};

const getSelectionHandlePoints = (bounds) => {
  if (!bounds) {
    return null;
  }

  return {
    nw: { x: bounds.minX, y: bounds.minY },
    ne: { x: bounds.maxX, y: bounds.minY },
    se: { x: bounds.maxX, y: bounds.maxY },
    sw: { x: bounds.minX, y: bounds.maxY },
  };
};

const getOppositeHandleKey = (handleKey) => {
  if (handleKey === 'nw') {
    return 'se';
  }

  if (handleKey === 'ne') {
    return 'sw';
  }

  if (handleKey === 'se') {
    return 'nw';
  }

  return 'ne';
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
  const renderedPdfPagesRef = useRef({});
  const pdfRenderTasksRef = useRef({});
  const drawingStateRef = useRef({ isDrawing: false, pageNumber: null, stroke: null });
  const originalPdfBytesRef = useRef(null);
  const documentIdRef = useRef('');
  const previousInkToolRef = useRef(DEFAULT_TOOL);
  const recentInkToolsRef = useRef(['pen', 'highlighter']);
  const toolbarDragStateRef = useRef({ active: false, startX: 0, startY: 0, originX: 0, originY: 0 });
  const lastTapRef = useRef({ timestamp: 0, x: 0, y: 0, pointerType: '' });
  const miniToolbarHideTimerRef = useRef(null);
  const lassoSelectionRef = useRef(null);
  const lassoInteractionRef = useRef({ mode: null, pageNumber: null, startPoint: null, baseStrokes: null, selectedIndexes: [] });
  const lassoClipboardRef = useRef([]);
  const speechRecognitionRef = useRef(null);
  const voiceAudioContextRef = useRef(null);
  const voiceAudioStreamRef = useRef(null);
  const voiceAnalyserRef = useRef(null);
  const voiceVolumeFrameRef = useRef(null);
  const transcriptToastTimerRef = useRef(null);

  const [pages, setPages] = useState(createNotebookPages(1));
  const [currentView, setCurrentView] = useState('home');
  const [libraryItems, setLibraryItems] = useState([]);
  const [isLibraryLoading, setIsLibraryLoading] = useState(true);
  const [sourceType, setSourceType] = useState('notebook');
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
  const [shapeType, setShapeType] = useState('line');
  const [shapeSize, setShapeSize] = useState(4);
  const [eraserSize, setEraserSize] = useState(ERASER_SIZES[1]);
  const [penFavorites, setPenFavorites] = useState(() => readStoredFavorites('pen'));
  const [highlighterFavorites, setHighlighterFavorites] = useState(() => readStoredFavorites('highlighter'));
  const [toolbarPosition, setToolbarPosition] = useState(() => readStoredToolbarPosition());
  const [isToolbarCollapsed, setIsToolbarCollapsed] = useState(() => readStoredToolbarCollapsed());
  const [autoSelectPreviousTool, setAutoSelectPreviousTool] = useState(false);
  const [eraseEntireStroke, setEraseEntireStroke] = useState(true);
  const [eraseHighlighterOnly, setEraseHighlighterOnly] = useState(false);
  const [erasePenOnly, setErasePenOnly] = useState(false);
  const [annotationVersion, setAnnotationVersion] = useState(0);
  const [shortcutConfig, setShortcutConfig] = useState(() => readStoredShortcutConfig());
  const [miniToolbar, setMiniToolbar] = useState({ visible: false, x: 0, y: 0 });
  const [bookmarkedPages, setBookmarkedPages] = useState([]);
  const [lassoPath, setLassoPath] = useState(null);
  const [lassoSelection, setLassoSelection] = useState(null);
  const [viewScale, setViewScale] = useState(DEFAULT_VIEW_SCALE);
  const [rotationDegrees, setRotationDegrees] = useState(DEFAULT_ROTATION_DEGREES);
  const [splitViewEnabled, setSplitViewEnabled] = useState(DEFAULT_SPLIT_VIEW);
  const [paperTemplate, setPaperTemplate] = useState(DEFAULT_PAPER_TEMPLATE);
  const [notesLayout, setNotesLayout] = useState(DEFAULT_NOTES_LAYOUT);
  const [scrollDirection, setScrollDirection] = useState(DEFAULT_SCROLL_DIRECTION);
  const [pageBackgroundColor, setPageBackgroundColor] = useState(DEFAULT_PAGE_BACKGROUND);
  const [coverVariant, setCoverVariant] = useState(DEFAULT_COVER_VARIANT);
  const [showNotebookSettings, setShowNotebookSettings] = useState(false);
  const [applyTemplateToAllPages, setApplyTemplateToAllPages] = useState(true);
  const [themePreset, setThemePreset] = useState(DEFAULT_THEME_PRESET);
  const [librarySearchQuery, setLibrarySearchQuery] = useState('');
  const [libraryTypeFilter, setLibraryTypeFilter] = useState('all');
  const [libraryDateFilter, setLibraryDateFilter] = useState('all-time');
  const [librarySmartView, setLibrarySmartView] = useState('all');
  const [todoInput, setTodoInput] = useState('');
  const [todoItems, setTodoItems] = useState([]);
  const [focusSecondsLeft, setFocusSecondsLeft] = useState(25 * 60);
  const [focusTimerRunning, setFocusTimerRunning] = useState(false);
  const [focusModeEnabled, setFocusModeEnabled] = useState(false);
  const [dailyDashboardNote, setDailyDashboardNote] = useState('Plan your top 3 outcomes for today.');
  const [showOnboardingTour, setShowOnboardingTour] = useState(false);
  const [uiScale, setUiScale] = useState(1);
  const [largeHitTargets, setLargeHitTargets] = useState(false);
  const [highContrastMode, setHighContrastMode] = useState(false);
  const [toastItems, setToastItems] = useState([]);
  const [isVoiceListening, setIsVoiceListening] = useState(false);
  const [lastVoiceTranscript, setLastVoiceTranscript] = useState('');
  const [voiceLanguage, setVoiceLanguage] = useState('en-US');
  const [voiceLevel, setVoiceLevel] = useState(0);
  const [interimVoiceTranscript, setInterimVoiceTranscript] = useState('');
  const [voiceTranscriptToast, setVoiceTranscriptToast] = useState('');

  const pageCount = pages.length;
  const hasPdf = sourceType === 'pdf';
  const selectedPenPreset = PEN_VARIANTS.find((preset) => preset.id === penVariant) ?? PEN_VARIANTS[1];
  const selectedHighlighterPreset = HIGHLIGHTER_VARIANTS.find((preset) => preset.id === highlighterVariant) ?? HIGHLIGHTER_VARIANTS[0];
  const activeVoiceKeywordMap = VOICE_COMMAND_KEYWORD_MAPS[voiceLanguage] ?? VOICE_COMMAND_KEYWORD_MAPS['en-US'];

  const filteredLibraryItems = useMemo(() => {
    const now = Date.now();

    return libraryItems.filter((item) => {
      const name = (item.documentName || '').toLowerCase();
      const sourceTypeName = (item.sourceType || '').toLowerCase();
      const query = librarySearchQuery.trim().toLowerCase();

      if (query && !name.includes(query) && !sourceTypeName.includes(query)) {
        return false;
      }

      if (libraryTypeFilter !== 'all' && item.sourceType !== libraryTypeFilter) {
        return false;
      }

      const updatedAtMs = new Date(item.updatedAt || 0).getTime();
      const ageMs = now - updatedAtMs;

      if (libraryDateFilter === 'today' && ageMs > 24 * 60 * 60 * 1000) {
        return false;
      }

      if (libraryDateFilter === 'week' && ageMs > 7 * 24 * 60 * 60 * 1000) {
        return false;
      }

      if (librarySmartView === 'recently-annotated' && (item.strokeCount ?? 0) < 3) {
        return false;
      }

      if (librarySmartView === 'needs-review' && (item.pageCount ?? 0) > 0 && (item.strokeCount ?? 0) === 0) {
        return true;
      }

      if (librarySmartView === 'needs-review' && (item.strokeCount ?? 0) > 0) {
        return false;
      }

      return true;
    });
  }, [libraryDateFilter, libraryItems, librarySearchQuery, librarySmartView, libraryTypeFilter]);

  const clearLassoSelection = useCallback(() => {
    lassoSelectionRef.current = null;
    lassoInteractionRef.current = { mode: null, pageNumber: null, startPoint: null, baseStrokes: null, selectedIndexes: [] };
    setLassoPath(null);
    setLassoSelection(null);
  }, []);

  const computeSelectionBoundsForIndexes = useCallback((pageStrokes, strokeIndexes) => {
    const selectedPoints = strokeIndexes.flatMap((strokeIndex) => pageStrokes[strokeIndex]?.points ?? []);
    return getBoundsFromPoints(selectedPoints);
  }, []);

  const drawShapePrimitive = useCallback((context, shapeKind, startPoint, endPoint, strokeWidth) => {
    const deltaX = endPoint.x - startPoint.x;
    const deltaY = endPoint.y - startPoint.y;

    context.lineWidth = strokeWidth;

    if (shapeKind === 'rectangle') {
      context.beginPath();
      context.rect(startPoint.x, startPoint.y, deltaX, deltaY);
      context.stroke();
      context.closePath();
      return;
    }

    if (shapeKind === 'circle') {
      const centerX = startPoint.x + deltaX / 2;
      const centerY = startPoint.y + deltaY / 2;
      const radiusX = Math.abs(deltaX) / 2;
      const radiusY = Math.abs(deltaY) / 2;

      context.beginPath();
      context.ellipse(centerX, centerY, Math.max(radiusX, 0.5), Math.max(radiusY, 0.5), 0, 0, Math.PI * 2);
      context.stroke();
      context.closePath();
      return;
    }

    context.beginPath();
    context.moveTo(startPoint.x, startPoint.y);
    context.lineTo(endPoint.x, endPoint.y);
    context.stroke();
    context.closePath();

    if (shapeKind === 'arrow') {
      const angle = Math.atan2(deltaY, deltaX);
      const arrowLength = Math.max(10, strokeWidth * 3.2);
      const spread = Math.PI / 7;

      context.beginPath();
      context.moveTo(endPoint.x, endPoint.y);
      context.lineTo(
        endPoint.x - arrowLength * Math.cos(angle - spread),
        endPoint.y - arrowLength * Math.sin(angle - spread),
      );
      context.moveTo(endPoint.x, endPoint.y);
      context.lineTo(
        endPoint.x - arrowLength * Math.cos(angle + spread),
        endPoint.y - arrowLength * Math.sin(angle + spread),
      );
      context.stroke();
      context.closePath();
    }
  }, []);

  const trackInkTool = useCallback((toolName) => {
    if (toolName !== 'pen' && toolName !== 'highlighter') {
      return;
    }

    const [latestTool = 'pen', previousTool = latestTool === 'pen' ? 'highlighter' : 'pen'] = recentInkToolsRef.current;

    if (latestTool === toolName) {
      return;
    }

    recentInkToolsRef.current = [toolName, latestTool || previousTool];
  }, []);

  const activateTool = useCallback((toolName, openPanel = true) => {
    if (toolName !== 'eraser') {
      previousInkToolRef.current = toolName;
      trackInkTool(toolName);
    }

    setActiveTool(toolName);
    if (openPanel) {
      setToolPanel((currentPanel) => currentPanel === toolName ? null : toolName);
      return;
    }

    setToolPanel(null);
  }, [trackInkTool]);

  const swapRecentInkTools = useCallback(() => {
    const [latestTool = 'pen', previousTool = latestTool === 'pen' ? 'highlighter' : 'pen'] = recentInkToolsRef.current;
    const targetTool = activeTool === latestTool ? previousTool : latestTool;

    activateTool(targetTool, true);
    setStatusMessage(`Quick switch: ${targetTool} ready.`);
  }, [activateTool, activeTool]);

  const showMiniToolbarAtPointer = useCallback((event) => {
    if (miniToolbarHideTimerRef.current) {
      window.clearTimeout(miniToolbarHideTimerRef.current);
      miniToolbarHideTimerRef.current = null;
    }

    setMiniToolbar({
      visible: true,
      x: event.clientX + 16,
      y: event.clientY + 16,
    });
  }, []);

  const scheduleMiniToolbarHide = useCallback(() => {
    if (miniToolbarHideTimerRef.current) {
      window.clearTimeout(miniToolbarHideTimerRef.current);
    }

    miniToolbarHideTimerRef.current = window.setTimeout(() => {
      setMiniToolbar((currentValue) => ({ ...currentValue, visible: false }));
      miniToolbarHideTimerRef.current = null;
    }, MINI_TOOLBAR_HIDE_DELAY_MS);
  }, []);

  const handleToolSelect = useCallback((toolName) => {
    activateTool(toolName, true);
  }, [activateTool]);

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
    activateTool(toolName, true);

    if (toolName === 'pen') {
      setPenVariant(favorite.variant);
      setPenSize(favorite.size);
      setStatusMessage(`Applied pen favorite ${favorite.variant}.`);
      return;
    }

    setHighlighterVariant(favorite.variant);
    setHighlighterSize(favorite.size);
    setStatusMessage(`Applied highlighter favorite ${favorite.variant}.`);
  }, [activateTool]);

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
    nextBookmarks = bookmarkedPages,
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
        bookmarks: nextBookmarks,
      });
      setStorageLabel('IndexedDB primary storage');
      setLastSavedAt(formatSavedTime(updatedAt));
      refreshLibrary();
    } catch (error) {
      setStorageLabel('Unable to persist session');
      setStatusMessage('IndexedDB is unavailable in this browser session.');
    }
  }, [bookmarkedPages, documentName, pages, refreshLibrary, sourceType]);

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

    if (stroke.tool === 'shape') {
      const startPoint = stroke.points[0];
      const endPoint = stroke.points[stroke.points.length - 1] ?? startPoint;

      drawShapePrimitive(
        context,
        stroke.shapeType ?? 'line',
        { x: startPoint.x * width, y: startPoint.y * height },
        { x: endPoint.x * width, y: endPoint.y * height },
        Math.max(stroke.width, 1),
      );

      context.globalAlpha = 1;
      return;
    }

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
  }, [applyStrokeStyle, drawShapePrimitive, getPageSize]);

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

    if (lassoSelectionRef.current?.pageNumber === pageNumber && lassoSelectionRef.current.bounds) {
      const bounds = lassoSelectionRef.current.bounds;
      const { width, height } = getPageSize(pageNumber);
      const handles = getSelectionHandlePoints(bounds);

      context.save();
      context.setLineDash([8, 6]);
      context.lineWidth = 1.5;
      context.strokeStyle = '#f59e0b';
      context.globalAlpha = 0.9;
      context.strokeRect(
        bounds.minX * width,
        bounds.minY * height,
        Math.max(bounds.width * width, 1),
        Math.max(bounds.height * height, 1),
      );

      if (handles) {
        const halfHandleSize = LASSO_HANDLE_RENDER_SIZE_PX / 2;

        context.setLineDash([]);
        context.fillStyle = '#f8fafc';
        context.strokeStyle = '#0f172a';
        Object.values(handles).forEach((handlePoint) => {
          const handleX = handlePoint.x * width - halfHandleSize;
          const handleY = handlePoint.y * height - halfHandleSize;
          context.fillRect(handleX, handleY, LASSO_HANDLE_RENDER_SIZE_PX, LASSO_HANDLE_RENDER_SIZE_PX);
          context.strokeRect(handleX, handleY, LASSO_HANDLE_RENDER_SIZE_PX, LASSO_HANDLE_RENDER_SIZE_PX);
        });
      }
      context.restore();
    }

    if (lassoPath?.pageNumber === pageNumber && Array.isArray(lassoPath.points) && lassoPath.points.length > 1) {
      const { width, height } = getPageSize(pageNumber);

      context.save();
      context.setLineDash([6, 4]);
      context.lineWidth = 1.5;
      context.strokeStyle = '#38bdf8';
      context.globalAlpha = 0.9;
      context.beginPath();
      context.moveTo(lassoPath.points[0].x * width, lassoPath.points[0].y * height);
      for (let index = 1; index < lassoPath.points.length; index += 1) {
        context.lineTo(lassoPath.points[index].x * width, lassoPath.points[index].y * height);
      }
      context.stroke();
      context.closePath();
      context.restore();
    }
  }, [clearCanvas, drawStroke, getPageSize, lassoPath]);

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

  const cancelPdfRenderTask = useCallback(async (pageNumber) => {
    const activeRenderEntry = pdfRenderTasksRef.current[pageNumber];

    if (!activeRenderEntry) {
      return;
    }

    activeRenderEntry.task.cancel();

    try {
      await activeRenderEntry.promise;
    } catch (error) {
      // Cancellation is expected during rapid navigation and rerenders.
    }
  }, []);

  const renderPdfPage = useCallback(async ({ pageNumber, width, height }) => {
    const pdf = pdfDocumentRef.current;
    const canvas = pdfCanvasRefs.current[pageNumber];

    if (!pdf || !canvas) {
      return;
    }

    const signature = `${width}x${height}`;
    const activeRenderEntry = pdfRenderTasksRef.current[pageNumber];

    if (activeRenderEntry?.signature === signature) {
      await activeRenderEntry.promise;
      return;
    }

    if (!activeRenderEntry && renderedPdfPagesRef.current[pageNumber] === signature) {
      return;
    }

    if (activeRenderEntry) {
      await cancelPdfRenderTask(pageNumber);
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

    const renderTask = page.render({ canvasContext: context, viewport });
    const renderEntry = {
      signature,
      task: renderTask,
      promise: null,
    };

    renderEntry.promise = renderTask.promise
      .then(() => {
        if (pdfRenderTasksRef.current[pageNumber] === renderEntry) {
          renderedPdfPagesRef.current[pageNumber] = signature;
        }
      })
      .catch((error) => {
        if (error?.name !== 'RenderingCancelledException') {
          throw error;
        }
      })
      .finally(() => {
        if (pdfRenderTasksRef.current[pageNumber] === renderEntry) {
          delete pdfRenderTasksRef.current[pageNumber];
        }
      });

    pdfRenderTasksRef.current[pageNumber] = renderEntry;

    await renderEntry.promise;
  }, [cancelPdfRenderTask]);

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
    const nextBookmarks = sanitizeStoredBookmarks(record?.bookmarks, nextPages.length);

    pdfDocumentRef.current = null;
    originalPdfBytesRef.current = null;
    documentIdRef.current = documentId;
    annotationsRef.current = sanitizeStoredAnnotations(record?.annotations ?? {});
    Object.values(pdfRenderTasksRef.current).forEach((renderEntry) => {
      renderEntry?.task?.cancel();
    });
    pdfRenderTasksRef.current = {};
    renderedPdfPagesRef.current = {};
    lassoSelectionRef.current = null;
    lassoInteractionRef.current = { mode: null, pageNumber: null, startPoint: null, baseStrokes: null, selectedIndexes: [] };
    setPages(nextPages);
    setSourceType('notebook');
    setDocumentName(notebookName);
    setLassoPath(null);
    setLassoSelection(null);
    setBookmarkedPages(nextBookmarks);
    setActivePageNumber(1);
    setPageJumpValue('1');
    setLastSavedAt(formatSavedTime(record?.updatedAt));
    setStorageLabel('IndexedDB primary storage');
    setStatusMessage(`Opened ${notebookName}. Start writing or upload a PDF.`);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setCurrentView('editor');
    updateRenderWindow(1);
    if (fresh) {
      persistSession({}, nextPages, 'notebook', notebookName, documentId, null, []);
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

      const nextBookmarks = sanitizeStoredBookmarks(record?.bookmarks, nextPages.length);

      pdfDocumentRef.current = pdf;
      originalPdfBytesRef.current = pdfBytes;
      documentIdRef.current = documentId;
      annotationsRef.current = sanitizeStoredAnnotations(record?.annotations ?? {});
      Object.values(pdfRenderTasksRef.current).forEach((renderEntry) => {
        renderEntry?.task?.cancel();
      });
      pdfRenderTasksRef.current = {};
      renderedPdfPagesRef.current = {};
      lassoSelectionRef.current = null;
      lassoInteractionRef.current = { mode: null, pageNumber: null, startPoint: null, baseStrokes: null, selectedIndexes: [] };
      setPages(nextPages);
      setSourceType('pdf');
      setDocumentName(fileName);
      setLassoPath(null);
      setLassoSelection(null);
      setBookmarkedPages(nextBookmarks);
      setActivePageNumber(1);
      setPageJumpValue('1');
      setLastSavedAt(formatSavedTime(record?.updatedAt));
      setStorageLabel('IndexedDB primary storage');
      setStatusMessage('PDF ready. Scroll, jump, annotate, or export.');
      setAnnotationVersion((currentValue) => currentValue + 1);
      setCurrentView('editor');
      updateRenderWindow(1);
      persistSession(record?.annotations ?? {}, nextPages, 'pdf', fileName, documentId, pdfBytes, nextBookmarks);
    } catch (error) {
      setStatusMessage('Unable to render that PDF. Try another file.');
    }
  }, [buildPdfPages, persistSession, updateRenderWindow]);

  useEffect(() => {
    refreshLibrary();
  }, [refreshLibrary]);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    const onboardingComplete = window.localStorage.getItem(ONBOARDING_STORAGE_KEY) === 'true';
    const storedDailyCard = window.localStorage.getItem(DAILY_CARD_STORAGE_KEY);
    const storedPrefsRaw = window.localStorage.getItem(UI_PREFS_STORAGE_KEY);

    if (storedDailyCard) {
      setDailyDashboardNote(storedDailyCard);
    }

    if (storedPrefsRaw) {
      try {
        const storedPrefs = JSON.parse(storedPrefsRaw);

        if (typeof storedPrefs.paperTemplate === 'string') {
          setPaperTemplate(storedPrefs.paperTemplate);
        }

        if (storedPrefs.notesLayout === 'classic' || storedPrefs.notesLayout === 'tablet') {
          setNotesLayout(storedPrefs.notesLayout);
        }

        if (storedPrefs.scrollDirection === 'horizontal' || storedPrefs.scrollDirection === 'vertical') {
          setScrollDirection(storedPrefs.scrollDirection);
        }

        if (typeof storedPrefs.pageBackgroundColor === 'string') {
          setPageBackgroundColor(storedPrefs.pageBackgroundColor);
        }

        if (typeof storedPrefs.coverVariant === 'string') {
          setCoverVariant(storedPrefs.coverVariant);
        }

        if (typeof storedPrefs.applyTemplateToAllPages === 'boolean') {
          setApplyTemplateToAllPages(storedPrefs.applyTemplateToAllPages);
        }

        if (Number.isFinite(storedPrefs.uiScale)) {
          setUiScale(clampNumber(storedPrefs.uiScale, 0.85, 1.25));
        }

        if (typeof storedPrefs.largeHitTargets === 'boolean') {
          setLargeHitTargets(storedPrefs.largeHitTargets);
        }

        if (typeof storedPrefs.highContrastMode === 'boolean') {
          setHighContrastMode(storedPrefs.highContrastMode);
        }
      } catch (error) {
        // Ignore malformed preferences and keep defaults.
      }
    }

    setShowOnboardingTour(!onboardingComplete);
  }, []);

  useEffect(() => {
    if (!focusTimerRunning) {
      return undefined;
    }

    const timerId = window.setInterval(() => {
      setFocusSecondsLeft((currentValue) => {
        if (currentValue <= 1) {
          window.clearInterval(timerId);
          setFocusTimerRunning(false);
          setStatusMessage('Focus timer completed. Nice work.');
          return 0;
        }

        return currentValue - 1;
      });
    }, 1000);

    return () => {
      window.clearInterval(timerId);
    };
  }, [focusTimerRunning]);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    window.localStorage.setItem(DAILY_CARD_STORAGE_KEY, dailyDashboardNote);
  }, [dailyDashboardNote]);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    window.localStorage.setItem(UI_PREFS_STORAGE_KEY, JSON.stringify({
      themePreset,
      paperTemplate,
      notesLayout,
      scrollDirection,
      pageBackgroundColor,
      coverVariant,
      applyTemplateToAllPages,
      uiScale,
      largeHitTargets,
      highContrastMode,
    }));
  }, [applyTemplateToAllPages, coverVariant, highContrastMode, largeHitTargets, notesLayout, pageBackgroundColor, paperTemplate, scrollDirection, themePreset, uiScale]);

  useEffect(() => {
    if (!statusMessage) {
      return;
    }

    const now = Date.now();
    const nextToast = { id: `${now}-${Math.random().toString(36).slice(2, 7)}`, message: statusMessage };

    setToastItems((currentToasts) => [...currentToasts, nextToast].slice(-4));

    const dismissTimerId = window.setTimeout(() => {
      setToastItems((currentToasts) => currentToasts.filter((toast) => toast.id !== nextToast.id));
    }, 2800);

    return () => {
      window.clearTimeout(dismissTimerId);
    };
  }, [statusMessage]);

  useEffect(() => {
    lassoSelectionRef.current = lassoSelection;
  }, [lassoSelection]);

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

      cancelPdfRenderTask(pageNumber);
      releaseCanvasResources(inkCanvasRefs.current[pageNumber], width, height);
      releaseCanvasResources(pdfCanvasRefs.current[pageNumber], width, height);
      delete renderedPdfPagesRef.current[pageNumber];
    });
  }, [cancelPdfRenderTask, pages, releaseCanvasResources, renderWindowPageNumbers, syncInkCanvas]);

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
    return () => {
      if (miniToolbarHideTimerRef.current) {
        window.clearTimeout(miniToolbarHideTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    writeStoredShortcutConfig(shortcutConfig);
  }, [shortcutConfig]);

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
    if (lassoSelectionRef.current?.pageNumber === pageNumber) {
      lassoSelectionRef.current = null;
      setLassoSelection(null);
      setLassoPath(null);
    }
    redrawInkPage(pageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Erased ink on page ${pageNumber}.`);
    if (autoSelectPreviousTool) {
      activateTool(previousInkToolRef.current, false);
    }
  }, [activateTool, autoSelectPreviousTool, buildStrokeSegments, ensurePageAnnotations, eraseEntireStroke, eraseHighlighterOnly, erasePenOnly, eraserSize, getPageSize, persistSession, redrawInkPage]);

  const handlePointerDown = useCallback((pageNumber) => (event) => {
    if (event.pointerType === 'mouse' && event.button !== 0) {
      return;
    }

    showMiniToolbarAtPointer(event);

    if (event.pointerType !== 'mouse') {
      const now = Date.now();
      const previousTap = lastTapRef.current;
      const tapDistance = Math.hypot(event.clientX - previousTap.x, event.clientY - previousTap.y);
      const isDoubleTap = previousTap.pointerType === event.pointerType
        && now - previousTap.timestamp <= DOUBLE_TAP_WINDOW_MS
        && tapDistance <= DOUBLE_TAP_MAX_DISTANCE;

      lastTapRef.current = {
        timestamp: now,
        x: event.clientX,
        y: event.clientY,
        pointerType: event.pointerType,
      };

      if (isDoubleTap) {
        swapRecentInkTools();
        drawingStateRef.current = { isDrawing: false, pageNumber: null, stroke: null };
        return;
      }
    }

    const pageLayout = pageLayoutMapRef.current[pageNumber];
    const canvas = inkCanvasRefs.current[pageNumber];

    if (!pageLayout || !canvas) {
      return;
    }

    const point = getPoint(pageNumber, event);
    const normalizedPoint = getNormalizedPoint(pageLayout, point, event);

    if (activeTool === 'lasso') {
      const pageAnnotations = ensurePageAnnotations(pageNumber);
      const existingSelection = lassoSelectionRef.current;
      const existingHandles = getSelectionHandlePoints(existingSelection?.bounds);
      const hitHandleKey = existingHandles
        ? Object.entries(existingHandles).find(([, handlePoint]) => {
          const distanceFromHandle = Math.hypot(
            (handlePoint.x - normalizedPoint.x) * pageLayout.width,
            (handlePoint.y - normalizedPoint.y) * pageLayout.height,
          );

          return distanceFromHandle <= LASSO_HANDLE_HIT_RADIUS_PX;
        })?.[0] ?? null
        : null;
      const canMoveSelection = existingSelection?.pageNumber === pageNumber
        && Array.isArray(existingSelection.strokeIndexes)
        && existingSelection.strokeIndexes.length > 0
        && isPointWithinBounds(normalizedPoint, existingSelection.bounds, 0.012);

      if (hitHandleKey && existingSelection?.pageNumber === pageNumber && existingHandles) {
        lassoInteractionRef.current = {
          mode: 'resize',
          pageNumber,
          startPoint: normalizedPoint,
          baseStrokes: pageAnnotations.strokes.map((stroke) => cloneStroke(stroke)),
          selectedIndexes: [...existingSelection.strokeIndexes],
          baseBounds: existingSelection.bounds,
          handleKey: hitHandleKey,
          anchorPoint: existingHandles[getOppositeHandleKey(hitHandleKey)],
          baseCorner: existingHandles[hitHandleKey],
          pathPoints: null,
        };
        setStatusMessage('Resizing selected strokes...');
      } else if (canMoveSelection) {
        lassoInteractionRef.current = {
          mode: 'move',
          pageNumber,
          startPoint: normalizedPoint,
          baseStrokes: pageAnnotations.strokes.map((stroke) => cloneStroke(stroke)),
          selectedIndexes: [...existingSelection.strokeIndexes],
          pathPoints: null,
        };
        setStatusMessage('Moving selected strokes...');
      } else {
        const initialPath = [normalizedPoint];

        lassoSelectionRef.current = null;
        setLassoSelection(null);
        setLassoPath({ pageNumber, points: initialPath });
        lassoInteractionRef.current = {
          mode: 'lasso',
          pageNumber,
          startPoint: normalizedPoint,
          baseStrokes: null,
          selectedIndexes: [],
          pathPoints: initialPath,
        };
        setStatusMessage('Draw a loop around strokes to select them.');
      }

      redrawInkPage(pageNumber);
      setActivePageNumber(pageNumber);
      return;
    }

    if (activeTool === 'eraser') {
      eraseStrokeAtPoint(pageNumber, normalizedPoint);
      scheduleMiniToolbarHide();
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
      variant: activeTool === 'highlighter' ? selectedHighlighterPreset.id : activeTool === 'shape' ? 'shape' : selectedPenPreset.id,
      shapeType,
      color: inkColor,
      width: activeTool === 'shape'
        ? Math.max(shapeSize, 1)
        : activeTool === 'highlighter'
        ? Math.max(highlighterSize * selectedHighlighterPreset.widthScale, TOOL_PRESETS.highlighter.width * 0.8)
        : Math.max(penSize * selectedPenPreset.widthScale, 1.5),
      opacity: activeTool === 'shape' ? 1 : activeTool === 'highlighter' ? selectedHighlighterPreset.opacity : selectedPenPreset.opacity,
      pressureEnabled: activeTool !== 'shape' && event.pointerType === 'pen' && Number.isFinite(event.pressure) && event.pressure > 0,
    };

    drawingStateRef.current = { isDrawing: true, pageNumber, stroke };
    canvas.setPointerCapture?.(event.pointerId);
    redrawInkPage(pageNumber);
    drawStroke(context, pageNumber, stroke);
    setActivePageNumber(pageNumber);
  }, [activeTool, drawStroke, ensurePageAnnotations, eraseStrokeAtPoint, getNormalizedPoint, getPoint, highlighterSize, inkColor, penSize, redrawInkPage, renderWindowPageNumbers, scheduleMiniToolbarHide, selectedHighlighterPreset, selectedPenPreset, shapeSize, shapeType, showMiniToolbarAtPointer, swapRecentInkTools, syncInkCanvas, updateRenderWindow]);

  const handlePointerMove = useCallback((pageNumber) => (event) => {
    const interaction = lassoInteractionRef.current;
    const pageLayout = pageLayoutMapRef.current[pageNumber];

    if (!pageLayout) {
      return;
    }

    const point = getPoint(pageNumber, event);
    const normalizedPoint = getNormalizedPoint(pageLayout, point, event);

    if (activeTool === 'lasso' && interaction.pageNumber === pageNumber && interaction.mode) {
      if (interaction.mode === 'lasso') {
        const nextPath = [...(interaction.pathPoints ?? []), normalizedPoint];
        interaction.pathPoints = nextPath;
        setLassoPath({ pageNumber, points: nextPath });
        redrawInkPage(pageNumber);
        return;
      }

      if (interaction.mode === 'move' && interaction.startPoint && Array.isArray(interaction.baseStrokes)) {
        const deltaX = normalizedPoint.x - interaction.startPoint.x;
        const deltaY = normalizedPoint.y - interaction.startPoint.y;
        const selectedIndexes = new Set(interaction.selectedIndexes);
        const nextStrokes = interaction.baseStrokes.map((stroke, strokeIndex) => {
          if (!selectedIndexes.has(strokeIndex)) {
            return cloneStroke(stroke);
          }

          return translateStrokePoints(stroke, deltaX, deltaY);
        });

        const nextAnnotations = {
          ...annotationsRef.current,
          [pageNumber]: {
            ...ensurePageAnnotations(pageNumber),
            strokes: nextStrokes,
          },
        };

        annotationsRef.current = nextAnnotations;
        const nextBounds = computeSelectionBoundsForIndexes(nextStrokes, interaction.selectedIndexes);
        const nextSelection = {
          pageNumber,
          strokeIndexes: [...interaction.selectedIndexes],
          bounds: nextBounds,
        };

        lassoSelectionRef.current = nextSelection;
        setLassoSelection(nextSelection);
        redrawInkPage(pageNumber);
        return;
      }

      if (
        interaction.mode === 'resize'
        && interaction.anchorPoint
        && interaction.baseCorner
        && Array.isArray(interaction.baseStrokes)
      ) {
        const baseDeltaX = interaction.baseCorner.x - interaction.anchorPoint.x;
        const baseDeltaY = interaction.baseCorner.y - interaction.anchorPoint.y;
        const rawScaleX = Math.abs(baseDeltaX) < Number.EPSILON
          ? 1
          : (normalizedPoint.x - interaction.anchorPoint.x) / baseDeltaX;
        const rawScaleY = Math.abs(baseDeltaY) < Number.EPSILON
          ? 1
          : (normalizedPoint.y - interaction.anchorPoint.y) / baseDeltaY;
        const scaleX = clampNumber(Math.abs(rawScaleX), 0.1, 4);
        const scaleY = clampNumber(Math.abs(rawScaleY), 0.1, 4);
        const selectedIndexes = new Set(interaction.selectedIndexes);
        const nextStrokes = interaction.baseStrokes.map((stroke, strokeIndex) => {
          if (!selectedIndexes.has(strokeIndex)) {
            return cloneStroke(stroke);
          }

          return scaleStrokePointsXY(stroke, interaction.anchorPoint, scaleX, scaleY);
        });
        const nextAnnotations = {
          ...annotationsRef.current,
          [pageNumber]: {
            ...ensurePageAnnotations(pageNumber),
            strokes: nextStrokes,
          },
        };

        annotationsRef.current = nextAnnotations;
        const nextBounds = computeSelectionBoundsForIndexes(nextStrokes, interaction.selectedIndexes);
        const nextSelection = {
          pageNumber,
          strokeIndexes: [...interaction.selectedIndexes],
          bounds: nextBounds,
        };

        lassoSelectionRef.current = nextSelection;
        setLassoSelection(nextSelection);
        redrawInkPage(pageNumber);
        return;
      }
    }

    if (!drawingStateRef.current.isDrawing || drawingStateRef.current.pageNumber !== pageNumber) {
      return;
    }

    const context = inkCanvasRefs.current[pageNumber]?.getContext('2d');

    if (!context) {
      return;
    }

    if (drawingStateRef.current.stroke.tool === 'shape') {
      const originPoint = drawingStateRef.current.stroke.points[0];
      drawingStateRef.current.stroke.points = [originPoint, normalizedPoint];
    } else {
      drawingStateRef.current.stroke.points.push(normalizedPoint);
      drawingStateRef.current.stroke.pressureEnabled = drawingStateRef.current.stroke.pressureEnabled
        || (event.pointerType === 'pen' && Number.isFinite(event.pressure) && event.pressure > 0);
    }

    setMiniToolbar((currentValue) => ({
      ...currentValue,
      visible: true,
      x: event.clientX + 16,
      y: event.clientY + 16,
    }));

    redrawInkPage(pageNumber);
    drawStroke(context, pageNumber, drawingStateRef.current.stroke);
  }, [activeTool, computeSelectionBoundsForIndexes, drawStroke, ensurePageAnnotations, getNormalizedPoint, getPoint, redrawInkPage]);

  const handlePointerUp = useCallback((pageNumber) => (event) => {
    const interaction = lassoInteractionRef.current;

    if (activeTool === 'lasso' && interaction.pageNumber === pageNumber && interaction.mode) {
      if (interaction.mode === 'lasso') {
        const polygonPoints = interaction.pathPoints ?? [];
        const pageStrokes = ensurePageAnnotations(pageNumber).strokes;
        const selectedIndexes = polygonPoints.length < 3
          ? []
          : pageStrokes.reduce((nextIndexes, stroke, strokeIndex) => {
            const isSelected = stroke.points.some((strokePoint) => isPointInsidePolygon(strokePoint, polygonPoints));

            if (isSelected) {
              nextIndexes.push(strokeIndex);
            }

            return nextIndexes;
          }, []);

        if (selectedIndexes.length) {
          const nextSelection = {
            pageNumber,
            strokeIndexes: selectedIndexes,
            bounds: computeSelectionBoundsForIndexes(pageStrokes, selectedIndexes),
          };

          lassoSelectionRef.current = nextSelection;
          setLassoSelection(nextSelection);
          setStatusMessage(`Selected ${selectedIndexes.length} stroke${selectedIndexes.length > 1 ? 's' : ''}. Drag to move or use resize.`);
        } else {
          lassoSelectionRef.current = null;
          setLassoSelection(null);
          setStatusMessage('No strokes found inside that lasso.');
        }

        setLassoPath(null);
        lassoInteractionRef.current = { mode: null, pageNumber: null, startPoint: null, baseStrokes: null, selectedIndexes: [] };
        redrawInkPage(pageNumber);
        return;
      }

      if (interaction.mode === 'move') {
        lassoInteractionRef.current = { mode: null, pageNumber: null, startPoint: null, baseStrokes: null, selectedIndexes: [] };
        persistSession(annotationsRef.current);
        setAnnotationVersion((currentValue) => currentValue + 1);
        scheduleMiniToolbarHide();
        setStatusMessage('Moved selected strokes.');
        return;
      }

      if (interaction.mode === 'resize') {
        lassoInteractionRef.current = { mode: null, pageNumber: null, startPoint: null, baseStrokes: null, selectedIndexes: [] };
        persistSession(annotationsRef.current);
        setAnnotationVersion((currentValue) => currentValue + 1);
        scheduleMiniToolbarHide();
        setStatusMessage('Resized selected strokes from corner handle.');
        return;
      }
    }

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
    scheduleMiniToolbarHide();
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Stroke saved on page ${pageNumber}.`);
  }, [activeTool, computeSelectionBoundsForIndexes, ensurePageAnnotations, persistSession, redrawInkPage, scheduleMiniToolbarHide]);

  const handleShortcutChange = useCallback((actionName, value) => {
    const normalizedKey = normalizeShortcutKey(value);

    if (!normalizedKey && value !== '') {
      return;
    }

    setShortcutConfig((currentConfig) => ({
      ...currentConfig,
      [actionName]: normalizedKey,
    }));
  }, []);

  const resetShortcutConfig = useCallback(() => {
    setShortcutConfig(DEFAULT_SHORTCUT_CONFIG);
    setStatusMessage('Keyboard shortcuts reset to defaults.');
  }, []);

  const applyZoomPreset = useCallback((preset) => {
    if (preset === 'fit-width') {
      setViewScale(1);
      setStatusMessage('Zoom preset: fit width.');
      return;
    }

    if (preset === 'fit-page') {
      setViewScale(0.82);
      setStatusMessage('Zoom preset: fit page.');
      return;
    }

    setViewScale(DEFAULT_VIEW_SCALE);
    setStatusMessage('Zoom reset to 100%.');
  }, []);

  const rotateViewClockwise = useCallback(() => {
    setRotationDegrees((currentValue) => (currentValue + 90) % 360);
    setStatusMessage('Rotated page view 90 degrees.');
  }, []);

  const handleAddTodo = useCallback(() => {
    const trimmedInput = todoInput.trim();

    if (!trimmedInput) {
      return;
    }

    setTodoItems((currentItems) => [{ id: `${Date.now()}`, text: trimmedInput, done: false }, ...currentItems]);
    setTodoInput('');
  }, [todoInput]);

  const toggleTodoItem = useCallback((todoId) => {
    setTodoItems((currentItems) => currentItems.map((item) => {
      if (item.id !== todoId) {
        return item;
      }

      return { ...item, done: !item.done };
    }));
  }, []);

  const removeTodoItem = useCallback((todoId) => {
    setTodoItems((currentItems) => currentItems.filter((item) => item.id !== todoId));
  }, []);

  const launchSampleNotebook = useCallback(() => {
    const sampleDocumentId = `sample:${Date.now()}`;
    loadNotebookSession({ documentId: sampleDocumentId, notebookName: 'Sample Project Notebook', fresh: true });
    setShowOnboardingTour(false);
    if (typeof window !== 'undefined') {
      window.localStorage.setItem(ONBOARDING_STORAGE_KEY, 'true');
    }
  }, [loadNotebookSession]);

  const applyTemplateSelection = useCallback((templateId) => {
    setPaperTemplate(templateId);
    setStatusMessage(`${templateId.replace('-', ' ')} template selected.`);
  }, []);

  const resizeSelectedStrokes = useCallback((scaleFactor) => {
    const selection = lassoSelectionRef.current;

    if (!selection || selection.pageNumber !== activePageNumber || !selection.strokeIndexes.length || !selection.bounds) {
      setStatusMessage('Select strokes with lasso before resizing.');
      return;
    }

    const pageAnnotations = ensurePageAnnotations(activePageNumber);
    const selectedIndexes = new Set(selection.strokeIndexes);
    const centerPoint = {
      x: selection.bounds.minX + selection.bounds.width / 2,
      y: selection.bounds.minY + selection.bounds.height / 2,
    };
    const nextStrokes = pageAnnotations.strokes.map((stroke, strokeIndex) => {
      if (!selectedIndexes.has(strokeIndex)) {
        return stroke;
      }

      return scaleStrokePoints(stroke, centerPoint, scaleFactor);
    });
    const nextAnnotations = {
      ...annotationsRef.current,
      [activePageNumber]: {
        ...pageAnnotations,
        strokes: nextStrokes,
      },
    };

    annotationsRef.current = nextAnnotations;
    const nextSelection = {
      ...selection,
      bounds: computeSelectionBoundsForIndexes(nextStrokes, selection.strokeIndexes),
    };
    lassoSelectionRef.current = nextSelection;
    setLassoSelection(nextSelection);
    redrawInkPage(activePageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Resized selected strokes (${Math.round(scaleFactor * 100)}%).`);
  }, [activePageNumber, computeSelectionBoundsForIndexes, ensurePageAnnotations, persistSession, redrawInkPage]);

  const copySelectedStrokes = useCallback(() => {
    const selection = lassoSelectionRef.current;

    if (!selection || selection.pageNumber !== activePageNumber || !selection.strokeIndexes.length) {
      setStatusMessage('Select strokes with lasso before copying.');
      return;
    }

    const pageAnnotations = ensurePageAnnotations(activePageNumber);
    const copiedStrokes = selection.strokeIndexes
      .map((strokeIndex) => pageAnnotations.strokes[strokeIndex])
      .filter(Boolean)
      .map((stroke) => cloneStroke(stroke));

    if (!copiedStrokes.length) {
      setStatusMessage('No selected strokes available to copy.');
      return;
    }

    lassoClipboardRef.current = copiedStrokes;
    setStatusMessage(`Copied ${copiedStrokes.length} stroke${copiedStrokes.length > 1 ? 's' : ''}.`);
  }, [activePageNumber, ensurePageAnnotations]);

  const pasteSelectedStrokes = useCallback(() => {
    const clipboardStrokes = lassoClipboardRef.current;

    if (!Array.isArray(clipboardStrokes) || !clipboardStrokes.length) {
      setStatusMessage('Clipboard is empty. Copy a lasso selection first.');
      return;
    }

    const pageAnnotations = ensurePageAnnotations(activePageNumber);
    const translatedStrokes = clipboardStrokes.map((stroke) => translateStrokePoints(stroke, 0.02, 0.02));
    const insertionStartIndex = pageAnnotations.strokes.length;
    const nextStrokes = [...pageAnnotations.strokes, ...translatedStrokes];
    const nextIndexes = translatedStrokes.map((_, index) => insertionStartIndex + index);
    const nextAnnotations = {
      ...annotationsRef.current,
      [activePageNumber]: {
        ...pageAnnotations,
        strokes: nextStrokes,
        redoStack: [],
      },
    };

    annotationsRef.current = nextAnnotations;
    const nextSelection = {
      pageNumber: activePageNumber,
      strokeIndexes: nextIndexes,
      bounds: computeSelectionBoundsForIndexes(nextStrokes, nextIndexes),
    };

    lassoSelectionRef.current = nextSelection;
    setLassoSelection(nextSelection);
    setLassoPath(null);
    redrawInkPage(activePageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Pasted ${translatedStrokes.length} stroke${translatedStrokes.length > 1 ? 's' : ''}.`);
  }, [activePageNumber, computeSelectionBoundsForIndexes, ensurePageAnnotations, persistSession, redrawInkPage]);

  const deleteSelectedStrokes = useCallback(() => {
    const selection = lassoSelectionRef.current;

    if (!selection || selection.pageNumber !== activePageNumber || !selection.strokeIndexes.length) {
      setStatusMessage('Select strokes with lasso before deleting.');
      return;
    }

    const pageAnnotations = ensurePageAnnotations(activePageNumber);
    const selectedIndexes = new Set(selection.strokeIndexes);
    const nextStrokes = pageAnnotations.strokes.filter((_, strokeIndex) => !selectedIndexes.has(strokeIndex));
    const nextAnnotations = {
      ...annotationsRef.current,
      [activePageNumber]: {
        ...pageAnnotations,
        strokes: nextStrokes,
        redoStack: [],
      },
    };

    annotationsRef.current = nextAnnotations;
    lassoSelectionRef.current = null;
    setLassoSelection(null);
    setLassoPath(null);
    redrawInkPage(activePageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage('Deleted selected strokes.');
  }, [activePageNumber, ensurePageAnnotations, persistSession, redrawInkPage]);

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
    if (lassoSelectionRef.current?.pageNumber === activePageNumber) {
      lassoSelectionRef.current = null;
      setLassoSelection(null);
      setLassoPath(null);
    }
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
    if (lassoSelectionRef.current?.pageNumber === activePageNumber) {
      lassoSelectionRef.current = null;
      setLassoSelection(null);
      setLassoPath(null);
    }
    redrawInkPage(activePageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage(`Restored a stroke on page ${activePageNumber}.`);
  }, [activePageNumber, ensurePageAnnotations, persistSession, redrawInkPage]);

  useEffect(() => {
    const handleKeyDown = (event) => {
      if (currentView !== 'editor') {
        return;
      }

      const target = event.target;
      const isTypingTarget = target instanceof HTMLElement
        && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable);

      if (isTypingTarget || event.metaKey || event.ctrlKey || event.altKey) {
        return;
      }

      const pressedKey = normalizeShortcutKey(event.key);

      if (!pressedKey) {
        return;
      }

      if (pressedKey === shortcutConfig.pen) {
        event.preventDefault();
        activateTool('pen', true);
        return;
      }

      if (pressedKey === shortcutConfig.highlighter) {
        event.preventDefault();
        activateTool('highlighter', true);
        return;
      }

      if (pressedKey === shortcutConfig.eraser) {
        event.preventDefault();
        activateTool('eraser', true);
        return;
      }

      if (pressedKey === shortcutConfig.undo) {
        event.preventDefault();
        handleUndo();
        return;
      }

      if (pressedKey === shortcutConfig.redo) {
        event.preventDefault();
        handleRedo();
      }
    };

    window.addEventListener('keydown', handleKeyDown);

    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [activateTool, currentView, handleRedo, handleUndo, shortcutConfig]);

  useEffect(() => {
    const handleLassoShortcuts = (event) => {
      if (currentView !== 'editor') {
        return;
      }

      const target = event.target;
      const isTypingTarget = target instanceof HTMLElement
        && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable);

      if (isTypingTarget || activeTool !== 'lasso') {
        return;
      }

      const pressedKey = event.key.toLowerCase();

      if ((event.metaKey || event.ctrlKey) && pressedKey === 'c') {
        event.preventDefault();
        copySelectedStrokes();
        return;
      }

      if ((event.metaKey || event.ctrlKey) && pressedKey === 'v') {
        event.preventDefault();
        pasteSelectedStrokes();
        return;
      }

      if (!event.metaKey && !event.ctrlKey && !event.altKey && (event.key === 'Delete' || event.key === 'Backspace')) {
        event.preventDefault();
        deleteSelectedStrokes();
      }
    };

    window.addEventListener('keydown', handleLassoShortcuts);

    return () => {
      window.removeEventListener('keydown', handleLassoShortcuts);
    };
  }, [activeTool, copySelectedStrokes, currentView, deleteSelectedStrokes, pasteSelectedStrokes]);

  useEffect(() => {
    return () => {
      const recognition = speechRecognitionRef.current;

      if (recognition) {
        recognition.stop();
      }

      if (transcriptToastTimerRef.current) {
        window.clearTimeout(transcriptToastTimerRef.current);
      }

      if (voiceVolumeFrameRef.current) {
        window.cancelAnimationFrame(voiceVolumeFrameRef.current);
      }

      if (voiceAudioStreamRef.current) {
        voiceAudioStreamRef.current.getTracks().forEach((track) => {
          track.stop();
        });
      }

      if (voiceAudioContextRef.current) {
        voiceAudioContextRef.current.close();
      }
    };
  }, []);

  const handleClearPage = useCallback(() => {
    const nextAnnotations = {
      ...annotationsRef.current,
      [activePageNumber]: createEmptyPageAnnotations(),
    };

    annotationsRef.current = nextAnnotations;
    if (lassoSelectionRef.current?.pageNumber === activePageNumber) {
      lassoSelectionRef.current = null;
      setLassoSelection(null);
      setLassoPath(null);
    }
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

  const scrollViewerByVoice = useCallback((delta, directionLabel) => {
    const viewerNode = viewerShellRef.current;

    if (!viewerNode) {
      setStatusMessage('Viewer is not ready for scrolling yet.');
      return false;
    }

    viewerNode.scrollTop += delta;
    setStatusMessage(`Voice: scrolled ${directionLabel}.`);
    return true;
  }, []);

  const getCenterViewportTarget = useCallback(() => {
    const viewerNode = viewerShellRef.current;

    if (!viewerNode) {
      return null;
    }

    const viewerRect = viewerNode.getBoundingClientRect();
    const centerX = viewerRect.left + viewerRect.width / 2;
    const centerY = viewerRect.top + viewerRect.height / 2;
    const centerElement = document.elementFromPoint(centerX, centerY);

    if (!centerElement) {
      return null;
    }

    const pageStage = centerElement.closest('.page-stage');
    const stackStage = pageStage?.querySelector('.stack-stage') ?? null;
    const stackRect = stackStage?.getBoundingClientRect() ?? null;
    const textElement = centerElement.closest('span, p, strong, label, h1, h2, h3, h4, li, a, button');
    const parsedPage = Number.parseInt(pageStage?.getAttribute('data-page-number') || '', 10);
    const pageNumber = Number.isFinite(parsedPage) ? parsedPage : activePageNumber;

    let normalizedCenter = { x: 0.5, y: 0.5 };

    if (stackRect) {
      normalizedCenter = {
        x: clampNumber((centerX - stackRect.left) / Math.max(stackRect.width, 1), 0.02, 0.98),
        y: clampNumber((centerY - stackRect.top) / Math.max(stackRect.height, 1), 0.02, 0.98),
      };
    }

    return {
      pageNumber,
      stackRect,
      textElement,
      textRect: textElement?.getBoundingClientRect() ?? null,
      normalizedCenter,
    };
  }, [activePageNumber]);

  const applyCenterFocusHighlight = useCallback(() => {
    const centerTarget = getCenterViewportTarget();

    if (!centerTarget) {
      setStatusMessage('Voice: unable to detect center focus target.');
      return false;
    }

    const pageNumber = centerTarget.pageNumber;
    const pageAnnotations = ensurePageAnnotations(pageNumber);
    const baseCenter = centerTarget.normalizedCenter;
    let minX = clampNumber(baseCenter.x - 0.16, 0.02, 0.94);
    let maxX = clampNumber(baseCenter.x + 0.16, 0.06, 0.98);
    let minY = clampNumber(baseCenter.y - 0.028, 0.02, 0.94);
    let maxY = clampNumber(baseCenter.y + 0.028, 0.06, 0.98);

    if (centerTarget.textRect && centerTarget.stackRect) {
      const { textRect, stackRect } = centerTarget;
      minX = clampNumber((textRect.left - stackRect.left - 6) / Math.max(stackRect.width, 1), 0.02, 0.96);
      maxX = clampNumber((textRect.right - stackRect.left + 6) / Math.max(stackRect.width, 1), 0.04, 0.98);
      minY = clampNumber((textRect.top - stackRect.top - 6) / Math.max(stackRect.height, 1), 0.02, 0.96);
      maxY = clampNumber((textRect.bottom - stackRect.top + 6) / Math.max(stackRect.height, 1), 0.04, 0.98);
    }

    if (maxX - minX < 0.05) {
      maxX = clampNumber(minX + 0.11, 0.07, 0.98);
    }

    if (maxY - minY < 0.012) {
      maxY = clampNumber(minY + 0.03, 0.04, 0.98);
    }

    const highlightStroke = cloneStroke({
      id: `voice-highlight-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      tool: 'highlighter',
      variant: selectedHighlighterPreset.id,
      color: inkColor,
      width: Math.max(highlighterSize * selectedHighlighterPreset.widthScale, TOOL_PRESETS.highlighter.width * 0.8),
      opacity: selectedHighlighterPreset.opacity,
      pressureEnabled: false,
      shapeType: null,
      createdAt: new Date().toISOString(),
      points: [
        { x: minX, y: minY, pressure: 1 },
        { x: maxX, y: minY, pressure: 1 },
        { x: maxX, y: maxY, pressure: 1 },
        { x: minX, y: maxY, pressure: 1 },
        { x: minX, y: minY, pressure: 1 },
      ],
    });

    const nextAnnotations = {
      ...annotationsRef.current,
      [pageNumber]: {
        ...pageAnnotations,
        strokes: [...pageAnnotations.strokes, highlightStroke],
        redoStack: [],
      },
    };

    annotationsRef.current = nextAnnotations;
    redrawInkPage(pageNumber);
    persistSession(nextAnnotations);
    setAnnotationVersion((currentValue) => currentValue + 1);
    setStatusMessage('Voice: smart center highlight applied.');
    return true;
  }, [ensurePageAnnotations, getCenterViewportTarget, highlighterSize, inkColor, persistSession, redrawInkPage, selectedHighlighterPreset]);

  const jumpToPageFromVoice = useCallback((requestedPageNumber) => {
    if (!Number.isFinite(requestedPageNumber)) {
      return false;
    }

    const nextPageNumber = Math.max(1, Math.min(pageCount, Math.trunc(requestedPageNumber)));
    setPageJumpValue(String(nextPageNumber));
    scrollToPage(nextPageNumber);
    setStatusMessage(`Voice: jumped to page ${nextPageNumber}.`);
    return true;
  }, [pageCount, scrollToPage]);

  const saveReminderFromTranscript = useCallback(async (rawTranscript) => {
    const parsedReminder = parseReminderTask(rawTranscript);

    if (!parsedReminder) {
      return false;
    }

    const reminderRecord = {
      reminderId: `reminder:${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      taskDescription: parsedReminder.taskDescription,
      createdAt: new Date().toISOString(),
      dueAt: parsedReminder.dueAt ? parsedReminder.dueAt.toISOString() : null,
      sourceTranscript: rawTranscript,
    };

    try {
      await saveReminder(reminderRecord);
      const dueLabel = parsedReminder.dueAt ? ` for ${parsedReminder.dueAt.toLocaleString()}` : '';
      setStatusMessage(`Voice reminder saved: ${parsedReminder.taskDescription}${dueLabel}.`);
      return true;
    } catch (error) {
      setStatusMessage('Unable to save reminder right now.');
      return false;
    }
  }, []);

  const stopVoiceVisualizer = useCallback(() => {
    if (voiceVolumeFrameRef.current) {
      window.cancelAnimationFrame(voiceVolumeFrameRef.current);
      voiceVolumeFrameRef.current = null;
    }

    if (voiceAudioStreamRef.current) {
      voiceAudioStreamRef.current.getTracks().forEach((track) => {
        track.stop();
      });
      voiceAudioStreamRef.current = null;
    }

    if (voiceAudioContextRef.current) {
      voiceAudioContextRef.current.close();
      voiceAudioContextRef.current = null;
    }

    voiceAnalyserRef.current = null;
    setVoiceLevel(0);
  }, []);

  const startVoiceVisualizer = useCallback(async () => {
    if (typeof window === 'undefined' || !navigator.mediaDevices?.getUserMedia) {
      return;
    }

    stopVoiceVisualizer();

    try {
      const mediaStream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const audioContext = new (window.AudioContext || window.webkitAudioContext)();
      const analyser = audioContext.createAnalyser();
      analyser.fftSize = 256;
      analyser.smoothingTimeConstant = 0.75;
      const source = audioContext.createMediaStreamSource(mediaStream);
      source.connect(analyser);

      voiceAudioStreamRef.current = mediaStream;
      voiceAudioContextRef.current = audioContext;
      voiceAnalyserRef.current = analyser;

      const frequencyBuffer = new Uint8Array(analyser.frequencyBinCount);

      const tick = () => {
        const activeAnalyser = voiceAnalyserRef.current;

        if (!activeAnalyser) {
          return;
        }

        activeAnalyser.getByteFrequencyData(frequencyBuffer);
        const total = frequencyBuffer.reduce((sum, value) => sum + value, 0);
        const normalizedLevel = clampNumber(total / (frequencyBuffer.length * 180), 0, 1);
        setVoiceLevel(normalizedLevel);
        voiceVolumeFrameRef.current = window.requestAnimationFrame(tick);
      };

      voiceVolumeFrameRef.current = window.requestAnimationFrame(tick);
    } catch (error) {
      setStatusMessage('Microphone permission is required for waveform feedback.');
    }
  }, [stopVoiceVisualizer]);

  const dispatchVoiceCommand = useCallback((rawTranscript) => {
    const transcript = String(rawTranscript || '').trim().toLowerCase();

    if (!transcript) {
      return false;
    }

    setLastVoiceTranscript(transcript);

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.scrollDown)) {
      return scrollViewerByVoice(VOICE_SCROLL_STEP_PX, 'down');
    }

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.scrollUp)) {
      return scrollViewerByVoice(-VOICE_SCROLL_STEP_PX, 'up');
    }

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.nextPage)) {
      return jumpToPageFromVoice(activePageNumber + 1);
    }

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.previousPage)) {
      return jumpToPageFromVoice(activePageNumber - 1);
    }

    const goToPageMatch = transcript.match(/\b(?:go to|jump to|open)?\s*page\s*(\d{1,4})\b/i)
      ?? transcript.match(/पान\s*(\d{1,4})/i)
      ?? transcript.match(/\b(\d{1,4})\b/);
    if (goToPageMatch && transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.jumpToPage)) {
      const parsedPageNumber = Number.parseInt(goToPageMatch[1], 10);
      if (Number.isFinite(parsedPageNumber)) {
        return jumpToPageFromVoice(parsedPageNumber);
      }
    }

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.remind)) {
      saveReminderFromTranscript(rawTranscript);
      return true;
    }

    const matchedColor = Object.entries(VOICE_COLOR_MAP).find(([colorName]) => {
      return new RegExp(`\\b${colorName}\\b`, 'i').test(transcript);
    });

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.highlighter)) {
      activateTool('highlighter', true);

      if (matchedColor) {
        setInkColor(matchedColor[1]);
        setStatusMessage(`Voice: highlighter set to ${matchedColor[0]}.`);
      } else {
        setStatusMessage('Voice: highlighter selected.');
      }

      if (transcript.includes('highlight') || transcript.includes('हायलाइट')) {
        applyCenterFocusHighlight();
      }

      return true;
    }

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.pen)) {
      activateTool('pen', true);

      if (matchedColor) {
        setInkColor(matchedColor[1]);
        setStatusMessage(`Voice: pen set to ${matchedColor[0]}.`);
      } else {
        setStatusMessage('Voice: pen selected.');
      }

      return true;
    }

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.eraser)) {
      activateTool('eraser', true);
      setStatusMessage('Voice: eraser selected.');
      return true;
    }

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.lasso)) {
      activateTool('lasso', true);
      setStatusMessage('Voice: lasso selected.');
      return true;
    }

    if (transcriptIncludesAnyKeyword(transcript, activeVoiceKeywordMap.shape)) {
      activateTool('shape', true);
      setStatusMessage('Voice: shape tool selected.');
      return true;
    }

    if (transcript.includes('undo')) {
      handleUndo();
      return true;
    }

    if (transcript.includes('redo')) {
      handleRedo();
      return true;
    }

    setStatusMessage(`Voice command not recognized: "${transcript}".`);
    return false;
  }, [activateTool, activePageNumber, activeVoiceKeywordMap, applyCenterFocusHighlight, handleRedo, handleUndo, jumpToPageFromVoice, saveReminderFromTranscript, scrollViewerByVoice]);

  const stopVoiceRecognition = useCallback(() => {
    const recognition = speechRecognitionRef.current;

    if (recognition) {
      recognition.stop();
      speechRecognitionRef.current = null;
    }

    stopVoiceVisualizer();
    setInterimVoiceTranscript('');
    setIsVoiceListening(false);
  }, [stopVoiceVisualizer]);

  const startVoiceRecognition = useCallback((language = voiceLanguage) => {
    if (typeof window === 'undefined') {
      return;
    }

    const RecognitionApi = window.SpeechRecognition || window.webkitSpeechRecognition;

    if (!RecognitionApi) {
      setStatusMessage('Speech recognition is unavailable in this browser.');
      return;
    }

    if (speechRecognitionRef.current) {
      speechRecognitionRef.current.stop();
      speechRecognitionRef.current = null;
    }

    const recognition = new RecognitionApi();
    recognition.lang = language;
    recognition.continuous = true;
    recognition.interimResults = true;

    recognition.onresult = (event) => {
      const nextTranscripts = [];
      const interimParts = [];

      for (let index = event.resultIndex; index < event.results.length; index += 1) {
        const result = event.results[index];

        if (result.isFinal && result[0]?.transcript) {
          nextTranscripts.push(result[0].transcript);
          continue;
        }

        if (result[0]?.transcript) {
          interimParts.push(result[0].transcript);
        }
      }

      const interimTranscript = interimParts.join(' ').trim();
      setInterimVoiceTranscript(interimTranscript);

      if (interimTranscript) {
        setVoiceTranscriptToast(interimTranscript);

        if (transcriptToastTimerRef.current) {
          window.clearTimeout(transcriptToastTimerRef.current);
        }

        transcriptToastTimerRef.current = window.setTimeout(() => {
          setVoiceTranscriptToast('');
          transcriptToastTimerRef.current = null;
        }, 1500);
      }

      nextTranscripts.forEach((transcript) => {
        dispatchVoiceCommand(transcript);
      });
    };

    recognition.onerror = (event) => {
      setStatusMessage(`Voice error: ${event.error || 'unknown speech error'}.`);
      stopVoiceVisualizer();
      setIsVoiceListening(false);
    };

    recognition.onend = () => {
      speechRecognitionRef.current = null;
      stopVoiceVisualizer();
      setInterimVoiceTranscript('');
      setIsVoiceListening(false);
    };

    recognition.start();
    startVoiceVisualizer();
    speechRecognitionRef.current = recognition;
    setIsVoiceListening(true);
    setStatusMessage(`Voice commands are listening (${language}).`);
  }, [dispatchVoiceCommand, startVoiceVisualizer, stopVoiceVisualizer, voiceLanguage]);

  const toggleVoiceRecognition = useCallback(() => {
    if (isVoiceListening) {
      stopVoiceRecognition();
      setStatusMessage('Voice commands stopped.');
      return;
    }

    startVoiceRecognition();
  }, [isVoiceListening, startVoiceRecognition, stopVoiceRecognition]);

  const handleVoiceLanguageChange = useCallback((event) => {
    const nextLanguage = event.target.value;
    setVoiceLanguage(nextLanguage);

    if (isVoiceListening) {
      stopVoiceRecognition();
      window.setTimeout(() => {
        startVoiceRecognition(nextLanguage);
      }, 160);
    }
  }, [isVoiceListening, startVoiceRecognition, stopVoiceRecognition]);

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

    const fileName = file.name || '';
    const isPdf = file.type === 'application/pdf' || fileName.toLowerCase().endsWith('.pdf');

    if (!isPdf) {
      setStatusMessage('Only PDF files are supported for import.');
      event.target.value = '';
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

    if (stroke.tool === 'shape') {
      const startPoint = stroke.points[0];
      const endPoint = stroke.points[stroke.points.length - 1] ?? startPoint;
      const startX = startPoint.x * width;
      const startY = height - startPoint.y * height;
      const endX = endPoint.x * width;
      const endY = height - endPoint.y * height;

      if (stroke.shapeType === 'rectangle') {
        page.drawRectangle({
          x: Math.min(startX, endX),
          y: Math.min(startY, endY),
          width: Math.abs(endX - startX),
          height: Math.abs(endY - startY),
          borderWidth: stroke.width,
          borderColor: rgbColor,
          opacity: stroke.opacity,
        });
        return;
      }

      if (stroke.shapeType === 'circle') {
        const centerX = (startX + endX) / 2;
        const centerY = (startY + endY) / 2;
        page.drawEllipse({
          x: centerX,
          y: centerY,
          xScale: Math.max(Math.abs(endX - startX) / 2, 0.5),
          yScale: Math.max(Math.abs(endY - startY) / 2, 0.5),
          borderWidth: stroke.width,
          borderColor: rgbColor,
          opacity: stroke.opacity,
        });
        return;
      }

      page.drawLine({
        start: { x: startX, y: startY },
        end: { x: endX, y: endY },
        thickness: stroke.width,
        color: rgbColor,
        opacity: stroke.opacity,
      });

      if (stroke.shapeType === 'arrow') {
        const deltaX = endX - startX;
        const deltaY = endY - startY;
        const angle = Math.atan2(deltaY, deltaX);
        const arrowLength = Math.max(10, stroke.width * 3.2);
        const spread = Math.PI / 7;

        page.drawLine({
          start: { x: endX, y: endY },
          end: {
            x: endX - arrowLength * Math.cos(angle - spread),
            y: endY - arrowLength * Math.sin(angle - spread),
          },
          thickness: stroke.width,
          color: rgbColor,
          opacity: stroke.opacity,
        });
        page.drawLine({
          start: { x: endX, y: endY },
          end: {
            x: endX - arrowLength * Math.cos(angle + spread),
            y: endY - arrowLength * Math.sin(angle + spread),
          },
          thickness: stroke.width,
          color: rgbColor,
          opacity: stroke.opacity,
        });
      }

      return;
    }

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
  const pageStrokeCountMap = pages.reduce((nextCounts, page) => {
    nextCounts[page.pageNumber] = annotationsRef.current[page.pageNumber]?.strokes?.length ?? 0;
    return nextCounts;
  }, {});
  const totalStrokeCount = Object.values(annotationsRef.current).reduce((strokeCount, pageAnnotations) => {
    return strokeCount + (pageAnnotations?.strokes?.length ?? 0);
  }, 0);
  const _annotationVersion = annotationVersion;
  void _annotationVersion;

  return (
    <div className={`app-shell theme-${themePreset} notes-layout-${notesLayout} ${highContrastMode ? 'high-contrast' : ''} ${largeHitTargets ? 'large-targets' : ''} ${focusModeEnabled ? 'focus-mode' : ''}`} style={{ '--ui-scale': uiScale }}>
      <header className="app-header">
        <div>
          <p className="eyebrow">Tablet-style notes</p>
          <h1>VoxNotes AI</h1>
          <p className="header-copy">
            Start from a notes library, then open a ruled notebook or imported PDF. Write with pen or highlighter, erase strokes, add blank pages, jump through long documents, and export everything back to PDF.
          </p>
        </div>

        <div className="global-controls" aria-label="Reading and accessibility controls">
          <label className="global-control-row">
            <span>Paper</span>
            <select value={paperTemplate} onChange={(event) => setPaperTemplate(event.target.value)}>
              <option value="ruled">Ruled</option>
              <option value="grid">Grid</option>
              <option value="dot">Dot</option>
              <option value="cornell">Cornell</option>
              <option value="planner">Planner</option>
            </select>
          </label>
          <label className="global-control-row">
            <span>Notes Layout</span>
            <select value={notesLayout} onChange={(event) => setNotesLayout(event.target.value)}>
              <option value="tablet">Tablet</option>
              <option value="classic">Classic</option>
            </select>
          </label>
          <div className="global-control-inline">
            <button type="button" className="overlay-button compact" onClick={() => applyZoomPreset('fit-width')}>Fit Width</button>
            <button type="button" className="overlay-button compact" onClick={() => applyZoomPreset('fit-page')}>Fit Page</button>
            <button type="button" className="overlay-button compact" onClick={rotateViewClockwise}>Rotate</button>
            <button type="button" className={`overlay-button compact ${splitViewEnabled ? 'selected' : ''}`} onClick={() => setSplitViewEnabled((currentValue) => !currentValue)}>Split</button>
          </div>
          <label className="global-control-row">
            <span>UI Scale</span>
            <input type="range" min="0.85" max="1.25" step="0.05" value={uiScale} onChange={(event) => setUiScale(Number.parseFloat(event.target.value))} />
          </label>
          <div className="global-control-inline">
            <button type="button" className={`overlay-button compact ${largeHitTargets ? 'selected' : ''}`} onClick={() => setLargeHitTargets((currentValue) => !currentValue)}>Large Targets</button>
            <button type="button" className={`overlay-button compact ${highContrastMode ? 'selected' : ''}`} onClick={() => setHighContrastMode((currentValue) => !currentValue)}>High Contrast</button>
          </div>
        </div>
      </header>

      <main className="workspace-card">
        {currentView === 'home' ? (
          <div className="library-shell">
            <aside className="library-sidebar">
              <h2 className="library-brand">Noteshelf</h2>

              <div className="library-tile-grid">
                {LIBRARY_TILES.map((tile) => {
                  const TileIcon = tile.icon;

                  return (
                    <button key={tile.key} type="button" className={`library-tile ${tile.accentClass} ${tile.active ? 'active' : ''}`}>
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

              <div className="library-folder-card">
                <div className="library-section-head">
                  <h3>Content</h3>
                </div>
                <button type="button" className="folder-row">
                  <Image size={16} />
                  <span>Photos</span>
                </button>
                <button type="button" className="folder-row">
                  <Mic size={16} />
                  <span>Recordings</span>
                </button>
                <button type="button" className="folder-row">
                  <Bookmark size={16} />
                  <span>Bookmarks</span>
                </button>
              </div>

              <div className="library-upgrade-card">
                <span className="upgrade-badge"><Crown size={14} /> Upgrade to Premium</span>
                <strong>0 notes left</strong>
                <button type="button" className="upgrade-button">Upgrade Now</button>
              </div>
            </aside>

            <section className="library-main">
              <div className="library-main-head">
                <div>
                  <h2>All Notes</h2>
                </div>
                <div className="library-main-actions" aria-label="Library quick controls">
                  <button type="button" className="main-action-icon" onClick={handleCreateFreshNotebook} aria-label="Create note">
                    <Plus size={18} />
                  </button>
                  <button type="button" className="main-action-icon" aria-label="Search notes">
                    <Search size={18} />
                  </button>
                  <button type="button" className="main-action-icon" aria-label="More options">
                    <EllipsisVertical size={18} />
                  </button>
                </div>
              </div>

              <div className="library-actions-row">
                <button type="button" className="library-action-card" onClick={() => loadNotebookSession()}>
                  <span className="library-action-icon plus">+</span>
                  <strong>Quick Note</strong>
                </button>
                <button type="button" className="library-action-card" onClick={handleCreateFreshNotebook}>
                  <span className="library-action-icon notebook"><NotebookPen size={18} /></span>
                  <strong>Notebook</strong>
                </button>
                <button type="button" className="library-action-card" onClick={() => fileInputRef.current?.click()}>
                  <span className="library-action-icon import"><Download size={18} /></span>
                  <strong>Import File</strong>
                </button>
              </div>

              <div className="library-note-grid">
                {filteredLibraryItems.length ? filteredLibraryItems.map((item) => (
                  <button key={item.documentId} type="button" className="note-card" onClick={() => handleOpenStoredItem(item)}>
                    <div className={`note-card-cover ${item.accentClass} ${item.sourceType === 'pdf' ? 'pdf' : 'notebook'}`}>
                      {item.sourceType === 'pdf' ? <FileText size={22} /> : <NotebookPen size={22} />}
                    </div>
                    <strong>{item.documentName}</strong>
                    <span>{item.updatedAt ? new Date(item.updatedAt).toLocaleString([], { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' }) : 'Just now'}</span>
                    <span>Unfiled</span>
                  </button>
                )) : (
                  <div className="library-empty-state">
                    <NotebookPen size={28} />
                    <strong>No notes matched your filters</strong>
                    <p>Try another smart view or clear filters to discover your note library.</p>
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
                      <button type="button" className={`tool-rail-button ${activeTool === 'shape' ? 'selected' : ''}`} onClick={() => handleToolSelect('shape')} aria-label="Shapes">
                        <span className="tool-rail-icon simple">
                          <RailGlyph kind="shape" />
                        </span>
                      </button>
                      <button type="button" className={`tool-rail-button ${activeTool === 'lasso' ? 'selected' : ''}`} onClick={() => handleToolSelect('lasso')} aria-label="Lasso select">
                        <span className="tool-rail-icon simple">
                          <RailGlyph kind="lasso" />
                        </span>
                      </button>
                      <button type="button" className="tool-rail-button action" onClick={handleUndo} disabled={!activePageAnnotations.strokes.length} aria-label="Undo">
                        <Undo2 size={13} />
                      </button>
                      <button type="button" className="tool-rail-button action" onClick={handleRedo} disabled={!activePageAnnotations.redoStack.length} aria-label="Redo">
                        <Redo2 size={13} />
                      </button>
                      <button type="button" className={`tool-rail-button action ${toolPanel === 'shortcuts' ? 'selected' : ''}`} onClick={() => setToolPanel((currentValue) => currentValue === 'shortcuts' ? null : 'shortcuts')} aria-label="Keyboard shortcuts">
                        <Keyboard size={13} />
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

                        {toolPanel === 'shape' ? (
                          <>
                            <div className="tool-popover-head">
                              <strong>Shape tools</strong>
                              <span className="tool-hint">Draw straight guides, callouts, and geometric shapes.</span>
                            </div>

                            <div className="shape-preset-grid">
                              {SHAPE_VARIANTS.map((shapePreset) => (
                                <button
                                  key={shapePreset.id}
                                  type="button"
                                  className={`shape-preset-chip ${shapeType === shapePreset.id ? 'selected' : ''}`}
                                  onClick={() => {
                                    setShapeType(shapePreset.id);
                                    setStatusMessage(`${shapePreset.label} selected.`);
                                  }}
                                >
                                  {shapePreset.label}
                                </button>
                              ))}
                            </div>

                            <div className="tool-slider-row">
                              <button type="button" className="mini-icon-button" onClick={() => setShapeSize((currentValue) => Math.max(1, currentValue - 1))}>-</button>
                              <div className="tool-slider-pill">{shapeSize}</div>
                              <input type="range" min="1" max="16" value={shapeSize} onChange={(event) => setShapeSize(Number.parseInt(event.target.value, 10))} />
                              <button type="button" className="mini-icon-button" onClick={() => setShapeSize((currentValue) => Math.min(16, currentValue + 1))}>+</button>
                            </div>

                            <div className="color-group tool-popover-colors" aria-label="Shape colors">
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
                              <input type="color" value={inkColor} onChange={(event) => setInkColor(event.target.value)} aria-label="Custom shape color" />
                              <strong>{inkColor.toUpperCase()}</strong>
                            </label>
                          </>
                        ) : null}

                        {toolPanel === 'lasso' ? (
                          <>
                            <div className="tool-popover-head">
                              <strong>Lasso select</strong>
                              <span className="tool-hint">Loop strokes to select, drag the selection box to move, then resize.</span>
                            </div>

                            <div className="lasso-actions-grid">
                              <button
                                type="button"
                                className="shape-preset-chip"
                                onClick={() => {
                                  clearLassoSelection();
                                  redrawInkPage(activePageNumber);
                                  setStatusMessage('Cleared lasso selection.');
                                }}
                              >
                                Clear Selection
                              </button>
                              <button type="button" className="shape-preset-chip" onClick={copySelectedStrokes}>
                                Copy Selected
                              </button>
                              <button type="button" className="shape-preset-chip" onClick={pasteSelectedStrokes}>
                                Paste Selected
                              </button>
                              <button type="button" className="shape-preset-chip" onClick={deleteSelectedStrokes}>
                                Delete Selected
                              </button>
                              <button type="button" className="shape-preset-chip" onClick={() => resizeSelectedStrokes(0.9)}>
                                Resize 90%
                              </button>
                              <button type="button" className="shape-preset-chip" onClick={() => resizeSelectedStrokes(1.1)}>
                                Resize 110%
                              </button>
                              <button type="button" className="shape-preset-chip" onClick={() => resizeSelectedStrokes(1.25)}>
                                Resize 125%
                              </button>
                            </div>
                          </>
                        ) : null}

                        {toolPanel === 'shortcuts' ? (
                          <>
                            <div className="tool-popover-head">
                              <strong>Keyboard shortcuts</strong>
                              <span className="tool-hint">Set one-letter keys for faster tool switching.</span>
                            </div>

                            <div className="shortcut-grid" role="group" aria-label="Keyboard shortcut settings">
                              <label className="shortcut-row">
                                <span>Pen</span>
                                <input type="text" value={shortcutConfig.pen.toUpperCase()} onChange={(event) => handleShortcutChange('pen', event.target.value)} maxLength={1} />
                              </label>
                              <label className="shortcut-row">
                                <span>Highlighter</span>
                                <input type="text" value={shortcutConfig.highlighter.toUpperCase()} onChange={(event) => handleShortcutChange('highlighter', event.target.value)} maxLength={1} />
                              </label>
                              <label className="shortcut-row">
                                <span>Eraser</span>
                                <input type="text" value={shortcutConfig.eraser.toUpperCase()} onChange={(event) => handleShortcutChange('eraser', event.target.value)} maxLength={1} />
                              </label>
                              <label className="shortcut-row">
                                <span>Undo</span>
                                <input type="text" value={shortcutConfig.undo.toUpperCase()} onChange={(event) => handleShortcutChange('undo', event.target.value)} maxLength={1} />
                              </label>
                              <label className="shortcut-row">
                                <span>Redo</span>
                                <input type="text" value={shortcutConfig.redo.toUpperCase()} onChange={(event) => handleShortcutChange('redo', event.target.value)} maxLength={1} />
                              </label>
                            </div>

                            <div className="shortcut-actions-row">
                              <button type="button" className="text-action-button" onClick={resetShortcutConfig}>Reset defaults</button>
                              <button type="button" className="text-action-button" onClick={swapRecentInkTools}>Swap last tools</button>
                            </div>
                          </>
                        ) : null}
                      </div>
                    ) : null}
                  </div>

                  <button type="button" className="section-fab" onClick={() => setShowNotebookSettings(true)}>
                    <Settings2 size={15} />
                    <span>Section</span>
                  </button>

                  <button
                    type="button"
                    className={`voice-pill-floating ${isVoiceListening ? 'listening' : ''}`}
                    onClick={toggleVoiceRecognition}
                    aria-pressed={isVoiceListening}
                  >
                    <Mic size={14} />
                    <span>{isVoiceListening ? 'Listening' : 'Voice'}</span>
                  </button>

                  {!isToolbarCollapsed ? (
                    <div className="overlay-stack">
                      {notesLayout === 'tablet' ? (
                        <div className="notes-topbar" aria-label="Tablet notes top bar">
                          <div className="notes-topbar-left">
                            <button type="button" className="overlay-button compact" onClick={handleOpenHome}>
                              <ChevronLeft size={16} />
                            </button>
                            <strong className="notes-topbar-title">{documentName || 'Title'}</strong>
                          </div>
                          <div className="notes-topbar-right">
                            <button type="button" className="overlay-button compact" onClick={() => setShowNotebookSettings(true)}>
                              <Settings2 size={15} />
                              <span>Section</span>
                            </button>
                            <button type="button" className="overlay-button compact" onClick={handleAddPage} disabled={sourceType !== 'notebook'}>
                              <Plus size={16} />
                            </button>
                            <button type="button" className="overlay-button compact" onClick={handleExportPdf} disabled={isExporting}>
                              <Download size={15} />
                            </button>
                            <button type="button" className={`overlay-button compact ${isVoiceListening ? 'selected' : ''}`} onClick={toggleVoiceRecognition} aria-pressed={isVoiceListening}>
                              <Mic size={15} />
                            </button>
                            <button type="button" className="overlay-button compact" onClick={handleClearPage}>
                              <Eraser size={15} />
                            </button>
                          </div>
                        </div>
                      ) : (
                        <div className="overlay-toolbar">
                          <div className="mode-group">
                            <button type="button" className="overlay-button" onClick={handleOpenHome}>
                              <Home size={16} />
                              <span>Library</span>
                            </button>
                            <button type="button" className="overlay-button" onClick={() => setShowNotebookSettings(true)}>
                              <Settings2 size={16} />
                              <span>Section</span>
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
                          <button type="button" className={`overlay-button compact ${isVoiceListening ? 'selected' : ''}`} onClick={toggleVoiceRecognition} aria-pressed={isVoiceListening}>
                            <Mic size={15} />
                            <span>{isVoiceListening ? 'Listening' : 'Voice'}</span>
                          </button>

                          <div className="save-pill compact-pill">
                            <Save size={14} />
                            <span>{lastSavedAt ? `Saved ${lastSavedAt}` : 'Auto-save ready'}</span>
                          </div>

                          <div className="status-pill compact-pill">
                            <span className="status-dot" />
                            <span>{statusMessage}</span>
                          </div>
                        </div>
                      )}

                      <div className={`voice-control-panel ${isVoiceListening ? 'listening' : ''}`} aria-live="polite">
                        <div className="voice-control-top">
                          <div className="voice-control-status">
                            <span className={`voice-status-dot ${isVoiceListening ? 'active' : ''}`} />
                            <strong>{isVoiceListening ? 'Listening now' : 'Voice paused'}</strong>
                          </div>
                          <button type="button" className={`overlay-button compact ${isVoiceListening ? 'selected' : ''}`} onClick={toggleVoiceRecognition} aria-pressed={isVoiceListening}>
                            <Mic size={15} />
                            <span>{isVoiceListening ? 'Stop' : 'Start'}</span>
                          </button>
                        </div>

                        <div className="voice-control-row">
                          <label htmlFor="voice-language-select">Language</label>
                          <select id="voice-language-select" value={voiceLanguage} onChange={handleVoiceLanguageChange}>
                            {VOICE_LANGUAGE_OPTIONS.map((languageOption) => (
                              <option key={languageOption.id} value={languageOption.id}>{languageOption.label}</option>
                            ))}
                          </select>
                        </div>

                        <div className="voice-waveform" aria-hidden="true">
                          {Array.from({ length: 12 }).map((_, index) => {
                            const phase = ((index % 6) + 1) / 6;
                            const barScale = isVoiceListening ? clampNumber(0.28 + voiceLevel * (0.5 + phase), 0.2, 1) : 0.18;
                            return (
                              <span
                                key={`voice-bar-${index}`}
                                className="voice-wave-bar"
                                style={{ '--voice-bar-scale': barScale, '--voice-bar-delay': `${index * 60}ms` }}
                              />
                            );
                          })}
                        </div>

                        <div className="voice-control-caption">
                          {interimVoiceTranscript || lastVoiceTranscript || 'Try: "Highlight", "Go to page 15", "Remind me tomorrow to review notes", "पुढचे पान".'}
                        </div>
                      </div>
                    </div>
                  ) : null}
                </div>
              </div>

              {currentView === 'editor' && miniToolbar.visible ? (
                <div className="mini-tool-float" style={{ left: `${miniToolbar.x}px`, top: `${miniToolbar.y}px` }} role="toolbar" aria-label="Quick mini toolbar">
                  <button type="button" className={`mini-tool-button ${activeTool === 'pen' ? 'selected' : ''}`} onClick={() => activateTool('pen', false)}>
                    Pen
                  </button>
                  <button type="button" className={`mini-tool-button ${activeTool === 'highlighter' ? 'selected' : ''}`} onClick={() => activateTool('highlighter', false)}>
                    Mark
                  </button>
                  <button type="button" className={`mini-tool-button ${activeTool === 'eraser' ? 'selected' : ''}`} onClick={() => activateTool('eraser', false)}>
                    Erase
                  </button>
                  <button type="button" className={`mini-tool-button ${activeTool === 'shape' ? 'selected' : ''}`} onClick={() => activateTool('shape', false)}>
                    Shape
                  </button>
                  <button type="button" className={`mini-tool-button ${activeTool === 'lasso' ? 'selected' : ''}`} onClick={() => activateTool('lasso', false)}>
                    Select
                  </button>
                  <button type="button" className="mini-tool-button" onClick={handleUndo} disabled={!activePageAnnotations.strokes.length}>
                    Undo
                  </button>
                </div>
              ) : null}

              <div className={`viewer-layout notes-layout-${notesLayout}`}>
                {notesLayout === 'classic' ? (
                  <aside className="page-thumbnail-strip" aria-label="Page thumbnails">
                  <div className="thumbnail-strip-head">
                    <span>Pages</span>
                    <strong>{pageCount}</strong>
                  </div>

                  <div className="thumbnail-list">
                    {pages.map((page) => {
                      const pageStrokeCount = pageStrokeCountMap[page.pageNumber] ?? 0;

                      return (
                        <button
                          key={`thumb-${page.pageNumber}`}
                          type="button"
                          className={`thumbnail-card ${activePageNumber === page.pageNumber ? 'active' : ''}`}
                          onClick={() => scrollToPage(page.pageNumber)}
                          aria-label={`Go to page ${page.pageNumber}`}
                        >
                          <span className={`thumbnail-preview ${page.kind === 'pdf' ? 'pdf' : 'notebook'}`}>
                            <span className="thumbnail-page-number">{page.pageNumber}</span>
                          </span>
                          <span className="thumbnail-meta">
                            <strong>Page {page.pageNumber}</strong>
                            <span>{page.kind === 'pdf' ? 'PDF page' : 'Notebook page'}</span>
                            <span>{pageStrokeCount ? `${pageStrokeCount} stroke${pageStrokeCount > 1 ? 's' : ''}` : 'No ink yet'}</span>
                          </span>
                        </button>
                      );
                    })}
                  </div>
                  </aside>
                ) : null}

                <div className={`viewer-shell ${splitViewEnabled ? 'split-view' : ''} ${scrollDirection === 'horizontal' ? 'scroll-horizontal' : ''}`} ref={viewerShellRef} onScroll={handleViewerScroll}>
                  <div className="page-stack" style={{ transform: `scale(${viewScale}) rotate(${rotationDegrees}deg)` }}>
                    {pages.map((page) => {
                      const isInRenderWindow = renderWindowPageNumbers.includes(page.pageNumber);

                      return (
                        <section
                          key={page.pageNumber}
                          ref={(node) => {
                            pageStageRefs.current[page.pageNumber] = node;
                          }}
                          data-page-number={page.pageNumber}
                          className={`page-stage ${activePageNumber === page.pageNumber ? 'active' : ''}`}
                          onPointerEnter={() => setActivePageNumber(page.pageNumber)}
                          onClick={() => setActivePageNumber(page.pageNumber)}
                        >
                          <div className="page-badge">Page {page.pageNumber}</div>
                          <div className="stack-stage" style={{ width: `${page.width}px`, height: `${page.height}px`, '--page-bg-color': pageBackgroundColor }}>
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
                              <div className={`stack-layer notebook-layer template-${paperTemplate}`} aria-label={`Blank ruled page ${page.pageNumber}`} />
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
                <strong>{activeTool === 'highlighter' ? 'Highlighter active' : activeTool === 'eraser' ? 'Stroke eraser active' : activeTool === 'shape' ? `Shape tool: ${shapeType}` : activeTool === 'lasso' ? 'Lasso select active' : 'Pen active'}</strong>
              </div>
              <div>
                <span className="meta-label">Voice transcript</span>
                <strong>{lastVoiceTranscript || 'No voice command captured yet'}</strong>
              </div>
              <div>
                <span className="meta-label">Persistence</span>
                <strong>{storageLabel}</strong>
              </div>
              <div>
                <span className="meta-label">Ink summary</span>
                <strong>{totalStrokeCount} saved strokes</strong>
              </div>
            </div>
          </>
        )}

        {showOnboardingTour ? (
          <div className="onboarding-backdrop" role="dialog" aria-label="Welcome tour">
            <div className="onboarding-card">
              <p className="eyebrow">Welcome</p>
              <h2>Build your workflow in VoxNotes</h2>
              <p>Start with the sample notebook, tune your toolbar shortcuts, and bookmark key pages for quick revision.</p>
              <div className="onboarding-actions">
                <button type="button" className="overlay-button" onClick={launchSampleNotebook}>Open Sample</button>
                <button
                  type="button"
                  className="overlay-button compact"
                  onClick={() => {
                    setShowOnboardingTour(false);
                    if (typeof window !== 'undefined') {
                      window.localStorage.setItem(ONBOARDING_STORAGE_KEY, 'true');
                    }
                  }}
                >
                  Skip tour
                </button>
              </div>
            </div>
          </div>
        ) : null}

        {showNotebookSettings ? (
          <div className="settings-sheet-backdrop" role="dialog" aria-label="Notebook settings" onClick={() => setShowNotebookSettings(false)}>
            <section className="settings-sheet" onClick={(event) => event.stopPropagation()}>
              <div className="settings-sheet-head">
                <button type="button" className="overlay-button compact" onClick={() => setShowNotebookSettings(false)}>
                  <ChevronLeft size={16} />
                </button>
                <h3>Page template</h3>
              </div>

              <label className="sheet-toggle-row">
                <span>Apply to all pages</span>
                <input type="checkbox" checked={applyTemplateToAllPages} onChange={(event) => setApplyTemplateToAllPages(event.target.checked)} />
              </label>

              <div className="settings-sheet-section">
                <div className="settings-section-title-row">
                  <strong>Default templates</strong>
                </div>
                <div className="template-grid">
                  {PAGE_TEMPLATE_OPTIONS.map((templateOption) => (
                    <button
                      key={templateOption.id}
                      type="button"
                      className={`template-card ${paperTemplate === templateOption.id ? 'selected' : ''}`}
                      onClick={() => applyTemplateSelection(templateOption.id)}
                    >
                      <span className={`template-preview template-${templateOption.id}`} />
                    </button>
                  ))}
                </div>
              </div>

              <div className="settings-sheet-section">
                <strong>Scroll direction</strong>
                <div className="segmented-control">
                  <button type="button" className={scrollDirection === 'vertical' ? 'selected' : ''} onClick={() => setScrollDirection('vertical')}>Vertical</button>
                  <button type="button" className={scrollDirection === 'horizontal' ? 'selected' : ''} onClick={() => setScrollDirection('horizontal')}>Horizontal</button>
                </div>
              </div>

              <div className="settings-sheet-section">
                <strong>Background colour</strong>
                <div className="background-swatch-row">
                  {PAGE_BACKGROUND_SWATCHES.map((swatch) => (
                    <button
                      key={swatch}
                      type="button"
                      className={`background-swatch ${pageBackgroundColor === swatch ? 'selected' : ''}`}
                      style={{ backgroundColor: swatch }}
                      onClick={() => setPageBackgroundColor(swatch)}
                    />
                  ))}
                </div>
              </div>

              <div className="settings-sheet-section">
                <strong>Cover</strong>
                <div className="cover-grid">
                  {COVER_VARIANTS.map((coverOption) => (
                    <button
                      key={coverOption.id}
                      type="button"
                      className={`cover-card ${coverOption.className} ${coverVariant === coverOption.id ? 'selected' : ''}`}
                      onClick={() => setCoverVariant(coverOption.id)}
                      aria-label={coverOption.label}
                    />
                  ))}
                </div>
              </div>
            </section>
          </div>
        ) : null}

        <div className="toast-tray" aria-live="polite" aria-label="Notifications">
          {toastItems.slice(-4).map((toastItem) => (
            <div key={toastItem.id} className="toast-item">
              <span>{toastItem.message}</span>
              <button type="button" className="mini-icon-button" onClick={() => setToastItems((currentItems) => currentItems.filter((candidate) => candidate.id !== toastItem.id))}>×</button>
            </div>
          ))}
        </div>

        {voiceTranscriptToast ? (
          <div className="voice-transcript-toast" aria-live="polite">
            <span>{voiceTranscriptToast}</span>
          </div>
        ) : null}

        <input ref={fileInputRef} type="file" accept=".pdf,application/pdf" hidden onChange={handleFileChange} />
      </main>
    </div>
  );
}

export default App;