import {
  createContext,
  useCallback,
  useContext,
  useMemo,
  type ReactNode,
} from 'react';

export type ReaderLang = 'en' | 'zh';

const EN = {
  readerSubtitle: 'Virtual Journal Reader',
  docTitleSuffix: 'Virtual Journal Reader',
  errLoadData: 'Could not load journal data.',
  errSaveFailed: 'Save failed.',
  errNetworkSave: 'Network error while saving.',
  errSketchSave: 'Could not save sketch.',
  errNetworkSketch: 'Network error while saving sketch.',
  navPage: 'Page',
  navJumpDate: 'Jump to date',
  navPenOptions: 'Pen Options',
  navDrawSketch: 'Insert',
  navEditJournal: 'Edit Journal',
  navEditSTT: 'Edit speech transcript',
  navEditAI: 'Edit AI report',
  navSaveChanges: 'Save changes',
  navSortOldest: 'Oldest first',
  navSortNewest: 'Newest first',
  navSortPrefix: 'Sort:',
  coverSubtitle: 'Virtual Journal',
  coverBegin: 'Interact to begin',
  tabJournal: 'JOURNAL',
  tabSpeech: 'SPEECH',
  tabAi: 'AI REPORT',
  footerPrev: 'Previous',
  footerNext: 'Next',
  footerFlipHint: 'Use A / D or Arrows to flip pages',
  ariaPrevPage: 'Previous page',
  ariaNextPage: 'Next page',
  pageEmpty: 'The rest is still unwritten',
  pageSketchCaption: 'Digital Sketch',
  pageSketchAlt: 'Full page sketch',
  pageBlank: '(This space intentionally left blank)',
  pageJournalEntry: 'Journal entry',
  pageDailyReflection: 'Daily Reflection',
  pageJournal: 'Journal',
  pageVoiceTranscript: 'Voice Transcript',
  pageIntelAnalysis: 'Intelligence Analysis',
  pageContSuffix: '(cont.)',
  pagePlaceholderEdit: 'Edit text…',
  pageQuote: '“The machine sees what we often overlook…”',
  sketchpad: 'Sketchpad',
  clearCanvas: 'Clear Canvas',
  saveSketch: 'Save Sketch',
  settingsTitle: 'Settings',
  settingsCoverTheme: 'Cover Theme',
  settingsBgTheme: 'Background Theme',
  themePickerTitle: 'Choose Theme',
  themePickerCover: 'Cover Theme',
  themePickerBackground: 'Background Theme',
  sketchPlacerTitle: 'Insert Manager',
  sketchInsertHere: 'Click to insert here',
  sketchDeleteConfirm: 'Delete this sketch?',
  sketchDeleteBtn: 'Delete',
  sketchCancelBtn: 'Cancel',
  insertChoicePage: 'New Page',
  insertChoiceSketch: 'New Sketch',
  insertPageDateLabel: 'Date',
  insertPageTimeLabel: 'Time',
  insertPageCreate: 'Create',
  errCreatePage: 'Could not create page.',
  deleteEntry: 'Delete',
  deleteEntryConfirm: 'Delete this entry?',
  errDeleteEntry: 'Could not delete entry.',
  rightTabSketch: 'Sketch',
  rightTabAi: 'AI Report',
  rightTabStt: 'Speech',
} as const;

const ZH: { [K in keyof typeof EN]: string } = {
  readerSubtitle: '阅读器',
  docTitleSuffix: '阅读器',
  errLoadData: '无法加载日记数据。',
  errSaveFailed: '保存失败。',
  errNetworkSave: '保存时出现网络错误。',
  errSketchSave: '无法保存涂鸦。',
  errNetworkSketch: '保存涂鸦时出现网络错误。',
  navPage: '页',
  navJumpDate: '按日期跳转',
  navPenOptions: '笔选项',
  navDrawSketch: '插入',
  navEditJournal: '编辑日记',
  navEditSTT: '编辑语音转写',
  navEditAI: '编辑 AI 报告',
  navSaveChanges: '保存更改',
  navSortOldest: '最旧在前',
  navSortNewest: '最新在前',
  navSortPrefix: '排序：',
  coverSubtitle: '电子日记',
  coverBegin: '点击开始',
  tabJournal: '日记',
  tabSpeech: '语音转写',
  tabAi: 'AI 报告',
  footerPrev: '上一页',
  footerNext: '下一页',
  footerFlipHint: '使用 A / D 或方向键翻页',
  ariaPrevPage: '上一页',
  ariaNextPage: '下一页',
  pageEmpty: '余白待书',
  pageSketchCaption: '数字涂鸦',
  pageSketchAlt: '整页涂鸦',
  pageBlank: '（留白）',
  pageJournalEntry: '日记条目',
  pageDailyReflection: '当日随想',
  pageJournal: '日记',
  pageVoiceTranscript: '语音转写',
  pageIntelAnalysis: '智能分析',
  pageContSuffix: '（续）',
  pagePlaceholderEdit: '编辑正文…',
  pageQuote: '「机器所见，常为吾辈所忽……」',
  sketchpad: '涂鸦板',
  clearCanvas: '清空画布',
  saveSketch: '保存涂鸦',
  settingsTitle: '设置',
  settingsCoverTheme: '封面主题',
  settingsBgTheme: '背景主题',
  themePickerTitle: '选择主题',
  themePickerCover: '封面主题',
  themePickerBackground: '背景主题',
  sketchPlacerTitle: '插入管理',
  sketchInsertHere: '点击此处插入',
  sketchDeleteConfirm: '确定删除此涂鸦？',
  sketchDeleteBtn: '删除',
  sketchCancelBtn: '取消',
  insertChoicePage: '新页面',
  insertChoiceSketch: '新涂鸦',
  insertPageDateLabel: '日期',
  insertPageTimeLabel: '时间',
  insertPageCreate: '创建',
  errCreatePage: '无法创建页面。',
  deleteEntry: '删除',
  deleteEntryConfirm: '确定删除此条目？',
  errDeleteEntry: '无法删除条目。',
  rightTabSketch: '涂鸦',
  rightTabAi: 'AI 报告',
  rightTabStt: '语音',
};

export type ReaderStringKey = keyof typeof EN;

function langFromSearch(search: string): ReaderLang {
  const raw = new URLSearchParams(search).get('lang')?.trim().toLowerCase() ?? '';
  if (raw === 'zh' || raw === 'zh-cn' || raw === 'cn' || raw === 'chinese' || raw === '中文') {
    return 'zh';
  }
  return 'en';
}

type ReaderI18nValue = {
  lang: ReaderLang;
  t: (key: ReaderStringKey) => string;
};

const ReaderI18nContext = createContext<ReaderI18nValue | null>(null);

export function ReaderI18nProvider({ children }: { children: ReactNode }) {
  const lang = useMemo(() => langFromSearch(window.location.search), []);
  const table = lang === 'zh' ? ZH : EN;
  const t = useCallback((key: ReaderStringKey) => table[key], [table]);
  const value = useMemo(() => ({ lang, t }), [lang, t]);
  return (
    <ReaderI18nContext.Provider value={value}>{children}</ReaderI18nContext.Provider>
  );
}

export function useReaderT(): ReaderI18nValue {
  const ctx = useContext(ReaderI18nContext);
  if (!ctx) {
    throw new Error('useReaderT must be used within ReaderI18nProvider');
  }
  return ctx;
}
