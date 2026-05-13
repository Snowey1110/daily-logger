
export interface JournalTheme {
  id: string;
  name: string;
  colors: {
    bg: string;
    nav: string;
    bookInner: string;
    text: string;
    textMuted: string;
    border: string;
    spine: string;
    tabs: {
      journal: { bg: string; active: string };
      stt: { bg: string; active: string };
      ai: { bg: string; active: string };
    };
  };
  texture: string;
}
