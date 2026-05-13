import JournalBook from './components/JournalBook';
import { ReaderI18nProvider } from './readerI18n';
import { ThemeProvider } from './components/ThemeProvider';

export default function App() {
  return (
    <ReaderI18nProvider>
      <ThemeProvider>
        <div className="w-full h-screen overflow-hidden">
          <JournalBook />
        </div>
      </ThemeProvider>
    </ReaderI18nProvider>
  );
}
