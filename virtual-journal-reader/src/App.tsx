import JournalBook from './components/JournalBook';
import { ReaderI18nProvider } from './readerI18n';

export default function App() {
  return (
    <ReaderI18nProvider>
      <div className="w-full h-screen overflow-hidden">
        <JournalBook />
      </div>
    </ReaderI18nProvider>
  );
}
