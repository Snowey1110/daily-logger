import JournalBook from './components/JournalBook';
import { ThemeProvider } from './components/ThemeProvider';

export default function App() {
  return (
    <ThemeProvider>
      <div className="w-full h-screen overflow-hidden">
        <JournalBook />
      </div>
    </ThemeProvider>
  );
}
