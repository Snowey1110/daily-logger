import React from 'react';
import { motion } from 'motion/react';
import { Book as BookIcon } from 'lucide-react';
import { useReaderT } from '../readerI18n';
import type { JournalTheme } from '../types/theme';

interface CoverProps {
  title: string;
  onClick: () => void;
  theme: JournalTheme;
}

const Cover: React.FC<CoverProps> = ({ title, onClick, theme }) => {
  const { t } = useReaderT();
  const c = theme.cover;

  return (
    <motion.div
      onClick={onClick}
      className="relative w-full h-full cursor-pointer group flex items-center justify-center overflow-hidden rounded-r-xl border-l-[12px] transition-colors duration-500"
      initial={{ rotateY: 0 }}
      whileHover={{ scale: 1.01 }}
      transition={{ type: 'spring', stiffness: 300, damping: 20 }}
      style={{
        borderColor: c.borderColor,
        backgroundImage: c.gradient,
        boxShadow: 'inset -20px 0 30px rgba(0,0,0,0.5), 20px 20px 60px rgba(0,0,0,0.4)',
        transformStyle: 'preserve-3d',
      }}
    >
      <div className="absolute inset-0 opacity-20 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/leather.png')]" />

      <div
        className="absolute inset-4 border rounded-lg pointer-events-none transition-colors duration-500"
        style={{ borderColor: c.accentBorder }}
      />
      <div
        className="absolute inset-8 border-2 rounded-lg pointer-events-none transition-colors duration-500"
        style={{ borderColor: `${c.accentBorder}50` }}
      />

      <div className="flex flex-col items-center gap-8 z-10 text-center px-8">
        <div
          className="w-24 h-24 rounded-full flex items-center justify-center border-2 group-hover:scale-110 transition-transform duration-500 shadow-lg"
          style={{ backgroundColor: c.accentBg, borderColor: c.accentBorder }}
        >
          <BookIcon style={{ color: c.accentText }} size={48} strokeWidth={1.5} />
        </div>

        <div className="space-y-4">
          <h1
            className="text-4xl font-serif font-light tracking-[0.2em] uppercase drop-shadow-lg"
            style={{ color: c.accentText }}
          >
            {title}
          </h1>
          <div className="h-[1px] w-24 mx-auto" style={{ backgroundColor: c.accentBorder }} />
          <p
            className="font-serif italic text-xs tracking-[0.3em] uppercase"
            style={{ color: c.subtitleText }}
          >
            {t('coverSubtitle')}
          </p>
        </div>
      </div>

      <div className="absolute bottom-12 right-0 left-0 text-center">
        <p style={{ color: c.subtitleText }} className="font-sans text-[9px] tracking-[0.4em] uppercase opacity-60">
          {t('coverBegin')}
        </p>
      </div>

      <div className="absolute top-0 left-0 bottom-0 w-3 bg-white/5 shadow-inner" />
    </motion.div>
  );
};

export default Cover;
