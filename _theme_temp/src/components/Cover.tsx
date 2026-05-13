import React from 'react';
import { motion } from 'motion/react';
import { Book as BookIcon } from 'lucide-react';

import { cn } from '../lib/utils';
import { JournalTheme } from '../types/theme';

interface CoverProps {
  title: string;
  onClick: () => void;
  theme: JournalTheme;
}

const Cover: React.FC<CoverProps> = ({ title, onClick, theme }) => {
  const isDark = theme.id === 'midnight' || theme.id === 'royal';

  return (
    <motion.div
      onClick={onClick}
      className={cn(
        "relative w-full h-full cursor-pointer group flex items-center justify-center overflow-hidden rounded-r-xl border-l-[12px] transition-colors duration-500",
        isDark ? "border-zinc-800" : "border-zinc-950"
      )}
      initial={{ rotateY: 0 }}
      whileHover={{ scale: 1.01 }}
      transition={{ type: 'spring', stiffness: 300, damping: 20 }}
      style={{
        backgroundColor: theme.id === 'parchment' ? '#2D3436' : (theme.id === 'midnight' ? '#0F172A' : (theme.id === 'botanist' ? '#3A5A40' : '#4A3B3B')),
        backgroundImage: theme.id === 'parchment' 
          ? 'linear-gradient(135deg, #2d3436 0%, #000000 100%)' 
          : (theme.id === 'midnight' 
            ? 'linear-gradient(135deg, #1e293b 0%, #0f172a 100%)' 
            : (theme.id === 'botanist' 
              ? 'linear-gradient(135deg, #3a5a40 0%, #1b261a 100%)' 
              : 'linear-gradient(135deg, #4a3b3b 0%, #1a0f0f 100%)')),
        boxShadow: 'inset -20px 0 30px rgba(0,0,0,0.5), 10px 10px 30px rgba(0,0,0,0.3)',
        transformStyle: 'preserve-3d'
      }}
    >
      {/* TextureOverlay */}
      <div className="absolute inset-0 opacity-10 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/leather.png')]" />

      {/* Golden Accents */}
      <div className={cn(
        "absolute inset-4 border rounded-lg pointer-events-none transition-colors duration-500",
        isDark ? "border-indigo-500/20" : "border-yellow-600/30"
      )} />
      <div className={cn(
        "absolute inset-8 border-2 rounded-lg pointer-events-none transition-colors duration-500",
        isDark ? "border-indigo-500/10" : "border-yellow-600/10"
      )} />

      <div className="flex flex-col items-center gap-8 z-10 text-center px-8">
        <div className={cn(
          "w-24 h-24 rounded-full flex items-center justify-center border-2 transition-all duration-500 shadow-lg",
          isDark ? "bg-indigo-600/20 border-indigo-500/40 shadow-indigo-600/10" : "bg-yellow-600/20 border-yellow-600/40 shadow-yellow-600/10"
        )}>
          <BookIcon className={isDark ? "text-indigo-400" : "text-yellow-500"} size={48} strokeWidth={1.5} />
        </div>

        <div className="space-y-3">
          <h1 className={cn(
            "text-4xl font-serif font-bold tracking-widest uppercase drop-shadow-md transition-colors duration-500",
            isDark ? "text-slate-100" : "text-yellow-500"
          )}>
            {title || "Daily Logger"}
          </h1>
          <div className={cn(
            "h-[2px] w-24 mx-auto transition-colors duration-500",
            isDark ? "bg-indigo-500/40" : "bg-yellow-600/40"
          )} />
          <p className={cn(
            "font-serif italic text-sm tracking-widest transition-colors duration-500",
            isDark ? "text-slate-400/60" : "text-yellow-600/60"
          )}>
            VIRTUAL JOURNAL ADDON
          </p>
        </div>
      </div>

      <div className="absolute bottom-12 right-0 left-0 text-center">
        <p className="text-zinc-500 font-sans text-[10px] tracking-tighter uppercase opacity-30 mt-4">
          Click cover to open or use arrow keys
        </p>
      </div>
      
      {/* Spine highlight */}
      <div className="absolute top-0 left-0 bottom-0 w-2 bg-white/5" />
    </motion.div>
  );
};

export default Cover;
