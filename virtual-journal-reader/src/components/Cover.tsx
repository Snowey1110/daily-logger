import React from 'react';
import { motion } from 'motion/react';
import { Book as BookIcon } from 'lucide-react';
import { useReaderT } from '../readerI18n';

interface CoverProps {
  title: string;
  onClick: () => void;
}

const Cover: React.FC<CoverProps> = ({ title, onClick }) => {
  const { t } = useReaderT();
  return (
    <motion.div
      onClick={onClick}
      className="relative w-full h-full cursor-pointer group flex items-center justify-center overflow-hidden rounded-r-xl border-l-[12px] border-[#1a130e]"
      initial={{ rotateY: 0 }}
      whileHover={{ scale: 1.01 }}
      transition={{ type: 'spring', stiffness: 300, damping: 20 }}
      style={{
        backgroundColor: '#3d2b1f',
        backgroundImage: 'linear-gradient(135deg, #3d2b1f 0%, #1a130e 100%)',
        boxShadow: 'inset -20px 0 30px rgba(0,0,0,0.5), 20px 20px 60px rgba(0,0,0,0.4)',
        transformStyle: 'preserve-3d'
      }}
    >
      {/* TextureOverlay */}
      <div className="absolute inset-0 opacity-20 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/leather.png')]" />

      {/* Golden Accents */}
      <div className="absolute inset-4 border border-[#d9c5b2]/20 rounded-lg pointer-events-none" />
      <div className="absolute inset-8 border-2 border-[#d9c5b2]/5 rounded-lg pointer-events-none" />

      <div className="flex flex-col items-center gap-8 z-10 text-center px-8">
        <div className="w-24 h-24 rounded-full bg-[#d9c5b2]/10 flex items-center justify-center border-2 border-[#d9c5b2]/30 group-hover:scale-110 transition-transform duration-500 shadow-lg shadow-black/20">
          <BookIcon className="text-[#d9c5b2]" size={48} strokeWidth={1.5} />
        </div>

        <div className="space-y-4">
          <h1 className="text-4xl font-serif text-[#d9c5b2] font-light tracking-[0.2em] uppercase drop-shadow-lg">
            {title}
          </h1>
          <div className="h-[1px] w-24 bg-[#d9c5b2]/20 mx-auto" />
          <p className="text-[#d9c5b2]/40 font-serif italic text-xs tracking-[0.3em] uppercase">
            {t('coverSubtitle')}
          </p>
        </div>
      </div>

      <div className="absolute bottom-12 right-0 left-0 text-center">
        <p className="text-[#d9c5b2]/20 font-sans text-[9px] tracking-[0.4em] uppercase">
          {t('coverBegin')}
        </p>
      </div>
      
      {/* Spine highlight */}
      <div className="absolute top-0 left-0 bottom-0 w-3 bg-white/5 shadow-inner" />
    </motion.div>
  );
};

export default Cover;
