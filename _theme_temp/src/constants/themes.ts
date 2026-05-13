import { JournalTheme } from '../types/theme';

export const THEMES: JournalTheme[] = [
  {
    id: 'parchment',
    name: 'Classic Parchment',
    colors: {
      bg: 'bg-[#FDFCF0]',
      nav: 'bg-white/80',
      bookInner: 'bg-[#FCFBF4]',
      text: 'text-zinc-900',
      textMuted: 'text-zinc-400',
      border: 'border-zinc-200',
      spine: 'bg-zinc-800/20',
      tabs: {
        journal: { bg: 'bg-[#E9E1CC]', active: 'bg-[#D4A373]' },
        stt: { bg: 'bg-[#DDE5B6]', active: 'bg-[#A3B18A]' },
        ai: { bg: 'bg-[#CCD5AE]', active: 'bg-[#8A9A5B]' },
      }
    },
    texture: 'handmade-paper.png'
  },
  {
    id: 'midnight',
    name: 'Midnight Archive',
    colors: {
      bg: 'bg-[#0F172A]',
      nav: 'bg-[#1E293B]/80',
      bookInner: 'bg-[#1E293B]',
      text: 'text-slate-100',
      textMuted: 'text-slate-500',
      border: 'border-slate-700',
      spine: 'bg-indigo-500/20',
      tabs: {
        journal: { bg: 'bg-slate-700', active: 'bg-indigo-600' },
        stt: { bg: 'bg-slate-700', active: 'bg-emerald-600' },
        ai: { bg: 'bg-slate-700', active: 'bg-purple-600' },
      }
    },
    texture: 'carbon-fibre.png'
  },
  {
    id: 'botanist',
    name: 'Forest Botanist',
    colors: {
      bg: 'bg-[#E3E8E2]',
      nav: 'bg-[#F1F3F0]/80',
      bookInner: 'bg-[#F1F3F0]',
      text: 'text-stone-800',
      textMuted: 'text-stone-400',
      border: 'border-stone-300',
      spine: 'bg-green-800/20',
      tabs: {
        journal: { bg: 'bg-[#C2C5AA]', active: 'bg-[#6B705C]' },
        stt: { bg: 'bg-[#A3B18A]', active: 'bg-[#3A5A40]' },
        ai: { bg: 'bg-[#DDE5B6]', active: 'bg-[#588157]' },
      }
    },
    texture: 'natural-paper.png'
  },
  {
    id: 'royal',
    name: 'Royal Library',
    colors: {
      bg: 'bg-[#1A1A1A]',
      nav: 'bg-[#2A2A2A]/80',
      bookInner: 'bg-[#2A2323]',
      text: 'text-[#E5D5C5]',
      textMuted: 'text-[#8B7D7B]',
      border: 'border-[#4A3B3B]',
      spine: 'bg-red-800/30',
      tabs: {
        journal: { bg: 'bg-[#582F0E]', active: 'bg-[#936639]' },
        stt: { bg: 'bg-[#333D29]', active: 'bg-[#656D4A]' },
        ai: { bg: 'bg-[#414833]', active: 'bg-[#A68A64]' },
      }
    },
    texture: 'paper-fibers.png'
  }
];
