import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export interface JournalEntry {
  id: string;
  date: string; // MM/DD/YYYY
  time: string;
  journal: string;
  speechToText: string;
  aiReport: string;
  /** YYYY-MM-DD from sheet tab; used for stable sort */
  isoDate?: string;
  rowIndex?: number;
}

export interface PositionedSketch {
  id: string;
  afterEntryId: string;
  dataUrl: string;
  createdAt: string;
}

export type RightPageSetting = 'none' | 'ai' | 'stt';

export type JournalSection = 'journal' | 'stt' | 'ai';
