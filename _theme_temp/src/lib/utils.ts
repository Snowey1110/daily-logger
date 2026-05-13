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
  sketch?: string; // Data URL for canvas sketch
}

export type JournalSection = 'journal' | 'stt' | 'ai';
