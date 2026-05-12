import { JournalEntry } from '../lib/utils';

export const MOCK_ENTRIES: JournalEntry[] = [
  {
    id: '1',
    date: '05/10/2026',
    time: '09:00 AM',
    journal: 'Started the morning with a fresh cup of coffee and a clear mind. Working on the new journal reader UI today. The goal is to make it feel tangible and intuitive.',
    speechToText: 'I am starting the morning. Coffee is good. Working on the UI. Hope it looks like a real book.',
    aiReport: 'Today the user focuses on UI development and ritualistic starting habits. Sentiment is positive and productive.',
  },
  {
    id: '2',
    date: '05/11/2026',
    time: '10:30 PM',
    journal: 'Evening reflection. The flipping animations are coming together. It is interesting how physical metaphors in software can make things feel more personal.',
    speechToText: 'It is late now. Animations are working. Metaphors are cool.',
    aiReport: 'User is reflecting on design philosophy late at night. Shows interest in skeuomorphic tendencies in modern web apps.',
  },
  {
    id: '3',
    date: '05/12/2026',
    time: '02:15 PM',
    journal: 'Finalizing the search and calendar features. Being able to jump through time in a journal is the key utility.',
    speechToText: 'Finalizing search and calendar. Jumping through time is important.',
    aiReport: 'User is focusing on utility and navigation efficiency. The project is nearing completion.',
  }
];
