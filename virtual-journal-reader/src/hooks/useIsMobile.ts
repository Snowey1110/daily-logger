import { useState, useEffect } from 'react';

const MD_BREAKPOINT = 768;

interface MobileState {
  isMobile: boolean;
  isLandscape: boolean;
}

function measure(): MobileState {
  return {
    isMobile: window.innerWidth < MD_BREAKPOINT,
    isLandscape: window.innerWidth > window.innerHeight,
  };
}

export function useIsMobile(): MobileState {
  const [state, setState] = useState<MobileState>(measure);

  useEffect(() => {
    let raf = 0;
    const update = () => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => setState(measure()));
    };
    window.addEventListener('resize', update);
    window.addEventListener('orientationchange', update);
    return () => {
      cancelAnimationFrame(raf);
      window.removeEventListener('resize', update);
      window.removeEventListener('orientationchange', update);
    };
  }, []);

  return state;
}
