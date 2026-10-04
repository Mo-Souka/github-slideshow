import * as React from 'react';

/**
 * Width of the web part itself (not the browser window). A web part can sit in
 * a narrow column on a desktop page, so the layout follows its own width.
 */
export function useContainerWidth(ref: React.RefObject<HTMLElement>): number {
  const [width, setWidth] = React.useState<number>(() => (ref.current ? ref.current.offsetWidth : 1024));

  React.useEffect(() => {
    const element = ref.current;
    if (!element) return undefined;
    setWidth(element.offsetWidth);

    if (typeof ResizeObserver === 'undefined') {
      const onResize = (): void => setWidth(element.offsetWidth);
      window.addEventListener('resize', onResize);
      return () => window.removeEventListener('resize', onResize);
    }
    const observer = new ResizeObserver((entries) => {
      const entry = entries[0];
      if (entry) setWidth(Math.round(entry.contentRect.width));
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, [ref]);

  return width;
}
