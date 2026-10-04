import * as React from 'react';

/**
 * Minimal hash-based navigation so that records have shareable links
 * (e.g. .../SitePages/Receiving.aspx#gr/record/123) and the browser Back
 * button works, without interfering with the SharePoint page itself.
 */
export type Route =
  | { name: 'list' }
  | { name: 'pending' }
  | { name: 'new' }
  | { name: 'view'; id: number }
  | { name: 'edit'; id: number };

const PREFIX = '#gr/';

export function parseHash(hash: string): Route {
  if (!hash || hash.indexOf(PREFIX) !== 0) return { name: 'list' };
  const parts = hash.substring(PREFIX.length).split('/');
  if (parts[0] === 'new') return { name: 'new' };
  if (parts[0] === 'pending') return { name: 'pending' };
  if (parts[0] === 'record') {
    const id = parseInt(parts[1], 10);
    if (id > 0) return parts[2] === 'edit' ? { name: 'edit', id } : { name: 'view', id };
  }
  return { name: 'list' };
}

export function toHash(route: Route): string {
  switch (route.name) {
    case 'new':
      return `${PREFIX}new`;
    case 'pending':
      return `${PREFIX}pending`;
    case 'view':
      return `${PREFIX}record/${route.id}`;
    case 'edit':
      return `${PREFIX}record/${route.id}/edit`;
    default:
      return `${PREFIX}list`;
  }
}

export function useHashRoute(): [Route, (route: Route) => void] {
  const [route, setRoute] = React.useState<Route>(() => parseHash(window.location.hash));

  React.useEffect(() => {
    const onHashChange = (): void => setRoute(parseHash(window.location.hash));
    window.addEventListener('hashchange', onHashChange);
    return () => window.removeEventListener('hashchange', onHashChange);
  }, []);

  const navigate = React.useCallback((next: Route) => {
    const hash = toHash(next);
    if (window.location.hash === hash) {
      setRoute(parseHash(hash));
    } else {
      window.location.hash = hash;
    }
  }, []);

  return [route, navigate];
}
