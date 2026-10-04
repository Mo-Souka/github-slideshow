import * as React from 'react';
import { MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
import { IAppConfig } from '../../../models/IFieldConfig';
import { IUserContext } from '../../../models/IUserContext';
import { AppContext, IAppContext, IAppServices, INotification } from './AppContext';
import { Route, toHash, useHashRoute } from './useHashRoute';
import { useContainerWidth } from './useContainerWidth';
import { ErrorMessage } from './common/ErrorMessage';
import { ErrorBoundary } from './common/ErrorBoundary';
import { RecordForm } from './form/RecordForm';
import { RecordList } from './list/RecordList';
import { RecordDetail } from './detail/RecordDetail';
import styles from './App.module.scss';

export interface IAppProps {
  config: IAppConfig;
  services: IAppServices;
}

const MESSAGE_BAR_TYPE: Record<INotification['type'], MessageBarType> = {
  success: MessageBarType.success,
  info: MessageBarType.info,
  warning: MessageBarType.warning
};

export const App: React.FC<IAppProps> = ({ config, services }) => {
  const rootRef = React.useRef<HTMLDivElement>(null);
  const width = useContainerWidth(rootRef);
  const [route, navigate] = useHashRoute();
  const [user, setUser] = React.useState<IUserContext | undefined>();
  const [initError, setInitError] = React.useState<unknown>();
  const [attempt, setAttempt] = React.useState(0);
  const [notification, setNotification] = React.useState<INotification | undefined>();
  const notificationTarget = React.useRef<string | undefined>();
  const isFirstRoute = React.useRef(true);

  React.useEffect(() => {
    let cancelled = false;
    setInitError(undefined);
    Promise.all([services.records.ensureListExists(), services.users.getCurrentUser()])
      .then(([, currentUser]) => {
        if (!cancelled) setUser(currentUser);
      })
      .catch((error) => {
        if (!cancelled) setInitError(error);
      });
    return () => {
      cancelled = true;
    };
  }, [services, attempt]);

  // A notification is shown on the screen it was meant for and cleared when the user moves on.
  React.useEffect(() => {
    const hash = toHash(route);
    if (notificationTarget.current === hash) {
      notificationTarget.current = undefined;
    } else {
      setNotification(undefined);
    }
    // Bring the top of the web part into view when switching screens (not on first load).
    if (isFirstRoute.current) {
      isFirstRoute.current = false;
    } else if (rootRef.current && rootRef.current.getBoundingClientRect().top < 0) {
      rootRef.current.scrollIntoView({ block: 'start' });
    }
  }, [route]);

  const notify = React.useCallback(
    (next: INotification, then?: Route) => {
      setNotification(next);
      if (then) {
        notificationTarget.current = toHash(then);
        navigate(then);
      }
    },
    [navigate]
  );

  let content: React.ReactNode;
  if (initError) {
    content = <ErrorMessage error={initError} onRetry={() => setAttempt(attempt + 1)} />;
  } else if (!user) {
    content = <Spinner className={styles.loading} size={SpinnerSize.large} label="Loading..." />;
  } else {
    switch (route.name) {
      case 'new':
        content = <RecordForm key="new" />;
        break;
      case 'edit':
        content = <RecordForm key={`edit-${route.id}`} recordId={route.id} />;
        break;
      case 'view':
        content = <RecordDetail key={`view-${route.id}`} recordId={route.id} />;
        break;
      case 'pending':
        content = <RecordList key="pending" pendingOnly={true} />;
        break;
      default:
        content = <RecordList key="list" pendingOnly={false} />;
        break;
    }
  }

  const context: IAppContext | undefined = user ? { config, user, services, navigate, notify, width } : undefined;

  return (
    <div className={styles.app} ref={rootRef}>
      <ErrorBoundary>
        {notification && (
          <div className={styles.notification}>
            <MessageBar
              messageBarType={MESSAGE_BAR_TYPE[notification.type]}
              isMultiline={true}
              onDismiss={() => setNotification(undefined)}
              dismissButtonAriaLabel="Close"
            >
              {notification.text}
            </MessageBar>
          </div>
        )}
        {context ? <AppContext.Provider value={context}>{content}</AppContext.Provider> : content}
      </ErrorBoundary>
    </div>
  );
};
