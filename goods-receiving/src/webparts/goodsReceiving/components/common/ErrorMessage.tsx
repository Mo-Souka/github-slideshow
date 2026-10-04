import * as React from 'react';
import { MessageBar, MessageBarType, Link, MessageBarButton } from '@fluentui/react';
import { AppError } from '../../../../models/AppError';
import { toAppError } from '../../../../services/errors';

export interface IErrorMessageProps {
  error: unknown;
  onRetry?: () => void;
  onDismiss?: () => void;
}

/** Shows an error in plain language, with technical details on request. */
export const ErrorMessage: React.FC<IErrorMessageProps> = ({ error, onRetry, onDismiss }) => {
  const [showDetails, setShowDetails] = React.useState(false);
  if (!error) return null;
  const appError: AppError = toAppError(error);
  const isWarning = appError.kind === 'Validation' || appError.kind === 'Conflict' || appError.kind === 'ThresholdExceeded';

  return (
    <MessageBar
      messageBarType={isWarning ? MessageBarType.warning : MessageBarType.error}
      isMultiline={true}
      onDismiss={onDismiss}
      dismissButtonAriaLabel="Close"
      actions={onRetry ? <MessageBarButton onClick={onRetry}>Try again</MessageBarButton> : undefined}
    >
      {appError.userMessage}
      {appError.details && appError.details !== appError.userMessage && (
        <>
          {' '}
          <Link onClick={() => setShowDetails(!showDetails)}>{showDetails ? 'Hide details' : 'Details'}</Link>
          {showDetails && <div style={{ marginTop: 4, whiteSpace: 'pre-wrap', fontSize: 12 }}>{appError.details}</div>}
        </>
      )}
    </MessageBar>
  );
};
