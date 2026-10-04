import { AppError } from '../models/AppError';

export interface IErrorContext {
  /** What the user was doing, e.g. "save the record". Used in permission messages. */
  action?: string;
  /** Title of the list or library involved, used in "not found" messages. */
  listTitle?: string;
}

interface IHttpLikeError {
  isHttpRequestError?: boolean;
  status?: number;
  message?: string;
}

/** Extracts SharePoint's own error text from a PnPjs HttpRequestError message. */
export function extractSharePointMessage(rawMessage: string): string {
  const marker = '::>';
  const index = rawMessage.indexOf(marker);
  const body = index >= 0 ? rawMessage.substring(index + marker.length).trim() : rawMessage;
  try {
    const parsed = JSON.parse(body);
    const odataError = parsed['odata.error'] || parsed.error;
    if (odataError && odataError.message) {
      const message = odataError.message;
      return typeof message === 'string' ? message : String(message.value || '');
    }
  } catch {
    // Not JSON; fall through.
  }
  return body;
}

function isNetworkFailure(error: unknown): boolean {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
    return true;
  }
  const message = error instanceof Error ? error.message : String(error);
  return /failed to fetch|networkerror|network request failed|load failed/i.test(message);
}

/**
 * Converts anything thrown by PnPjs, fetch or our own code into an AppError
 * with a message a warehouse user can act on.
 */
export function toAppError(error: unknown, context: IErrorContext = {}): AppError {
  if (error instanceof AppError) {
    return error;
  }

  const action = context.action || 'do this';
  const httpError = error as IHttpLikeError;

  if (httpError && httpError.isHttpRequestError && typeof httpError.status === 'number') {
    const spMessage = extractSharePointMessage(httpError.message || '');
    const status = httpError.status;

    if (status === 401 || status === 403) {
      return new AppError(
        'PermissionDenied',
        `You don't have permission to ${action}. Ask your site owner to add you to the right SharePoint group.`,
        spMessage
      );
    }
    if (status === 404) {
      if (/list .* does not exist|list does not exist/i.test(spMessage) || /does not exist at site/i.test(spMessage)) {
        const name = context.listTitle ? `"${context.listTitle}"` : 'needed by this app';
        return new AppError(
          'ListNotFound',
          `The SharePoint list or library ${name} was not found on this site. Ask your administrator to run the provisioning script (see DEPLOYMENT.md) or check the web part settings.`,
          spMessage
        );
      }
      return new AppError('NotFound', 'The item was not found. It may have been deleted or moved.', spMessage);
    }
    if (status === 412) {
      return new AppError(
        'Conflict',
        'Someone else changed this record after you opened it. Reload the record to see the latest version, then make your changes again.',
        spMessage
      );
    }
    if (status === 429 || status === 503) {
      return new AppError('Throttled', 'SharePoint is busy right now. Please wait a minute and try again.', spMessage);
    }
    if (/SPQueryThrottledException|exceeds the list view threshold|list view threshold/i.test(spMessage)) {
      return new AppError(
        'ThresholdExceeded',
        'Too many records match this search for SharePoint to process. Choose a shorter date range or add more filters.',
        spMessage
      );
    }
    if (status === 400 && /column .* does not exist|field or property .* does not exist|does not exist on type/i.test(spMessage)) {
      return new AppError(
        'ConfigurationMismatch',
        'A column the app needs is missing in SharePoint. Ask your administrator to re-run the provisioning script.',
        spMessage
      );
    }
    if (status === 400 && /validation|value does not fall within the expected range|invalid/i.test(spMessage)) {
      return new AppError('Validation', `SharePoint rejected the data: ${spMessage}`, spMessage);
    }
    return new AppError('Unknown', `SharePoint returned an error (${status}). ${spMessage}`.trim(), spMessage);
  }

  if (isNetworkFailure(error)) {
    return new AppError(
      'Network',
      'Cannot reach SharePoint. Check the network connection (Wi-Fi) and try again.',
      error instanceof Error ? error.message : String(error)
    );
  }

  const message = error instanceof Error ? error.message : String(error);
  return new AppError('Unknown', `Something went wrong: ${message}`, message);
}
