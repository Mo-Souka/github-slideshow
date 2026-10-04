export type AppErrorKind =
  | 'PermissionDenied'
  | 'ListNotFound'
  | 'NotFound'
  | 'Conflict'
  | 'Throttled'
  | 'ThresholdExceeded'
  | 'ConfigurationMismatch'
  | 'Network'
  | 'Validation'
  | 'Unknown';

/**
 * Error with a message that can be shown to warehouse users as-is.
 * `details` keeps the technical message for support staff.
 */
export class AppError extends Error {
  public readonly kind: AppErrorKind;
  public readonly userMessage: string;
  public readonly details?: string;

  public constructor(kind: AppErrorKind, userMessage: string, details?: string) {
    super(userMessage);
    // Required for `instanceof` to work when compiling to ES5.
    Object.setPrototypeOf(this, AppError.prototype);
    this.name = 'AppError';
    this.kind = kind;
    this.userMessage = userMessage;
    this.details = details;
  }
}
