/**
 * Runs a promise from an event handler. The promise is expected to handle its
 * own errors (show them in the UI); this only prevents unhandled rejections.
 */
export function fireAndForget(promise: Promise<unknown>): void {
  promise.catch(() => undefined);
}
