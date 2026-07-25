/**
 * Thrown when the signed-in operator lacks the app role/permission required
 * for a workflow action.
 *
 * This is a business-logic authorization check that sits on top of Entra's
 * delegated permission enforcement. UI gating should prevent most callers from
 * ever hitting this, but the engine enforces it defensively.
 */
export class AuthorizationError extends Error {
  public readonly permission: string;

  public constructor(permission: string) {
    super(`Operation requires permission: ${permission}`);
    Object.setPrototypeOf(this, AuthorizationError.prototype);
    this.name = 'AuthorizationError';
    this.permission = permission;
  }
}
