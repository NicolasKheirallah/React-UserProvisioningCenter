import { redactSecrets, summarizeRequest, AuditService } from './AuditService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { TelemetryService } from '../telemetry/TelemetryService';
import type { IAuditEntry, IAuditInput } from '../../models';

describe('redactSecrets', () => {
  it('redacts a top-level secret-shaped key', () => {
    expect(redactSecrets({ password: 'hunter2' })).toEqual({ password: '***redacted***' });
  });

  it('redacts every secret-shaped key regardless of case', () => {
    const input = { Password: 'a', TAP: 'b', Secret: 'c', Credential: 'd', Token: 'e', ok: 'f' };
    expect(redactSecrets(input)).toEqual({
      Password: '***redacted***',
      TAP: '***redacted***',
      Secret: '***redacted***',
      Credential: '***redacted***',
      Token: '***redacted***',
      ok: 'f'
    });
  });

  it('redacts the whole nested object when its own container key looks secret-shaped', () => {
    // Graph's passwordProfile body: { passwordProfile: { password, forceChangePasswordNextSignIn } }
    // — the exact shape runCreateUser posts. "passwordProfile" itself matches
    // the secret-key pattern (contains "pass"), so the whole sub-object is
    // redacted rather than recursed into — the conservative, safe direction.
    const body = {
      accountEnabled: true,
      passwordProfile: {
        password: 'S3cr3t!',
        forceChangePasswordNextSignIn: true
      }
    };
    expect(redactSecrets(body)).toEqual({
      accountEnabled: true,
      passwordProfile: '***redacted***'
    });
    expect(JSON.stringify(redactSecrets(body))).not.toContain('S3cr3t!');
  });

  it('recurses into a nested object whose own key is not secret-shaped', () => {
    const body = { profile: { password: 'S3cr3t!', displayName: 'Anna' } };
    expect(redactSecrets(body)).toEqual({
      profile: { password: '***redacted***', displayName: 'Anna' }
    });
  });

  it('redacts secret-shaped keys inside arrays of objects', () => {
    const body = [{ token: 'a' }, { ok: 'b' }];
    expect(redactSecrets(body)).toEqual([{ token: '***redacted***' }, { ok: 'b' }]);
  });

  it('leaves non-secret-shaped keys and primitives untouched', () => {
    expect(redactSecrets({ userPrincipalName: 'anna@contoso.com', count: 3 })).toEqual({
      userPrincipalName: 'anna@contoso.com',
      count: 3
    });
    expect(redactSecrets('plain string')).toBe('plain string');
    expect(redactSecrets(42)).toBe(42);
    expect(redactSecrets(null)).toBe(null);
  });
});

describe('summarizeRequest', () => {
  it('never includes a raw secret value in the summary it hands to the audit log', () => {
    const summary: string = summarizeRequest('POST', '/users', {
      passwordProfile: { password: 'hunter2' }
    });
    expect(summary).not.toContain('hunter2');
    expect(summary).toContain('***redacted***');
  });
});

describe('AuditService', () => {
  function makeService(addAuditEntry: (entry: IAuditEntry) => Promise<void>): AuditService {
    const data = { addAuditEntry } as unknown as SharePointDataService;
    return new AuditService(data, 'operator@contoso.com');
  }

  const input: IAuditInput = {
    jobId: 'job-1',
    action: 'create-user',
    targetUser: 'anna@contoso.com',
    graphEndpoint: '/users',
    requestSummary: '{}',
    responseCode: 201,
    durationMs: 120,
    result: 'Success',
    correlationId: 'corr-1'
  };

  it('stamps actor and timestamp and writes exactly one entry on success', async () => {
    const written: IAuditEntry[] = [];
    const svc = makeService(async (entry) => {
      written.push(entry);
    });
    await svc.log(input);
    expect(written).toHaveLength(1);
    expect(written[0].actor).toBe('operator@contoso.com');
    expect(written[0].timestampUtc).toEqual(expect.any(String));
    expect(written[0].jobId).toBe('job-1');
  });

  it('does not throw when the underlying write fails, and reports it to telemetry when wired', async () => {
    const svc = makeService(async () => {
      throw new Error('list unavailable');
    });
    const trackError = jest.fn();
    svc.setTelemetry({ trackError } as unknown as TelemetryService);

    await expect(svc.log(input)).resolves.toBeUndefined();
    expect(trackError).toHaveBeenCalledTimes(1);
    expect(trackError.mock.calls[0][1]).toMatchObject({ scope: 'audit.write', jobId: 'job-1' });
  });

  it('does not throw when the write fails and no telemetry is wired', async () => {
    const svc = makeService(async () => {
      throw new Error('list unavailable');
    });
    await expect(svc.log(input)).resolves.toBeUndefined();
  });
});
