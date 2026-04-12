import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
  type DeviceCodeRequest,
  type InteractiveRequest
} from '@azure/msal-node';
import { shell } from 'electron';
import type { AuthLoginResult } from '../pim/types';

const AUTHORITY_HOST = 'https://login.microsoftonline.com';

export class DeviceCodeAuth {
  private app: PublicClientApplication | null = null;
  private account: AccountInfo | null = null;
  private readonly tenantAccounts = new Map<string, AccountInfo>();
  private clientId = '';

  async login(clientId: string, tenantId?: string): Promise<AuthLoginResult> {
    const tenant = tenantId && tenantId.trim() ? tenantId.trim() : 'organizations';
    this.configureApp(clientId.trim());

    const result = await this.acquireInteractiveWithFallback({
      authority: `${AUTHORITY_HOST}/${tenant}`,
      scopes: ['openid', 'profile', 'offline_access'],
      prompt: 'select_account',
      openBrowser: async (url) => {
        await shell.openExternal(url);
      },
      successTemplate: '<h2>Sign in completed. You can close this window.</h2>',
      errorTemplate: '<h2>Sign in failed. Return to the app and retry.</h2>'
    });

    if (!result || !result.account) {
      throw new Error('Interactive login failed: no account returned');
    }

    this.account = result.account;
    if (result.account.tenantId) {
      this.tenantAccounts.set(result.account.tenantId, result.account);
    }

    return {
      accountId: result.account.homeAccountId,
      tenantId: result.account.tenantId,
      username: result.account.username
    };
  }

  async getGraphToken(): Promise<string> {
    return this.getToken([
      'RoleEligibilitySchedule.Read.Directory',
      'RoleAssignmentSchedule.ReadWrite.Directory',
      'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup'
    ]);
  }

  async getGraphTokenForTenant(tenantId: string): Promise<string> {
    return this.getToken(
      [
        'RoleEligibilitySchedule.Read.Directory',
        'RoleAssignmentSchedule.ReadWrite.Directory',
        'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup'
      ],
      tenantId
    );
  }

  async getArmToken(): Promise<string> {
    return this.getToken(['https://management.azure.com/user_impersonation']);
  }

  async getArmTokenForTenant(tenantId: string): Promise<string> {
    return this.getToken(['https://management.azure.com/user_impersonation'], tenantId);
  }

  private async getToken(scopes: string[], tenantId?: string): Promise<string> {
    if (!this.account) {
      throw new Error('Not signed in. Call login first.');
    }

    const resolvedTenant = tenantId && tenantId.trim() ? tenantId.trim() : this.account.tenantId || 'organizations';
    const accountForTenant = (await this.getAccountForTenant(resolvedTenant)) ?? this.account;

    let result: AuthenticationResult | null = null;

    try {
      result = await this.getApp().acquireTokenSilent({
        account: accountForTenant,
        authority: `${AUTHORITY_HOST}/${resolvedTenant}`,
        scopes
      });
    } catch {
      result = await this.acquireInteractiveWithFallback({
        authority: `${AUTHORITY_HOST}/${resolvedTenant}`,
        scopes,
        prompt: 'select_account',
        loginHint: accountForTenant.username,
        openBrowser: async (url) => {
          await shell.openExternal(url);
        },
        successTemplate: '<h2>Token refresh completed. You can close this window.</h2>',
        errorTemplate: '<h2>Token refresh failed. Return to the app and retry.</h2>'
      });
    }

    if (!result?.accessToken) {
      throw new Error('Failed to acquire access token');
    }

    if (result.account?.tenantId) {
      this.tenantAccounts.set(result.account.tenantId, result.account);
      this.account = result.account;
    }

    return result.accessToken;
  }

  private async getAccountForTenant(tenantId: string): Promise<AccountInfo | null> {
    const fromCache = this.tenantAccounts.get(tenantId);
    if (fromCache) {
      return fromCache;
    }

    const baseAccount = this.account;
    if (!baseAccount) {
      return null;
    }

    const allAccounts = await this.getApp().getTokenCache().getAllAccounts();
    const matchingTenant = allAccounts.find(
      (candidate: AccountInfo) =>
        candidate.tenantId === tenantId &&
        (candidate.homeAccountId === baseAccount.homeAccountId || candidate.username === baseAccount.username)
    );

    if (matchingTenant) {
      this.tenantAccounts.set(tenantId, matchingTenant);
      return matchingTenant;
    }

    return null;
  }

  private configureApp(clientId: string): void {
    if (!clientId) {
      throw new Error('Application (client) ID is required. Configure settings first.');
    }

    if (this.app && this.clientId === clientId) {
      return;
    }

    const config: Configuration = {
      auth: {
        clientId,
        authority: `${AUTHORITY_HOST}/organizations`
      }
    };

    this.app = new PublicClientApplication(config);
    this.clientId = clientId;
    this.account = null;
    this.tenantAccounts.clear();
  }

  private getApp(): PublicClientApplication {
    if (!this.app) {
      throw new Error('Auth client not configured. Save settings first.');
    }

    return this.app;
  }

  private async acquireTokenByDeviceCodeWithGuidance(request: DeviceCodeRequest): Promise<AuthenticationResult | null> {
    try {
      return await this.getApp().acquireTokenByDeviceCode(request);
    } catch (error) {
      throw new Error(this.toActionableAuthMessage(error));
    }
  }

  private async acquireInteractiveWithFallback(request: InteractiveRequest): Promise<AuthenticationResult | null> {
    try {
      return await this.getApp().acquireTokenInteractive(request);
    } catch (interactiveError) {
      if (!request.scopes || request.scopes.length === 0) {
        throw new Error(this.toActionableAuthMessage(interactiveError));
      }

      const fallbackRequest: DeviceCodeRequest = {
        authority: request.authority,
        scopes: request.scopes,
        deviceCodeCallback: (response) => {
          console.log(response.message);
        }
      };

      try {
        return await this.acquireTokenByDeviceCodeWithGuidance(fallbackRequest);
      } catch {
        throw new Error(this.toActionableAuthMessage(interactiveError));
      }
    }
  }

  private toActionableAuthMessage(error: unknown): string {
    const raw = String(error ?? 'Unknown authentication error');
    const normalized = raw.toLowerCase();

    if (normalized.includes('invalid_client')) {
      return [
        'Azure returned invalid_client.',
        'Check app registration: enable Public client flows, ensure Client ID is correct, and ensure Supported account types include multitenant if needed.'
      ].join(' ');
    }

    if (normalized.includes('unauthorized_client')) {
      return 'Azure returned unauthorized_client. Verify tenant consent and delegated API permissions for this app.';
    }

    if (normalized.includes('invalid_scope')) {
      return 'Azure returned invalid_scope. Verify Graph/ARM delegated permissions and admin consent in the target tenant.';
    }

    if (normalized.includes('redirect_uri')) {
      return 'Azure returned redirect_uri error. Add http://localhost as a Mobile and desktop redirect URI in app registration.';
    }

    return raw;
  }
}
