import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  Configuration,
  AccountInfo,
  AuthenticationResult,
  SilentFlowRequest,
} from "@azure/msal-node";
import { shell } from "electron";

interface TokenRequest {
  scopes: string[];
  account?: AccountInfo;
}

export default class AuthProvider {
  private msalConfig;
  private clientApplication;
  private account: null | AccountInfo;
  private cache;
  readonly initialAccount = {
    homeAccountId: "",
    environment: "",
    tenantId: "",
    username: "",
    localAccountId: "",
  };

  constructor(msalConfig: Configuration) {
    this.msalConfig = msalConfig;
    this.clientApplication = new PublicClientApplication(this.msalConfig);
    this.account = null;
    this.cache = this.clientApplication.getTokenCache();
  }

  /**
   * If there are scopes that you would like users to consent up front, add them
   * in the scopes array.
   */
  async login(): Promise<AccountInfo | null> {
    const authResponse: AuthenticationResult | null = await this.getToken({
      scopes: [],
      account: this.initialAccount,
    });
    if (authResponse) return this.handleResponse(authResponse);
    return null;
  }

  async logout() {
    if (!this.account) return;
    if (!this.account.idTokenClaims) return;
    try {
      /**
       * If you would like to end the session with AAD, use the logout endpoint. You'll need to enable
       * the optional token claim 'login_hint' for this to work as expected.
       */
      if (!this.account.idTokenClaims.login_hint) return;
      if (this.account.idTokenClaims.hasOwnProperty("login_hint")) {
        await shell.openExternal(
          `${this.msalConfig.auth.authority}/oauth2/v2.0/logout?logout_hint=${encodeURIComponent(
            this.account.idTokenClaims.login_hint
          )}`
        );
      }
      await this.cache.removeAccount(this.account);
      this.account = null;
    } catch (error) {
      console.log(error);
    }
  }

  async getToken(tokenRequest: SilentFlowRequest) {
    let authResponse;
    const account = this.account || (await this.getAccount());
    if (account) {
      tokenRequest.account = account;
      authResponse = await this.getTokenSilent(tokenRequest);
    } else {
      authResponse = await this.getTokenInteractive(tokenRequest);
    }
    return authResponse || null;
  }

  async getTokenSilent(tokenRequest: SilentFlowRequest) {
    try {
      return await this.clientApplication.acquireTokenSilent(tokenRequest);
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        console.log("Silent token acquisition failed, acquiring token interactive");
        return await this.getTokenInteractive(tokenRequest);
      }
      console.log(error);
    }
  }

  async getTokenInteractive(tokenRequest: TokenRequest) {
    try {
      const openBrowser = async (url: string) => {
        await shell.openExternal(url);
      };
      const authResponse = await this.clientApplication.acquireTokenInteractive({
        ...tokenRequest,
        openBrowser,
        successTemplate: "You have succesfully signed in. You can close this window.",
        errorTemplate: "Something went wrong, please try signing in again.",
      });
      return authResponse;
    } catch (error) {
      throw error;
    }
  }

  /**
   * Handles the response from a popup or redirect. If response is null, will check if we have any accounts and attempt to sign in.
   * @param response
   */
  async handleResponse(response: AuthenticationResult) {
    console.log("AuthProvider handleResponse(response)", response);
    if (response !== null) {
      this.account = response.account;
    } else {
      this.account = await this.getAccount();
    }
    return this.account;
  }

  /**
   * Calls getAllAccounts and determines the correct account to sign into, currently defaults to first account found in cache.
   * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
   */
  async getAccount() {
    const currentAccounts = await this.cache.getAllAccounts();
    if (!currentAccounts) return null;
    if (currentAccounts.length > 1) {
      // Add choose account code here
      console.log("Multiple accounts detected, need to add choose account code.");
      return currentAccounts[0];
    } else if (currentAccounts.length === 1) {
      return currentAccounts[0];
    } else {
      return null;
    }
  }
}
