import {
  BrowserAuthorizationCallbackHandler,
  BrowserAuthorizationClient,
  BrowserAuthorizationClientConfiguration,
} from "@bentley/frontend-authorization-client";
import { FrontendRequestContext } from "@bentley/imodeljs-frontend";

class AuthorizationClient {
  private static _oidcClient: BrowserAuthorizationClient;

  public static get oidcClient(): BrowserAuthorizationClient {
    return this._oidcClient;
  }

  public static async initializeOidc(): Promise<void> {
    if (this._oidcClient) {
      return;
    }

    /*
IMJS_AUTH_CLIENT_CLIENT_ID=spa-uqhpUbc1pk70uNDmr7lG8nHRR
IMJS_AUTH_CLIENT_REDIRECT_URI=http://localhost:3000/signin-callback
IMJS_AUTH_CLIENT_LOGOUT_URI=http://localhost:3000/logout
IMJS_AUTH_CLIENT_SCOPES = "email openid profile organization assets:modify assets:read connections:modify connections:read context-registry-service:read-only forms:modify forms:read general-purpose-imodeljs-backend imodelhub imodeljs-router imodels:modify imodels:read insights:read issues:modify issues:read library:read product-settings-service projects:modify projects:read projectwise-share rbac-user:external-client realitydata:read storage:modify storage:read urlps-third-party users:read validation:modify validation:read"
# ---- Test ids ----

  

    const scope = process.env.IMJS_AUTH_CLIENT_SCOPES ?? "";
    const clientId = process.env.IMJS_AUTH_CLIENT_CLIENT_ID ?? "";
    const redirectUri = process.env.IMJS_AUTH_CLIENT_REDIRECT_URI ?? "";
    const postSignoutRedirectUri = process.env.IMJS_AUTH_CLIENT_LOGOUT_URI;
*/

    const scope = "email openid profile organization assets:modify assets:read connections:modify connections:read context-registry-service:read-only forms:modify forms:read general-purpose-imodeljs-backend imodelhub imodeljs-router imodels:modify imodels:read insights:read issues:modify issues:read library:read product-settings-service projects:modify projects:read projectwise-share rbac-user:external-client realitydata:read storage:modify storage:read urlps-third-party users:read validation:modify validation:read";
    const clientId = "spa-xbXROps01bjtnyzjy1Z76aHWO";
    const redirectUri = "https://issuesimpexp.herokuapp.com/signin-callback";
    const postSignoutRedirectUri = "https://issuesimpexp.herokuapp.com/logout";

    // authority is optional and will default to Production IMS
    const oidcConfiguration: BrowserAuthorizationClientConfiguration = {
      clientId,
      redirectUri,
      postSignoutRedirectUri,
      scope,
      responseType: "code",
    };

    await BrowserAuthorizationCallbackHandler.handleSigninCallback(
      oidcConfiguration.redirectUri
    );

    this._oidcClient = new BrowserAuthorizationClient(oidcConfiguration);
  }

  public static async signIn(): Promise<void> {
    await this.oidcClient.signIn(new FrontendRequestContext());
  }

  public static async signInSilent(): Promise<void> {
    await this.oidcClient.signInSilent(new FrontendRequestContext());
  }

  public static async signOut(): Promise<void> {
    await this.oidcClient.signOut(new FrontendRequestContext());
  }
}

export default AuthorizationClient;
