import { ClientRequestContext } from "@bentley/bentleyjs-core";
import { BrowserAuthorizationCallbackHandler, BrowserAuthorizationClient, BrowserAuthorizationClientConfiguration } from "@bentley/frontend-authorization-client";

export class IMSLoginProper{
    public static async getAccessTokenForBentleyAPI(){

    const oidcConfiguration: BrowserAuthorizationClientConfiguration = {
        authority: "https://ims.bentley.com",
        clientId: "spa-xbXROps01bjtnyzjy1Z76aHWO",
        redirectUri: `${window.location.origin}/signin-callback`,
        //scope: "email openid profile organization assets:modify assets:read connections:modify connections:read context-registry-service:read-only forms:modify forms:read general-purpose-imodeljs-backend imodelhub imodeljs-router imodels:modify imodels:read insights:read issues:modify issues:read library:read product-settings-service projects:modify projects:read projectwise-share rbac-user:external-client realitydata:read storage:modify storage:read urlps-third-party users:read validation:modify validation:read",
        scope: "itwinjs email openid profile organization issues:modify issues:read projects:read urlps-third-party users:read",
        responseType: "code",
    };
    await BrowserAuthorizationCallbackHandler.handleSigninCallback(oidcConfiguration.redirectUri);
    const browserClient = new BrowserAuthorizationClient(oidcConfiguration);
    await browserClient.signInSilent(new ClientRequestContext());
    const currentToken =  await browserClient.getAccessToken();
    return currentToken.toTokenString();
    // end of system login process
    }
}