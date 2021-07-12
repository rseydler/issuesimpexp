import { ClientRequestContext } from "@bentley/bentleyjs-core";
import { BrowserAuthorizationCallbackHandler, BrowserAuthorizationClient, BrowserAuthorizationClientConfiguration } from "@bentley/frontend-authorization-client";

export class IMSLoginProper{
    public static async getAccessTokenForBentleyAPI(){

        /*
  const scope = "email openid profile organization assets:modify assets:read connections:modify connections:read context-registry-service:read-only forms:modify forms:read general-purpose-imodeljs-backend imodelhub imodeljs-router imodels:modify imodels:read insights:read issues:modify issues:read library:read product-settings-service projects:modify projects:read projectwise-share rbac-user:external-client realitydata:read storage:modify storage:read urlps-third-party users:read validation:modify validation:read";
    const clientId = "spa-uqhpUbc1pk70uNDmr7lG8nHRR";
    const redirectUri = "http://localhost:3000/signin-callback";
    const postSignoutRedirectUri = "http://localhost:3000/logout";
        */

    const oidcConfiguration: BrowserAuthorizationClientConfiguration = {
        authority: "https://ims.bentley.com",
        clientId: "spa-xbXROps01bjtnyzjy1Z76aHWO",
        redirectUri: "https://issuesimpexp.herokuapp.com/signin-callback",
        scope: "email openid profile organization assets:modify assets:read connections:modify connections:read context-registry-service:read-only forms:modify forms:read general-purpose-imodeljs-backend imodelhub imodeljs-router imodels:modify imodels:read insights:read issues:modify issues:read library:read product-settings-service projects:modify projects:read projectwise-share rbac-user:external-client realitydata:read storage:modify storage:read urlps-third-party users:read validation:modify validation:read",
        responseType: "code",
    };
    await BrowserAuthorizationCallbackHandler.handleSigninCallback(oidcConfiguration.redirectUri);
    const browserClient = new BrowserAuthorizationClient(oidcConfiguration);
    await browserClient.signInSilent(new ClientRequestContext);
    const currentToken =  await browserClient.getAccessToken();
    return currentToken.toTokenString();
    // end of system login process
    }
}