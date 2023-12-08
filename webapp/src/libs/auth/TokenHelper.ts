import {
    IPublicClientApplication,
    InteractionRequiredAuthError,
    InteractionStatus,
    PopupRequest,
} from '@azure/msal-browser';
import * as microsoftTeams from '@microsoft/teams-js';
import { AuthHelper } from './AuthHelper';
import { getClientSideToken, getServerSideToken } from './TeamsAuthHelper';
enum TokenErrors {
    InteractionInProgress = 'interaction_in_progress',
}

/*
 * This implementation follows incremental consent, and token acquisition is limited to one
 * resource at a time (scopes), but user can consent to many resources upfront (extraScopesToConsent)
 */
export const getAccessTokenUsingMsal = async (
    inProgress: InteractionStatus,
    msalInstance: IPublicClientApplication,
    scopes: string[],
    extraScopesToConsent?: string[],
) => {
    const url = new URL(window.location.href);
    //get params from url
    const params = new URLSearchParams(url.search);
    if (params.get('inTeams')) {
        await microsoftTeams.app.initialize();
        const teamsAuthToken: string = await getClientSideToken();


        if (teamsAuthToken && teamsAuthToken.length > 0) {
            console.log(`I am in teams token ${teamsAuthToken}`);
            // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
            const teamsAccessToken: string = await getServerSideToken(teamsAuthToken);
            return teamsAccessToken;
        }
    } else {
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        const account = msalInstance.getActiveAccount()!;
        const authority = AuthHelper.getAuthConfig()?.aadAuthority;
        const accessTokenRequest: PopupRequest = {
            authority,
            scopes,
            extraScopesToConsent,
            account,
        };

        return await acquireToken(accessTokenRequest, msalInstance, inProgress).catch(async (e) => {
            if (e instanceof Error && e.message === (TokenErrors.InteractionInProgress as string)) {
                return await interactionInProgressHandler(inProgress, msalInstance, accessTokenRequest);
            }

            throw e;
        });
    }
    return '';
};

const acquireToken = async (
    accessTokenRequest: PopupRequest,
    msalInstance: IPublicClientApplication,
    interactionStatus: InteractionStatus,
) => {
    return await msalInstance
        .acquireTokenSilent(accessTokenRequest)
        .then(function (accessTokenResponse) {
            // Acquire token silent success
            return accessTokenResponse.accessToken;
        })
        .catch(async (error) => {
            if (error instanceof InteractionRequiredAuthError) {
                // Since app can trigger concurrent interactive requests, first check
                // if any other interaction is in progress proper to invoking a new one
                if (interactionStatus !== InteractionStatus.None) {
                    // throw a new error to be handled in the caller above
                    throw new Error(TokenErrors.InteractionInProgress);
                } else {
                    return await msalInstance
                        .acquireTokenPopup({ ...accessTokenRequest })
                        .then(function (accessTokenResponse) {
                            // Acquire token interactive success
                            return accessTokenResponse.accessToken;
                        })
                        .catch(function (error) {
                            // Acquire token interactive failure
                            throw new Error(`Received error while retrieving access token: ${error as string}`);
                        });
                }
            }
            throw new Error(`Received error while retrieving access token: ${error as string}`);
        });
};

const interactionInProgressHandler = async (
    interactionStatus: InteractionStatus,
    msalInstance: IPublicClientApplication,
    accessTokenRequest: PopupRequest,
) => {
    // Polls the interaction status from the application
    // state and resolves when it's equal to "None".
    waitFor(() => interactionStatus === InteractionStatus.None);

    // Wait is over, call acquireToken again to re-try acquireTokenSilent
    return await acquireToken(accessTokenRequest, msalInstance, interactionStatus);
};

const waitFor = (hasInteractionCompleted: () => boolean) => {
    const checkInteraction = () => {
        if (!hasInteractionCompleted()) {
            setTimeout(checkInteraction, 500);
        }
    };

    checkInteraction();
};

export const TokenHelper = {
    getAccessTokenUsingMsal,
};
