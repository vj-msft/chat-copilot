/* eslint-disable @typescript-eslint/restrict-template-expressions */
import * as microsoftTeams from '@microsoft/teams-js';

let consentToken = '';
// 1. Get auth token
// Ask Teams to get us a token from AAD
export const getClientSideToken = (): Promise<string> => {
    return new Promise((resolve, reject) => {
        display('1. Get auth token from Microsoft Teams');

        microsoftTeams.authentication
            .getAuthToken()
            .then((result) => {
                display(result);
                resolve(result);
                consentToken = result;
            })
            .catch((error) => {
                reject('Error getting token: ' + error);
            });
    });
};

export const isTeamsAuthenticated = () => {
    console.log('isTeamsAuthenticated reached');
    let token = false;
    return microsoftTeams.app.initialize()
        .then(() => getClientSideToken())
        .then((result) => {
            token = !!result;
            return token;
        })
        .catch((error) => {
            console.error('Error checking Teams authentication:', error);
            return token;
        });
};

// 2. Exchange that token for a token with the required permissions
//    using the web service (see /auth/token handler in app.js)
export const getServerSideToken = async (clientSideToken: string)=> {
    try {
        const context = await microsoftTeams.app.getContext();
        if (context.user?.tenant) {
            const response = await fetch(
                `/getProfileOnBehalfOf?token=${clientSideToken}&tid=${context.user.tenant.id}`,
                {
                    method: 'get',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    mode: 'cors',
                    cache: 'default',
                },
            );

            if (response.ok) {
                // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
                const profile = await response.json();

                // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-assignment
                consentToken = profile.accessToken;
            } else {
                const errorText = await response.text();
                // eslint-disable-next-line @typescript-eslint/no-throw-literal
                throw errorText;
            }
        } else {
            // eslint-disable-next-line @typescript-eslint/no-throw-literal
            throw 'Error: Unable to get user tenant information from context';
        }
    } catch (error) {
        throw error;
    }
};

function showConsentDialog() {
    microsoftTeams.authentication
        .authenticate({
            url: `${window.location.origin}/auth-start`,
            width: 600,
            height: 535,
        })
        .then((result) => {
            consentToken = result;
        })
        .catch((error) => {
            console.error('Consent failed: ', JSON.stringify(error));
        });
}

// Add text to the display in a <p> or other HTML element
const display = (text: string, elementTag?: string) => {
    const logDiv = document.getElementById('logs');
    const p = document.createElement(elementTag ? elementTag : 'p');
    p.innerText = text;
    logDiv?.append(p);
    console.log('ssoDemo: ' + text);
    return p;
};

const ssoAuth = async () => {
    try {
        await microsoftTeams.app.initialize();
        await getClientSideToken();
    } catch (error) {
        if ((JSON.parse(error as string) as { error?: string }).error === 'consent_required') {
            display(`Error: ${error} - user or admin consent required`);
            try {
                showConsentDialog();
                return consentToken;
            } catch (error) {
                display(`Error: ${error}`);
            }
        } else {
            display(`Error from web service: ${error}`);
        }
    }
    return consentToken;
};

export const TeamsAuthHelper = {
    ssoAuth,
    getClientSideToken,
    display,
    isTeamsAuthenticated,
};
