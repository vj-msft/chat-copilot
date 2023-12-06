/* eslint-disable @typescript-eslint/restrict-template-expressions */
import * as microsoftTeams from '@microsoft/teams-js';

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
            })
            .catch((error) => {
                reject('Error getting token: ' + error);
            });
    });
};

export const isTeamsAuthenticated = async (): Promise<boolean> => {
    let token = false;
    try {
        await getClientSideToken().then((result) => (token = !!result));
    } catch (error) {
        console.error('Error checking Teams authentication:', error);
    }
    return token;
};

// 2. Exchange that token for a token with the required permissions
//    using the web service (see /auth/token handler in app.js)
export const getServerSideToken = async (clientSideToken: string): Promise<any> => {
    try {
        const context = await microsoftTeams.app.getContext();
        if (context.user?.tenant) {
            const response = await fetch('/getProfileOnBehalfOf', {
                method: 'post',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    tid: context.user.tenant.id,
                    token: clientSideToken,
                }),
                mode: 'cors',
                cache: 'default',
            });

            if (response.ok) {
                // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
                const profile = await response.json();
                // eslint-disable-next-line @typescript-eslint/no-unsafe-return
                return profile;
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
// 3. Get the server side token and use it to call the Graph API
/* const useServerSideTokenWithGraph = (data: any) => {
        display('2. Call https://graph.microsoft.com/v1.0/me/ with the server side token');
        return display(JSON.stringify(data, undefined, 4), 'pre');
    };*/
// Show the consent pop-up
const requestConsent = (): Promise<string> => {
    return new Promise((resolve, reject) => {
        microsoftTeams.authentication
            .authenticate({
                url: window.location.origin + '/teamsAuthStart',
                width: 600,
                height: 535,
            })
            .then((result) => {
                const tokenData = result;
                resolve(tokenData);
            })
            .catch((reason) => {
                reject(JSON.stringify(reason));
            });
    });
};
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
        console.log('ssoAuth reached');
        await microsoftTeams.app.initialize();
        await getClientSideToken();
        console.log('I am done with getClientSideToken');

        // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
        // await getServerSideToken(clientSideToken);
        // eslint-disable-next-line react-hooks/rules-of-hooks
        // useServerSideTokenWithGraph(profile);
    } catch (error) {
        if (error === 'invalid_grant') {
            display(`Error: ${error} - user or admin consent required`);
            const button = display('Consent', 'button');

            button.onclick = async () => {
                try {
                    await requestConsent();
                    display(`Consent succeeded`);
                    if (error === 'invalid_grant') {
                        display(`Error: ${error} - user or admin consent required`);
                        const button = display('Consent', 'button');

                        button.onclick = async () => {
                            try {
                                await requestConsent();
                                const refreshButton = display('Refresh page', 'button') as HTMLButtonElement;
                                refreshButton.onclick = () => {
                                    window.location.reload();
                                };
                            } catch (error) {
                                display(`Error: ${error}`);
                            }
                        };
                    } else {
                        display(`Error from web service: ${error}`);
                    }
                } catch (error) {
                    display(`Error: ${error}`);
                }
            };
        } else {
            display(`Error from web service: ${error}`);
        }
    }
};

export const TeamsAuthHelper = {
    ssoAuth,
    getClientSideToken,
    getServerSideToken,
    requestConsent,
    display,
    isTeamsAuthenticated,
};
