// Import necessary types from the Teams SDK and MSAL
import * as msal from '@azure/msal-browser';
import * as microsoftTeams from '@microsoft/teams-js';
import { store } from '../../redux/app/store';
import React from 'react';

export const TeamsAuthEnd: React.FC = () => {
    React.useEffect(() => {
        const fetchData = async () => {
            try {
                await microsoftTeams.app.initialize();

                // Get the tab context, and use the information to handle the redirect promise
                const context = await microsoftTeams.app.getContext();
                const getAuthConfig = () => store.getState().app.authConfig;
                console.log('I am in authend');


                const msalConfig: msal.Configuration = {
                    auth: {
                        clientId: getAuthConfig()?.aadClientId ?? '',
                        authority: `https://login.microsoftonline.com/${context.user?.tenant?.id}`,
                        navigateToLoginRequestUrl: false,
                    },
                    cache: {
                        cacheLocation: 'sessionStorage',
                    },
                };

                const msalInstance = new msal.PublicClientApplication(msalConfig);

                // Handle the redirect promise from MSAL
                const tokenResponse = await msalInstance.handleRedirectPromise();

                if (tokenResponse !== null) {
                    // Notify Teams about the successful authentication
                    microsoftTeams.authentication.notifySuccess(
                        JSON.stringify({
                            sessionStorage: sessionStorage,
                        }),
                    );
                } else {
                    // Notify Teams about the empty response
                    microsoftTeams.authentication.notifyFailure('Get empty response.');
                }
            } catch (error) {
                // Notify Teams about the failure with the error details
                // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
                console.error(`Error: ${error}`);
                microsoftTeams.authentication.notifyFailure(JSON.stringify(error));
            }
        }
            void fetchData();
    }, []);

    return <div>Teams Auth End</div>;
};
