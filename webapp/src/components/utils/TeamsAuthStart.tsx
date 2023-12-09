// Import necessary types from the Teams SDK and MSAL
import * as msal from '@azure/msal-browser';
import * as microsoftTeams from '@microsoft/teams-js';
import { store } from '../../redux/app/store';
import React from 'react';

export const TeamsAuthStart: React.FC = () => {
    React.useEffect(() => {
        const fetchData = async () => {
             try {
                 await microsoftTeams.app.initialize();

                 // Get the tab context, and use the information to navigate to Azure AD login page
                 const context = await microsoftTeams.app.getContext();
                 const getAuthConfig = () => store.getState().app.authConfig;
                 console.log('I am in authstart');
                 console.log(getAuthConfig()?.aadClientId ?? '');

                 //const currentURL = new URL(window.location.href);
                 const scope = 'access_as_user User.Read email openid profile offline_access';
                 // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
                 const loginHint = context.user?.userPrincipalName; // Assuming loginHint comes from userPrincipalName

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
                 const scopesArray = scope.split(' ');
                 const scopesRequest: msal.RedirectRequest = {
                     scopes: scopesArray,
                     redirectUri: window.location.origin + `/teamsAuthEnd`,
                     loginHint: loginHint,
                 };

                 // Use loginRedirect to navigate to the Azure AD login page
                 await msalInstance.loginRedirect(scopesRequest);
             } catch (error) {
                 // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
                 console.error(`Error: ${error}`);
             }
        };

        void fetchData();
    }, []);

    return <div>Teams Auth Start</div>;
};
/*
export const TeamsAuthStart = async () => {

};*/

/*export const TeamsAuthStart = () => {
    startAuth;
};*/
