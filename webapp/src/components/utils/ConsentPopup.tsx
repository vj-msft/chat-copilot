/* eslint-disable @typescript-eslint/no-floating-promises */
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import * as msal from '@azure/msal-browser';

/**
 * This component is loaded to grant consent for graph permissions.
 */
class ConsentPopup extends React.Component {
    componentDidMount() {
         microsoftTeams.app.initialize().then(() => {
            // Get the tab context, and use the information to navigate to Azure AD login page
             // eslint-disable-next-line @typescript-eslint/no-floating-promises
             microsoftTeams.app.getContext().then(async (context) => {
                const scope = 'User.Read';
                const loginHint = context.user?.loginHint;
console.log(`this is tenantid ${context.user?.tenant?.id}`);
                const msalConfig: msal.Configuration = {
                    auth: {
                        clientId:'0e730660-a590-40ae-9974-35f1b95ced0f',
                         //   process.env.REACT_APP_AZURE_APP_REGISTRATION_ID ?? '0e730660-a590-40ae-9974-35f1b95ced0f',
                        // authority: `https://login.microsoftonline.com/${context.user?.tenant?.id}`,
                        authority: `https://login.microsoftonline.com/2ec5b44c-6459-40b1-9429-42e529afc36d`,
                        navigateToLoginRequestUrl: false,
                    },
                    cache: {
                        cacheLocation: 'sessionStorage',
                    },
                };

                const msalInstance = new msal.PublicClientApplication(msalConfig);

                const scopesArray = scope.split(' ');
                const scopesRequest = {
                    scopes: scopesArray,
                    redirectUri: window.location.origin + `/auth-end`,
                    loginHint: loginHint,
                };

                    await msalInstance.loginRedirect(scopesRequest);
            });
        });
    }

    render() {
        return (
            <div>
                <h1>Redirecting to consent page.</h1>
            </div>
        );
    }
}

export default ConsentPopup;
