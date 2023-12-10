/* eslint-disable @typescript-eslint/no-floating-promises */
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import * as msal from '@azure/msal-browser';

/**
 * This component is used to redirect the user to the Azure authorization endpoint from a popup.
 */
class CloseConsentPopup extends React.Component {
    componentDidMount() {
         microsoftTeams.app.initialize().then(() => {
            // eslint-disable-next-line @typescript-eslint/require-await, @typescript-eslint/no-floating-promises
             microsoftTeams.app.getContext().then(async (context) => {
               console.log( context.user?.displayName);
                const msalConfig: msal.Configuration = {
                    auth: {
                        clientId:
                            process.env.REACT_APP_AZURE_APP_REGISTRATION_ID ?? '0e730660-a590-40ae-9974-35f1b95ced0f',
                        authority: `https://login.microsoftonline.com/${context.user?.tenant?.id}`,
                        navigateToLoginRequestUrl: false,
                    },
                    cache: {
                        cacheLocation: 'sessionStorage',
                    },
                };

                const msalInstance = new msal.PublicClientApplication(msalConfig);

                msalInstance
                    .handleRedirectPromise()
                    .then((tokenResponse) => {
                        if (tokenResponse !== null) {
                            microsoftTeams.authentication.notifySuccess('Authentication succedded');
                        } else {
                            microsoftTeams.authentication.notifyFailure('Get empty response.');
                        }
                    })
                    .catch((error) => {
                        microsoftTeams.authentication.notifyFailure(JSON.stringify(error));
                    });
            });
        });
    }

    render() {
        return (
            <div>
                <h1>Consent flow complete.</h1>
            </div>
        );
    }
}

export default CloseConsentPopup;
