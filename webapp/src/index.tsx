import { PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import ReactDOM from 'react-dom/client';
import { Provider as ReduxProvider } from 'react-redux';
import App from './App';
import { Constants } from './Constants';
import './index.css';
import { AuthConfig, AuthHelper } from './libs/auth/AuthHelper';
import { store } from './redux/app/store';

import * as microsoftTeams from '@microsoft/teams-js';
import React from 'react';
import { TeamsAuthHelper } from './libs/auth/TeamsAuthHelper';
import { BackendServiceUrl } from './libs/services/BaseService';
import { setAuthConfig } from './redux/features/app/appSlice';

if (!localStorage.getItem('debug')) {
    localStorage.setItem('debug', `${Constants.debug.root}:*`);
}

let container: HTMLElement | null = null;
let root: ReactDOM.Root | undefined = undefined;
let msalInstance: PublicClientApplication | undefined;

document.addEventListener('DOMContentLoaded', () => {
    if (!container) {
        container = document.getElementById('root');
        if (!container) {
            throw new Error('Could not find root element');
        }
        root = ReactDOM.createRoot(container);

        renderApp();
    }
});

export function renderApp() {
    fetch(new URL('authConfig', BackendServiceUrl))
        .then((response) => (response.ok ? (response.json() as Promise<AuthConfig>) : Promise.reject()))
        .then(async (authConfig) => {
            store.dispatch(setAuthConfig(authConfig));

            if (AuthHelper.isAuthAAD()) {
                console.log('AAD Auth Enabled');

                //get window.location urllogin.microsoftonline.com
                const url = new URL(window.location.href);
                console.log('url: ${ url }');
                //get params from url
                const params = new URLSearchParams(url.search);
                console.log('inTeams: ${params}');
                if (params.get('inTeams')) {
                    // const appExpress: any = express();
                    //  setup(appExpress);

                    // Initialize the Microsoft Teams SDK
                    void microsoftTeams.app.initialize();
                    console.log('I am in teams auth');
                    //sso.tsinTeams
                    await TeamsAuthHelper.ssoAuth();
                    // render with the Teams auth if AAD is enabled
                    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                    root!.render(
                        <React.StrictMode>
                            <ReduxProvider store={store}>
                                <App />
                            </ReduxProvider>
                        </React.StrictMode>,
                    );
                } else {
                    if (!msalInstance) {
                        msalInstance = new PublicClientApplication(AuthHelper.getMsalConfig(authConfig));
                        void msalInstance.handleRedirectPromise().then((response) => {
                            if (response) {
                                msalInstance?.setActiveAccount(response.account);
                            }
                        });
                    }

                    // render with the MsalProvider if AAD is enabled
                    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
                    root!.render(
                        <React.StrictMode>
                            <ReduxProvider store={store}>
                                <MsalProvider instance={msalInstance}>
                                    <App />
                                </MsalProvider>
                            </ReduxProvider>
                        </React.StrictMode>,
                    );
                }
            }
        })
        .catch(() => {
            store.dispatch(setAuthConfig(undefined));
        });

    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    root!.render(
        <React.StrictMode>
            <ReduxProvider store={store}>
                <App />
            </ReduxProvider>
        </React.StrictMode>,
    );
}
