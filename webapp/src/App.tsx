// Copyright (c) Microsoft. All rights reserved.

import { AuthenticatedTemplate, UnauthenticatedTemplate, useIsAuthenticated, useMsal } from '@azure/msal-react';
import { FluentProvider, Subtitle1, makeStyles, shorthands, tokens } from '@fluentui/react-components';

import * as microsoftTeams from '@microsoft/teams-js';
import * as React from 'react';
import { useEffect } from 'react';
import { UserSettingsMenu } from './components/header/UserSettingsMenu';
import { PluginGallery } from './components/open-api-plugins/PluginGallery';
import { BackendProbe, ChatView, Error, Loading, Login } from './components/views';
import { AuthHelper } from './libs/auth/AuthHelper';
import { TeamsAuthHelper } from './libs/auth/TeamsAuthHelper';
import { useChat, useFile } from './libs/hooks';
import { AlertType } from './libs/models/AlertType';
import { useAppDispatch, useAppSelector } from './redux/app/hooks';
import { RootState } from './redux/app/store';
import { FeatureKeys } from './redux/features/app/AppState';
import { addAlert, setActiveUserInfo, setServiceInfo } from './redux/features/app/appSlice';
import { semanticKernelDarkTheme, semanticKernelLightTheme } from './styles';
//import { setup } from './libs/auth/TeamsAuthServer';
//import express from 'express';

export const useClasses = makeStyles({
    container: {
        display: 'flex',
        flexDirection: 'column',
        height: '100vh',
        width: '100%',
        ...shorthands.overflow('hidden'),
    },
    header: {
        alignItems: 'center',
        backgroundColor: tokens.colorBrandForeground2,
        color: tokens.colorNeutralForegroundOnBrand,
        display: 'flex',
        '& h1': {
            paddingLeft: tokens.spacingHorizontalXL,
            display: 'flex',
        },
        height: '48px',
        justifyContent: 'space-between',
        width: '100%',
    },
    persona: {
        marginRight: tokens.spacingHorizontalXXL,
    },
    cornerItems: {
        display: 'flex',
        ...shorthands.gap(tokens.spacingHorizontalS),
    },
});

enum AppState {
    ProbeForBackend,
    SettingUserInfo,
    ErrorLoadingChats,
    ErrorLoadingUserInfo,
    LoadingChats,
    Chat,
    SigningOut,
}

const App = () => {
    const classes = useClasses();

    const [appState, setAppState] = React.useState(AppState.ProbeForBackend);
    const dispatch = useAppDispatch();

    const { instance, inProgress } = useMsal();
    const { features, isMaintenance } = useAppSelector((state: RootState) => state.app);
    const isAuthenticated = useIsAuthenticated();

    const isTeamsAuthenticated = async () => {
        try {
            const result = await TeamsAuthHelper.isTeamsAuthenticated();
            console.log('Am i in the right place');
            console.log(result);
            return result;
        } catch (error) {
            console.error('Error checking Teams authentication:', error);
            return false;
        }
    };

   // const isTeamsAuthenticatedValue = isTeamsAuthenticated();

    async function teamsUserContext() {
        try {
            console.log('teamsUserContext');
            const context = await microsoftTeams.app.getContext();
            //check if context.user is null

            if (context.user) {
                return {
                    tid: context.user.tenant?.id,
                    email: context.user.userPrincipalName,
                    name: context.user.displayName?context.user.displayName:context.user.userPrincipalName,
                };
            } else {
                return null; // Handle the case when context.user is null
            }
        } catch (error) {
            console.error('Error getting Teams user context:', error);
            return null;
        }
    }

    const chat = useChat();
    const file = useFile();

    useEffect(() => {
        console.log('useEffect');

        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const fetchData = async () => {

            console.log('fetchData called');
            if (isMaintenance && appState !== AppState.ProbeForBackend) {
                setAppState(AppState.ProbeForBackend);
                return;
            }
            const teamsAuthenticated = await isTeamsAuthenticated();
            console.log(`the teamsAuthenticated is ${teamsAuthenticated}`);
            if (teamsAuthenticated) {
                // Add the route for handling tabs
              //  const expressApp = express();
               // setup(expressApp);
            }

            if ((isAuthenticated || teamsAuthenticated) && appState === AppState.SettingUserInfo) {
                console.log('SettingUserInfo');
                const account = instance.getActiveAccount();
                if (!account && isAuthenticated) {
                    setAppState(AppState.ErrorLoadingUserInfo);
                } else {
                    let activeUserInfoPayload;
                    if (teamsAuthenticated) {
                        console.log('SettingUserInfo with teams info');
                        teamsUserContext()
                            .then((result) => {
                                console.log(result);
                                if (result?.tid && result.email && result.name) {
                                    activeUserInfoPayload = {
                                        id: result.tid,
                                        email: result.email,
                                        username: result.name,
                                    };
                                } else {
                                    // Handle the case when teamsUserContext returns incomplete data
                                    setAppState(AppState.ErrorLoadingUserInfo);
                                    return;
                                }
                                dispatch(setActiveUserInfo(activeUserInfoPayload));
                            })
                            .catch((error) => {
                                console.error('Error getting Teams user context:', error);
                                setAppState(AppState.ErrorLoadingUserInfo);
                            });
                    } else {
                        console.log('SettingUserInfo with browser info');
                        if (account) {
                            activeUserInfoPayload = {
                                id: `${account.localAccountId}.${account.tenantId}`,
                                email: account.username, // username is the email address
                                username: account.name ?? account.username,
                            };

                            dispatch(setActiveUserInfo(activeUserInfoPayload));
                        }
                    }
                    // Privacy disclaimer for internal Microsoft users
                    if (account?.username.split('@')[1] === 'microsoft.com') {
                        dispatch(
                            addAlert({
                                message:
                                    'By using Chat Copilot, you agree to protect sensitive data, not store it in chat, and allow chat history collection for service improvements. This tool is for internal use only.',
                                type: AlertType.Info,
                            }),
                        );
                    }

                    setAppState(AppState.LoadingChats);
                }
            }

            if (
                (isAuthenticated || !AuthHelper.isAuthAAD() || teamsAuthenticated) &&
                appState === AppState.LoadingChats
            ) {
                void Promise.all([
                    // Load all chats from memory
                    chat
                        .loadChats()
                        .then(() => {
                            setAppState(AppState.Chat);
                        })
                        .catch(() => {
                            setAppState(AppState.ErrorLoadingChats);
                        }),

                    // Check if content safety is enabled
                    file.getContentSafetyStatus(),

                    // Load service information
                    chat.getServiceInfo().then((serviceInfo) => {
                        if (serviceInfo) {
                            dispatch(setServiceInfo(serviceInfo));
                        }
                    }),
                ]);
            }
        };
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        fetchData();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [instance, inProgress, isAuthenticated, appState, isMaintenance]);

    const content = <Chat classes={classes} appState={appState} setAppState={setAppState} />;
    return (
        <FluentProvider
            className="app-container"
            theme={features[FeatureKeys.DarkMode].enabled ? semanticKernelDarkTheme : semanticKernelLightTheme}
        >
            {AuthHelper.isAuthAAD() && isAuthenticated? (
                <>
                    <UnauthenticatedTemplate>
                        <div className={classes.container}>
                            <div className={classes.header}>
                                <Subtitle1 as="h1">Chat Copilot</Subtitle1>
                            </div>
                            {appState === AppState.SigningOut && <Loading text="Signing you out..." />}
                            {appState !== AppState.SigningOut && <Login />}
                        </div>
                    </UnauthenticatedTemplate>
                    <AuthenticatedTemplate>{content}</AuthenticatedTemplate>
                </>
            ) : (
                content
            )}
        </FluentProvider>
    );
};

const Chat = ({
    classes,
    appState,
    setAppState,
}: {
    classes: ReturnType<typeof useClasses>;
    appState: AppState;
    setAppState: (state: AppState) => void;
}) => {
    const onBackendFound = React.useCallback(() => {
        setAppState(
            AuthHelper.isAuthAAD()
                ? // if AAD is enabled, we need to set the active account before loading chats
                  AppState.SettingUserInfo
                : // otherwise, we can load chats immediately
                  AppState.LoadingChats,
        );
    }, [setAppState]);
    return (
        <div className={classes.container}>
            <div className={classes.header}>
                <Subtitle1 as="h1">Chat Copilot</Subtitle1>
                {appState > AppState.SettingUserInfo && (
                    <div className={classes.cornerItems}>
                        <div className={classes.cornerItems}>
                            <PluginGallery />
                            <UserSettingsMenu
                                setLoadingState={() => {
                                    setAppState(AppState.SigningOut);
                                }}
                            />
                        </div>
                    </div>
                )}
            </div>
            {appState === AppState.ProbeForBackend && <BackendProbe onBackendFound={onBackendFound} />}
            {appState === AppState.SettingUserInfo && (
                <Loading text={'Hang tight while we fetch your information...'} />
            )}
            {appState === AppState.ErrorLoadingUserInfo && (
                <Error text={'Unable to load user info. Please try signing out and signing back in.'} />
            )}
            {appState === AppState.ErrorLoadingChats && (
                <Error text={'Unable to load chats. Please try refreshing the page.'} />
            )}
            {appState === AppState.LoadingChats && <Loading text="Loading chats..." />}
            {appState === AppState.Chat && <ChatView />}
        </div>
    );
};

export default App;
