/* eslint-disable @typescript-eslint/require-await */
/* eslint-disable @typescript-eslint/no-throw-literal */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-misused-promises */
/* eslint-disable @typescript-eslint/no-unsafe-return */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-var-requires */
const express = require('express');
const cors = require('cors');
const jwt_decode = require('jwt-decode');
const msal = require('@azure/msal-node');
const app = express();
const path = require('path');
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

const clientId = process.env.CLIENT_ID ?? '0e730660-a590-40ae-9974-35f1b95ced0f';
const clientSecret = process.env.CLIENT_SECRET ?? 'Z~F8Q~sigRT8D_frXC71vJrqh2CzHglPDFIM6b-n';
const graphScopes = ['https://graph.microsoft.com/User.Read'];
let handleQueryError = function (err) {
    console.log('handleQueryError called: ', err);
    return new Response(
        JSON.stringify({
            code: 400,
            message: 'network Error',
        }),
    );
};
/*
const corsOptions = {
    origin: 'https://8449-86-20-192-252.ngrok-free.app/', // Replace with the actual origin of your React app
    credentials: true, // Enable credentials (cookies, authorization headers)
};
// Enable CORS
app.use(cors(corsOptions));
*/
app.get('/getProfileOnBehalfOf', async (req, res) => {
    const msalClient = new msal.ConfidentialClientApplication({
        auth: {
            clientId: clientId,
            clientSecret: clientSecret,
        },
    });

    msalClient
        .acquireTokenOnBehalfOf({
            authority: `https://login.microsoftonline.com/${req.query.tid}`,
            oboAssertion: req.query.token,
            scopes: graphScopes,
            skipCache: true,
        })
        .then(async (result) => {
            console.log(result.accessToken);
            res.status(200).json({ accessToken: result.accessToken });
        })
        .catch((error) => {
            console.log('error' + error.errorCode);
            res.status(403).json({ error: 'consent_required' });
        });
});

// Handles any requests that don't match the ones above
app.get('*', (req, res) => {
    console.log('Unhandled request: ', req);
    res.status(404).send('Path not defined');
});
// Pop-up dialog to ask for additional permissions, redirects to AAD page
/*app.get('/teamsAuthStart', function (_req, res) {
    res.render('TeamsAuthStart', { clientId: clientId });
});

// End of the pop-up dialog auth flow, returns the results back to the parent window
app.get('/teamsAuthEnd', function (_req, res) {
    res.render('TeamsAuthEnd', { clientId: clientId });
});
*/
const port = process.env.PORT ?? 3001;
app.listen(port);

console.log('API server is listening on port ' + port);
