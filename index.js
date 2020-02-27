import { UserAgentApplication } from 'msal';

// Class used to log user with msal library
// This class provide a way to fetch <access_token> used to call API
export default class MsalLogin {
    _msalConfig;
    _tokenRequest;

    constructor({ clientId, tenantId, cache = { cacheLocation: "localStorage", storeAuthStateInCookie: true }, scopes = [] }) {
        this._msalConfig = {
            auth: {
                clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
            },
            cache
        }; // Configure base config
        this._tokenRequest = {
            scopes
        }; // Configure token request with scopes
    }

    // retrieve access token
    getAccessToken() {
        return new Promise((resolve, reject) => {
            // Configure base instance of msal library
            const msalInstance = new UserAgentApplication(this._msalConfig);
            // Callback when login
            msalInstance.handleRedirectCallback(function loginCallBack(val) {
            });

            // Test user is auth
            if (msalInstance.getAccount()) {
                msalInstance.acquireTokenSilent(this._tokenRequest)
                    .then(response => {
                        // get access token from response
                        // response.accessToken
                        resolve(`Bearer ${response.accessToken}`);
                    })
                    .catch(err => {
                        // could also check if err instance of InteractionRequiredAuthError if you can import the class.
                        if (err.name === "InteractionRequiredAuthError") {
                            return msalInstance.acquireTokenRedirect(this._tokenRequest)
                                .then(response => {
                                    // get access token from response
                                    // response.accessToken
                                    this._acquireTokenSilent(); // Used to remove old data from storage
                                    resolve(`Bearer ${response.accessToken}`);
                                })
                                .catch(err => {
                                    // handle error.
                                    reject(err); // Error happend
                                });
                        }
                    });
            } else {
                // user is not logged in, you will need to log them in to acquire a token;
                msalInstance.loginRedirect(this._tokenRequest)
                    .then(response => {
                        msalInstance.acquireTokenSilent(this._tokenRequest)
                            .then(response => {
                                // get access token from response
                                // response.accessToken
                                resolve(`Bearer ${response.accessToken}`);
                            })
                            .catch(err => {
                                // could also check if err instance of InteractionRequiredAuthError if you can import the class.
                                if (err.name === "InteractionRequiredAuthError") {
                                    return msalInstance.acquireTokenRedirect(this._tokenRequest)
                                        .then(response => {
                                            // get access token from response
                                            // response.accessToken
                                            resolve(`Bearer ${response.accessToken}`);
                                        })
                                        .catch(err => {
                                            // handle error.
                                            reject(err); // Error happend
                                        });
                                }
                            });
                    });
            }
        })
    }
    // Workaround found on https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/759
    // msal js acquire token silent sometimes fail (token expired)
    // It is a manual acquire token
    _acquireTokenSilent = () => {
        const timestamp = Math.floor((new Date()).getTime() / 1000);
        let token = null;

        for (const key of Object.keys(localStorage)) {
            if (key.includes('"authority":')) {
                const val = JSON.parse(localStorage.getItem(key));

                if (val && val.expiresIn) {
                    // We have a (possibly expired) token

                    if (val.expiresIn > timestamp && val.idToken === val.accessToken) {
                        // Found the correct token
                        token = val.accessToken;
                    }
                    else {
                        // Clear old data
                        localStorage.removeItem(key);
                    }
                }
            }
        }
        if (token) return token; // If token exist, return it
        throw new Error('No valid token found'); // Else trhow error, there was no token
    };
}

// Workaround found on https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/759
// msal js acquire token silent sometimes fail (token expired)
// It is a manual acquire token
const acquireTokenSilent = () => {
    const timestamp = Math.floor((new Date()).getTime() / 1000);
    let token = null;

    for (const key of Object.keys(localStorage)) {
        if (key.includes('"authority":')) {
            const val = JSON.parse(localStorage.getItem(key));

            if (val && val.expiresIn) {
                // We have a (possibly expired) token

                if (val.expiresIn > timestamp && val.idToken === val.accessToken) {
                    // Found the correct token
                    token = val.idToken;
                }
                else {
                    // Clear old data
                    localStorage.removeItem(key);
                }
            }
        }
    }
    if (token) return token; // If token exist, return it
    throw new Error('No valid token found'); // Else trhow error, there was no token
};

/** Mise en cache des jetons
 * Fonction retournant une promesse (async)
 */
const tokenToCache = (token) => {
    return asyncLocalStorage.setItem('msal-access-token', `Bearer ${token}`); // Mise en local storage
}

/** Fonction asyncrone du local storage */
const asyncLocalStorage = {
    setItem: function (key, value) {
        return Promise.resolve().then(function () {
            localStorage.setItem(key, value);
        });
    },
    getItem: function (key) {
        return Promise.resolve().then(function () {
            return localStorage.getItem(key);
        });
    }
};