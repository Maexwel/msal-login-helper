import { UserAgentApplication } from 'msal'
import { msalConfig } from './config'
// Ce module permet d'encapsuler la librairie msal.js
// Grâce à ce module, la connexion et récupération de token est simplifiée
// Les token récupéré (msal-access-token) permettent d'effectuer des requêtes
// Sur l'API GRAPH


/** Méthode de login en utilisant les fonctions msal
 * Le login doit renvoyer un objet contenant les ID et token
 * Il faut savoir que la config est faite pour mettre en cache (localstorage)
 * les informations et token récupéré de la libraire.
 * Cette fonction est une promesse (async)
 * Cette fonction a besoin pour fonctionner du clientId(app graph), le tenantId(domain)
 *  et des scopes authorisés
 */
export const msalLogin = (clientId, tenantId, scopes) => {
    return new Promise((resolve, reject) => {
        const msalLoginAgent = new UserAgentApplication(msalConfig(clientId, tenantId)) // Déclaration de l'agent de login
        msalLoginAgent.handleRedirectCallback((error, response) => {
            return getAccessToken(msalLoginAgent, scopes) // Récupération du jeton d'accès
                .then(response => {
                    tokenToCache(response).then(() => {
                        return resolve({ accessToken: `Bearer ${response}` }) // on renvoit la réponse
                    })
                })
                .catch(err => {
                    return reject(err)
                })
        }) // Déclaration de la gestion de callback
        if (msalLoginAgent.getAccount()) {
            // Le compte est disponible
            return getAccessToken(msalLoginAgent, scopes) // Récupération du jeton d'accès
                .then(response => {
                    tokenToCache(response).then(() => {
                        return resolve({ accessToken: `Bearer ${response}` }) // on renvoit la réponse
                    })
                })
                .catch(err => {
                    return reject(err)
                })
        } else {
            // Il faut un login
            return msalLoginAgent.loginRedirect(scopes)
        }
    })
}

/** Récupération du token d'accès à utiliser pour les appels rest, graph, ... 
 * Cette fonction est une promesse
*/
const getAccessToken = (msalLoginAgent, scopes) => {
    return new Promise((resolve, reject) => {
        try {
            const token = acquireTokenSilent(); // Custom call to acquiretoken silent function
            resolve(token); // Token foudn, resolve it
        } catch (err) {
            return msalLoginAgent.loginRedirect(scopes)
                .then(response => {
                    return resolve(response.accessToken) // response contient accessToken
                })
                .catch(err => {
                    return reject(err)
                });
        }
    });
};

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
}

/** Mise en cache des jetons
 * Fonction retournant une promesse (async)
 */
const tokenToCache = (token) => {
    return asyncLocalStorage.setItem('msal-access-token', `Bearer ${token}`) // Mise en local storage
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