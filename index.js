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
        const msalLoginAgent = new UserAgentApplication(msalConfig(clientId, tenantId))
        if (msalLoginAgent.getAccount()) {
            // Le compte est disponible
            return getAccessToken(msalLoginAgent, scopes) // Récupération du jeton d'accès
                .then(response => {
                    tokenToCache(response.accessToken)
                    return resolve({ ...response, accessToken: `Bearer ${response.accessToken}` }) // on renvoit la réponse
                })
                .catch(err => {
                    return reject(err)
                })
        } else {
            // Il faut un login par popUp
            return msalLoginAgent.loginPopup(scopes) // Pop up de login
                .then(response => {
                    //new UserAgentApplication(msalConfig) // Pour fermer la popup
                    return getAccessToken(msalLoginAgent, scopes) // Récupération du jeton d'accès
                        .then(response => {
                            tokenToCache(response.accessToken).then(() => {
                                return resolve({ ...response, accessToken: `Bearer ${response.accessToken}` }) // on renvoit la réponse
                            })
                        })
                        .catch(err => {
                            return reject(err)
                        })
                })
                .catch(err => {
                    return reject(err)
                })
        }
    })
}

/** Récupération du token d'accès à utiliser pour les appels rest, graph, ... 
 * Cette fonction est une promesse
*/
const getAccessToken = (msalLoginAgent, scopes) => {
    return new Promise((resolve, reject) => {
        return msalLoginAgent.acquireTokenSilent(scopes)
            .then(response => {
                return resolve(response) // response contient accessToken
            })
            .catch(err => {
                if (err.name === "InteractionRequiredAuthError") {
                    return msalLoginAgent.acquireTokenPopup(scopes)
                        .then(response => {
                            return resolve(response) // response contient accessToken
                        })
                        .catch(err => {
                            return reject(err)
                        })
                } else {
                    return reject(err)
                }
            })
    })

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