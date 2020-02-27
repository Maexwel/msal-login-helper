/** Configuration générale de la librairie msal.js */
export const msalConfig = (clientId, tenantId) => {
    return {
        auth: {
            clientId,
            authority: `https://login.microsoftonline.com/${tenantId}`,
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: true,
        }
    }
}
