import { Configuration, LogLevel } from '@azure/msal-node';
import { getCloud, GRAPH_SCOPES, qualifyScopes } from './cloudConfig.js';

const cloud = getCloud();

/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/configuration.md
 *
 * Environment variables:
 * - CLIENT_ID: The Application (client) ID from your Azure AD app registration (required)
 * - AUTHORITY: Your Azure AD tenant ID (required)
 * - M365_CLOUD: Selects the Microsoft 365 national cloud (optional; see cloudConfig.ts)
 * - GRAPH_BASE_URL: Overrides the Microsoft Graph origin (optional; see cloudConfig.ts)
 * - LOGIN_AUTHORITY_HOST: Overrides the Entra ID login origin (optional; see cloudConfig.ts)
 */
export const msalConfig: Configuration = {
    auth: {
        clientId: process.env.CLIENT_ID!,
        authority: `${cloud.authorityHost}/${process.env.AUTHORITY!}`,
        // MSAL already ships instance-discovery metadata for the preset login
        // hosts; knownAuthorities is only needed to trust a custom LOGIN_AUTHORITY_HOST.
        ...(cloud.customAuthority ? { knownAuthorities: [cloud.authorityHostname] } : {}),
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel: any, message: any, containsPii: any) {
                // Show log messages if DEBUG is enabled. stdout is reserved for the MCP
                // stdio transport, so MSAL logs must go to stderr.
                if (process.env.DEBUG && !containsPii) {
                    console.error(`MSAL: ${message}`);
                }
            },
            piiLoggingEnabled: false,
            logLevel: LogLevel.Error,
        },
    },
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const loginRequest = {
    scopes: qualifyScopes(GRAPH_SCOPES, cloud.graphBaseUrl),
};