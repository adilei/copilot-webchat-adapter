/**
 * MSAL authentication for Copilot Studio agents.
 * Based on the official sample from microsoft/Agents SDK.
 */

import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client'

export async function acquireToken(settings) {
  const msalInstance = new window.msal.PublicClientApplication({
    auth: {
      clientId: settings.appClientId,
      authority: `https://login.microsoftonline.com/${settings.tenantId}`,
    },
  })

  await msalInstance.initialize()

  const loginRequest = {
    scopes: [CopilotStudioClient.scopeFromSettings(settings)],
    redirectUri: window.location.origin,
  }

  // Try silent token acquisition first, fall back to popup
  try {
    const accounts = await msalInstance.getAllAccounts()
    if (accounts.length > 0) {
      const response = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      })
      return response.accessToken
    }
  } catch (e) {
    if (!(e instanceof window.msal.InteractionRequiredAuthError)) {
      throw e
    }
  }

  const response = await msalInstance.loginPopup(loginRequest)
  return response.accessToken
}
