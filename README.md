# copilot-webchat-adapter

> **Experimental** -- This is an unofficial, community-driven adapter for learning and prototyping purposes. It is not supported by Microsoft and is not intended for production use. Use at your own risk.

A custom DirectLine-compatible adapter that connects [Copilot Studio](https://copilotstudio.microsoft.com) agents to [BotFramework WebChat](https://github.com/microsoft/BotFramework-WebChat) using the published `@microsoft/agents-copilotstudio-client` SDK.

## Why this exists

The official `CopilotStudioWebChat.createConnection()` in the M365 Agents SDK does not support:

- **Conversation resume** -- passing a `conversationId` to rehydrate an existing conversation
- **Controlling the start conversation event** -- choosing whether to call `startConversation` (e.g., skipping it when resuming)
- **Activity history** -- fetching past activities or listing previous conversations

This adapter fills those gaps while implementing the same DirectLine protocol surface.

## Quick start

```bash
npm install copilot-webchat-adapter @microsoft/agents-copilotstudio-client
```

```typescript
import { createConnection } from 'copilot-webchat-adapter'
import { CopilotStudioClient, ConnectionSettings } from '@microsoft/agents-copilotstudio-client'
import { createLocalStorageStore } from './activityStore.js'

// 1. Configure the SDK client
const settings = new ConnectionSettings()
settings.environmentId = 'your-environment-id'
settings.agentIdentifier = 'your-agent-identifier'
settings.tenantId = 'your-tenant-id'
settings.appClientId = 'your-app-client-id'
settings.cloud = 'Prod'

const client = new CopilotStudioClient(settings, accessToken)

// 2. Create the DirectLine connection (with optional resume + history)
const activityStore = createLocalStorageStore()

const directLine = createConnection(client, {
  conversationId: savedConversationId,           // omit for a new conversation
  getHistoryFromExternalStorage: (id) =>         // omit if you don't need history
    activityStore.getActivities(id),
  showTyping: true,
})

// 3. Set up a WebChat Redux store to save activities as they arrive
const store = WebChat.createStore({}, () => next => action => {
  if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
    const { activity } = action.payload
    if (activity.type === 'message' && directLine.conversationId) {
      activityStore.saveActivity(directLine.conversationId, activity)
    }
  }
  return next(action)
})

// 4. Render
WebChat.renderWebChat({ directLine, store }, document.getElementById('webchat'))
```

## Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `conversationId` | `string` | `undefined` | Existing conversation ID to resume |
| `startConversation` | `boolean` | `!conversationId` | Whether to call `startConversationStreaming()` on init |
| `showTyping` | `boolean` | `false` | Emit a synthetic typing indicator before each request |
| `getHistoryFromExternalStorage` | `(id: string) => Promise<Activity[]>` | `undefined` | Callback to fetch stored activities on resume |

## Conversation history

The adapter and the consumer split responsibility:

- **Adapter** fetches and replays stored activities through `activity$` on resume, maintaining proper sequence numbering so WebChat renders everything in order. If the callback throws, the adapter continues without history.
- **Consumer** saves activities as they arrive (via a WebChat Redux middleware) and provides the retrieval function.

The adapter is agnostic about storage. You can use localStorage, IndexedDB, a server-side API, or anything else. See `test-page/activityStore.js` for a working localStorage implementation.

The example above only saves `message` activities. Depending on your scenario, you may want to store other activity types too (e.g., adaptive card submissions, events). Keep in mind that non-message activities may need special handling when replayed, for example disabling adaptive cards that were already submitted.

> **Note:** `getHistoryFromExternalStorage` is intentionally optional. When the SDK adds native activity fetching, the adapter can call it internally by default and this callback becomes an override.

## Test page

The `test-page/` directory contains a standalone browser test harness with no build step required.

### Prerequisites

1. A published agent in [Copilot Studio](https://copilotstudio.microsoft.com)
2. An Entra ID app registration (public client / SPA)

### Create an app registration in Entra ID

1. Open [Azure Portal](https://portal.azure.com) and navigate to **Entra ID**
2. Create a new **App Registration**:
   - Name: anything you like
   - Supported account types: "Accounts in this organization directory only"
   - Platform: **Single-page application (SPA)**
   - Redirect URI: `http://localhost:5500` (match your local server port)
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Go to **API Permissions** > **Add a permission**:
   - Tab: "APIs my organization uses"
   - Search for **Power Platform API**
   - Delegated permissions > **CopilotStudio** > check **CopilotStudio.Copilots.Invoke**
   - Click "Add Permissions"
5. (Optional) Click "Grant admin consent"

> If "Power Platform API" doesn't appear, follow [Step 2 in the Power Platform API docs](https://learn.microsoft.com/power-platform/admin/programmability-authentication-v2#step-2-configure-api-permissions) to add it to your tenant.

### Configure and run

```bash
npm install
npm run build
cp test-page/agents.sample.js test-page/agents.js
```

Edit `test-page/agents.js` with your agent's metadata (found in Copilot Studio > Settings > Advanced > Metadata):

```js
export const agents = {
  'my-agent': {
    name: 'My Agent',
    environmentId: '<environment-id>',
    agentIdentifier: '<schema-name>',
    tenantId: '<directory-tenant-id>',
    appClientId: '<application-client-id>',
  },
}
export const defaultAgent = 'my-agent'
```

Then serve from the project root (the test page imports the built adapter from `dist/`):

```bash
npx serve . -l 5500
```

Open `http://localhost:5500/test-page/`, select an agent, and click **Connect**. To test resume, copy the conversation ID from the status bar, reload, paste it, and connect again.

## Project structure

```
src/
  createConnection.ts   # Main adapter (TypeScript, ~120 LOC)
  types.ts              # WebChatConnection and CreateConnectionOptions interfaces
  index.ts              # Public exports
test-page/              # Browser test harness (no build step)
  activityStore.js      # Sample localStorage-backed activity store
dist/                   # Compiled output
```

## Commands

```bash
npm install    # Install dependencies
npm run build  # Build TypeScript to dist/
```
