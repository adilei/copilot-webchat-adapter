# copilot-webchat-adapter

> **Experimental** -- This is an unofficial, community-driven adapter for learning and prototyping purposes. It is not supported by Microsoft and is not intended for production use. Use at your own risk.

A custom DirectLine-compatible adapter that connects [Copilot Studio](https://copilotstudio.microsoft.com) agents to [BotFramework WebChat](https://github.com/microsoft/BotFramework-WebChat) using the published `@microsoft/agents-copilotstudio-client` SDK.

## Why this exists

The official `CopilotStudioWebChat.createConnection()` in the M365 Agents SDK does not support:

- **Conversation resume** -- passing a `conversationId` to rehydrate an existing conversation
- **Controlling the start conversation event** -- choosing whether or not to call `startConversation` (e.g., skipping it when resuming)
- **Activity history** -- the underlying API has no way to fetch past activities or list previous conversations

This adapter fills those gaps while implementing the same DirectLine protocol surface, using the SDK's streaming async generators for real-time activity delivery.

## How it works

```
 ┌──────────┐   postActivity()   ┌──────────────────┐  sendActivityStreaming()  ┌──────────────┐
 │  WebChat  │ ───────────────>  │  createConnection │ ──────────────────────> │ Copilot Studio│
 │           │ <───────────────  │  (DirectLine shim)│ <────────────────────── │    (SSE)      │
 └──────────┘   activity$        └──────────────────┘   yield Activity          └──────────────┘
```

The adapter exposes `connectionStatus$`, `activity$`, `postActivity()`, and `end()` -- the four members WebChat needs from a DirectLine connection. Internally it:

1. Streams `startConversationStreaming()` on first subscription (unless `startConversation: false`), emitting each activity as it arrives
2. Echoes user activities back to the transcript (WebChat expects this)
3. Streams `sendActivityStreaming()` with the full Activity object and tracked `conversationId`, emitting each response chunk in real-time
4. Tracks `conversationId` from response activities for conversation resume

## Installation

```bash
npm install copilot-webchat-adapter @microsoft/agents-copilotstudio-client
```

## Usage

### TypeScript / Node.js (bundled app)

```typescript
import { createConnection } from 'copilot-webchat-adapter'
import { CopilotStudioClient, ConnectionSettings } from '@microsoft/agents-copilotstudio-client'

// Configure
const settings = new ConnectionSettings()
settings.environmentId = 'your-environment-id'
settings.agentIdentifier = 'your-agent-identifier'
settings.tenantId = 'your-tenant-id'
settings.appClientId = 'your-app-client-id'
settings.cloud = 'Prod'

const client = new CopilotStudioClient(settings, accessToken)

// New conversation
const directLine = createConnection(client, { showTyping: true })

// Resume existing conversation
const directLine = createConnection(client, {
  conversationId: savedConversationId,
  showTyping: true,
})

// Render WebChat
WebChat.renderWebChat({ directLine }, document.getElementById('webchat'))
```

### Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `conversationId` | `string` | `undefined` | Existing conversation ID to resume |
| `startConversation` | `boolean` | `!conversationId` | Whether to call `startConversationStreaming()` on init |
| `showTyping` | `boolean` | `false` | Emit a synthetic typing indicator before each request |
| `getHistoryFromExternalStorage` | `(id: string) => Promise<Activity[]>` | `undefined` | Optional callback to fetch stored activities on resume |

### Conversation history

The adapter can replay stored activities on resume via the `getHistoryFromExternalStorage` callback. The adapter is agnostic about storage -- it only needs a function that returns activities for a given conversation ID. You can use localStorage, IndexedDB, a server-side API, or anything else.

**Saving activities** -- Use a WebChat Redux middleware to intercept incoming activities and save them to your store:

```javascript
const store = WebChat.createStore({}, () => next => action => {
  if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
    const { activity } = action.payload
    if (activity.type === 'message' && directLine.conversationId) {
      myStore.save(directLine.conversationId, activity)
    }
  }
  return next(action)
})

WebChat.renderWebChat({ directLine, store }, document.getElementById('webchat'))
```

**Loading activities on resume** -- Pass your retrieval function when creating the connection:

```typescript
const directLine = createConnection(client, {
  conversationId: savedConversationId,
  getHistoryFromExternalStorage: (id) => myStore.getActivities(id),
})
```

The adapter replays the returned activities through `activity$` before any new messages stream in, maintaining proper sequence numbering so WebChat renders everything in order.

The example above only saves `message` activities. Depending on your scenario, you may want to store other activity types too (e.g., adaptive card submissions, events). Keep in mind that non-message activities may need special handling or transformation when replayed, for example, disabling adaptive cards that were already submitted.

If `getHistoryFromExternalStorage` throws, the adapter continues without history (graceful degradation).

> **Note:** `getHistoryFromExternalStorage` is intentionally optional. When the SDK adds native activity fetching, the adapter can call it internally by default, and this callback becomes an override.

See `test-page/` for a complete working example using localStorage.

### Browser (no build step)

See `test-page/` for a complete working example that runs directly in the browser using ES modules and CDN imports. No bundler required.

```bash
cd test-page
npx serve .
```

## Streaming

The adapter uses the SDK's async generator methods for real-time streaming. Each call to `sendActivityStreaming()` yields 100-300+ activities as individual SSE chunks:

```
typing                              → typing indicator
typing  "Generating plan..."        → agent status update
event                               → search/reasoning/citation metadata
typing  "Searching through..."      → agent status update
typing  ""                          → streaming text begins
typing  "**The"                     → incremental text (accumulated)
typing  "**The Wizard"              → incremental text (accumulated)
typing  "**The Wizard of Oz"        → incremental text (accumulated)
...                                 → (hundreds of incremental chunks)
message "**The Wizard of Oz..."     → final complete message
```

Every chunk is emitted to `activity$` as it arrives -- WebChat receives activities in real-time.

The SDK's browser ESM build exposes two levels of API:

| Method | Type | Behavior |
|--------|------|----------|
| `startConversationStreaming()` | `async*` generator | Yields each greeting activity as it arrives. **Used by this adapter.** |
| `sendActivityStreaming(activity)` | `async*` generator | Yields each response activity as it arrives. **Used by this adapter.** |
| `startConversationAsync()` | `async` | Convenience wrapper -- collects all streaming chunks into an array. |
| `askQuestionAsync(text, convId)` | `async` | Convenience wrapper -- text-only, collects all chunks into an array. |

> **Note:** The streaming generators are also used by the [official SDK adapter](https://github.com/microsoft/Agents-for-js/blob/main/packages/agents-copilotstudio-client/src/copilotStudioWebChat.ts) and [public samples](https://github.com/microsoft/Agents/blob/main/samples/nodejs/copilotstudio-client/src/index.ts).

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

> **Tip:** If "Power Platform API" doesn't appear, follow [Step 2 in the Power Platform API docs](https://learn.microsoft.com/power-platform/admin/programmability-authentication-v2#step-2-configure-api-permissions) to add it to your tenant.

### Get your agent's metadata

1. In [Copilot Studio](https://copilotstudio.microsoft.com), open your agent
2. Go to **Settings** > **Advanced** > **Metadata** and note:
   - **Schema name** (this is the `agentIdentifier`)
   - **Environment ID**

### Configure and run

```bash
npm install
npm run build
cp test-page/agents.sample.js test-page/agents.js
```

Edit `test-page/agents.js` with your values:

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

### Test flow

1. Open `http://localhost:5500/test-page/`, select an agent, click **Connect**
2. MSAL popup authenticates, WebChat renders with streaming responses
3. Copy the conversation ID from the status bar
4. Reload, paste the conversation ID, click Connect to test resume

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

