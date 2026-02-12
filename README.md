# copilot-webchat-adapter

> **Experimental** -- This is an unofficial, community-driven adapter for learning and prototyping purposes. It is not supported by Microsoft and is not intended for production use. Use at your own risk.

A custom DirectLine-compatible adapter that connects [Copilot Studio](https://copilotstudio.microsoft.com) agents to [BotFramework WebChat](https://github.com/microsoft/BotFramework-WebChat) using the published `@microsoft/agents-copilotstudio-client` SDK.

## Why this exists

The official M365 Agents SDK includes a `CopilotStudioWebChat.createConnection()` helper, but it is **not published to npm** -- it only exists in the unreleased source tree. This adapter fills the gap by implementing the same DirectLine protocol surface using the SDK's API:

- `startConversationStreaming()` -- async generator that yields greeting activities as SSE chunks arrive
- `sendActivityStreaming(activity, conversationId)` -- async generator that yields response activities in real-time

Additionally, this adapter supports **passing a `conversationId`** to resume an existing conversation, which the official helper does not expose.

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
npm install copilot-webchat-adapter
```

Peer dependencies (install separately):
```bash
npm install @microsoft/agents-copilotstudio-client @microsoft/agents-activity rxjs
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

> **Note:** The streaming generators exist in the published v1.2.3 browser ESM build but are not declared in the TypeScript `.d.ts` files. The local Agents-for-js source tree has a different (older) implementation without streaming generators.

## Architecture

### What the adapter is

- A thin translation layer (~120 LOC) between CopilotStudioClient's streaming API and WebChat's Observable-based DirectLine protocol
- Stateful: tracks `conversationId`, `sequence` counter, `ended` flag, and the single `activitySubscriber`
- No external dependencies beyond `rxjs`, `uuid`, and the SDK itself

### Comparison with the official (unreleased) adapter

| Feature | This adapter | `CopilotStudioWebChat.createConnection()` |
|---------|-------------|-------------------------------------------|
| Available on npm | Yes (or can be) | No -- unreleased source only |
| Conversation resume | Yes | No |
| Real-time streaming | Yes (`sendActivityStreaming`) | No (uses buffered `sendActivity`) |
| Sends full Activity objects | Yes | Yes |
| `ended` guard | Yes | No |
| Code size | ~120 LOC | ~150 LOC (+ ~180 LOC comments/docs) |

### Current limitations

| Limitation | Details |
|------------|---------|
| **Single subscriber** | One `activity$` subscriber at a time. Could use a `Subject` if multiple subscribers are needed. |
| **No reconnection** | If a request fails, the connection errors out. Retry logic can be added at the application level. |
| **Streaming generators not typed** | `sendActivityStreaming` and `startConversationStreaming` exist in the browser ESM build but are not in the published `.d.ts` declarations. |

### Server-side considerations

- Conversation resume routes the request to the right server-side conversation via the URL path, but Copilot Studio sessions expire after inactivity. There is no explicit error for an expired session -- the agent simply responds without prior context.
- Each `postActivity` call is a separate HTTP request to the SSE endpoint. There is no persistent WebSocket connection.

## Test page

The `test-page/` directory contains a standalone browser test harness:

```
test-page/
  index.html           # UI shell (WebChat CDN, MSAL CDN, importmap)
  app.js               # Entry point, Connect flow
  agents.sample.js     # Template for agent configs (copy to agents.js)
  acquireToken.js      # MSAL popup auth
  createConnection.js  # Browser ES module version of the adapter
```

### Setup

```bash
cd test-page
cp agents.sample.js agents.js
# Edit agents.js with your Copilot Studio agent configs
npx serve .
```

### Features
- Agent dropdown populated from `agents.js`
- Optional conversation ID input for resume testing (disables start conversation when set)
- Start conversation toggle (controls whether greeting is triggered)
- Status bar with copyable conversation ID
- Real-time streaming responses
- No build step -- pure ES modules via importmap

### Test flow
1. Select an agent, click Connect
2. MSAL popup authenticates, WebChat renders
3. Send messages, observe streaming responses
4. Copy the conversation ID from the status bar
5. Reload, paste the conversation ID, Connect again to test resume

## Project structure

```
src/
  createConnection.ts   # Main adapter (TypeScript, ~120 LOC)
  types.ts              # WebChatConnection and CreateConnectionOptions interfaces
  index.ts              # Public exports
test-page/              # Browser test harness (no build step)
dist/                   # Compiled output
```

## Commands

```bash
npm install    # Install dependencies
npm run build  # Build TypeScript to dist/
```

## Known issues

- The TypeScript `src/createConnection.ts` uses the buffered `startConversationAsync`/`askQuestionAsync` methods. The browser `test-page/createConnection.js` uses the streaming generators. The TypeScript version should be updated to use streaming once the generators are available in the `.d.ts` declarations.
