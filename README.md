# copilot-webchat-adapter

A custom DirectLine-compatible adapter that connects [Copilot Studio](https://copilotstudio.microsoft.com) agents to [BotFramework WebChat](https://github.com/microsoft/BotFramework-WebChat) using the published `@microsoft/agents-copilotstudio-client` SDK.

## Why this exists

The official M365 Agents SDK includes a `CopilotStudioWebChat.createConnection()` helper, but it is **not published to npm** -- it only exists in the unreleased source tree. This adapter fills the gap by implementing the same DirectLine protocol surface using only the SDK's public API:

- `CopilotStudioClient.startConversationAsync()` -- begin a new conversation
- `CopilotStudioClient.askQuestionAsync(text, conversationId)` -- send a message

Additionally, this adapter supports **passing a `conversationId`** to resume an existing conversation, which the official helper does not expose.

## How it works

```
 ┌──────────┐   postActivity()   ┌──────────────────┐  askQuestionAsync()  ┌──────────────┐
 │  WebChat  │ ───────────────>  │  createConnection │ ──────────────────> │ Copilot Studio│
 │           │ <───────────────  │  (DirectLine shim)│ <────────────────── │    (SSE)      │
 └──────────┘   activity$        └──────────────────┘   Activity[]         └──────────────┘
```

The adapter exposes `connectionStatus$`, `activity$`, `postActivity()`, and `end()` -- the four members WebChat needs from a DirectLine connection. Internally it:

1. Calls `startConversationAsync()` on first subscription (unless `startConversation: false`)
2. Echoes user activities back to the transcript (WebChat expects this)
3. Forwards user text to `askQuestionAsync()` with the tracked `conversationId`
4. Emits agent response activities back through `activity$`

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
| `startConversation` | `boolean` | `!conversationId` | Whether to call `startConversationAsync()` on init |
| `showTyping` | `boolean` | `false` | Emit typing indicators while agent processes |

### Browser (no build step)

See `test-page/` for a complete working example that runs directly in the browser using ES modules and CDN imports. No bundler required.

```bash
cd test-page
npx serve .
# Open http://localhost:3000
```

## Architecture review

### What the adapter is (~120 LOC TypeScript)

- A thin translation layer between CopilotStudioClient's request/response API and WebChat's Observable-based DirectLine protocol
- Stateful: tracks `conversationId`, `sequence` counter, `ended` flag, and the single `activitySubscriber`
- No external dependencies beyond `rxjs`, `uuid`, and the SDK itself

### Production considerations

**Pros**

- **Minimal surface area**: one function, three options, four DirectLine members. Easy to reason about and maintain.
- **No server component**: runs entirely client-side. Auth token is acquired via MSAL in the browser.
- **Streaming support**: the SDK's browser ESM build (v1.2.3) includes async generator methods (`sendActivityStreaming`, `startConversationStreaming`) that yield each SSE chunk as it arrives. The adapter can use these to emit activities to WebChat in real-time. Currently the adapter uses `askQuestionAsync` which is a convenience wrapper that collects all chunks into an array first -- switching to `sendActivityStreaming` would enable true real-time streaming.
- **Conversation resume**: supports passing a `conversationId` to resume sessions, which the official (unreleased) helper does not expose.

**Streaming activity protocol**

The Copilot Studio SSE API streams activities as individual chunks. A typical response produces 100-300+ activities:

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

The SDK's browser ESM build exposes two levels of API:

| Method | Type | Behavior |
|--------|------|----------|
| `sendActivityStreaming(activity)` | `async*` generator | Yields each activity as it arrives from SSE. Use this for real-time streaming. |
| `askQuestionAsync(text, convId)` | `async` | Collects all streaming activities into an array, returns when complete. |
| `startConversationStreaming()` | `async*` generator | Same as above, for the initial greeting. |
| `startConversationAsync()` | `async` | Collects greeting activities into array. |

To enable real-time streaming in the adapter, replace `askQuestionAsync` with a `for await` loop over `sendActivityStreaming`, emitting each activity to `activity$` as it yields.

> **Note:** The local Agents-for-js source tree has a different implementation of `CopilotStudioClient` that does NOT have the streaming generators. The published v1.2.3 browser.mjs is a newer build with the async generator architecture.

**Current limitations**

| Limitation | Details |
|------------|---------|
| **Text-only sends** | `askQuestionAsync(text, conversationId)` only sends text. Attachments and card actions are reduced to `activity.text`. The API accepts full Activity payloads -- could bypass `askQuestionAsync` to send them. |
| **Single subscriber** | One `activity$` subscriber at a time. Could use a `Subject` if multiple subscribers are needed. |
| **No reconnection** | If a request fails, the connection errors out. Retry logic can be added in `postActivity` or at the application level. |
| **SDK ESM/CJS mismatch** | `startConversationAsync` returns `Activity[]` in browser ESM but `Activity` in CJS. Handled with `Array.isArray()` in the browser version. |
| **Buffered responses (easy fix)** | Currently uses `askQuestionAsync` which collects all streaming chunks before returning. Switch to `sendActivityStreaming` async generator to emit activities in real-time. |

**Server-side considerations**

- Conversation resume routes the request to the right server-side conversation via the URL path, but Copilot Studio sessions expire after inactivity. There is no explicit error for an expired session -- the agent simply responds without prior context.
- Each `postActivity` call is a separate HTTP request to the SSE endpoint. There is no persistent WebSocket connection.

### Comparison with the official (unreleased) adapter

| Feature | This adapter | `CopilotStudioWebChat.createConnection()` |
|---------|-------------|-------------------------------------------|
| Published to npm | Yes (or can be) | No -- only in unreleased source tree |
| Conversation resume | Yes | No |
| Streaming | Buffered (uses `askQuestionAsync`); can switch to `sendActivityStreaming` for real-time | Buffered (uses `sendActivity` from old source) |
| Send full Activity objects | Not yet (`askQuestionAsync` is text-only) | Yes (`sendActivity` accepts full Activity) |
| `ended` guard | Yes | No |
| Code size | ~120 LOC | ~150 LOC (+ ~180 LOC comments/docs) |
| Dependencies | rxjs, uuid, SDK | rxjs, uuid, SDK |

## Test page

The `test-page/` directory contains a standalone browser test harness:

```
test-page/
  index.html           # UI shell (WebChat CDN, MSAL CDN, importmap)
  app.js               # Entry point, agent configs, Connect flow
  acquireToken.js       # MSAL popup auth
  createConnection.js   # Browser ES module port of src/createConnection.ts
```

### Features
- Agent dropdown with preconfigured Copilot Studio agents
- Optional conversation ID input for resume testing
- Start conversation toggle (controls whether greeting is triggered)
- Status bar with copyable conversation ID
- No build step -- pure ES modules via importmap

### Test flow
1. Select an agent, click Connect
2. MSAL popup authenticates, WebChat renders
3. Send messages, observe responses
4. Copy the conversation ID from the status bar
5. Reload, paste the conversation ID, Connect again to test resume

## Commands

```bash
npm install    # Install dependencies
npm run build  # Build TypeScript to dist/
```

## Project structure

```
src/
  createConnection.ts   # Main adapter (TypeScript, ~120 LOC)
  types.ts              # WebChatConnection and CreateConnectionOptions interfaces
  index.ts              # Public exports
test-page/              # Browser test harness (no build step)
dist/                   # Compiled output
```

## Known issues

- The TypeScript `src/createConnection.ts` does not yet handle the array return type from the browser ESM build of `startConversationAsync`. The `test-page/createConnection.js` browser version does.
- Debug logging (`console.log('[adapter]...')`) is present in `test-page/createConnection.js` for development. Remove before shipping.
