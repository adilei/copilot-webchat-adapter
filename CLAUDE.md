# Project: copilot-webchat-adapter

Custom DirectLine-compatible WebChat adapter for CopilotStudioClient with conversation resume support.

## Tech Stack

- TypeScript
- RxJS (Observables for DirectLine protocol)
- `@microsoft/agents-copilotstudio-client` (published npm v1.2.3)
- `@microsoft/agents-activity` (Activity types)

## Purpose

Replaces the unreleased `CopilotStudioWebChat.createConnection()` with a custom implementation that:
1. Supports passing a `conversationId` to resume existing conversations
2. Configurable `startConversation` behavior (skip when resuming)
3. Uses only published public API (`startConversationAsync`, `askQuestionAsync`)

## Key Files

| File | Purpose |
|------|---------|
| `src/createConnection.ts` | Main adapter - creates DirectLine-compatible connection |
| `src/types.ts` | `WebChatConnection` and `CreateConnectionOptions` interfaces |
| `src/index.ts` | Public exports |

## Usage

```typescript
import { createConnection } from 'copilot-webchat-adapter'
import { CopilotStudioClient, ConnectionSettings } from '@microsoft/agents-copilotstudio-client'

const client = new CopilotStudioClient(settings, token)

// New conversation
const directLine = createConnection(client)

// Resume existing conversation
const directLine = createConnection(client, {
  conversationId: 'saved-id',
  showTyping: true,
})
```

## Commands

```bash
npm install    # Install dependencies
npm run build  # Build TypeScript to dist/
```

## Test Page

`test-page/` contains a standalone browser test harness (no build step). Run with `npx serve test-page`.

## Important Notes

- Built against published npm `@microsoft/agents-copilotstudio-client` v1.2.3
- The published SDK only has `askQuestionAsync(text, conversationId?)` - no `sendActivity()`
- When the SDK publishes `sendActivity()`, this adapter could be updated to support sending non-text activities
- The `CopilotStudioWebChat` class exists only in the unreleased local Agents-for-js clone
- **SDK browser ESM vs CJS difference**: `startConversationAsync` returns `Activity[]` in the browser ESM build but `Activity` (single) in the CJS build. The test-page `createConnection.js` handles both; the TypeScript version does not yet.
- **ConnectionSettings requires manual property assignment**: the published v1.2.3 has a no-arg constructor only. Must set `environmentId`, `agentIdentifier`, `tenantId`, `appClientId`, and `cloud` ('Prod') individually.
- **Streaming**: the published v1.2.3 browser.mjs has `async*` generator methods (`sendActivityStreaming`, `startConversationStreaming`) that yield each SSE chunk in real-time. `askQuestionAsync`/`startConversationAsync` are convenience wrappers that collect into arrays. To enable real-time streaming, use the generators directly. The local Agents-for-js source tree does NOT have these generators -- it's a fundamentally different (older) implementation that buffers everything and only returns `message` activities.
- **Official adapter comparison**: read and analyzed `copilotStudioWebChat.ts` from the unreleased source tree. Same architecture as ours, but uses `client.sendActivity(activity)` (unpublished method, accepts full Activity) instead of `askQuestionAsync` (text-only). Does NOT support conversationId resume. Does NOT have an `ended` guard.
- **Conversation resume**: works at the HTTP level (conversationId in URL path) but server-side session expiry can cause the agent to lose context silently (responds with fallback instead of contextual answer).
