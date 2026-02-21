import { strict as assert } from 'assert'
import { describe, it, beforeEach, afterEach } from 'node:test'
import { createSandbox, SinonSandbox, SinonStub } from 'sinon'
import { Activity } from '@microsoft/agents-activity'
import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client'
import { firstValueFrom } from 'rxjs'

import { createConnection } from '../src/createConnection.js'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Creates a minimal Activity-like object. */
function makeActivity (overrides: Partial<Activity> = {}): Activity {
  return Activity.fromObject({
    type: 'message',
    text: 'hello',
    ...overrides,
  })
}

/** Creates a fake CopilotStudioClient with stubbed streaming methods. */
function createMockClient (sandbox: SinonSandbox, opts: {
  greetingActivities?: Partial<Activity>[]
  responseActivities?: Partial<Activity>[]
} = {}) {
  const greetingActivities = opts.greetingActivities ?? [
    {
      type: 'message',
      text: 'Hi there!',
      conversation: { id: 'conv-from-server', name: 'Bot' },
      replyToId: 'should-be-stripped',
    },
  ]
  const responseActivities = opts.responseActivities ?? [
    { type: 'message', text: 'Response', conversation: { id: 'conv-from-server' } },
  ]

  async function * fakeStartConversationStreaming (): AsyncGenerator<Activity> {
    for (const a of greetingActivities) {
      yield Activity.fromObject(a)
    }
  }

  async function * fakeSendActivityStreaming (): AsyncGenerator<Activity> {
    for (const a of responseActivities) {
      yield Activity.fromObject(a)
    }
  }

  const client = {
    startConversationStreaming: sandbox.stub().callsFake(fakeStartConversationStreaming),
    sendActivityStreaming: sandbox.stub().callsFake(fakeSendActivityStreaming),
  }

  return client as unknown as CopilotStudioClient & {
    startConversationStreaming: SinonStub
    sendActivityStreaming: SinonStub
  }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('createConnection', function () {
  let sandbox: SinonSandbox

  beforeEach(function () {
    sandbox = createSandbox()
  })

  afterEach(function () {
    sandbox.restore()
  })

  // =========================================================================
  // New conversation (default behavior)
  // =========================================================================
  describe('new conversation (default)', function () {
    it('should call startConversationStreaming and emit greeting activities', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      const activities: Partial<Activity>[] = []
      const done = new Promise<void>((resolve) => {
        conn.activity$.subscribe({
          next: (a) => activities.push(a),
          complete: () => resolve(),
        })
      })

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()
      await done

      assert(client.startConversationStreaming.calledOnce, 'startConversationStreaming should be called once')

      const messageActivities = activities.filter((a) => a.type === 'message')
      assert(messageActivities.length >= 1, 'should emit at least one message activity')
      assert.strictEqual(messageActivities[0].text, 'Hi there!')
    })

    it('should add timestamp and webchat:sequence-id to emitted activities', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      const activities: Partial<Activity>[] = []
      const done = new Promise<void>((resolve) => {
        conn.activity$.subscribe({
          next: (a) => activities.push(a),
          complete: () => resolve(),
        })
      })

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()
      await done

      for (const a of activities) {
        assert(a.timestamp, 'activity should have a timestamp')
        assert(
          (a.channelData as Record<string, unknown>)?.['webchat:sequence-id'] !== undefined,
          'activity should have webchat:sequence-id'
        )
      }
    })

    it('should transition connectionStatus$ to 2 on subscribe', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      const statuses: number[] = []
      conn.connectionStatus$.subscribe((s) => statuses.push(s))
      conn.activity$.subscribe({})

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()

      assert(statuses.includes(2), 'connectionStatus$ should reach 2 (connected)')
    })

    it('should strip replyToId from greeting activities', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      const activities: Partial<Activity>[] = []
      const done = new Promise<void>((resolve) => {
        conn.activity$.subscribe({
          next: (a) => activities.push(a),
          complete: () => resolve(),
        })
      })

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()
      await done

      const messageActivities = activities.filter((a) => a.type === 'message')
      for (const a of messageActivities) {
        assert.strictEqual(a.replyToId, undefined, 'replyToId should be stripped')
      }
    })

    it('should capture conversationId from first response activity', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      assert.strictEqual(conn.conversationId, undefined, 'conversationId should be undefined before subscribe')

      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert.strictEqual(conn.conversationId, 'conv-from-server', 'conversationId should be captured from response')
      conn.end()
    })
  })

  // =========================================================================
  // Conversation resume
  // =========================================================================
  describe('conversation resume', function () {
    it('should NOT call startConversationStreaming when conversationId is provided', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
      })

      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert.strictEqual(
        client.startConversationStreaming.callCount, 0,
        'startConversationStreaming should NOT be called when resuming'
      )
      conn.end()
    })

    it('should return the provided conversationId from the getter', function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
      })

      assert.strictEqual(conn.conversationId, 'existing-conv-123')
      conn.end()
    })

    it('should pass conversationId to sendActivityStreaming on postActivity', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
      })

      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      const activity = makeActivity()
      const id = await firstValueFrom(conn.postActivity(activity))

      assert(typeof id === 'string' && id.length > 0, 'postActivity should return an activity ID')
      assert(client.sendActivityStreaming.calledOnce, 'sendActivityStreaming should be called')

      const [, convIdArg] = client.sendActivityStreaming.firstCall.args
      assert.strictEqual(convIdArg, 'existing-conv-123', 'conversationId should be passed to sendActivityStreaming')
      conn.end()
    })

    it('should transition connectionStatus$ to 2 even when resuming', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
      })

      const statuses: number[] = []
      conn.connectionStatus$.subscribe((s) => statuses.push(s))
      conn.activity$.subscribe({})

      await new Promise((resolve) => setTimeout(resolve, 50))

      assert(statuses.includes(2), 'connectionStatus$ should reach 2 when resuming')
      conn.end()
    })
  })

  // =========================================================================
  // startConversation control
  // =========================================================================
  describe('startConversation setting', function () {
    it('startConversation: false should skip greeting even without conversationId', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        startConversation: false,
      })

      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert.strictEqual(
        client.startConversationStreaming.callCount, 0,
        'startConversationStreaming should NOT be called when startConversation is false'
      )
      conn.end()
    })

    it('startConversation: true with conversationId should call startConversationStreaming', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
        startConversation: true,
      })

      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert.strictEqual(
        client.startConversationStreaming.callCount, 1,
        'startConversationStreaming should be called when startConversation is explicitly true'
      )
      conn.end()
    })
  })

  // =========================================================================
  // Typing indicators
  // =========================================================================
  describe('typing indicators', function () {
    it('showTyping: true should emit typing activity before greeting stream', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, { showTyping: true })

      const activities: Partial<Activity>[] = []
      const done = new Promise<void>((resolve) => {
        conn.activity$.subscribe({
          next: (a) => activities.push(a),
          complete: () => resolve(),
        })
      })

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()
      await done

      const typingActivities = activities.filter((a) => a.type === 'typing')
      assert(typingActivities.length >= 1, 'should emit at least one typing activity')

      // Typing should come before the first message
      const firstTypingIdx = activities.findIndex((a) => a.type === 'typing')
      const firstMessageIdx = activities.findIndex((a) => a.type === 'message')
      assert(firstTypingIdx < firstMessageIdx, 'typing should be emitted before the first message')
    })

    it('showTyping: false (default) should not emit typing activities', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      const activities: Partial<Activity>[] = []
      const done = new Promise<void>((resolve) => {
        conn.activity$.subscribe({
          next: (a) => activities.push(a),
          complete: () => resolve(),
        })
      })

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()
      await done

      const typingActivities = activities.filter((a) => a.type === 'typing')
      assert.strictEqual(typingActivities.length, 0, 'should not emit typing activities by default')
    })

    it('typing should use server-provided conversation.name when available', async function () {
      const client = createMockClient(sandbox, {
        greetingActivities: [
          { type: 'message', text: 'First', conversation: { id: 'conv-1', name: 'MyBot' } },
        ],
      })

      const conn = createConnection(client, { showTyping: true })

      const activities: Partial<Activity>[] = []
      conn.activity$.subscribe({ next: (a) => activities.push(a) })
      await new Promise((resolve) => setTimeout(resolve, 50))

      // After greeting sets conversation, postActivity's notifyTyping should use the bot name
      await new Promise<void>((resolve, reject) => {
        conn.postActivity(makeActivity()).subscribe({
          complete: () => resolve(),
          error: (e) => reject(e),
        })
      })

      // Find the typing activity emitted by postActivity (after the greeting phase)
      // Greeting phase emits: typing (fallback), message. postActivity emits: user message, typing (bot name), response.
      const typingActivities = activities.filter((a) => a.type === 'typing')
      const postActivityTyping = typingActivities.find((a) => a.from?.name === 'MyBot')
      assert(postActivityTyping, 'should emit typing with server-provided conversation name')
      assert.strictEqual(postActivityTyping!.from!.id, 'conv-1')
      conn.end()
    })

    it('should forward server-sent typing activities from greeting stream', async function () {
      const client = createMockClient(sandbox, {
        greetingActivities: [
          { type: 'typing', conversation: { id: 'conv-1' } },
          { type: 'message', text: 'Hi there!', conversation: { id: 'conv-1' } },
        ],
      })

      const conn = createConnection(client)

      const activities: Partial<Activity>[] = []
      const done = new Promise<void>((resolve) => {
        conn.activity$.subscribe({
          next: (a) => activities.push(a),
          complete: () => resolve(),
        })
      })

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()
      await done

      const typingActivities = activities.filter((a) => a.type === 'typing')
      assert(typingActivities.length >= 1, 'should forward server-sent typing from greeting stream')
      // Typing should have timestamp and sequence-id like any other activity
      assert(typingActivities[0].timestamp, 'server typing should have timestamp')
      assert(
        (typingActivities[0].channelData as Record<string, unknown>)?.['webchat:sequence-id'] !== undefined,
        'server typing should have webchat:sequence-id'
      )
    })

    it('should forward server-sent typing activities from response stream', async function () {
      const client = createMockClient(sandbox, {
        responseActivities: [
          { type: 'typing', conversation: { id: 'conv-from-server' } },
          { type: 'message', text: 'Response', conversation: { id: 'conv-from-server' } },
        ],
      })

      const conn = createConnection(client)

      const activities: Partial<Activity>[] = []
      conn.activity$.subscribe({ next: (a) => activities.push(a) })
      await new Promise((resolve) => setTimeout(resolve, 50))

      // Send a message to trigger the response stream
      await new Promise<void>((resolve, reject) => {
        conn.postActivity(makeActivity()).subscribe({
          complete: () => resolve(),
          error: (e) => reject(e),
        })
      })

      // Filter to typing activities that came from the response stream (after the greeting)
      // The response stream yields: typing, then message
      const responseTyping = activities.filter(
        (a) => a.type === 'typing' && a.conversation?.id === 'conv-from-server'
      )
      assert(responseTyping.length >= 1, 'should forward server-sent typing from response stream')
      conn.end()
    })

    it('typing should fall back to default Agent identity before first response', async function () {
      const client = createMockClient(sandbox, {
        greetingActivities: [],  // No greeting activities — conversation never set
      })
      const conn = createConnection(client, { showTyping: true })

      const activities: Partial<Activity>[] = []
      const done = new Promise<void>((resolve) => {
        conn.activity$.subscribe({
          next: (a) => activities.push(a),
          complete: () => resolve(),
        })
      })

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()
      await done

      // Even with no greeting activities, showTyping emits a typing indicator using fallback
      const typingActivities = activities.filter((a) => a.type === 'typing')
      assert(typingActivities.length >= 1, 'should emit a typing activity even with no greeting')
      assert.deepStrictEqual(typingActivities[0].from, { id: 'agent', name: 'Agent' }, 'should use fallback Agent identity')
    })
  })

  // =========================================================================
  // History replay (getHistoryFromExternalStorage)
  // =========================================================================
  describe('history replay (getHistoryFromExternalStorage)', function () {
    it('should emit stored activities before live stream', async function () {
      const storedActivities: Partial<Activity>[] = [
        { type: 'message', text: 'stored-1', channelData: { 'webchat:sequence-id': 0 } },
        { type: 'message', text: 'stored-2', channelData: { 'webchat:sequence-id': 1 } },
      ]

      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
        getHistoryFromExternalStorage: sandbox.stub().resolves(storedActivities),
      })

      const activities: Partial<Activity>[] = []
      conn.activity$.subscribe({ next: (a) => activities.push(a) })
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert(activities.length >= 2, 'should emit stored activities')
      // Stored activities are emitted raw (not through notifyActivity)
      assert.strictEqual(activities[0].text, 'stored-1')
      assert.strictEqual(activities[1].text, 'stored-2')
      conn.end()
    })

    it('should continue sequence numbering from stored history', async function () {
      const storedActivities: Partial<Activity>[] = [
        { type: 'message', text: 'stored-1', channelData: { 'webchat:sequence-id': 0 } },
        { type: 'message', text: 'stored-2', channelData: { 'webchat:sequence-id': 5 } },
      ]

      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
        startConversation: true,
        getHistoryFromExternalStorage: sandbox.stub().resolves(storedActivities),
      })

      const activities: Partial<Activity>[] = []
      const done = new Promise<void>((resolve) => {
        conn.activity$.subscribe({
          next: (a) => activities.push(a),
          complete: () => resolve(),
        })
      })

      await new Promise((resolve) => setTimeout(resolve, 50))
      conn.end()
      await done

      // After stored history (max seq=5), live activities should start at 6
      const liveActivities = activities.filter(
        (a) => a.text !== 'stored-1' && a.text !== 'stored-2'
      )
      assert(liveActivities.length >= 1, 'should have live activities after history')
      const firstLiveSeq = (liveActivities[0].channelData as Record<string, unknown>)?.['webchat:sequence-id']
      assert.strictEqual(firstLiveSeq, 6, 'live activities should continue sequence from stored history')
    })

    it('should transition connectionStatus$ to 2 after history replay', async function () {
      const storedActivities: Partial<Activity>[] = [
        { type: 'message', text: 'stored', channelData: { 'webchat:sequence-id': 0 } },
      ]

      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
        getHistoryFromExternalStorage: sandbox.stub().resolves(storedActivities),
      })

      const statuses: number[] = []
      conn.connectionStatus$.subscribe((s) => statuses.push(s))
      conn.activity$.subscribe({})

      await new Promise((resolve) => setTimeout(resolve, 50))

      assert(statuses.includes(2), 'connectionStatus$ should reach 2 after history replay')
      conn.end()
    })

    it('should gracefully handle storage errors and continue', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: 'existing-conv-123',
        getHistoryFromExternalStorage: sandbox.stub().rejects(new Error('Storage unavailable')),
      })

      const statuses: number[] = []
      conn.connectionStatus$.subscribe((s) => statuses.push(s))
      conn.activity$.subscribe({})

      await new Promise((resolve) => setTimeout(resolve, 50))

      // Should still transition to connected despite storage error
      assert(statuses.includes(2), 'connectionStatus$ should reach 2 even after storage error')
      conn.end()
    })

    it('should not call getHistoryFromExternalStorage without conversationId', async function () {
      const getHistory = sandbox.stub().resolves([])
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        getHistoryFromExternalStorage: getHistory,
      })

      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert.strictEqual(getHistory.callCount, 0, 'should not call getHistoryFromExternalStorage without conversationId')
      conn.end()
    })
  })

  // =========================================================================
  // Error handling
  // =========================================================================
  describe('error handling', function () {
    it('should throw when postActivity is called after end()', function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      conn.activity$.subscribe({})
      conn.end()

      assert.throws(
        () => conn.postActivity(makeActivity()),
        /Connection has been ended/,
        'postActivity after end() should throw'
      )
    })

    it('should throw when postActivity is called with null activity', function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      conn.activity$.subscribe({})

      assert.throws(
        () => conn.postActivity(null as unknown as Activity),
        /Activity cannot be null/,
        'postActivity with null should throw'
      )
      conn.end()
    })
  })

  // =========================================================================
  // Edge cases
  // =========================================================================
  describe('edge cases', function () {
    it('multiple subscriptions should not trigger duplicate startConversation calls', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client)

      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert.strictEqual(
        client.startConversationStreaming.callCount, 1,
        'startConversationStreaming should only be called once despite multiple subscriptions'
      )
      conn.end()
    })

    it('conversationId captured from sendActivityStreaming response when not set upfront', async function () {
      const client = createMockClient(sandbox, {
        greetingActivities: [
          { type: 'message', text: 'Hello' },
        ],
        responseActivities: [
          { type: 'message', text: 'Response', conversation: { id: 'captured-conv-id' } },
        ],
      })

      const conn = createConnection(client)
      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert.strictEqual(conn.conversationId, undefined, 'conversationId should be undefined before sendActivity response')

      const activity = makeActivity()
      await new Promise<void>((resolve, reject) => {
        conn.postActivity(activity).subscribe({
          complete: () => resolve(),
          error: (e) => reject(e),
        })
      })

      assert.strictEqual(conn.conversationId, 'captured-conv-id', 'conversationId should be captured from sendActivity response')
      conn.end()
    })

    it('should normalize conversationId by trimming whitespace', function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: '  conv-123  ',
      })

      assert.strictEqual(conn.conversationId, 'conv-123', 'conversationId should be trimmed')
      conn.end()
    })

    it('should treat empty string conversationId as undefined', async function () {
      const client = createMockClient(sandbox)
      const conn = createConnection(client, {
        conversationId: '   ',
      })

      assert.strictEqual(conn.conversationId, undefined, 'whitespace-only conversationId should be treated as undefined')

      // Should behave as a new conversation (call startConversationStreaming)
      conn.activity$.subscribe({})
      await new Promise((resolve) => setTimeout(resolve, 50))

      assert.strictEqual(
        client.startConversationStreaming.callCount, 1,
        'should start new conversation when conversationId is empty'
      )
      conn.end()
    })
  })
})
