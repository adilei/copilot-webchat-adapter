import { v4 as uuid } from 'uuid'
import { Activity } from '@microsoft/agents-activity'
import { Observable, BehaviorSubject, Subscriber } from 'rxjs'
import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client'
import { CreateConnectionOptions, WebChatConnection } from './types'

/**
 * Creates a DirectLine-compatible connection for integrating CopilotStudioClient with WebChat.
 *
 * Unlike `CopilotStudioWebChat.createConnection()` (unreleased), this adapter supports:
 * - Passing a `conversationId` to resume an existing conversation
 * - Configuring whether `startConversationAsync()` is called on initialization
 *
 * Uses only CopilotStudioClient's published public API:
 * - `startConversationAsync()` to begin a new conversation
 * - `askQuestionAsync(text, conversationId)` to send messages with explicit conversationId
 *
 * @example New conversation (default behavior)
 * ```ts
 * const directLine = createConnection(client)
 * ```
 *
 * @example Resume existing conversation
 * ```ts
 * const directLine = createConnection(client, {
 *   conversationId: savedConversationId,
 *   showTyping: true,
 * })
 * ```
 *
 * @example Resume but also call startConversation (get greeting)
 * ```ts
 * const directLine = createConnection(client, {
 *   conversationId: savedConversationId,
 *   startConversation: true,
 * })
 * ```
 */
export function createConnection(
  client: CopilotStudioClient,
  options: CreateConnectionOptions = {}
): WebChatConnection {
  const {
    conversationId,
    showTyping = false,
  } = options

  // Default: start conversation if no conversationId; skip if resuming
  const shouldStart = options.startConversation ?? !conversationId

  let sequence = 0
  let activitySubscriber: Subscriber<Partial<Activity>> | undefined
  // Track the active conversationId (set from options or from startConversation response)
  let activeConversationId: string | undefined = conversationId
  let ended = false
  let started = false

  const connectionStatus$ = new BehaviorSubject(0)

  const activity$ = new Observable<Partial<Activity>>((subscriber) => {
    activitySubscriber = subscriber

    // Mark as connected when WebChat subscribes (matches SDK behavior)
    if (connectionStatus$.value < 2) {
      connectionStatus$.next(2)
    }

    // Guard against duplicate startConversation calls on re-subscription
    if (!shouldStart || started) {
      return
    }
    started = true

    // Start a new conversation (or re-greet an existing one)
    ;(async () => {
      try {
        sequence = 0
        emitTyping()
        const greeting = await client.startConversationAsync()
        // Capture conversationId from the server response (if not already set)
        if (!activeConversationId && greeting.conversation?.id) {
          activeConversationId = greeting.conversation.id
        }
        emitActivity(greeting)
      } catch (error) {
        subscriber.error(error)
      }
    })()

    // Cleanup on unsubscribe
    return () => {
      if (activitySubscriber === subscriber) {
        activitySubscriber = undefined
      }
    }
  })

  function emitActivity(activity: Partial<Activity>): void {
    if (ended || !activitySubscriber) return
    activitySubscriber.next({
      ...activity,
      timestamp: new Date().toISOString(),
      channelData: {
        ...activity.channelData,
        'webchat:sequence-id': sequence++,
      },
    })
  }

  function emitTyping(): void {
    if (!showTyping) return
    const from = activeConversationId
      ? { id: activeConversationId, name: 'Agent' }
      : { id: 'agent', name: 'Agent' }
    emitActivity({ type: 'typing', from })
  }

  return {
    connectionStatus$,
    activity$,

    postActivity(activity: Activity): Observable<string> {
      if (!activity) {
        throw new Error('Activity cannot be null.')
      }
      if (ended) {
        throw new Error('Connection has been ended.')
      }
      if (!activitySubscriber) {
        throw new Error('Activity subscriber is not initialized.')
      }

      return new Observable<string>((subscriber) => {
        ;(async () => {
          try {
            const id = uuid()

            // Echo the user's activity back to the transcript
            emitActivity({ ...activity, id })
            emitTyping()

            // Send to Copilot Studio using askQuestionAsync with explicit conversationId
            const text = activity.text || ''
            const responses = await client.askQuestionAsync(
              text,
              activeConversationId
            )

            // Capture conversationId from response if we don't have one yet
            for (const response of responses) {
              if (!activeConversationId && response.conversation?.id) {
                activeConversationId = response.conversation.id
              }
              emitActivity(response)
            }

            subscriber.next(id)
            subscriber.complete()
          } catch (error) {
            subscriber.error(error)
          }
        })()
      })
    },

    end(): void {
      ended = true
      connectionStatus$.complete()
      if (activitySubscriber) {
        activitySubscriber.complete()
        activitySubscriber = undefined
      }
    },
  }
}
