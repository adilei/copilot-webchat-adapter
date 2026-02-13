import { Activity } from '@microsoft/agents-activity'
import { Observable, BehaviorSubject, Subscriber } from 'rxjs'
import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client'
import { CreateConnectionOptions, WebChatConnection } from './types.js'

// The streaming generators exist in the published browser ESM build (v1.2.3+)
// but are not declared in the .d.ts files. Cast to access them.
interface StreamingClient {
  startConversationStreaming(): AsyncIterable<Partial<Activity>>
  sendActivityStreaming(activity: Partial<Activity>, conversationId?: string): AsyncIterable<Partial<Activity>>
}

/**
 * Creates a DirectLine-compatible connection for integrating CopilotStudioClient with WebChat.
 *
 * Uses the SDK's streaming async generators for real-time activity delivery:
 * - `startConversationStreaming()` to stream greeting activities
 * - `sendActivityStreaming(activity, conversationId)` to stream response activities
 *
 * Extends the official CopilotStudioWebChat.createConnection() pattern with:
 * - `conversationId` to resume an existing conversation
 * - `startConversation` to control whether the greeting is triggered
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
 */
export function createConnection(
  client: CopilotStudioClient,
  options: CreateConnectionOptions = {}
): WebChatConnection {
  const {
    conversationId,
    showTyping = false,
  } = options

  const shouldStart = options.startConversation ?? !conversationId
  const streaming = client as unknown as StreamingClient

  let sequence = 0
  let activitySubscriber: Subscriber<Partial<Activity>> | undefined
  let activeConversationId: string | undefined = conversationId
  let ended = false
  let started = false

  const connectionStatus$ = new BehaviorSubject(0)

  const activity$ = new Observable<Partial<Activity>>((subscriber) => {
    activitySubscriber = subscriber

    if (connectionStatus$.value < 2) {
      connectionStatus$.next(2)
    }

    if (!shouldStart || started) {
      return
    }
    started = true

    ;(async () => {
      try {
        sequence = 0
        emitTyping()
        for await (const activity of streaming.startConversationStreaming()) {
          if (!activeConversationId && activity.conversation?.id) {
            activeConversationId = activity.conversation.id
          }
          emitActivity(activity)
        }
      } catch (error) {
        subscriber.error(error)
      }
    })()

    return () => {
      if (activitySubscriber === subscriber) {
        activitySubscriber = undefined
      }
    }
  })

  function emitActivity(activity: Partial<Activity>): void {
    if (ended || !activitySubscriber || !activity) return
    if (!activity.type) return
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

    get conversationId() {
      return activeConversationId
    },

    postActivity(activity: Activity): Observable<string> {
      if (!activity) throw new Error('Activity cannot be null.')
      if (ended) throw new Error('Connection has been ended.')
      if (!activitySubscriber) throw new Error('Activity subscriber is not initialized.')

      return new Observable<string>((subscriber) => {
        ;(async () => {
          try {
            const id = crypto.randomUUID()

            emitActivity({ ...activity, id })
            emitTyping()

            const outgoing = {
              ...activity,
              conversation: { id: activeConversationId || '' },
            }
            for await (const response of streaming.sendActivityStreaming(outgoing, activeConversationId)) {
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
