/**
 * Browser-compatible DirectLine adapter for CopilotStudioClient + WebChat.
 *
 * Uses the streaming async generators from the published browser.mjs build:
 *   - client.startConversationStreaming()
 *   - client.sendActivityStreaming(activity, conversationId)
 *
 * Extends the official CopilotStudioWebChat.createConnection() pattern with:
 *   - conversationId resume support
 *   - configurable startConversation behavior
 *   - real-time streaming (yields activities as SSE chunks arrive)
 */

import { Observable, BehaviorSubject } from 'rxjs'

export function createConnection(client, options = {}) {
  const { conversationId, showTyping = false } = options
  const shouldStart = options.startConversation ?? !conversationId

  let sequence = 0
  let activitySubscriber
  let activeConversationId = conversationId
  let ended = false
  let started = false

  const connectionStatus$ = new BehaviorSubject(0)

  const activity$ = new Observable((subscriber) => {
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
        for await (const activity of client.startConversationStreaming()) {
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

  function emitActivity(activity) {
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

  function emitTyping() {
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

    postActivity(activity) {
      if (!activity) throw new Error('Activity cannot be null.')
      if (ended) throw new Error('Connection has been ended.')
      if (!activitySubscriber) throw new Error('Activity subscriber is not initialized.')

      return new Observable((subscriber) => {
        ;(async () => {
          try {
            const id = crypto.randomUUID()

            emitActivity({ ...activity, id })
            emitTyping()

            const outgoing = {
              ...activity,
              conversation: { id: activeConversationId || '' },
            }
            for await (const response of client.sendActivityStreaming(outgoing, activeConversationId)) {
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

    end() {
      ended = true
      connectionStatus$.complete()
      if (activitySubscriber) {
        activitySubscriber.complete()
        activitySubscriber = undefined
      }
    },
  }
}
