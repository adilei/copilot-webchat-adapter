import { Activity, Attachment } from '@microsoft/agents-activity'
import { Observable, BehaviorSubject, Subscriber } from 'rxjs'
import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client'
import { CreateConnectionOptions, WebChatConnection } from './types.js'

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
  const { showTyping = false, getHistoryFromExternalStorage } = options

  const normalizedConversationId =
    options.conversationId && options.conversationId.trim() !== ''
      ? options.conversationId.trim()
      : undefined
  const shouldStart = options.startConversation ?? !normalizedConversationId

  let sequence = 0
  let activitySubscriber: Subscriber<Partial<Activity>> | undefined
  let activeConversationId: string | undefined = normalizedConversationId
  let ended = false
  let started = false


  const connectionStatus$ = new BehaviorSubject(0)

  const activity$ = new Observable<Partial<Activity>>((subscriber) => {
    activitySubscriber = subscriber

    if (connectionStatus$.value < 2) {
      connectionStatus$.next(2)
    }

    ;(async () => {
      // Replay stored activities on resume
      if (getHistoryFromExternalStorage && normalizedConversationId) {
        try {
          const stored = await getHistoryFromExternalStorage(normalizedConversationId)
          for (const activity of stored) {
            if (!ended && activitySubscriber) {
              // Emit directly — stored activities are already enriched with
              // timestamps and sequence IDs from when the consumer saved them
              activitySubscriber.next(activity)
            }
          }
          // Continue sequence numbering from where stored history left off
          const lastSeq = stored.reduce((max, a) => {
            const seq = (a.channelData as Record<string, unknown>)?.['webchat:sequence-id']
            return typeof seq === 'number' && seq > max ? seq : max
          }, -1)
          if (lastSeq >= sequence) {
            sequence = lastSeq + 1
          }
        } catch {
          // History unavailable — continue without it
        }
      }

      // Stream greeting activities for new conversations
      if (shouldStart && !started) {
        started = true
        try {
          emitTyping()
          for await (const activity of client.startConversationStreaming()) {
            if (activity.conversation?.id) {
              activeConversationId = activity.conversation.id
            }
            delete activity.replyToId
            emitActivity(activity)
          }
        } catch (error) {
          subscriber.error(error)
        }
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

            const outgoing = Activity.fromObject({
              ...activity,
              id,
              conversation: { id: activeConversationId || '' },
              attachments: await processAttachments(activity),
            })

            emitActivity(outgoing)
            emitTyping()

            // Notify WebChat immediately that the message was sent
            subscriber.next(id)

            // Stream the agent's response
            for await (const response of client.sendActivityStreaming(outgoing, activeConversationId)) {
              if (!activeConversationId && response.conversation?.id) {
                activeConversationId = response.conversation.id
              }
              emitActivity(response)
            }

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

/**
 * Processes activity attachments, converting blob URLs to data URLs.
 */
async function processAttachments(activity: Activity): Promise<Attachment[]> {
  if (activity.type !== 'message' || !activity.attachments?.length) {
    return activity.attachments || []
  }

  const attachments: Attachment[] = []
  for (const attachment of activity.attachments) {
    attachments.push(await processBlobAttachment(attachment))
  }
  return attachments
}

/**
 * Converts a blob: content URL to a data: URL so it can be sent to the server.
 */
async function processBlobAttachment(attachment: Attachment): Promise<Attachment> {
  if (!attachment.contentUrl?.startsWith('blob:')) {
    return attachment
  }

  try {
    const response = await fetch(attachment.contentUrl)
    if (!response.ok) {
      throw new Error(`Failed to fetch blob URL: ${response.status} ${response.statusText}`)
    }
    const blob = await response.blob()
    const arrayBuffer = await blob.arrayBuffer()
    const base64 = arrayBufferToBase64(arrayBuffer)
    return { ...attachment, contentUrl: `data:${blob.type};base64,${base64}` }
  } catch {
    return attachment
  }
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  // Node.js
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const B = (globalThis as any).Buffer
  if (typeof B === 'function') {
    return B.from(buffer).toString('base64')
  }
  // Browser
  let binary = ''
  for (const byte of new Uint8Array(buffer)) {
    binary += String.fromCharCode(byte)
  }
  return btoa(binary)
}
