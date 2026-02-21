import { v4 as uuid } from 'uuid'

import { Activity, Attachment, ConversationAccount } from '@microsoft/agents-activity'
import { Observable, BehaviorSubject, type Subscriber } from 'rxjs'
import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client'
import { CreateConnectionOptions, WebChatConnection } from './types.js'

/**
 * Creates a wrapper that invokes `fn` at most once.
 * On the first call the wrapper invokes `fn(value)` and returns whatever `fn` returns.
 * Subsequent calls do nothing and return `undefined`.
 */
function once<T = void> (fn: (value: T) => Promise<void>): (value: T) => Promise<void> | void {
  let called = false

  return value => {
    if (!called) {
      called = true

      return fn(value)
    }
  }
}

/**
 * Creates an RxJS Observable that wraps an asynchronous function execution.
 */
function createObservable<T> (fn: (subscriber: Subscriber<T>) => void): Observable<T> {
  return new Observable<T>((subscriber: Subscriber<T>) => {
    Promise.resolve(fn(subscriber)).catch((error) => subscriber.error(error))
  })
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
 * - `getHistoryFromExternalStorage` to replay stored activities on resume
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
  let conversation: ConversationAccount | undefined = normalizedConversationId
    ? ({ id: normalizedConversationId } as ConversationAccount)
    : undefined
  let activeConversationId: string | undefined = normalizedConversationId
  let ended = false
  let started = false

  const connectionStatus$ = new BehaviorSubject(0)
  const activity$ = createObservable<Partial<Activity>>(async (subscriber) => {
    activitySubscriber = subscriber

    const handleAcknowledgementOnce = once(async (): Promise<void> => {
      connectionStatus$.next(2)
      await Promise.resolve() // Webchat requires an extra tick to process the connection status change
    })

    // Replay stored activities on resume
    if (getHistoryFromExternalStorage && normalizedConversationId) {
      try {
        const stored = await getHistoryFromExternalStorage(normalizedConversationId)
        for (const activity of stored) {
          if (activitySubscriber) {
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
      await handleAcknowledgementOnce()
    }

    // When resuming (shouldStart === false), transition straight to connected
    if (!shouldStart || started) {
      await handleAcknowledgementOnce()
      return
    }
    started = true

    notifyTyping()

    for await (const activity of (client as any).startConversationStreaming()) {
      delete activity.replyToId
      if (!conversation && activity.conversation) {
        conversation = activity.conversation
      }
      if (activity.conversation?.id) {
        activeConversationId = activity.conversation.id
      }
      await handleAcknowledgementOnce()
      notifyActivity(activity)
    }
    // If no activities received from bot, we should still acknowledge.
    await handleAcknowledgementOnce()
  })

  const notifyActivity = (activity: Partial<Activity>) => {
    const newActivity = {
      ...activity,
      timestamp: new Date().toISOString(),
      channelData: {
        ...activity.channelData,
        'webchat:sequence-id': sequence,
      },
    }
    sequence++
    activitySubscriber?.next(newActivity)
  }

  const notifyTyping = () => {
    if (!showTyping) {
      return
    }

    const from = conversation
      ? { id: conversation.id, name: conversation.name }
      : { id: 'agent', name: 'Agent' }
    notifyActivity({ type: 'typing', from })
  }

  return {
    connectionStatus$,
    activity$,

    get conversationId () {
      return activeConversationId
    },

    postActivity (activity: Activity) {
      if (!activity) {
        throw new Error('Activity cannot be null.')
      }

      if (ended) {
        throw new Error('Connection has been ended.')
      }

      if (!activitySubscriber) {
        throw new Error('Activity subscriber is not initialized.')
      }

      return createObservable<string>(async (subscriber) => {
        try {
          const newActivity = Activity.fromObject({
            ...activity,
            id: uuid(),
            attachments: await processAttachments(activity)
          })

          notifyActivity(newActivity)
          notifyTyping()

          // Notify WebChat immediately that the message was sent
          subscriber.next(newActivity.id!)

          // Stream the agent's response, passing activeConversationId for URL routing
          for await (const responseActivity of (client as any).sendActivityStreaming(newActivity, activeConversationId)) {
            if (!activeConversationId && responseActivity.conversation?.id) {
              activeConversationId = responseActivity.conversation.id
            }
            notifyActivity(responseActivity)
          }

          subscriber.complete()
        } catch (error) {
          console.warn('Error sending Activity to Copilot Studio:', error)
          subscriber.error(error)
        }
      })
    },

    end () {
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
 * Processes activity attachments.
 */
async function processAttachments (activity: Activity): Promise<Attachment[]> {
  if (activity.type !== 'message' || !activity.attachments?.length) {
    return activity.attachments || []
  }

  const attachments: Attachment[] = []
  for (const attachment of activity.attachments) {
    const processed = await processBlobAttachment(attachment)
    attachments.push(processed)
  }

  return attachments
}

/**
 * Converts a blob: content URL to a data: URL so it can be sent to the server.
 */
async function processBlobAttachment (attachment: Attachment): Promise<Attachment> {
  let newContentUrl = attachment.contentUrl
  if (!newContentUrl?.startsWith('blob:')) {
    return attachment
  }

  try {
    const response = await fetch(newContentUrl)
    if (!response.ok) {
      throw new Error(`Failed to fetch blob URL: ${response.status} ${response.statusText}`)
    }

    const blob = await response.blob()
    const arrayBuffer = await blob.arrayBuffer()
    const base64 = arrayBufferToBase64(arrayBuffer)
    newContentUrl = `data:${blob.type};base64,${base64}`
  } catch (error) {
    newContentUrl = attachment.contentUrl
    console.warn('Error processing blob attachment:', newContentUrl, error)
  }

  return { ...attachment, contentUrl: newContentUrl }
}

/**
 * Converts an ArrayBuffer to a base64 string.
 */
function arrayBufferToBase64 (buffer: ArrayBuffer): string {
  // Node.js environment
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const BufferClass = typeof (globalThis as any).Buffer === 'function' ? (globalThis as any).Buffer : undefined
  if (BufferClass && typeof BufferClass.from === 'function') {
    return BufferClass.from(buffer).toString('base64')
  }

  // Browser environment
  let binary = ''
  for (const byte of new Uint8Array(buffer)) {
    binary += String.fromCharCode(byte)
  }
  return btoa(binary)
}
