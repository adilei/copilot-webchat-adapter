import { Activity } from '@microsoft/agents-activity'
import { BehaviorSubject, Observable } from 'rxjs'

/**
 * DirectLine-compatible connection interface for WebChat integration.
 * Pass this object as the `directLine` prop to WebChat's Composer or renderWebChat.
 */
export interface WebChatConnection {
  connectionStatus$: BehaviorSubject<number>
  activity$: Observable<Partial<Activity>>
  postActivity(activity: Activity): Observable<string>
  end(): void
}

/**
 * Options for creating a WebChat connection with conversation support.
 */
export interface CreateConnectionOptions {
  /**
   * Existing conversation ID to resume.
   * When provided, the adapter sends all activities with this conversationId
   * via the client's public `sendActivity(activity, conversationId)` method.
   */
  conversationId?: string

  /**
   * Whether to call `startConversationAsync()` when the connection initializes.
   *
   * - `true`: Always call startConversationAsync (default when no conversationId)
   * - `false`: Skip the start call entirely (default when conversationId is provided)
   *
   * When not specified, defaults to `!conversationId` (skip start when resuming).
   */
  startConversation?: boolean

  /**
   * Show typing indicators while the agent processes a response.
   * @default false
   */
  showTyping?: boolean
}
