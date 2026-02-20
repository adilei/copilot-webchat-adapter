import { Activity } from '@microsoft/agents-activity'

/**
 * Abstraction for storing and retrieving conversation activities.
 * Consumers provide an implementation to persist activities and
 * pass `store.getActivities` as the adapter's `getHistoryFromExternalStorage`.
 */
export interface ActivityStore {
  getActivities(conversationId: string): Promise<Partial<Activity>[]>
  saveActivity(conversationId: string, activity: Partial<Activity>): Promise<void>
  clear(conversationId: string): Promise<void>
}

/**
 * Creates an ActivityStore backed by localStorage.
 *
 * Each conversation's activities are stored as a JSON array under
 * `${prefix}:${conversationId}`. Silently degrades on quota or parse errors.
 *
 * @param prefix - localStorage key prefix (default: `'webchat-activities'`)
 */
export function createLocalStorageStore(prefix = 'webchat-activities'): ActivityStore {
  function key(conversationId: string): string {
    return `${prefix}:${conversationId}`
  }

  function readActivities(conversationId: string): Partial<Activity>[] {
    try {
      const raw = localStorage.getItem(key(conversationId))
      if (!raw) return []
      const parsed = JSON.parse(raw)
      return Array.isArray(parsed) ? parsed : []
    } catch {
      return []
    }
  }

  function writeActivities(conversationId: string, activities: Partial<Activity>[]): void {
    try {
      localStorage.setItem(key(conversationId), JSON.stringify(activities))
    } catch {
      // Quota exceeded or other storage error — silently ignore
    }
  }

  return {
    async getActivities(conversationId: string): Promise<Partial<Activity>[]> {
      return readActivities(conversationId)
    },

    async saveActivity(conversationId: string, activity: Partial<Activity>): Promise<void> {
      const activities = readActivities(conversationId)
      activities.push(activity)
      writeActivities(conversationId, activities)
    },

    async clear(conversationId: string): Promise<void> {
      try {
        localStorage.removeItem(key(conversationId))
      } catch {
        // Silently ignore
      }
    },
  }
}
