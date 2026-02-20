/**
 * Creates an ActivityStore backed by localStorage.
 *
 * Each conversation's activities are stored as a JSON array under
 * `${prefix}:${conversationId}`. Silently degrades on quota or parse errors.
 *
 * @param {string} [prefix='webchat-activities'] - localStorage key prefix
 * @returns {{ getActivities: (id: string) => Promise<Array>, saveActivity: (id: string, activity: Object) => Promise<void>, clear: (id: string) => Promise<void> }}
 */
export function createLocalStorageStore(prefix = 'webchat-activities') {
  function key(conversationId) {
    return `${prefix}:${conversationId}`
  }

  function readActivities(conversationId) {
    try {
      const raw = localStorage.getItem(key(conversationId))
      if (!raw) return []
      const parsed = JSON.parse(raw)
      return Array.isArray(parsed) ? parsed : []
    } catch {
      return []
    }
  }

  function writeActivities(conversationId, activities) {
    try {
      localStorage.setItem(key(conversationId), JSON.stringify(activities))
    } catch {
      // Quota exceeded or other storage error — silently ignore
    }
  }

  return {
    async getActivities(conversationId) {
      return readActivities(conversationId)
    },

    async saveActivity(conversationId, activity) {
      const activities = readActivities(conversationId)
      activities.push(activity)
      writeActivities(conversationId, activities)
    },

    async clear(conversationId) {
      try {
        localStorage.removeItem(key(conversationId))
      } catch {
        // Silently ignore
      }
    },
  }
}
