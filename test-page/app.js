/**
 * Test page entry point.
 * Populates agent dropdown, handles Connect flow, renders WebChat.
 */

import { CopilotStudioClient, ConnectionSettings } from '@microsoft/agents-copilotstudio-client'
import { acquireToken } from './acquireToken.js'
import { createConnection } from './createConnection.js'
import { agents, defaultAgent } from './agents.js'

// DOM elements
const agentSelect = document.getElementById('agentSelect')
const convIdInput = document.getElementById('convIdInput')
const connectBtn = document.getElementById('connectBtn')
const statusText = document.getElementById('statusText')
const convIdDisplay = document.getElementById('convIdDisplay')
const startConvCheckbox = document.getElementById('startConvCheckbox')
const webchatEl = document.getElementById('webchat')

// Populate agent dropdown
for (const [key, agent] of Object.entries(agents)) {
  const option = document.createElement('option')
  option.value = key
  option.textContent = agent.name
  agentSelect.appendChild(option)
}

agentSelect.value = defaultAgent

// Disable "Start conversation" checkbox when a conversationId is entered
convIdInput.addEventListener('input', () => {
  const hasConvId = convIdInput.value.trim().length > 0
  startConvCheckbox.disabled = hasConvId
  if (hasConvId) startConvCheckbox.checked = false
})

let currentConnection = null

connectBtn.addEventListener('click', async () => {
  const agentKey = agentSelect.value
  const agent = agents[agentKey]
  if (!agent) return

  const conversationId = convIdInput.value.trim() || undefined

  connectBtn.disabled = true
  statusText.textContent = 'Authenticating...'
  convIdDisplay.innerHTML = ''

  // Clean up previous connection
  if (currentConnection) {
    currentConnection.end()
    currentConnection = null
    webchatEl.innerHTML = ''
  }

  try {
    const settings = new ConnectionSettings()
    settings.environmentId = agent.environmentId
    settings.agentIdentifier = agent.agentIdentifier
    settings.tenantId = agent.tenantId
    settings.appClientId = agent.appClientId
    settings.cloud = 'Prod'

    statusText.textContent = 'Acquiring token...'
    const token = await acquireToken(settings)

    statusText.textContent = 'Connecting...'
    const client = new CopilotStudioClient(settings, token)

    const directLine = createConnection(client, {
      conversationId,
      startConversation: startConvCheckbox.checked,
      showTyping: true,
    })
    currentConnection = directLine

    window.WebChat.renderWebChat(
      { directLine },
      webchatEl
    )

    document.querySelector('#webchat > *')?.focus()

    statusText.textContent = 'Connected'
    connectBtn.disabled = false

    // Watch for conversationId to appear
    pollConversationId(directLine)

    // If we already have a conversationId from input, show it
    if (conversationId) {
      showConversationId(conversationId)
    }
  } catch (err) {
    console.error('Connection failed:', err)
    statusText.textContent = `Error: ${err.message}`
    connectBtn.disabled = false
  }
})

function pollConversationId(directLine) {
  const interval = setInterval(() => {
    const id = directLine.conversationId
    if (id) {
      showConversationId(id)
      clearInterval(interval)
    }
    // Stop polling if connection ended
    if (!currentConnection || currentConnection !== directLine) {
      clearInterval(interval)
    }
  }, 500)
}

function showConversationId(id) {
  convIdDisplay.innerHTML = ''

  const label = document.createElement('span')
  label.textContent = 'ConvID: '
  convIdDisplay.appendChild(label)

  const idSpan = document.createElement('span')
  idSpan.className = 'convid'
  idSpan.textContent = id
  convIdDisplay.appendChild(idSpan)

  const copyBtn = document.createElement('button')
  copyBtn.className = 'copy-btn'
  copyBtn.textContent = 'Copy'
  copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(id).then(() => {
      copyBtn.textContent = 'Copied!'
      setTimeout(() => { copyBtn.textContent = 'Copy' }, 1500)
    })
  })
  convIdDisplay.appendChild(copyBtn)
}
