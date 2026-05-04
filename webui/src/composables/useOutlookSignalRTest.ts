import * as signalR from '@microsoft/signalr'
import { computed, onBeforeUnmount, ref } from 'vue'
import type {
  OutlookSignalRTestCommand,
  OutlookSignalRTestConnectionEvent,
  OutlookSignalRTestMessage,
  OutlookSignalRTestResult,
  SignalRState,
} from '../models/outlook'

type TestEvent = {
  id: string
  kind: 'connection' | 'command' | 'message' | 'result' | 'system'
  text: string
  timestamp: string
}

export function useOutlookSignalRTest() {
  const state = ref<SignalRState>('disconnected')
  const webConnectionId = ref('')
  const addinConnections = ref<OutlookSignalRTestConnectionEvent[]>([])
  const events = ref<TestEvent[]>([])
  const commandType = ref('ping')
  const commandPayload = ref('{"message":"hello from Web UI"}')
  const sending = ref(false)
  let connection: signalR.HubConnection | null = null

  const connectedAddinCount = computed(() => addinConnections.value.length)

  function addEvent(kind: TestEvent['kind'], text: string, timestamp = new Date().toISOString()) {
    events.value = [
      {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        kind,
        text,
        timestamp,
      },
      ...events.value,
    ].slice(0, 100)
  }

  async function connect() {
    if (connection && state.value !== 'disconnected') return

    connection = new signalR.HubConnectionBuilder()
      .withUrl('/hub/outlook-test')
      .withAutomaticReconnect()
      .build()

    connection.onreconnecting(() => {
      state.value = 'reconnecting'
      addEvent('system', 'Web UI test client reconnecting')
    })
    connection.onreconnected((connectionId?: string) => {
      state.value = 'connected'
      webConnectionId.value = connectionId ?? ''
      addEvent('system', `Web UI test client reconnected: ${connectionId ?? '-'}`)
    })
    connection.onclose(() => {
      state.value = 'disconnected'
      webConnectionId.value = ''
      addEvent('system', 'Web UI test client disconnected')
    })

    connection.on('OutlookSignalRTestAddinConnected', (info: OutlookSignalRTestConnectionEvent) => {
      addinConnections.value = [
        info,
        ...addinConnections.value.filter((item) => item.connectionId !== info.connectionId),
      ]
      addEvent('connection', `AddIn connected: ${info.clientName || 'Outlook AddIn'} ${info.workstation || ''}`, info.timestamp)
    })
    connection.on('OutlookSignalRTestAddinDisconnected', (info: OutlookSignalRTestConnectionEvent) => {
      addinConnections.value = addinConnections.value.filter((item) => item.connectionId !== info.connectionId)
      addEvent('connection', `AddIn disconnected: ${info.connectionId}`, info.timestamp)
    })
    connection.on('OutlookSignalRTestCommandDispatched', (command: OutlookSignalRTestCommand) => {
      addEvent('command', `Command dispatched: ${command.type} (${command.id})`, command.createdAt)
    })
    connection.on('OutlookSignalRTestMessage', (message: OutlookSignalRTestMessage) => {
      addEvent('message', `[${message.level}] ${message.source}: ${message.text}`, message.timestamp)
    })
    connection.on('OutlookSignalRTestResult', (result: OutlookSignalRTestResult) => {
      addEvent('result', `Result ${result.success ? 'success' : 'failed'} for ${result.commandId}: ${result.message}`, result.timestamp)
    })

    try {
      await connection.start()
      state.value = 'connected'
      webConnectionId.value = connection.connectionId ?? ''
      addEvent('system', `Web UI test client connected: ${connection.connectionId ?? '-'}`)
    } catch (error) {
      state.value = 'disconnected'
      addEvent('system', error instanceof Error ? error.message : 'SignalR test connection failed')
    }
  }

  async function disconnect() {
    await connection?.stop()
    connection = null
    state.value = 'disconnected'
    webConnectionId.value = ''
  }

  async function sendCommand() {
    if (!connection || state.value !== 'connected') await connect()
    if (!connection) return

    sending.value = true
    try {
      await connection.invoke('SendOutlookSignalRTestCommand', {
        id: crypto.randomUUID(),
        type: commandType.value.trim() || 'ping',
        payload: commandPayload.value,
        createdAt: new Date().toISOString(),
      })
    } finally {
      sending.value = false
    }
  }

  onBeforeUnmount(() => {
    void disconnect()
  })

  return {
    addinConnections,
    commandPayload,
    commandType,
    connect,
    connectedAddinCount,
    disconnect,
    events,
    sendCommand,
    sending,
    state,
    webConnectionId,
  }
}
