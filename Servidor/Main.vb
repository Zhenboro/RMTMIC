Imports System.Text
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.Threading
Public Class Main
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Init()
        LoadMemory()
        ReadParameters(Command())
    End Sub
    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            m_IsFormMain = False
            StopRecordingFromSounddevice_Client()
            DisconnectClient()
        Catch ex As Exception
            AddToLog("Main_FormClosing@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Sub ReadParameters(ByVal parameter As String)
        Try
            If parameter <> Nothing Then
                Dim parametros() As String = parameter.Split(" ")
                For Each item As String In parametros
                    If item.ToLower Like "*/startmicstreaminglisten*" Then
                        'Comenzar streaming TCP/IP del microfono
                        StartMicStreamingListen()

                    ElseIf item.ToLower Like "*--ip*" Then
                        Dim args As String() = item.Split("-")
                        ServerIP = Integer.Parse(args(3))

                    ElseIf item.ToLower Like "*--port*" Then
                        Dim args As String() = item.Split("-")
                        ServerPort = Integer.Parse(args(3))

                    ElseIf item.ToLower Like "*/stopmicstreamlisten*" Then
                        'Detiene el streaming del microfono
                        StopMicStreamingListen()

                    ElseIf item.ToLower Like "*/stop*" Then
                        'Detiene todo y se cierra
                        End
                    End If
                Next
            End If
        Catch ex As Exception
            AddToLog("ReadParameters@Main", "Error: " & ex.Message, True)
        End Try
    End Sub

    Sub StartMicStreamingListen()
        Try
            Button1.Enabled = False
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            Dim threadmicStreaming = New Thread(Sub() Starter(ServerIP, ServerPort))
            threadmicStreaming.Start()
        Catch ex As Exception
            AddToLog("StartMicStreamingListen@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Sub StopMicStreamingListen()
        Try
            Stopper()
        Catch ex As Exception
            AddToLog("StopMicStreamingListen@Main", "Error: " & ex.Message, True)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = Nothing Or TextBox2.Text = Nothing Then
            MsgBox("Faltan datos", MsgBoxStyle.Critical, Me.Text)
        Else
            ServerIP = TextBox1.Text
            ServerPort = TextBox2.Text
            SaveMemory()
            StartMicStreamingListen()
        End If
    End Sub
    Dim IPSERVIDOR As String
    Dim PUERTO As Integer
    Dim m_Client As TCPCLIENTE
    Dim m_SoundBufferCount As Integer = 8
    Dim m_PrototolClient As New WinSound.Protocol(WinSound.ProtocolTypes.LH, Encoding.[Default])
    Dim m_Recorder_Client As WinSound.Recorder
    Dim m_PlayerClient As WinSound.Player
    Dim m_RecorderFactor As UInteger = 4
    Dim m_JitterBufferClientRecording As WinSound.JitterBuffer
    Dim m_JitterBufferClientPlaying As WinSound.JitterBuffer
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
    Dim m_Encoding As Encoding = Encoding.GetEncoding(1252)
    Dim RecordingJitterBufferCount As Integer = 8
    Dim SamplesPerSecondClient As Integer = 8000
    Dim BitsPerSampleClient As Integer = 16
    Dim ChannelsClient As Integer = 1
    Dim JitterBufferCountClient As UInteger = 20
    Dim UseJitterBufferClientRecording As Boolean = True
    Dim MuteClientPlaying As Boolean = False
    Dim ClientNoSpeakAll As Boolean = False
    Public Sub Starter(ByVal ip As String, ByVal port As Integer)
        Try
            IPSERVIDOR = ip
            PUERTO = port
            Try
                InitJitterBufferClientRecording()
                InitJitterBufferClientPlaying()
                InitTimerShowProgressBarPlayingClient()
                InitProtocolClient()
            Catch ex As Exception
                AddToLog("Starter(0)@Main", "Error: " & ex.Message, True)
                MsgBox(ex.Message)
            End Try
            Try
                If IsClientConnected Then
                    DisconnectClient()
                    StopRecordingFromSounddevice_Client()
                Else
                    ConnectClient()
                End If
                System.Threading.Thread.Sleep(100)
            Catch ex As Exception
                AddToLog("Starter(1)@Main", "Error: " & ex.Message, True)
                MsgBox(ex.Message)
            End Try
        Catch ex As Exception
            AddToLog("Starter(2)@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
            Button1.Enabled = True
        End Try
    End Sub
    Public Sub Stopper()
        Try
            m_IsFormMain = False
            StopRecordingFromSounddevice_Client()
            DisconnectClient()
        Catch ex As Exception
            AddToLog("Stopper@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub InitJitterBufferClientRecording()
        If m_JitterBufferClientRecording IsNot Nothing Then
            RemoveHandler m_JitterBufferClientRecording.DataAvailable, AddressOf OnJitterBufferClientDataAvailableRecording
        End If
        m_JitterBufferClientRecording = New WinSound.JitterBuffer(Nothing, RecordingJitterBufferCount, 20)
        AddHandler m_JitterBufferClientRecording.DataAvailable, AddressOf OnJitterBufferClientDataAvailableRecording
    End Sub
    Private Sub InitJitterBufferClientPlaying()
        If m_JitterBufferClientPlaying IsNot Nothing Then
            RemoveHandler m_JitterBufferClientPlaying.DataAvailable, AddressOf OnJitterBufferClientDataAvailablePlaying
        End If
        m_JitterBufferClientPlaying = New WinSound.JitterBuffer(Nothing, JitterBufferCountClient, 20)
        AddHandler m_JitterBufferClientPlaying.DataAvailable, AddressOf OnJitterBufferClientDataAvailablePlaying
    End Sub
    Private Sub InitTimerShowProgressBarPlayingClient()
        m_TimerProgressBarPlayingClient = New System.Windows.Forms.Timer()
        m_TimerProgressBarPlayingClient.Interval = 60
        AddHandler m_TimerProgressBarPlayingClient.Tick, AddressOf OnTimerProgressPlayingClient
    End Sub
    Private Sub InitProtocolClient()
        If m_PrototolClient IsNot Nothing Then
            AddHandler m_PrototolClient.DataComplete, AddressOf OnProtocolClient_DataComplete
        End If
    End Sub
    Private Sub FillRTPBufferWithPayloadData(ByVal header As WinSound.WaveFileHeader)
        m_RTPPartsLength = WinSound.Utils.GetBytesPerInterval(header.SamplesPerSecond, header.BitsPerSample, header.Channels)
        m_FilePayloadBuffer = header.Payload
    End Sub
    Private Sub OnTimerProgressPlayingClient(ByVal obj As [Object], ByVal e As EventArgs)
        Try
            If m_PlayerClient IsNot Nothing Then
            End If
        Catch ex As Exception
            m_TimerProgressBarPlayingClient.[Stop]()
            AddToLog("OnTimerProgressPlayingClient@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub StartRecordingFromSounddevice_Client()
        Try
            If IsRecorderFromSounddeviceStarted_Client = False Then
                Dim bufferSize As Integer = 0
                If UseJitterBufferClientRecording Then
                    bufferSize = WinSound.Utils.GetBytesPerInterval(CUInt(SamplesPerSecondClient), BitsPerSampleClient, ChannelsClient) * CInt(m_RecorderFactor)
                Else
                    bufferSize = WinSound.Utils.GetBytesPerInterval(CUInt(SamplesPerSecondClient), BitsPerSampleClient, ChannelsClient)
                End If
                If bufferSize > 0 Then
                    m_Recorder_Client = New WinSound.Recorder()
                    AddHandler m_Recorder_Client.DataRecorded, AddressOf OnDataReceivedFromSoundcard_Client
                    AddHandler m_Recorder_Client.RecordingStopped, AddressOf OnRecordingStopped_Client
                    If m_Recorder_Client.Start(Nothing, SamplesPerSecondClient, BitsPerSampleClient, ChannelsClient, m_SoundBufferCount, bufferSize) Then
                        ShowStreamingFromSounddeviceStarted_Client()
                        If UseJitterBufferClientRecording Then
                            m_JitterBufferClientRecording.Start()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            AddToLog("StartRecordingFromSounddevice_Client@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub StopRecordingFromSounddevice_Client()
        Try
            If IsRecorderFromSounddeviceStarted_Client Then
                m_Recorder_Client.[Stop]()
                RemoveHandler m_Recorder_Client.DataRecorded, AddressOf OnDataReceivedFromSoundcard_Client
                RemoveHandler m_Recorder_Client.RecordingStopped, AddressOf OnRecordingStopped_Client
                m_Recorder_Client = Nothing
                If UseJitterBufferClientRecording Then
                    m_JitterBufferClientRecording.[Stop]()
                End If
                ShowStreamingFromSounddeviceStopped_Client()
            End If
        Catch ex As Exception
            AddToLog("StopRecordingFromSounddevice_Client@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub OnRecordingStopped_Client()
        Try
            Me.Invoke(New MethodInvoker(Sub() ShowStreamingFromSounddeviceStopped_Client()))
        Catch ex As Exception
            AddToLog("OnRecordingStopped_Client@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub OnDataReceivedFromSoundcard_Client(ByVal data As [Byte]())
        Try
            SyncLock Me
                If IsClientConnected Then
                    If ClientNoSpeakAll = False Then
                        Dim bytesPerInterval As Integer = WinSound.Utils.GetBytesPerInterval(CUInt(SamplesPerSecondClient), BitsPerSampleClient, ChannelsClient)
                        Dim count As Integer = data.Length / bytesPerInterval
                        Dim currentPos As Integer = 0
                        For i As Integer = 0 To count - 1
                            Dim partBytes As [Byte]() = New [Byte](bytesPerInterval - 1) {}
                            Array.Copy(data, currentPos, partBytes, 0, bytesPerInterval)
                            currentPos += bytesPerInterval
                            Dim rtp As WinSound.RTPPacket = ToRTPPacket(partBytes, BitsPerSampleClient, ChannelsClient)
                            If UseJitterBufferClientRecording Then
                                m_JitterBufferClientRecording.AddData(rtp)
                            Else
                                Dim rtpBytes As [Byte]() = ToRTPData(data, BitsPerSampleClient, ChannelsClient)
                                m_Client.Send(m_PrototolClient.ToBytes(rtpBytes))
                            End If
                        Next
                    End If
                End If
            End SyncLock
        Catch ex As Exception
            AddToLog("OnDataReceivedFromSoundcard_Client@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub OnJitterBufferClientDataAvailableRecording(ByVal sender As [Object], ByVal rtp As WinSound.RTPPacket)
        Try
            If rtp IsNot Nothing AndAlso m_Client IsNot Nothing AndAlso rtp.Data IsNot Nothing AndAlso rtp.Data.Length > 0 Then
                If IsClientConnected Then
                    If m_IsFormMain Then
                        Dim rtpBytes As [Byte]() = rtp.ToBytes()
                        m_Client.Send(m_PrototolClient.ToBytes(rtpBytes))
                    End If
                End If
            End If
        Catch ex As Exception
            Dim sf As New System.Diagnostics.StackFrame(True)
            AddToLog("OnJitterBufferClientDataAvailableRecording@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub OnJitterBufferClientDataAvailablePlaying(ByVal sender As [Object], ByVal rtp As WinSound.RTPPacket)
        Try
            If m_PlayerClient IsNot Nothing Then
                If m_PlayerClient.Opened Then
                    If m_IsFormMain Then
                        If MuteClientPlaying = False Then
                            Dim linearBytes As [Byte]() = WinSound.Utils.MuLawToLinear(rtp.Data, BitsPerSampleClient, ChannelsClient)
                            m_PlayerClient.PlayData(linearBytes, False)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Dim sf As New System.Diagnostics.StackFrame(True)
            AddToLog("OnJitterBufferClientDataAvailablePlaying@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
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
    Private ReadOnly Property IsRecorderFromSounddeviceStarted_Client() As Boolean
        Get
            If m_Recorder_Client IsNot Nothing Then
                Return m_Recorder_Client.Started
            End If
            Return False
        End Get
    End Property
    Private Sub ConnectClient()
        Try
            If IsClientConnected = False Then
                If IPSERVIDOR.Length > 0 AndAlso PUERTO > 0 Then
                    m_Client = New TCPCLIENTE(IPSERVIDOR, PUERTO)
                    AddHandler m_Client.ClientDisconnected, AddressOf OnClientDisconnected
                    AddHandler m_Client.ExceptionAppeared, AddressOf OnClientExceptionAppeared
                    AddHandler m_Client.DataReceived, AddressOf OnClientDataReceived
                    m_Client.Connect()
                End If
            End If
        Catch ex As Exception
            m_Client = Nothing
            AddToLog("ConnectClient@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
            Me.Close()
        End Try
    End Sub
    Private Sub DisconnectClient()
        Try
            StopRecordingFromSounddevice_Client()
            If m_Client IsNot Nothing Then
                m_Client.Disconnect()
                RemoveHandler m_Client.ClientDisconnected, AddressOf OnClientDisconnected
                RemoveHandler m_Client.ExceptionAppeared, AddressOf OnClientExceptionAppeared
                RemoveHandler m_Client.DataReceived, AddressOf OnClientDataReceived
                m_Client = Nothing
            End If
        Catch ex As Exception
            AddToLog("DisconnectClient@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub OnClientDisconnected(ByVal client As TCPCLIENTE, ByVal info As String)
        StopPlayingToSounddevice_Client()
        StopRecordingFromSounddevice_Client()
        If m_Client IsNot Nothing Then
            RemoveHandler m_Client.ClientDisconnected, AddressOf OnClientDisconnected
            RemoveHandler m_Client.ExceptionAppeared, AddressOf OnClientExceptionAppeared
            RemoveHandler m_Client.DataReceived, AddressOf OnClientDataReceived
        End If
    End Sub
    Private Sub OnClientExceptionAppeared(ByVal client As TCPCLIENTE, ByVal ex As Exception)
        DisconnectClient()
    End Sub
    Private Sub OnClientDataReceived(ByVal client As TCPCLIENTE, ByVal bytes As [Byte]())
        Try
            If m_PrototolClient IsNot Nothing Then
                m_PrototolClient.Receive_LH(client, bytes)
            End If
        Catch ex As Exception
            AddToLog("OnClientDataReceived@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub OnProtocolClient_DataComplete(ByVal sender As [Object], ByVal data As [Byte]())
        Try
            If m_PlayerClient IsNot Nothing Then
                If m_PlayerClient.Opened Then
                    Dim rtp As New WinSound.RTPPacket(data)
                    If rtp.Data IsNot Nothing Then
                        If m_JitterBufferClientPlaying IsNot Nothing Then
                            m_JitterBufferClientPlaying.AddData(rtp)
                        End If
                    End If
                End If
            Else
                OnClientConfigReceived(sender, data)
            End If
        Catch ex As Exception
            AddToLog("OnProtocolClient_DataComplete@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private Sub OnClientConfigReceived(ByVal sender As [Object], ByVal data As [Byte]())
        Try
            Dim msg As [String] = m_Encoding.GetString(data)
            If msg.Length > 0 Then
                Dim values As [String]() = msg.Split(":"c)
                Dim cmd As [String] = values(0)
                Select Case cmd.ToUpper()
                    Case "SAMPLESPERSECOND"
                        Dim samplePerSecond As Integer = Convert.ToInt32(values(1))
                        SamplesPerSecondClient = samplePerSecond
                        Me.Invoke(New MethodInvoker(Sub()
                                                        StartPlayingToSounddevice_Client()
                                                        StartRecordingFromSounddevice_Client()
                                                    End Sub))
                        Exit Select
                End Select
            End If
        Catch ex As Exception
            AddToLog("OnClientConfigReceived@Main", "Error: " & ex.Message, True)
        End Try
    End Sub
    Private ReadOnly Property IsClientConnected() As Boolean
        Get
            If m_Client IsNot Nothing Then
                Return m_Client.Connected
            End If
            Return False
        End Get
    End Property
    Private Sub ShowStreamingFromSounddeviceStarted_Client()
        Try
            If Me.InvokeRequired Then
                Me.Invoke(New MethodInvoker(Sub() ShowStreamingFromSounddeviceStarted_Client()))
            Else
            End If
        Catch ex As Exception
            AddToLog("ShowStreamingFromSounddeviceStarted_Client@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ShowStreamingFromSounddeviceStopped_Client()
        Try
            If Me.InvokeRequired Then
                Me.Invoke(New MethodInvoker(Sub() ShowStreamingFromSounddeviceStopped_Client()))
            Else
            End If
        Catch ex As Exception
            AddToLog("ShowStreamingFromSounddeviceStopped_Client@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ShowStreamingFromFileStarted()
        Try
            If Me.InvokeRequired Then
                Me.Invoke(New MethodInvoker(Sub() ShowStreamingFromFileStarted()))
            Else
            End If
        Catch ex As Exception
            AddToLog("ShowStreamingFromFileStarted@Main", "Error: " & ex.Message, True)
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub StopStreamSounddevice_Client()
        StopRecordingFromSounddevice_Client()
    End Sub
    Private Sub StartPlayingToSounddevice_Client()
        If m_JitterBufferClientPlaying IsNot Nothing Then
            InitJitterBufferClientPlaying()
            m_JitterBufferClientPlaying.Start()
        End If
        If m_PlayerClient Is Nothing Then
            m_PlayerClient = New WinSound.Player()
            m_PlayerClient.Open(Nothing, SamplesPerSecondClient, BitsPerSampleClient, ChannelsClient, CInt(JitterBufferCountClient))
        End If
        m_TimerProgressBarPlayingClient.Start()
    End Sub
    Private Sub StopPlayingToSounddevice_Client()
        If m_PlayerClient IsNot Nothing Then
            m_PlayerClient.Close()
            m_PlayerClient = Nothing
        End If
        If m_JitterBufferClientPlaying IsNot Nothing Then
            m_JitterBufferClientPlaying.[Stop]()
        End If
        m_TimerProgressBarPlayingClient.[Stop]()
        Me.Invoke(New MethodInvoker(Sub()
                                    End Sub))
    End Sub
End Class
Public Class TCPCLIENTE
    Public Sub New(server As [String], port As Integer)
        Me.m_Server = server
        Me.m_Port = port
    End Sub
    Public Client As TcpClient
    Private m_NetStream As NetworkStream
    Private m_ByteBuffer As Byte()
    Private m_Server As [String]
    Private m_Port As Integer
    Private m_AutoConnect As Boolean = False
    Private m_TimerAutoConnect As System.Threading.Timer
    Private m_AutoConnectInterval As Integer = 10
    Public Overrides Function ToString() As String
        Return [String].Format("{0} {1}:{2}", Me.[GetType](), Me.m_Server, Me.m_Port)
    End Function
    Private Class Locker_AutoConnectClass
    End Class
    Private Locker_AutoConnect As New Locker_AutoConnectClass()
    Public Delegate Sub DelegateDataReceived(client As TCPCLIENTE, bytes As [Byte]())
    Public Delegate Sub DelegateDataSend(client As TCPCLIENTE, bytes As [Byte]())
    Public Delegate Sub DelegateDataReceivedComplete(client As TCPCLIENTE, message As [String])
    Public Delegate Sub DelegateConnection(client As TCPCLIENTE, Info As String)
    Public Delegate Sub DelegateException(client As TCPCLIENTE, ex As Exception)
    Public Event DataReceived As DelegateDataReceived
    Public Event DataSend As DelegateDataSend
    Public Event ClientConnected As DelegateConnection
    Public Event ClientDisconnected As DelegateConnection
    Public Event ExceptionAppeared As DelegateException
    Private Sub InitTimerAutoConnect()
        If m_AutoConnect Then
            If m_TimerAutoConnect Is Nothing Then
                If m_AutoConnectInterval > 0 Then
                    m_TimerAutoConnect = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf OnTimer_AutoConnect), Nothing, m_AutoConnectInterval * 1000, m_AutoConnectInterval * 1000)
                End If
            End If
        End If
    End Sub

    Public Sub Send(data As [Byte]())
        Try
            m_NetStream.Write(data, 0, data.Length)
            RaiseEvent DataSend(Me, data)
        Catch ex As Exception
            RaiseEvent ExceptionAppeared(Me, ex)
        End Try
    End Sub
    Private Sub StartReading()
        Try
            m_ByteBuffer = New Byte(1023) {}
            m_NetStream.BeginRead(m_ByteBuffer, 0, m_ByteBuffer.Length, New AsyncCallback(AddressOf OnDataReceived), m_NetStream)
        Catch ex As Exception
            RaiseEvent ExceptionAppeared(Me, ex)
        End Try
    End Sub
    Private Sub OnDataReceived(ar As IAsyncResult)
        Try
            Dim myNetworkStream As NetworkStream = DirectCast(ar.AsyncState, NetworkStream)
            If myNetworkStream.CanRead Then
                Dim numberOfBytesRead As Integer = myNetworkStream.EndRead(ar)
                If numberOfBytesRead > 0 Then
                    Dim data As [Byte]() = New Byte(numberOfBytesRead - 1) {}
                    System.Array.Copy(m_ByteBuffer, 0, data, 0, numberOfBytesRead)
                    RaiseEvent DataReceived(Me, data)
                Else
                    RaiseEvent ClientDisconnected(Me, "FIN")
                    If m_AutoConnect = False Then
                        Me.disconnect_intern()
                    Else
                        Me.Disconnect_ButAutoConnect()
                    End If
                    Return
                End If
                myNetworkStream.BeginRead(m_ByteBuffer, 0, m_ByteBuffer.Length, New AsyncCallback(AddressOf OnDataReceived), myNetworkStream)
            End If
        Catch ex As Exception
            RaiseEvent ExceptionAppeared(Me, ex)
        End Try
    End Sub
    Public Sub ReConnect()
        Me.Disconnect()
        Me.Connect()
    End Sub
    Public Sub Connect()
        Try
            InitTimerAutoConnect()
            Client = New TcpClient(Me.m_Server, Me.m_Port)
            m_NetStream = Client.GetStream()
            Me.StartReading()
            RaiseEvent ClientConnected(Me, [String].Format("server: {0} port: {1}", Me.m_Server, Me.m_Port))
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub Ping()
        Dim ping__1 As New Net.NetworkInformation.Ping()
        Dim reply As PingReply = ping__1.Send(m_Server)
        If reply.Status <> IPStatus.Success Then
            Throw New Exception([String].Format("Server {0} NO RESPONDE AL PING ", m_Server))
        End If
    End Sub
    Public Sub Ping(waitTimeout As Int32)
        Dim ping__1 As New Ping()
        Dim reply As PingReply = ping__1.Send(m_Server, waitTimeout)
        If reply.Status <> IPStatus.Success Then
            Throw New Exception([String].Format("Server {0} NO RESPONDE AL PING ", m_Server))
        End If
    End Sub
    Public Sub Disconnect()
        disconnect_intern()
        If m_TimerAutoConnect IsNot Nothing Then
            m_TimerAutoConnect.Dispose()
            m_TimerAutoConnect = Nothing
        End If
        RaiseEvent ClientDisconnected(Me, "CONEXION TERMINADA")
    End Sub
    Private Sub Disconnect_ButAutoConnect()
        disconnect_intern()
    End Sub
    Private Sub disconnect_intern()
        If Client IsNot Nothing Then
            Client.Close()
        End If
        If m_NetStream IsNot Nothing Then
            m_NetStream.Close()
        End If
    End Sub
    Private Sub OnTimer_AutoConnect(ob As [Object])
        Try
            SyncLock Locker_AutoConnect
                If m_AutoConnect Then
                    If Client Is Nothing OrElse Client.Connected = False Then
                        Client = New TcpClient(Me.m_Server, Me.m_Port)
                        m_NetStream = Client.GetStream()
                        Me.StartReading()
                        RaiseEvent ClientConnected(Me, [String].Format("server: {0} port: {1}", Me.m_Server, Me.m_Port))
                    End If
                Else
                    If m_TimerAutoConnect IsNot Nothing Then
                        m_TimerAutoConnect.Dispose()
                        m_TimerAutoConnect = Nothing
                    End If
                End If
            End SyncLock
        Catch ex As Exception
            RaiseEvent ExceptionAppeared(Me, ex)
        End Try
    End Sub
    Public Property AutoConnectInterval() As Int32
        Get
            Return m_AutoConnectInterval
        End Get
        Set(value As Int32)
            m_AutoConnectInterval = value
            If value > 0 Then
                Try
                    If m_TimerAutoConnect IsNot Nothing Then
                        m_TimerAutoConnect.Change(value * 1000, value * 1000)
                    End If
                Catch ex As Exception
                    RaiseEvent ExceptionAppeared(Me, ex)
                End Try
            End If
        End Set
    End Property
    Public Property AutoConnect() As Boolean
        Get
            Return m_AutoConnect
        End Get
        Set(value As Boolean)
            m_AutoConnect = value
            If value = True Then
                InitTimerAutoConnect()
            End If
        End Set
    End Property
    Public ReadOnly Property IsRunning() As Boolean
        Get
            Return m_TimerAutoConnect IsNot Nothing
        End Get
    End Property
    Public ReadOnly Property Connected() As Boolean
        Get
            If Me.Client IsNot Nothing Then
                Return Me.Client.Connected
            Else
                Return False
            End If
        End Get
    End Property
End Class