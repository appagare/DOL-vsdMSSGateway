VERSION 5.00
Object = "{C7212F93-30E8-11D2-B450-0020AFD69DE6}#1.0#0"; "SocketX.OCX"
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.ocx"
Begin VB.Form frmService 
   Caption         =   "Form1"
   ClientHeight    =   750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1815
   Icon            =   "frmService.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   750
   ScaleWidth      =   1815
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrQueueRetry 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   480
   End
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer tmrHeartbeat 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   0
   End
   Begin VB.Timer tmrRetry 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   0
   End
   Begin SocketXCtl.SocketXCtl ctlSocket1 
      Left            =   360
      Top             =   0
      AcceptTimeout   =   0
      BlockingMode    =   0
      Blocking        =   0   'False
      BroadcastEnabled=   -1  'True
      ConnectTimeout  =   0
      EventMask       =   63
      KeepAliveEnabled=   0   'False
      LibraryName     =   "WSOCK32.DLL"
      LingerEnabled   =   0   'False
      LingerMode      =   0
      LingerTime      =   0
      LocalAddress    =   ""
      LocalPort       =   0
      OutOfBandEnabled=   0   'False
      ReceiveBufferSize=   8192
      ReceiveTimeout  =   0
      RemoteAddress   =   ""
      RemoteName      =   ""
      ReuseAddressEnabled=   0   'False
      RemotePort      =   0
      RouteEnabled    =   -1  'True
      SendTimeout     =   0
      SendBufferSize  =   8192
      SocketType      =   0
      TcpNoDelayEnabled=   0   'False
      Secure          =   0   'False
      SecureProtocol  =   0
      SecureKeyExchange=   0
      SecureCertName  =   "MY\"
      SecureForceRemoteAuth=   0   'False
   End
   Begin NTService.NTService ctlNTService1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "Simple"
      StartMode       =   3
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DB connection info
'User ID=vsMSSGatewayUser
'Password=?M2S4S6g8u?

'Variable to detect Unload attempts so
'provide for an orderly clean up and shutdown
'routine.
Private mblnUnload As Boolean
'
'Variable to detect whether the queue should
'be processed. Used to ignore queue events when not
'connected to the MSS or during shutdown.
Private mblnProcessQueue As Boolean
'
'Variable to track the number of retry attempts
'compared with mudtParam.lngRetryCount
Private mlngConnectCount As Long

'Variable to track the number of retry attempts
'compared with mudtParam.lngQueueRetryCount
Private mlngQueueRetryCount As Long

'
'Basic operating parameters
Private Type udtParams
    strServer As String 'MSS IP
    intPort As Integer 'MSS port
    strMnemonic As String 'DOL terminal ID for MSS
    intHeartbeatInterval As Integer
    intRetryInterval As Integer
    lngRetryCount As Long
    intConnectTimeout As Integer 'socket connect timeout
    intSendTimeout As Integer 'socket send timeout
    strDBConnectString As String
    strQServer As String 'MSMQ server name
    strQTx As String 'common MSMQ output (tx) queue
    lngQRxTimeout As Long 'receive timeout for MSMQ Rx queues
    strEmail As String 'e-mail notification string
    blnLogMessage As Boolean 'debug flag to indicate whether to log sent messages to the event log.
    lngQueueRetryCount As Long
    intQueueRetryInterval As Integer
    blnAckUnsavedMessages As Boolean ' PTR 4057 -- debug switch that prevents ACCESS switch from queuing unsaved messages.
    strSourceIP As String 'to support machines with multiple NICs or bound IP addresses
    blnStandByMode As Boolean 'determines active/stand-by mode of each instance of the service
End Type
Private mudtParam As udtParams

Private mrsQueueList As ADODB.Recordset
Private mcnMSSGateway As ADODB.Connection
Private mstrSent As String          'temp. string to hold the sent
                                    'message in order to handle the reception
                                    'of the ACK in a multiple MSSGateway
                                    'situation.
                                    
Private mstrBuffer As String        'temp. string to hold the received
                                    'message until all winsock pieces
                                    'are recieved. This needs to be a
                                    'form level variable rather than a
                                    'static procedure level variable because
                                    'the value needs to be cleared in the event
                                    'of a communication error when only part of
                                    'the message has been received. The MSS Switch
                                    'resends the entire message under this
                                    'error condition.
                                    
Private mstrQueueMsg As String      'used by QueuException to save messages
                                    'when the queue service was yanked from beneath us
                                    
'Internal DOL Queue wrapper for the common Tx queue
Private WithEvents mobjTxQueue As vsQueueWrapper.clsQueueWrapper
Attribute mobjTxQueue.VB_VarHelpID = -1

'Internal array of DOL Queue wrappers for the Rx queue(s)
Private mobjRxQueue() As vsQueueWrapper.clsQueueWrapper
Attribute mobjRxQueue.VB_VarHelpID = -1

Private Sub ctlNTService1_Start(Success As Boolean)
    On Error Resume Next
    Dim blnTest As Boolean
    Dim strErrMsg As String
    
    'Open the database
    blnTest = OpenConnection
    
    'Determine if we have enough data
    'to proceed.
    If mudtParam.strServer = "" Or mudtParam.intPort < 1 Or _
        mudtParam.intHeartbeatInterval < 1 Or mudtParam.strMnemonic = "" Or _
        mudtParam.intConnectTimeout < 0 Or mudtParam.intSendTimeout < 0 Or _
        mudtParam.intRetryInterval < 0 Or mrsQueueList.BOF = True Then
        'This should only happen if the parameterized values in the
        'db are bad values.
        strErrMsg = "Insufficient parameters to proceed (ctlNTService1.Start)."
        
        Call LogEvent(svcEventError, svcMessageError, "[" & Err.Number & "] " & _
            Err.Description & " [" & strErrMsg & "] ")
        
        'Send Email
        SendMail mudtParam.strEmail, strErrMsg
        
        'Add a delay so the Service Controller can complete the initialization
        'before we terminate ourselves.
        tmrUnload.Interval = 5000
        tmrUnload.Enabled = True
        Exit Sub
    End If
    
    'We can proceed
    Success = True
    
    'Open the queues
    blnTest = OpenQueues
    If blnTest = False Then
        'One or more of the queues failed to open. An event was logged
        'for each failure, but we'll send one e-mail indicating this
        'failure.
        strErrMsg = "One or more queues failed to open. See the Application event log for details (ctlNTService1.Start)."
        SendMail mudtParam.strEmail, strErrMsg
    End If
    
    'Attempt a connection
    'ConnectSocket
    
    'Beginning the process directly from the Start service
    'event could possibly cause unstable behavior if
    'the service controller doesn't detect that the service starts
    'until the connection completes.
    '
    'However, starting the process after a short delay
    'eliminates this problem so we perform the initial connection
    'in the tmrUnload (misnomer) event which fires 1 second
    'after the service starts.
    tmrUnload.Tag = "startup"
    tmrUnload.Interval = 1000
    tmrUnload.Enabled = True
    
End Sub

Private Sub ctlNTService1_Stop()
    On Error GoTo Err_Stop
    Unload Me
    Exit Sub
    
Err_Stop:
    Call LogEvent(svcEventError, svcMessageError, "[" & Err.Number & "] " & Err.Description & " [Service stop event] ")
    'Shouldn't happen but try again...
    tmrUnload.Enabled = True
End Sub

Private Sub ctlSocket1_Close(ByVal ErrorCode As Integer)
    On Error GoTo Err_Close
    
    Dim rsSent As ADODB.Recordset
    Dim lngReturn As Long
    Dim blnReturn As Boolean
        
    'This event will fire when the remote server closes
    'the socket connection. It MAY fire under certain
    'network failure conditions. In either case, the
    'socket must be re-initialized and a connect attempt
    'should be made.
    
    'Check to see if we were waiting for an ACK
    'If so, put the message back in the queue so it
    'can be re-submitted per
    If mstrSent <> "" Then
        Set rsSent = New ADODB.Recordset
        Set rsSent = MsgFromMSS(mstrSent)
        blnReturn = QueueWrite(mudtParam.strQTx, rsSent)
        'log the message
        Call LogEvent(svcEventError, svcMessageError, _
            "Message not acked [" & mstrSent & "]")
        
        If (blnReturn) Then
            mstrSent = ""
        End If
        'If Not blnReturn Then
            'failed to write to Queue (possibly no match or queue disabled)
            'write to the unknown table
        '    lngReturn = WriteUnknownMessage(rsSent)
        'End If
        'mstrSent = ""
        Set rsSent = Nothing
    End If
    
    Call LogEvent(svcEventInformation, svcMessageInfo, _
                "Socket Disconnected.")
    
    mstrBuffer = ""
    ReInitializeSocket
    Exit Sub
    
Err_Close:
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & _
        " [Socket Close event.] ")
    ReInitializeSocket
End Sub

Private Sub ctlSocket1_Connect(ByVal ErrorCode As Integer)
    On Error GoTo Err_Connect
    
    If ErrorCode <> WSAENOERROR Then
        'Error occurred, must re-init socket
        'increment counters, etc.
        Call LogEvent(svcEventError, svcMessageError, "[Socket Connect error (" & _
            CStr(ErrorCode) & ") " & ctlSocket1.LastErrorString & " Remote IP:" & _
            ctlSocket1.RemoteAddress & " Remote Port:" & CStr(ctlSocket1.RemotePort) & "] ")
        ConnectionRetry
        Exit Sub
    End If
    
    Call LogEvent(svcEventInformation, svcMessageInfo, _
                "Socket Connected.")
    
    'Send the config string upon connection
    SendData Chr(STX) & "mode=t" & vbLf & _
        "heartbeat=" & CStr(mudtParam.intHeartbeatInterval) & _
        vbLf & "version=1" & Chr$(0) & Chr(ETX)
    
    Exit Sub
Err_Connect:
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & _
        " [Socket Connect event.] ")
    ConnectionRetry
End Sub

Private Sub ctlSocket1_Receive(ByVal ErrorCode As Integer)
    On Error GoTo Err_Receive
    
    Dim strData As String
    Dim rsMessage As ADODB.Recordset
    Dim rsSent As ADODB.Recordset
    Dim lngReturn As Long
    Dim blnReturn As Boolean
    Dim strErrMsg As String
    Dim strTemp As String
    
    strData = ctlSocket1.Receive(0)
    
    If InStr(1, strData, Chr(ACK), vbBinaryCompare) > 0 Then
        'Clear our temp string since the last message
        'was ACK'ed
        mstrSent = ""
        mlngConnectCount = 0
        If strData = Chr(ACK) Then
            'The only thing we received was an ACK, so
            'process the queue.
            Call ProcessQueue
            Exit Sub
        Else
            'Remove the ACK and continue processing
            strData = Replace(strData, Chr(ACK), "", , , vbBinaryCompare)
        End If
    End If
    
    
    
    
    If InStr(1, strData, Chr(ETX), vbBinaryCompare) > 0 Then
        'We have an ETX in the data
        mstrBuffer = mstrBuffer & Left(strData, InStr(1, strData, Chr(ETX), vbBinaryCompare))
                    
        Set rsMessage = New ADODB.Recordset
        Set rsMessage = MsgFromMSS(mstrBuffer)
        
        'PTR 4057 -- add rsMessage.EOF = True check for garbage detection
        If rsMessage.State = adStateClosed Or rsMessage.EOF = True Then
            'Note: We don't ACK the junk message.
            '
            'Shouldn't happen, but if so we can't really do
            'anything but notify someone that we got junk
            'E-Mail
            strErrMsg = "Received MSS Message had an unknown header (ctlSocket1.Receive [1])." & vbCr & vbCrLf
            If mudtParam.blnAckUnsavedMessages = False Then
                strErrMsg = strErrMsg & "Message not Ack'ed. Potential ACCESS switch queuing could occur (set AckUnsavedMessages to '1' in vsMSSGateway.tblMSSConfig to prevent ACCESS queuing)." & vbCr & vbCrLf
            Else
                strErrMsg = strErrMsg & "DO NOT DELETE THIS E-MAIL UNLESS PROBLEM IS RESOLVED! Message WAS Ack'ed and this e-mail is the ONLY copy of the message (set AckUnsavedMessages to '0' in vsMSSGateway.tblMSSConfig to allow ACCESS queuing)." & vbCr & vbCrLf
            End If
            strErrMsg = strErrMsg & mstrBuffer
            
            'start standby code
            If mudtParam.blnStandByMode = False Then
                'only send this e-mail if this instance is NOT in standby mode
                'Send an e-mail if we can
                SendMail mudtParam.strEmail, strErrMsg
                Call LogEvent(svcEventError, svcMessageError, "[" & strErrMsg & "] ")
            End If
            'end standby code
        Else
            If IsNull(rsMessage("QueueLabel")) Then
                'Unknown message or error response
                lngReturn = WriteUnknownMessage(rsMessage)
            Else
                'Put the message in the queue
                blnReturn = QueueWrite(rsMessage("QueueLabel"), rsMessage)
                
                'Commented out because failure of QueueWrite will cause
                'QueueException to re-attempt the write function. If
                'successful, it will ACK the switch for us. If not, it
                'will re-try QueueRetryCount times, after which it will
                'shutdown.
                '
                'If Not blnReturn Then
                '    'Failed to write to Queue (possibly no match or queue disabled).
                '    'Write to the unknown table
                '    lngReturn = WriteUnknownMessage(rsMessage)
                '    Call LogEvent(svcEventError, svcMessageError, "Unable to write to queue " & rsMessage("QueueLabel"))
                'End If
            End If
        End If
        
        'PTR 4057
        'make a copy in case we need to send an e-mail below
        strTemp = mstrBuffer
        mstrBuffer = Right(strData, Len(strData) - InStr(1, strData, Chr(ETX), vbBinaryCompare))
        'Release the resource
        Set rsMessage = Nothing
        
        'We only send an ACK if we've written the message to either
        'a queue or to the "unknown" table.
        If blnReturn Or lngReturn > 0 Then
            Call SendData(Chr(ACK))
        Else
            'shouldn't happen but log some message. We check strErrMsg because
            'a junk message could land us here but we've already sent the
            'e-mail for that.
            If strErrMsg = "" Then
                
                'PTR 4057 -- modify e-mail messsage for more descriptive info.
                strErrMsg = "Message not written to queue or table. (ctlSocket1.Receive [2])." & vbCr & vbCrLf
                If mudtParam.blnAckUnsavedMessages = False Then
                    strErrMsg = strErrMsg & "Message not Ack'ed. Potential ACCESS switch queuing could occur (set AckUnsavedMessages to '1' in vsMSSGateway.tblMSSConfig to prevent ACCESS queuing)." & vbCr & vbCrLf
                Else
                    strErrMsg = strErrMsg & "DO NOT DELETE THIS E-MAIL UNTIL PROBLEM IS IDENTIFIED! Message WAS Ack'ed and this e-mail is the ONLY copy of the message (set AckUnsavedMessages to '0' in vsMSSGateway.tblMSSConfig to allow ACCESS queuing)." & vbCr & vbCrLf
                End If
                strErrMsg = strErrMsg & strTemp
                strTemp = ""
                
                'Send an e-mail if we can
                SendMail mudtParam.strEmail, strErrMsg
                Call LogEvent(svcEventError, svcMessageError, "[" & strErrMsg & "] ctlSocket1_Receive")
            End If
            
            'PTR 4057 -- Ack it to prevent ACCESS queuing.
            If mudtParam.blnAckUnsavedMessages = True Then
                Call SendData(Chr(ACK))
            End If
            
        End If
    Else
        'Probably won't happen, but append the new data to our receive
        'buffer.
        mstrBuffer = mstrBuffer & strData
    End If
    Call ProcessQueue
    Exit Sub
Err_Receive:
    'Receive Timeout. Return the message to the Tx queue.
    If mstrSent <> "" Then
        Set rsSent = New ADODB.Recordset
        Set rsSent = MsgFromMSS(mstrSent)
        blnReturn = QueueWrite(mudtParam.strQTx, rsSent)
        'Log the message
        Call LogEvent(svcEventError, svcMessageError, "Message not acked [" & mstrSent & "]")
        If Not blnReturn Then
            'Failed to write to Queue (possibly no match or queue disabled).
            'Write to the unknown table
            lngReturn = WriteUnknownMessage(rsSent)
        End If
        mstrSent = ""
        Set rsSent = Nothing
    End If
    Call LogEvent(svcEventError, svcMessageError, "[" & Err.Number & "] " & Err.Description & " [Socket Receive event.] ")
    ReInitializeSocket
End Sub

Private Sub ctlSocket1_Send(ByVal ErrorCode As Integer)
    On Error GoTo Err_Send
    If ErrorCode <> WSAENOERROR Then
        'A Send timeout error occurred
        Call LogEvent(svcEventError, svcMessageError, "[Socket Send error (" & _
            CStr(ErrorCode) & ") " & ctlSocket1.LastErrorString & "] ")
        GoTo Err_Send
    End If
    
    Exit Sub
Err_Send:
    'Shouldn't happen but log an event if it does.
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & " [Socket Send event - Unknown cause.] ")
    ReInitializeSocket
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Load
    
    'set the service name to the App.EXEName so
    'multiple instances of the service can be
    'installed if necessary
    ctlNTService1.DisplayName = "vsMSSGateway (" & App.EXEName & ")"
    ctlNTService1.ServiceName = App.EXEName
    
    If InStr(1, Command, "-install", vbTextCompare) > 0 Then
        'the service must be installed before proceeding
        
        'enable Interactive mode in order to
        'display the message box for start-up or error
        ctlNTService1.Interactive = True
        ctlNTService1.StartMode = svcStartAutomatic
        If ctlNTService1.Install Then
            
            'check for the command line parameters
            mudtParam.strServer = GetValue("-s", "")
            mudtParam.intPort = GetValue("-p", 0)
            mudtParam.strMnemonic = GetValue("-n", "")
            mudtParam.strDBConnectString = GetValue("-d", "")
            mudtParam.strQTx = GetValue("-q", "")
            
            'save the values in the registry along with
            'our service info
            SetRegistryValue "Parameters", "Remote Server", mudtParam.strServer
            SetRegistryValue "Parameters", "Remote Port", mudtParam.intPort
            SetRegistryValue "Parameters", "Mnemonic", mudtParam.strMnemonic
            SetRegistryValue "Parameters", "DB Connect String", mudtParam.strDBConnectString
            SetRegistryValue "Parameters", "Tx Queue", mudtParam.strQTx
            
            MsgBox ctlNTService1.DisplayName & " installed successfully"
        Else
            MsgBox ctlNTService1.DisplayName & " failed to install"
        End If
        End
        'Unload Me
        'Exit Sub
    ElseIf InStr(1, Command, "-uninstall", vbTextCompare) > 0 Then
        ctlNTService1.Interactive = True
        If ctlNTService1.Uninstall Then
            MsgBox ctlNTService1.DisplayName & " uninstalled successfully"
        Else
            MsgBox ctlNTService1.DisplayName & " failed to uninstall"
        End If
        End
        'Unload Me
        'Exit Sub
    ElseIf InStr(1, Command, "-ide", vbTextCompare) > 0 Then
        ctlNTService1.Debug = True
    End If
    
    mudtParam.strServer = GetRegistryValue("Parameters", "Remote Server", "")
    mudtParam.intPort = CInt(GetRegistryValue("Parameters", "Remote Port", 0))
    mudtParam.strMnemonic = GetRegistryValue("Parameters", "Mnemonic", "")
    mudtParam.strDBConnectString = GetRegistryValue("Parameters", "DB Connect String", "Provider=sqloledb;Data Source=DOLDBOLYPROD01;Initial Catalog=vsMSSGateway;User id=vsMSSGatewayUser;Password=?M2S4S6g8u?;")
    mudtParam.strQTx = GetRegistryValue("Parameters", "Tx Queue", "")
    mudtParam.strSourceIP = GetRegistryValue("Parameters", "Local Server", "0.0.0.0")
    
    'start standby code
    'assume "not-standby" (i.e. - active/normal processing) mode - 0 = active mode, 1 = standby mode
    mudtParam.blnStandByMode = CBool(GetRegistryValue("Parameters", "StandByMode", "0"))
    'end standby code
        
    ' Set service options. Must be set before StartService
    ' is called (or in design mode).
    ctlNTService1.ControlsAccepted = svcCtrlStartStop
    
    'Instantiate module level objects to be used.
    Set mcnMSSGateway = New ADODB.Connection
    Set mrsQueueList = New ADODB.Recordset
        
    ' connect service to Windows NT services controller
    ctlNTService1.Interactive = False
    ctlNTService1.StartService
    
    Exit Sub
Err_Load:
    Call LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description & "]")
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    'gracefully close the socket.
    If mblnUnload = False Then
        mblnProcessQueue = False
        mblnUnload = True
        'Cancel = True
        ReInitializeSocket
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CloseQueues
    CloseConnection
End Sub

Private Sub mobjTxQueue_Arrived()
    On Error Resume Next
    If mblnProcessQueue = False Then
        'Ignore the event. We'll pick it up later
        'after we re-connect, etc.
        Exit Sub
    End If
    Call ProcessQueue
End Sub

Private Sub tmrHeartbeat_Timer()
    On Error Resume Next
    
    Static intSeconds As Integer
    If mudtParam.intHeartbeatInterval > intSeconds Then
        'increment and wait for the next event
        intSeconds = intSeconds + 1
    Else
        'disable the timer, reset the second offset,
        'and send the ENQ
        intSeconds = 0
        'Disable the heartbeat timer.
        tmrHeartbeat.Enabled = False
        'Send our ENQ command to the MSS.
        If ctlSocket1.State = soxConnected Then
            SendData Chr(ENQ)
        End If
        
        're-init the queues in the event the queue service
        'was yanked from beneath us
        CloseQueues
        If OpenQueues = False Then
            'Queue(s) failed to open;
            tmrQueueRetry.Interval = 1000
            tmrQueueRetry.Enabled = True
            Exit Sub
        Else
            mlngQueueRetryCount = 0
        End If
    End If
    'check the queue for any messages
    Call ProcessQueue

End Sub

Private Sub tmrQueueRetry_Timer()
    On Error Resume Next
    Static intSeconds As Integer
    Dim blnReturn As Boolean
    
    If tmrHeartbeat.Enabled = False Then
        tmrHeartbeat.Enabled = True
    End If
    
    If mudtParam.intQueueRetryInterval > intSeconds Then
        'increment and wait for the next event
        intSeconds = intSeconds + 1
    Else
        'disable the timer, reset the second offset,
        'and attempt to open
        tmrHeartbeat.Enabled = False
        tmrQueueRetry.Enabled = False
        intSeconds = 0
        Call QueueException
    End If
End Sub

Private Sub tmrRetry_Timer()
    On Error Resume Next
    Static intSeconds As Integer
    If mudtParam.intRetryInterval > intSeconds Then
        'increment and wait for the next event
        intSeconds = intSeconds + 1
    Else
        'disable the timer, reset the second offset,
        'and attempt to connect
        tmrRetry.Enabled = False
        intSeconds = 0
        Call LogEvent(svcEventInformation, svcMessageInfo, _
                "Connection retry.")
        ReInitializeSocket
    End If
End Sub

Private Sub tmrUnload_Timer()
    On Error Resume Next
    tmrUnload.Enabled = False
    If tmrUnload.Tag = "startup" Then
        tmrUnload.Tag = ""
        tmrUnload.Interval = 5000
        'initial start up connection attempt
        ConnectSocket
    Else
        If ctlNTService1.Debug = False Then
            ctlNTService1.StopService
        Else
            Unload Me
        End If
    End If
    
End Sub
'Purpose: Clean up DB resources
Private Sub CloseConnection()
    On Error Resume Next
    'Release the queuelist
    Set mrsQueueList = Nothing
    
    'Close and release the connection
    mcnMSSGateway.Close
    Set mcnMSSGateway = Nothing
End Sub
'Purpose: Close all queues and release resources.
Private Sub CloseQueues()
    On Error Resume Next
    Dim intIdx As Integer
    
    'Close the Tx queue
    mobjTxQueue.QueueClose
    
    'Close the Rx queue(s)
    For intIdx = 0 To UBound(mobjRxQueue)
        mobjRxQueue(intIdx).QueueClose
    Next intIdx
    
    'Release resources
    Set mobjTxQueue = Nothing
    
    intIdx = 0
    For intIdx = 0 To UBound(mobjRxQueue)
        Set mobjRxQueue(intIdx) = Nothing
    Next intIdx
End Sub
'Purpose:   Common entry point for re-initializing the socket
'           connection. This will be called either during an error
'           from the socket Connect event or failure of the MSS
'           to ACK the configuration command.
Private Sub ConnectionRetry()
    On Error Resume Next
    Dim strErrMsg As String
    mlngConnectCount = mlngConnectCount + 1
    If mlngConnectCount < mudtParam.lngRetryCount Then
        'if connected, close and start retry counter
        If ctlSocket1.State = soxConnected Then
            ctlSocket1.Close
        End If
        tmrRetry.Enabled = True
    Else
        strErrMsg = "Unable to establish socket connection. Tried " & _
            CStr(mudtParam.lngRetryCount) & " times."
        'Unable to connect within RetryCount times so
        'log event, notify, bail
        Call LogEvent(svcEventError, svcMessageError, strErrMsg)
        
        'Send an e-mail if we can
        SendMail mudtParam.strEmail, strErrMsg
        
        tmrUnload.Interval = 8000
        tmrUnload.Enabled = True
        Exit Sub
    End If
End Sub
'Purpose: Common location to initialize
'         the socket and initiate a connection
Private Sub ConnectSocket()

    On Error GoTo Err_ConnectSocket
    'configure and create the local socket
    ctlSocket1.Blocking = False
    ctlSocket1.LocalAddress = mudtParam.strSourceIP
    ctlSocket1.LocalPort = 0
    ctlSocket1.Create
    
    'set the remote server properties
    ctlSocket1.ConnectTimeout = mudtParam.intConnectTimeout
    ctlSocket1.SendTimeout = mudtParam.intSendTimeout
    ctlSocket1.RemoteAddress = mudtParam.strServer
    ctlSocket1.RemotePort = mudtParam.intPort
    
    'request a connection
    ctlSocket1.Connect
    Exit Sub
    
Err_ConnectSocket:
    'ok to ignore WSAEWOULDBLOCK errors
    If Err.Number <> WSAEWOULDBLOCK Then
        Call LogEvent(svcEventError, svcMessageError, "[" & Err.Number & "] " & Err.Description & " [Connect Socket] ")
    End If
    Resume Next
End Sub
'Purpose:   This function is called in an attempt to create each
'           queue because the QueueWrapper only opens queues.
Private Sub CreateQueues()
    On Error Resume Next
    Dim objQueueInfo As MSMQ.MSMQQueueInfo
    'Create the common Tx queue.
    Set objQueueInfo = New MSMQ.MSMQQueueInfo
    
    'default to the local machine if no server
    'is specified.
    If mudtParam.strQServer = "" Then
        mudtParam.strQServer = ".\PRIVATE$\"
    End If
    
    If Right(mudtParam.strQServer, 1) <> "\" Then
        mudtParam.strQServer = mudtParam.strQServer & "\"
    End If
    objQueueInfo.PathName = mudtParam.strQServer & mudtParam.strQTx
    objQueueInfo.Create False, True
    
    'Create the application specific Rx Queues
    mrsQueueList.MoveFirst
    Do While Not mrsQueueList.EOF
        Set objQueueInfo = Nothing
        Set objQueueInfo = New MSMQ.MSMQQueueInfo
        
        objQueueInfo.PathName = mudtParam.strQServer & mrsQueueList("QueueLabel")
        objQueueInfo.Create False, True
        
        mrsQueueList.MoveNext
    Loop
    mrsQueueList.MoveFirst
    Set objQueueInfo = Nothing
End Sub
'Purpose: Encapsulate the event registry read code.
'         Currently uses the integral registry
'         methods of the NTService Control which
'         retrieves the values from the Service
'         registry key.
Private Function GetRegistryValue(ByVal pstrSection As String, _
    ByVal pstrKey As String, _
    ByVal pstrDefaultValue As String) As String
    On Error Resume Next
    GetRegistryValue = ctlNTService1.GetSetting(pstrSection, _
        pstrKey, pstrDefaultValue)
    
End Function
'Purpose: Parse the command line parameters
'         to locate a specific value.
'
'Inputs: pstrKey=parameter key name
'        pvarDefaultValue -- optional default value
'        to return if key not found.
'
'        Possible Keys
'        -s = remote server
'        -p = remote port
'        -n = mnemonic name
'        -q = Tx queue name
Private Function GetValue(pstrKey As String, _
    Optional pvarDefaultValue)
    
    On Error GoTo Err_ParseCommand
    
    Dim intStartPos As Integer
    Dim intNextSpace As Integer
    
    If Not IsMissing(pvarDefaultValue) Then
        GetValue = pvarDefaultValue
    End If
    
    'shouldn't happen but handle it anyways
    If Command = "" Or pstrKey = "" Then
        Exit Function
    End If
    
    'locate the starting position of the key in the
    'command line parameters
    intStartPos = InStr(1, Command, pstrKey, vbTextCompare)
    If intStartPos > 0 Then
        'found the key, see if its the last parameter in the string
        intNextSpace = InStr(intStartPos + Len(pstrKey) + 1, _
            Command, " ")
        'if its the last position, set our end position to the length
        'of the command line; otherwise, the ending position will be
        'the location of the next space char.
        If intNextSpace = 0 Then
            intNextSpace = Len(Command) + 1
        End If
        GetValue = Trim(Mid(Command, intStartPos + Len(pstrKey) + 1, _
            intNextSpace - (Len(pstrKey) + intStartPos + 1)))
    End If
    Exit Function

Err_ParseCommand:
    Call LogEvent(svcMessageError, svcEventError, _
        "[" & Err.Number & "] " & Err.Description)
    Resume Next
End Function

'Purpose: Encapsulate the event logging code.
'         Currently uses the integral event logging
'         methods of the NTService Control.
'
'Inputs: pintEventType=enumerated integer value indicating
'        event type (error, information, success, message)
'        pintEventID=
'        pstrEventMessage=event description.
'
'Note:  Using VB to create an NT service requires that all functions
'       that return values either accept the return value or explicitly
'       ignore the value.
'
'       Because the NT Service control returns a value, procedures
'       which call this function MUST either accept the return value
'       or invoke the function using the CALL keyword which instructs
'       VB to accept and ignore the return value.
'
Private Function LogEvent(ByVal pintEventType As SvcEventType, _
    ByVal pintEventID As SvcEventId, _
    ByVal pstrEventMessage As String) As Boolean
    On Error Resume Next
    
    pstrEventMessage = "[Ver. " & CStr(App.Major) & _
                "." & CStr(App.Minor) & "." & CStr(App.Revision) & "]" & vbCrLf & pstrEventMessage
    
    If ctlNTService1.Debug = True Then
        Debug.Print Now() & " " & pstrEventMessage
    Else
        LogEvent = ctlNTService1.LogEvent(pintEventType, _
            pintEventID, pstrEventMessage)
    End If
End Function

'Purpose:   Parse a MSS Message and return a recordset of the
'           fields.
'Input:     String containing an MSS message.
'Output:    Recordset consisting of
'           QueueLabel (if known)
'           MsgDate
'           OrigID
'           Aux
'           Mnem
'           Delimiter ("." or ")")
'           Body
Private Function MsgFromMSS(ByVal pstrMSSMessage As String) As ADODB.Recordset
    On Error GoTo Err_MsgFromMSS
    
    Dim rs As ADODB.Recordset
    Dim intTest As Integer
    
    Set rs = New ADODB.Recordset
    
    'Create an empty recordset to return
    rs.Fields.Append "MsgDate", adDate, 8, adFldIsNullable
    rs.Fields.Append "QueueLabel", adVarChar, 8, adFldIsNullable
    rs.Fields.Append "OrigID", adChar, 5, adFldIsNullable
    rs.Fields.Append "Aux", adChar, 4, adFldIsNullable
    rs.Fields.Append "Mnem", adVarChar, 255, adFldIsNullable
    rs.Fields.Append "Delimiter", adChar, 1
    'PTR 4057 -- make sure Body is long enough
    rs.Fields.Append "Body", adLongVarWChar, Len(pstrMSSMessage) + 1
    rs.Open
    
    pstrMSSMessage = Trim(pstrMSSMessage)
    
       
    'strip the STX/ETX
    pstrMSSMessage = Replace(pstrMSSMessage, Chr(STX), "", , , vbBinaryCompare)
    pstrMSSMessage = Replace(pstrMSSMessage, Chr(ETX), "", , , vbBinaryCompare)
    
    If Len(pstrMSSMessage) < 16 Then
        'Not enough characters to do anything
        'so return an empty recordset
        Set MsgFromMSS = rs
        Exit Function
    End If
    
    'Determine if the message is an error or a meaningful message
    If InStr(1, pstrMSSMessage, "REJ FLD ERR", vbTextCompare) > 0 Then
        'Error
        rs.AddNew
        rs("MsgDate") = Now
        rs("Delimiter") = Mid(pstrMSSMessage, InStr(1, pstrMSSMessage, "REJ FLD ERR", vbTextCompare) - 1, 1)
        If rs("Delimiter") <> "." And rs("Delimiter") <> ")" Then
            rs("Delimiter") = "."
        End If
        
        'PTR 4057 -- pad Body
        rs("Body") = pstrMSSMessage & " "
        rs.Update
    Else
        'Meaningful message
        rs.AddNew
        rs("MsgDate") = Now
        rs("OrigID") = Left(pstrMSSMessage, 5)
        rs("Aux") = Mid(pstrMSSMessage, 6, 4)
        
        If Mid(pstrMSSMessage, 15, 1) = "." Or Mid(pstrMSSMessage, 15, 1) = ")" Then
            'typical value or db connection
            intTest = 15
        Else
            'possible regional connection and multiple destination mnems
            intTest = InStr(10, pstrMSSMessage, ".")
            If intTest < 15 Then
                'Shouldn't happen, but fill the fields assuming
                'the req'd delimiter was in pos. 15
                intTest = 15
            End If
        End If
        
        rs("Mnem") = Mid(pstrMSSMessage, 10, intTest - 10)
        rs("Delimiter") = Mid(pstrMSSMessage, intTest, 1)
        rs("Body") = Right(pstrMSSMessage, Len(pstrMSSMessage) - intTest)
        
        'Attempt to locate the QueueID
        If mrsQueueList.EOF = False Then
            mrsQueueList.MoveFirst
            Do While Not mrsQueueList.EOF
                If Mid(pstrMSSMessage, _
                    mrsQueueList("IDStartPos") + (intTest - mrsQueueList("IDStartPos") + 1), _
                    Len(Trim(mrsQueueList("MsgKey")))) = Trim(mrsQueueList("MsgKey")) Then
                    'Found a match in the Queue List
                    rs("QueueLabel") = Trim(mrsQueueList("QueueLabel"))
                    Exit Do
                ElseIf Left(pstrMSSMessage, 4) = "NCIC" _
                    And Trim(rs("Aux")) = Trim(mrsQueueList("MsgKey")) Then
                    'Special case for the NCIC.
                    'NCIC does not return the Message Key in its responses
                    'so check to see if the initiating application stored it in the
                    'Aux field.
                    'Found a match in the Queue List
                    rs("QueueLabel") = Trim(mrsQueueList("QueueLabel"))
                    Exit Do
                End If
                mrsQueueList.MoveNext
            Loop
        End If
        rs.Update
    End If
    
    'last ditch attempt to ID the message
    If IsNull(rs("QueueLabel")) And Trim(rs("Aux")) <> "" Then
        'Try the Aux field
        rs("QueueLabel") = Trim(rs("Aux"))
    End If
    
    Set MsgFromMSS = rs.Clone
    Set rs = Nothing
    
    Exit Function
Err_MsgFromMSS:
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & " [MsgFromMSS] ")
    'return an empty recordset
    Set MsgFromMSS = rs
End Function
'Purpose:   Return a MSS Message from a recordset
'Input:     Recordset consisting of
'           MsgDate
'           QueueID (if known)
'           OrigID
'           Aux
'           Mnem
'           Body
'Output:    String formatted as a MSS message.
Private Function MSSFromMsg(ByVal prsMessage As ADODB.Recordset) As String
    On Error GoTo Err_MSSFromMsg
    Dim strTemp As String
    
    If prsMessage.BOF = True Then
        'Empty recordset; shouldn't happen but bail out if
        'it occurs
        Exit Function
    End If
    
    'IMPORTANT NOTE:
    'The wrong number of delimiters can bring down the
    'NCIC connection (trust me), so strip them out to
    'insure no delimiters are in the fields.
    prsMessage("OrigID") = Replace(prsMessage("OrigID"), ".", " ")
    prsMessage("Aux") = Replace(prsMessage("Aux"), ".", " ")
    prsMessage("Mnem") = Replace(prsMessage("Mnem"), ".", " ")
    
    prsMessage("OrigID") = Replace(prsMessage("OrigID"), ")", " ")
    prsMessage("Aux") = Replace(prsMessage("Aux"), ")", " ")
    prsMessage("Mnem") = Replace(prsMessage("Mnem"), ")", " ")
           
    strTemp = Left(Format(Trim(prsMessage("OrigID")), "!@@@@@"), 5)
    'Handle cases where the calling app. does not insert
    'the Originating Terminal ID
    If Trim(strTemp) = "" Then
        strTemp = Left(Format(Trim(mudtParam.strMnemonic), "!@@@@@"), 5)
    End If
    
    If Trim(prsMessage("Aux")) = "" Then
        'no Aux
        strTemp = strTemp & Space(4)
    Else
        'Aux
        strTemp = strTemp & Left(Format(Trim(prsMessage("Aux")), "!@@@@"), 4)
    End If
    
    If Trim(prsMessage("Mnem")) = "" Then
        'shouldn't happen
        strTemp = strTemp & Space(5)
    Else
        strTemp = strTemp & Left(Format(Trim(prsMessage("Mnem")), "!@@@@@"), 5)
    End If
    
    strTemp = strTemp & prsMessage("Delimiter")
    strTemp = strTemp & prsMessage("Body")
    
    'Bookend the message with STX/ETX
    MSSFromMsg = Chr(STX) & strTemp & Chr(ETX)
    Exit Function
Err_MSSFromMsg:
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & " [MSSFromMsg] ")
    'Return nothing rather than chance sending a malformed message
    'and bringing down the state's crime computers (again).
    MSSFromMsg = ""
End Function
'Purpose:   Database startup code. Open the DB connection object,
'           get the QueueList, read the parameters
Private Function OpenConnection() As Boolean
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim intState As Integer
    
    'Open the connection object.
    On Error Resume Next
    mcnMSSGateway.Close
    On Error GoTo Err_OpenConnection
    mcnMSSGateway.Open mudtParam.strDBConnectString
        
    intState = intState + 1
    
    'Create a local command object
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = mcnMSSGateway
    
    'Open the QueueList recordset for readonly
    cmd.CommandText = "usp_GetQueues"
    cmd.CommandType = adCmdStoredProc
    mrsQueueList.CursorLocation = adUseClient
    mrsQueueList.Open cmd, , adOpenStatic, adLockReadOnly
    
    intState = intState + 1
    
    'Get the configuration values
    cmd.CommandText = "usp_GetConfigurationValues"
    cmd.CommandType = adCmdStoredProc
    
    Set rs = cmd.Execute
    
    Do While Not rs.EOF
        Select Case UCase(rs("ParamName"))
            Case "ACKUNSAVEDMESSAGES"
                'PTR 4057
                'Prevent ACCESS switch from queuing by acking
                'messages that were not saved or were considered garbage.
                mudtParam.blnAckUnsavedMessages = False
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.blnAckUnsavedMessages = CBool(rs("ParamValue"))
                End If
            Case "CONNECTTIMEOUT"
                'Socket Connect Timeout
                mudtParam.intConnectTimeout = 30
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.intConnectTimeout = CInt(rs("ParamValue"))
                End If
            Case "EMAIL"
                'E-mail notification string
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.strEmail = rs("ParamValue")
                End If
            Case "HEARTBEATINTERVAL"
                'Heartbeat Interval
                mudtParam.intHeartbeatInterval = 60
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.intHeartbeatInterval = CInt(rs("ParamValue"))
                End If
            Case "LOGSENTMESSAGES"
                'Log mode for sent messages
                mudtParam.blnLogMessage = False
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.blnLogMessage = CBool(rs("ParamValue"))
                End If
            Case "QUEUERETRYCOUNT"
                'Retry Count
                mudtParam.lngQueueRetryCount = 10
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.lngQueueRetryCount = CLng(rs("ParamValue"))
                End If
            Case "QUEUERETRYINTERVAL"
                'Retry Interval
                mudtParam.intQueueRetryInterval = 20
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.intQueueRetryInterval = CInt(rs("ParamValue"))
                End If
            Case "QRXTIMEOUT"
                'Queue Timeout
                mudtParam.lngQRxTimeout = 100
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.lngQRxTimeout = CLng(rs("ParamValue"))
                End If
            Case "QSERVERNAME"
                'Queue server
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.strQServer = rs("ParamValue")
                End If
            Case "QTX"
                'Output Queue
                If mudtParam.strQTx = "" Then
                    'if we didn't get the value from the registry,
                    'get the value from the db
                    mudtParam.strQTx = "QTx"
                    If Not IsNull(rs("ParamValue")) Then
                        mudtParam.strQTx = rs("ParamValue")
                    End If
                End If
            Case "RETRYCOUNT"
                'Retry Count
                mudtParam.lngRetryCount = 1440
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.lngRetryCount = CLng(rs("ParamValue"))
                End If
            Case "RETRYINTERVAL"
                'Retry Interval
                mudtParam.intRetryInterval = 60
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.intRetryInterval = CInt(rs("ParamValue"))
                End If
            Case "SENDTIMEOUT"
                'Socket Send Timeout
                mudtParam.intSendTimeout = 30
                If Not IsNull(rs("ParamValue")) Then
                    mudtParam.intSendTimeout = CInt(rs("ParamValue"))
                End If
        End Select
        rs.MoveNext
    Loop
    
    OpenConnection = True
    
    'Release the unnecessary object resources
    rs.Close
    Set rs = Nothing
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    
    Exit Function
Err_OpenConnection:
    
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & " [Opening DB connection.] ")
    
    'Create an empty recordset in the event accessing the
    'QueueList fails.
    If intState < 2 Then
        mrsQueueList.Fields.Append "MsgKey", adVarChar, 8, adFldKeyColumn
        mrsQueueList.Fields.Append "QueueLabel", adVarChar, 8
        mrsQueueList.Fields.Append "IdStartPos", adInteger, 4
        mrsQueueList.Fields.Append "QueueName", adVarChar, 255
        mrsQueueList.Fields.Append "Disabled", adBoolean
        mrsQueueList.Open
    End If
    
    Resume Next
    
End Function

'Purpose:   Open the common Tx queue and all of the application
'           specific Rx queues. Notify someone if an enabled queue
'           cannot be opened.
Private Function OpenQueues() As Boolean
    On Error Resume Next
    Dim intIdx As Integer
    
    Set mobjTxQueue = New vsQueueWrapper.clsQueueWrapper
    
    If mrsQueueList.BOF Then
        'no queues, bail
        Call LogEvent(svcEventError, svcMessageError, "No queues available to open!")
        Exit Function
    End If
    
    'The Queue wrapper only opens queues. This procedure
    'will create all of the necessary queues.
    CreateQueues
    
    'Assume true.
    OpenQueues = True
    
    'Open the Tx queue
    mobjTxQueue.QueueOpen mudtParam.strQTx, _
        MQ_RECEIVE_ACCESS, MQ_DENY_NONE, _
        mudtParam.strQServer, mudtParam.lngQRxTimeout, False
    
    If mobjTxQueue.IsOpen = False Then
        'Fail function call but proceed.
        OpenQueues = False
        Call LogEvent(svcEventError, svcMessageError, _
            "Unable to open queue [" & mudtParam.strQTx & "] " & _
            CStr(Err.Number) & " " & Err.Description)
            
    End If
        
    'Open the Rx queue(s)
    mrsQueueList.MoveFirst
    ReDim mobjRxQueue(intIdx) As vsQueueWrapper.clsQueueWrapper
    Set mobjRxQueue(intIdx) = New vsQueueWrapper.clsQueueWrapper
    Do While Not mrsQueueList.EOF
        If Not mrsQueueList("Disabled") Then
            'If enabled, create a slot in the object array
            If intIdx <> 0 Then
                ReDim Preserve mobjRxQueue(UBound(mobjRxQueue) + 1) As vsQueueWrapper.clsQueueWrapper
            End If
            Set mobjRxQueue(UBound(mobjRxQueue)) = New vsQueueWrapper.clsQueueWrapper
            'Open the queue
            mobjRxQueue(UBound(mobjRxQueue)).QueueOpen mrsQueueList("QueueLabel"), _
                MQ_SEND_ACCESS, MQ_DENY_NONE, mudtParam.strQServer
                
            If mobjRxQueue(UBound(mobjRxQueue)).IsOpen = True Then
                intIdx = intIdx + 1
            Else
                'Fail function call but proceed.
                OpenQueues = False
                Call LogEvent(svcEventError, svcMessageError, _
                    "Unable to open queue [" & mrsQueueList("QueueLabel") & "] " & _
                    CStr(Err.Number) & " " & Err.Description)
            End If
        End If
        mrsQueueList.MoveNext
    Loop
    mrsQueueList.MoveFirst
    
End Function

'Purpose:   Check in common Tx queue, read a message if
'           present, format the message for transmission
'           to the MSS, and invoke the socket Send method.
Private Sub ProcessQueue()
    On Error Resume Next
    Dim strMSS As String
    Dim rsMessage As ADODB.Recordset
    Set rsMessage = New ADODB.Recordset
    
    mblnProcessQueue = False
    
    'start standby code
    If mudtParam.blnStandByMode = True Then
        'if instance is in standby mode, simulate empty queue actions
        Set rsMessage = Nothing
        mblnProcessQueue = True
        tmrHeartbeat.Enabled = True
        Exit Sub
    End If
    'end standby code
    
    'Read the queue
    Set rsMessage = mobjTxQueue.ReceiveMessage(False, mudtParam.lngQRxTimeout)
    
    If rsMessage.EOF = False Then
        'Queue had a message; convert it to a string
        strMSS = MSSFromMsg(rsMessage)
        'Shut off the heartbeat for a moment
        tmrHeartbeat.Enabled = False
        'Hold a copy in memory until we get our ACK
        mstrSent = strMSS
        Set rsMessage = Nothing
        
        'Debug utility to log sent messages
        If mudtParam.blnLogMessage = True Then
            Call LogEvent(svcEventInformation, svcMessageDebug, _
                strMSS)
        End If
        
        Call SendData(strMSS)
        Exit Sub
    Else
        Set rsMessage = Nothing
        mblnProcessQueue = True
        'No message in queue.
        tmrHeartbeat.Enabled = True
    End If
    
End Sub

'Purpose:   Writes a message to a queue.
'Inputs:    pstrQueueLabel indicates the queue in which to
'           write the message.
'           prsMessage contains a recordset which contains the
'           message to be written.
Private Function QueueWrite(ByVal pstrQueueLabel, _
    ByVal prsMessage As ADODB.Recordset) As Boolean
    On Error GoTo Err_QueueWrite
    Dim intIdx As Integer
    Dim lngReturn As Long
    
    'start standby code
    If mudtParam.blnStandByMode = True Then
        'if instance is in standby mode, simulate incoming Queue Write successful
        QueueWrite = True
        Exit Function
    End If
    'end standby code
    
    'Returning a message to the Tx Queue
    If pstrQueueLabel = mudtParam.strQTx Then
        'MSS failed to ACK a message. Stop processing,
        'return the message to the queue and attempt to
        're-connect.
        mblnProcessQueue = False
        
        'create a writable Tx queue object
        Dim objTxQueue As clsQueueWrapper
        Set objTxQueue = New clsQueueWrapper
        objTxQueue.QueueOpen mudtParam.strQTx, MQ_SEND_ACCESS, MQ_DENY_NONE, mudtParam.strQServer
        
        'write the message
        objTxQueue.SendMessage prsMessage
        
        On Error Resume Next
        objTxQueue.QueueClose
        Set objTxQueue = Nothing
        QueueWrite = True
        ReInitializeSocket
        Exit Function
    Else
        'Write the message to an incoming queue.
        For intIdx = 0 To UBound(mobjRxQueue)
            If mobjRxQueue(intIdx).QueueLabel = pstrQueueLabel Then
                'if queue is open, attempt to send the message
                If mobjRxQueue(intIdx).IsOpen = True Then
                    mobjRxQueue(intIdx).SendMessage prsMessage
                    'message written safely
                    QueueWrite = True
                    Exit Function
                End If
                Exit For
            End If
        Next intIdx
    End If
    'don't allow QueueException to occur when we
    'know the queue isn't open or we can't write to a queue
    Call LogEvent(svcEventError, svcMessageError, _
        " Unable to locate queue or queue not open. [QueueWrite] " & pstrQueueLabel & ")")
    lngReturn = WriteUnknownMessage(prsMessage)
    If lngReturn > 0 Then
        'return true indicating we captured the record and
        'the ACCESS switch can stop sending it.
        QueueWrite = True
    End If
    Exit Function
Err_QueueWrite:
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & _
        " Error writing to queue " & prsMessage("QueueLabel") & ".")
    
    'if this is the first time this error occured, then
    'execute QueueException immediately. Otherwise, start the
    'timer.
    If mlngQueueRetryCount = 0 Then
        If mstrSent <> "" Then
            mstrQueueMsg = mstrSent
        ElseIf mstrBuffer <> "" Then
            mstrQueueMsg = mstrBuffer
        End If
        Call QueueException
    Else
        tmrQueueRetry.Enabled = True
    End If
    
End Function
'Purpose: Common location to close
'         the socket.
Private Sub ReInitializeSocket()
    On Error Resume Next
    '
    'clear the process queue
    mblnProcessQueue = False
    '
    'disable the retry timer in the
    'event that this procedure was
    'called by QueryUnload and tmrRetry
    'was enabled.
    tmrRetry.Enabled = False
    '
    'close the socket
    ctlSocket1.Close
    '
    'enable normal error handling
    On Error GoTo Err_SocketInit
    
    If Not mblnUnload Then
        ConnectSocket
    End If
    
    Exit Sub
Err_SocketInit:
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & _
        " [Socket Initialization] ")
    Resume Next
End Sub

'Purpose: Common location to send data out
'         the socket.
'Effects: If an error occurs, it may be trapped
'         by the Send event.
Private Sub SendData(pstrData As String)
    On Error GoTo Err_SendData
    
    'Only proceed if we have something to send
    If pstrData = "" Then
        Exit Sub
    End If
    
    If ctlSocket1.State <> soxConnected Then
        'if the connection dropped undetected,
        'force an error
        Error WSAECONNRESET
    End If
    
    ctlSocket1.SendBuffer = pstrData
    ctlSocket1.Send
    Exit Sub
    
Err_SendData:
    Call LogEvent(svcEventError, svcMessageError, _
        "[" & Err.Number & "] " & Err.Description & _
        " [Send Data] " & pstrData)
    ReInitializeSocket
End Sub

'Purpose:   Wrapper for the stored procedure which will send
'           an e-mail. Used to supplement the event logging.
Private Sub SendMail(ByVal pstrTo As String, ByVal pstrMessage As String)
    On Error GoTo Err_SendMail
    Dim cmd As ADODB.Command
    
    If pstrTo = "" Then
        'Shouldn't happen, but make sure we have a recipient
        Err.Raise 327, _
        , "The data value could not be found. SendMail procedure is missing a recipient address."
        
        GoTo Err_SendMail
    End If
    
    'Create a local command object
    Set cmd = New ADODB.Command
    
    'Utilize our open connection
    cmd.ActiveConnection = mcnMSSGateway
    
    'Set the parameters
    cmd.CommandText = "usp_SendMail"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@recipients", _
        adVarChar, adParamInput, Len(pstrTo), pstrTo)
    
    cmd.Parameters.Append cmd.CreateParameter("@subject", _
        adVarChar, adParamInput, Len(App.EXEName & " Error"), _
        App.EXEName & " Error")
    
    cmd.Parameters.Append cmd.CreateParameter("@message", _
        adVarChar, adParamInput, Len(pstrMessage), pstrMessage)
    
    'Execute without a resulting recordset
    cmd.Execute , , ADODB.adExecuteNoRecords
    
    'Release resources
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Exit Sub
    
Err_SendMail:
    If Err.Number = 3704 Then
        If Not OpenConnection Then
            'Couldn't open
            Call LogEvent(svcEventError, svcMessageError, "[" & Err.Number & "] " & _
                Err.Description & " [Connection to database was lost and could not be re-established.] ")
        End If
    Else
        Call LogEvent(svcEventError, svcMessageError, _
            "[" & Err.Number & "] " & Err.Description & _
            " [SendMail] ")
    End If
End Sub
'Purpose: Encapsulate the event registry write code.
'         Currently uses the integral registry
'         methods of the NTService Control which
'         stores the values with the Service registry
'         key.
Private Sub SetRegistryValue(ByVal pstrSection As String, _
    ByVal pstrKey As String, ByVal pstrValue As String)
    
    On Error Resume Next
    ctlNTService1.SaveSetting pstrSection, pstrKey, pstrValue
End Sub
'Purpose:   Saves unknown messages OR messages whose queue is
'           not open.
'Inputs:    Recordset containing the parsed message
'Outputs:   Record Identity value
Private Function WriteUnknownMessage(ByVal prsMessage As ADODB.Recordset) As Long
    On Error GoTo Err_WriteUnknownMessage
    Dim strMessage As String
    Dim cmd As ADODB.Command
    Dim intTest As Integer
    
    'start standby code
    If mudtParam.blnStandByMode = True Then
        'if instance is in standby mode, simulate successful write to unknown table
        'note - not sure why this would happen but let active not write it's copy of the record
        WriteUnknownMessage = 1
        Exit Function
    End If
    'end standby code
    
    'PTR 4057 make sure Body size is > zero
    Dim lngBodySize As Long
    lngBodySize = 0
    If Not IsNull(prsMessage("Body")) Then
        lngBodySize = Len(prsMessage("Body"))
    End If
    lngBodySize = lngBodySize + 1
    
    'Create a local command object
    Set cmd = New ADODB.Command
    
    'Utilize our open connection
    cmd.ActiveConnection = mcnMSSGateway
    
    'Set the parameters
    cmd.CommandText = "usp_AddUnknownMessage"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@intUnkMsgID", _
        adInteger, adParamOutput, 4)
    
    cmd.Parameters.Append cmd.CreateParameter("@pdatMsgDate", _
        adDate + adEmpty, adParamInput, 8, prsMessage("MsgDate"))
    
    cmd.Parameters.Append cmd.CreateParameter("@pstrQueueLabel", _
        adVarChar + adEmpty, adParamInput, 8, _
        prsMessage("QueueLabel"))
    
    cmd.Parameters.Append cmd.CreateParameter("@pstrOrigID", _
        adChar, adParamInput, 5, prsMessage("OrigID"))
    
    cmd.Parameters.Append cmd.CreateParameter("@pstrAux", _
        adChar, adParamInput, 4, prsMessage("Aux"))
    
    'PTR 4057 -- set Mnemonic size
    cmd.Parameters.Append cmd.CreateParameter("@pstrMnem", _
        adVarChar + adEmpty, adParamInput, 5, _
        prsMessage("Mnem"))
    
    cmd.Parameters.Append cmd.CreateParameter("@pstrDelimiter", _
        adChar, adParamInput, 1, prsMessage("Delimiter"))
    
    'PTR 4057 -- set Body size
    cmd.Parameters.Append cmd.CreateParameter("@pstrBody", _
        adLongVarWChar, adParamInput, lngBodySize, _
        prsMessage("Body"))
    
    'Execute without a resulting recordset and pull out the "return value" parameter
    cmd.Execute , , ADODB.adExecuteNoRecords
    WriteUnknownMessage = cmd.Parameters("@intUnkMsgID").Value
    
    Call LogEvent(svcEventInformation, svcMessageInfo, _
        "Unknown Message or queue not open. UnkMsgID=" & _
        CStr(cmd.Parameters("@intUnkMsgID").Value))
    
    'Construct an e-mail message with the contents
    strMessage = "Message written to unknown table. Message details: " & vbCrLf & _
        "Date:" & prsMessage("MsgDate") & vbCrLf
    If Not IsNull(prsMessage("QueueLabel")) Then
        strMessage = strMessage & "Label: " & prsMessage("QueueLabel") & vbCrLf
    Else
        strMessage = strMessage & "Label: " & vbCrLf
    End If
    strMessage = strMessage & "Originating ID: " & prsMessage("OrigID") & vbCrLf & _
    "Aux: " & prsMessage("Aux") & vbCrLf & _
    "Destination Mnemonic: " & prsMessage("Mnem") & vbCrLf & _
    "Message Body: " & prsMessage("Body")
    'Send the e-mail
    SendMail mudtParam.strEmail, strMessage
    
    Set cmd = Nothing
    Exit Function
Err_WriteUnknownMessage:
    If Err.Number = 3704 Then
        If Not OpenConnection Then
            'Couldn't open
            Call LogEvent(svcEventError, svcMessageError, "[" & Err.Number & "] " & _
                Err.Description & " [Connection to database was lost and could not be re-established.] ")
        End If
    Else
        Call LogEvent(svcEventError, svcMessageError, _
            "[" & Err.Number & "] " & Err.Description & _
            " [WriteUnknownMessage] ")
    End If
    Resume Next
End Function





'Purpose:   Handler to account for cases where the queue service
'           is disabled w/out notifiying the MSSGateway.
'
Private Sub QueueException()

    On Error Resume Next
    Dim strErrMsg As String
    Dim rsMessage As ADODB.Recordset
    Dim lngReturn As Long
    
    Set rsMessage = New ADODB.Recordset
    
    'if here, we have a message that failed to be written
    'to a queue. Close the queues, re-open, and attempt to
    'write. If successful and we were writing to a rcv queue,
    'send the ack and clear the temp string.
    
    'If successful and
    'we were writing to the Tx queue, clear the temp. stuff and
    'proceed with the re-init.
    
    'If not successful, start the timer. when the timer fires,
    'it'll send us here and we can retry
    
    If mstrQueueMsg <> "" Then
        Set rsMessage = MsgFromMSS(mstrQueueMsg)
    End If
    
    mlngQueueRetryCount = mlngQueueRetryCount + 1
    If mlngQueueRetryCount <= mudtParam.lngQueueRetryCount Then
        'close the queues
        mblnProcessQueue = False
        CloseQueues
        'attempt to re-open
        If OpenQueues = True Then
            If mstrQueueMsg = "" Then
                mlngQueueRetryCount = 0
                Call ProcessQueue
                Exit Sub
            End If
            
            'attempt to write to the queue
            If QueueWrite(rsMessage("QueueLabel"), rsMessage) = True Then
                'success, now what?
                mlngQueueRetryCount = 0
                tmrHeartbeat.Enabled = False
                mstrQueueMsg = ""
                If rsMessage("QueueLabel") = mudtParam.strQTx Then
                    'Tx queue
                    'MSS failed to ACK a message, attempt to
                    're-connect.
                    mstrSent = ""
                    ReInitializeSocket
                Else
                    'Rx queue
                    'we were able to save the message to the appropriate
                    'queue, so clear the buffer, ACK the switch and continue.
                    mstrBuffer = ""
                    Call SendData(Chr(ACK))
                    Call ProcessQueue
                End If
                Exit Sub
            Else
                'do nothing. If QueueWrite fails, it will call us again.
                Exit Sub
            End If
            
        Else
            strErrMsg = "Failed to re-open queues (attempt " & CStr(mlngQueueRetryCount) & _
            "). Will try " & CStr(mudtParam.lngQueueRetryCount) & " times."
            'Unable to re-open within QueueRetryCount times so
            'log event, bail
            Call LogEvent(svcEventError, svcMessageError, strErrMsg)
        
            'Send an e-mail if we can
            SendMail mudtParam.strEmail, strErrMsg
        
            'still unable to open, so start the timer
            tmrQueueRetry.Enabled = True
            Exit Sub
        End If
    Else
        tmrQueueRetry.Enabled = False
        tmrHeartbeat.Enabled = False
        mblnProcessQueue = False
        'Write to the unknown table
        If mstrQueueMsg <> "" Then
            lngReturn = WriteUnknownMessage(rsMessage)
        End If
        
        strErrMsg = "Unable to re-open queues. Tried " & _
            CStr(mudtParam.lngQueueRetryCount) & " times. Service is stopping."
        'Unable to re-open within QueueRetryCount times so
        'log event, notify, bail
        Call LogEvent(svcEventError, svcMessageError, strErrMsg)
        
        'Send an e-mail if we can
        SendMail mudtParam.strEmail, strErrMsg
        
        'close the socket
        ctlSocket1.Close
        
        'stop the service
        tmrUnload.Interval = 8000
        tmrUnload.Enabled = True
        Exit Sub
    End If
End Sub
