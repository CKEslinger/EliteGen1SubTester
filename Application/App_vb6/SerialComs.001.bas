Attribute VB_Name = "SerialComs"
Option Explicit

Const MAX_NUM_RECEIVED_LINES As Integer = 6

Public Const strCOM4LogPath As String = ".\"
Public Const strCOM4LogFName As String = "COM4_Log"
Public Const strCOM4LogFExt As String = "txt"

Public Const REG_MODE_AUTO As String = "0"
Public Const REG_MODE_MANUAL As String = "1"

Public Const BM_REG_MODE_SET As Byte = 1
Public Const BM_REG_MODE_VERIFIED As Byte = 2
Public Const BM_SELECTED_OPERATOR_SET As Byte = 4
Public Const BM_SELECTED_OPERATOR_VERIFIED As Byte = 8

Public m_bDevPOST_OK As Boolean

Private Const CME_ERR_527_DELAY As Double = 10
'Private Const MAX_WAIT_FOR_AT_RESPONSE As Integer = 31000
Private Const MAX_WAIT_FOR_AT_RESPONSE As Integer = 5000

Private m_AddCmd As String
Private m_bKillTimers As Boolean
Private m_szPSReceive As String
Private m_szATCommand As String
Private m_iParameter As Integer
Private m_szParameter As String
Private m_szUnsolicitated As String
Private m_aszReceivedLines(MAX_NUM_RECEIVED_LINES) As String
Private m_iReceivedIndex As Integer
Private m_bPSOutput As Boolean
Private m_fhNumEliteComPortLogFile As Integer
Private m_inputStream As String
Private m_szDevPOST_OK_Message As String
Private m_bRcvUnsollicited As Boolean
Public m_szDevPOST_Message As String

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
      dwFileAttributes As Long
      ftCreationTime As FILETIME
      ftLastAccessTime As FILETIME
      ftLastWriteTime As FILETIME
      nFileSizeHigh As Long
      nFileSizeLow As Long
      dwReserved0 As Long
      dwReserved1 As Long
      cFileName As String * 260
      cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" _
          Alias "FindFirstFileA" (ByVal lpFileName As String, _
          lpFindFileData As WIN32_FIND_DATA) As Long


Public Property Get PSOutput() As Boolean

    PSOutput = m_bPSOutput

End Property

Public Property Let PSOutput(bTurnPowerOn As Boolean)

    m_bPSOutput = bTurnPowerOn
    If bTurnPowerOn Then
        TransmitToPS ePS_TURN_ON
        '   Start the Power Up timer going in the main tester.
        g_fMainTester.tmrWaitPowerUp.Enabled = True
    Else
        '   Ensure that anything waiting for a response from the
        '   device is put out of its misery.
        
        g_eElite_Status = eSTAT_ERROR
        TransmitToPS ePS_TURN_OFF
        g_bOpenATRunning = False
        '   Ensure that the Power Up timer in the main tester is turned off.
        g_fMainTester.tmrWaitPowerUp.Enabled = False
    End If
    
    While g_ePS_COM_Status = eSTAT_RUNNING
        DoEvents
    Wend
    
    If g_ePS_COM_Status <> eSTAT_SUCCESS Then
        If bTurnPowerOn Then
            MsgBox "Failed to turn Power Supply ON"
        Else
            MsgBox "Failed to turn Power Supply OFF"
        End If
    End If


End Property

Public Sub ComPortToPowerSupplyEventProcessor()

    Dim cCR As Byte
    Dim cLF As Byte

    '   Something has happened so stop the time-out (give the COM port the
    '   benefit of the doubt).
    g_fMainTester.tmrPSCom.Enabled = False
    Select Case g_fMainTester.ComPortToPowerSupply.CommEvent
        Case comEvSend
            cCR = 0
            cLF = 0
            Select Case g_ePS_State
                Case ePS_VERIFY
                    '   Nothing to do here but wait for the response.

                Case ePS_CHECK_FOR_ERROR
                    '   Nothing to do here but wait for the response.

                Case ePS_GET_ERROR
                    '   Nothing to do here but wait for the response.

                Case ePS_CLEAR_ERROR
                    g_ePS_COM_Status = eSTAT_ERROR
                
                Case ePS_INIT_1
                    '   There is no response back from the Power Supply in
                    '   response to the command that was sent out.  Just go
                    '   straight to the next command in this sequence.
                    TransmitToPS ePS_INIT_2

                Case ePS_INIT_2
                    '   See comment above
                    TransmitToPS ePS_INIT_3

                Case ePS_INIT_3
                    '   See comment above
                    TransmitToPS ePS_INIT_4

                Case ePS_INIT_4
                    '   See comment above
                    TransmitToPS ePS_INIT_5

                Case ePS_INIT_5
                    '   See comment above
                    TransmitToPS ePS_INIT_6

                Case ePS_INIT_6
                    '   See comment above
                    TransmitToPS ePS_INIT_7

                Case ePS_INIT_7
                    '   See comment above
                    TransmitToPS ePS_INIT_8

                Case ePS_INIT_8
                    '   Power Supply has been initialized
                    g_ePS_COM_Status = eSTAT_SUCCESS

                Case ePS_TURN_ON
                    '   Power Supply has been turned ON
                    g_ePS_COM_Status = eSTAT_SUCCESS

                Case ePS_TURN_OFF
                    '   Power Supply has been turned OFF
                    g_ePS_COM_Status = eSTAT_SUCCESS

                Case ePS_MEAS_CURRENT
                    '   Nothing to do here but wait for the response.

                Case ePS_MEAS_VOLTS
                    '   Nothing to do here but wait for the response.

                Case Else
                    MsgBox "Invalid state after message sent to PS: (" & g_ePS_State & ")"
                    g_ePS_COM_Status = eSTAT_ERROR

            End Select

        Case comEvReceive
            '   Go get the character that has just come in from the Power Supply
            '   and add it to the receive buffer.
            While g_fMainTester.ComPortToPowerSupply.InBufferCount > 0
                cCR = cLF
                cLF = Asc(g_fMainTester.ComPortToPowerSupply.Input)
                If cCR = &HD And cLF = &HA Then
                    '   A complete response has been received from the PS.  Go Step
                    '   to the next PS state.
                    Select Case g_ePS_State
                        Case ePS_NULL
                            '       Just ignore whatever comes back

                        Case ePS_VERIFY
                            g_ePS_State = ePS_NULL
                            If m_szPSReceive = "1997.0" Then
                                g_ePS_COM_Status = eSTAT_SUCCESS
                            Else
                                g_ePS_COM_Status = eSTAT_ERROR
                            End If

                        Case ePS_CHECK_FOR_ERROR
                            If m_szPSReceive = "" Or "+0" Then
                                g_ePS_COM_Status = eSTAT_SUCCESS
                            Else
                                TransmitToPS ePS_GET_ERROR
                            End If

                        Case ePS_GET_ERROR
                            MsgBox "Power Supply Error: " & m_szPSReceive
                            TransmitToPS ePS_CLEAR_ERROR

                        Case ePS_CLEAR_ERROR
                            '   Just ignore whatever comes back

                        Case ePS_MEAS_CURRENT
                            g_ePS_COM_Status = eSTAT_SUCCESS

                        Case ePS_MEAS_VOLTS
                            g_ePS_COM_Status = eSTAT_SUCCESS

                        Case Else
                            MsgBox "An illegal state (" & g_ePS_State & ") in the PS Com state machine has been reached" & vbCrLf & vbCrLf & _
                                   "  -Received: '" & m_szPSReceive & "'"
                            g_ePS_COM_Status = eSTAT_ERROR

                    End Select
                    If g_ePS_State <> ePS_NULL Then
                    End If

                ElseIf cLF = &HD And (g_ePS_State = ePS_MEAS_CURRENT Or g_ePS_State = ePS_MEAS_VOLTS) Then
                    '   Some Agilent Power Supplies do not return the <CR><LF> pair, just a <CR>.
                    g_ePS_COM_Status = eSTAT_SUCCESS
                ElseIf cLF <> &HD And cLF <> &HA Then
                    '   Add on whatever was received to the PS receive string.
                    m_szPSReceive = m_szPSReceive + Chr(cLF)
                End If
            Wend

        Case Else
            '   Ignore all other events from the COM port

    End Select
    If g_ePS_COM_Status = eSTAT_RUNNING Then
        '   OK, what ever has happened, things haven't finished yet so start
        '   the timeout going again.
        g_fMainTester.tmrPSCom.Enabled = True
    End If

End Sub

Public Sub ComPortToEliteEventProcessor()

    Dim i As Integer
    Dim cChar As Byte
    Dim sCME_ErrorStr As String
    Dim eCME_ErrorNum As etCME_ERR

    On Error Resume Next

    '   Something has happened so stop the time-out (give the COM port the
    '   benefit of the doubt).
    g_fMainTester.tmrEliteCom.Enabled = False
    Select Case g_fMainTester.ComPortToElite.CommEvent
        Case comEvSend
            '   There is no special processing that is required once a command
            '   is sent, other than resetting a few control variables and just
            '   sitting here and waiting.
            cChar = 0
'            m_iReceivedIndex = 0
'            m_aszReceivedLines(m_iReceivedIndex) = ""

        Case comEvReceive
            '   Go get the character that has just come in from the device
            '   and add it to the receive buffer.
            
            While g_fMainTester.ComPortToElite.InBufferCount > 0
                cChar = 0
                cChar = Asc(g_fMainTester.ComPortToElite.Input)
                '   If the current state is eElite_NULL then just ignore this
                '   data.  It is simply unsolicited debug stream data coming
                '   from the Spectrum application.
                If g_eElite_State <> eElite_NULL Then
                    If cChar = &H0 Or cChar = &HD Or cChar = &HA Then
                        '   Strip out all control characters from the input
                        '   stream.
                        If "" <> m_inputStream Then
                            LogEliteComPortEventsToFile (m_inputStream)
                            m_inputStream = ""
                        End If
                        If m_aszReceivedLines(m_iReceivedIndex) <> "" Then
                            '   We have just received a line of some sorts.
                            '   Check to see if it is the line we expect (from m_szStateResponses)
                            If _
                                InStr(m_aszReceivedLines(m_iReceivedIndex), m_szStateResponses(g_eElite_State)) > 0 _
                                Or _
                                m_aszReceivedLines(m_iReceivedIndex) = "OK" > 0 _
                                Or _
                                InStr(1, m_aszReceivedLines(m_iReceivedIndex), "+CME ERROR:") > 0 _
                            Then
                                ParseResponse
                                If False = m_bRcvUnsollicited Then
                                    g_eElite_State = eElite_NULL
                                Else
                                    m_bRcvUnsollicited = False
                                End If
                            ElseIf m_iReceivedIndex >= 5 Then
                                If Not InStr(m_aszReceivedLines(m_iReceivedIndex - 5), "$SERVER[1-5]") > 0 Then
                                    '   This ugly jump has to be put in because
                                    '   the above three conditionals can't be
                                    '   concatenated into one expression with
                                    '   an And between them because (unlike C)
                                    '   Visual Basic goes and calculates ALL
                                    '   terms of an expression regardless if
                                    '   leading terms have already invalidated
                                    '   the entire expression.  Thus we get
                                    '   forced down this rabbit hole if the
                                    '   second expression is False, thus
                                    '   requiring the following 'ugly jump' out
                                    '   of the hole.
                                    GoTo UglyJump
                                End If
                                ParseResponse
                                g_eElite_State = eElite_NULL
                            Else
UglyJump:
                                '   Add the received line to the queue.  If the
                                '   queue is full then start shifting lines
                                '   down in the queue getting rid of line that
                                '   was received MAX_NUM_RECEIVED_LINES ago.
                                If m_iReceivedIndex < MAX_NUM_RECEIVED_LINES Then
                                    m_iReceivedIndex = m_iReceivedIndex + 1
                                Else
                                    '   Start flushing out lines
                                    For i = 0 To MAX_NUM_RECEIVED_LINES - 1
                                        m_aszReceivedLines(i) = m_aszReceivedLines(i + 1)
                                    Next i
                                End If
                                '   Prepare the next spot in the queue to
                                '   receive the next line from the device
                                m_aszReceivedLines(m_iReceivedIndex) = ""
                            End If
                        End If
                    Else
                        m_inputStream = m_inputStream & Chr(cChar)
                        m_aszReceivedLines(m_iReceivedIndex) = m_aszReceivedLines(m_iReceivedIndex) + Chr(cChar)
                    End If
                Else
                    ' Save the unsolicitated data stream
                    If &HD = cChar Or &HA = cChar Then
                        If "" <> m_inputStream Then
                            LogEliteComPortEventsToFile (m_inputStream)
                            If False = m_bDevPOST_OK And "" <> m_szDevPOST_OK_Message Then
                                If InStr(m_inputStream, m_szDevPOST_OK_Message) > 0 Then
                                    ' Device POST message found in unsolicitated data stream
                                    m_szDevPOST_Message = m_inputStream
                                    m_bDevPOST_OK = True
                                End If
                            End If
                            m_inputStream = ""
                        End If
                        m_szUnsolicitated = ""
                    Else
                        m_inputStream = m_inputStream & Chr(cChar)
                        m_szUnsolicitated = m_szUnsolicitated + Chr(cChar)
                    End If
                End If
            Wend

        Case comEventBreak ' Check for Break character
            ' Break character was hosing something with the MSComm control or
            ' the driver.  It stopped reception of new data causing a timeout error.
            ' If port is open, wait a couple of seconds before proceeding
            If True = g_fMainTester.ComPortToElite.PortOpen Then
                Delay (2#)
            End If
            
        Case Else
            '   Ignore all other events from the COM port

    End Select
    If g_eElite_Status = eSTAT_RUNNING Then
        '   OK, what ever has happened, things haven't finished yet so start
        '   the timeout going again.
        g_fMainTester.tmrEliteCom.Enabled = True
    End If

End Sub

Private Function ParseResponse() As Boolean
    Dim iMyIndex As Integer
    Dim iComma As Integer
    Dim iCommaPrev As Integer
    Dim i As Integer

    ' If command needs to be aborted then we want the event processor to not set
    ' the state to null until the response of the abort command is processed. By
    ' setting the ParseResponse return value to false the event processor will
    ' know that the previous command has been aborted by the current command.
    ParseResponse = True
    
    If InStr(m_aszReceivedLines(m_iReceivedIndex), "+CME ERROR: 3") > 0 Then
        g_eElite_Status = eSTAT_CME_3
    ElseIf InStr(m_aszReceivedLines(m_iReceivedIndex), "+CME ERROR: 515") > 0 Then
        g_eElite_Status = eSTAT_CME_515
        g_fMainTester.LogMessage (g_eElite_State & ": System is busy.")
    ElseIf m_aszReceivedLines(m_iReceivedIndex) = "ERROR" Then
        '   Parse out the response based on the issued
        '   command.

        g_fMainTester.LogMessage (g_eElite_State & ": Comd=" & m_szATCommand)
        g_fMainTester.LogMessage (g_eElite_State & ": Parm=" & m_szParameter & "," & m_iParameter)
        g_fMainTester.LogMessage (g_eElite_State & ": (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
        g_fMainTester.LogMessage (g_eElite_State & ": 'ERROR' returned from OpenAT")
        g_eElite_Status = eSTAT_ERROR
    Else
        Select Case g_eElite_State
            Case eElite_NULL
                '   Do nothing.  This happens when the
                '   first response line is not the
                '   command that was sent.

            Case eElite_verify, eElite_CHECK_BUZZER, eElite_SET_STATION_AUTOMATIC, eElite_Elite_Rly0, eElite_Elite_Rly1, eElite_ADDX
                If m_iReceivedIndex < 1 Then
                    g_fMainTester.LogMessage (g_eElite_State & ": Index=" & m_iReceivedIndex & " (supposed to >=1)")
                    g_fMainTester.LogMessage (g_eElite_State & ": Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage (g_eElite_State & ": Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage (g_eElite_State & ": (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                ElseIf m_aszReceivedLines(m_iReceivedIndex - 1) = m_szATCommand And _
                   m_aszReceivedLines(m_iReceivedIndex) = "OK" Then
                    g_eElite_Status = eSTAT_SUCCESS
                Else
                    g_fMainTester.LogMessage (g_eElite_State & ": Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage (g_eElite_State & ": Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage (g_eElite_State & ": (-1)=" & m_aszReceivedLines(m_iReceivedIndex - 1))
                    g_fMainTester.LogMessage (g_eElite_State & ": (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                End If

            
            Case eElite_Grn_LED, eElite_Red_LED, eElite_Grn_LED_Off, eElite_Red_LED_Off, eElite_GPS_INT_ANT, eElite_GPS_EXT_ANT
                If m_iReceivedIndex > 0 Then
                    If InStr(m_aszReceivedLines(m_iReceivedIndex), m_szStateResponses(g_eElite_State)) > 0 Then
                        g_eElite_Status = eSTAT_SUCCESS
                    Else
                        g_eElite_Status = eSTAT_ERROR
                    End If
                Else
                    g_eElite_Status = eSTAT_ERROR
                End If

            Case eElite_GET_IMEI, eElite_GET_IMSI, eElite_GET_CCID, eElite_GET_VOLTS, eElite_SET_SERIAL_NUM, eElite_GET_SERIAL_NUM, eElite_LOCK_GPS, _
                 eElite_SET_APN, eElite_APPVER, eElite_GET_PORT, eElite_SET_PORT, eElite_Elite_Rly2, eElite_RF_RCVR, eElite_ADD, _
                 eElite_SET_SMS_ADDRESS, eElite_SET_SMS_RESET_MIN, eElite_SET_SMS_MODE, eElite_GET_SMS_MODE, eElite_ANTITHEFT, _
                 eElite_CHECK_IGN, eElite_CHECK_STARTER, eElite_GET_SIGNAL_STRENGTH, eElite_GET_GPS_LOCATION
                If m_iReceivedIndex < 2 Then
                    g_fMainTester.LogMessage (g_eElite_State & ": Index=" & m_iReceivedIndex & " (supposed to >=2)")
                    g_fMainTester.LogMessage (g_eElite_State & ": Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage (g_eElite_State & ": Parm=" & m_szParameter & "," & m_iParameter)
                    If m_iReceivedIndex = 1 Then
                        g_fMainTester.LogMessage (g_eElite_State & ": (-1)=" & m_aszReceivedLines(m_iReceivedIndex - 1))
                    End If
                    g_fMainTester.LogMessage (g_eElite_State & ": (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                ElseIf m_aszReceivedLines(m_iReceivedIndex - 2) = m_szATCommand And _
                   m_aszReceivedLines(m_iReceivedIndex) = "OK" Then
                    m_szParameter = m_aszReceivedLines(m_iReceivedIndex - 1)
                    g_eElite_Status = eSTAT_SUCCESS
                Else
                    g_fMainTester.LogMessage (g_eElite_State & ": Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage (g_eElite_State & ": Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage (g_eElite_State & ": (-2)=" & m_aszReceivedLines(m_iReceivedIndex - 2))
                    g_fMainTester.LogMessage (g_eElite_State & ": (-1)=" & m_aszReceivedLines(m_iReceivedIndex - 1))
                    g_fMainTester.LogMessage (g_eElite_State & ": (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                End If


            Case eElite_SET_STATION_ID_FORMAT
                If m_iReceivedIndex < 1 Then
                    g_fMainTester.LogMessage ("COPS=3,2: Index=" & m_iReceivedIndex & " (supposed to >=1)")
                    g_fMainTester.LogMessage ("COPS=3,2: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                ElseIf (m_aszReceivedLines(m_iReceivedIndex - 1) = m_szATCommand) _
                        And _
                       (m_aszReceivedLines(m_iReceivedIndex) = m_szStateResponses(g_eElite_State)) Then
                    m_szParameter = m_aszReceivedLines(m_iReceivedIndex)
                    g_eElite_Status = eSTAT_SUCCESS
                Else
                    g_fMainTester.LogMessage ("COPS=3,2: Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage ("COPS=3,2: Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage ("COPS=3,2: (-1)=" & m_aszReceivedLines(m_iReceivedIndex - 1))
                    g_fMainTester.LogMessage ("COPS=3,2: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                End If


            Case eElite_SET_MANUAL_REG_MODE_AND_SELECT_OPERATOR
                If m_iReceivedIndex < 1 Then
                    g_fMainTester.LogMessage ("COPS=1,2,?: Index=" & m_iReceivedIndex & " (supposed to >=1)")
                    g_fMainTester.LogMessage ("COPS=1,2,?: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                ElseIf (m_aszReceivedLines(m_iReceivedIndex - 1) = m_szATCommand) _
                        And _
                       (m_aszReceivedLines(m_iReceivedIndex) = m_szStateResponses(g_eElite_State)) Then
                    m_szParameter = m_aszReceivedLines(m_iReceivedIndex)
                    g_eElite_Status = eSTAT_SUCCESS
                Else
                    If 0 = InStr(m_aszReceivedLines(m_iReceivedIndex), "+CME ERROR: ") Then
                        g_fMainTester.LogMessage ("COPS=1,2,?: Comd=" & m_szATCommand)
                        g_fMainTester.LogMessage ("COPS=1,2,?: Parm=" & m_szParameter & "," & m_iParameter)
                        g_fMainTester.LogMessage ("COPS=1,2,?: (-1)=" & m_aszReceivedLines(m_iReceivedIndex - 1))
                        g_fMainTester.LogMessage ("COPS=1,2,?: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    End If
                    g_eElite_Status = eSTAT_ERROR
                End If

            Case eElite_ABORT_SMS_SS_PLMN_SELECT
                If m_iReceivedIndex < 1 Then
                    g_fMainTester.LogMessage ("WAC: Index=" & m_iReceivedIndex & " (supposed to >=1)")
                    g_fMainTester.LogMessage ("WAC: Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage ("WAC: Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage ("WAC: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                ElseIf (m_aszReceivedLines(m_iReceivedIndex - 1) = m_szATCommand) _
                        And _
                       (m_aszReceivedLines(m_iReceivedIndex) = m_szStateResponses(g_eElite_State)) Then
                    m_szParameter = m_aszReceivedLines(m_iReceivedIndex)
                    g_eElite_Status = eSTAT_SUCCESS
                Else
                    g_fMainTester.LogMessage ("WAC: Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage ("WAC: Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage ("WAC: (-1)=" & m_aszReceivedLines(m_iReceivedIndex - 1))
                    g_fMainTester.LogMessage ("WAC: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                End If

            Case eElite_GET_GSM_REG_STATUS
                If m_iReceivedIndex < 1 Then
                    g_fMainTester.LogMessage ("CREG?: Index=" & m_iReceivedIndex & " (supposed to >=1)")
                    g_fMainTester.LogMessage ("CREG?: Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage ("CREG?: Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage ("CREG?: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                ElseIf (m_aszReceivedLines(m_iReceivedIndex - 1) = m_szATCommand) _
                        And _
                       (1 = InStr(m_aszReceivedLines(m_iReceivedIndex), m_szStateResponses(g_eElite_State))) Then
                    m_szParameter = m_aszReceivedLines(m_iReceivedIndex)
                    g_eElite_Status = eSTAT_SUCCESS
                Else
                    g_fMainTester.LogMessage ("CREG?: Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage ("CREG?: Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage ("CREG?: (-1)=" & m_aszReceivedLines(m_iReceivedIndex - 1))
                    g_fMainTester.LogMessage ("CREG?: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                End If

            Case eElite_SET_GSM_MODEM_FLOW_CTRL, eElite_SET_GSM_MODEM_BAUD_RATE, eElite_CLR_GSM_MODEM_BAUD_RATE
                If m_iReceivedIndex < 1 Then
                    g_fMainTester.LogMessage ("g_eElite_State: Index=" & m_iReceivedIndex & " (supposed to >=1)")
                    g_fMainTester.LogMessage ("g_eElite_State: Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage ("g_eElite_State: Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage ("g_eElite_State: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                ElseIf (m_aszReceivedLines(m_iReceivedIndex - 1) = m_szATCommand) _
                        And _
                       (1 = InStr(m_aszReceivedLines(m_iReceivedIndex), m_szStateResponses(g_eElite_State))) Then
                    m_szParameter = m_aszReceivedLines(m_iReceivedIndex)
                    g_eElite_Status = eSTAT_SUCCESS
                Else
                    g_fMainTester.LogMessage ("g_eElite_State: Comd=" & m_szATCommand)
                    g_fMainTester.LogMessage ("g_eElite_State: Parm=" & m_szParameter & "," & m_iParameter)
                    g_fMainTester.LogMessage ("g_eElite_State: (-1)=" & m_aszReceivedLines(m_iReceivedIndex - 1))
                    g_fMainTester.LogMessage ("g_eElite_State: (-0)=" & m_aszReceivedLines(m_iReceivedIndex))
                    g_eElite_Status = eSTAT_ERROR
                End If

            Case eElite_GET_APN

                iComma = 0
                For i = 0 To m_iParameter
                    iComma = InStr(iComma + 1, m_aszReceivedLines(m_iReceivedIndex - 1), ";")
                Next i
                
                m_szParameter = Right(m_aszReceivedLines(m_iReceivedIndex - 1), Len(m_aszReceivedLines(m_iReceivedIndex - 1)) - iComma)
                iComma = InStr(1, m_szParameter, ",")
                m_szParameter = Left(m_szParameter, iComma - 1)
                g_eElite_Status = eSTAT_SUCCESS
            

            Case eElite_GET_SERVER_IP
                iCommaPrev = 0
                iComma = 0
                For i = 0 To m_iParameter - 1
                iCommaPrev = iComma
                iComma = InStr(iCommaPrev + 1, m_aszReceivedLines(m_iReceivedIndex - 1), ";")
                Next i
                
                m_szParameter = Mid(m_aszReceivedLines(m_iReceivedIndex - 1), iCommaPrev + 1, iComma - iCommaPrev - 1)
                g_eElite_Status = eSTAT_SUCCESS
        

            Case eElite_SET_SERVER_IP
                If Mid(m_aszReceivedLines(m_iReceivedIndex - 1), 1, 1) = m_iParameter Then
                    g_eElite_Status = eSTAT_SUCCESS
                Else
                    g_eElite_Status = eSTAT_ERROR
                End If

            Case eElite_GET_SMS_REPLY_ADDR, eElite_MODEM_STATUS, eElite_SAVE_FLASH_PARAMS
                m_szParameter = m_aszReceivedLines(m_iReceivedIndex - 1)
                g_eElite_Status = eSTAT_SUCCESS
            
            Case eElite_MODEM_ID
                If "" = m_szParameter Then
                    m_szParameter = m_aszReceivedLines(m_iReceivedIndex - 1)
                End If
                g_eElite_Status = eSTAT_SUCCESS
            
            Case eElite_SET_LOW_POWER
                g_eElite_Status = eSTAT_ERROR
                If m_iReceivedIndex > 1 And _
                   InStr(m_aszReceivedLines(m_iReceivedIndex - 1), m_szParameter & ",") > 0 Then
                        g_eElite_Status = eSTAT_SUCCESS
                End If

            Case eelite_verboseoff
                g_eElite_Status = eSTAT_ERROR
                If m_iReceivedIndex > 0 And _
                   InStr(m_aszReceivedLines(m_iReceivedIndex), m_szStateResponses(g_eElite_State)) > 0 Then
                        m_szParameter = m_aszReceivedLines(m_iReceivedIndex)
                        g_eElite_Status = eSTAT_SUCCESS
                End If

            Case etELITE_STATES.eElite_ENABLE_BYPASS, etELITE_STATES.eElite_ENABLE_BYPASS
                g_fMainTester.LogMessage ("Non AT command was got by parser!")
                g_eElite_Status = eSTAT_ERROR
            
            Case etELITE_STATES.eElite_INIT_COMPLETE
                g_eElite_Status = eSTAT_SUCCESS
            
            Case Else
'                MsgBox "An illegal state (" & g_eElite_State & ") in the Elite Com state machine has been reached" & vbCrLf & vbCrLf & _
'                       "  -Received: '" & m_aszReceivedLines(m_iReceivedIndex) & "'"
                g_fMainTester.LogMessage ("An illegal state (" & g_eElite_State & ") in the Elite Com state machine has been reached" & vbCrLf & vbCrLf & _
                       "  -Received: '" & m_aszReceivedLines(m_iReceivedIndex) & "'")
                g_eElite_Status = eSTAT_ERROR
        End Select
    End If

End Function

Public Sub Init(Optional szPostMsg As String = "")

    Dim i As Integer

    m_szDevPOST_OK_Message = szPostMsg
    m_bDevPOST_OK = False
    m_bKillTimers = False
    m_iReceivedIndex = 0
    For i = 0 To MAX_NUM_RECEIVED_LINES
        m_aszReceivedLines(i) = ""
    Next i

End Sub

Public Sub Destroy()

    g_ePS_COM_Status = eSTAT_SHUTTING_DOWN
    g_eElite_Status = eSTAT_SHUTTING_DOWN
    m_bKillTimers = True
    g_fMainTester.tmrPSCom.Enabled = False
    g_fMainTester.tmrPSCom.Interval = 0
    g_fMainTester.tmrEliteCom.Enabled = False
    g_fMainTester.tmrEliteCom.Interval = 0
    
    If m_bPSOutput Then
        '   Turn the Power Supply OFF
        TransmitToPS ePS_TURN_OFF
        While g_ePS_COM_Status = eSTAT_RUNNING
            DoEvents
        Wend
        If g_ePS_COM_Status <> eSTAT_SUCCESS Then
            MsgBox "Failed to turn the Power Supply OFF"
        End If
    End If

End Sub

'   This function transmits a string to the Power Supply and starts the PS COM
'   state machine going to handle all of the events that will during this COM
'   session.
Public Sub TransmitToPS(eState As etPS_STATES)

    g_ePS_COM_Status = eSTAT_RUNNING
    g_ePS_State = eState
    m_szPSReceive = ""
    '   Start the timer going
    g_fMainTester.tmrPSCom.Enabled = True
    If g_fMainTester.ComPortToPowerSupply.PortOpen = True Then
        Select Case g_ePS_State
            Case ePS_NULL
                '   Ignore this state

            Case ePS_VERIFY
                g_fMainTester.ComPortToPowerSupply.Output = "SYST:VERS?" & vbCrLf

            Case ePS_CHECK_FOR_ERROR
                g_fMainTester.ComPortToPowerSupply.Output = "*ESR?" & vbCrLf

            Case ePS_GET_ERROR
                g_fMainTester.ComPortToPowerSupply.Output = "SYST:ERR?" & vbCrLf

            Case ePS_CLEAR_ERROR
                g_fMainTester.ComPortToPowerSupply.Output = "*CLS" & vbCrLf

            Case ePS_INIT_1
                g_fMainTester.ComPortToPowerSupply.Output = "*RST" & vbCrLf

            Case ePS_INIT_2
                g_fMainTester.ComPortToPowerSupply.Output = "INST:NSEL 1" & vbCrLf

            Case ePS_INIT_3
                g_fMainTester.ComPortToPowerSupply.Output = "VOLT:RANG HIGH" & vbCrLf

            Case ePS_INIT_4
                g_fMainTester.ComPortToPowerSupply.Output = "APPL 12.0,0.5" & vbCrLf

            Case ePS_INIT_5
                g_fMainTester.ComPortToPowerSupply.Output = "INST:NSEL 2" & vbCrLf

            Case ePS_INIT_6
                g_fMainTester.ComPortToPowerSupply.Output = "VOLT:RANG HIGH" & vbCrLf

            Case ePS_INIT_7
                g_fMainTester.ComPortToPowerSupply.Output = "APPL 9.0,0.1" & vbCrLf

            Case ePS_INIT_8
                g_fMainTester.ComPortToPowerSupply.Output = "INST:NSEL 1" & vbCrLf

            Case ePS_TURN_ON
                g_fMainTester.ComPortToPowerSupply.Output = "OUTP ON" & vbCrLf

            Case ePS_TURN_OFF
                g_fMainTester.ComPortToPowerSupply.Output = "OUTP OFF" & vbCrLf

            Case ePS_MEAS_CURRENT
                g_fMainTester.ComPortToPowerSupply.Output = "MEAS:CURR?" & vbCrLf

            Case ePS_MEAS_VOLTS
                g_fMainTester.ComPortToPowerSupply.Output = "MEAS:VOLT?" & vbCrLf

            Case Else
                MsgBox "PS transmit, illegal state (" & g_ePS_State & ")"
                g_ePS_State = ePS_NULL
        End Select
    Else
        g_ePS_COM_Status = eSTAT_NOT_CONNECTED
    End If
    If g_ePS_State = ePS_NULL Then
        '   Stop the timer, nothing got sent out
        g_fMainTester.tmrPSCom.Enabled = False
        g_ePS_COM_Status = eSTAT_SUCCESS
    End If

End Sub

'   This function transmits a string to the device and starts the Elite COM state
'   machine going to handle all of the events that will during this COM session.
Public Sub TransmitToElite(eState As etELITE_STATES, Optional lSerialComsTimeout As Long = SERIAL_COMS_TIMEOUT)

    g_eElite_Status = eSTAT_RUNNING
    g_eElite_State = eState
    
    If True = g_fMainTester.ComPortToElite.PortOpen Then
        Select Case g_eElite_State
            Case eElite_NULL
                '   Ignore this state

            Case eElite_ENABLE_BYPASS
                m_szATCommand = "BYPASS"
                g_fMainTester.ComPortToElite.Output = m_szATCommand & vbCr
                g_fMainTester.LogMessage ("Transmitted 'BYPASS' to device. Response not verifed")
                g_eElite_State = eElite_NULL
            
            Case eElite_DISABLE_BYPASS
                m_szATCommand = Chr(27)
                g_fMainTester.ComPortToElite.Output = m_szATCommand & vbCr
                g_fMainTester.LogMessage ("Transmitted 'Esc' character to device. Response not verifed")
                g_eElite_State = eElite_NULL

            Case eelite_verboseoff
                m_szATCommand = "Mf"
                If 1 = lSerialComsTimeout Then
                ' We don't want to verify the AT command
                ' First time this is exectued we could get unsollicited messages
                ' so need to execute a second time to verify command worked
                    g_fMainTester.ComPortToElite.Output = m_szATCommand & vbCr
                    g_fMainTester.LogMessage ("Transmitted '" & m_szATCommand & "' to device. Response not verifed")
                    g_eElite_State = eElite_NULL
                End If

            Case eElite_verify
                m_szATCommand = "AT"

            Case eElite_GET_IMEI
                m_szATCommand = "AT+CGSN"

            Case eElite_GET_IMSI
'                 m_szATCommand = "AT+CIMI"
                m_szATCommand = "M3"

            Case eElite_GET_CCID
'                m_szATCommand = "AT^SCID"
                m_szATCommand = "M6"

            Case eElite_LOCK_GPS
                m_szATCommand = "MS" & m_szParameter

            Case eElite_CHECK_IGN
                m_szATCommand = "MI"

            Case eElite_CHECK_STARTER
                m_szATCommand = "MH"

            Case eElite_CHECK_BUZZER
                m_szATCommand = "tt4"
                If 1 = lSerialComsTimeout Then
                ' We don't want to verify the AT command
                ' If the buzzer output is verified through hardware
                    g_fMainTester.ComPortToElite.Output = m_szATCommand & vbCr
                    g_fMainTester.LogMessage ("Transmitted 'tt4' to device. Response not verifed")
                    g_eElite_State = eElite_NULL
                End If

            Case eElite_SET_MANUAL_REG_MODE_AND_SELECT_OPERATOR
                m_szATCommand = "AT+COPS=1,2," & m_szParameter

            Case eElite_SET_STATION_ID_FORMAT
                m_szATCommand = "AT+COPS=3,2"

            Case eElite_GET_SIGNAL_STRENGTH
                m_szATCommand = "AT+CSQ"

            Case eElite_SET_STATION_AUTOMATIC
                m_szATCommand = "AT+COPS=0"

            Case eElite_ABORT_SMS_SS_PLMN_SELECT
                m_szATCommand = "AT+WAC"

            Case eElite_GET_GPS_LOCATION
                m_szATCommand = "AT$GPSLOC"

            Case eElite_GET_SERIAL_NUM
                m_szATCommand = "AT$SERIAL"

            Case eElite_SET_SERIAL_NUM
                m_szATCommand = "AT$SERIAL=" & m_szParameter

            Case eElite_GET_APN
                m_szATCommand = "AT$APN"

            Case eElite_SET_APN
                m_szATCommand = "AT$APN=" & m_szParameter

            Case eElite_GET_SERVER_IP
                m_szATCommand = "AT$SERVER?"

            Case eElite_SET_SERVER_IP
                m_szATCommand = "AT$SERVER=" & m_szParameter

            Case eElite_GET_PORT
                m_szATCommand = "AT$PORT"

            Case eElite_SET_PORT
                m_szATCommand = "AT$PORT=" & m_szParameter

            Case eElite_GET_VOLTS
                m_szATCommand = "AT$VOLT"
                              
            Case eElite_GET_SMS_MODE
                ' Get SMS state and "Hello packet timer" value
                m_szATCommand = "AT$SMSSTATE?"
                
            Case eElite_SET_SMS_MODE
                ' Set SMS state and "Hello packet timer" value
                m_szATCommand = "AT$SMSSTATE=" & m_szParameter
                
            Case eElite_SET_SMS_ADDRESS
                m_szATCommand = "AT$SMSREPLY=" & m_szParameter
                
            Case eElite_GET_SMS_REPLY_ADDR
                m_szATCommand = "AT$SMSREPLY"
                
            Case eElite_MODEM_ID
                m_szATCommand = "M2" & m_szParameter
                
            Case eElite_MODEM_STATUS
                m_szATCommand = "M?"
                
            Case eElite_SAVE_FLASH_PARAMS
                m_szATCommand = "Mw"
                
            Case eElite_SET_SMS_RESET_MIN
                m_szATCommand = "AT$RESETMIN=" & m_szParameter

            Case eElite_SET_LOW_POWER
                m_szATCommand = "AT$GPSMODE=" & m_szParameter
            
            Case eElite_GET_GSM_REG_STATUS
                m_szATCommand = "AT+CREG?"
                                      
            Case eElite_SET_GSM_MODEM_BAUD_RATE
                m_szATCommand = "AT+IPR=115200"
                                      
            Case eElite_SET_GSM_MODEM_FLOW_CTRL
                m_szATCommand = "AT\Q3"
                                      
            Case eElite_CLR_GSM_MODEM_BAUD_RATE
                m_szATCommand = "AT+IPR=0"
                                      
            ' AE added 8-25-09
            Case eElite_APPVER
                m_szATCommand = "AT$APVER"
                
            ' AE added 8-26-09
            Case eElite_Grn_LED
                m_szATCommand = "MLG1"
                
            ' AE added 8-26-09
            Case eElite_Red_LED
                m_szATCommand = "MLR1"
                
            ' AE added 8-26-09
            Case eElite_Grn_LED_Off
                m_szATCommand = "MLG0"
                
            ' AE added 8-26-09
            Case eElite_Red_LED_Off
                m_szATCommand = "MLR0"
                
            ' AE added 8-27-09
            Case eElite_Elite_Rly0
                m_szATCommand = "MR0"
                
            ' AE added 8-27-09
            Case eElite_Elite_Rly1
                m_szATCommand = "MR1"
                
            ' AE added 8-27-09
            Case eElite_Elite_Rly2
                m_szATCommand = "MR2"
                
            ' AE added 8-27-09
            Case eElite_RF_RCVR
                m_szATCommand = "MK"
                
            ' JB added 8-30-09
            Case eElite_GPS_INT_ANT
                m_szATCommand = "M4"
                
            ' JB added 8-30-09
            Case eElite_GPS_EXT_ANT
                m_szATCommand = "M5"
                
            ' AE added 8-27-09
            Case eElite_ADD
                m_szATCommand = m_AddCmd
                
            ' AE added 8-27-09
            Case eElite_ADDX
                m_szATCommand = m_AddCmd
                
            ' JMB modified 11-11-09
            Case eElite_ANTITHEFT
                m_szATCommand = "AT$PERMVALET=" & m_szParameter
            ' SFS added 2/17/2012
            Case 1000
                m_szATCommand = "AT+CREG?"
            Case Else
            ' MsgBox "Elite transmit, illegal state (" & g_eElite_State & ")"
                g_fMainTester.LogMessage ("Elite transmit, illegal state (" & g_eElite_State & ")")
                g_eElite_State = eElite_NULL
        End Select
        
        If g_eElite_State <> eElite_NULL Then
            g_fMainTester.LogMessage ("Transmitted '" & m_szATCommand & "' to device")
            m_iReceivedIndex = 0
            m_aszReceivedLines(m_iReceivedIndex) = ""
            
            '   Start the timer going
            ' If lSerialComsTimeout is zero, just restart the timer.
            ' NB: you better be sure you are in control of lSerialComsTimeout
            If Not (0 = lSerialComsTimeout) Then
                g_fMainTester.tmrEliteCom.Interval = lSerialComsTimeout
            End If
            g_fMainTester.tmrEliteCom.Enabled = True
            
            '   Send the AT command to the device
            g_fMainTester.ComPortToElite.Output = m_szATCommand & vbCr
        Else
            '   Stop the timer, nothing got sent out
            g_fMainTester.tmrEliteCom.Enabled = False
            g_eElite_Status = eSTAT_SUCCESS
        End If
    Else
        g_fMainTester.LogMessage (g_eElite_State & ": COM port not yet connected!")
        g_eElite_Status = eSTAT_NOT_CONNECTED
    End If

End Sub

Private Function CheckForPSError() As Boolean

    CheckForPSError = False
    TransmitToPS ePS_CHECK_FOR_ERROR
    While g_ePS_COM_Status = eSTAT_RUNNING
        DoEvents
    Wend
    If g_ePS_COM_Status = eSTAT_SUCCESS Or g_ePS_COM_Status = eSTAT_TIMEOUT Then
        '   If the Power Supply doesn't respond then it doesn't have any error
        '   messages for us (so this is a good thing).
        g_ePS_COM_Status = eSTAT_SUCCESS
        CheckForPSError = True
    Else
        '   Query the Power Supply for whatever error it has detected
        TransmitToPS ePS_GET_ERROR
        While g_ePS_COM_Status = eSTAT_RUNNING
            DoEvents
        Wend
    End If

End Function

Public Function InitPowerSupply() As Boolean

    InitPowerSupply = False
    While Not InitPowerSupply And Not g_bQuitProgram
        DoEvents
        TransmitToPS ePS_INIT_PS
        While g_ePS_COM_Status = eSTAT_RUNNING
            DoEvents
        Wend
        If g_ePS_COM_Status = eSTAT_SUCCESS Then
            '   Check for any errors in the Power Supply
            If CheckForPSError Then
                InitPowerSupply = True
            End If
        Else
            MsgBox "Failed to intialize the power supply: initialization phase = " & g_ePS_State - ePS_INIT_PS + 1 & "; status = " & g_ePS_COM_Status, vbCritical
            g_bQuitProgram = True
        End If
    Wend

End Function

Public Function PowerSupply1_Current() As Double
    TransmitToPS ePS_MEAS_CURRENT
    While g_ePS_COM_Status = eSTAT_RUNNING
        DoEvents
    Wend
    If g_ePS_COM_Status = eSTAT_SUCCESS Then
        PowerSupply1_Current = CDbl(m_szPSReceive)
    Else
        MsgBox "Error Receiving Measured Current From Power Supply" & vbCrLf & vbCrLf & _
               "Status = " & g_ePS_COM_Status & ", received = '" & m_szPSReceive, vbCritical
        PowerSupply1_Current = -9999
        TransmitToPS ePS_CLEAR_ERROR
        While g_ePS_COM_Status = eSTAT_RUNNING
            DoEvents
        Wend
    End If

End Function

Public Function PowerSupply1_Volts() As Double

    TransmitToPS ePS_MEAS_VOLTS
    While g_ePS_COM_Status = eSTAT_RUNNING
        DoEvents
    Wend
    If g_ePS_COM_Status = eSTAT_SUCCESS Then
        PowerSupply1_Volts = CDbl(m_szPSReceive)
    Else
        MsgBox "Error Receiving Measured Voltage From Power Supply" & vbCrLf & vbCrLf & _
               "Status = " & g_ePS_COM_Status & ", received = '" & m_szPSReceive, vbCritical
        PowerSupply1_Volts = -9999
        TransmitToPS ePS_CLEAR_ERROR
        While g_ePS_COM_Status = eSTAT_RUNNING
            DoEvents
        Wend
    End If

End Function

Public Function VerifyOpenAT() As Boolean
    ' This function replaces OpenATQuery.
    ' By handling CME error 515 (system busy error) at the
    ' event processor level, we can call eElite_VERIFY with a long
    ' timeout (20 seconds) and it will not fail if the system is
    ' just busy. It may be the case, however, that a previous
    ' commands response will finally be sent out from the device
    ' and cause our eElite_VERIFY to fail. To make sure that the
    ' eElite_VERIFY command really failed call it a second time
    ' with a short or default (3 second) timeout
    VerifyOpenAT = SendATCommand(eElite_verify, MAX_WAIT_FOR_AT_RESPONSE)
    If Not VerifyOpenAT Then
        Delay 1#
        VerifyOpenAT = SendATCommand(eElite_verify)
        If Not VerifyOpenAT Then
            g_fMainTester.LogMessage " Response: " & m_szParameter
        End If
    End If
End Function
   
Public Function SetSerialNum(lSerialNum As Long) As Boolean

    m_szParameter = Format(lSerialNum, "0000000") '& Chr(&H22)
    SetSerialNum = SendATCommand(eElite_SET_SERIAL_NUM)

End Function

Public Function SetAPN(iAPNnumber As Integer, szAPN As String, szPassword As String) As Boolean


    m_szParameter = Chr(&H22) & szAPN & Chr(&H22) & "," & Chr(&H22) & szPassword & Chr(&H22) & "," & iAPNnumber & "," & Chr(&H22) & Chr(&H22) & "," & Chr(&H22) & Chr(&H22)
    'm_szStateResponses(eElite_SET_APN) = "APN=" & iAPNnumber & ";"
    SetAPN = False
    If SendATCommand(eElite_SET_APN) Then
        If m_szParameter = iAPNnumber & "," & szAPN & ",," Then
            SetAPN = True
        Else
            g_fMainTester.LogMessage "Set APN failed, was expecting '" & iAPNnumber & "," & szAPN & ",," & "' but received '" & m_szParameter & "'"
        End If
    End If

End Function

Public Function SetServerIP(iServer As Integer, szIP As String) As Boolean

    m_iParameter = iServer
    m_szParameter = Chr(&H22) & Str(iServer) & Chr(&H22) & "," & Chr(&H22) & szIP & Chr(&H22)
    SetServerIP = SendATCommand(eElite_SET_SERVER_IP)

End Function

Public Function SetPort(szPort As String) As Boolean

    m_szParameter = Chr(&H22) & szPort & Chr(&H22)
    SetPort = SendATCommand(eElite_SET_PORT)

End Function



Public Function GetAPN(iAPNnumber As Integer, szAPNname As String) As Boolean

    'm_szStateResponses(eElite_GET_APN) = "APN " & iAPNnumber & ":"
    m_iParameter = iAPNnumber
    GetAPN = SendATCommand(eElite_GET_APN)
    If GetAPN Then
        szAPNname = m_szParameter
    End If

End Function

Public Function GetServerIP(iServer As Integer, szIP As String) As Boolean

    '   The Get Server IP command actually returns 5 IP strings (one for each
    '   server).  The Response Parser provides a special case for this command
    '   in that it picks which of the 5 returned strings to place into
    '   m_szParameter based on the input value placed in m_iParameter.
    m_iParameter = iServer
    GetServerIP = SendATCommand(eElite_GET_SERVER_IP)
    If GetServerIP Then
        szIP = m_szParameter
    End If

End Function

Public Function GetPort(szPort As String) As Boolean

    GetPort = SendATCommand(eElite_GET_PORT)
    If GetPort Then
        szPort = m_szParameter
    End If

End Function

Public Function GetVolts(iVolts As Integer) As Boolean

    GetVolts = SendATCommand(eElite_GET_VOLTS)
    If GetVolts Then
        iVolts = m_szParameter
    End If

End Function


Public Function GetSerialNum(szModemID As String) As Boolean

    GetSerialNum = SendATCommand(eElite_GET_SERIAL_NUM)
    If GetSerialNum Then
        szModemID = m_szParameter
    End If

End Function


Public Function CheckEliteRelay(bState As String, bInt As Integer) As Boolean

    'AE 8/26/09 check Elite relay
    Dim iRetries As Integer
    For iRetries = 1 To 3
    
    Select Case bInt
    Case 0
    'off
    CheckEliteRelay = SendATCommand(eElite_Elite_Rly0)
    Case 1
    'on
    CheckEliteRelay = SendATCommand(eElite_Elite_Rly1)
    Case 2
    'check
    CheckEliteRelay = SendATCommand(eElite_Elite_Rly2)
    End Select
    
    If CheckEliteRelay Then
        bState = m_szParameter
        Exit For
    End If
     
    Next iRetries


End Function

Public Function SetAntiTheft(Optional szEnable As String = "0") As Boolean

    'AE 9/4/09 SetAntiTheft
    'JMB Modified to accept "szEnable" parameter 11/11/09
    Dim iRetries As Integer
    
    SetAntiTheft = False
    ' TODO: Why does this need to retry
    For iRetries = 1 To 3
        m_szParameter = szEnable
    
        If True = SendATCommand(eElite_ANTITHEFT) Then
            SetAntiTheft = True
            Exit For
        End If

    Next iRetries

End Function

Public Function CheckIgnition(bState As String) As Boolean

    'AE 8/25/09 check ignition line
    Dim iRetries As Integer
    For iRetries = 1 To 3
    CheckIgnition = SendATCommand(eElite_CHECK_IGN)
    If CheckIgnition Then
        bState = m_szParameter
        Exit For
    End If
     
    Next iRetries
     
'    Dim iRetries As Integer
'
'    For iRetries = 1 To 3
'        If SendATCommand(eElite_CHECK_IGN) Then
'            CheckIgnition = True
'            If m_szParameter = "0" Then
'                '   Ignition relay is in the right position
'                bState = True
'                Exit For
'            Else
'                bState = False
'            End If
'        Else
'            CheckIgnition = False
'        End If
'        '   Hold off a bit before retrying the command again
'        Delay 0.5
'    Next iRetries

End Function

Public Function CheckStarter(bState As String) As Boolean

    'AE 8/25/09 check starter line
    Dim iRetries As Integer
    For iRetries = 1 To 3
    CheckStarter = SendATCommand(eElite_CHECK_STARTER)
    If CheckStarter Then
        bState = m_szParameter
        Exit For
    End If
    
    Next iRetries
    
'    Dim iRetries As Integer
'
'    For iRetries = 1 To 3
'        If SendATCommand(eElite_CHECK_STARTER) Then
'            CheckStarter = True
'            If m_szParameter = "0" Then
'                '   Starter relay is in the right position
'                bState = True
'                Exit For
'            Else
'                bState = False
'            End If
'        Else
'            CheckStarter = False
'        End If
'        '   Hold off a bit before retrying the command again
'        Delay 0.5
'    Next iRetries

End Function

Public Function CheckLED(iLED As Integer) As Boolean

    'AE 8/26/09 check LEDs
    Select Case iLED
    Case 0
        CheckLED = SendATCommand(eElite_Grn_LED)
        Delay (0.5)
    Case 1
        CheckLED = SendATCommand(eElite_Red_LED)
        Delay (0.5)
    Case 2
        CheckLED = SendATCommand(eElite_Grn_LED_Off)
    Case 3
        CheckLED = SendATCommand(eElite_Red_LED_Off)
    End Select
    
End Function

Public Function AddTest(szAdd As String, szTO As Long, szResult As String) As Boolean

    Dim iRetries As Integer

    'AE 8/28/09 for additional tests
    m_AddCmd = szAdd
    
    For iRetries = 1 To 3
    If szResult = "" Or szResult = "OK" Then
    AddTest = SendATCommand(eElite_ADDX, szTO)
        If AddTest Then
        g_szAddResult = "OK"
        Exit For
    End If
    Else
    AddTest = SendATCommand(eElite_ADD, szTO)
        If AddTest Then
        g_szAddResult = m_szParameter
        Exit For
    End If
    End If
    
    Next iRetries

End Function

Public Function GetRfRcvrInput(szRfRcvdInput As String) As Boolean

    'AE 8/27/09 get last input received by rf receiver on the dut
    GetRfRcvrInput = SendATCommand(eElite_RF_RCVR)
    If GetRfRcvrInput Then
        If Len(m_szParameter) > 1 Then
            szRfRcvdInput = Right(m_szParameter, 1)
        Else
            szRfRcvdInput = m_szParameter
        End If
    End If
End Function

Public Function GetAppVer(szVER As String) As Boolean
    'AE 8/25/09 check firmware version downloaded with filename in XML
    GetAppVer = SendATCommand(eElite_APPVER)
    If GetAppVer Then
        szVER = m_szParameter
    End If

End Function

Public Function GetIMEI(szIMEI As String) As Boolean

    '   Get the serial # of the modem
    GetIMEI = SendATCommand(eElite_GET_IMEI)
    If GetIMEI Then
        szIMEI = m_szParameter
    End If

End Function

Public Function GetIMSI(szIMSI As String) As Boolean

'AE 10/19/09 no longer suspend/resume
    GetIMSI = SendATCommand(eElite_GET_IMSI)
    If GetIMSI Then
        szIMSI = m_szParameter
    End If

End Function
    
Public Function GetCCID(szCCID As String) As Boolean

'AE 10/21/09 no longer suspend/resume
    GetCCID = SendATCommand(eElite_GET_CCID)
    If GetCCID Then
        szCCID = m_szParameter
    End If

End Function
    

Public Function FindGPS(szGPS As String) As Boolean
    
    'AE 10/21/09 New GPS method
    m_szParameter = szGPS
    szGPS = ""
    FindGPS = SendATCommand(eElite_LOCK_GPS)
    If FindGPS Then
        szGPS = m_szParameter
    End If

End Function

Public Function SetGpsAntenna(eGpsAnt As etELITE_STATES) As Boolean
    
    ' Select antennna specified by the eGpsAnt paramter.
    If eElite_GPS_INT_ANT = eGpsAnt Or eElite_GPS_EXT_ANT = eGpsAnt Then
        SetGpsAntenna = SendATCommand(eGpsAnt)
    Else
        ' Parameter can only be either eElite_GPS_INT_ANT or eElite_GPS_EXT_ANT
        SetGpsAntenna = False
    End If

End Function

Public Function VerifySignalStrengthGPRS(szMin As String, szMax As String, Optional bMainTester As Boolean = True) As Boolean
    ' Much more likely to correctly read AT command responses if application is stopped or suspended
    Dim iSigStrength As Integer
    Dim bAtCmdStatus As Boolean
    Dim eCME_ErrorNum As etCME_ERR
    
    VerifySignalStrengthGPRS = False

    ' Send AT command to device to read the signal strength
    iSigStrength = 0
    bAtCmdStatus = SendATCommand(eElite_GET_SIGNAL_STRENGTH)
    If bAtCmdStatus Then
        VerifySignalStrengthGPRS = True
        If InStr(1, m_szParameter, "+CSQ: ") = 1 Then
            ' Parse out the actual signal strength values
            iSigStrength = CInt(Val(Right(m_szParameter, Len(m_szParameter) - Len("+CSQ: "))))
            If iSigStrength < CInt(Val(szMin)) Or iSigStrength > CInt(Val(szMax)) Then
                g_fMainTester.LogMessage "Signal strength is bad, got " & iSigStrength & ", range = (" & szMin & "," & szMax & ")"
            Else
                ' GPRS signal strength is in expected range
                szMin = Str(iSigStrength)
                szMax = szMin
            End If
        Else
            g_fMainTester.LogMessage "Unexpected response to AT command."
        End If
    Else
        g_fMainTester.LogMessage "AT command to get signal strength failed."
    End If
    
End Function

Public Function FindSimulator(Optional bMainTester As Boolean = True) As Boolean
    ' This function works best if the OpenAT application is either stopped or suspended
    ' The GSM should also be restarted prior to calling this function, in most cases

    Dim bAtCmdStatus As Boolean
    Dim bSetCREG_Retries As Byte

    FindSimulator = False
    bSetCREG_Retries = 20
    
    Do
        ' Analyze device registration state
        bAtCmdStatus = SendATCommand(eElite_GET_GSM_REG_STATUS)
        If bAtCmdStatus Then
            ' AT Command succeeded now check the command response
            
            If Not InStr(1, m_szParameter, "+CREG: 2,5,") = 1 Then
                ' Not registered yet
                ' TODO: Make sure it is the cell site simulator. Returns +CREG: 5,"00FA","1010"
                If InStr(1, m_szParameter, "+CREG: 2,0") = 1 Then
                    ' Probably needs more time
                    
                    g_fMainTester.LogMessage (" Not registered. Not searching for new operator (" & m_szParameter & ").")
                     Delay 1#
                Else
                    If 1 = bSetCREG_Retries Then
                        ' Somebody forgot to push the "F2" (LOC UPD) button on the simulator or device wasn't ready
                        g_fMainTester.LogMessage (" Not registered. Searching for new operator (" & m_szParameter & ").")
                        ' TODO: MessageBox "Push 'F2' (LOC UPD) button on the cell site simulator
                        ' and then CFUN=1
                    Else
                        Delay 1#
                    End If
               End If
            Else
                ' Excellent! We found the cell site simulator and are registered in roaming mode
                g_fMainTester.LogMessage (" Registered with cell site simulator (" & m_szParameter & ").")
                FindSimulator = True
                Exit Do
            End If
        End If
        bSetCREG_Retries = bSetCREG_Retries - 1
    Loop Until 0 = bSetCREG_Retries
End Function

Public Function SetStationIdFormatDefault(Optional bMainTester As Boolean = True) As Boolean
    ' This function works best if the OpenAT application is either stopped or suspended
    
    ' Set the format for reading/selecting the operator
    ' TODO: Read station ID format and only set if necessary
    ' TODO: Handle eElite_SET_STATION_ID_FORMAT error condition
    If Not SendATCommand(eElite_SET_STATION_ID_FORMAT) Then
        If bMainTester Then
            g_fMainTester.LogMessage "Failed to set station ID format"
            Exit Function
        End If
    End If
End Function

Public Function SetModeAndOperator(szMode As String, Optional szPLMN As String = "00000", Optional bMainTester As Boolean = True) As Byte
    ' This function works best if the OpenAT application is either stopped or suspended
    Dim bAtCmdStatus As Boolean
    Dim bCME_ErrorRetry As Boolean
    Dim bSetPLMN_Retries As Byte

    bCME_ErrorRetry = False
    SetModeAndOperator = False
    
    bSetPLMN_Retries = 6
    Do
        If REG_MODE_MANUAL = szMode Then
            ' Try to select the cell site simulator as the operator and set the device to manual registration mode
            m_szParameter = szPLMN
            bAtCmdStatus = SendATCommand(eElite_SET_MANUAL_REG_MODE_AND_SELECT_OPERATOR, 3000)
        Else
            bAtCmdStatus = SendATCommand(eElite_SET_STATION_AUTOMATIC, 3000)
        End If
    
        ' Check to see if we got back a CME error. If not then eCME_ErrorNum will be eCME_ERR_FALSE
        Select Case (ProcessCME_Error(m_aszReceivedLines(m_iReceivedIndex)))
        Case eCME_ERR_FALSE
            If True = bAtCmdStatus Then
                ' It's a miracle the COPS command actually succeeded without a CME Error
                g_fMainTester.LogMessage (" Selected operator and registration mode set to manual.")
                SetModeAndOperator = True
                Exit Do
            Else
                ' AT Command failed. Error message should be handled by event processor/parser
                ' This is a little tricky. We will probably get a lot of 515 errors if we now issue
                ' a new command because the previous command is still processing so lets try and
                ' get back to a known good state by just issuing AT (verify) commands until either
                ' a long timeout expires or we get back the response from the last command.
                If Not SerialComs.VerifyOpenAT() Then
                    ' Give up.
                    g_fMainTester.LogMessage (" Unable to proceed. Previous command still processing.")
                    g_fMainTester.LogMessage " Aborting previous command..."
                    bAtCmdStatus = SendATCommand(eElite_verify)
                    If Not bAtCmdStatus Then
                        bAtCmdStatus = SendATCommand(eElite_ABORT_SMS_SS_PLMN_SELECT)
                        If Not bAtCmdStatus Then
                            SendATCommand (eElite_verify)
                        End If
                    End If
                    SetModeAndOperator = BM_REG_MODE_SET Or SetModeAndOperator
                    SetModeAndOperator = (Not BM_REG_MODE_VERIFIED) And SetModeAndOperator
                    If Not REG_MODE_AUTO = szMode Then
                        SetModeAndOperator = BM_SELECTED_OPERATOR_SET Or SetModeAndOperator
                        SetModeAndOperator = (Not BM_SELECTED_OPERATOR_VERIFIED) And SetModeAndOperator
                        g_fMainTester.LogMessage "Selected operator (PLMN) " & szPLMN & "."
                        g_fMainTester.LogMessage " New PLMN can't be verified until modem reset."
                    End If
                End If
                Exit Do
                ' Delay 1
            End If
        Case eCME_ERR_032
            g_fMainTester.LogMessage (" +CME ERROR: 32 is good. It means we connected to the cell site simulator")
            ' Try calling FindSimulator to give the tester time to finish registering
            ' bAtCmdStatus = SerialComs.FindSimulator()
            ' The above was commented out because it led to passing but with a very weak
            ' signal strength (Signal Strength = 20)
            Exit Do
        Case eCME_ERR_515
            ' This means that the previous command is still processing and that our command was never sent
            g_fMainTester.LogMessage (" Unable to execute command. Previous command still processing.")
            g_fMainTester.LogMessage (" This should have been handled by the Elite event processor and response parser.")
            Delay 1
        Case eCME_ERR_527
            ' This means that our command was unable to process and we need to try it again
            If 1 = bSetPLMN_Retries Then
                ' Somebody forgot to push the "F2" (LOC UPD) button on the simulator
                g_fMainTester.LogMessage (" Unable to execute command. Device GPRS is busy.")
            Else
                Delay CME_ERR_527_DELAY
            End If
        Case eCME_ERR_529
            ' Somebody forgot to push the "F2" (LOC UPD) button on the simulator
            Exit Do
        Case eCME_ERR_547
            ' Not quite sure what this means. Could mean the SIM card is not present or is not seated correctly
            ' g_fMainTester.LogMessage (" ")
            Delay 1
        Case eCME_ERR_UNKNOWN
            Exit Do
        Case Else
            ' It's a known error code but not handled in this case
            ' Exit Do
        End Select
                    
        bSetPLMN_Retries = bSetPLMN_Retries - 1
    Loop Until (0 = bSetPLMN_Retries)
End Function


Public Function GetSMSignalStrength(iStrength As Integer) As Boolean
    ' This function currently only used by tester diagnostics, not the actual testing
    Dim strStrength As String

    If g_eElite_State = eElite_NULL Then
        '   Get the GSM signal strength.
        GetSMSignalStrength = SendATCommand(eElite_GET_SIGNAL_STRENGTH)
        If GetSMSignalStrength Then
            If InStr(1, m_szParameter, "+CSQ: ") = 1 Then
                iStrength = CInt(Val(Right(m_szParameter, Len(m_szParameter) - Len("+CSQ: "))))
            End If
'        Else
'            GetSMSignalStrength = False
        End If
    Else
        GetSMSignalStrength = False
    End If

End Function
Public Function GetSMSmodeString() As String
    Dim iStrIdx As Integer
    GetSMSmodeString = ""
    
    If g_eElite_State = eElite_NULL Then
        '   send AT command to get the SMS state and "hello packet timer" value
        If SendATCommand(eElite_GET_SMS_MODE) Then
            ' Parse the return string
            'iStrIdx = InStr(1, m_szParameter, ":")
            'm_szParameter = Trim(Right(m_szParameter, Len(m_szParameter) - iStrIdx))
            GetSMSmodeString = m_szParameter
        End If
    End If

End Function
Public Function SendSMSmodeString(strMode As String) As Boolean

    SendSMSmodeString = False
    m_szParameter = strMode
    If g_eElite_State = eElite_NULL Then
        '   send AT command to set the SMS state and "hello packet timer" value
        SendSMSmodeString = SendATCommand(eElite_SET_SMS_MODE)
    End If

End Function
Public Function SendSMSaddressString(strAddress As String) As Boolean

'AE 9/1/09 set SMSaddress
    m_szParameter = strAddress
    SendSMSaddressString = SendATCommand(eElite_SET_SMS_ADDRESS)
    If SendSMSaddressString Then
        strAddress = m_szParameter
    End If

'    SendSMSaddressString = False
'    m_szParameter = strAddress
'    If g_eElite_State = eElite_NULL Then
'        '   send AT command to set the SMS address
'        SendSMSaddressString = SendATCommand(eElite_SET_SMS_ADDRESS)
'    End If

End Function
Public Function GetSMSReplyAddress(szSMSReplyAddr As String, Optional iAddrIdx As Integer = 0) As Boolean
    m_iParameter = iAddrIdx
    GetSMSReplyAddress = False
    GetSMSReplyAddress = SendATCommand(eElite_GET_SMS_REPLY_ADDR)
    If GetSMSReplyAddress Then
        szSMSReplyAddr = m_szParameter
    End If
End Function
Public Function SendResetMin(strResetMin As String) As Boolean

    SendResetMin = False
    m_szParameter = strResetMin
    If g_eElite_State = eElite_NULL Then
        '   send AT command to set the SMS reset minimum
        SendResetMin = SendATCommand(eElite_SET_SMS_RESET_MIN)
    End If

End Function

Public Function SendLowPowerCommandString(Optional szGPS_AlwaysOn As String = "0") As Boolean
    
    SendLowPowerCommandString = False
    m_szParameter = szGPS_AlwaysOn
    If g_eElite_State = eElite_NULL Then
        '   send AT command to turn off the "always-on GPS" feature
        SendLowPowerCommandString = SendATCommand(eElite_SET_LOW_POWER)
    End If

End Function


Public Function GetGPSCoordinates(lLongitude As Long, lLatitude As Long, iNoSatellites As Integer) As Boolean

    Dim szResponse As String

    GetGPSCoordinates = False
    On Error GoTo FormatError
    If SendATCommand(eElite_GET_GPS_LOCATION) Then
        '   Parse the coordinates and number of satellites out of the returned
        '   parameter
        szResponse = m_szParameter
        iNoSatellites = CInt(Val(Right(szResponse, Len(szResponse) - InStrRev(szResponse, ","))))
        szResponse = Left(szResponse, InStrRev(szResponse, ",") - 1)
        lLongitude = Val(Right(szResponse, Len(szResponse) - InStrRev(szResponse, ",")))
        szResponse = Left(szResponse, InStrRev(szResponse, ",") - 1)
        lLatitude = Right(szResponse, Len(szResponse) - InStrRev(szResponse, ","))
        GetGPSCoordinates = True
    End If
    Exit Function

FormatError:
    g_fMainTester.LogMessage "Format error in AT$GPSLOC response: '" & m_szParameter & "'"

End Function

Public Function SendATCommand(eATCommand As etELITE_STATES, Optional lSerialComsTimeout As Long = SERIAL_COMS_TIMEOUT) As Boolean

    Dim szMessage As String

    SendATCommand = False
    Do
        TransmitToElite eATCommand, lSerialComsTimeout
        While g_eElite_Status = eSTAT_RUNNING
            DoEvents
            If Not g_fIO.IsTesterClosed Then
                g_eElite_Status = eSTAT_ERROR
            End If
        Wend
        lSerialComsTimeout = 0
        Delay 0.25
    Loop Until ((eSTAT_TIMEOUT = g_eElite_Status) Or (Not (eSTAT_CME_515 = g_eElite_Status)))
    If g_eElite_Status = eSTAT_SUCCESS Or g_eElite_Status = eSTAT_DOWNLOADING Then
        '   The AT command was sent and executed successfully.
        SendATCommand = True
    Else
        szMessage = "Command " & eATCommand & " failed; status = "
        If g_eElite_Status = eSTAT_ERROR Then
            szMessage = szMessage + "Unrecognized response or unknown error"
        ElseIf g_eElite_Status = eSTAT_TIMEOUT Then
            szMessage = szMessage + "Timed out waiting for response"
        Else
            szMessage = szMessage + Str(g_eElite_Status)
        End If
        g_fMainTester.LogMessage szMessage
        SendATCommand = False
    End If

End Function

Public Function OpenEliteComPort(bOpenElite As Boolean, Optional bWarnOnError As Boolean = False) As Boolean
    OpenEliteComPort = False
    On Error GoTo OpenEliteComPort_Err
    
    If True = bOpenElite Then
        If 0 = m_fhNumEliteComPortLogFile Then
            ' Open COM port I/O logging file
            m_fhNumEliteComPortLogFile = OpenEliteComPortLogFile
        Else
            g_fMainTester.LogMessage ("Device COM Port log file already open")
        End If
        
        If False = g_fMainTester.ComPortToElite.PortOpen Then
            ' Open the COM port
            g_fMainTester.ComPortToElite.PortOpen = True
        Else
            g_fMainTester.LogMessage ("Device COM Port already open")
        End If
    Else
        If True = g_fMainTester.ComPortToElite.PortOpen Then
            ' Close the COM port
            g_fMainTester.ComPortToElite.PortOpen = False
        Else
            g_fMainTester.LogMessage ("Device COM Port already closed")
        End If
        
        If Not 0 = m_fhNumEliteComPortLogFile Then
            ' Close COM port I/O logging file
            Close #m_fhNumEliteComPortLogFile
            m_fhNumEliteComPortLogFile = 0
        Else
            g_fMainTester.LogMessage ("Device COM Port log file already closed")
        End If
    End If
    
    If bOpenElite = g_fMainTester.ComPortToElite.PortOpen Then
        OpenEliteComPort = True
    End If

    Exit Function

OpenEliteComPort_Err:
    If True = bWarnOnError Then
        MsgBox "Error opening COM port to Elilte II", vbExclamation
    Else
        If g_fMainTester Is Nothing Then
            Exit Function
        End If
        g_fMainTester.LogMessage ("Error opening COM port to Elilte II")
    End If

End Function

Public Function LogEliteComPortEventsToFile(StringToWrite As String)
    
    Dim ErrMsg As VbMsgBoxResult
    Dim fhNumber As Integer
    
    fhNumber = m_fhNumEliteComPortLogFile

    ' LogEliteComPortEventsToFile = False

    If Not fhNumber = 0 Then
        On Error GoTo FileWriteErrorCOM4:
        
        Print #fhNumber, StringToWrite
            
        ' LogEliteComPortEventsToFile = True
    End If
    Exit Function

FileWriteErrorCOM4:
    ErrMsg = MsgBox("Error writing to COM4 log file!", vbCritical)
    m_fhNumEliteComPortLogFile = 0

End Function

' Make backup copy of existing log file and
' delete existing log file if it exists
Public Function InitEliteComPortLogFile()
    ' If file exists, make a backup and delete the file
    Dim strLogPathAndFile As String
    Dim strLogBackupPathAndFile As String
    Dim ErrMsg As VbMsgBoxResult
    Dim lHandle As Long
    Dim WFD As WIN32_FIND_DATA

    strLogPathAndFile = strCOM4LogPath & strCOM4LogFName & "." & strCOM4LogFExt
    strLogBackupPathAndFile = strCOM4LogPath & strCOM4LogFName & ".bak"
    
    On Error GoTo FileInitErrorCOM4:
    
    lHandle = FindFirstFile(strLogPathAndFile, WFD)

    If lHandle > 0 Then
        ' TODO: Don't make zero length file backups
        FileCopy strLogPathAndFile, strLogBackupPathAndFile
        Kill strLogPathAndFile
    End If
    Exit Function

FileInitErrorCOM4:
    ErrMsg = MsgBox("Error initializing COM4 log file:" & vbCrLf & "   " & Err.Description, vbCritical)

End Function

Private Function OpenEliteComPortLogFile() As Integer

    Dim strLogPathAndFile As String
    Dim ErrMsg As VbMsgBoxResult
    Dim fhNumber As Integer

    OpenEliteComPortLogFile = 0

    On Error GoTo FileOpenErrorCOM4:
    
    fhNumber = FreeFile
    
    strLogPathAndFile = strCOM4LogPath & strCOM4LogFName & "." & strCOM4LogFExt
    Open strLogPathAndFile For Append Access Write As #fhNumber
    
    ' Return the file handle number
    OpenEliteComPortLogFile = fhNumber

    Exit Function
    
FileOpenErrorCOM4:
    ErrMsg = MsgBox("Error opening COM4 log file!", vbCritical)

End Function

Public Function WaitForPost(iWaitForPOST As Integer) As Boolean
    WaitForPost = False
    
    If False = g_fMainTester.ComPortToElite.PortOpen Then
        ' Open the COM port
        If False = OpenEliteComPort(True) Then
            g_fMainTester.LogMessage ("Error Opening COM port")
            Exit Function
        End If
    End If
    
    While False = m_bDevPOST_OK And 0 <> iWaitForPOST
        iWaitForPOST = iWaitForPOST - 1
        Delay (1#)
    Wend
    
    If False = m_bDevPOST_OK Then
        g_fMainTester.LogMessage ("Failed to receive message '" & m_szDevPOST_OK_Message & "' from device")
    Else
        m_bDevPOST_OK = False
        If True = ReceiveUnsollicitedMessage(etELITE_STATES.eElite_INIT_COMPLETE, 30000) Then
            WaitForPost = True
        End If
    End If
End Function

Public Function ReceiveUnsollicitedMessage(eState As etELITE_STATES, Optional lSerialComsTimeout As Long = SERIAL_COMS_TIMEOUT) As Boolean
    Dim szMessage As String
    Dim szSaveResponseString As String
    
    szSaveResponseString = m_szStateResponses(eElite_INIT_COMPLETE)
    ReceiveUnsollicitedMessage = False
    g_eElite_Status = eSTAT_RUNNING
    g_eElite_State = eState
    m_iReceivedIndex = 0
    m_aszReceivedLines(m_iReceivedIndex) = ""
    
    If False = g_fMainTester.ComPortToElite.PortOpen Then
        g_fMainTester.LogMessage (g_eElite_State & ": COM port not yet connected!")
        g_eElite_Status = eSTAT_NOT_CONNECTED
        Exit Function
    End If
    
    If etELITE_STATES.eElite_INIT_COMPLETE <> g_eElite_State Then
        g_fMainTester.LogMessage ("Elite receive, illegal state (" & g_eElite_State & ")")
        g_eElite_State = eElite_NULL
    End If
    
    If g_eElite_State <> eElite_NULL Then
'        g_fMainTester.tmrEliteCom.Enabled = False
        g_fMainTester.LogMessage ("Waiting for 'done initializing' message from device")
                
        If InStr(m_szDevPOST_Message, "AIRPLANE MODE") > 0 Then
            m_szStateResponses(eElite_INIT_COMPLETE) = "Modem is in airplane mode."
        End If

        m_iReceivedIndex = 0
        m_aszReceivedLines(m_iReceivedIndex) = ""
        
        DoEvents
        m_bRcvUnsollicited = True
        
        '   Start the timer going
        ' If lSerialComsTimeout is zero, just restart the timer.
        ' NB: you better be sure you are in control of lSerialComsTimeout
        If Not (0 = lSerialComsTimeout) Then
            g_fMainTester.tmrEliteCom.Interval = lSerialComsTimeout
        End If

        g_fMainTester.tmrEliteCom.Enabled = True
        While g_eElite_Status = eSTAT_RUNNING
            DoEvents
            If Not g_fIO.IsTesterClosed Then
                g_eElite_Status = eSTAT_ERROR
            End If
        Wend
        
        m_bRcvUnsollicited = False
        m_szStateResponses(eElite_INIT_COMPLETE) = szSaveResponseString
        
        If g_eElite_Status = eSTAT_SUCCESS Then
            '   The unsollicited message was received
            ReceiveUnsollicitedMessage = True
            szMessage = "Received message '" & m_szStateResponses(eState) & "' from device"
        Else
            szMessage = "Wait for 'done initializing' message failed: status = "
            If g_eElite_Status = eSTAT_ERROR Then
                szMessage = szMessage + "Unknown error"
            ElseIf g_eElite_Status = eSTAT_TIMEOUT Then
                szMessage = szMessage + "Timed out"
            Else
                szMessage = szMessage + Str(g_eElite_Status)
            End If
        End If
        g_fMainTester.LogMessage szMessage
    Else
        '   Stop the timer, nothing got sent out
        g_fMainTester.tmrEliteCom.Enabled = False
        g_eElite_Status = eSTAT_SUCCESS
    End If

End Function

Public Function ModemID(szModemID As String) As Boolean
    ModemID = False
    If "" = szModemID Then
        ' Read the modem ID
        m_szParameter = ""
        ModemID = SendATCommand(eElite_MODEM_ID)
        If ModemID Then
            szModemID = m_szParameter
        End If
    Else
        m_szParameter = "=" & szModemID
        ModemID = SendATCommand(eElite_MODEM_ID)
    End If
    
End Function

Public Function ModemStatus(szModemStatus As String) As Boolean
    ModemStatus = False
    m_szParameter = ""
    ModemStatus = SendATCommand(eElite_MODEM_STATUS)
    If ModemStatus Then
        szModemStatus = m_szParameter
    End If
End Function

Public Function SaveFlashParams() As Boolean
    SaveFlashParams = False
    SaveFlashParams = SendATCommand(eElite_SAVE_FLASH_PARAMS)
End Function

