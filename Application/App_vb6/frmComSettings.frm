VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmComSettings 
   Caption         =   "COM port settings"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrXMODEM 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6240
      Top             =   4080
   End
   Begin VB.Timer tmrPSCom 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5040
      Top             =   4080
   End
   Begin VB.Timer tmrWCCom 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5640
      Top             =   4080
   End
   Begin VB.Frame frRequired 
      ForeColor       =   &H000000C0&
      Height          =   1212
      Left            =   4440
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   3852
      Begin VB.Label lbRequired 
         Caption         =   "Selections required before Apply can be pressed!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   612
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   3372
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6600
      TabIndex        =   10
      Top             =   4560
      Width           =   1572
   End
   Begin VB.CommandButton cmdApplyWCSettings 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4440
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CommandButton cmdApplyPSSettings 
      Caption         =   "APPLY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4440
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Frame frComError 
      Height          =   2292
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8052
      Begin VB.Label lbComErrorMessage 
         Alignment       =   2  'Center
         Caption         =   "Failed to find Agilant Power Supply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   372
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   7572
      End
      Begin VB.Label lbComErrorDescription 
         Caption         =   $"frmComSettings.frx":0000
         Height          =   1092
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   7452
      End
   End
   Begin MSCommLib.MSComm ComPortToPowerSupply 
      Left            =   6840
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      StopBits        =   2
   End
   Begin MSCommLib.MSComm ComPortToWavecom 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin VB.Label lbRequired 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   252
      Index           =   4
      Left            =   2880
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label lbRequired 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   252
      Index           =   3
      Left            =   2880
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label lbRequired 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   252
      Index           =   2
      Left            =   3480
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label lbRequired 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   252
      Index           =   1
      Left            =   3480
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label lbRequired 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   252
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Label5 
      Caption         =   "Stop bits:"
      Height          =   252
      Left            =   960
      TabIndex        =   7
      Top             =   4680
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "Data bits:"
      Height          =   252
      Left            =   960
      TabIndex        =   6
      Top             =   4200
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "Parity:"
      Height          =   252
      Left            =   960
      TabIndex        =   5
      Top             =   3720
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Baud Rate:"
      Height          =   252
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "COM Port #:"
      Height          =   252
      Left            =   960
      TabIndex        =   3
      Top             =   2760
      Width           =   972
   End
End
Attribute VB_Name = "frmComSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_NUM_RECEIVED_LINES As Integer = 6

''''    These are the states for communication with the Power Supply
Private Enum etPS_STATES
    ePS_NULL                        '   0
    ePS_VERIFY                      '   1
    ePS_CHECK_FOR_ERROR             '   2
    ePS_GET_ERROR                   '   3
    ePS_CLEAR_ERROR                 '   4
    ePS_INIT_PS                     '   5
    ePS_INIT_1 = ePS_INIT_PS        '   5
    ePS_INIT_2                      '   6
    ePS_INIT_3                      '   7
    ePS_INIT_4                      '   8
    ePS_INIT_5                      '   9
    ePS_INIT_6                      '   10
    ePS_INIT_7                      '   11
    ePS_INIT_8                      '   12
    ePS_TURN_ON                     '   13
    ePS_TURN_OFF                    '   14
    ePS_MEAS_CURRENT                '   15
    ePS_MEAS_VOLTS                  '   16
    '   Any new PS state definitions must be added BEFORE this comment
    ePS_NUM_OF_STATES               '   17
End Enum

''''    These are the states for communication with the Wavecom Module
Private Enum etWC_STATES
    '   The assignment of eWC_NULL to ePS_NUM_OF_STATES ensures that there is
    '   no overlap between the values defined in these two steps (just in case
    '   one of the states gets passed off to the wrong serial link in the code.
    eWC_NULL = ePS_NUM_OF_STATES    '   17
    eWC_SUCCESS                     '   18
    eWC_ERROR                       '   19
    eWC_VERIFY                      '   20
    eWC_FW_VERSION                  '   21
    eWC_XMODEM_VERSION              '   22
    eWC_DOWNLOAD_APP                '   23
    eWC_RESTART_GSM                 '   24
    eWC_STOP_APP                    '   25
    eWC_ERASE_OBJECTS               '   26
    eWC_START_APP                   '   27
    eWC_GET_IMSI                    '   28, sn of the simm card
    eWC_GET_CCID                    '   29, prefered network serial number (AT+ Command)
    eWC_GET_IMEI                    '   30, serial # for Wavecomm modem
    eWC_LOCK_GPS                    '   31
    eWC_REGISTER_GPRS               '   32
    eWC_CHECK_IGN                   '   33
    eWC_CHECK_STARTER               '   34
    eWC_CHECK_PPAY                  '   35
    eWC_CHECK_BUZZER                '   36
    eWC_GET_SIGNAL_STRENGTH         '   37
    eWC_GET_GPS_LOCATION            '   38
    eWC_GET_FREQUENCY               '   39
    eWC_SET_FREQUENCY               '   40
    eWC_GET_DATA_STORAGE            '   41
    eWC_SET_DATA_STORAGE            '   42
    eWC_GET_MODEM_ID                '   43
    eWC_SET_MODEM_ID                '   44
    eWC_GET_APN                     '   45
    eWC_SET_APN                     '   46
    eWC_GET_SERVER_IP               '   47
    eWC_SET_SERVER_IP               '   48
    eWC_GET_PORT                    '   49
    eWC_SET_PORT                    '   50
    eWC_GET_APVER                   '   51
    eWC_RESET_STATE                 '   52
    eWC_GET_VOLTS                   '   53
End Enum

''''    These are the various status' of the COM links
Private Enum etCOM_STATUS
    eSTAT_SUCCESS
    eSTAT_RUNNING
    eSTAT_DOWNLOADING
    eSTAT_TIMEOUT
    eSTAT_NOT_CONNECTED
    eSTAT_ERROR
    eSTAT_SHUTTING_DOWN
End Enum

Private m_iApply As Integer
Private m_bKillTimers As Boolean
Private m_ePS_State As etPS_STATES
Private m_ePS_Status As etCOM_STATUS
Private m_eWC_State As etWC_STATES
Private m_eWC_Status As etCOM_STATUS
Private m_szPSReceive As String
Private m_szATCommand As String
Private m_iParameter As Integer
Private m_szParameter As String
Private m_szUnsolicitated As String
Private m_aszWCReceivedLines(MAX_NUM_RECEIVED_LINES) As String
Private m_iWCReceivedIndex As Integer
Private m_bPSOutput As Boolean

Public Property Get PSOutput() As Boolean

    PSOutput = m_bPSOutput

End Property

Public Property Let PSOutput(bTurnPowerOn As Boolean)

    m_bPSOutput = bTurnPowerOn
    If bTurnPowerOn Then
        TransmitToPS ePS_TURN_ON
        If Not g_fMainTester.txtTestNumber(0).Enabled Or _
           Not g_fMainTester.txtTestNumber(1).Enabled Then
            '   The power is being turned on to the DUT but the basic sound
            '   and current draw tests are not being executed.  We need to
            '   insert a delay here else any subsequent commands to the
            '   Wavecom unit might fail.
            Delay 4
        End If
    Else
        TransmitToPS ePS_TURN_OFF
    End If
    While m_ePS_Status = eSTAT_RUNNING
        DoEvents
    Wend
    If m_ePS_Status <> eSTAT_SUCCESS Then
        If bTurnPowerOn Then
            MsgBox "Failed to turn Power Suply ON"
        Else
            MsgBox "Failed to turn Power Suply OFF"
        End If
    End If

End Property

Public Sub ShowDAQ()

    lbComErrorMessage.Caption = "Failed to find DAQ card"
    lbComErrorDescription.Caption = "The Digital/Analog I/O card did not respond to its initialization messages.  Please ensure that the card has been installed correctly and that the appropriate driver software has been loaded successfully into the test PC."
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
'    ckApplySerialSettings.Visible = False
    Show vbModal

End Sub

Public Sub ShowPS()

    lbComErrorMessage.Caption = "Failed to find Agilant Power Supply"
    lbComErrorDescription.Caption = "The Agilant Power Supply did not respond to its hail message.  Please ensure that the power supply has been connected to the correct COM port on the test PC and that it has been powered up successfully." & vbCrLf & vbCrLf & "Adjust the COM port settings below if necessary."
'    cbxComPortPS.Visible = True
'    cbxBaudRatePS.Visible = True
'    cbxParityPS.Visible = True
'    cbxDataBitsPS.Visible = True
'    cbxStopBitsPS.Visible = True
'    cmdApplyPSSettings.Visible = True
    Show vbModal

End Sub

Public Sub ShowWC()

    lbComErrorMessage.Caption = "Failed to find Wavecom module"
    lbComErrorDescription.Caption = "The Wavecom module did not respond to its hail message.  Please ensure that the Elite device has been inserted into the test fixture correctly, that the serial cable to the test fixture is connected to the correct COM port on the test PC and that the powered supply is delivering the correct voltage to the DUT." & vbCrLf & vbCrLf & "Adjust the COM port settings below if necessary."
'    cbxComPortWC.Visible = True
'    cbxBaudRateWC.Visible = True
'    cbxParityWC.Visible = True
'    cbxDataBitsWC.Visible = True
'    cbxStopBitsWC.Visible = True
    cmdApplyWCSettings.Visible = True
    Show vbModal

End Sub

'Private Sub cbxBaudRatePS_Change()

'    g_szComBaudPS = cbxBaudRatePS.List(cbxBaudRatePS.ListIndex)

'End Sub

'Private Sub cbxBaudRateWC_Change()

'    g_szComBaudWC = cbxBaudRateWC.List(cbxBaudRatePS.ListIndex)

'End Sub

'Private Sub cbxComPortPS_Change()

'    g_iComPortPS = cbxComPortPS.ListIndex + 1

'End Sub

'Private Sub cbxComPortWC_Change()

'    g_iComPortWC = cbxComPortWC.ListIndex + 1

'End Sub

'Private Sub cbxDataBitsPS_Change()

'    g_szComDataBitsPS = cbxDataBitsPS.List(cbxDataBitsPS.ListIndex)

'End Sub

'Private Sub cbxDataBitsWC_Change()

'    g_szComDataBitsWC = cbxDataBitsWC.List(cbxDataBitsPS.ListIndex)

'End Sub

'Private Sub cbxParityPS_Change()

'    g_szComParityPS = cbxParityPS.List(cbxParityPS.ListIndex)

'End Sub

'Private Sub cbxParityWC_Change()

'    g_szComParityWC = cbxParityWC.List(cbxParityPS.ListIndex)

'End Sub

'Private Sub cbxStopBitsPS_Change()

'    g_szComStopBitsPS = cbxStopBitsPS.List(cbxStopBitsPS.ListIndex)

'End Sub

'Private Sub cbxStopBitsWC_Change()

'    g_szComStopBitsWC = cbxStopBitsWC.List(cbxStopBitsPS.ListIndex)

'End Sub

'Private Sub ClearRequired()

'    frRequired.Visible = False
'    lbRequired(0).Visible = False
'    lbRequired(1).Visible = False
'    lbRequired(2).Visible = False
'    lbRequired(3).Visible = False
'    lbRequired(4).Visible = False

'End Sub

'Private Sub cmdApplyPSSettings_Click()

'    ClearRequired
'    g_iComPortPS = cbxComPortPS.ListIndex + 1
'    g_szComBaudPS = cbxBaudRatePS.List(cbxBaudRatePS.ListIndex)
'    g_szComParityPS = cbxParityPS.List(cbxParityPS.ListIndex)
'    g_szComDataBitsPS = cbxDataBitsPS.List(cbxDataBitsPS.ListIndex)
'    g_szComStopBitsPS = cbxStopBitsPS.List(cbxStopBitsPS.ListIndex)
    '   Sanity check
'    If g_iComPortPS = 0 Then
'        lbRequired(0).Visible = True
'        frRequired.Visible = True
'    End If
'    If g_szComBaudPS = "" Then
'        lbRequired(1).Visible = True
'        frRequired.Visible = True
'    End If
'    If g_szComParityPS = "" Then
'        lbRequired(2).Visible = True
'        frRequired.Visible = True
'    End If
'    If g_szComDataBitsPS = "" Then
'        lbRequired(3).Visible = True
'        frRequired.Visible = True
'    End If
'    If g_szComStopBitsPS = "" Then
'        lbRequired(4).Visible = True
'        frRequired.Visible = True
'    End If
'    If Not frRequired.Visible Then
'        cbxComPortPS.Visible = False
'        cbxBaudRatePS.Visible = False
'        cbxParityPS.Visible = False
'        cbxDataBitsPS.Visible = False
'        cbxStopBitsPS.Visible = False
'        cmdApplyPSSettings.Visible = False
'        Hide
'    End If

'End Sub

'Private Sub cmdApplyWCSettings_Click()

'    ClearRequired
'    g_iComPortWC = cbxComPortWC.ListIndex + 1
'    g_szComBaudWC = cbxBaudRateWC.List(cbxBaudRateWC.ListIndex)
'    g_szComParityWC = cbxParityWC.List(cbxParityWC.ListIndex)
'    g_szComDataBitsWC = cbxDataBitsWC.List(cbxDataBitsWC.ListIndex)
'    g_szComStopBitsWC = cbxStopBitsWC.List(cbxStopBitsWC.ListIndex)
    '   Sanity check
'    If g_iComPortWC = 0 Then
'        lbRequired(0).Visible = True
'        frRequired.Visible = True
'    End If
'    If g_szComBaudWC = "" Then
'        lbRequired(1).Visible = True
'        frRequired.Visible = True
'    End If
'    If g_szComParityWC = "" Then
'        lbRequired(2).Visible = True
'        frRequired.Visible = True
'    End If
'    If g_szComDataBitsWC = "" Then
'        lbRequired(3).Visible = True
'        frRequired.Visible = True
'    End If
'    If g_szComStopBitsWC = "" Then
'        lbRequired(4).Visible = True
'        frRequired.Visible = True
'    End If
'    If Not frRequired.Visible Then
'        cbxComPortWC.Visible = False
'        cbxBaudRateWC.Visible = False
'        cbxParityWC.Visible = False
'        cbxDataBitsWC.Visible = False
'        cbxStopBitsWC.Visible = False
'        cmdApplyWCSettings.Visible = False
'        Hide
'    End If

'End Sub

Private Sub cmdQuit_Click()

    g_bQuitProgram = True
    Hide

End Sub

Private Sub ComPortToPowerSupply_OnComm()

    Dim cCR As Byte
    Dim cLF As Byte

    '   Something has happened so stop the time-out (give the COM port the
    '   benefit of the doubt).
    tmrPSCom.Enabled = False
    Select Case ComPortToPowerSupply.CommEvent
        Case comEvSend
            cCR = 0
            cLF = 0
            Select Case m_ePS_State
                Case ePS_VERIFY
                    '   Nothing to do here but wait for the response.

                Case ePS_CHECK_FOR_ERROR
                    '   Nothing to do here but wait for the response.

                Case ePS_GET_ERROR
                    '   Nothing to do here but wait for the response.

                Case ePS_CLEAR_ERROR
                    m_ePS_Status = eSTAT_ERROR
                
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
                    m_ePS_Status = eSTAT_SUCCESS

                Case ePS_TURN_ON
                    '   Power Supply has been turned ON
                    m_ePS_Status = eSTAT_SUCCESS

                Case ePS_TURN_OFF
                    '   Power Supply has been turned ON
                    m_ePS_Status = eSTAT_SUCCESS

                Case ePS_MEAS_CURRENT
                    '   Nothing to do here but wait for the response.

                Case ePS_MEAS_VOLTS
                    '   Nothing to do here but wait for the response.

                Case Else
                    MsgBox "Invalid state after message sent to PS: (" & m_ePS_State & ")"
                    m_ePS_Status = eSTAT_ERROR

            End Select

        Case comEvReceive
            '   Go get the character that has just come in from the Power Supply
            '   and add it to the receive buffer.
            While ComPortToPowerSupply.InBufferCount > 0
                cCR = cLF
                cLF = Asc(ComPortToPowerSupply.Input)
                If cCR = &HD And cLF = &HA Then
                    '   A complete response has been received from the PS.  Go Step
                    '   to the next PS state.
                    Select Case m_ePS_State
                        Case ePS_NULL
                            '       Just ignore whatever comes back

                        Case ePS_VERIFY
                            m_ePS_State = ePS_NULL
                            If m_szPSReceive = "1997.0" Then
                                m_ePS_Status = eSTAT_SUCCESS
                            Else
                                m_ePS_Status = eSTAT_ERROR
                            End If

                        Case ePS_CHECK_FOR_ERROR
                            If m_szPSReceive = "" Or "+0" Then
                                m_ePS_Status = eSTAT_SUCCESS
                            Else
                                TransmitToPS ePS_GET_ERROR
                            End If

                        Case ePS_GET_ERROR
                            MsgBox "Power Supply Error: " & m_szPSReceive
                            TransmitToPS ePS_CLEAR_ERROR

                        Case ePS_CLEAR_ERROR
                            '   Just ignore whatever comes back

                        Case ePS_MEAS_CURRENT
                            m_ePS_Status = eSTAT_SUCCESS

                        Case ePS_MEAS_VOLTS
                            m_ePS_Status = eSTAT_SUCCESS

                        Case Else
                            MsgBox "An illegal state (" & m_ePS_State & ") in the PS Com state machine has been reached" & vbCrLf & vbCrLf & _
                                   "  -Received: '" & m_szPSReceive & "'"
                            m_ePS_Status = eSTAT_ERROR

                    End Select
                    If m_ePS_State <> ePS_NULL Then
                    End If

                ElseIf cLF = &HD And (m_ePS_State = ePS_MEAS_CURRENT Or m_ePS_State = ePS_MEAS_VOLTS) Then
                    '   Some Agilent Power Supplies do not return the <CR><LF> pair, just a <CR>.
                    m_ePS_Status = eSTAT_SUCCESS
                ElseIf cLF <> &HD And cLF <> &HA Then
                    '   Add on whatever was received to the PS receive string.
                    m_szPSReceive = m_szPSReceive + Chr(cLF)
                End If
            Wend

        Case Else
            '   Ignore all other events from the COM port

    End Select
    If m_ePS_Status = eSTAT_RUNNING Then
        '   OK, what ever has happened, things haven't finished yet so start
        '   the timeout going again.
        tmrPSCom.Enabled = True
    End If

End Sub

Private Sub ComPortToWavecom_OnComm()

    Dim i As Integer
    Dim cChar As Byte

    '   Something has happened so stop the time-out (give the COM port the
    '   benefit of the doubt).
    tmrWCCom.Enabled = False
    Select Case ComPortToWavecom.CommEvent
        Case comEvSend
            '   There is no special processing that is required once a command
            '   is sent, other than resetting a few control variables and just
            '   sitting here and waiting.
            cChar = 0
            m_iWCReceivedIndex = 0
            m_aszWCReceivedLines(m_iWCReceivedIndex) = ""

        Case comEvReceive
            '   Go get the character that has just come in from the Wavecom
            '   module and add it to the receive buffer.
            While ComPortToWavecom.InBufferCount > 0
                cChar = Asc(ComPortToWavecom.Input)
                '   If the current state is eWC_NULL then just ignore this
                '   data.  It is simply unsolicited debug stream data coming
                '   from the Spectrum application.
                If m_eWC_State <> eWC_NULL Then
                    If cChar = &H0 Or cChar = &HD Or cChar = &HA Then
                        '   Strip out all control characters from the input
                        '   stream.
                        If m_aszWCReceivedLines(m_iWCReceivedIndex) <> "" Then
                            '   We have just received a line of some sorts.
                            '   Check to see if it is the OK line.
                            If InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "+WDWL: 0") > 0 Then
                                i = 0
                            End If
                            If m_aszWCReceivedLines(m_iWCReceivedIndex) = "+WIND:  3" Then
                                If g_bOpenATRunning Then
                                    g_bOpenATRebooted = True
                                Else
                                    g_bOpenATRunning = True
                                    g_bOpenATRebooted = False
                                End If
                                m_aszWCReceivedLines(m_iWCReceivedIndex) = ""
                            ElseIf m_aszWCReceivedLines(m_iWCReceivedIndex) = "OK" Or _
                                   m_aszWCReceivedLines(m_iWCReceivedIndex) = "ERROR" Or _
                                   InStr(1, m_aszWCReceivedLines(m_iWCReceivedIndex), "+CME ERROR:") > 0 Or _
                                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "+WDWL: 0") > 0 Or _
                                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$PPAY 1234") > 0 Or _
                                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$SERVER ") > 0 Or _
                                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "ModemId is:") > 0 Or _
                                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$APN ") > 0 Or _
                                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$PORT ") > 0 Or _
                                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "Version is: ") > 0 Or _
                                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "Sending Reset State: ") > 0 Then
                                ParseResponse
                                m_eWC_State = eWC_NULL
                            ElseIf m_iWCReceivedIndex >= 5 Then
                                If Not InStr(m_aszWCReceivedLines(m_iWCReceivedIndex - 5), "$SERVER[1-5]") > 0 Then
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
                                m_eWC_State = eWC_NULL
                            Else
UglyJump:
                                '   Add the received line to the queue.  If the
                                '   queue is full then start shifting lines
                                '   down in the queue getting rid of line that
                                '   was received MAX_NUM_RECEIVED_LINES ago.
                                If m_iWCReceivedIndex < MAX_NUM_RECEIVED_LINES Then
                                    m_iWCReceivedIndex = m_iWCReceivedIndex + 1
                                Else
                                    '   Start flushing out lines
                                    For i = 0 To MAX_NUM_RECEIVED_LINES - 1
                                        m_aszWCReceivedLines(i) = m_aszWCReceivedLines(i + 1)
                                    Next i
                                End If
                                '   Prepare the next spot in the queue to
                                '   receive the next line from the Wavecom
                                '   module.
                                m_aszWCReceivedLines(m_iWCReceivedIndex) = ""
                            End If
                        End If
                    Else
                        m_aszWCReceivedLines(m_iWCReceivedIndex) = m_aszWCReceivedLines(m_iWCReceivedIndex) + Chr(cChar)
                    End If
                Else
                    '   Check the unsolicitated data stream for the "+WIND:  3" message
                    If cChar = &H0 Or cChar = &HD Or cChar = &HA Then
                        If m_szUnsolicitated = "+WIND:  3" Then
                            If g_bOpenATRunning Then
                                g_bOpenATRebooted = True
                            Else
                                g_bOpenATRunning = True
                                g_bOpenATRebooted = False
                            End If
                        Else
                            '   Just ignore this unsolicitated message from the Wavecom unit
                            m_szUnsolicitated = ""
                        End If
                    Else
                        m_szUnsolicitated = m_szUnsolicitated + Chr(cChar)
                    End If
                End If
            Wend
        
        Case Else
            '   Ignore all other events from the COM port

    End Select
    If m_eWC_Status = eSTAT_RUNNING Then
        '   OK, what ever has happened, things haven't finished yet so start
        '   the timeout going again.
        tmrWCCom.Enabled = True
    End If

End Sub

Private Sub ParseResponse()

    '   Parse out the response based on the issued
    '   command.
    If m_aszWCReceivedLines(m_iWCReceivedIndex) = "ERROR" Then
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Parm=" & m_szParameter & "," & m_iParameter
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": 'ERROR' returned from OpenAT"
        m_eWC_Status = eSTAT_ERROR
    Else
        Select Case m_eWC_State
            Case eWC_NULL
                '   Do nothing.  This happens when the
                '   first response line is not the
                '   command that was sent.

            Case eWC_VERIFY, eWC_RESTART_GSM, eWC_STOP_APP, eWC_START_APP, _
                 eWC_CHECK_BUZZER, eWC_SET_FREQUENCY
                If m_iWCReceivedIndex < 1 Then
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Index=" & m_iWCReceivedIndex & " (<1)"
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Parm=" & m_szParameter & "," & m_iParameter
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": (-0)=" & m_aszWCReceivedLines(m_iWCReceivedIndex)
                    m_eWC_Status = eSTAT_ERROR
                ElseIf m_aszWCReceivedLines(m_iWCReceivedIndex - 1) = m_szATCommand And _
                   m_aszWCReceivedLines(m_iWCReceivedIndex) = "OK" Then
                    m_eWC_Status = eSTAT_SUCCESS
                Else
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Parm=" & m_szParameter & "," & m_iParameter
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": (-1)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 1)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": (-0)=" & m_aszWCReceivedLines(m_iWCReceivedIndex)
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_FW_VERSION, eWC_XMODEM_VERSION, eWC_GET_CCID, eWC_GET_IMSI, _
                 eWC_GET_IMEI, eWC_LOCK_GPS, eWC_REGISTER_GPRS, eWC_CHECK_IGN, _
                 eWC_CHECK_STARTER, eWC_GET_SIGNAL_STRENGTH, eWC_GET_GPS_LOCATION, _
                 eWC_GET_FREQUENCY, eWC_GET_DATA_STORAGE, eWC_SET_DATA_STORAGE
                If m_iWCReceivedIndex < 2 Then
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Index=" & m_iWCReceivedIndex & " (<2)"
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Parm=" & m_szParameter & "," & m_iParameter
If m_iWCReceivedIndex = 1 Then
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": (-1)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 1)
End If
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": (-0)=" & m_aszWCReceivedLines(m_iWCReceivedIndex)
                    m_eWC_Status = eSTAT_ERROR
                ElseIf m_aszWCReceivedLines(m_iWCReceivedIndex - 2) = m_szATCommand And _
                   m_aszWCReceivedLines(m_iWCReceivedIndex) = "OK" Then
                    m_szParameter = m_aszWCReceivedLines(m_iWCReceivedIndex - 1)
                    m_eWC_Status = eSTAT_SUCCESS
                Else
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": Parm=" & m_szParameter & "," & m_iParameter
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": (-2)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 2)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": (-1)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 1)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & ": (-0)=" & m_aszWCReceivedLines(m_iWCReceivedIndex)
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_DOWNLOAD_APP
                If m_iWCReceivedIndex < 1 Then
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & "WDWL: Index=" & m_iWCReceivedIndex & " (<1)"
                    m_eWC_Status = eSTAT_ERROR
                ElseIf m_aszWCReceivedLines(m_iWCReceivedIndex - 1) = m_szATCommand And _
                   m_aszWCReceivedLines(m_iWCReceivedIndex) = "+WDWL: 0" Then
                    m_eWC_Status = eSTAT_DOWNLOADING
                Else
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & "WDWL: Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & "WDWL: Parm=" & m_szParameter & "," & m_iParameter
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & "WDWL: (-1)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 1)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & m_eWC_State & "WDWL: (-0)=" & m_aszWCReceivedLines(m_iWCReceivedIndex)
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_ERASE_OBJECTS
                If m_iWCReceivedIndex < 1 Then
                    m_eWC_Status = eSTAT_ERROR
                ElseIf m_aszWCReceivedLines(m_iWCReceivedIndex - 1) = m_szATCommand And _
                       (m_aszWCReceivedLines(m_iWCReceivedIndex) = "OK" Or _
                       m_aszWCReceivedLines(m_iWCReceivedIndex) = "+CME ERROR: 548") Then
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

           Case eWC_CHECK_PPAY
                If m_aszWCReceivedLines(m_iWCReceivedIndex) = "$PPAY 1234" Then
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$PPAY: Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$PPAY: Parm=" & m_szParameter & "," & m_iParameter
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$PPAY: (-0)=" & m_aszWCReceivedLines(m_iWCReceivedIndex)
                End If

            Case eWC_GET_MODEM_ID
                If InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "ModemId is:") > 0 Then
                    m_szParameter = Right(m_aszWCReceivedLines(m_iWCReceivedIndex), Len(m_aszWCReceivedLines(m_iWCReceivedIndex)) - InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), ":"))
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_SET_MODEM_ID
                If InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "ModemId is:") > 0 And _
                   Chr(&H22) & Trim(Right(m_aszWCReceivedLines(m_iWCReceivedIndex), Len(m_aszWCReceivedLines(m_iWCReceivedIndex)) - (InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), ":")))) & Chr(&H22) = m_szParameter Then
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_GET_APN
                If InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$APN") > 0 Then
                    m_szParameter = Right(m_aszWCReceivedLines(m_iWCReceivedIndex), Len(m_aszWCReceivedLines(m_iWCReceivedIndex)) - InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), " "))
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_SET_APN
                If m_iWCReceivedIndex < 1 Then
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$APN: Index=" & m_iWCReceivedIndex & " (<1)"
                    m_eWC_Status = eSTAT_ERROR
                ElseIf m_aszWCReceivedLines(m_iWCReceivedIndex - 1) = "Password checks out" And _
                   InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$APN ") > 0 Then
                    m_eWC_Status = eSTAT_SUCCESS
                Else
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$APN: Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$APN: Parm=" & m_szParameter & "," & m_iParameter
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$APN: (-1)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 1)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$APN: (-0)=" & m_aszWCReceivedLines(m_iWCReceivedIndex)
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_GET_SERVER_IP
                If m_iWCReceivedIndex < 5 Then
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: Index=" & m_iWCReceivedIndex & " (<5)"
                    m_eWC_Status = eSTAT_ERROR
                ElseIf InStr(m_aszWCReceivedLines(m_iWCReceivedIndex - 5), "$SERVER[1-5]") > 0 Then
                    m_szParameter = Right(m_aszWCReceivedLines(m_iWCReceivedIndex - 5 + m_iParameter), Len(m_aszWCReceivedLines(m_iWCReceivedIndex - 5 + m_iParameter)) - 2)
                    m_eWC_Status = eSTAT_SUCCESS
                Else
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: Comd=" & m_szATCommand
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: Parm=" & m_szParameter & "," & m_iParameter
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: (-5)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 5)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: (-4)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 4)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: (-3)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 3)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: (-2)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 2)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: (-1)=" & m_aszWCReceivedLines(m_iWCReceivedIndex - 1)
g_fMainTester.txtMessage.Text = g_fMainTester.txtMessage.Text + vbCrLf & "$SERVER: (-0)=" & m_aszWCReceivedLines(m_iWCReceivedIndex)
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_SET_SERVER_IP
                If InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$SERVER ") > 0 And _
                    Val(Trim(Right(m_aszWCReceivedLines(m_iWCReceivedIndex), 2))) = m_iParameter Then
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_GET_PORT
                If InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$PORT ") > 0 Then
                    m_aszWCReceivedLines(m_iWCReceivedIndex) = Trim(m_aszWCReceivedLines(m_iWCReceivedIndex))
                    m_szParameter = Right(m_aszWCReceivedLines(m_iWCReceivedIndex), Len(m_aszWCReceivedLines(m_iWCReceivedIndex)) - InStrRev(m_aszWCReceivedLines(m_iWCReceivedIndex), " "))
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_SET_PORT
                m_aszWCReceivedLines(m_iWCReceivedIndex) = Trim(m_aszWCReceivedLines(m_iWCReceivedIndex))
                If InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "$PORT ") > 0 And _
                    Chr(&H22) & Trim(Right(m_aszWCReceivedLines(m_iWCReceivedIndex), Len(m_aszWCReceivedLines(m_iWCReceivedIndex)) - InStrRev(m_aszWCReceivedLines(m_iWCReceivedIndex), " "))) & Chr(&H22) = m_szParameter Then
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_GET_APVER
                If InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "Version is: ") > 0 Then
                    m_szParameter = Trim(Right(m_aszWCReceivedLines(m_iWCReceivedIndex), Len(m_aszWCReceivedLines(m_iWCReceivedIndex)) - InStrRev(m_aszWCReceivedLines(m_iWCReceivedIndex), " ")))
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_RESET_STATE
                If m_iWCReceivedIndex < 1 Then
                    m_eWC_Status = eSTAT_ERROR
                ElseIf InStr(m_aszWCReceivedLines(m_iWCReceivedIndex), "Sending Reset State: ") > 0 Then
                    m_szParameter = m_aszWCReceivedLines(m_iWCReceivedIndex - 1)
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case eWC_GET_VOLTS
                If m_iWCReceivedIndex < 2 Then
                    m_eWC_Status = eSTAT_ERROR
                ElseIf m_aszWCReceivedLines(m_iWCReceivedIndex) = "OK" And _
                       m_aszWCReceivedLines(m_iWCReceivedIndex - 2) = m_szATCommand Then
                    m_iParameter = Val(m_aszWCReceivedLines(m_iWCReceivedIndex - 1))
                    m_eWC_Status = eSTAT_SUCCESS
                Else
                    m_eWC_Status = eSTAT_ERROR
                End If

            Case Else
                MsgBox "An illegal state (" & m_eWC_State & ") in the WC Com state machine has been reached" & vbCrLf & vbCrLf & _
                       "  -Received: '" & m_aszWCReceivedLines(m_iWCReceivedIndex) & "'"
                m_eWC_Status = eSTAT_ERROR

        End Select
    End If

End Sub

Private Sub Form_Load()

    Dim i As Integer

    '   Initialize all of the contols on the form
'    ClearRequired
'    cbxComPortPS.ListIndex = -1
'    cbxBaudRatePS.ListIndex = -1
'    cbxParityPS.ListIndex = -1
'    cbxDataBitsPS.ListIndex = -1
'    cbxStopBitsPS.ListIndex = -1
'    cbxComPortWC.ListIndex = -1
'    cbxBaudRateWC.ListIndex = -1
'    cbxParityWC.ListIndex = -1
'    cbxDataBitsWC.ListIndex = -1
'    cbxStopBitsWC.ListIndex = -1
'    If m_iApply = vbChecked Then
        '       Set up the Power Supply COM setting combo boxes
'        cbxComPortPS.ListIndex = g_iComPortPS - 1
'        For i = 0 To cbxBaudRatePS.ListCount - 1
'            If cbxBaudRatePS.List(i) = g_szComBaudPS Then
'                cbxBaudRatePS.ListIndex = i
'                Exit For
'            End If
'        Next i
'        For i = 0 To cbxParityPS.ListCount - 1
'            If Left(cbxParityPS.List(i), 1) = g_szComParityPS Then
'                cbxParityPS.ListIndex = i
'                Exit For
'            End If
'        Next i
'        For i = 0 To cbxDataBitsPS.ListCount - 1
'            If cbxDataBitsPS.List(i) = g_szComDataBitsPS Then
'                cbxDataBitsPS.ListIndex = i
'                Exit For
'            End If
'        Next i
'        For i = 0 To cbxStopBitsPS.ListCount - 1
'            If cbxStopBitsPS.List(i) = g_szComStopBitsPS Then
'                cbxStopBitsPS.ListIndex = i
'                Exit For
'            End If
'        Next i
        '       Set up the Wavecom COM setting combo boxes
'        cbxComPortWC.ListIndex = g_iComPortWC - 1
'        For i = 0 To cbxBaudRateWC.ListCount - 1
'            If cbxBaudRateWC.List(i) = g_szComBaudWC Then
'                cbxBaudRateWC.ListIndex = i
'                Exit For
'            End If
'        Next i
'        For i = 0 To cbxParityWC.ListCount - 1
'            If Left(cbxParityWC.List(i), 1) = g_szComParityWC Then
'                cbxParityWC.ListIndex = i
'                Exit For
'            End If
'        Next i
'        For i = 0 To cbxDataBitsWC.ListCount - 1
'            If cbxDataBitsWC.List(i) = g_szComDataBitsWC Then
'                cbxDataBitsWC.ListIndex = i
'                Exit For
'            End If
'        Next i
'        For i = 0 To cbxStopBitsWC.ListCount - 1
'            If cbxStopBitsWC.List(i) = g_szComStopBitsWC Then
'                cbxStopBitsWC.ListIndex = i
'                Exit For
'            End If
'        Next i
'    End If
    '   Finally, initialize the "Appy at startup" check box
'    ckApplySerialSettings.value = m_iApply
    m_bKillTimers = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    m_ePS_Status = eSTAT_SHUTTING_DOWN
    m_eWC_Status = eSTAT_SHUTTING_DOWN
    m_bKillTimers = True
    tmrPSCom.Enabled = False
    tmrPSCom.Interval = 0
    tmrWCCom.Enabled = False
    tmrWCCom.Interval = 0
    tmrXMODEM.Enabled = False
    tmrXMODEM.Interval = 0
    If m_bPSOutput Then
        '   Turn the Power Supply OFF
        TransmitToPS ePS_TURN_OFF
        While m_ePS_Status = eSTAT_RUNNING
            DoEvents
        Wend
        If m_ePS_Status <> eSTAT_SUCCESS Then
            MsgBox "Failed to turn the Power Supply OFF"
        End If
    End If
    If ComPortToPowerSupply.PortOpen = True Then
        ComPortToPowerSupply.PortOpen = False
    End If
    If ComPortToWavecom.PortOpen = True Then
        ComPortToWavecom.PortOpen = False
    End If

End Sub

'   This function connects a COM port to the Power Supply and verifies that the
'   instrument is responding correctly.

Public Function ConnectToPowerSupply() As Boolean

    On Error GoTo ComPSError
    '   We are stuck in this loop until the user either connects to the Power
    '   Supply successfully or they press the Quit button
    ConnectToPowerSupply = False
    While Not ConnectToPowerSupply And Not g_bQuitProgram And m_ePS_Status <> eSTAT_SHUTTING_DOWN
        DoEvents
        If g_iComPortPS <> 0 And g_szComBaudPS <> "" And g_szComParityPS <> "" And _
           g_szComDataBitsPS <> "" And g_szComStopBitsPS <> "" Then
            '   If the COM port to the Power Supply is already open then close
            '   it.
            If ComPortToPowerSupply.PortOpen = True Then
                ComPortToPowerSupply.PortOpen = False
            End If
            ' Initialize the COM port settings then open it.
            ComPortToPowerSupply.CommPort = g_iComPortPS
            ComPortToPowerSupply.RThreshold = 1
            ComPortToPowerSupply.SThreshold = 1
            ComPortToPowerSupply.Handshaking = comNone
            ComPortToPowerSupply.ParityReplace = Chr$(255)
            ComPortToPowerSupply.CTSTimeout = 0
            ComPortToPowerSupply.DSRTimeout = 0
            ComPortToPowerSupply.DTREnable = True
            ComPortToPowerSupply.NullDiscard = True
            ComPortToPowerSupply.InBufferSize = 1024
            ComPortToPowerSupply.InputLen = 1
            ComPortToPowerSupply.OutBufferSize = 512
            ComPortToPowerSupply.InBufferCount = 0
            ComPortToPowerSupply.OutBufferCount = 0
            ComPortToPowerSupply.Settings = g_szComBaudPS & "," & _
                                            Left(g_szComParityPS, 1) & "," & _
                                            g_szComDataBitsPS & "," & _
                                            g_szComStopBitsPS
            ComPortToPowerSupply.PortOpen = True
            '   Attempt to communicate with the Power Supply and verify correct response.
            TransmitToPS ePS_VERIFY
            While m_ePS_Status = eSTAT_RUNNING
                DoEvents
            Wend
            If m_ePS_Status = eSTAT_SUCCESS Then
                ConnectToPowerSupply = True
            Else
                TransmitToPS ePS_CLEAR_ERROR
                While m_ePS_Status = eSTAT_RUNNING
                    DoEvents
                Wend
            End If
        End If
        GoTo SkipPSError
        
ComPSError:
        Select Case Err.Number
            Case comPortAlreadyOpen
                MsgBox "COM port " & g_iComPortPS & " is already open!", vbCritical
    
            Case comPortInvalid
                MsgBox "COM port " & g_iComPortPS & " is invalid!" & vbCrLf & vbCrLf & "Try another port number.", vbCritical
    
            Case Else
                MsgBox "Error " & Err.Number & " while attempting to open COM port " & g_iComPortPS & " to Power Supply!", vbCritical
    
        End Select

SkipPSError:
        If Not ConnectToPowerSupply Then
            ShowPS
        End If
    Wend

End Function

' This function connects the RS-232 connection to the Wavecom module inside the Elite device.

Public Function ConnectToWavecom() As Boolean

    On Error GoTo ComWCError
    '   We are stuck in this loop until the user either connects to the Wavecom
    '   module successfully or they press the Quit button
    ConnectToWavecom = False
    While Not ConnectToWavecom And Not g_bQuitProgram And m_eWC_Status <> eSTAT_SHUTTING_DOWN
        DoEvents
        If g_iComPortWC <> 0 And g_szComBaudWC <> "" And g_szComParityWC <> "" And _
           g_szComDataBitsWC <> "" And g_szComStopBitsWC <> "" Then
            '   If the COM port to the Wavecom module is already open then
            '   close it.
            If ComPortToWavecom.PortOpen = True Then
                ComPortToWavecom.PortOpen = False
            End If
            ' Initialize the COM port settings then open it.
            ComPortToWavecom.CommPort = g_iComPortWC
            ComPortToWavecom.RThreshold = 1
            ComPortToWavecom.SThreshold = 1
            ComPortToWavecom.Handshaking = comNone
            ComPortToWavecom.ParityReplace = Chr$(255)
            ComPortToWavecom.CTSTimeout = 0
            ComPortToWavecom.DSRTimeout = 0
            ComPortToWavecom.DTREnable = True
            ComPortToWavecom.NullDiscard = True
            ComPortToWavecom.InBufferSize = 1024
            ComPortToWavecom.InputLen = 1
            ComPortToWavecom.OutBufferSize = 512
            ComPortToWavecom.InBufferCount = 0
            ComPortToWavecom.OutBufferCount = 0
            ComPortToWavecom.Settings = g_szComBaudWC & "," & _
                                        Left(g_szComParityWC, 1) & "," & _
                                        g_szComDataBitsWC & "," & _
                                        g_szComStopBitsWC
            ComPortToWavecom.PortOpen = True
            ConnectToWavecom = True
        End If
        GoTo SkipWCError

ComWCError:
        Select Case Err.Number
            Case comPortAlreadyOpen
                MsgBox "COM port " & g_iComPortWC & " is already open!", vbCritical
    
            Case comPortInvalid
                MsgBox "COM port " & g_iComPortWC & " is invalid!" & vbCrLf & vbCrLf & "Try another port number.", vbCritical
    
            Case Else
                MsgBox "Error " & Err.Number & " while attempting to open COM port " & g_iComPortWC & " to Wavecom module!", vbCritical
    
        End Select

SkipWCError:
        If Not ConnectToWavecom Then
            ShowWC
        End If
    Wend

End Function

'   This function transmits a string to the Power Supply and starts the PS COM
'   state machine going to handle all of the events that will during this COM
'   session.
Private Sub TransmitToPS(eState As etPS_STATES)

    m_ePS_Status = eSTAT_RUNNING
    m_ePS_State = eState
    m_szPSReceive = ""
    '   Start the timer going
    tmrPSCom.Enabled = True
    If ComPortToPowerSupply.PortOpen = True Then
        Select Case m_ePS_State
            Case ePS_NULL
                '   Ignore this state

            Case ePS_VERIFY
                ComPortToPowerSupply.Output = "SYST:VERS?" & vbCrLf

            Case ePS_CHECK_FOR_ERROR
                ComPortToPowerSupply.Output = "*ESR?" & vbCrLf

            Case ePS_GET_ERROR
                ComPortToPowerSupply.Output = "SYST:ERR?" & vbCrLf

            Case ePS_CLEAR_ERROR
                ComPortToPowerSupply.Output = "*CLS" & vbCrLf

            Case ePS_INIT_1
                ComPortToPowerSupply.Output = "*RST" & vbCrLf

            Case ePS_INIT_2
                ComPortToPowerSupply.Output = "INST:NSEL 1" & vbCrLf

            Case ePS_INIT_3
                ComPortToPowerSupply.Output = "VOLT:RANG HIGH" & vbCrLf

            Case ePS_INIT_4
                ComPortToPowerSupply.Output = "APPL 12.0,0.5" & vbCrLf

            Case ePS_INIT_5
                ComPortToPowerSupply.Output = "INST:NSEL 2" & vbCrLf

            Case ePS_INIT_6
                ComPortToPowerSupply.Output = "VOLT:RANG HIGH" & vbCrLf

            Case ePS_INIT_7
                ComPortToPowerSupply.Output = "APPL 9.0,0.1" & vbCrLf

            Case ePS_INIT_8
                ComPortToPowerSupply.Output = "INST:NSEL 1" & vbCrLf

            Case ePS_TURN_ON
                ComPortToPowerSupply.Output = "OUTP ON" & vbCrLf

            Case ePS_TURN_OFF
                ComPortToPowerSupply.Output = "OUTP OFF" & vbCrLf

            Case ePS_MEAS_CURRENT
                ComPortToPowerSupply.Output = "MEAS:CURR?" & vbCrLf

            Case ePS_MEAS_VOLTS
                ComPortToPowerSupply.Output = "MEAS:VOLT?" & vbCrLf

            Case Else
                MsgBox "PS transmit, illegal state (" & m_ePS_State & ")"
                m_ePS_State = ePS_NULL
        End Select
    Else
        m_ePS_Status = eSTAT_NOT_CONNECTED
    End If
    If m_ePS_State = ePS_NULL Then
        '   Stop the timer, nothing got sent out
        tmrPSCom.Enabled = False
        m_ePS_Status = eSTAT_SUCCESS
    End If

End Sub

'   This function transmits a string to the Wavecom module and starts the WC
'   COM state machine going to handle all of the events that will during this
'   COM session.
Private Sub TransmitToWC(eState As etWC_STATES)

    m_eWC_Status = eSTAT_RUNNING
    m_eWC_State = eState
    '   Start the timer going
    tmrWCCom.Enabled = True
    If ComPortToWavecom.PortOpen = True Then
        Select Case m_eWC_State
            Case eWC_NULL
                '   Ignore this state

            Case eWC_VERIFY
                m_szATCommand = "AT"

            Case eWC_FW_VERSION
                m_szATCommand = "ATI3"

            Case eWC_XMODEM_VERSION
                m_szATCommand = "AT+WDWL?"

            Case eWC_DOWNLOAD_APP
                m_szATCommand = "AT+WDWL"

            Case eWC_RESTART_GSM
                m_szATCommand = "AT+CFUN=1"

            Case eWC_STOP_APP
                m_szATCommand = "AT+WOPEN=0"

            Case eWC_ERASE_OBJECTS
                m_szATCommand = "AT+WOPEN=3"

            Case eWC_START_APP
                m_szATCommand = "AT+WOPEN=1"

            Case eWC_GET_IMSI
                m_szATCommand = "AT+CIMI"

            Case eWC_GET_CCID
                m_szATCommand = "AT+CCID?"

            Case eWC_GET_IMEI
                m_szATCommand = "AT+CGSN"

            Case eWC_LOCK_GPS
                m_szATCommand = "AT$GPSLOCK"

            Case eWC_REGISTER_GPRS
                m_szATCommand = "AT$GPRSFOUND"

            Case eWC_CHECK_IGN
                m_szATCommand = "AT$IGNSTATE"

            Case eWC_CHECK_STARTER
                m_szATCommand = "AT$STARTERSTATE"

            Case eWC_CHECK_PPAY
                m_szATCommand = "AT$PPAY=1234"

            Case eWC_CHECK_BUZZER
                m_szATCommand = "AT+WTONE=1,2,500,15,5"

            Case eWC_GET_SIGNAL_STRENGTH
                m_szATCommand = "AT+CSQ"

            Case eWC_GET_GPS_LOCATION
                m_szATCommand = "AT$GPSLOC"

            Case eWC_GET_FREQUENCY
                m_szATCommand = "AT+WMBS?"

            Case eWC_SET_FREQUENCY
                m_szATCommand = "AT+WMBS=" & m_szParameter

            Case eWC_GET_DATA_STORAGE
                m_szATCommand = "AT+WOPEN=6"

            Case eWC_SET_DATA_STORAGE
                m_szATCommand = "AT+WOPEN=6," & m_szParameter

            Case eWC_GET_MODEM_ID
                m_szATCommand = "AT$MODID?"

            Case eWC_SET_MODEM_ID
                m_szATCommand = "AT$MODID=" & m_szParameter

            Case eWC_GET_APN
                m_szATCommand = "AT$APN"

            Case eWC_SET_APN
                m_szATCommand = "AT$APN=" & m_szParameter

            Case eWC_GET_SERVER_IP
                m_szATCommand = "AT$SERVER?"

            Case eWC_SET_SERVER_IP
                m_szATCommand = "AT$SERVER=" & m_szParameter

            Case eWC_GET_PORT
                m_szATCommand = "AT$PORT"

            Case eWC_SET_PORT
                m_szATCommand = "AT$PORT=" & m_szParameter

            Case eWC_GET_APVER
                m_szATCommand = "AT$APVER"

            Case eWC_RESET_STATE
                m_szATCommand = "AT$RESETSTATE"

            Case eWC_GET_VOLTS
                m_szATCommand = "AT$VOLT"

            Case Else
                MsgBox "WC transmit, illegal state (" & m_eWC_State & ")"
                m_eWC_State = eWC_NULL
        End Select
        If m_eWC_State <> eWC_NULL Then
            '   Send the AT command to the Wavecom module
            ComPortToWavecom.Output = m_szATCommand & vbCr
        End If
    Else
        m_eWC_Status = eSTAT_NOT_CONNECTED
    End If
    If m_eWC_State = eWC_NULL Then
        '   Stop the timer, nothing got sent out
        tmrWCCom.Enabled = False
        m_eWC_Status = eSTAT_SUCCESS
    End If

End Sub

Private Function CheckForPSError() As Boolean

    CheckForPSError = False
    TransmitToPS ePS_CHECK_FOR_ERROR
    While m_ePS_Status = eSTAT_RUNNING
        DoEvents
    Wend
    If m_ePS_Status = eSTAT_SUCCESS Or m_ePS_Status = eSTAT_TIMEOUT Then
        '   If the Power Supply doesn't respond then it doesn't have any error
        '   messages for us (so this is a good thing).
        m_ePS_Status = eSTAT_SUCCESS
        CheckForPSError = True
    Else
        '   Query the Power Supply for whatever error it has detected
        TransmitToPS ePS_GET_ERROR
        While m_ePS_Status = eSTAT_RUNNING
            DoEvents
        Wend
    End If

End Function

Public Function InitPowerSupply() As Boolean

    InitPowerSupply = False
    While Not InitPowerSupply And Not g_bQuitProgram
        DoEvents
        TransmitToPS ePS_INIT_PS
        While m_ePS_Status = eSTAT_RUNNING
            DoEvents
        Wend
        If m_ePS_Status = eSTAT_SUCCESS Then
            '   Check for any errors in the Power Supply
            If CheckForPSError Then
                InitPowerSupply = True
            End If
        Else
            MsgBox "Failed to intialize the power supply: initialization phase = " & m_ePS_State - ePS_INIT_PS + 1 & "; status = " & m_ePS_Status, vbCritical
            ShowPS
        End If
    Wend

End Function

Public Function PowerSupply1_Current() As Double

    TransmitToPS ePS_MEAS_CURRENT
    While m_ePS_Status = eSTAT_RUNNING
        DoEvents
    Wend
    If m_ePS_Status = eSTAT_SUCCESS Then
        PowerSupply1_Current = CDbl(m_szPSReceive)
    Else
        MsgBox "Error Receiving Measured Current From Power Supply" & vbCrLf & vbCrLf & _
               "Status = " & m_ePS_Status & ", received = '" & m_szPSReceive, vbCritical
        PowerSupply1_Current = -9999
        TransmitToPS ePS_CLEAR_ERROR
        While m_ePS_Status = eSTAT_RUNNING
            DoEvents
        Wend
    End If

End Function

Public Function PowerSupply1_Volts() As Double

    TransmitToPS ePS_MEAS_VOLTS
    While m_ePS_Status = eSTAT_RUNNING
        DoEvents
    Wend
    If m_ePS_Status = eSTAT_SUCCESS Then
        PowerSupply1_Volts = CDbl(m_szPSReceive)
    Else
        MsgBox "Error Receiving Measured Voltage From Power Supply" & vbCrLf & vbCrLf & _
               "Status = " & m_ePS_Status & ", received = '" & m_szPSReceive, vbCritical
        PowerSupply1_Volts = -9999
        TransmitToPS ePS_CLEAR_ERROR
        While m_ePS_Status = eSTAT_RUNNING
            DoEvents
        Wend
    End If

End Function

Public Function OpenATQuery() As Boolean

    OpenATQuery = SendATCommand(eWC_VERIFY)

End Function

'   *** Currently not used ***
'Public Function GetFWVersion(szFWVersion As String) As Boolean

'    GetFWVersion = SendATCommand(eWC_FW_VERSION)
'    If GetFWVersion Then
'        szFWVersion = m_szParameter
'    End If

'End Function

'   *** Currently not used ***
'Public Function GetFrequency(iFrequency As Integer)

'    GetFrequency = SendATCommand(eWC_GET_FREQUENCY)
'    If GetFrequency Then
'        iFrequency = Val(m_szParameter)
'    End If

'End Function

'   *** Currently not used ***
'Public Function SetFrequency(iFrequency As Integer)

'    m_szParameter = Str(iFrequency)
'    SetFrequency = SendATCommand(eWC_SET_FREQUENCY)

'End Function

Public Function GetDataStorage(iStorage As Integer)

    GetDataStorage = SendATCommand(eWC_GET_DATA_STORAGE)
    If GetDataStorage Then
        iStorage = Val(Right(m_szParameter, Len(m_szParameter) - Len("+WOPEN: 6,")))
    End If

End Function

Public Function SetDataStorage(iStorage As Integer)

    m_szParameter = Str(iStorage)
    SetDataStorage = SendATCommand(eWC_SET_DATA_STORAGE)

End Function

Public Function InitiateOpenATAppDownload() As Boolean

    InitiateOpenATAppDownload = SendATCommand(eWC_DOWNLOAD_APP)
    If InitiateOpenATAppDownload Then
        '   We must turn on RTS handshaking then close the COM port to the
        '   Wavecom module first because the XMODEM DLL is about to take it
        '   over.
        ComPortToWavecom.PortOpen = False
    End If

End Function

'   *** Currently not used ***
'Public Function GetXMODEMVersion(szXMODEMVersion As String) As Boolean

'    GetXMODEMVersion = SendATCommand(eWC_XMODEM_VERSION)
'    If GetXMODEMVersion Then
'        szXMODEMVersion = m_szParameter
'    End If

'End Function

Public Function RestartGSMRegistration() As Boolean

    RestartGSMRegistration = SendATCommand(eWC_RESTART_GSM)

End Function

Public Function StopOpenATApp() As Boolean

    StopOpenATApp = SendATCommand(eWC_STOP_APP)

End Function

Public Function EraseOpenATFlashObjects() As Boolean

    EraseOpenATFlashObjects = SendATCommand(eWC_ERASE_OBJECTS)

End Function

Public Function StartOpenATApp() As Boolean

    StartOpenATApp = SendATCommand(eWC_START_APP)

End Function

Public Sub FinishOpenATAppDownload()

    '   Reopen the WC port so that regular AT commands can be sent out then
    '   fake a successful AT command completion in the status (so that other
    '   routines will think that all is happy and right with the world).
    ComPortToWavecom.Handshaking = comNone
    ComPortToWavecom.PortOpen = True
    m_eWC_Status = eSTAT_SUCCESS

End Sub

Public Function SetModemID(lModemID As Long)

    m_szParameter = Chr(&H22) & Format(lModemID, "00000000") & Chr(&H22)
    SetModemID = SendATCommand(eWC_SET_MODEM_ID)

End Function

Public Function SetAPN(szAPN As String, szPassword As String)


    m_szParameter = Chr(&H22) & szAPN & Chr(&H22) & "," & Chr(&H22) & szPassword & Chr(&H22)
    SetAPN = SendATCommand(eWC_SET_APN)

End Function

Public Function SetServerIP(iServer As Integer, szIP As String)

    m_iParameter = iServer
    m_szParameter = Chr(&H22) & Str(iServer) & Chr(&H22) & "," & Chr(&H22) & szIP & Chr(&H22)
    SetServerIP = SendATCommand(eWC_SET_SERVER_IP)

End Function

Public Function SetPort(szPort As String) As Boolean

    m_szParameter = Chr(&H22) & szPort & Chr(&H22)
    SetPort = SendATCommand(eWC_SET_PORT)

End Function

Public Function GetAPVersion(szVersion As String) As Boolean

    GetAPVersion = SendATCommand(eWC_GET_APVER)
    If GetAPVersion Then
        szVersion = m_szParameter
    End If

End Function

Public Function GetResetState(bPowerUp As Boolean) As Boolean

    GetResetState = SendATCommand(eWC_RESET_STATE)
    If GetResetState Then
        If InStr(m_szParameter, "ADL_INIT_POWER_ON") > 0 Then
            bPowerUp = True
        Else
            bPowerUp = False
        End If
    End If

End Function

Public Function GetAPN(szAPN As String)

    GetAPN = SendATCommand(eWC_GET_APN)
    If GetAPN Then
        szAPN = m_szParameter
    End If

End Function

Public Function GetServerIP(iServer As Integer, szIP As String)

    '   The Get Server IP command actually returns 5 IP strings (one for each
    '   server).  The Response Parser provides a special case for this command
    '   in that it picks which of the 5 returned strings to place into
    '   m_szParameter based on the input value placed in m_iParameter.
    m_iParameter = iServer
    GetServerIP = SendATCommand(eWC_GET_SERVER_IP)
    If GetServerIP Then
        szIP = m_szParameter
    End If

End Function

Public Function GetPort(szPort As String)

    GetPort = SendATCommand(eWC_GET_PORT)
    If GetPort Then
        szPort = m_szParameter
    End If

End Function

Public Function GetVolts(iVolts As Integer)

    GetVolts = SendATCommand(eWC_GET_VOLTS)
    If GetVolts Then
        iVolts = m_iParameter
    End If

End Function

Public Function GetModemID(szModemID As String)

    GetModemID = SendATCommand(eWC_GET_MODEM_ID)
    If GetModemID Then
        szModemID = m_szParameter
    End If

End Function

Public Function CheckPIC_IF() As Integer

    Dim dStart As Double
    Dim iRetries As Integer

    CheckPIC_IF = 0
    '   Three strikes and then you're out
    For iRetries = 0 To 3
        g_dSoundLevel = 0
        g_dSoundFreq = 0
        If Not SendATCommand(eWC_CHECK_PPAY) Then
            '   PPAY command failed
            CheckPIC_IF = 1
        Else
            dStart = Timer
            Do
                DoEvents
                '   Start listening
                g_fMicIn.RecordWav
            Loop While g_dSoundLevel < MIN_SOUND_LEVEL And Timer < dStart + 1
            If g_dSoundLevel >= MIN_SOUND_LEVEL Then
                CheckPIC_IF = 0
                Exit For
            End If
            '   PPAY beeps failed
            CheckPIC_IF = 2
        End If
    Next iRetries
    g_fMainTester.txtMessage = g_fMainTester.txtMessage + vbCrLf & "PPAY beep = " & g_dSoundLevel & "@" & g_dSoundFreq & "Hz"
    If CheckPIC_IF = 0 Then
        For iRetries = 0 To 3
            If Not SendATCommand(eWC_CHECK_IGN) Then
                '   Ignition relay command failure
                CheckPIC_IF = 3
            ElseIf m_szParameter <> "0" Then
                '   Ignition relay failure
                CheckPIC_IF = 4
            Else
                Exit For
            End If
        Next iRetries
        If CheckPIC_IF = 0 Then
            If Not SendATCommand(eWC_CHECK_STARTER) Then
                '   Starter relay command failure
                CheckPIC_IF = 5
            ElseIf m_szParameter <> "0" Then
                '   Starter relay failure
                CheckPIC_IF = 6
            ElseIf Not SendATCommand(eWC_CHECK_BUZZER) Then
                '   Wavecom tone command failed
                CheckPIC_IF = 7
            Else
                '   Start listening
                g_fMicIn.RecordWav
                If g_dSoundLevel < MIN_SOUND_LEVEL Then
                    '   Wavecom generated tone failed
                    CheckPIC_IF = 8
                End If
            End If
        End If
    End If

End Function

Public Function FindGPRS(bGPRSRegistered As Boolean) As Boolean

    '   Start attempting to register to a local cell tower
    FindGPRS = SendATCommand(eWC_REGISTER_GPRS)
    If FindGPRS And m_szParameter = "1" Then
        bGPRSRegistered = True
    Else
        bGPRSRegistered = False
    End If

End Function

Public Function GetIMEI(szIMEI As String) As Boolean

    '   Get the serial # for Wavecomm modem
    GetIMEI = SendATCommand(eWC_GET_IMEI)
    If GetIMEI Then
        szIMEI = m_szParameter
    End If

End Function

Public Function GetIMSI(szIMSI As String) As Boolean

    '   Get the SN of the simm card
    GetIMSI = SendATCommand(eWC_GET_IMSI)
    If GetIMSI Then
        szIMSI = m_szParameter
    End If

End Function
    
Public Function GetCCID(szCCID As String) As Boolean

    GetCCID = SendATCommand(eWC_GET_CCID)
    If GetCCID Then
        If InStr(1, m_szParameter, "+CCID: ") = 1 Then
            '   We need to take the last 21 characters from the input message
            '   because there is a trailing double quote at the end of the
            '   line.
            m_szParameter = Right(m_szParameter, Len(m_szParameter) - Len("+CCID: "))
            '   And this trims off the trailing quote
            m_szParameter = Left(m_szParameter, 20)
            If Left(m_szParameter, 1) = Chr(&H22) Then
                '   There's a leading double (because this CCID is only 19
                '   characters long).  Trim it off.
                m_szParameter = Right(m_szParameter, 19)
            End If
            szCCID = m_szParameter
        Else
            GetCCID = False
        End If
    End If

End Function
    
Public Function FindGPS(bGPSLock As Boolean) As Boolean

    '   Start attempting to lock onto a GPS satellite
    FindGPS = SendATCommand(eWC_LOCK_GPS)
    If FindGPS Then
        If m_szParameter = "1" Then
            bGPSLock = True
        Else
            bGPSLock = False
        End If
    End If

End Function

Public Function GetGPRSSignalStrength(iStrength As Integer) As Boolean

    Dim strStrength As String

    If m_eWC_State = eWC_NULL Then
        '   Get the GSM signal strength.
        GetGPRSSignalStrength = SendATCommand(eWC_GET_SIGNAL_STRENGTH)
        If GetGPRSSignalStrength And InStr(1, m_szParameter, "+CSQ: ") = 1 Then
            iStrength = CInt(Val(Right(m_szParameter, Len(m_szParameter) - Len("+CSQ: "))))
        End If
    Else
        GetGPRSSignalStrength = False
    End If

End Function

Public Function GetGPSCoordinates(lLongitude As Long, lLatitude As Long, iNoSatellites As Integer)

    GetGPSCoordinates = SendATCommand(eWC_GET_GPS_LOCATION)
    If GetGPSCoordinates Then
        '   Parse the coordinates and number of satellites out of the returned
        '   parameter
        iNoSatellites = CInt(Val(Right(m_szParameter, Len(m_szParameter) - InStrRev(m_szParameter, ","))))
        m_szParameter = Left(m_szParameter, InStrRev(m_szParameter, ",") - 1)
        lLongitude = Val(Right(m_szParameter, Len(m_szParameter) - InStrRev(m_szParameter, ",")))
        m_szParameter = Left(m_szParameter, InStrRev(m_szParameter, ",") - 1)
        lLatitude = Right(m_szParameter, Len(m_szParameter) - InStrRev(m_szParameter, ","))
    End If

End Function

Private Function SendATCommand(eATCommand As etWC_STATES) As Boolean

    SendATCommand = False
    TransmitToWC eATCommand
    While m_eWC_Status = eSTAT_RUNNING
        DoEvents
    Wend
    If m_eWC_Status = eSTAT_SUCCESS Or m_eWC_Status = eSTAT_DOWNLOADING Then
        '   The AT command was sent and executed successfully.
        SendATCommand = True
    Else
        g_fMainTester.txtMessage = g_fMainTester.txtMessage + vbCrLf & "AT Command " & eATCommand & " failed: status = "
        If m_eWC_Status = eSTAT_ERROR Then
            g_fMainTester.txtMessage = g_fMainTester.txtMessage + "OpenAT returned an ERROR"
        ElseIf m_eWC_Status = eSTAT_TIMEOUT Then
            g_fMainTester.txtMessage = g_fMainTester.txtMessage + "Response from OpenAT timed out"
        Else
            g_fMainTester.txtMessage = g_fMainTester.txtMessage + m_eWC_Status
        End If
    End If

End Function

Private Sub tmrPSCom_Timer()

    '   Turn the timer off
    tmrPSCom.Enabled = False
    m_ePS_Status = eSTAT_TIMEOUT

End Sub

Private Sub tmrWCCom_Timer()

    '   Turn the timer off
    tmrWCCom.Enabled = False
    If m_eWC_Status = eSTAT_RUNNING Then
        '   Flag a timeout, unless we're downloading
        m_eWC_Status = eSTAT_TIMEOUT
    End If

End Sub

Private Sub tmrXMODEM_Timer()

    Dim i As Integer
    Dim iCode As Integer
    Dim iPacket As Integer
    Dim strBuffer As String * 82

    '   Turn the timer off
    tmrXMODEM.Enabled = False
    If g_fDownloadProgress.g_bCancel Then
        '   Abort the download
        iCode = xyAbort(g_clsOpenAT.g_iPort)
    Else
        '   Execute 10 xyDriver states
        For i = 1 To 10
            iCode = xyDriver(g_clsOpenAT.g_iPort)
        Next i
        '   Flush out any messages returned from the xyDriver (and then
        '   discard them).
        strBuffer = Space(81)
        While xyGetMessage(g_clsOpenAT.g_iPort, strBuffer, 80) > 0
            DoEvents
'            g_fMainTester.txtMessage = g_fMainTester.txtMessage + vbCrLf & strBuffer
        Wend
    End If
    If iCode = XY_IDLE Then
        If g_fDownloadProgress.g_bCancel Then
            g_fMainTester.txtMessage = g_fMainTester.txtMessage + vbCrLf & "Download aborted!"
        ElseIf xyGetParameter(g_clsOpenAT.g_iPort, XY_GET_ERROR_CODE) = 0 Then
            g_fMainTester.txtMessage = g_fMainTester.txtMessage + vbCrLf & "Download completed!"
        Else
            g_fMainTester.txtMessage = g_fMainTester.txtMessage + vbCrLf & "Download aborted!" & vbCrLf & _
                                       "error code = " & xyGetParameter(g_clsOpenAT.g_iPort, XY_GET_ERROR_CODE) & _
                                       ", error state = " & xyGetParameter(g_clsOpenAT.g_iPort, XY_GET_ERROR_STATE)
        End If
        g_fDownloadProgress.g_bCancel = True
        g_fDownloadProgress.cmdCancel_Click
        g_clsOpenAT.g_bXMODEMDownloading = False
    Else
        '   xyDriver is still running
        iPacket = xyGetParameter(g_clsOpenAT.g_iPort, XY_GET_PACKET)
        g_fDownloadProgress.UpdateProgress (iPacket)
        '   Restart the timer.  Step the xyDriver another 10 states.
        If Not m_bKillTimers Then
            tmrXMODEM.Enabled = True
        End If
    End If

End Sub
