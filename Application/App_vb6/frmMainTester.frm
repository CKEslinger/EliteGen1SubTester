VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainTester 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PassTime Elite II Functional Tester"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   -4440
   ClientWidth     =   17760
   FillColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleMode       =   0  'User
   ScaleWidth      =   17760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCBCom 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4920
      Top             =   4680
   End
   Begin VB.Timer tmrWaitPowerUp 
      Interval        =   2000
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000002&
      Height          =   9255
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   12135
      Begin MSComctlLib.ListView lvwTestResults 
         Height          =   7335
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   11890
         _ExtentX        =   20981
         _ExtentY        =   12938
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblPassFail 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "***Device not installed***"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   11890
      End
      Begin VB.Label lblTestStatus 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TEST STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Timer tmrEliteCom 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4920
      Top             =   4080
   End
   Begin VB.Timer tmrPSCom 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4920
      Top             =   3480
   End
   Begin VB.ComboBox cbxSIMCards 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3000
      Width           =   4092
   End
   Begin VB.Frame frOperations 
      BackColor       =   &H8000000E&
      Caption         =   "Operations to perform"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   24
      Top             =   3600
      Width           =   4095
      Begin VB.CheckBox chkTrialRun 
         BackColor       =   &H8000000E&
         Caption         =   "Trial Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox chkConfigureUnit 
         BackColor       =   &H8000000E&
         Caption         =   "Configure Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chkDownload 
         BackColor       =   &H8000000E&
         Caption         =   "Download Firmware"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox chkTestFunctions 
         BackColor       =   &H8000000E&
         Caption         =   "Test Functions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.ComboBox cbxModelTypes 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2160
      Width           =   4092
   End
   Begin VB.Timer tmrWaitWireless 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   2280
   End
   Begin VB.Timer tmrAutoTest 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4920
      Top             =   1680
   End
   Begin VB.CommandButton cmdStartTest 
      Appearance      =   0  'Flat
      Caption         =   "START TEST"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   17
      Top             =   8880
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000002&
      Height          =   3372
      Left            =   720
      TabIndex        =   1
      Top             =   5280
      Width           =   4095
      Begin VB.CommandButton cmdResetCounter 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSerialNumber 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtTotalParts 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtPassedParts 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtFailedParts 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtCycleTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblSerialNumber 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1572
      End
      Begin VB.Label lblTotalParts 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Parts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label lblPassedParts 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Passed Parts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label lblFailedParts 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Failed Parts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label lblCycleTime 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Test Cycle Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1812
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   1215
      Left            =   240
      Picture         =   "frmMainTester.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   4995
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   5055
   End
   Begin MSCommLib.MSComm ComPortToPowerSupply 
      Left            =   4920
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      StopBits        =   2
      SThreshold      =   1
   End
   Begin MSCommLib.MSComm ComPortToElite 
      Left            =   4920
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
      InBufferSize    =   8192
      InputLen        =   1
      NullDiscard     =   -1  'True
      RThreshold      =   1
      BaudRate        =   115200
      SThreshold      =   1
   End
   Begin MSCommLib.MSComm ComPortToCallBox 
      Left            =   4920
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
      InBufferSize    =   8192
      InputLen        =   1
      NullDiscard     =   -1  'True
      RThreshold      =   1
      BaudRate        =   115200
      SThreshold      =   1
   End
   Begin VB.Label lblSIMCard 
      BackColor       =   &H80000009&
      Caption         =   "Installed SIM Card"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      TabIndex        =   26
      Top             =   2640
      Width           =   3012
   End
   Begin VB.Label lblModelType 
      BackColor       =   &H80000009&
      Caption         =   "Model Type Under Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      TabIndex        =   25
      Top             =   1800
      Width           =   3012
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSetup 
         Caption         =   "&Setup..."
      End
      Begin VB.Menu mnuSerialNumber 
         Caption         =   "Serial &Numbers..."
      End
      Begin VB.Menu mnuIMEI 
         Caption         =   "&IMEI Numbers..."
      End
   End
   Begin VB.Menu mnuDiag 
      Caption         =   "&Diagnostics"
      Begin VB.Menu mnuGSM 
         Caption         =   "&GSM..."
      End
      Begin VB.Menu mnuGPS 
         Caption         =   "G&PS..."
      End
      Begin VB.Menu mnuDAQ 
         Caption         =   "&DAQ..."
      End
      Begin VB.Menu mnuMicIn 
         Caption         =   "&Mic Input..."
      End
      Begin VB.Menu mnuPrintLabel 
         Caption         =   "&Label Printer..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMainTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' Constants
Private Const GPRS_TIMEOUT As Integer = 45
Private Const VOLTS_TIMEOUT As Integer = 5
Private Const CURRENT_TIMEOUT As Integer = 2
Private Const AUTO_TEST_DEBOUNCE As Integer = 10
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const IMEI_WRITE_WAIT As Double = 10#
Private Const SPI_WRITE_WAIT As Integer = 3#

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

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

''' Enums
Private Enum etTEST_STATE
    eTESTING
    eTEST_COMPLETE
    eTESTER_OPEN
    eTEST_APP_STARTED
    eTEST_DIAGNOSTIC
End Enum

Private Enum etRESULTS
    eRESULT_PASSED
    eRESULT_FAILED
    eRESULT_TESTING
    eRESULT_NOT_TESTED
    eRESULT_ERROR
End Enum

Private Enum etITEMS
    eITEM_DESCRIPTION = 1
    eITEM_RESULTS
    eITEM_PASS_FAIL
    eITEM_NO_OF_ITEMS
End Enum

''' Variables
Private m_bGPSTimeout As Boolean
Private m_bGPRSTimeout As Boolean
Private m_iCurrentTest As Integer
Private m_aszMessages() As String
Private m_bVoltsTimeout As Boolean
Private m_szOldModeltype As String
Private m_bCurrentTimeout As Boolean
Private m_bDeviceFailure As Boolean
Private m_bDownloadFailed As Boolean
Private m_bDownloadFW As Boolean
Private m_iWirelessTimeout  As Integer
Private m_eTestingState As etTEST_STATE
Private m_bApplicationRunning As Boolean
Private m_bConfigureSuccessful As Boolean
Private m_bSuspendWirelessTimeout As Boolean
Private m_bApplicationHasBeenReset As Boolean

Private m_iAutoTestDebounce As Integer
Private m_bTesterClosed As Boolean
Private m_szSerialNum As String

Private m_ExecutedTests As String
Private m_AllTests() As String

''' Variables added by SFS to collect GPS/GSM/Current values and add to log file (2/11/12)
Private m_test_Current As String
Private m_test_GPS_Sig As String
Private m_test_GSM_Sig As String

''' Forms
Private m_fAbout As frmAbout
Private m_fDiagDAQ As frmDiagDAQ
Private m_fDiagGPS As frmDiagGPS
Private m_fDiagGSM As frmDiagGSM
Private m_fContents As frmContents
Private m_fPassword As frmPassword
Private m_fSetupForm As frmTesterSetup
Private m_fDiagPrinter As frmDiagPrinter
Private m_fDetailedResults As frmDetailedResults
Private m_fUpdateSerialNumbers As frmUpdateSerialNumbers

Private Declare Function FindFirstFile Lib "kernel32" _
          Alias "FindFirstFileA" (ByVal lpFileName As String, _
          lpFindFileData As WIN32_FIND_DATA) As Long
          
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long
   
Private fso1 As FileSystemObject
Private ts As TextStream

Private m_bCntrlKeyDown As Boolean

Public m_dStartTime As Double

Public alldata As FileSystemObject
Public ad As TextStream


Private Sub cbxModelTypes_Click()

    Dim iEnd As Integer
    Dim iStart As Integer
    Dim oSIM As IXMLDOMNode
    Dim strSIMCards As String
    Dim modelName() As String

    If Len(cbxModelTypes.Text) = 0 Then
        Exit Sub
    End If
    '   Get the selected model's XML node
    modelName = Split(cbxModelTypes.Text, ":", 2)
    Set g_oSelectedModel = g_oModels.selectSingleNode(Trim(modelName(0)))
    '   Sanity check
    If g_oSelectedModel Is Nothing Then
        MsgBox "CRITICAL ERROR: Model type '" & cbxModelTypes.Text & "' can't be found!", vbCritical
        Exit Sub
    End If
    If m_szOldModeltype = g_oSelectedModel.baseName Then
        '   The window is probably being refreshed.  Don't continue else the
        '   SIM card selection will get lost (it would be a waste of time
        '   anyway).
        Exit Sub
    End If
    m_szOldModeltype = g_oSelectedModel.baseName
    '   Refresh which tests are going to be run for this model type.
    FormRefresh
    '   Access the Serial Numbers for this model
    g_lNextSerialNumber = CLng(Val(g_oSelectedModel.selectSingleNode("SerialNo").Attributes.getNamedItem("Next").Text))
    g_lEndingSerialNumber = CLng(Val(g_oSelectedModel.selectSingleNode("SerialNo").Attributes.getNamedItem("End").Text))
    txtSerialNumber.Text = Format(g_lNextSerialNumber, SERIAL_NUM_FORMAT)
    '   Clear out the SIM's combo box
    lblSIMCard.Enabled = False
    cbxSIMCards.Enabled = False
    While cbxSIMCards.ListCount > 0
        cbxSIMCards.RemoveItem (0)
    Wend
    '   Parse out the SIM Cards configured for this model type and then add
    '   them to the SIM Card combo box.
    strSIMCards = g_oSelectedModel.selectSingleNode("SIMs").Text
    iStart = 1
    While iStart < Len(strSIMCards)
        iEnd = InStr(iStart, strSIMCards, ",")
        If iEnd = 0 Then
            iEnd = Len(strSIMCards) + 1
        End If
        Set oSIM = g_oSIMs.selectSingleNode(Mid(strSIMCards, iStart, iEnd - iStart))
        If oSIM Is Nothing Then
            MsgBox "Invalid SIM card found in " & m_szOldModeltype & "'s XML profile", vbCritical
        Else
            cbxSIMCards.AddItem oSIM.baseName
            iStart = iEnd + 1
            lblSIMCard.Enabled = True
            cbxSIMCards.Enabled = True
        End If
    Wend
    
    cbxSIMCards.ListIndex = 0
    
    '   Retrieve the firmware file name for this model type
    g_szFirmwareFileName = g_oSelectedModel.selectSingleNode("Firmware").Text
    FirmwareVerFromFN (g_szFirmwareFileName)

End Sub

Private Sub cbxSIMCards_Click()

    '   Get the selected SIM card's XML node
    Set g_oSelectedSIM = g_oSIMs.selectSingleNode(cbxSIMCards.Text)
    '   Sanity check
    If g_oSelectedSIM Is Nothing Then
        MsgBox "CRITICAL ERROR: SIM card '" & cbxSIMCards.Text & "' can't be found!", vbCritical
        Exit Sub
    End If

End Sub

Private Sub chkConfigureUnit_Click()

    FormRefresh

End Sub

Private Sub chkDownload_Click()

    FormRefresh

End Sub

Private Sub chkTestFunctions_Click()

    FormRefresh

End Sub

Private Sub chkTrialRun_Click()

    FormRefresh
    
End Sub

Private Sub cmdStartTest_Click()

Dim runNum As Integer
Dim dlyCnt As Integer

Do
    Set ts = fso1.OpenTextFile(TEST_LOG_FN, ForAppending)
    ts.WriteLine "*****Test Run Number " & runNum & "*****"
    ts.Close
    
    '   Disable various controls on the Main Tester screen
    cmdStartTest.Enabled = False
    lblModelType.Enabled = False
    cbxModelTypes.Enabled = False
    lblSIMCard.Enabled = False
    cbxSIMCards.Enabled = False
    frOperations.Enabled = False
    chkConfigureUnit.Enabled = False
    chkDownload.Enabled = False
    chkTestFunctions.Enabled = False
    chkTrialRun.Enabled = False
    cmdResetCounter.Enabled = False
    mnuFile.Enabled = False
    mnuTools.Enabled = False
    mnuDiag.Enabled = False
    mnuHelp.Enabled = False
    lvwTestResults.Enabled = False
    '   Start the tests going
    StartTest
    '   Re-enable the controls
    If Not g_bAutoTest Then
        '   Allow the user to manually start the next test
        cmdStartTest.Enabled = True
    End If
    lblModelType.Enabled = True
    cbxModelTypes.Enabled = True
    lblSIMCard.Enabled = True
    cbxSIMCards.Enabled = True
    frOperations.Enabled = True
    chkConfigureUnit.Enabled = True
    chkDownload.Enabled = True
    chkTestFunctions.Enabled = True
    chkTrialRun.Enabled = True
    cmdResetCounter.Enabled = True
    mnuFile.Enabled = True
    mnuTools.Enabled = True
    mnuDiag.Enabled = True
    mnuHelp.Enabled = True
    lvwTestResults.Enabled = True
    
    If vbChecked = m_fSetupForm.chkAutotest.value Or False = g_bDebugOn Then
        Exit Do
    Else
        If "PTE-II.S" = g_oSelectedModel.baseName Or "PTE-II.NS" = g_oSelectedModel.baseName Then
            dlyCnt = 60
        Else
            dlyCnt = 5
        End If
        runNum = runNum + 1
        If "PASS" <> Left(lblPassFail.Caption, 4) Then
            lblPassFail.Caption = "Auto Test Failure on Run Number " & runNum
            lblPassFail.BackColor = vbRed
            Exit Do
        Else
            Do
                lblPassFail.Caption = "Run " & runNum & " in " & dlyCnt & "s (Ctrl-Double Click to Exit)"
                dlyCnt = dlyCnt - 1
                startup.Delay 1#
                If eTESTER_OPEN = m_eTestingState Then
                    dlyCnt = 0
                    Exit Sub
                End If
            Loop Until 0 = dlyCnt
        End If
    End If
Loop

End Sub

Private Sub ComPortToCallBox_OnComm()

    SerialComs.ComPortToCallBoxEventProcessor
    
End Sub

Private Sub ComPortToPowerSupply_OnComm()

    SerialComs.ComPortToPowerSupplyEventProcessor

End Sub

Private Sub ComPortToElite_OnComm()

    SerialComs.ComPortToEliteEventProcessor

End Sub

Private Sub Form_Activate()

    '   The Main Tester form has finally become visiable.  Start polling the
    '   tester continually to see if a DUT has been inserted.
    tmrAutoTest.Enabled = True

End Sub

Private Sub Form_Click()
    Dim i
    i = i + 1
End Sub

Private Sub Form_DblClick()
    Dim i
    If vbUnchecked = m_fSetupForm.chkAutotest.value And True = m_bCntrlKeyDown Then
        ' Cntrl key is down
        m_eTestingState = eTESTER_OPEN
'        m_bDeviceFailure = True
'        m_bDownloadFailed = False
'        m_bApplicationRunning = False
'        m_bApplicationHasBeenReset = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyControl = KeyCode Then
        m_bCntrlKeyDown = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If vbKeyControl = KeyCode Then
        m_bCntrlKeyDown = False
    End If
End Sub

Private Sub Form_Load()

    Dim i, j As Integer
    Dim ltmRow As ListItem
    Dim lsiItem As ListSubItem
    Dim oModel As IXMLDOMNode
    Dim strModelType As String
    Dim chrHeader As ColumnHeader

    Set m_fAbout = New frmAbout
    Set m_fContents = New frmContents
    Set m_fDiagDAQ = New frmDiagDAQ
    Set m_fDiagPrinter = New frmDiagPrinter
    Set m_fDiagGPS = New frmDiagGPS
    Set m_fDiagGSM = New frmDiagGSM
    Set m_fSetupForm = New frmTesterSetup
    Set m_fUpdateSerialNumbers = New frmUpdateSerialNumbers
    Set g_fSerialNumber = New frmSerialNumber
    Set g_fMicIn = New frmMicIn
    Set m_fPassword = New frmPassword
    Set m_fDetailedResults = New frmDetailedResults

    '@RWL
    If False = startup.LoadResultsLogIntoCollection Then
        startup.g_bStartupError = True
        Exit Sub
    End If
    '@RWL
                
    Load m_fAbout
    Load m_fContents
    Load m_fDiagDAQ
    Load m_fDiagPrinter
    Load m_fDiagGPS
    Load m_fDiagGSM
    Load m_fSetupForm
    If g_bAutoTest Then
        m_fSetupForm.chkAutotest.value = vbChecked
    Else
        m_fSetupForm.chkAutotest.value = vbUnchecked
    End If
    Load m_fUpdateSerialNumbers
    Load g_fSerialNumber
    Load g_fMicIn
    Load m_fPassword
    Load m_fDetailedResults
    Me.Top = g_iMainTop
    Me.Left = g_iMainLeft
    Me.Height = g_iMainHeight
    Me.Width = g_iMainWidth
    cmdResetCounter_Click
    chkDownload.value = g_iDownloadCheckSetting
    chkConfigureUnit.value = g_iConfigureModemIDCheckSetting
    chkTestFunctions.value = g_iTestFunctionsCheckSetting
    m_eTestingState = eTEST_APP_STARTED
    '   Load up all the currently defined Model Types
    For Each oModel In g_oModels.childNodes
        DoEvents
        If oModel.nodeType = NODE_ELEMENT Then
            cbxModelTypes.AddItem oModel.baseName & "   :" & oModel.selectSingleNode("ProdNo").Text
        End If
    Next oModel
    '   Create all of the columns for the ListView of Model Types.
    Set chrHeader = lvwTestResults.ColumnHeaders.Add()
    chrHeader.Text = "Test No:"
    'chrHeader.Width = lvwTestResults.Width \ 8 + 30
    chrHeader.Width = 0
    Set chrHeader = lvwTestResults.ColumnHeaders.Add()
    chrHeader.Text = "Description"
    chrHeader.Width = (lvwTestResults.Width \ 7) * 3 '+ 250
    Set chrHeader = lvwTestResults.ColumnHeaders.Add()
    chrHeader.Text = "Results"
    chrHeader.Width = (lvwTestResults.Width \ 7) * 3 - 520 '- 800
    Set chrHeader = lvwTestResults.ColumnHeaders.Add()
    chrHeader.Text = "Pass/Fail"
    chrHeader.Width = lvwTestResults.Width \ 7 + 150
    '   Create one row for each defined test.
    ReDim m_aszMessages(g_iNumOfTests) As String
    For i = 1 To g_iNumOfTests
        Set ltmRow = lvwTestResults.ListItems.Add
        ltmRow.Text = ""
        '   Create the sub-items for this row.
        For j = 1 To eITEM_NO_OF_ITEMS - 1
            Set lsiItem = ltmRow.ListSubItems.Add(, , "")
        Next j
        m_aszMessages(i) = ""
    Next i
    lblPassFail.Caption = "***Device not inserted***"
    lblPassFail.BackColor = vbDarkGrey
    m_szOldModeltype = ""
    m_bSuspendWirelessTimeout = False

    Caption = Caption & " v." & App.Major & "." & App.Minor & "." & App.Revision
    
    Set fso1 = New FileSystemObject   ''' Create the file handling object
    
    ' Make sure the test log file exists and is zero bytes long
    If fso1.FileExists(TEST_LOG_FN) Then
        ' Wipe the file contents
        Set ts = fso1.OpenTextFile(TEST_LOG_FN, ForWriting)
    Else
        ' FNF so create the file
        Set ts = fso1.CreateTextFile(TEST_LOG_FN)
    End If
            
    ts.WriteLine "Computer Name: " & ComputerName
    
    ' creates a new log file each day
    Set alldata = New FileSystemObject
    Dim Today As Date
    Dim FileName As String
    Dim LogDirPath As String
    Today = DateValue(Now)
    LogDirPath = DEFAULT_DATA_DIR + "daily_logs/"
    If Dir(LogDirPath, vbDirectory) = "" Then
        MkDir LogDirPath
    End If
    FileName = LogDirPath + Replace$(CStr(Today), "/", "-") + ".log"
    AllDataLogName = FileName
    If alldata.FileExists(AllDataLogName) Then
        'open the file, do not wipe contents
        Set ad = alldata.OpenTextFile(AllDataLogName, ForAppending)
    Else
        Set ad = alldata.CreateTextFile(AllDataLogName, ForAppending)
    End If
    
    ad.Close
    ts.Close

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode <> vbFormCode Then
        '   Tell the main routine to stop the application then let it take things
        '   down in a more controlled fashion.
        g_bStopApplication = True
        Cancel = 1
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set fso1 = Nothing   ''' Invalidate the file handling object
    
    tmrAutoTest.Enabled = False
    tmrAutoTest.Interval = 0
    tmrWaitWireless.Enabled = False
    tmrWaitWireless.Interval = 0
    If ComPortToPowerSupply.PortOpen = True Then
        ComPortToPowerSupply.PortOpen = False
    End If
    
    '   Close connections to the device if necessary
    OpenEliteComPort False

    g_iDownloadCheckSetting = chkDownload.value
    g_iConfigureModemIDCheckSetting = chkConfigureUnit.value
    g_iTestFunctionsCheckSetting = chkTestFunctions.value
    g_iMainTop = Me.Top
    g_iMainLeft = Me.Left
    g_iMainHeight = Me.Height
    g_iMainWidth = Me.Width
    '@RWL
    Set TestLogCollection = Nothing
    '@RWL
    Unload m_fAbout
    Set m_fAbout = Nothing
    Unload m_fContents
    Set m_fContents = Nothing
    Unload m_fDiagDAQ
    Set m_fDiagDAQ = Nothing
    Unload m_fDiagPrinter
    Set m_fDiagPrinter = Nothing
    Unload m_fDiagGPS
    Set m_fDiagGPS = Nothing
    Unload m_fDiagGSM
    Set m_fDiagGSM = Nothing
    Unload m_fSetupForm
    Set m_fSetupForm = Nothing
    Unload m_fUpdateSerialNumbers
    Set m_fUpdateSerialNumbers = Nothing
    If Not g_fSerialNumber Is Nothing Then
        Unload g_fSerialNumber
        Set g_fSerialNumber = Nothing
    End If
    If Not g_fMicIn Is Nothing Then
        Unload g_fMicIn
        Set g_fMicIn = Nothing
    End If
    Unload m_fPassword
    Set m_fPassword = Nothing
    Unload m_fDetailedResults
    Set m_fDetailedResults = Nothing

End Sub

Private Sub StartTest()
        
    Dim lstRow As ListItem
    Dim bOverallTestResult As Boolean
    m_ExecutedTests = ""
     
    Dim conf As IXMLDOMNode
    Dim buff As String
    
    ' get all possible tests from xml file
    Set conf = g_oTests.selectSingleNode("Configure")
    buff = conf.Text
    Set conf = g_oTests.selectSingleNode("Test")
    buff = buff + "," + conf.Text
    
    m_AllTests = Split(buff, ",")
    
    'open up log files before all the tests start
    
    Set ad = alldata.OpenTextFile(AllDataLogName, ForAppending)
    
    m_test_GSM_Sig = ""
    m_test_GPS_Sig = ""
    m_test_Current = ""
    
    '   Verify tester model, SIM and options settings
    If cbxModelTypes.Text = "" Then
        MsgBox "You MUST select a Model Type and SIM card before starting any tests."
        m_eTestingState = eTEST_COMPLETE
        Exit Sub
    End If
    If cbxSIMCards.Text = "" Then
        MsgBox "You MUST select a SIM card before starting any tests."
        m_eTestingState = eTEST_COMPLETE
        ' close log file then exit
        ad.Close
        Exit Sub
    End If
    
    If chkDownload.value = vbUnchecked And chkConfigureUnit.value = vbUnchecked And chkTestFunctions.value = vbUnchecked Then
        MsgBox "You MUST check at least one of either" & vbCrLf & _
               "'Download Firmware', 'Configure Unit' or 'Test Functions'" & vbCrLf & _
               "before starting any tests."
        m_eTestingState = eTEST_COMPLETE
        Exit Sub
    End If
    If chkDownload.value = vbChecked Then
        m_bDownloadFW = True
    End If
    
    '   Check if next serial number is valid
    If False = CheckNextIdNumber(etID_TYPE.eSERIAL_NUM) Then
        m_eTestingState = eTEST_COMPLETE
        Exit Sub
    End If
    
    '   Check if next IMEI number is valid
    If False = CheckNextIdNumber(etID_TYPE.eIMEI_NUM) Then
        m_eTestingState = eTEST_COMPLETE
        Exit Sub
    ElseIf String(Len(IMEI_NUM_FORMAT) + 1, "0") = g_szIMEI Then
        ' just in case something went wrong on last IMEI write
        g_szIMEI = ""
    End If
    
    ' Reset test run variables
    g_fIO.StarterInput = False
    g_fIO.IgnitionInput = False
    g_eElite_State = eElite_NULL
    g_szIMSI = ""
    
    m_eTestingState = eTESTING
    m_bDownloadFailed = False
    m_bConfigureSuccessful = True
    m_bDeviceFailure = False
    m_bApplicationRunning = False
    m_bApplicationHasBeenReset = False

    ' Reset timers
    m_iWirelessTimeout = 0
    m_bGPSTimeout = False
    m_bGPRSTimeout = False
    m_bVoltsTimeout = False
    m_bCurrentTimeout = False

    bOverallTestResult = True
    
    ' Reset form
    FormClear
    lblPassFail.BackColor = vbYellow
    If Not (chkTrialRun = vbChecked) Then
        lblPassFail.Caption = "TESTING"
    Else
        lblPassFail.Caption = "TESTING -- TRIAL RUN"
    End If
    
    ' Backup log file if it exists and overwrite/create new log file
    SerialComs.InitEliteComPortLogFile
    
    ' Start the timers running for the GPS, GPRS, and test time accumulated
    txtCycleTime.Text = ""
    tmrWaitWireless.Enabled = True
    m_dStartTime = Timer
    
    DoEvents
    ' Main loop for running the test
    For Each lstRow In lvwTestResults.ListItems
        
        If lstRow.Text <> "" Then
            '   Make certain that this item is visible
            lstRow.EnsureVisible
        End If
        
        m_iCurrentTest = Val(lstRow.Tag) - 1
        
        '   The timer set up for polling whether or not the test jig has
        '   been opened doesn't work when the tests are running (don't know
        '   why???).  So, we have to do our own polling here instead.
        If Not g_fIO.IsTesterClosed Then
            ' Test jig has been opened.  Force all left over tests to abort
            m_eTestingState = eTESTER_OPEN
        End If
        
        'AE 8/28/09 had to add this due to me doing the loops in my function instead of this for loop
        If m_iCurrentTest = g_iNumOfTests - 1 Then
            If RunTest(lstRow, , , lstRow.Index) = eRESULT_FAILED Then
                bOverallTestResult = False
            End If
            Exit For
        Else
            If RunTest(lstRow, , , lstRow.Index) = eRESULT_FAILED Then
                bOverallTestResult = False
            End If
            UpdateTestCycleTime
        End If

    Next lstRow
    
    ' If necessary close device COM log file and tester COM port to device
    OpenEliteComPort False
    
    If m_eTestingState = eTESTING Then
        If bOverallTestResult Then
            '   The DUT has passed all of the tests
            If Not (chkTrialRun = vbChecked) Then
                lblPassFail.Caption = "PASS"
                lblPassFail.BackColor = vbGreen
            Else
                lblPassFail.Caption = "PASS -- TRIAL RUN"
                lblPassFail.BackColor = vbGreen
            End If
            Globals.PassedParts = Globals.PassedParts + 1
            FormUpdateCounter
            ' Check if in Trial Run Mode.  If so, then do not write results,
            ' update the serial number, pr save anything in the XML file.
            If chkConfigureUnit.value = vbChecked And Not chkTrialRun.value = vbChecked Then
                '   Write out the results, print the labels, increment the serial
                '   number then save the serial number back to the XML tree.
                InventoryWriteResults
                g_clsBarcode.PrintScan 1
                g_lNextSerialNumber = g_lNextSerialNumber + 1
                txtSerialNumber.Text = Format(g_lNextSerialNumber, SERIAL_NUM_FORMAT)
                g_oSelectedModel.selectSingleNode("SerialNo").Attributes.getNamedItem("Next").Text = Format(g_lNextSerialNumber, "0000000")
                g_oEliteTester.Save ("ModelTypes.xml")
            End If
        Else
            '   The DUT has failed one or more tests
            If Not (chkTrialRun = vbChecked) Then
                lblPassFail.Caption = "FAIL"
                lblPassFail.BackColor = vbRed
            Else
                lblPassFail.Caption = "FAIL -- TRIAL RUN"
                lblPassFail.BackColor = vbRed
            End If
            Globals.FailedParts = Globals.FailedParts + 1
            FormUpdateCounter
        End If
        m_eTestingState = eTEST_COMPLETE
    Else
        '   The test jig has been opened in the middle of testing.  Testing has
        '   been aborted and the display has already been updated to show the
        '   open jig state.
    End If
    UpdateTestCycleTime
    
    lvwTestResults.SelectedItem.Selected = False

    tmrWaitWireless.Enabled = False
    
    'handle all end-of-test stuff with log files
    Dim X As Integer
    Dim y As Integer
    Dim found As Boolean
    Dim model_tests() As String
    
    ad.Write CStr(m_szSerialNum) + ","
    
    model_tests = Split(m_ExecutedTests, ",")
    
    For X = LBound(m_AllTests) To UBound(m_AllTests)
        Dim time_and_test() As String
        Dim time
        Dim test
        found = False
        For y = LBound(model_tests) To UBound(model_tests)
            If (model_tests(y) = "") Then
                Exit For
            End If
            time_and_test = Split(model_tests(y), "-")
            test = Trim$(time_and_test(0))
            time = time_and_test(1)
            Dim both As String
            both = Trim$(model_tests(y))
            If (m_AllTests(X) = test) Then
                found = True
                Exit For
            End If
        Next y
        If found Then
            ad.Write both + ","
        Else
            ad.Write ","
        End If
    Next X
    
    'reserve 5 spots for future tests
    ad.Write "r,r,r,r,r,"
    
    Dim testStatus As String
    If bOverallTestResult Then
        testStatus = "P"
    Else
        testStatus = "F"
    End If
    
    Dim timeStamp As String
 
    timeStamp = CStr(DateValue(Now)) + " " + CStr(TimeValue(Now))
    ad.Write CStr(g_szIMEI) + ","
    ad.Write CStr(g_szIMSI) + ","
    ad.Write CStr(g_szCCID) + ","
    ad.Write CStr(cbxModelTypes.Text) + ","
    ad.Write testStatus + ","
    ad.Write CStr(timeStamp) + ","
    ad.Write m_test_Current + ","
    ad.Write m_test_GSM_Sig + ","
    ad.Write m_test_GPS_Sig
    ad.Write vbCrLf
    ad.Close
    
    ' Turn off the power supply and close the COM port after each test cycle
#If DISABLE_EQUIP_CHECKS = 0 Then
    SerialComs.PSOutput = False
#Else
    If etPS_CTL.ePS_CTL_RS232 = g_ePS_Control Then
        SerialComs.PSOutput = False
'    Else
        '   No communications to PS
'        If vbOK = MsgBox("Power DUT off?", vbOKCancel) Then
'            '
'        Else
'            '
'        End If
    End If
#End If

End Sub

Private Function RunTest(lstRow As ListItem, Optional bRunTest As Boolean = True, Optional AddTestCnt As Integer = 1, Optional ByVal iLstRow As Integer = 0) As etRESULTS
    ' m_dStartTime used to keep track of elapsed time for each test
    Dim m_dStartTime As Double
    Dim TestName As String
    
    If lstRow.Tag = "" Then
        RunTest = eRESULT_NOT_TESTED
        Exit Function
    End If

    m_dStartTime = Timer
    
    Select Case Val(lstRow.Tag)
        Case 3
            TestName = "TEST- Download the firmware"
            '   Download the Firmware
            RunTest = FirmwareDownload(lstRow, bRunTest)
            m_bDownloadFW = False
            
        Case 7
            TestName = "TEST- Configure SMS Reply-to address"
            '   Configure SMS Reply-to address
            RunTest = ConfigureSMSreplyAddr(lstRow, bRunTest)
            
        Case 8
            TestName = "TEST- Configure SMS as the default comm. mode"
            '   Configure SMS as the default comm. mode
            RunTest = ConfigureSMSmode(lstRow, bRunTest)

        Case 9
            TestName = "TEST- Configure low power mode"
            '   Configure Low Power Mode
            RunTest = ConfigurePwrMode(lstRow, bRunTest)
            
        Case 10
            TestName = "TEST- Configure antitheft mode"
            ' Configure AntiTheft mode
            m_bApplicationRunning = True
            RunTest = ConfigureAntiTheft(lstRow, bRunTest)
            
        Case 11
            TestName = "TEST- Configure the APN and password"
            '   Configure the APN and password
            m_bApplicationRunning = True
            RunTest = ConfigureAPN(lstRow, bRunTest)

        Case 12
            TestName = "TEST- Configure the server IP addresses"
            '   Configure the Server IP addresses
            m_bApplicationRunning = True
            RunTest = ConfigureServers(lstRow, bRunTest)

        Case 14
            TestName = "TEST- Configure the service center address"
            '   Configure the Service Center Address
            m_bApplicationRunning = True
            RunTest = ConfigureSrvcCntrAddr(lstRow, bRunTest)

        Case 15
            TestName = "TEST- Configure the port address"
            '   Configure the Port Address
            m_bApplicationRunning = True
            RunTest = ConfigurePort(lstRow, bRunTest)

        Case 16
            TestName = "TEST- Configure the Serial #"
            '   Configure the Serial #
            m_bApplicationRunning = True
            RunTest = ConfigureIDs(lstRow, bRunTest, eSERIAL_NUM)

        Case 17
            TestName = "TEST- Configure the IMEI on the modem"
            ' Configure the IMEI on the Modem
            m_bSuspendWirelessTimeout = True
            RunTest = ConfigureIDs(lstRow, bRunTest, eIMEI_NUM)
            m_bSuspendWirelessTimeout = False

        Case 18
            TestName = "TEST- Save Changes to SPI Flash"
            RunTest = ConfigureFlash(lstRow, bRunTest)
            
        Case 19
            TestName = "TEST- Test Application firmware version"
            ' Test Application Firmware Version
            RunTest = TestAppVersion(lstRow, bRunTest)

        Case 20
            TestName = "TEST- Test the Serial #"
            '   Configure the Serial #
            m_bApplicationRunning = True
            RunTest = ConfigureIDs(lstRow, bRunTest, eSERIAL_NUM_READ_ONLY)

        Case 21
            TestName = "TEST- Test the IMEI from the modem"
            ' Test the IMEI from the Modem
            '   Read the IMEI number from the modem.  We need to suspend
            '   the wireless timeout because this test suspends the OpenAT
            '   application (which causes all wireless operations to be
            '   suspended as well).
            m_bApplicationRunning = True
            m_bSuspendWirelessTimeout = True
            RunTest = ConfigureIDs(lstRow, bRunTest, eIMEI_NUM_READ_ONLY)
            '   Resume the wireless timeout.
            m_bSuspendWirelessTimeout = False

        Case 22
            TestName = "TEST- Test the IMSI from the SIM card"
            ' Test the IMSI from the SIM Card
            '   Extract the IMSI number from the SIM card.  We need to suspend
            '   the wireless timeout because this test suspends the OpenAT
            '   application (which causes all wireless operations to be
            '   suspended as well).
            m_bApplicationRunning = True
            m_bSuspendWirelessTimeout = True
            RunTest = TestIMSI(lstRow, bRunTest)
            '   Resume the wireless timeout.
            m_bSuspendWirelessTimeout = False

        Case 23
            TestName = "TEST- Test the ICCID from the SIM card"
            ' Test the ICCID from the SIM card
            '   Extract the ICCID number from the SIM card.  We need to suspend
            '   the wireless timeout because this test suspends the OpenAT
            '   application (which causes all wireless operations to be
            '   suspended as well).
            m_bApplicationRunning = True
            m_bSuspendWirelessTimeout = True
            RunTest = TestICCID(lstRow, bRunTest)
            '   Resume the wireless timeout.
            m_bSuspendWirelessTimeout = False

        Case 24
            TestName = "TEST- Test the buzzer"
            '   Test the Buzzer
            m_bApplicationRunning = True
            RunTest = TestBuzzer(lstRow, bRunTest)

        Case 25
            TestName = "TEST- Test RF receiver"
            '   Test RF Receiver
            RunTest = TestRfRcvr(lstRow, bRunTest)

        Case 26
            TestName = "TEST- Test LEDs"
            '   Test LEDs
            RunTest = TestLED(lstRow, bRunTest)
            
        Case 27
            TestName = "TEST- Test ignition relay "
            '   Test Ignition Relay
            m_bApplicationRunning = True
            RunTest = TestIgnition(lstRow, bRunTest)

        Case 28
            TestName = "TEST- Test starter relay"
            '   Test Starter Relay
            m_bApplicationRunning = True
            RunTest = TestStarter(lstRow, bRunTest)

        Case 29
            TestName = "TEST- Test relays"
            '   Test Relays
            RunTest = TestRelays(lstRow, bRunTest)

        Case 30
            TestName = "TEST- Test the voltage regulator"
            '   Test the Voltage Regulator
            m_bApplicationRunning = True
            RunTest = TestVoltageRegulator(lstRow, bRunTest)

        Case 31
            TestName = "TEST- Test GPS Signal Strength"
            '   Test GPS Signal Strength
            m_bApplicationRunning = True
            If InStr(UCase(g_oSelectedModel.baseName), "-2X") > 0 Then
                RunTest = TestGpsAntenna(lstRow, bRunTest, True)
            Else
                RunTest = TestGpsAntenna(lstRow, bRunTest, False)
            End If

        Case 32
            TestName = "TEST- Test GSM Signal Strength"
            '   Test GSM Signal Strength
            '   We need to suspend the wireless timeout because this test
            '   suspends the OpenAT application (which causes all wireless
            '   operations to be suspended as well).
            m_bApplicationRunning = True
            m_bSuspendWirelessTimeout = True
            RunTest = VerifySignalStrengthGPRS(lstRow, bRunTest)
            '   Resume the wireless timeout.
            m_bSuspendWirelessTimeout = False

        Case 40, 41, 42
            TestName = "TEST- Power ON"
            ' Initial Power On
            RunTest = ConfigurePowerOn(lstRow, bRunTest)
        
        Case 43, 44
            TestName = "TEST- Wait for Initialization"
            '   TODO: Nab usollicited text from debug/test stream
            '   For now just wait
            RunTest = WaitForInitAfterReset(lstRow, bRunTest)
        
        Case 45, 46, 47, 48
            TestName = "TEST- Test DUT current draw"
            '   Test DUT Current Draw
            RunTest = TestCurrent(lstRow, bRunTest)

        Case 50, 51, 52
            TestName = "TEST- Setup and Test Serial Comms to Device"
            '   Configure Serial Comms to Device
            RunTest = ConfigDeviceCommunications(lstRow, bRunTest)
        
        Case 54
            TestName = "TEST- Setup and Test GSM Modem Serial Comms"
            '   Configure Serial Comms to GSM Modem
            RunTest = ConfigureGSM_Communications(lstRow, bRunTest)
        
        Case 55, 56
            TestName = "TEST- Test if modem initialized correctly"
            '   Test if modem initialized correctly
            RunTest = TestModemInit(lstRow, bRunTest)
        
        Case 60, 61, 62
            TestName = "TEST- Turn power OFF"
            '   Turn Power Off
            RunTest = ConfigurePowerOff(lstRow, bRunTest)
            
        Case 63
            TestName = "TEST- Reset"
            If bRunTest Then
                If m_eTestingState = eTESTER_OPEN Then
                    UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                    RunTest = eRESULT_NOT_TESTED
                ElseIf m_bDeviceFailure Then
                    UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
                    RunTest = eRESULT_NOT_TESTED
                Else
                    g_fMainTester.ComPortToElite.Output = "Mj" & vbCr
                    UpdateResults lstRow, "Sent reset command to DUT", eRESULT_PASSED
        '            UpdateResults lstRow, "Reset command sent to DUT", eRESULT_PASSED", eRESULT_PASSED
                    RunTest = eRESULT_PASSED
                End If
            End If
        
        Case 65
            TestName = "TEST- Test 'Fail Safe' mode"
            ' Test "Fail Safe" mode (Test relays after power off)
            '   Check to make certain that the relays have fallen open with the
            '   removal of the power.
            m_bApplicationRunning = False
            RunTest = TestRelayAfterPowerDown(lstRow, bRunTest)
            
        Case 66
            TestName = "TEST- Test Super Cap Voltage"
            '   Turn Power Off
            RunTest = TestSuperCap(lstRow, bRunTest)
            
        Case Else
            RunTest = AddTests(lstRow, bRunTest, AddTestCnt, iLstRow)
            '   This should never happen
            'MsgBox "Invalid test tag (" & lstRow.Tag & ")"
            'lstRow.ListSubItems(eITEM_DESCRIPTION).Text = "***<Internal error>"
            'RunTest = eRESULT_ERROR
    End Select
    
    ' Save the elapsed time to the tests log
    m_dStartTime = Timer - m_dStartTime
    If m_dStartTime < 0 Then
        m_dStartTime = m_dStartTime + 86400
    End If
            
    If bRunTest Then
        'SFS log file
        Dim time As Double
        time = Format(m_dStartTime, "0.00")
        m_ExecutedTests = m_ExecutedTests + lstRow.Tag + "-" + CStr(time) + ","
        
        'JMB log file
        Set ts = fso1.OpenTextFile(TEST_LOG_FN, ForAppending)
        m_aszMessages(Val(lstRow.Tag) - 1) = m_aszMessages(Val(lstRow.Tag) - 1) + vbCrLf & "Elapsed Time: " & Format(m_dStartTime, "0.00") & "s"
        ts.WriteLine m_aszMessages(Val(lstRow.Tag) - 1)
        
        ts.Close
    Else
        lstRow.SubItems(eITEM_DESCRIPTION) = Right(TestName, Len(TestName) - 6)
        ' Unit test for test names
        ' Dim szTestNames(CInt(g_oTests.selectSingleNode("NumOfTests").Text))
        Dim szTestNames(70)
        Dim oNode As IXMLDOMNode
        For Each oNode In g_oTests.childNodes
            If 1 = InStr(1, Trim(oNode.Text), Trim(lstRow.Tag)) Then
                If 0 <> InStr(1, oNode.Text, Right(TestName, Len(TestName) - 6)) Then
                ' Found Test number and TestName
                    Exit For
                End If
            End If
        Next oNode
        If Not oNode Is Nothing Then
            szTestNames(CInt(lstRow.Tag)) = g_oTests.selectSingleNode("NumOfTests").Text
        Else
            ' Failed unit test
            startup.Delay 0.1
        End If
    End If

    ToggleEliteComPortLogFileCloseThenOpen
    
End Function

Private Function CheckNextIdNumber(Optional eIdType As etID_TYPE = eSERIAL_NUM) As Boolean

    Dim SNLeft As Long
    Dim blnRecycle As Boolean
    Dim RecycledSerNum As Long
    Dim RecycledSN As Long
    Dim lNextIdNum As Long
    Dim szIdNum As String
    Dim lEndIdNum As Long
    Dim szIdType As String
    Dim oNode As IXMLDOMNode
    Dim oAttributes As IXMLDOMNamedNodeMap
    
    CheckNextIdNumber = False
    
    If etID_TYPE.eSERIAL_NUM = eIdType Then
        lNextIdNum = g_lNextSerialNumber
        lEndIdNum = g_lEndingSerialNumber
        szIdType = "Serial"
    Else
        Set oNode = g_oSettings.selectSingleNode("TestSystem").childNodes(g_iTestStationID)
        Set oNode = oNode.selectSingleNode("IMEI")
        If oNode.selectSingleNode("Input_Method").Text = "BY_MODEL" Then
            Set oNode = g_oSelectedModel.selectSingleNode("IMEI")
        End If
        
        lNextIdNum = oNode.selectSingleNode("SNR").Attributes.getNamedItem("Next").Text
        lEndIdNum = oNode.selectSingleNode("SNR").Attributes.getNamedItem("End").Text
        szIdType = "IMEI"
    End If
    
    If lNextIdNum > lEndIdNum Then
        Beep
        MsgBox "No More " & szIdType & " Numbers Available." & vbCrLf & vbCrLf & _
               "Configuration File Must Be Updated With New Customer Approved " & szIdType & " Numbers." & vbCrLf & vbCrLf & _
               "Test Aborted!", vbCritical
    ElseIf lNextIdNum + 100 > lEndIdNum Then
        SNLeft = (lEndIdNum - lNextIdNum) + 1
        Beep
        MsgBox "There Are Only " & CStr(SNLeft) & " Serial Numbers Left." & vbCrLf & vbCrLf & _
               "Configuration File Needs To Be Updated With New Customer Approved " & szIdType & " Numbers." & vbCrLf & vbCrLf & _
               "Continuing with test.", vbCritical
        CheckNextIdNumber = True
    Else
        CheckNextIdNumber = True
    End If

End Function

'AE 10/17/09 added to test additional tests, if any
Private Function AddTests(lstRow As ListItem, bRunTest As Boolean, AddTestCnt As Integer, iLstRow As Integer) As etRESULTS
    Dim result As String
    Dim command As String
    Dim Timeout As Long
    Dim overallpass As Integer
    overallpass = 0
    Dim i As Integer
    i = 1
    Dim j As IXMLDOMNode
    Dim oAttributes As IXMLDOMNamedNodeMap
    Set oAttributes = g_oAddTests.selectSingleNode("ADD" & AddTestCnt).Attributes
    Dim tmplstRow As ListItem
     
 If Not bRunTest Then
    lstRow.SubItems(eITEM_DESCRIPTION) = oAttributes.getNamedItem("Description").Text
    AddTests = eRESULT_NOT_TESTED
    Else
        'check to see if there is a command
        command = oAttributes.getNamedItem("Command").Text
        If command = "" Then
        Exit Function
        End If
         
        Set tmplstRow = lstRow
        'turn on power again and wait for POST
        UpdateResults tmplstRow, "Starting up DUT again", eRESULT_TESTING
        SerialComs.PSOutput = True
        startup.Delay 12#
        UpdateResults tmplstRow, "Opening COM4 again", eRESULT_TESTING
        
        For Each j In g_oAddTests.childNodes
        
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults tmplstRow, "JIG OPENED: Abort test", eRESULT_FAILED
            overallpass = 1
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults tmplstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
            overallpass = 1
        Else
        
         Set oAttributes = g_oAddTests.selectSingleNode("ADD" & i).Attributes
         command = oAttributes.getNamedItem("Command").Text
         If command = "" Then
         Exit Function
         Else
         
            If oAttributes.getNamedItem("Time").Text = "" Then
            Timeout = SERIAL_COMS_TIMEOUT
            Else
            Timeout = CLng(oAttributes.getNamedItem("Time").Text)
            End If

            result = oAttributes.getNamedItem("Result").Text
            
            If SerialComs.AddTest(command, Timeout, result) Then
                If result = g_szAddResult Then
                    UpdateResults tmplstRow, "Result: " & g_szAddResult, eRESULT_PASSED
                Else
                    UpdateResults tmplstRow, "Result: " & g_szAddResult, eRESULT_FAILED
                    overallpass = 1
                End If
            Else
                UpdateResults tmplstRow, "Failed to execute command: " & command, eRESULT_FAILED
                overallpass = 1
            End If
   
         End If
         'get next lstRow
         Set tmplstRow = lvwTestResults.ListItems(iLstRow + i)
         i = i + 1
                
        End If
        Next j
        
        If overallpass <> 1 Then
        AddTests = eRESULT_PASSED
        Else
        AddTests = eRESULT_FAILED
        End If
End If

End Function
'NO LONGER USED
Private Function TestCurrentOLDTEST(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    Dim dMinCurrent As Double
    Dim dMaxCurrent As Double
    Dim dDUTCurrentDraw As Double
    Dim oMinAttribute As IXMLDOMAttribute
    Dim oMaxAttribute As IXMLDOMAttribute
    Dim oAttributes As IXMLDOMNamedNodeMap

    TestCurrentOLDTEST = eRESULT_NOT_TESTED
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test DUT Current Draw"
        TestCurrentOLDTEST = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
        Else
            'Measure Current Draw
            If etPS_CTL.ePS_CTL_RS232 <> g_ePS_Control Then
                '   No communications to PS
                UpdateResults lstRow, "Measuring current", eRESULT_NOT_TESTED
                TestCurrentOLDTEST = eRESULT_NOT_TESTED
            Else
                    UpdateResults lstRow, "Measuring current...", eRESULT_TESTING
                    While Not m_bCurrentTimeout
                        DoEvents
                    Wend
                    dDUTCurrentDraw = SerialComs.PowerSupply1_Current
                    '   Get the minimum and maximum current draws allowable for the selected Model Type
                    Set oAttributes = g_oSelectedModel.selectSingleNode("Current").Attributes
                    Set oMinAttribute = oAttributes.getNamedItem("Min")
                    Set oMaxAttribute = oAttributes.getNamedItem("Max")
                    dMinCurrent = CDbl(oMinAttribute.Text)
                    dMaxCurrent = CDbl(oMaxAttribute.Text)
                    If dDUTCurrentDraw <= dMaxCurrent And dDUTCurrentDraw >= dMinCurrent Then
                        TestCurrentOLDTEST = eRESULT_PASSED
                        UpdateResults lstRow, Format(dDUTCurrentDraw * 1000, "0.00" & " mA"), eRESULT_PASSED
                    Else
                        TestCurrentOLDTEST = eRESULT_FAILED
                        UpdateResults lstRow, Format(dDUTCurrentDraw * 1000, "0.00" & " mA") & " not between " & dMinCurrent & " and " & dMaxCurrent & "A", eRESULT_FAILED
                        m_bDeviceFailure = True
                    End If
                End If
        End If
    End If

End Function

Private Function TestCurrent(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    
    Dim dMinCurrent, dMaxCurrent, dWait As Double
    Dim dMaxInRushCurrent, dInRushDuration, dDUTCurrentDraw As Double
    Dim oAttributes As IXMLDOMNamedNodeMap
    Dim Index As Integer
    Dim total_current As Double
    Dim avg_good_current As Double
    
    m_test_Current = "'Not Measured'"
    avg_good_current = 0
    total_current = 0
    
    TestCurrent = eRESULT_NOT_TESTED
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test DUT Current Draw"
        TestCurrent = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
        Else
            'Measure Current Draw
            If etPS_CTL.ePS_CTL_RS232 <> g_ePS_Control Then
                '   No communications to PS
                UpdateResults lstRow, "Measuring current", eRESULT_NOT_TESTED
                TestCurrent = eRESULT_NOT_TESTED
            Else
                While Not m_bCurrentTimeout
                    DoEvents
                Wend
                
                ' get min/max current criteria from XML file
                Set oAttributes = g_oSelectedModel.selectSingleNode("Power").selectSingleNode("Nominal_Current").Attributes
                dMinCurrent = CDbl(oAttributes.getNamedItem("Min").Text)
                dMaxCurrent = CDbl(oAttributes.getNamedItem("Max").Text)
                
                ' get duration, in seconds, of how long to average the current readings
                dWait = CDbl(oAttributes.getNamedItem("Wait").Text)
                
                Index = 0
                UpdateResults lstRow, "Measuring nominal current...", eRESULT_TESTING
                
                total_current = 0#
                Do While Index < dWait
                    ' get current reading from power suply
                    dDUTCurrentDraw = SerialComs.PowerSupply1_Current
                    If -9999 = dDUTCurrentDraw Then
                        ' Throw out the reading from the average
                        dWait = dWait - 1
                        dDUTCurrentDraw = 0
                    Else
                        UpdateResults lstRow, "Nominal current measurement " & Index & ": " & Format(dDUTCurrentDraw * 1000, "0.00" & " mA"), eRESULT_TESTING
                    End If
                    total_current = total_current + dDUTCurrentDraw

                    If 0 <> dWait And Index = dWait - 1 Then
                        avg_good_current = total_current / dWait
                        m_test_Current = Format(avg_good_current * 1000, "0.00" & " mA")
                        If avg_good_current <= dMaxCurrent And avg_good_current >= dMinCurrent Then
                            TestCurrent = eRESULT_PASSED
                            UpdateResults lstRow, "Nominal current: " & m_test_Current, eRESULT_PASSED
                            Exit Function
                        End If
                    End If
                    Index = Index + 1
                    startup.Delay 1#
                Loop
                If 0 <> dWait Then
                    m_test_Current = Format(total_current / dWait * 1000, "0.00" & " mA")
                End If
                TestCurrent = eRESULT_FAILED
                UpdateResults lstRow, m_test_Current & " is not in nominal current range", eRESULT_FAILED
                m_bDeviceFailure = True
            End If
        End If
    End If

End Function

Private Function WaitForInitAfterReset(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    Dim szEventMsg(2) As String
    Dim szEventDisplayMsg(2) As String
    Dim iWaitForPostTimeout(2) As Integer
    Dim iIdx As Integer
    Dim iCnt As Integer

    ' m_szDevPOST_OK_Message set by serialcoms.init
    szEventMsg(0) = ""
    szEventMsg(1) = SerialComs.m_szDevPOST_Message
    szEventMsg(2) = m_szStateResponses(eElite_INIT_COMPLETE)
    
    szEventDisplayMsg(0) = "Waiting for reset to complete..."
    szEventDisplayMsg(1) = "System Initialization"
    szEventDisplayMsg(2) = "Modem Initialization"
    
    iWaitForPostTimeout(0) = 15
    iWaitForPostTimeout(1) = 30
    iWaitForPostTimeout(2) = 30
    
    WaitForInitAfterReset = eRESULT_NOT_TESTED
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Wait for " & szEventMsg(0)
        WaitForInitAfterReset = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            WaitForInitAfterReset = eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        Else
            WaitForInitAfterReset = eRESULT_TESTING
            UpdateResults lstRow, "Reset DUT and wait for event...", eRESULT_TESTING
            
            iIdx = 0
'            Do
                If 0 = iIdx Then
                    ' Just wait
                    For iCnt = iWaitForPostTimeout(0) To 0 Step -1
                        UpdateResults lstRow, szEventDisplayMsg(iIdx) & " " & iCnt, eRESULT_TESTING
                        startup.Delay 1#
                    Next iCnt
                    UpdateResults lstRow, "Reset wait timer expired", eRESULT_PASSED
                    WaitForInitAfterReset = eRESULT_PASSED
                Else
                    ' TODO: Get nabbing working so we can detect these events
                    If False = SerialComs.WaitForPost(iIdx) Then
                        WaitForInitAfterReset = eRESULT_FAILED
                        m_bDeviceFailure = True
                        UpdateResults lstRow, "Failed to detect " & szEventDisplayMsg(iIdx) & " event", eRESULT_FAILED
                    Else
                        g_fMainTester.LogMessage ("Detected " & szEventDisplayMsg(iIdx) & " event")
                    End If
                End If
'            Loop While iIdx = UBound(szEventMsg)
        End If
        
        If eRESULT_TESTING = WaitForInitAfterReset Then
            WaitForInitAfterReset = eRESULT_PASSED
            UpdateResults lstRow, "Initialization verified, eRESULT_PASSED", eRESULT_PASSED
        End If
    End If
End Function

Private Function ConfigDeviceCommunications(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    ConfigDeviceCommunications = eRESULT_NOT_TESTED
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Configure Serial Comms to Device"
        ConfigDeviceCommunications = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigDeviceCommunications = eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        Else
            Dim iRetries As Integer
            
            '   Connect the serial COM port to the device
            ConfigDeviceCommunications = eRESULT_TESTING
            UpdateResults lstRow, "Opening COM port...", eRESULT_TESTING
            While Not g_bOpenATRunning
                '   Wait here until the OpenAT OS has had a chance to boot up.
                DoEvents
            Wend
            
            If False = SerialComs.OpenEliteComPort(True) Or True = g_bQuitProgram Then
                'Just try again to be sure
                startup.Delay 1#
                Debug.Print "OpenEliteComPort function failed or g_bQuitProgram flag is set"
                If False = SerialComs.OpenEliteComPort(True) Or True = g_bQuitProgram Then
                    ' Give up
                    m_bDeviceFailure = True
                    ConfigDeviceCommunications = eRESULT_FAILED
                    UpdateResults lstRow, "Serial COM Port is not open", eRESULT_FAILED
                    Exit Function
                End If
            End If
            
            ' Com port is open so see if we can talk to the device
            
            UpdateResults lstRow, "Configuring serial comms to device...", eRESULT_TESTING
            If False = ConfigureDeviceForTesting Or True = g_bQuitProgram Then
                ' May indicate programming failure but still possible that we just need to wait
                Debug.Print "ConfigureDeviceForTesting function failed or g_bQuitProgram flag is set"
                g_fMainTester.LogMessage ("No comms. Toggling COM port")
                
                If False = SerialComs.OpenEliteComPort(True) Then
                    m_bDeviceFailure = True
                    ConfigDeviceCommunications = eRESULT_FAILED
                    UpdateResults lstRow, "Serial COM Port is not open", eRESULT_FAILED
                    Exit Function
                ElseIf False = ConfigureDeviceForTesting Then
                    m_bDeviceFailure = True
                    ConfigDeviceCommunications = eRESULT_FAILED
                    UpdateResults lstRow, "Failed to configure comms to device", eRESULT_FAILED
                    Exit Function
                End If
            End If
            
            If eRESULT_TESTING = ConfigDeviceCommunications Then
                ConfigDeviceCommunications = eRESULT_PASSED
                UpdateResults lstRow, "Configured serial comms to device", eRESULT_PASSED
            End If
        End If
    End If
    
End Function
                    
Public Function UpdateTestCycleTime()
    ' Limited to 24 hour max
    Dim dCurrentTimer As Double
    
    dCurrentTimer = Timer
    If dCurrentTimer < m_dStartTime Then
        ' Timer rolled over midnight and reset to 0
        m_dStartTime = m_dStartTime - 86400
    End If
    txtCycleTime.Text = Format(dCurrentTimer - m_dStartTime, "0.00") & " Sec"
End Function

Private Function TestModemInit(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    TestModemInit = eRESULT_NOT_TESTED
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify Serial Comms to GSM Modem"
    Else
        
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        Else
            Dim iRetries As Integer
            Dim bSerialCmdStatus As Boolean
            TestModemInit = eRESULT_TESTING

            If eRESULT_TESTING = TestModemInit Then
                UpdateResults lstRow, "Querying GSM modem...", eRESULT_TESTING
                If False = SerialComs.VerifyOpenAT Then
                    ' We have waited at least 20 seconds already and still no response
                    ' to the AT command but we did get a response from the non-AT command
                    ' to configure from testing
                    m_bDeviceFailure = True
                    TestModemInit = eRESULT_FAILED
                    UpdateResults lstRow, "GSM modem not responding", eRESULT_FAILED
                Else
                    TestModemInit = eRESULT_TESTING
                End If
            End If
             
            If eRESULT_TESTING = TestModemInit Then
                ' This used to only take one second but now will take up to 3 seconds to fail
                For iRetries = MODEM_INIT_MAX_WAIT / 3 To 1 Step -1
                    UpdateResults lstRow, "Waiting for modem..." & iRetries, eRESULT_TESTING
                    If TestModemStatus Then
                        UpdateResults lstRow, "GSM modem responding", eRESULT_TESTING
                        TestModemInit = eRESULT_PASSED
                        Exit For
                    End If
                    UpdateTestCycleTime
                    DoEvents
                    ' DoEvents not catching tester open here so calling
                    ' g_fIO.IsTesterClosed directly
                    If eTESTER_OPEN = m_eTestingState Or False = g_fIO.IsTesterClosed Then
                        UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                        TestModemInit = eRESULT_FAILED
                        Exit For
                    Else
                        startup.Delay 1#
                    End If
                Next iRetries
            End If
            
            If eRESULT_PASSED <> TestModemInit Then
                ' Modem failed to initialize
                m_bDeviceFailure = True
                TestModemInit = eRESULT_FAILED
                UpdateResults lstRow, "GSM modem not responding", eRESULT_FAILED
            Else
                UpdateResults lstRow, "Verified Serial Comms to GSM Modem", eRESULT_PASSED
                TestModemInit = eRESULT_PASSED
            End If
        End If
    End If

End Function


Public Function PicFW(szAction As String) As Long
    Dim picProgBatchFN As String
    Dim lHandle As Long
    Dim WFD As WIN32_FIND_DATA
    Dim szFullyQualifiedFN

    On Error GoTo FileInitErrorPicProg:

    picProgBatchFN = PIC_PROG_BATCH_FN
    lHandle = FindFirstFile(picProgBatchFN, WFD)
    If lHandle < 1 Then
        GoTo FileInitErrorPicProg:
    End If

    szFullyQualifiedFN = g_szDataDirPath & g_szFirmwareFileName & ".hex"
    PicFW = ExecCmd(Chr(34) & picProgBatchFN & Chr(34) & Chr(32) & szAction & Chr(32) & "24FJ256GA106" & Chr(32) & szFullyQualifiedFN)
    
    Exit Function
FileInitErrorPicProg:
    ' MsgBox "Failed to execute batch file " & szFullyQualifiedFN, vbCritical
    g_fMainTester.LogMessage ("Failed to execute batch file " & szFullyQualifiedFN)

End Function

Private Function FirmwareDownload(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Download Firmware"
        FirmwareDownload = eRESULT_NOT_TESTED
    Else
        FirmwareDownload = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
        ElseIf Not (chkTrialRun.value = vbChecked) Then
            '   Download the firmware
            Dim iRetries As Integer

            FirmwareDownload = eRESULT_TESTING
            UpdateResults lstRow, "Downloading...", eRESULT_TESTING

            If 0 = PicFW("p") Then
                ' To be deterministic need to cycle power
                UpdateResults lstRow, "F/W=" & g_szFirmwareFileName, eRESULT_PASSED
                FirmwareDownload = eRESULT_PASSED
            Else
                UpdateResults lstRow, "F/W not downloaded", eRESULT_FAILED
                FirmwareDownload = eRESULT_FAILED
                m_bDeviceFailure = True
            End If
        Else
            'trial run selected
            UpdateResults lstRow, "Trial run - skipped", eRESULT_NOT_TESTED
            FirmwareDownload = eRESULT_NOT_TESTED
        End If
    End If

End Function


Private Function TestVoltageRegulator(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim dVolts As Double
    Dim iVolts As Integer
    Dim iRetries As Integer

    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test the Voltage Regulator"
        TestVoltageRegulator = eRESULT_NOT_TESTED
    Else
        TestVoltageRegulator = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestVoltageRegulator = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            TestVoltageRegulator = eRESULT_NOT_TESTED
#If DISABLE_EQUIP_CHECKS = 1 Then
        ElseIf etPS_CTL.ePS_CTL_RS232 <> g_ePS_Control Then
            '   No communications to PS
            UpdateResults lstRow, "Voltages acceptable", eRESULT_NOT_TESTED
            TestVoltageRegulator = eRESULT_NOT_TESTED
#End If
        Else
            '   Check the voltage regulator
            UpdateResults lstRow, "Getting voltage from regulator...", eRESULT_TESTING
            TestVoltageRegulator = eRESULT_TESTING
            '   Wait until the voltage regulator has had a chance to settle
            '   down since the application started.
            While Not m_bVoltsTimeout
                DoEvents
            Wend
            '   Start by getting the voltage directly from the power supply
            dVolts = SerialComs.PowerSupply1_Volts * 1000
            '   Now get the volts from the voltage regultor (via the unit).
            For iRetries = 1 To 3
                If SerialComs.GetVolts(iVolts) Then
                    If CDbl(iVolts) >= dVolts - 2000 And CDbl(iVolts) <= dVolts + 2000 Then
                        UpdateResults lstRow, "DUT: " & CDbl(iVolts) / 1000 & "V +-2V of PS: " & dVolts / 1000 & "V", eRESULT_PASSED
                        TestVoltageRegulator = eRESULT_PASSED
                        Exit For
                    End If
                Else
                    UpdateResults lstRow, "Failed to get voltage", eRESULT_FAILED
                    TestVoltageRegulator = eRESULT_FAILED
                    Exit For
                End If
                If m_eTestingState = eTESTER_OPEN Then
                    UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                    TestVoltageRegulator = eRESULT_FAILED
                    Exit For
                Else
                    startup.Delay 2#
                End If
            Next iRetries
            If TestVoltageRegulator = eRESULT_TESTING Then
                UpdateResults lstRow, "DUT: " & CDbl(iVolts) / 1000 & "V NOT +-2V of PS: " & dVolts / 1000 & "V", eRESULT_FAILED
                TestVoltageRegulator = eRESULT_FAILED
            End If
            If TestVoltageRegulator = eRESULT_FAILED Then
                '   Stop all further tests with the unit
                m_bDeviceFailure = True
            End If
        End If
    End If

End Function
Private Function TestModemStatus() As Boolean
    Dim szModemStatus As String
    
    ' Because this command turns off unsolicited messages, it is typical that
    ' the tester receives unsolicited messages before the device can process
    ' the command. Run with minimal timeout and then reissue the command
    ' to verify communications with the device
    
    ' TestModemStatus = SerialComs.SendATCommand(eElite_VerboseOff, 1)

    TestModemStatus = False
    If SerialComs.ModemStatus(szModemStatus) Then
        If "READY" = szModemStatus Then
            TestModemStatus = True
        ElseIf "WAIT" = szModemStatus Then
            g_fMainTester.LogMessage "Modem not ready"
        Else
            g_fMainTester.LogMessage "Undefined modem status"
        End If
    Else
        szModemStatus = ""
    End If
End Function
Private Function ConfigureSerialNum(lstRow As ListItem, Optional bReadWrite As Boolean = True) As etRESULTS
    Dim iRetries As Integer
    Dim szSerialNum As String
    Dim bReadOnly As Boolean
    
    bReadOnly = Not bReadWrite
    
    ConfigureSerialNum = eRESULT_TESTING
    
    If (True = bReadOnly Or vbChecked = chkTrialRun.value Or vbUnchecked = Me.chkConfigureUnit.value) Then
        iRetries = 0
        bReadOnly = True
    Else
        
        '   Configure the Serial Number
        szSerialNum = g_lNextSerialNumber
        UpdateResults lstRow, "Setting Serial Number...", eRESULT_TESTING
        For iRetries = 1 To 4
            If SerialComs.SetSerialNum(g_lNextSerialNumber) Then
                iRetries = 0
                Exit For
            End If
            If m_eTestingState = eTESTER_OPEN Then
                UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                iRetries = 1
                Exit For
            Else
                startup.Delay 1#
            End If
        Next iRetries
    End If
    
    szSerialNum = ""
    If iRetries <> 0 Then
        ConfigureSerialNum = eRESULT_FAILED
    Else
        If True = bReadOnly Then
            '   Read the Serial Number
            UpdateResults lstRow, "Reading Serial Number...", eRESULT_TESTING
        Else
            '   Confirm that the Serial Number is set properly
            UpdateResults lstRow, "Confirming Serial Number...", eRESULT_TESTING
        End If
        
        For iRetries = 1 To 4
            If SerialComs.GetSerialNum(szSerialNum) Then
                If CLng(szSerialNum) = g_lNextSerialNumber Or vbUnchecked = Me.chkConfigureUnit.value Then
                    If vbUnchecked = Me.chkConfigureUnit.value Then
                        UpdateResults lstRow, "Serial Number = " & Format(CLng(szSerialNum), "0000000"), eRESULT_PASSED
                    Else
                        UpdateResults lstRow, "Serial Number " & Format(CLng(szSerialNum), "0000000") & " verified", eRESULT_PASSED
                    End If
                    ' TODO: Don't think this is necessary (don't think compare ever happens)
                    m_szSerialNum = szSerialNum ' Set global m_szSerialNum so that it can be compared with serial number later on
                    ConfigureSerialNum = eRESULT_PASSED
                    Exit For
                End If
            End If
            If m_eTestingState = eTESTER_OPEN Then
                UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                ConfigureSerialNum = eRESULT_FAILED
                Exit For
            Else
                startup.Delay 1#
            End If
        Next iRetries
    End If
    
End Function

Private Function ConfigureBGS2ModemID(lstRow As ListItem, Optional eID_Type As etID_TYPE = eIMEI_NUM) As etRESULTS
    Dim iRetries As Integer
    Dim szIMEINum As String
    Dim iSNR_Suffix As Integer
    Dim bVerifyIMEI_Write As Boolean
    Dim lIMEINumber As stIMEI_NUM
    Dim szPassTimeIMEI_TAC As String
    Dim oNode As IXMLDOMNode
    Dim szOpStat As String
    
    szIMEINum = ""
    
    Set oNode = g_oSettings.selectSingleNode("TestSystem").childNodes(g_iTestStationID)
    Set oNode = oNode.selectSingleNode("IMEI")
    If oNode.selectSingleNode("Input_Method").Text = "BY_MODEL" Then
        Set oNode = g_oSelectedModel.selectSingleNode("IMEI")
    End If
    
    szPassTimeIMEI_TAC = oNode.selectSingleNode("TAC").Text

    ' Check if g_szIMEI is a string of all zeros
    If String(Len(IMEI_NUM_FORMAT) + 1, "0") = g_szIMEI Then
        ' String of all zeros indicates that we are verifying
        ' an IMEI number that was just written
        bVerifyIMEI_Write = True
    Else
        bVerifyIMEI_Write = False
    End If

    ' Read IMEI number and see if it needs to be set
    UpdateResults lstRow, "Reading IMEI Number...", eRESULT_TESTING
    
    For iRetries = 3 To 1 Step -1
        ConfigureBGS2ModemID = eRESULT_TESTING
        szIMEINum = ""
        
        If False = SerialComs.BGS2ModemID(szIMEINum, eElite_BGS2_MODEM_ID) Then
            ConfigureBGS2ModemID = eRESULT_ERROR
        ElseIf (vbUnchecked = chkTrialRun.value And vbChecked = Me.chkConfigureUnit.value And eIMEI_NUM_READ_ONLY <> eID_Type) Then
            ' Not in read only mode. Verify no serial IMEI number has been written to device yet
            If "NO IMEI" = szIMEINum Or ("^SGSN:" = Left(szIMEINum, 6) And ",NO IMEI" = Right(szIMEINum, 8)) Then
                '   Configure the Modem ID once and only once
                UpdateResults lstRow, "Setting IMEI Number...", eRESULT_TESTING
                g_lNextIMEINumber.TAC = szPassTimeIMEI_TAC
                g_lNextIMEINumber.SNR = CLng(Val(oNode.selectSingleNode("SNR").Attributes.getNamedItem("Next").Text))
                g_lNextIMEINumber.SNR_Suffix = CLng(Val(oNode.selectSingleNode("SV").Text))
                
                szIMEINum = Format(g_lNextIMEINumber.TAC, IMEI_NUM_TAC_FORMAT)
                szIMEINum = szIMEINum & Format(g_lNextIMEINumber.SNR, IMEI_NUM_SNR_FORMAT)
                szIMEINum = Format(szIMEINum, IMEI_NUM_FORMAT)
                ' szIMEINum = szIMEINum & Format(g_lNextIMEINumber.SNR_Suffix, IMEI_NUM_SV_FORMAT)
                
                ' ASSUMING we are testing with eElite_BGS2_MODEM_ID (i.e., not BGS3)
                ' so need to append check digit to IMEI number and set the svn
                ' Compute the check digit
                Dim szIMEINumAndChkDigit As String
                szIMEINumAndChkDigit = util_GetImeiCheckDigit(szIMEINum)
                
                If 0 = Len(szIMEINumAndChkDigit) Then
                    ConfigureBGS2ModemID = eRESULT_ERROR
                Else
                    ' For BGS2 need to send at commands directly to modem so use bypass mode
                    UpdateResults lstRow, "Entering bypass mode...", eRESULT_TESTING
                    SerialComs.SendATCommand etELITE_STATES.eElite_ENABLE_BYPASS, 1
                    
                    If False = SerialComs.BGS2ModemID(szIMEINum & szIMEINumAndChkDigit, eElite_BGS2_MODEM_ID) Then
                        ' Write the IMEI number to the modem
                        UpdateResults lstRow, "Failed to set IMEI", eRESULT_FAILED
                        ConfigureBGS2ModemID = eRESULT_FAILED
                    Else
                        Dim szSVN As String
                        szSVN = g_lNextIMEINumber.SNR_Suffix
                        If False = SerialComs.BGS2ModemSVN(szSVN) Then
                            ' TODO: This needs to be outside of the IMEI write function since if it failed we
                            ' may want to try again
                            UpdateResults lstRow, "Failed to set modem SVN", eRESULT_FAILED
                            ConfigureBGS2ModemID = eRESULT_FAILED
                        Else
                            g_fMainTester.LogMessage "Waiting 10 seconds for soft modem restart"
                
                            Dim iWaitForSoftReset As Integer
                            
                            ' Not sure what to do if reset command failed and since we
                            ' verify IMEI number later, not really necessarry
                            SerialComs.SendATCommand eElite_RESET_GSM_MODULE
                            
                            For iWaitForSoftReset = IMEI_WRITE_WAIT To 1 Step -1
                                UpdateResults lstRow, "Waiting for soft modem restart..." & iWaitForSoftReset, eRESULT_TESTING
                                startup.Delay 1#
                                If m_eTestingState = eTESTER_OPEN Then
                                    UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                                    ConfigureBGS2ModemID = eRESULT_FAILED
                                    Exit For
                                End If
                            Next iWaitForSoftReset
                            ConfigureBGS2ModemID = eRESULT_PASSED
                        End If
                        
                    End If
                    
                UpdateResults lstRow, "Exiting bypass mode...", eRESULT_TESTING
                SerialComs.SendATCommand etELITE_STATES.eElite_DISABLE_BYPASS, 1

                End If
                
                ' Set next IMEI number to write. This is the safest place
                ' to increment the IMEI number and insure no duplicates.
                ' We will have to decrement it locally when we go to check
                ' that the IMEI was written correctly but that's not a problem at all
                g_lNextIMEINumber.SNR = g_lNextIMEINumber.SNR + 1
                oNode.selectSingleNode("SNR").Attributes.getNamedItem("Next").Text = Format(g_lNextIMEINumber.SNR, IMEI_NUM_SNR_FORMAT)
                
                ' Set g_szIMEI to a string of all zeros so that we know to verify
                ' the IMEI write the next time the function is called
                g_szIMEI = String(Len(IMEI_NUM_FORMAT) + 1, "0")
                
                Exit For
            End If
        End If
        
        If eRESULT_TESTING = ConfigureBGS2ModemID Then
            ' We have succesfully read the IMEI
            If szPassTimeIMEI_TAC <> Left(szIMEINum, Len(IMEI_NUM_TAC_FORMAT)) Then
                ' TODO: Need direct way to determine if using S2210 variant but not
                ' sure if there is any way to check this directly. For now assume
                ' if the TAC isn't the PassTime TAC then we are using the S2210
                ' Can't read the SVN so assume it is zero. The SSVN command
                ' doesen't exist on the the S2210 variant of the BGS2 and
                ' occasionaly calling the unuupported command will lock up tester
                ' (reported from field but never verified)
                szSVN = IMEI_NUM_SV_FORMAT
            Else
                ' Read the SVN firmware version number
                szSVN = "?"
                SerialComs.BGS2ModemSVN szSVN
                If Len(szSVN) < 2 Then
                    szSVN = IMEI_NUM_SV_FORMAT
                Else
                    If IsNumeric(Right(szSVN, 2)) Then
                        szSVN = Trim(Right(szSVN, 2))
                        If 2 = Len(szSVN) Then
                            szSVN = Right(szSVN, 2)
                        ElseIf "0" = szSVN Then
                            szSVN = IMEI_NUM_SV_FORMAT
                        End If
                    Else
                        szSVN = IMEI_NUM_SV_FORMAT
                    End If
                End If
            End If
            Exit For
        End If
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
            ConfigureBGS2ModemID = eRESULT_FAILED
            Exit For
        Else
            startup.Delay 1#
        End If
    Next iRetries
    
    If eRESULT_TESTING <> ConfigureBGS2ModemID Then
        If eRESULT_FAILED <> ConfigureBGS2ModemID And String(Len(IMEI_NUM_FORMAT) + 1, "0") = g_szIMEI Then
            UpdateResults lstRow, "IMEI Number = " & szIMEINum, eRESULT_PASSED
        Else
            ConfigureBGS2ModemID = eRESULT_FAILED
        End If
    Else
        '   Confirm that the Modem ID is set properly
        UpdateResults lstRow, "Confirming BGS2 IMEI Number...", eRESULT_TESTING
               
        ' Verify length of return string. Need to add one to length
        ' because modem returns IMEI number plus check digit
        If Len(IMEI_NUM_FORMAT) + 1 <> Len(szIMEINum) Then
            ' Not an IMEI number
            g_fMainTester.LogMessage "IMEI number = " & szIMEINum
            If "NO IMEI" = szIMEINum Or ("^SGSN:" = Left(szIMEINum, 6) And ",NO IMEI" = Right(szIMEINum, 8)) Then
                If vbChecked = chkTrialRun.value And vbChecked = Me.chkConfigureUnit.value Then
                    UpdateResults lstRow, "Trial run - skipped", eRESULT_NOT_TESTED
                End If
            Else
                g_fMainTester.LogMessage "Malformed IMEI Number"
            End If
        Else
            If _
                Len(szIMEINum) < Len(IMEI_NUM_FORMAT) + 1 Or _
                False = IsNumeric(Right(szIMEINum, 1)) Then
'                1 <> InStr(szIMEINum, "IMEI") Or _

                g_fMainTester.LogMessage "Malformed IMEI Number"
            Else
                ' Parse out the IMEI number including the check digit
                szIMEINum = Right(szIMEINum, Len(IMEI_NUM_FORMAT) + 1)
                g_szIMEI = szIMEINum
'                If True = IsNumeric(Right(szIMEINum, 1)) Then
'                    iSNR_Suffix = Right(szIMEINum, 1)
'                End If
                szIMEINum = Left(szIMEINum, Len(IMEI_NUM_FORMAT))
        
                If True = bVerifyIMEI_Write Then
                    ' Get local IMEI number to verify that the number read
                    ' matches what we expected
                    lIMEINumber.TAC = g_lNextIMEINumber.TAC
                    
                    ' Need to subtract one because the XML has already burned the
                    ' serial number (i.e., incremented to next serial number)
                    ' even if our verify fails
                    lIMEINumber.SNR = g_lNextIMEINumber.SNR - 1
                    
                    lIMEINumber.SNR_Suffix = g_lNextIMEINumber.SNR_Suffix
                    
                    If lIMEINumber.TAC = CLng(Left(szIMEINum, Len(IMEI_NUM_TAC_FORMAT))) Then
                        If lIMEINumber.SNR = CLng(Right(szIMEINum, Len(IMEI_NUM_SNR_FORMAT))) Then
                            If lIMEINumber.SNR_Suffix = CLng(Right(szSVN, Len(IMEI_NUM_SV_FORMAT))) Then
                                ConfigureBGS2ModemID = eRESULT_PASSED
                            End If
                        End If
                    End If
                Else
                    ' Will land here typically when IMEI number is already set. In that case we
                    ' update the globabl "next" IMEI number variable with what we read and
                    ' then recheck that it matches after we power cycle/reset the DUT
                    g_lNextIMEINumber.TAC = CLng(Left(szIMEINum, Len(IMEI_NUM_TAC_FORMAT)))
                    g_lNextIMEINumber.SNR = CLng(Right(szIMEINum, Len(IMEI_NUM_SNR_FORMAT)))
                    g_lNextIMEINumber.SNR_Suffix = CLng(Right(szSVN, Len(IMEI_NUM_SV_FORMAT)))
                    ConfigureBGS2ModemID = eRESULT_PASSED
                    
                    ' 7/1/2013: Need to add a check here at least tempoarily to see if COPS=0.
                    startup.Delay 1#
                    szOpStat = ""
                    If True = SerialComs.GetOperatorStatus(szOpStat) Then
                        If Not 1 = InStr(1, szOpStat, "+COPS: 0") Then
                        ' If Not "+COPS: 0" = szOpStat Then
                            ' Not in automatic mode
                            Debug.Print "Not in automatic mode (" & szOpStat & ")"
                            g_fMainTester.LogMessage " Automatic registration not set"
                            If True = SendATCommand(etELITE_STATES.eElite_SET_STATION_AUTOMATIC, 20000) Then
                                Debug.Print "Sent COPS=0 to DUT"
                                Me.LogMessage " Set to automatic registration"
                            Else
                                Debug.Print "COPS=0 command failed"
                                Me.LogMessage " Failed to set automatic registration"
                            End If
                        Else
                            ' g_fMainTester.LogMessage (" DUT in automatic mode (" & m_szParameter & ").")
                        End If
                    Else
                        ' TODO: If we keep this check here need to do something
                    End If
                    
                End If
                
                If eRESULT_PASSED = ConfigureBGS2ModemID Then
                    UpdateResults lstRow, "IMEI # = " & g_szIMEI, eRESULT_PASSED
                    ' TODO: Even if not verify we should probably error if the TAC and SVN aren't correct
                End If
            End If
        End If

    End If
    
    If eRESULT_PASSED <> ConfigureBGS2ModemID Then
        If vbChecked = chkTrialRun.value Then
            ConfigureBGS2ModemID = eRESULT_NOT_TESTED
        Else
            ConfigureBGS2ModemID = eRESULT_FAILED
            If (vbChecked = Me.chkConfigureUnit.value And eIMEI_NUM_READ_ONLY <> eID_Type) Then
                If 0 <> InStr(lstRow.SubItems(eITEM_RESULTS), "Confirming") Then
                    UpdateResults lstRow, "Failed to verify modem IMEI number", eRESULT_FAILED
                Else
                    UpdateResults lstRow, "Failed to set modem IMEI number", eRESULT_FAILED
                End If
            Else
                UpdateResults lstRow, "IMEI number not set", eRESULT_FAILED
            End If
        End If
    End If

End Function

Private Function ConfigureModemID(lstRow As ListItem, Optional eID_Type As etID_TYPE = eIMEI_NUM) As etRESULTS
    Dim iRetries As Integer
    Dim szIMEINum As String
    Dim iSNR_Suffix As Integer
    Dim bVerifyIMEI_Write As Boolean
    Dim lIMEINumber As stIMEI_NUM
    Dim oNode As IXMLDOMNode
    
    szIMEINum = ""
    ' Check if g_szIMEI is a string of all zeros
    If String(Len(IMEI_NUM_FORMAT) + 1, "0") = g_szIMEI Then
        ' String of all zeros indicates that we are verifying
        ' an IMEI number that was just written
        bVerifyIMEI_Write = True
    Else
        bVerifyIMEI_Write = False
    End If

    ' Read IMEI number and see if it needs to be set
    UpdateResults lstRow, "Reading IMEI Number...", eRESULT_TESTING
    
    Set oNode = g_oSettings.selectSingleNode("TestSystem").childNodes(g_iTestStationID)
    Set oNode = oNode.selectSingleNode("IMEI")
    If oNode.selectSingleNode("Input_Method").Text = "BY_MODEL" Then
        Set oNode = g_oSelectedModel.selectSingleNode("IMEI")
    End If
    
    For iRetries = 3 To 1 Step -1
        ConfigureModemID = eRESULT_TESTING
        szIMEINum = ""
        
        If False = SerialComs.ModemID(szIMEINum) Then
            ConfigureModemID = eRESULT_ERROR
        ElseIf (vbUnchecked = chkTrialRun.value And vbChecked = Me.chkConfigureUnit.value And eIMEI_NUM_READ_ONLY <> eID_Type) Then
            ' Not in read only mode. See if we are in airplane mode
            ' 20130116 Apparently the Frankenstein (PTE-2.vT) boards can be
            ' prograammed but still report Airplane mode
            If "IMEI:0" = szIMEINum Or "IMEI:NO IMEI" = szIMEINum Then
                '   Configure the Modem ID once and only once
                UpdateResults lstRow, "Setting IMEI Number...", eRESULT_TESTING
                g_lNextIMEINumber.TAC = CLng(Val(oNode.selectSingleNode("TAC").Text))
                g_lNextIMEINumber.SNR = CLng(Val(oNode.selectSingleNode("SNR").Attributes.getNamedItem("Next").Text))
                g_lNextIMEINumber.SNR_Suffix = CLng(Val(oNode.selectSingleNode("SV").Text))
                
                szIMEINum = Format(g_lNextIMEINumber.TAC, IMEI_NUM_TAC_FORMAT)
                szIMEINum = szIMEINum & Format(g_lNextIMEINumber.SNR, IMEI_NUM_SNR_FORMAT)
                szIMEINum = Format(szIMEINum, IMEI_NUM_FORMAT)
                ' szIMEINum = szIMEINum & Format(g_lNextIMEINumber.SNR_Suffix, IMEI_NUM_SV_FORMAT)
                
                ' Write the IMEI number to the modem
                If False = SerialComs.ModemID(szIMEINum) Then
                    ConfigureModemID = eRESULT_ERROR
                Else
                    ConfigureModemID = eRESULT_PASSED
                End If
                
                ' Set next IMEI number to write. This is the safest place
                ' to increment the IMEI number and insure no duplicates.
                ' We will have to decrement it locally when we go to check
                ' that the IMEI was written correctly but that's not a problem at all
                g_lNextIMEINumber.SNR = g_lNextIMEINumber.SNR + 1
                oNode.selectSingleNode("SNR").Attributes.getNamedItem("Next").Text = Format(g_lNextIMEINumber.SNR, IMEI_NUM_SNR_FORMAT)
                
                ' Firmware will reset modem so give it a few seconds before
                ' attempting to restart
                ' TODO: Is this still necessary
                g_fMainTester.LogMessage "Waiting 10 seconds for soft modem restart"
                
                ' Set g_szIMEI to a string of all zeros so that we know to verify
                ' the IMEI write the next time the function is called
                g_szIMEI = String(Len(IMEI_NUM_FORMAT) + 1, "0")
                
                Dim iWaitForSoftReset As Integer
                For iWaitForSoftReset = IMEI_WRITE_WAIT To 1 Step -1
                    UpdateResults lstRow, "Waiting for soft modem restart..." & iWaitForSoftReset, eRESULT_TESTING
                    startup.Delay 1#
                    If m_eTestingState = eTESTER_OPEN Then
                        UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                        ConfigureModemID = eRESULT_FAILED
                        Exit For
                    End If
                Next iWaitForSoftReset
                Exit For
            End If
        End If
        If (eRESULT_TESTING = ConfigureModemID) Then
            ' We have succesfully read the IMEI
            Exit For
        End If
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
            ConfigureModemID = eRESULT_FAILED
            Exit For
        Else
            startup.Delay 1#
        End If
    Next iRetries
    
    If eRESULT_TESTING <> ConfigureModemID Then
        If eRESULT_FAILED <> ConfigureModemID And String(Len(IMEI_NUM_FORMAT) + 1, "0") = g_szIMEI Then
            UpdateResults lstRow, "IMEI Number = " & szIMEINum, eRESULT_PASSED
        Else
            ConfigureModemID = eRESULT_FAILED
        End If
    Else
        '   Confirm that the Modem ID is set properly
        UpdateResults lstRow, "Confirming IMEI Number...", eRESULT_TESTING
               
        ' Verify length of return string. Need to add one to length
        ' because modem returns IMEI number plus check digit
        If Len("IMEI:") + Len(IMEI_NUM_FORMAT) + 1 <> Len(szIMEINum) Then
            If "IMEI:0" = szIMEINum Or "IMEI:NO IMEI" = szIMEINum Then
                g_fMainTester.LogMessage "NO IMEI or modem still set to Airplane Mode"
                If vbChecked = chkTrialRun.value And vbChecked = Me.chkConfigureUnit.value Then
                    UpdateResults lstRow, "Trial run - skipped", eRESULT_NOT_TESTED
                End If
            Else
                g_fMainTester.LogMessage "Malformed IMEI Number"
            End If
        Else
            If _
                Len(szIMEINum) < Len(IMEI_NUM_FORMAT) + 1 Or _
                1 <> InStr(szIMEINum, "IMEI") Or _
                False = IsNumeric(Right(szIMEINum, 1)) Then
                
                g_fMainTester.LogMessage "Malformed IMEI Number"
            Else
                ' Parse out the IMEI number including the check digit
                szIMEINum = Right(szIMEINum, Len(IMEI_NUM_FORMAT) + 1)
                g_szIMEI = szIMEINum
                If False = IsNumeric(Right(szIMEINum, 1)) Then
                    iSNR_Suffix = Right(szIMEINum, 1)
                End If
                szIMEINum = Left(szIMEINum, Len(IMEI_NUM_FORMAT))
        
                If True = bVerifyIMEI_Write Then
                    ' Get local IMEI number to verify that the number read
                    ' matches what we expected
                    lIMEINumber.TAC = g_lNextIMEINumber.TAC
                    
                    ' Need to subtract one because the XML has already burned the
                    ' serial number (i.e., incremented to next serial number)
                    ' even if our verify fails
                    lIMEINumber.SNR = g_lNextIMEINumber.SNR - 1
                    
                    If lIMEINumber.TAC = CLng(Left(szIMEINum, Len(IMEI_NUM_TAC_FORMAT))) Then
                        If lIMEINumber.SNR = CLng(Right(szIMEINum, Len(IMEI_NUM_SNR_FORMAT))) Then
                            ConfigureModemID = eRESULT_PASSED
                        End If
                    End If
                Else
                    g_lNextIMEINumber.TAC = CLng(Left(szIMEINum, Len(IMEI_NUM_TAC_FORMAT)))
                    g_lNextIMEINumber.SNR = CLng(Right(szIMEINum, Len(IMEI_NUM_SNR_FORMAT)))
                    g_lNextIMEINumber.SNR_Suffix = iSNR_Suffix
                    ConfigureModemID = eRESULT_PASSED
                End If
                
                If eRESULT_PASSED = ConfigureModemID Then
                    UpdateResults lstRow, "IMEI Number = " & g_szIMEI, eRESULT_PASSED
                End If
            End If
        End If

    End If
    
    If eRESULT_PASSED <> ConfigureModemID Then
        If vbChecked = chkTrialRun.value Then
            ConfigureModemID = eRESULT_NOT_TESTED
        Else
            m_bConfigureSuccessful = False
            ConfigureModemID = eRESULT_FAILED
            UpdateResults lstRow, "Failed to set modem IMEI number", eRESULT_FAILED
        End If
    End If

End Function
Private Function ConfigureGSM_Communications(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    ' This is a work around for an issue in the GSM modem firmware that causes
    ' an unnecessary delay in th MCU firmware when trying to set the flow control
    If Not bRunTest Then
        If chkTrialRun.value <> vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Configure GSM Modem Serial Comms"
        Else
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify GSM Modem Serial Comms"
        End If
        ConfigureGSM_Communications = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureGSM_Communications = eRESULT_NOT_TESTED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            ConfigureGSM_Communications = eRESULT_NOT_TESTED
        Else
            
            If vbChecked = chkTrialRun.value Then
                '   Trial Run -- Just verify the AT and get the CREG status
                ConfigureGSM_Communications = eRESULT_NOT_TESTED
                UpdateResults lstRow, "Verifying GSM Modem Serial Comms...", eRESULT_TESTING
            Else
                ConfigureGSM_Communications = eRESULT_TESTING
                UpdateResults lstRow, "Querying GSM modem...", eRESULT_TESTING
            End If
                
            If False = SerialComs.VerifyOpenAT Then
                ' We have waited at least 20 seconds already and still no response
                ' to the AT command but we did get a response from the non-AT command
                ' to configure from testing
                m_bDeviceFailure = True
                ConfigureGSM_Communications = eRESULT_FAILED
                UpdateResults lstRow, "GSM modem not responding", eRESULT_FAILED
            End If
            
            ' set to bypass
            If eRESULT_FAILED <> ConfigureGSM_Communications Then
                UpdateResults lstRow, "Entering bypass mode...", eRESULT_TESTING
                ' TODO: Calling SendATCommand(eElite_ENABLE_BYPASS) always returns true
                ' so this conditional is never executed
                If False = SerialComs.SendATCommand(etELITE_STATES.eElite_ENABLE_BYPASS, 1) Then
                    UpdateResults lstRow, "Failed to enter bypass mode", eRESULT_FAILED
                    ConfigureGSM_Communications = eRESULT_FAILED
                End If
            End If
            
            ' TODO: Setting bypass mode may be screwing up parser on next CLI command
            ' 7/1/2013 Added a retry to see if the CLI failure clears up the parser
            If eRESULT_TESTING = ConfigureGSM_Communications Then
                ' Only executed if not in trial run mode
                UpdateResults lstRow, "Setting GSM modem flow control...", eRESULT_TESTING
                If False = SerialComs.SendATCommand(eElite_SET_GSM_MODEM_FLOW_CTRL) Then
                    Me.LogMessage " AT command failure. Retrying..."
                    ' Give it one more shot
                    startup.Delay 1#
                    If False = SerialComs.SendATCommand(eElite_SET_GSM_MODEM_FLOW_CTRL) Then
                        UpdateResults lstRow, "Failed to set flow control on GSM modem", eRESULT_FAILED
                        ConfigureGSM_Communications = eRESULT_FAILED
                    End If
                End If
            
                If eRESULT_TESTING = ConfigureGSM_Communications And False = SerialComs.SendATCommand(eElite_SET_GSM_MODEM_BAUD_RATE) Then
                    UpdateResults lstRow, "Failed to set baud rate on GSM modem", eRESULT_FAILED
                    ConfigureGSM_Communications = eRESULT_FAILED
                End If
    
            End If
                
            If eRESULT_FAILED <> ConfigureGSM_Communications Then
                ' Observed failure if delay is < 0.02
                startup.Delay 0.05
                
                UpdateResults lstRow, "Query GSM modem for registration status...", eRESULT_TESTING
                If False = SerialComs.SendATCommand(eElite_GET_GSM_REG_STATUS) Then
                    UpdateResults lstRow, "Failed to get gsm registration status", eRESULT_FAILED
                    ConfigureGSM_Communications = eRESULT_FAILED
                End If
            End If

            ' onced configured or verified, take out of bypass mode
            If eRESULT_FAILED <> ConfigureGSM_Communications Then
                UpdateResults lstRow, "Exiting bypass mode...", eRESULT_TESTING
                
                ' TODO: Calling SendATCommand(eElite_DISABLE_BYPASS) always returns true
                ' so this conditional is never executed
                If False = SerialComs.SendATCommand(etELITE_STATES.eElite_DISABLE_BYPASS, 1) Then
                    UpdateResults lstRow, "Failed to exit bypass mode", eRESULT_FAILED
                    ConfigureGSM_Communications = eRESULT_FAILED
                ElseIf eRESULT_NOT_TESTED = ConfigureGSM_Communications Then
                    ' Not testing but verified AT and CREG status
                    UpdateResults lstRow, "Verified GSM Modem Serial Comms", eRESULT_PASSED
                End If
            End If
            
            If eRESULT_TESTING = ConfigureGSM_Communications Then
                UpdateResults lstRow, "Configured GSM Modem Serial Comms", eRESULT_PASSED
                ConfigureGSM_Communications = eRESULT_PASSED
            ElseIf eRESULT_NOT_TESTED <> ConfigureGSM_Communications Then
                m_bConfigureSuccessful = False
            End If
            
        End If
    End If

End Function
Private Function ConfigureIDs(lstRow As ListItem, bRunTest As Boolean, Optional eID_Type As etID_TYPE = eSERIAL_NUM) As etRESULTS

    Dim iRetries As Integer
    Dim szModemID As String
    Dim szID As String
    Dim oNode As IXMLDOMNode
    
    If etID_TYPE.eSERIAL_NUM = eID_Type Or etID_TYPE.eSERIAL_NUM_READ_ONLY = eID_Type Then
        szID = "Serial"
    Else
        szID = "Modem IMEI"
        Set oNode = g_oSettings.selectSingleNode("TestSystem").childNodes(g_iTestStationID)
        Set oNode = oNode.selectSingleNode("IMEI")
        If oNode.selectSingleNode("Input_Method").Text = "BY_MODEL" Then
            Set oNode = g_oSelectedModel.selectSingleNode("IMEI")
        End If
    End If

    If Not bRunTest Then
        If (vbChecked = chkTrialRun.value Or vbUnchecked = Me.chkConfigureUnit.value Or eIMEI_NUM_READ_ONLY = eID_Type Or eSERIAL_NUM_READ_ONLY = eID_Type) Then
            ' example: Get Modem IMEI Number
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Get " & szID & " Number"
        Else
            ' example: Set Modem IMEI Number
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set " & szID & " Number"
        End If
        ConfigureIDs = eRESULT_NOT_TESTED
    Else
        ConfigureIDs = eRESULT_TESTING
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureIDs = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            ConfigureIDs = eRESULT_NOT_TESTED
        Else
            If etID_TYPE.eSERIAL_NUM = eID_Type Or etID_TYPE.eSERIAL_NUM_READ_ONLY = eID_Type Then
                ConfigureIDs = ConfigureSerialNum(lstRow, etID_TYPE.eSERIAL_NUM_READ_ONLY <> eID_Type)
            Else
                ' 20121221: Routines for setting/reading IMEI on BGS3 modem (used on older PassTime devices)
                ' are not the same on the BGS2 modems. Easiest path forward for now that maintains
                ' backward compatibility is to create new ConfigureModemID function instead of trying
                ' to shoehorn in the differences into one function
                
                ' Which modem are we using
                If "BGS2" = oNode.selectSingleNode("ModemModel").Text Then
                    ConfigureIDs = ConfigureBGS2ModemID(lstRow, eID_Type)
                    If ConfigureIDs <> eRESULT_PASSED Then
                        Debug.Print "Failed. Retrying modem configuration..."
                        LogMessage "Failed. Retrying modem configuration..."
                        ConfigureIDs = ConfigureBGS2ModemID(lstRow, eID_Type)
                    End If
                Else
                    ConfigureIDs = ConfigureModemID(lstRow, eID_Type)
                End If
            End If
            If ConfigureIDs <> eRESULT_PASSED Then
                If (vbChecked = chkTrialRun.value Or vbUnchecked = Me.chkConfigureUnit.value) Then
                    UpdateResults lstRow, "Failed to read " & szID & " number", eRESULT_FAILED
                Else
                    If 0 <> InStr(lstRow.SubItems(eITEM_RESULTS), "verify") Then
                        UpdateResults lstRow, "Failed to verify " & szID & " number", eRESULT_FAILED
                    Else
                        UpdateResults lstRow, "Failed to set " & szID & " number", eRESULT_FAILED
                    End If
                    m_bConfigureSuccessful = False
                    m_bDeviceFailure = True
                End If
            End If
        End If
    End If
    
End Function

Private Function ConfigureAPN(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    
    Dim oAttributes As IXMLDOMNamedNodeMap

    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set and Verify APN"
        Else
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify APN"
        End If
        ConfigureAPN = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
            ConfigureAPN = eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
            ConfigureAPN = eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureAPN = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            ConfigureAPN = eRESULT_NOT_TESTED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            ConfigureAPN = eRESULT_NOT_TESTED
        Else
            '   Configure the APN name and password
            UpdateResults lstRow, "Configuring APN...", eRESULT_TESTING
            If vbChecked = chkTrialRun.value Then
                UpdateResults lstRow, "Verifying APNs...", eRESULT_TESTING
            Else
                UpdateResults lstRow, "Configuring APNs...", eRESULT_TESTING
            End If
            ConfigureAPN = eRESULT_TESTING
            
            Dim iRetries As Integer
            Dim szParam As String
            Dim oDOM_Node As IXMLDOMNode
            Dim szSingleNodeTag As String
            Dim szAPN As String
            Dim szPassword As String
            Dim iApnIdx As Integer
            Dim iIdx As Integer

            iIdx = 0
            For Each oDOM_Node In g_oSelectedSIM.selectSingleNode("APN").childNodes
                If "PW" = oDOM_Node.baseName Then
                    szPassword = oDOM_Node.Attributes.getNamedItem("Pass").Text
                    Exit For
                End If
            Next oDOM_Node
            
            For Each oDOM_Node In g_oSelectedSIM.selectSingleNode("APN").childNodes
                If "APN" = oDOM_Node.baseName Then
                    szAPN = oDOM_Node.Attributes.getNamedItem("URL").Text
                    iIdx = CInt(oDOM_Node.Attributes.getNamedItem("Index").Text)
                
                    If Not vbChecked = chkTrialRun Then
                        ' Set the APN
                        For iRetries = 1 To 3
                            If SerialComs.SetAPN(iIdx, szAPN, szPassword) Then
                                ' AT command success
                                iRetries = 0
                                Exit For
                            End If
                            startup.Delay 0.25
                        Next iRetries
                        
                        If iRetries > 0 Then
                            ' AT command failed
                            UpdateResults lstRow, "Failed to set APN " & iIdx, eRESULT_FAILED
                            ConfigureAPN = eRESULT_FAILED
                            iIdx = -1
                            Exit For
    '                    Else
    '                        UpdateResults lstRow, "APN" & iIdx & " = " & szAPN, eRESULT_TESTING
                        End If
                    End If
                    
                    If eRESULT_TESTING = ConfigureAPN Then
                        ' Verify the APN got set
                        For iRetries = 1 To 3
                            If SerialComs.GetAPN(iIdx, szParam) Then
                                ' AT command success
                                iRetries = 0
                                Exit For
                            End If
                            startup.Delay 0.25
                        Next iRetries
                        
                        If iRetries > 0 Then
                            ' AT command failed
                            If Not vbChecked = chkTrialRun.value Then
                                ' Since AT command to set the APN executed successfully, allow testing to continue
                                UpdateResults lstRow, "Failed to verify APN " & iIdx, eRESULT_NOT_TESTED
                                ConfigureAPN = eRESULT_NOT_TESTED
                            Else
                                ' Trial Run selected. Fail if we can't verify
                                UpdateResults lstRow, "Failed to verify APN " & iIdx, eRESULT_FAILED
                                iIdx = -1
                                Exit For
                            End If
                            
                            If szAPN = szParam Then
                                UpdateResults lstRow, "APN" & iIdx & " = " & szAPN, eRESULT_TESTING
                            Else
                                iIdx = -1
                                Exit For
                            End If
                        ElseIf szParam <> szAPN Then
                            UpdateResults lstRow, "Expected " & szAPN & ", received " & szParam & ". Failed to set APN " & iIdx, eRESULT_FAILED
                            ConfigureAPN = eRESULT_FAILED
                            iIdx = -1
                            Exit For
                        End If
                    End If
                End If
            Next oDOM_Node
            
            If iIdx = -1 Then
                ConfigureAPN = eRESULT_FAILED
                m_bDeviceFailure = True
            Else
                If vbChecked = chkTrialRun Then
                    UpdateResults lstRow, "APNs verified", eRESULT_PASSED
                Else
                    UpdateResults lstRow, "APNs set", eRESULT_PASSED
                End If
                ConfigureAPN = eRESULT_PASSED
            End If

            If eRESULT_TESTING = ConfigureAPN Then
                UpdateResults lstRow, "APNs set", eRESULT_PASSED
                ConfigureAPN = eRESULT_PASSED
            End If
        End If
    End If
End Function

Private Function ConfigureServers(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set and Verify Server Address"
        Else
             ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify Server Address"
        End If
        
        ConfigureServers = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
            ConfigureServers = eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
            ConfigureServers = eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureServers = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            ConfigureServers = eRESULT_NOT_TESTED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            ConfigureServers = eRESULT_NOT_TESTED
        Else
            '   Configure the Server IP addresses
            If vbChecked = chkTrialRun.value Then
                UpdateResults lstRow, "Verifying IPServers...", eRESULT_TESTING
            Else
                UpdateResults lstRow, "Configuring IPServers...", eRESULT_TESTING
            End If
            ConfigureServers = eRESULT_TESTING
            
            Dim iRetries As Integer
            Dim szParam As String
            Dim oDOM_Node As IXMLDOMNode
            Dim szSingleNodeTag As String
            Dim szIP As String
            Dim iServerIpIdx As Integer
            Dim iIdx As Integer
            
            For Each oDOM_Node In g_oSelectedSIM.selectSingleNode("Server").childNodes
                ' Retrieve the IP address from the XML file
                szIP = oDOM_Node.Attributes.getNamedItem("IP").Text
                iIdx = CInt(oDOM_Node.Attributes.getNamedItem("Index").Text)
               
                If Not vbChecked = chkTrialRun Then
                    ' Set the server IP
                    For iRetries = 1 To 3
                        If SerialComs.SetServerIP(iIdx, szIP) Then
                            iRetries = 0
                            Exit For
                        End If
                        startup.Delay 0.25
                    Next iRetries
                    
                    If iRetries > 0 Then
                        ' AT command failed
                        UpdateResults lstRow, "Failed to set server " & iIdx & "'s IP", eRESULT_FAILED
                        ConfigureServers = eRESULT_FAILED
                        Exit For
                    End If
                End If
                
                If eRESULT_TESTING = ConfigureServers Then
                    ' Verify the server IP got set
                    For iRetries = 1 To 3
                        If SerialComs.GetServerIP(iIdx, szParam) Then
                            ' AT command success
                            iRetries = 0
                            Exit For
                        End If
                        startup.Delay 0.25
                    Next iRetries
                    
                    If iRetries > 0 Then
                        ' AT command failed
                        If Not vbChecked = chkTrialRun.value Then
                            ' Doesn't mean IP is not set just means AT command failed. Since AT command
                            ' to set the IP executed successfully, allow testing to continue
                            UpdateResults lstRow, "Failed to verify server IP " & iIdx, eRESULT_NOT_TESTED
                            ConfigureServers = eRESULT_NOT_TESTED
                        Else
                            ' Trial Run selected. Fail if we can't verify
                            UpdateResults lstRow, "Failed to verify server IP " & iIdx, eRESULT_FAILED
                            ConfigureServers = eRESULT_FAILED
                            Exit For
                        End If
                    ElseIf szParam <> szIP Then
                        UpdateResults lstRow, "Expected " & szIP & ", received " & szParam & ". Failed to set server " & iIdx & "'s IP", eRESULT_FAILED
                        ConfigureServers = eRESULT_FAILED
                        Exit For
                    End If
                End If
            Next oDOM_Node
            
            If eRESULT_TESTING = ConfigureServers Then
                If vbChecked = chkTrialRun Then
                    UpdateResults lstRow, "Verified IP addresses", eRESULT_PASSED
                Else
                    UpdateResults lstRow, "Server IP addresses set", eRESULT_PASSED
                End If
                ConfigureServers = eRESULT_PASSED
            End If
        End If
    End If

End Function

Private Function ConfigureFlash(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    ' Save Changes to SPI Flash
    Dim eResults As etRESULTS
    
    m_bApplicationRunning = False
    
    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            lstRow.ListSubItems(eITEM_DESCRIPTION).Text = "Save Values to SPI Flash"
        Else
             ' lstRow.SubItems(eITEM_DESCRIPTION) = "No Updates to SPI Flash"
        End If
        eResults = eRESULT_NOT_TESTED
        
    Else
        eResults = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            eResults = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            eResults = eRESULT_NOT_TESTED
        Else
            eResults = eRESULT_TESTING
            If vbChecked = chkConfigureUnit.value And vbChecked <> chkTrialRun.value Then
                If False = SerialComs.SaveFlashParams() Then
                    UpdateResults lstRow, "Failed to save values to SPI flash", eRESULT_FAILED
                    eResults = eRESULT_FAILED
                Else
                    Dim iWaitForSPIWrite As Integer
                    For iWaitForSPIWrite = SPI_WRITE_WAIT To 1 Step -1
                        UpdateResults lstRow, "Saving values to SPI flash..." & iWaitForSPIWrite, eRESULT_TESTING
                        startup.Delay 1#
                        If m_eTestingState = eTESTER_OPEN Then
                            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                            eResults = eRESULT_FAILED
                            Exit For
                        End If
                    Next iWaitForSPIWrite
                End If
            ElseIf vbChecked = chkTrialRun.value Then
                UpdateResults lstRow, "Skipped", eRESULT_NOT_TESTED
                eResults = eRESULT_NOT_TESTED
            End If
        End If
        If eRESULT_TESTING = eResults Then
            UpdateResults lstRow, "Saved values to SPI flash", eRESULT_PASSED
            eResults = eRESULT_PASSED
        ElseIf eRESULT_NOT_TESTED <> eResults Then
            m_bDeviceFailure = True
        End If
        
    End If
    ConfigureFlash = eResults

End Function
Private Function ConfigureSrvcCntrAddr(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set Service Center Address"
        Else
             ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify Service Center Address"
        End If
        
        ConfigureSrvcCntrAddr = eRESULT_NOT_TESTED
    Else
        ConfigureSrvcCntrAddr = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureSrvcCntrAddr = eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            ConfigureSrvcCntrAddr = eRESULT_NOT_TESTED
        Else
            Dim chkStr As String
            Dim scaStr As String
            Dim scaAddr As String
            Dim scaAddrType As String
            Dim smsHelloPktTimer As String
            Dim iStrIdx As Integer
            Dim oNamedNodeMap As IXMLDOMNamedNodeMap
            Dim oSmsAttribute As IXMLDOMAttribute
            Dim oTag As IXMLDOMNode
    
            Set oNamedNodeMap = g_oSelectedSIM.selectSingleNode("SrvcCntrAddr").selectSingleNode("SCA").Attributes
            
            ' ADDR is verified by XML parser that it has at least 12 chars,
            ' the first char is "+", and all remaining chars are numbers
            Set oSmsAttribute = oNamedNodeMap.getNamedItem("ADDR")
            scaAddr = oSmsAttribute.Text
            
            ' Set oSmsAttribute = oNamedNodeMap.getNamedItem("ADDR_TYPE")
            ' scaAddrType = oSmsAttribute.Text
            
            ' Since the XML parser verifies we have an international address
            ' the ADDR TYPE will be automatically set to 145
            scaAddrType = "145"
            
            chkStr = Chr(&H22) & scaAddr & Chr(&H22) & "," & scaAddrType
            ' chkStr = scaAddr & "," & scaAddrType
            UpdateResults lstRow, "Reading current Service Center Address...", eRESULT_TESTING
            scaStr = GetSCAString()
            ' UpdateResults lstRow, "SCA = " & scaStr, eRESULT_TESTING
            If "" = scaStr Then
                ' Could be waiting for SIM
                startup.Delay 5#
                scaStr = GetSCAString()
            End If
            If "" = scaStr Then
                ' Command failed
                UpdateResults lstRow, "Couldn't verify Service Center Address", eRESULT_FAILED
                ConfigureSrvcCntrAddr = eRESULT_FAILED
            ElseIf Chr(&H22) & "+0000000000000" & Chr(&H22) & "," & scaAddrType = chkStr Then
            ' ElseIf "+0000000000000," & scaAddrType = chkStr Then
                If InStr(scaStr, "+CSCA: ") > 0 Then
                    chkStr = Right(scaStr, Len(scaStr) - Len("+CSCA: "))
                    ' lstRow.SubItems(eITEM_DESCRIPTION) = "Get Service Center Address"
                    UpdateResults lstRow, "SCA = " & chkStr, eRESULT_PASSED
                    ConfigureSrvcCntrAddr = eRESULT_PASSED
                Else
                    UpdateResults lstRow, "Couldn't verify Service Center Address", eRESULT_FAILED
                    ConfigureSrvcCntrAddr = eRESULT_FAILED
                End If
                scaStr = ""
            ElseIf InStr(scaStr, chkStr) > 0 Then
                ' Service Center Address is already set
                ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify Service Center Address"
                UpdateResults lstRow, "SCA = " & chkStr, eRESULT_PASSED
                ConfigureSrvcCntrAddr = eRESULT_PASSED
                scaStr = ""
            ElseIf vbChecked = chkTrialRun Then
                ' Service Center Address not set correctly and trial run selected
                ' so can't write to device, therefore, test fails
                UpdateResults lstRow, "Got " & scaStr & " but expected " & chkStr, eRESULT_FAILED
                ConfigureSrvcCntrAddr = eRESULT_FAILED
            Else
'                UpdateResults lstRow, "Entering bypass mode...", eRESULT_TESTING
'                ' TODO: Calling SendATCommand(eElite_ENABLE_BYPASS) always returns true
'                ' so this conditional is never executed
'                If False = SerialComs.SendATCommand(etELITE_STATES.eElite_ENABLE_BYPASS, 1) Then
'                    UpdateResults lstRow, "Failed to enter bypass mode", eRESULT_FAILED
'                    ConfigureSrvcCntrAddr = eRESULT_FAILED
'                End If
                UpdateResults lstRow, "Setting Service Center Address...", eRESULT_TESTING
'                scaStr =  & scaAddr
                

                If SendSCAString(Chr(&H22) & scaAddr & Chr(&H22)) Then
                    UpdateResults lstRow, "Set Service Center Address", eRESULT_TESTING
                    scaStr = GetSCAString()
                    If InStr(scaStr, chkStr) > 0 Then
                        UpdateResults lstRow, "SCA = " & chkStr, eRESULT_PASSED
                        ConfigureSrvcCntrAddr = eRESULT_PASSED
                    End If
                End If
                
'                ' TODO: Calling SendATCommand(eElite_DISABLE_BYPASS) always returns true
'                ' so this conditional is never executed
'                UpdateResults lstRow, "Exiting bypass mode...", eRESULT_TESTING
'                If False = SerialComs.SendATCommand(etELITE_STATES.eElite_DISABLE_BYPASS, 1) Then
'                    UpdateResults lstRow, "Failed to exit bypass mode", eRESULT_FAILED
'                    ConfigureSrvcCntrAddr = eRESULT_FAILED
'                End If
                
                If eRESULT_PASSED <> ConfigureSrvcCntrAddr Then
                    UpdateResults lstRow, "Couldn't set Service Center Address", eRESULT_FAILED
                    ConfigureSrvcCntrAddr = eRESULT_FAILED
                End If
            End If
        End If
    End If

End Function

Private Function ConfigurePort(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim szPort As String
    Dim szParam As String
    Dim iRetries As Integer

    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set Port Address"
        Else
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Read Port Address"
        End If
        ConfigurePort = eRESULT_NOT_TESTED
    Else
        ConfigurePort = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigurePort = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            ConfigurePort = eRESULT_NOT_TESTED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            ConfigurePort = eRESULT_NOT_TESTED
        ElseIf Not chkTrialRun.value = vbChecked Then
            '   Set the port address then  verify its setting
            UpdateResults lstRow, "Setting Port address...", eRESULT_TESTING
            szPort = g_oSelectedSIM.selectSingleNode("Port").Text
            For iRetries = 1 To 3
'                startup.Delay 5#
                If SerialComs.SetPort(szPort) Then
                    iRetries = 0
                    Exit For
                End If
            Next iRetries
            If iRetries > 0 Then
                UpdateResults lstRow, "Failed to set Port address", eRESULT_FAILED
            Else
                For iRetries = 1 To 3
                    If SerialComs.GetPort(szParam) Then
                        iRetries = 0
                        Exit For
                    End If
                Next iRetries
                If iRetries > 0 Then
                    UpdateResults lstRow, "Failed to get Port number", eRESULT_FAILED
                ElseIf Mid(szParam, 1, 4) <> szPort Then
                    UpdateResults lstRow, "Port did not set properly, expecting " & Str(szPort) & ", received " & Str(szParam), eRESULT_FAILED
                Else
                    UpdateResults lstRow, "Port address = " & Str(szPort), eRESULT_PASSED
                    ConfigurePort = eRESULT_PASSED
                End If
            End If
            If ConfigurePort <> eRESULT_PASSED Then
                m_bConfigureSuccessful = False
                m_bDeviceFailure = False
                ConfigurePort = eRESULT_FAILED
            End If
        Else ' Trial Run -- just read and report Port Address
            UpdateResults lstRow, "Reading Port address...", eRESULT_TESTING
            szPort = g_oSelectedSIM.selectSingleNode("Port").Text
            For iRetries = 1 To 3
                If SerialComs.GetPort(szParam) Then
                    iRetries = 0
                    Exit For
                End If
            Next iRetries
            If iRetries > 0 Then
                UpdateResults lstRow, "Failed to read Port number", eRESULT_FAILED
            ElseIf Mid(szParam, 1, 4) <> szPort Then
                UpdateResults lstRow, "Port did not get set properly, expecting " & Str(szPort) & ", received " & Str(szParam), eRESULT_FAILED
            Else
                UpdateResults lstRow, "Port address = " & Str(szPort), eRESULT_PASSED
                ConfigurePort = eRESULT_PASSED
            End If
        End If
        If ConfigurePort <> eRESULT_PASSED Then
            m_bDeviceFailure = True
            ConfigurePort = eRESULT_FAILED
        End If
    End If

End Function


Private Function TestIgnition(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim bState As String

    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test Ignition Relay"
        TestIgnition = eRESULT_NOT_TESTED
    Else
        TestIgnition = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestIgnition = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            TestIgnition = eRESULT_NOT_TESTED
        Else
            '   Test the Ignition relay via the PIC I/F
            UpdateResults lstRow, "Testing Ignition relay...", eRESULT_TESTING
            'first check that nothing is on ignition line
            If SerialComs.CheckIgnition(bState) Then
                If bState = "0" Then
                    g_fMainTester.LogMessage "Ignition line initially off"
                Else
                    UpdateResults lstRow, "Ignition line is not initially 0", eRESULT_FAILED
                    Exit Function
                End If
            Else
                UpdateResults lstRow, "Ignition command failed", eRESULT_FAILED
                Exit Function
            End If
            'second check that something is on ignition line
            g_fIO.IgnitionInput = True
            startup.Delay 1#
            If SerialComs.CheckIgnition(bState) Then
                If bState = "1" Then
                    g_fMainTester.LogMessage "Ignition line on"
                Else
                    UpdateResults lstRow, "Ignition line is not on", eRESULT_FAILED
                    Exit Function
                End If
            Else
                UpdateResults lstRow, "Ignition command failed", eRESULT_FAILED
                Exit Function
            End If
            'third check that nothing is on ignition line
            g_fIO.IgnitionInput = False
            Delay 1#
            If SerialComs.CheckIgnition(bState) Then
                If bState = "0" Then
                    g_fMainTester.LogMessage "Ignition line off"
                    UpdateResults lstRow, "Ignition relay passed", eRESULT_PASSED
                    TestIgnition = eRESULT_PASSED
                Else
                    UpdateResults lstRow, "Ignition line is not off", eRESULT_FAILED
                    Exit Function
                End If
            Else
                UpdateResults lstRow, "Ignition command failed", eRESULT_FAILED
                Exit Function
            End If
        End If
    End If

End Function

Private Function TestStarter(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim bState As String

    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test Starter Relay"
        TestStarter = eRESULT_NOT_TESTED
    Else
        TestStarter = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestStarter = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            TestStarter = eRESULT_NOT_TESTED
        Else
            '   Test the Starter relay via the PIC I/F
            UpdateResults lstRow, "Testing Starter relay...", eRESULT_TESTING
            'first check that nothing is on starter line
            If SerialComs.CheckStarter(bState) Then
                If bState = "0" Then
                    g_fMainTester.LogMessage "Starter line initially off"
                Else
                    UpdateResults lstRow, "Starter line is not initially 0", eRESULT_FAILED
                    Exit Function
                End If
            Else
                UpdateResults lstRow, "Starter command failed", eRESULT_FAILED
                Exit Function
            End If
            'second check that something is on ignition line
            g_fIO.StarterInput = True
            startup.Delay 1#
            If SerialComs.CheckStarter(bState) Then
                If bState = "1" Then
                    g_fMainTester.LogMessage "Starter line on"
                Else
                    UpdateResults lstRow, "Starter line is not on", eRESULT_FAILED
                    Exit Function
                End If
            Else
                UpdateResults lstRow, "Starter command failed", eRESULT_FAILED
                Exit Function
            End If
            'third check that nothing is on ignition line
            g_fIO.StarterInput = False
            startup.Delay 1#
            If SerialComs.CheckStarter(bState) Then
                If bState = "0" Then
                    g_fMainTester.LogMessage "Starter line off"
                    UpdateResults lstRow, "Starter relay passed", eRESULT_PASSED
                    TestStarter = eRESULT_PASSED
                Else
                    UpdateResults lstRow, "Starter line is not off", eRESULT_FAILED
                    Exit Function
                End If
            Else
                UpdateResults lstRow, "Starter command failed", eRESULT_FAILED
                Exit Function
            End If
        End If
    End If

End Function

Private Function TestBuzzer(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    Dim iNumRetries As Integer
    iNumRetries = 3
    
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test Buzzer"
        TestBuzzer = eRESULT_NOT_TESTED
    Else
        TestBuzzer = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestBuzzer = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            TestBuzzer = eRESULT_NOT_TESTED
        Else
            TestBuzzer = eRESULT_TESTING
            '   Test the buzzer
            UpdateResults lstRow, "Testing buzzer...", eRESULT_TESTING
            
            Do
                ' Reset last recorded sound level and frequency to zero
                g_dSoundLevel = 0
                g_dSoundFreq = 0
                
                ' Stimulate the device to producde the sound we want to check
                SendATCommand etELITE_STATES.eElite_CHECK_BUZZER, 1
                
                ' Record the sound using the PC mic in
                If False = g_fMicIn.RecordWav Then
                    ' Check control panel on PC to make sure mic is setup correctly
                    UpdateResults lstRow, "Failed to find audio input device", eRESULT_FAILED
                    TestBuzzer = eRESULT_FAILED
                    Exit Do
                End If
                
                If g_dSoundLevel >= g_iMinBuzzerAmplitude Then
                    'Print #iFileNo, "CRITERIA: " + CStr(g_dSoundLevel) + ">" + CStr(g_iMinBuzzerAmplitude)
                    ' Recorded sound meets minimum criteria
                    TestBuzzer = eRESULT_PASSED
                    Exit Do
                Else
                    UpdateResults lstRow, Format(g_dSoundLevel, "0.00") & "@" & Format(g_dSoundFreq, "0") & "hz. Retesting...", eRESULT_TESTING
                End If
                iNumRetries = iNumRetries - 1
            Loop While (eRESULT_TESTING = TestBuzzer) And (1 <> iNumRetries)
            
            If eRESULT_PASSED = TestBuzzer Then
                UpdateResults lstRow, Format(g_dSoundLevel, "0.00") & "@" & Format(g_dSoundFreq, "0") & "hz", eRESULT_PASSED
            Else
                UpdateResults lstRow, Format(g_dSoundLevel, "0.00") & "@" & Format(g_dSoundFreq, "0") & "hz", eRESULT_FAILED
                TestBuzzer = eRESULT_FAILED
            End If
        End If
    End If

End Function

Private Function TestLED(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim LEDlimitmin As Double
    Dim LEDlimitmax As Double
    LEDlimitmin = 0.2
    LEDlimitmax = 0.6
    Dim GrnResult As Double
    Dim RedResult As Double
    Dim GrnResultPwrOff As Double
    Dim RedResultPwrOff As Double

    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test LEDs"
        TestLED = eRESULT_NOT_TESTED
    Else
        TestLED = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestLED = eRESULT_NOT_TESTED
        Else
                
            'AE 8/26/09 check LEDs
            'turn on green LED
            If True = SerialComs.CheckLED(0) Then
                LogMessage "Turned on green LED"
                startup.Delay 1#
                GrnResult = g_fIO.LedVoltage(0)
                
                'turn off green LED
                If True = SerialComs.CheckLED(2) Then
                    LogMessage "Turned off green LED"
                    startup.Delay 1#
                    GrnResultPwrOff = g_fIO.LedVoltage(0)
                Else
                    g_fMainTester.LogMessage "Failed to execute LED " & eElite_Grn_LED & " commands."
                    UpdateResults lstRow, "Failed to turn off Grn LED", eRESULT_FAILED
                    Exit Function
                End If
            Else
                g_fMainTester.LogMessage "Failed to execute LED " & eElite_Grn_LED & " command."
                UpdateResults lstRow, "Failed to turn on Grn LED", eRESULT_FAILED
                Exit Function
            End If
            
            'turn on red LED
            If True = SerialComs.CheckLED(1) Then
                LogMessage "Turned on red LED"
                startup.Delay 1#
                RedResult = g_fIO.LedVoltage(1)
            
                'turn off red LED
                If True = SerialComs.CheckLED(3) Then
                    LogMessage "Turned off red LED"
                    startup.Delay 1#
                    RedResultPwrOff = g_fIO.LedVoltage(1)
                Else
                  LogMessage "Failed to execute LED " & eElite_Red_LED & " commands."
                  UpdateResults lstRow, "Failed to turn off Red LED", eRESULT_FAILED
                  Exit Function
                End If
            Else
                g_fMainTester.LogMessage "Failed to execute LED " & eElite_Red_LED & " command."
                UpdateResults lstRow, "Failed to turn on Red LED", eRESULT_FAILED
                Exit Function
            End If
            
            'check results
            If GrnResult >= LEDlimitmin And GrnResult <= LEDlimitmax Then
                g_fMainTester.LogMessage "Green LED detected"
                If RedResult >= LEDlimitmin And RedResult <= LEDlimitmax Then
                    g_fMainTester.LogMessage "Red LED detected"
                    If GrnResultPwrOff < LEDlimitmin And RedResultPwrOff < LEDlimitmin Then
                        TestLED = eRESULT_PASSED
                        UpdateResults lstRow, "Grn LED: " & FormatNumber(GrnResult, 3) & ", Red LED: " & FormatNumber(RedResult, 3), eRESULT_PASSED
                    Else
                        g_fMainTester.LogMessage "Failed to turn off LED"
                        UpdateResults lstRow, "Grn LED: " & FormatNumber(GrnResult, 3) & ", Red LED: " & FormatNumber(RedResult, 3), eRESULT_FAILED
                    End If
                Else
                    g_fMainTester.LogMessage "Red LED not within .2-.6V"
                    UpdateResults lstRow, "Grn LED: " & FormatNumber(GrnResult, 3) & ", Red LED: " & FormatNumber(RedResult, 3), eRESULT_FAILED
                End If
            Else
              g_fMainTester.LogMessage "Green LED not within .2-.6V"
              UpdateResults lstRow, "Grn LED: " & FormatNumber(GrnResult, 3) & ", Red LED: " & FormatNumber(RedResult, 3), eRESULT_FAILED
            End If
        End If
    End If

End Function

' Get firmware version from file name
Public Function FirmwareVerFromFN(szFN As String) As Boolean
    Dim szVersionMajor As String
    Dim szVersionMinor As String
    
    FirmwareVerFromFN = False
    If 0 = InStr(szFN, ".") Then
        MsgBox "Failed to find version number in file name", vbCritical
        Exit Function
    End If
    
    szVersionMajor = Left(szFN, InStr(szFN, ".") - 1)
    
    If 0 < Len(szVersionMajor) - Len("elite") Then
        If 0 < Len(szFN) - Len(szVersionMajor) - 1 Then
            szVersionMinor = Right(szFN, Len(szFN) - Len(szVersionMajor) - 1)
        Else
            ' szVersionMinor = "0"
            MsgBox "Failed to find minor version number in file name", vbCritical
            Exit Function
        End If
        szVersionMajor = Right(szVersionMajor, Len(szVersionMajor) - Len("elite"))
    Else
        ' szVersionMajor = "0"
        MsgBox "Failed to find major version number in file name", vbCritical
        Exit Function
    End If
    
    
    While 0 < Len(szVersionMajor) And 1 = InStr(szVersionMajor, "0")
        'Trim leading zeros
        szVersionMajor = Right(szVersionMajor, Len(szVersionMajor) - 1)
    Wend
    If 0 = Len(szVersionMajor) Then
        szVersionMajor = "0"
    End If
    g_szFirmwareVersion = szVersionMajor & "." & szVersionMinor
    FirmwareVerFromFN = True
End Function
Private Function TestAppVersion(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Get Application Firmware Version"
        TestAppVersion = eRESULT_NOT_TESTED
    Else
        TestAppVersion = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestAppVersion = eRESULT_NOT_TESTED
        Else
            'AE 8/25/09 check firmware version downloaded with filename in XML
                If SerialComs.GetAppVer(g_szAPPVER) Then
                    'check version
                    If InStr(g_szAPPVER, g_szFirmwareVersion) > 0 Then
                        'correct version
                        TestAppVersion = eRESULT_PASSED
                        UpdateResults lstRow, "F/W=" & g_szAPPVER, eRESULT_PASSED
                    ElseIf (chkDownload.value <> vbChecked) And ("TRAX-II" = g_oSelectedModel.baseName Or "TRAX-II.N" = g_oSelectedModel.baseName) And _
                        ("Elite3.19" = g_szFirmwareFileName Or _
                         "Elite3.3" = g_szFirmwareFileName) Then
                        TestAppVersion = eRESULT_PASSED
                        If "Elite3.19" = g_szFirmwareFileName Then
                            UpdateResults lstRow, "F/W=" & g_szAPPVER, eRESULT_PASSED
                        Else
                            UpdateResults lstRow, "F/W=" & g_szAPPVER, eRESULT_PASSED
                        End If
                    Else
                        'not correct version
                        g_fMainTester.LogMessage g_szAPPVER & " -- " & g_szFirmwareFileName
                        UpdateResults lstRow, "Wrong firmware version detected", eRESULT_FAILED
                        m_bDeviceFailure = True
                    End If
                Else
                  g_fMainTester.LogMessage "AT command AT$APVER to get application version failed."
                  UpdateResults lstRow, "Couldn't find firmware version", eRESULT_FAILED
                  m_bDeviceFailure = True
                End If
        End If
    End If

End Function

Private Function VerifyVoltage(szDUT_Voltage As String, oNode As IXMLDOMNode, ByRef szResult As String) As etRESULTS

    Dim avg, min, max As Integer
    Dim units As String
    Dim pwr_off_wait As Double
    units = oNode.Attributes.getNamedItem("Units").Text
    avg = CInt(oNode.Attributes.getNamedItem("Avg").Text)
    min = CInt(oNode.Attributes.getNamedItem("Min").Text)
    max = CInt(oNode.Attributes.getNamedItem("Max").Text)
    
    'check voltage
    If False = IsNumeric(szDUT_Voltage) Then
        g_fMainTester.LogMessage "command to get DUT voltage failed."
        szResult = "Couldn't read voltage from DUT"
        VerifyVoltage = eRESULT_FAILED
    ElseIf CInt(szDUT_Voltage) < min Or CInt(szDUT_Voltage) > max Then
        ' voltage out of spec
        g_fMainTester.LogMessage "Expected " & avg & units & ". Read " & szDUT_Voltage & units & ". Voltage not in range of " & min & " to " & max & units
        szResult = "Voltage not in range of " & min & " to " & max & units
        VerifyVoltage = eRESULT_FAILED
    Else
        'correct voltage
        g_fMainTester.LogMessage "Expected " & avg & units & ". Read " & szDUT_Voltage & units & ". Voltage in range of " & min & " to " & max & units
        szResult = "DUT voltage = " & szDUT_Voltage
        VerifyVoltage = eRESULT_PASSED
    End If
End Function
Private Function TestSuperCap(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test Supercaps"
        TestSuperCap = eRESULT_NOT_TESTED
    Else
        TestSuperCap = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
            SerialComs.PSOutput = False
            LogMessage "Turned power supply off"
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
            SerialComs.PSOutput = False
            LogMessage "Turned power supply off"
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestSuperCap = eRESULT_NOT_TESTED
            SerialComs.PSOutput = False
            LogMessage "Turned power supply off"
        Else
            Dim szDUT_Voltage As String
            Dim oNode As IXMLDOMNode
            Dim szResult As String
            
            TestSuperCap = eRESULT_TESTING
               
            ' Check the operating voltage
            UpdateResults lstRow, "Reading voltage...", TestSuperCap
            Set oNode = g_oSelectedModel.selectSingleNode("Power").selectSingleNode("Nominal_Voltage")
            If SerialComs.GetVoltsDUT(szDUT_Voltage) Then
                TestSuperCap = VerifyVoltage(szDUT_Voltage, oNode, szResult)
                If TestSuperCap = eRESULT_PASSED Then
                    TestSuperCap = eRESULT_TESTING
                End If
            Else
                g_fMainTester.LogMessage "command to get DUT voltage failed."
                szResult = "Couldn't read voltage from DUT"
                TestSuperCap = eRESULT_FAILED
                UpdateResults lstRow, szResult, TestSuperCap
            End If
            
            If TestSuperCap = eRESULT_TESTING Then
                ' Power off the DUT, wait specified period and read voltage
                Set oNode = g_oSelectedModel.selectSingleNode("Power").selectSingleNode("PWR_Off_Voltage")
                szDUT_Voltage = ""
                
                g_fMainTester.LogMessage "Removing power and sleeping for " & oNode.Attributes.getNamedItem("Wait").Text & "s"
                SerialComs.PSOutput = False
                startup.Delay CDbl(oNode.Attributes.getNamedItem("Wait").Text)
                
                If SerialComs.GetVoltsDUT(szDUT_Voltage) Then
                    TestSuperCap = VerifyVoltage(szDUT_Voltage, oNode, szResult)
                Else
                    g_fMainTester.LogMessage "command to get DUT voltage failed."
                    szResult = "Couldn't read voltage from DUT"
                    TestSuperCap = eRESULT_FAILED
                End If
                If TestSuperCap = eRESULT_PASSED Then
                    szResult = "Confirmed SuperCap Operation"
                End If
                UpdateResults lstRow, szResult, TestSuperCap
            End If
            
            If TestSuperCap <> eRESULT_PASSED Then
                m_bDeviceFailure = True
            End If
            
        End If
    End If

End Function

Private Function TestIMSI(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim iRetries As Integer

    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Get the IMSI from the SIM Card"
        TestIMSI = eRESULT_NOT_TESTED
    Else
        TestIMSI = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestIMSI = eRESULT_NOT_TESTED
        Else
            '   Get the IMSI from the SIM card
            UpdateResults lstRow, "Getting IMSI...", eRESULT_TESTING
            
            For iRetries = 1 To 10
                If SerialComs.GetIMSI(g_szIMSI) Then
                    ' Remove the text tag from the IMSI
                    If Len("CIMI:") < Len(g_szIMSI) Then
                        g_szIMSI = Right(g_szIMSI, Len(g_szIMSI) - Len("CIMI:"))
                    End If
                    UpdateResults lstRow, "IMSI = " & g_szIMSI, eRESULT_PASSED
                    TestIMSI = eRESULT_PASSED
                    Exit For
                End If
            Next iRetries
            If TestIMSI <> eRESULT_PASSED Then
                UpdateResults lstRow, "Failed to get the IMSI", eRESULT_FAILED
            End If
        End If
    End If

End Function

Private Function TestICCID(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim iRetries As Integer

    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Get the ICCID from the SIM Card"
        TestICCID = eRESULT_NOT_TESTED
    Else
        TestICCID = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestICCID = eRESULT_NOT_TESTED
        Else
            '   Go and get the ICCID fron the SIM card
            UpdateResults lstRow, "Getting the ICCID...", eRESULT_TESTING
            
            For iRetries = 1 To 10
                If SerialComs.GetCCID(g_szCCID) Then
                    ' Remove the text tag from the ICCID
                    If Len("SCID:") < Len(g_szCCID) Then
                        g_szCCID = Right(g_szCCID, Len(g_szCCID) - Len("SCID:"))
                    End If
                    UpdateResults lstRow, "ICCID = " & g_szCCID, eRESULT_PASSED
                    TestICCID = eRESULT_PASSED
                    Exit For
                End If
            Next iRetries
            If TestICCID <> eRESULT_PASSED Then
                UpdateResults lstRow, "Failed to get the ICCID", eRESULT_FAILED
            End If
        End If
    End If

End Function

Private Function TestGpsAntenna(lstRow As ListItem, bRunTest As Boolean, Optional bExtAnt As Boolean = False) As etRESULTS

    Dim iRetries As Integer
    Dim iAntCnt As Integer
    Dim dMinGPS As Integer
    Dim dMaxGPS As Integer
    Dim dMinGPS_X As Integer
    Dim dMaxGPS_X As Integer
    Dim oMinAttribute As IXMLDOMAttribute
    Dim oMaxAttribute As IXMLDOMAttribute
    Dim oAttributes As IXMLDOMNamedNodeMap
    Dim eGpsAnt As etELITE_STATES
    Dim szGPS As String
    
    Set oAttributes = g_oSettings.selectSingleNode("GPSSimulator").Attributes
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Get GPS SV" + oAttributes.getNamedItem("SVID").Text + " Signal Strength"
        TestGpsAntenna = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestGpsAntenna = eRESULT_NOT_TESTED
        Else
            UpdateResults lstRow, "Getting signal strength...", eRESULT_TESTING
            g_szIntGPS = 0
            g_szExtGPS = 0
            
            ' Attempt to find GPS up to three times
            Do
                TestGpsAntenna = eRESULT_TESTING
                If True = bExtAnt Then
                    ' Signal coax switch in tester to disconnect from the tester's internal
                    ' antenna and connect to the DUT's external connector
                    If False = g_fIO.GPS_Ext Then
                        g_fIO.GPS_Ext = True
                    End If
                    
                    If Not SetGpsAntenna(eElite_GPS_EXT_ANT) Then
                        TestGpsAntenna = eRESULT_FAILED
                        Exit Do
                    End If
                    eGpsAnt = eElite_GPS_EXT_ANT
                
                    '   Get the minimum and maximum GPS limits for selected Model Type
                    Set oMinAttribute = oAttributes.getNamedItem("xMin")
                    Set oMaxAttribute = oAttributes.getNamedItem("xMax")
                    dMinGPS = CInt(oMinAttribute.Text)
                    dMaxGPS = CInt(oMaxAttribute.Text)
        
                Else
                    ' Signal coax switch in tester to disconnect from the DUT's external connector
                    ' and connect to the tester's antenna
                    If True = g_fIO.GPS_Ext Then
                        g_fIO.GPS_Ext = False
                    End If
                    
                    If Not SetGpsAntenna(eElite_GPS_INT_ANT) Then
                        TestGpsAntenna = eRESULT_FAILED
                        Exit Do
                    End If
                    eGpsAnt = eElite_GPS_INT_ANT
                
                    '   Get the minimum and maximum GPS limits for selected Model Type
                    Set oMinAttribute = oAttributes.getNamedItem("Min")
                    Set oMaxAttribute = oAttributes.getNamedItem("Max")
                    dMinGPS = CInt(oMinAttribute.Text)
                    dMaxGPS = CInt(oMaxAttribute.Text)
        
                End If
                
                szGPS = ""
                For iRetries = g_iGPSTimeout To 1 Step -1
                    If m_eTestingState = eTESTER_OPEN Then
                        UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                        TestGpsAntenna = eRESULT_FAILED
                        Exit For
                    ElseIf "" <> szGPS Then
                        Dim szGPS_Antenna
                        If True = g_fIO.GPS_Ext Then
                            szGPS_Antenna = "Ext."
                        Else
                            szGPS_Antenna = "Int."
                        End If
                        UpdateResults lstRow, g_szExtGPS & "dBm (" & szGPS_Antenna & "); Waiting for Good Signal: " & iRetries, eRESULT_TESTING
                    End If
                    startup.Delay 1#
                    UpdateTestCycleTime
                    DoEvents
                    szGPS = g_szGPSID
                    If SerialComs.FindGPS(szGPS) Then
                        If CInt(szGPS) >= dMinGPS And CInt(szGPS) <= dMaxGPS Then
                            Exit For
                        End If
                    Else
                        TestGpsAntenna = eRESULT_FAILED
                        Exit For
                    End If
                Next iRetries
                
                If "" = szGPS Or eRESULT_TESTING <> TestGpsAntenna Then
                    TestGpsAntenna = eRESULT_FAILED
                    szGPS = "0"
                End If
                
                If True = bExtAnt Then
                    ' Save GPS test results for the external antenna
                    g_szExtGPS = szGPS
                    
                    ' We are done testing the external antenna
                    bExtAnt = False
                Else
                    ' Save GPS test results for the internal antenna
                    g_szIntGPS = szGPS
                End If
                
                ' Verify results
                If eRESULT_FAILED <> TestGpsAntenna Then
                    ' Either not testing external antenna or external antenna didn't fail
                    If CInt(szGPS) >= dMinGPS And CInt(szGPS) <= dMaxGPS Then
                        TestGpsAntenna = eRESULT_PASSED
                    Else
                        TestGpsAntenna = eRESULT_FAILED
    
                    End If
                End If
            Loop While eElite_GPS_EXT_ANT = eGpsAnt
            
            m_test_GPS_Sig = g_szIntGPS & "dBm"
            UpdateResults lstRow, "GPS: " & g_szIntGPS & "dBm (Int.); " & g_szExtGPS & "dBm (Ext.)", TestGpsAntenna
        End If
    End If

End Function

Private Function VerifySignalStrengthGPRS(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim szMin As String
    Dim szMax As String
    Dim bIMSI As Boolean
    Dim szTestID As String
    Dim iRetries As Integer
    Dim oAttributes As IXMLDOMNamedNodeMap
    Dim bSerialCmdStatus As Boolean
    Dim bMaxWaitVerfiyCSQ As Integer
    Dim bUseCallBox As Boolean

    bUseCallBox = 0 <> CInt(Val(Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PORT)))
    
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Cell Site Simulator Test"
        VerifySignalStrengthGPRS = eRESULT_NOT_TESTED
    Else
        '   Now go find the simulator
        VerifySignalStrengthGPRS = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            VerifySignalStrengthGPRS = eRESULT_NOT_TESTED
        ElseIf m_bDownloadFailed Then
            UpdateResults lstRow, "Skipped due to download failure", eRESULT_NOT_TESTED
            VerifySignalStrengthGPRS = eRESULT_NOT_TESTED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            VerifySignalStrengthGPRS = eRESULT_NOT_TESTED
        Else
            VerifySignalStrengthGPRS = eRESULT_TESTING
            
            '   Start looking for the simualtor (we need to get the station ID
            '   and minimum signal strength from the XML file first).
            Set oAttributes = g_oSettings.selectSingleNode("CellSiteSimulator").Attributes
            szTestID = oAttributes.getNamedItem("Station").Text
            
            ' set to bypass
            ' No response expected so SendATCommand eElite_ENABLE_BYPASS always
            ' returns true
            SerialComs.SendATCommand etELITE_STATES.eElite_ENABLE_BYPASS, 1
            
            ' Previously would only execute following code if bUseCallBox and the
            ' model type was not PTE/C-2 or PTE/C-2X
            If True = bUseCallBox Then
                Dim bFindSimRetries As Byte
                
                UpdateResults lstRow, "Registering with call box...", eRESULT_TESTING
                
                ' Setup the test equipment
                ComPortToCallBox.PortOpen = True
                
                ' Deregister from network and remain unregistered
                ' Need a longer than normal timeout for response
                SerialComs.SendATCommand eElite_DISABLE_AUTO_GSM_REG, 5000
                
                ' Set the cops command format
                SerialComs.SetStationIdFormatDefault
                
                ' Tell call box to issue a location update
                ' This will cause the modem to communicate with the callbox
                SerialComs.TransmitToCB etCB_STATES.eCB_LOC_UPD
                While g_eCB_COM_Status = eSTAT_RUNNING
                    UpdateTestCycleTime
                    DoEvents
                Wend
                
                startup.Delay 1#
                bSerialCmdStatus = SerialComs.SendATCommand(eElite_RESET_GSM_MODULE)
                UpdateTestCycleTime
                
                startup.Delay 1#
                UpdateTestCycleTime
                
               If False = SendATCommand(eElite_SET_MANUAL_REG_MODE_AND_SELECT_OPERATOR, 3000) Then
                    g_fMainTester.LogMessage (" Set reg. mode and select operator command failed. Retrying...")
                    startup.Delay 1#
                    
                    If False = SendATCommand(eElite_SET_MANUAL_REG_MODE_AND_SELECT_OPERATOR, 3000) Then
                        ' Give up
                        g_fMainTester.LogMessage (" Set reg. mode and select operator command failed.")
                    End If
                End If
            
                ' Set how long to wait before assuming device won't register
                bFindSimRetries = g_iModemRegWait
                startup.Delay 1#
                UpdateTestCycleTime
                Do Until 0 = bFindSimRetries Or True = FindSimulator
                    UpdateResults lstRow, "Registering with call box..." & bFindSimRetries, eRESULT_TESTING
                    UpdateTestCycleTime
                    DoEvents
                    If eTESTER_OPEN = m_eTestingState Then
                        bFindSimRetries = 0
                        Exit Do
                    End If
                    bFindSimRetries = bFindSimRetries - 1
                    startup.Delay 1#
                Loop
                
                
                ComPortToCallBox.PortOpen = False
                
                ' Check if timeout occurred registering to the call box
                If 0 = bFindSimRetries Then
                    ' Give up
                    g_fMainTester.LogMessage (" Failed to register with call box")
                    VerifySignalStrengthGPRS = etRESULTS.eRESULT_FAILED
                End If
            End If
            
            If etRESULTS.eRESULT_TESTING = VerifySignalStrengthGPRS Then
                UpdateResults lstRow, "Testing with simulator...", eRESULT_TESTING
            
                If True = bUseCallBox Then
                    bMaxWaitVerfiyCSQ = g_iModemCSQ_Wait
                End If
                For iRetries = bMaxWaitVerfiyCSQ To 1 Step -1
                    szMin = oAttributes.getNamedItem("Min").Text
                    szMax = oAttributes.getNamedItem("Max").Text
                    
                    g_fMainTester.LogMessage ("Test Station " & ComputerName & " GSM Attenuation: " & CStr(g_iGSM_Attenuation))
    
                    ' Check the GPRS Signal Strength
                    If szMin = szMax Then
                        ' Need to have these values not equal to each other
                        szMin = CStr(CInt(szMin) - 1)
                    End If
                    
                    bSerialCmdStatus = SerialComs.VerifySignalStrengthGPRS(szMin, szMax)
                    
                    If True = bSerialCmdStatus Then
                        If szMax = szMin Then
                            VerifySignalStrengthGPRS = eRESULT_PASSED
                            Exit For
                        Else
                            If "99" = szMin Then
                                UpdateResults lstRow, "No signal. Testing with simulator..." & iRetries, eRESULT_TESTING
                            Else
                                UpdateResults lstRow, "Low signal (" & szMin & "). Testing with simulator..." & iRetries, eRESULT_TESTING
                            End If
                            startup.Delay 1#
                            If Not g_fIO.IsTesterClosed Then
                                Exit For
                            End If
                        End If
                    Else
                        ' AT command failed so no point retrying
                        VerifySignalStrengthGPRS = eRESULT_FAILED
                        Exit For
                    End If
                    
                Next iRetries
                
            End If
            
            ' bAtCmdStatus = SerialComs.SendATCommand(eElite_RESET_GSM_MODULE)
        
            If eRESULT_PASSED <> VerifySignalStrengthGPRS Then
                m_test_GSM_Sig = szMin
                g_fMainTester.LogMessage (" Cell Simulator test failed: Signal=" & m_test_GSM_Sig)
                UpdateResults lstRow, "Cell Simulator test failed", eRESULT_FAILED
                VerifySignalStrengthGPRS = eRESULT_FAILED
                m_bDeviceFailure = True
            End If
            
            ' Want to put the deivice back in automtatic mode when done
            ' Not necessary on TraxII and TraxIIC
            
            Dim loopCnt As Integer
            loopCnt = 10
            Do
                'This is a big deal. We need to be back in auto mode
                If True = SendATCommand(etELITE_STATES.eElite_SET_STATION_AUTOMATIC, 20000) Then
                    g_fMainTester.LogMessage (" Set registration mode to automatic.")
'                    If eRESULT_PASSED = VerifySignalStrengthGPRS Then
'                        UpdateResults lstRow, "Signal=" & szMin, eRESULT_PASSED
'                    Else
'                        UpdateResults lstRow, "Cell Simulator test failed. Low signal (" & szMin & ").", eRESULT_FAILED
'                    End If
                    Exit Do
                Else
                    If eRESULT_PASSED = VerifySignalStrengthGPRS Then
                        ' Only set if message if no previous failure message is being displayed
                        UpdateResults lstRow, "Setting registration mode to automatic..." & loopCnt, eRESULT_TESTING
                    End If
                End If
                startup.Delay 1#
                loopCnt = loopCnt - 1
            Loop Until 0 = loopCnt
            
            If 0 = loopCnt Then
                g_fMainTester.LogMessage (" Set station automatic mode failed.")
                If eRESULT_PASSED = VerifySignalStrengthGPRS Then
                    ' Only set if message if no previous failure message is being displayed
                    UpdateResults lstRow, "Set station automatic mode failed." & loopCnt, eRESULT_FAILED
                End If
                VerifySignalStrengthGPRS = eRESULT_FAILED
                m_bDeviceFailure = True
            Else
                ' BGS2 Modem had reported failure to retain automatic mode setting across power cycles
                ' Explicit shutoff command for modem appears to resolve issue per Passtime on 7/3/2013
                startup.Delay 1#
                If False = SendATCommand(etELITE_STATES.eElite_BGS2_SHUTOFF, 7000) Then
                    g_fMainTester.LogMessage (" Modem shutoff command failed.")
                    VerifySignalStrengthGPRS = eRESULT_FAILED
                    m_bDeviceFailure = True
                End If
            End If
            
            ' No response expected from SendATCommand(eElite_DISABLE_BYPASS); always returns true
             SerialComs.SendATCommand etELITE_STATES.eElite_DISABLE_BYPASS, 1
            
            If eRESULT_PASSED = VerifySignalStrengthGPRS Then
                m_test_GSM_Sig = szMin
                UpdateResults lstRow, "Signal=" & m_test_GSM_Sig, eRESULT_PASSED
            End If
        
        End If
    End If
End Function


 Private Function TestRfRcvr(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim bState As String
    
    If Not bRunTest Then
      ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test RF Receiver"
      TestRfRcvr = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestRfRcvr = eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        Else
            TestRfRcvr = eRESULT_FAILED
            UpdateResults lstRow, "Checking RF Receiver...", eRESULT_TESTING
            
            g_fIO.IgnitionInput = True
            startup.Delay 1#
            
            ' Call RF_Output True to toggle transmitter on for 1/4 sec. and then back off
            g_fIO.RF_Output = True
            
            If SerialComs.GetRfRcvrInput(bState) Then
                If bState = RF_XMT_CHAR Or bState = RF_ALT_RCV_CHAR Then
                    TestRfRcvr = eRESULT_PASSED
                    UpdateResults lstRow, "Received " & bState & " via RF", eRESULT_PASSED
                Else
                    UpdateResults lstRow, "Received " & bState & " via RF; expected " & RF_XMT_CHAR & " or " & RF_ALT_RCV_CHAR & ".", eRESULT_FAILED
                End If
            Else
                UpdateResults lstRow, "Get RF receiver input command failed.", eRESULT_FAILED
            End If
            g_fIO.IgnitionInput = False
            startup.Delay 1#
        End If
    End If

End Function

Private Function TestRelays(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim iRetries As Integer
    Dim dSoundLevel As Double
    Dim bState As String

    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test Relay"
        TestRelays = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            TestRelays = eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        Else
            '   Test the relays in the old Plus part of the device
            TestRelays = eRESULT_FAILED
            UpdateResults lstRow, "Checking relays...", eRESULT_TESTING
            
            ' check that relay is off
            If SerialComs.CheckEliteRelay(bState, 2) Then
                If bState <> "0" Or g_fIO.DUTRelayState <> RELAY_AFTER_POWER_DOWN Then
                UpdateResults lstRow, "Relay is not open", eRESULT_FAILED
                Exit Function
                End If
                
            Else
                UpdateResults lstRow, "Could not get relay status", eRESULT_FAILED
                Exit Function
            End If
            'turn on relay
            If SerialComs.CheckEliteRelay(bState, 1) Then
                g_fMainTester.LogMessage "Relay closed"
            Else
                UpdateResults lstRow, "Could not close relay", eRESULT_FAILED
                Exit Function
            End If
            startup.Delay 1#
            ' check that relay is closed
            If SerialComs.CheckEliteRelay(bState, 2) Then
                If bState <> "1" Or g_fIO.DUTRelayState <> RELAY_AFTER_IGNITION Then
                    UpdateResults lstRow, "Relay is not closed", eRESULT_FAILED
                    Exit Function
                End If
            Else
                UpdateResults lstRow, "Could not get relay status", eRESULT_FAILED
                Exit Function
            End If
             'turn off relay
            If SerialComs.CheckEliteRelay(bState, 0) Then
                g_fMainTester.LogMessage "Relay opened"
            Else
                UpdateResults lstRow, "Could not open relay", eRESULT_FAILED
                Exit Function
            End If
            startup.Delay 1#
            ' check that relay is open
            If SerialComs.CheckEliteRelay(bState, 2) Then
                If bState = "0" And g_fIO.DUTRelayState = RELAY_AFTER_POWER_DOWN Then
                    TestRelays = eRESULT_PASSED
                    UpdateResults lstRow, "Relay passed", eRESULT_PASSED
                Else
                    UpdateResults lstRow, "Relay is not open", eRESULT_FAILED
                End If
            Else
                UpdateResults lstRow, "Could not get relay status", eRESULT_FAILED
            End If
        End If
    End If

End Function

Private Function ConfigureAntiTheft(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    Dim szState As String
'    Dim oGpsModeAttribute As IXMLDOMAttribute
    Dim oAttributes As IXMLDOMNamedNodeMap

    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set AntiTheft mode"
        Else
             ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify AntiTheft Mode"
        End If
        
        ConfigureAntiTheft = eRESULT_NOT_TESTED
    Else
        ConfigureAntiTheft = eRESULT_FAILED
        
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureAntiTheft = eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        Else
            
            ConfigureAntiTheft = eRESULT_TESTING
            ' Only need to set antitheft on device not actually test functionality thus the commented out code below
            
'            ' check that relay is off
'            If SerialComs.CheckEliteRelay(bState, 2) Then
'                If bState <> "0" Or g_fIO.DUTRelayState <> RELAY_AFTER_POWER_DOWN Then
'                UpdateResults lstRow, "Relay is not open", eRESULT_FAILED
'                Exit Function
'                End If
'
'            Else
'            UpdateResults lstRow, "Could not get relay status", eRESULT_FAILED
'            Exit Function
'            End If
'            'check that nothing is on ignition line
'            g_fIO.IgnitionInput = False
'            startup.Delay 1#
'            If SerialComs.CheckIgnition(bState) Then
'                If bState = "0" Then
'                    g_fMainTester.LogMessage "Ignition line initially off"
'                Else
'                    UpdateResults lstRow, "Ignition line is not initially 0", eRESULT_FAILED
'                    Exit Function
'                End If
'            Else
'                UpdateResults lstRow, "Ignition command failed", eRESULT_FAILED
'                Exit Function
'            End If

'            oModel.selectSingleNode("Security").Attributes.getNamedItem ("Valet")
            Set oAttributes = g_oSelectedModel.selectSingleNode("Security").Attributes
'            If SerialComs.SetAntiTheft(oGpsModeAttribute.Text) Then
            If vbChecked = chkTrialRun.value Then
                UpdateResults lstRow, "Verifying Anti Theft Setting...", eRESULT_TESTING
            Else
                UpdateResults lstRow, "Setting Configuring Anti Theft...", eRESULT_TESTING
                If False = SerialComs.SetAntiTheft(oAttributes.getNamedItem("Valet").Text) Then
                    ConfigureAntiTheft = eRESULT_FAILED
                End If
            End If
            If eRESULT_TESTING = ConfigureAntiTheft Then
                If vbChecked = chkTrialRun.value Then
                    ConfigureAntiTheft = eRESULT_NOT_TESTED
                    UpdateResults lstRow, "Skipped", eRESULT_NOT_TESTED
               Else
                    ConfigureAntiTheft = eRESULT_PASSED
                    UpdateResults lstRow, "Configured anti-theft mode", eRESULT_PASSED
                End If
            Else
                ConfigureAntiTheft = eRESULT_FAILED
                UpdateResults lstRow, "Failed to configure anti-theft mode", eRESULT_FAILED
            End If
'
'             'start and check that something is on ignition line
'            g_fIO.IgnitionInput = True
'            startup.Delay 1#
'            If SerialComs.CheckIgnition(bState) Then
'                If bState = "1" Then
'                    g_fMainTester.LogMessage "Ignition line on"
'                Else
'                    UpdateResults lstRow, "Ignition line is not on", eRESULT_FAILED
'                    Exit Function
'                End If
'            Else
'                UpdateResults lstRow, "Ignition command failed", eRESULT_FAILED
'                Exit Function
'            End If
'
'            ' check that relay is open
'            If SerialComs.CheckEliteRelay(bState, 2) Then
'                If bState = "0" And g_fIO.DUTRelayState = RELAY_AFTER_POWER_DOWN Then
'                ConfigureAntiTheft = eRESULT_PASSED
'                UpdateResults lstRow, "ConfigureAntiTheft passed", eRESULT_PASSED
'                Else
'                UpdateResults lstRow, "Car started - ConfigureAntiTheft failed.", eRESULT_FAILED
'                End If
'            Else
'            UpdateResults lstRow, "Could not get relay status", eRESULT_FAILED
'            End If
'
'            'turn off ignition line but don't bother checking
'            g_fIO.IgnitionInput = False
'            startup.Delay 1#
       End If
    End If

End Function


Private Function ConfigureSMSmode(lstRow As ListItem, bRunTest) As etRESULTS
    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set SMS Mode as Default"
        Else
             ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify SMS Mode is Default"
        End If
        
        ConfigureSMSmode = eRESULT_NOT_TESTED
    Else
        ConfigureSMSmode = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureSMSmode = eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            ConfigureSMSmode = eRESULT_NOT_TESTED
        Else
            Dim smsResetMin As String
            Dim smsMode As String
            Dim smsHelloPktTimer As String
            Dim iStrIdx As Integer
            Dim oNamedNodeMap As IXMLDOMNamedNodeMap
            Dim oSmsAttribute As IXMLDOMAttribute
            Dim oTag As IXMLDOMNode
            Dim smsState As String
    
            ConfigureSMSmode = eRESULT_TESTING
            Set oNamedNodeMap = g_oSelectedSIM.selectSingleNode("SMS_Mode").selectSingleNode("Mode").Attributes
            
            Set oSmsAttribute = oNamedNodeMap.getNamedItem("HelloPktInterval")
            smsHelloPktTimer = oSmsAttribute.Text
            
            Set oSmsAttribute = oNamedNodeMap.getNamedItem("State")
            smsState = oSmsAttribute.Text
            
            If eRESULT_TESTING = ConfigureSMSmode Then
                UpdateResults lstRow, "Reading current SMS mode...", eRESULT_TESTING
                smsMode = GetSMSmodeString()
                ' Let's see if we even need to change the mode
                If InStr(smsMode, smsState & "," & smsHelloPktTimer) > 0 Then
                    UpdateResults lstRow, "Default is SMS Mode", eRESULT_PASSED
                    ConfigureSMSmode = eRESULT_PASSED
                    smsMode = ""
                End If
                If "" <> smsMode Then
                    ' Parse the response to get the first paramter (state paramter)
                    ' The two paramters are separated by a comma
                    iStrIdx = InStr(1, smsMode, ",")
                    If 0 < (Len(smsMode) - iStrIdx) Then
                        ' Trim away the right side of the string after the comma
                        smsMode = Left(smsMode, iStrIdx)
                        ' Concatenate the state parameter we read in with the
                        ' value we want to set the "hello packet timer" to
                        smsMode = smsMode & smsHelloPktTimer
                        UpdateResults lstRow, "Setting SMS hello packet timer...", eRESULT_TESTING
                        ' change the SMS mode,"hello packet timer" first then change the SMS state
                        If SendSMSmodeString(smsMode) Then
                            UpdateResults lstRow, "Setting SMS state...", eRESULT_TESTING
                            ' now change the SMS state
                            smsMode = smsState & "," & smsHelloPktTimer
                            If SendSMSmodeString(smsMode) Then
                                UpdateResults lstRow, "Default is now SMS Mode", eRESULT_PASSED
                                ConfigureSMSmode = eRESULT_PASSED
                            End If
                        End If
                    End If
                End If
            End If
            If Not (eRESULT_PASSED = ConfigureSMSmode) Then
                UpdateResults lstRow, "Couldn't set SMS mode", eRESULT_FAILED
                ConfigureSMSmode = eRESULT_FAILED
            End If
       End If
    End If
End Function

Private Function ConfigureSMSreplyAddr(lstRow As ListItem, bRunTest) As etRESULTS
    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set and Verify SMS Reply Address"
        Else
             ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify SMS Reply Address"
        End If
        
        ConfigureSMSreplyAddr = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
            ConfigureSMSreplyAddr = eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
            ConfigureSMSreplyAddr = eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureSMSreplyAddr = eRESULT_NOT_TESTED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            ConfigureSMSreplyAddr = eRESULT_NOT_TESTED
        Else
            '   Configure the SMS reply addresses
            If vbChecked = chkTrialRun.value Then
                UpdateResults lstRow, "Verifying SMS reply address...", eRESULT_TESTING
            Else
                UpdateResults lstRow, "Setting SMS reply address...", eRESULT_TESTING
            End If
            ConfigureSMSreplyAddr = eRESULT_TESTING
            
            Dim iRetries As Integer
            Dim szParam As String
            Dim oDOM_Node As IXMLDOMNode
            Dim szSingleNodeTag As String
            Dim smsReplyAddress As String
            Dim iIdx As Integer
            Dim szSMS_Reply(6) As String
            
            szSMS_Reply(6) = 1
            szSMS_Reply(5) = 1
            ' Maybe not necessary but just in case
            szSMS_Reply(4) = ""
            szSMS_Reply(3) = ""
            szSMS_Reply(2) = ""
            szSMS_Reply(1) = ""
            szSMS_Reply(0) = ""
            
            For Each oDOM_Node In g_oSelectedSIM.selectSingleNode("SMS_Reply").childNodes
                ' Retrieve the SMS reply addresses from the XML file
                smsReplyAddress = oDOM_Node.Attributes.getNamedItem("ADDR").Text
                iIdx = CInt(oDOM_Node.Attributes.getNamedItem("Index").Text)
                szSMS_Reply(iIdx - 1) = smsReplyAddress
                
                If Not vbChecked = chkTrialRun Then
                    ' SMS reply register holds up to five addresses.
                    ' Command to set and address relies on the third "mode" paramater being "1".
                    ' Param 2 is the index (1 to 5) of the address field to overwrite
                    ' Param 1 is the address that will be written to the register at the specified index
                    szParam = smsReplyAddress & "," & CStr(iIdx) & ",1"
                    For iRetries = 1 To 3
                        If SerialComs.SendSMSaddressString(szParam) Then
                            iRetries = 0
                            Exit For
                        End If
                        startup.Delay 0.25
                    Next iRetries
                    
                    If iRetries > 0 Then
                        ' AT command failed
                        g_fMainTester.LogMessage ("Failed to set SMS reply address " & iIdx & " to " & smsReplyAddress)
                        UpdateResults lstRow, "Failed to set SMS reply address", eRESULT_FAILED
                        ConfigureSMSreplyAddr = eRESULT_FAILED
                        Exit For
                    End If
                End If
            Next oDOM_Node
                
                
            ' Verify the SMS reply addresses
            ' Easiest way is to construct the expected return value
            ' and compare it to the return value in szParam
            smsReplyAddress = szSMS_Reply(0)
            For iIdx = 1 To UBound(szSMS_Reply)
                smsReplyAddress = smsReplyAddress & "," & szSMS_Reply(iIdx)
            Next iIdx
            
            If szParam = smsReplyAddress Then
                UpdateResults lstRow, "Verified SMS Reply = " & smsReplyAddress, eRESULT_PASSED
                If vbChecked = chkTrialRun.value Then
                    ConfigureSMSreplyAddr = eRESULT_PASSED
                End If
            Else
                g_fMainTester.LogMessage ("Failed to set SMS Reply. Expected " & smsReplyAddress & ", received " & szParam & ".")
                UpdateResults lstRow, "Failed to set SMS reply address", eRESULT_FAILED
                ConfigureSMSreplyAddr = eRESULT_FAILED
            End If
            
            If eRESULT_TESTING = ConfigureSMSreplyAddr Then
                UpdateResults lstRow, "Set SMS Reply = " & smsReplyAddress, eRESULT_PASSED
                ConfigureSMSreplyAddr = eRESULT_PASSED
            ElseIf eRESULT_FAILED = ConfigureSMSreplyAddr Then
                m_bDeviceFailure = True
            End If
        End If
    End If
End Function

Private Function ConfigureAutoResets(lstRow As ListItem, bRunTest) As etRESULTS
    If Not bRunTest Then
        ConfigureAutoResets = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
            ConfigureAutoResets = eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
            ConfigureAutoResets = eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigureAutoResets = eRESULT_NOT_TESTED
        ElseIf Not m_bConfigureSuccessful Then
            UpdateResults lstRow, "Skipped due to a previous configuration failure", eRESULT_NOT_TESTED
            ConfigureAutoResets = eRESULT_NOT_TESTED
        Else
            '   Configure the auto resets
            If vbChecked = chkTrialRun.value Then
                UpdateResults lstRow, "Verifying auto-resets...", eRESULT_TESTING
            Else
                UpdateResults lstRow, "Setting auto-resets...", eRESULT_TESTING
            End If
            ConfigureAutoResets = eRESULT_TESTING
            
            Dim iRetries As Integer
            Dim szParam As String
            Dim oDOM_Node As IXMLDOMNode
            Dim szAutoResetName As String
            Dim szAutoReset As String
            Dim iIdx As Integer
            
            For Each oDOM_Node In g_oSelectedSIM.selectSingleNode("Reset").childNodes
                ' Retrieve the auto-resets from the XML file
                szAutoReset = oDOM_Node.Attributes.getNamedItem("Minutes").Text
                iIdx = CInt(oDOM_Node.Attributes.getNamedItem("Index").Text)
                szAutoResetName = oDOM_Node.Attributes.getNamedItem("Name").Text
                
                If Not vbChecked = chkTrialRun Then
                    ' Set the Auto-reset
                    szParam = szAutoReset
                    For iRetries = 1 To 3
                        If SerialComs.SetResetMin(szParam) Then
                            iRetries = 0
                            Exit For
                        End If
                        startup.Delay 0.25
                    Next iRetries
                    
                    If iRetries > 0 Then
                        ' AT command failed
                        UpdateResults lstRow, "Failed to set " & szAutoResetName & " auto-reset", eRESULT_FAILED
                        ConfigureAutoResets = eRESULT_FAILED
                        Exit For
                    End If
                End If
                
                
                ' Verify the auto-reset
                If eRESULT_TESTING = ConfigureAutoResets Then
                    ' Verify the auto-reset got set
                    For iRetries = 1 To 3
                        szParam = ""
                        If True = SerialComs.GetResetMin(szParam) Then
                            ' AT command success
                            iRetries = 0
                            Exit For
                        End If
                        startup.Delay 0.25
                    Next iRetries
                    
                    If iRetries > 0 Then
                        ' AT command failed
                        UpdateResults lstRow, "Failed to verify " & szAutoResetName & " auto-reset", eRESULT_FAILED
                        ConfigureAutoResets = eRESULT_FAILED
                        Exit For
                    Else
                        If 0 < InStr(szParam, ",") Then
                            szParam = Left(szParam, InStr(szParam, ",") - 1)
                        End If
                        If szParam <> szAutoReset Then
                            UpdateResults lstRow, "Expected " & szAutoReset & ", received " & szParam & ". Failed to verify auto-reset " & iIdx & "'s IP", eRESULT_FAILED
                            ConfigureAutoResets = eRESULT_FAILED
                            Exit For
                        Else
                            UpdateResults lstRow, "Verified " & szAutoResetName & " auto-reset = " & szAutoReset, eRESULT_PASSED
                            If vbChecked = chkTrialRun.value Then
                                ConfigureAutoResets = eRESULT_PASSED
                            End If
                        End If
                    End If
                End If
                iIdx = iIdx + 1
            Next oDOM_Node
            
            If eRESULT_TESTING = ConfigureAutoResets Then
                UpdateResults lstRow, "Set " & szAutoResetName & " auto-reset " & szAutoReset, eRESULT_PASSED
                ConfigureAutoResets = eRESULT_PASSED
            ElseIf eRESULT_FAILED = ConfigureAutoResets Then
                m_bDeviceFailure = True
            End If
        End If
    End If
End Function

Private Function ConfigurePwrMode(lstRow As ListItem, bRunTest) As etRESULTS
    Dim oGpsModeAttribute As IXMLDOMAttribute
    Dim oAttributes As IXMLDOMNamedNodeMap
    
    If Not bRunTest Then
        If Not chkTrialRun.value = vbChecked Then
            ' lstRow.SubItems(eITEM_DESCRIPTION) = "Set Low Power Mode"
        Else
             ' lstRow.SubItems(eITEM_DESCRIPTION) = "Verify Low Power Mode"
        End If
        
        ConfigurePwrMode = eRESULT_NOT_TESTED
    Else
        ConfigurePwrMode = eRESULT_FAILED
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bDeviceFailure Then
            UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
            ConfigurePwrMode = eRESULT_NOT_TESTED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        Else
            ConfigurePwrMode = eRESULT_TESTING
            Set oAttributes = g_oSelectedModel.selectSingleNode("GPS").Attributes
            Set oGpsModeAttribute = oAttributes.getNamedItem("AlwaysOn")

            If chkTrialRun.value <> vbChecked Then
                UpdateResults lstRow, "Setting Low Power Mode...", eRESULT_TESTING
                If False = SendLowPowerCommandString(oGpsModeAttribute.Text) Then
                    UpdateResults lstRow, "Couldn't Turn off GPS System", eRESULT_FAILED
                    ConfigurePwrMode = eRESULT_FAILED
                End If
            End If
            If eRESULT_TESTING = ConfigurePwrMode Then
                If chkTrialRun.value <> vbChecked Then
                    UpdateResults lstRow, "Low Power Mode Set", eRESULT_PASSED
                    ConfigurePwrMode = eRESULT_PASSED
                Else
                    UpdateResults lstRow, "Skipped", eRESULT_NOT_TESTED
                    ConfigurePwrMode = eRESULT_NOT_TESTED
                End If
            End If
        End If
    End If
End Function

Private Function TestRelayAfterPowerDown(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    If Not bRunTest Then
        ' lstRow.SubItems(eITEM_DESCRIPTION) = "Test 'Fail Safe' mode"
        TestRelayAfterPowerDown = eRESULT_NOT_TESTED
    Else
        If m_eTestingState = eTESTER_OPEN Then
            UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
        ElseIf m_bApplicationHasBeenReset Then
            UpdateResults lstRow, "APP REBOOTED: Abort test", eRESULT_FAILED
        Else
            ' With the addition of super caps in spring 2013 this test was
            ' moved to the beginning of the test sequence. This still doesn't
            ' ensure that the supercaps are discharged before testing. The use
            ' case is that a test failed and the operator restarted the test.
            ' Added test to check voltage before allowing this test

            '   Ensure that the relays have fallen back to their fail-safe
            '   positions after power has been removed from the device.
            UpdateResults lstRow, "Checking relay...", eRESULT_TESTING
            g_fIO.StarterInput = False
            g_fIO.IgnitionInput = False
            startup.Delay 0.2
            If g_fIO.DUTRelayState = RELAY_AFTER_POWER_DOWN Then
                UpdateResults lstRow, "Relay in Fail-Safe position", eRESULT_PASSED
                TestRelayAfterPowerDown = eRESULT_PASSED
            Else
                UpdateResults lstRow, "Relay Stuck", eRESULT_FAILED
                TestRelayAfterPowerDown = eRESULT_FAILED
            End If
        End If
    End If

End Function

Private Sub cmdResetCounter_Click()
    
    Globals.PassedParts = 0
    Globals.FailedParts = 0
    FormUpdateCounter

End Sub

Private Sub FormUpdateCounter()
    
    txtTotalParts.Text = CStr(Globals.PassedParts + Globals.FailedParts)
    txtPassedParts.Text = CStr(Globals.PassedParts)
    txtFailedParts.Text = CStr(Globals.FailedParts)

End Sub

Private Sub FormRefresh()

    Dim i As Integer
    Dim iEnd As Integer
    Dim iRow As Integer
    Dim iTest As Integer
    Dim iStart As Integer
    Dim iAdd As Integer
    Dim szTests As String
    Dim lstRow As ListItem
    Dim szDownload As String
    Dim szWireless As String
    Dim szConfigure As String
    Dim abTests() As Boolean
    Dim fhNumber As Integer
    Dim szTestSeqFN As String
    
    ReDim abTests(g_iNumOfTests) As Boolean
    '   Start by clearing out all of the rows on the display screen
    For Each lstRow In lvwTestResults.ListItems
        lstRow.Text = ""
        lstRow.Tag = ""
        For i = 1 To eITEM_NO_OF_ITEMS - 1
            lstRow.SubItems(i) = ""
        Next i
        '   Clear out the test flags while we're at it
        abTests(lstRow.Index) = False
    Next lstRow
    If g_oSelectedModel Is Nothing Then
        '   We may be just starting up the application, in which case we do not
        '   want to try to enable any tests.
        Exit Sub
    End If
    '   Next, determine which tests are activated based on the enabled
    '   operations.
    szTests = ""
    szDownload = g_oTests.selectSingleNode("Download").Text
    szConfigure = g_oTests.selectSingleNode("Configure").Text
    szWireless = g_oTests.selectSingleNode("Wireless").Text
    If chkTestFunctions.value = vbChecked Then
        szTests = g_oTests.selectSingleNode("Test").Text
        If Len(szWireless) > 0 Then
            szTests = szTests + "," & szWireless
        End If
    End If
    If chkDownload.value = vbChecked And Len(szDownload) > 0 Then
        If Len(szTests) > 0 Then
            szTests = szTests + "," & szDownload
        Else
            szTests = szDownload
        End If
    End If
    If chkConfigureUnit.value = vbChecked And Len(szConfigure) > 0 Then
        If Len(szTests) > 0 Then
            szTests = szTests + "," & szConfigure
        Else
            szTests = szConfigure
        End If
    End If
    '   Turn on the activated tests
    iStart = 1
    While iStart < Len(szTests) + 1
        DoEvents
        iEnd = InStr(iStart, szTests, ",")
        If iEnd = 0 Then
            iEnd = Len(szTests) + 1
        End If
        
        ' Get the test number
        iTest = Val(Mid(szTests, iStart, iEnd - iStart))
        
        '   Activate the specific test
        abTests(iTest) = True
        iStart = iEnd + 1
    Wend
    
    '   Now go and get the tests that are to be run based on the selected Model
    '   Type.  Add only those tests from the Model Type that are also activated
    '   by the Operation switches.
    szTests = g_oSelectedModel.selectSingleNode("Tests").Text
    iStart = 1
    iRow = 1
    iAdd = 1
    
    Dim iPrevTest As Integer
    
    If True = g_bDebugOn Then
        ' strLogPathAndFile = strCOM4LogPath & strCOM4LogFName & "." & strCOM4LogFExt
        fhNumber = FreeFile
        szTestSeqFN = g_oSelectedModel.baseName & ".txt"
        ' If Not chkTrialRun.value = vbChecked Then
        ' If m_bDownloadFW = True Then
        If chkDownload.value = vbChecked Then
            szTestSeqFN = "Dnld_" & szTestSeqFN
        End If
        If chkConfigureUnit.value = vbUnchecked Then
            szTestSeqFN = "NoConfig_" & szTestSeqFN
        End If
        szTestSeqFN = ".\Test_Sequence\" & szTestSeqFN
        Open szTestSeqFN For Output Access Write As #fhNumber
    End If
    
    iPrevTest = 0
    While iStart < Len(szTests) + 1
        DoEvents
        iEnd = InStr(iStart, szTests, ",")
        If iEnd = 0 Then
            iEnd = Len(szTests) + 1
        End If
        ' Get the model specific test number
        iTest = Val(Mid(szTests, iStart, iEnd - iStart))
        
        If abTests(iTest) And iTest <> iPrevTest Then
            '   Calling RunTest with the second parameter set to False will not
            '   cause the test to be run but will, instead, cause the
            '   description for the test to be populated into the Test Results
            '   ListView window.
            If iTest = g_iNumOfTests And g_oAddTests.childNodes.length > 1 Then
                While iAdd <= g_oAddTests.childNodes.length
                    Set lstRow = lvwTestResults.ListItems(iRow)
                    lstRow.Text = "Test " & Str(iTest + iAdd - 1)
                    lstRow.Tag = Str(iTest + iAdd - 1)
                    RunTest lstRow, False, iAdd
                    iAdd = iAdd + 1
                    iRow = iRow + 1
                Wend
            Else
                Set lstRow = lvwTestResults.ListItems(iRow)
                lstRow.Text = "Test " & Str(iTest)
                lstRow.Tag = Str(iTest)
                RunTest lstRow, False
                
                If True = g_bDebugOn Then
                    Print #fhNumber, lstRow.Text & ": " & lstRow.SubItems(eITEM_DESCRIPTION)
                    
                End If
                iRow = iRow + 1
            End If
            iPrevTest = iTest
        End If
        iStart = iEnd + 1
    Wend
    
    If True = g_bDebugOn Then
        
        Close #fhNumber
        ' Kill szTestSeqFN
    End If
                    
    
    lvwTestResults.Refresh

End Sub

Private Sub UpdateResults(lstRow As ListItem, szTestData As String, eUpdate As etRESULTS)

    Dim Index As Integer
    Dim szMessage As String

    Index = Val(lstRow.Tag) - 1
    lstRow.SubItems(eITEM_RESULTS) = szTestData
    m_aszMessages(Index) = m_aszMessages(Index) + vbCrLf & szTestData
    szMessage = vbCrLf & "TEST " & lstRow.Tag & ": " & lstRow.SubItems(eITEM_DESCRIPTION)
    Select Case eUpdate
        Case eRESULT_PASSED
            lstRow.SubItems(eITEM_PASS_FAIL) = "Pass"
            lstRow.ListSubItems(eITEM_PASS_FAIL).ForeColor = vbDarkGreen
            m_aszMessages(Index) = m_aszMessages(Index) + szMessage & " (PASSED)"

        Case eRESULT_FAILED
            lstRow.SubItems(eITEM_PASS_FAIL) = "Fail"
            lstRow.ListSubItems(eITEM_PASS_FAIL).ForeColor = vbRed
            m_aszMessages(Index) = m_aszMessages(Index) + szMessage & " (FAILED) >" & szTestData

        Case eRESULT_TESTING
            lstRow.SubItems(eITEM_PASS_FAIL) = "Testing..."
            lstRow.ListSubItems(eITEM_PASS_FAIL).ForeColor = vbYellow

        Case eRESULT_NOT_TESTED
            lstRow.SubItems(eITEM_PASS_FAIL) = "N/T"
            lstRow.ListSubItems(eITEM_PASS_FAIL).ForeColor = vbLightGrey
            m_aszMessages(Index) = m_aszMessages(Index) + szMessage & " (NOT TESTED)"
    End Select
    DoEvents

End Sub

Private Sub FormClear()

    Dim i As Integer

    For i = 1 To g_iNumOfTests
        lvwTestResults.ListItems(i).SubItems(eITEM_RESULTS) = ""
        lvwTestResults.ListItems(i).SubItems(eITEM_PASS_FAIL) = ""
        lvwTestResults.ListItems(i).ListSubItems(eITEM_PASS_FAIL).ForeColor = vbBlack
        m_aszMessages(i - 1) = ""
    Next i

End Sub

Private Sub lvwTestResults_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

    Dim iTestNumber As Integer

    If Button = vbRightButton And Not lvwTestResults.SelectedItem Is Nothing Then
        iTestNumber = Val(lvwTestResults.SelectedItem.Tag)
        If iTestNumber > 0 And iTestNumber <= g_iNumOfTests Then
            '   Display the contents of the text messages stored up for the
            '   selected test.
            m_fDetailedResults.Caption = "TEST " & iTestNumber & " (Detailed Results)"
            m_fDetailedResults.txtMessages.Text = m_aszMessages(iTestNumber - 1)
            m_fDetailedResults.Show vbModal
        End If
    End If

End Sub

Private Sub mnuAbout_Click()

    m_fAbout.Show vbModal

End Sub

Private Sub mnuContents_Click()

    m_fContents.Show vbModal

End Sub

Private Sub mnuDAQ_Click()

    If m_fPassword.CheckPassword Then
        m_fDiagDAQ.Show vbModal
    End If

End Sub

Private Sub mnuExit_Click()
    
    g_bStopApplication = True

End Sub

Private Sub mnuGPS_Click()

    If m_fPassword.CheckPassword Then
        m_fDiagGPS.Show vbModal
    End If

End Sub

Private Sub mnuGSM_Click()

    If m_fPassword.CheckPassword Then
        m_fDiagGSM.Show vbModal
    End If

End Sub

Private Sub mnuIMEI_Click()
    Dim oNode As IXMLDOMNode
    
    ' There are two ways currently to track IMEI numbers--by model type and by test station
    ' Tracking by test station ID means that each test station has a pool of IMEI numbers
    ' and any device, regardless of the model type, simply gets the next unused IMEI number.
    ' Tracking by model means that each model type is assigned a range of IMEI numbers
    
    Set oNode = g_oSettings.selectSingleNode("TestSystem").childNodes(g_iTestStationID)
    Set oNode = oNode.selectSingleNode("IMEI")
    
    If m_fPassword.CheckPassword Then
        m_fUpdateSerialNumbers.UpdateType = etID_TYPE.eIMEI_NUM
        If oNode.selectSingleNode("Input_Method").Text = "BY_MODEL" Then
            ' Track IMEI BY_MODEL
            m_fUpdateSerialNumbers.Show vbModal
        
            If Not g_oSelectedModel Is Nothing Then
                Set oNode = g_oSelectedModel.selectSingleNode("IMEI").selectSingleNode("SNR")
                '   The IMEI numbers may have changed by the above launching of
                '   UpdateSerialNumbers.  Update the global variables just in case.
                g_lNextIMEINumber.SNR = CLng(Val(oNode.Attributes.getNamedItem("Next").Text))
                g_szIMEI = Format(g_lNextIMEINumber.SNR, IMEI_NUM_FORMAT)
                g_lEndingIMEINumber.SNR = CLng(Val(oNode.Attributes.getNamedItem("End").Text))
            End If
        Else
            ' Track IMEI BY_COMPUTER_NAME
            Set oNode = oNode.selectSingleNode("SNR")
            g_lNextIMEINumber.SNR = CLng(Val(oNode.Attributes.getNamedItem("Next").Text))
            g_szIMEI = Format(g_lNextIMEINumber.SNR, IMEI_NUM_FORMAT)
            g_lEndingIMEINumber.SNR = CLng(Val(oNode.Attributes.getNamedItem("End").Text))
            m_fUpdateSerialNumbers.EditSerialNumbers ("BY_COMPUTER_NAME")

        End If
    End If
    
End Sub

Private Sub mnuMicIn_Click()

    If m_fPassword.CheckPassword Then
        g_fMicIn.Show vbModal
    End If

End Sub

Private Sub mnuPrintLabel_Click()

    If m_fPassword.CheckPassword Then
        m_fDiagPrinter.Show vbModal
    End If

End Sub

Private Sub mnuSerialNumber_Click()

    If m_fPassword.CheckPassword Then
        m_fUpdateSerialNumbers.UpdateType = etID_TYPE.eSERIAL_NUM
        m_fUpdateSerialNumbers.Show vbModal
        If Not g_oSelectedModel Is Nothing Then
            '   The currently selected model may have had its serial numbers
            '   changed by the above launching of UpdateSerialNumbers.  Update
            '   the global variables and refresh the serial number text box
            '   just in case.
            g_lNextSerialNumber = CLng(Val(g_oSelectedModel.selectSingleNode("SerialNo").Attributes.getNamedItem("Next").Text))
            g_lEndingSerialNumber = CLng(Val(g_oSelectedModel.selectSingleNode("SerialNo").Attributes.getNamedItem("End").Text))
            txtSerialNumber.Text = Format(g_lNextSerialNumber, SERIAL_NUM_FORMAT)
        End If
    End If

End Sub

Private Sub mnuSetup_Click()

    If m_fPassword.CheckPassword Then
        m_fSetupForm.Show vbModal
        '   The serial numbers of the currently selected Model Type have just
        '   been change so this call will cause the Next Serial Number box to
        '   get refreshed.
        cbxModelTypes_Click
    End If

End Sub

Public Sub BarcodeRePrint()
    
    g_clsBarcode.PrintBarcode 1, 1
    g_clsBarcode.PrintBarcode 2, 1

End Sub

Public Sub InventoryWriteResults()

    Dim szTestDateAndTime As String
    Dim szTestVersion As String
    Dim szFirmwareVersion As String
    If Len(g_szAPPVER) > 2 Then
    szFirmwareVersion = Mid(g_szAPPVER, 1, Len(g_szAPPVER) - 2)
    Else
    szFirmwareVersion = ""
    End If
    szTestVersion = "Version " & App.Major & "." & App.Minor & "  Build: " & App.Revision
    szTestDateAndTime = Now
    
    Dim Entry As Variant
    ReDim Entry(0 To 8)
    
    Entry(0) = g_lNextSerialNumber
    Entry(1) = g_szIMSI
    Entry(2) = g_szIMEI
    Entry(3) = g_szCCID
    Entry(4) = cbxModelTypes.Text
    Entry(5) = szFirmwareVersion
    Entry(6) = szTestVersion
    Entry(7) = szTestDateAndTime
    Entry(8) = 1
    
    InventoryAddNewRemoveDupes Entry
    
End Sub

'   This function connects a COM port to the Power Supply and verifies that the
'   instrument is responding correctly.

Public Function InitCOM_PortToPowerSupply() As Boolean

    Dim iCOM_Port As Integer

    iCOM_Port = CInt(Val(Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PORT)))
#If DISABLE_EQUIP_CHECKS = 0 Then
    On Error GoTo ComPSError
    '   We are stuck in this loop until the user either connects to the Power
    '   Supply successfully or they press the Quit button
    InitCOM_PortToPowerSupply = False
    While Not InitCOM_PortToPowerSupply And Not g_bQuitProgram And g_ePS_COM_Status <> eSTAT_SHUTTING_DOWN
        DoEvents
        If _
        0 <> iCOM_Port And _
        "" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_BAUD) And _
        "" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PARITY) And _
        "" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_DATA_BITS) And _
        "" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_STOP_BITS) Then
            '   If the COM port to the Power Supply is already open then close it.
            If ComPortToPowerSupply.PortOpen = True Then
                ComPortToPowerSupply.PortOpen = False
            End If
            ' Initialize the COM port settings then open it.
            ComPortToPowerSupply.CommPort = iCOM_Port
            ComPortToPowerSupply.RThreshold = 1
            ComPortToPowerSupply.SThreshold = 1
            ComPortToPowerSupply.Handshaking = 1
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
            ComPortToPowerSupply.Settings = _
                Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_BAUD) & "," & _
                Left(Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PARITY), 1) & "," & _
                Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_DATA_BITS) & "," & _
                Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_STOP_BITS)
            ComPortToPowerSupply.PortOpen = True
            '   Attempt to communicate with the Power Supply and verify correct response.
            SerialComs.TransmitToPS ePS_VERIFY
            While g_ePS_COM_Status = eSTAT_RUNNING
                DoEvents
            Wend
            If g_ePS_COM_Status = eSTAT_SUCCESS Then
                InitCOM_PortToPowerSupply = True
            Else
                SerialComs.TransmitToPS ePS_CLEAR_ERROR
                While g_ePS_COM_Status = eSTAT_RUNNING
                    DoEvents
                Wend
            End If
        End If
        GoTo SkipPSError
#Else
    Dim msgBoxMessage As String
    Dim msgBoxResponse As Integer
    On Error GoTo ComPSError
    ' Loop until the operator:
    '   powers up the power supply and, if DISABLE_EQUIP_CHECKS =  1, presses retry,
    '   presses ignore and enters the system administrator password, or
    '   presses abort
    InitCOM_PortToPowerSupply = False
    While Not InitCOM_PortToPowerSupply And Not g_bQuitProgram And g_ePS_COM_Status <> eSTAT_SHUTTING_DOWN
        DoEvents
        If _
        "0" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PORT) And _
        "" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_BAUD) And _
        "" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PARITY) And _
        "" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_DATA_BITS) And _
        "" <> Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_STOP_BITS) Then
            '   If the COM port to the Power Supply is already open then close it.
               If ComPortToPowerSupply.PortOpen = True Then
                   ComPortToPowerSupply.PortOpen = False
               End If
               ' Initialize the COM port settings then open it.
               ComPortToPowerSupply.CommPort = CInt(Val(Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PORT)))
               ComPortToPowerSupply.RThreshold = 1
               ComPortToPowerSupply.SThreshold = 1
               ComPortToPowerSupply.Handshaking = 1
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
               ComPortToPowerSupply.Settings = _
                    Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_BAUD) & "," & _
                    Left(Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PARITY), 1) & "," & _
                    Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_DATA_BITS) & "," & _
                    Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_STOP_BITS)
            If 0 = Err.number Then
              ComPortToPowerSupply.PortOpen = True
              '   Attempt to communicate with the Power Supply and verify correct response.
              SerialComs.TransmitToPS ePS_VERIFY
              While g_ePS_COM_Status = eSTAT_RUNNING
                  DoEvents
              Wend
              If g_ePS_COM_Status = eSTAT_SUCCESS Then
                  InitCOM_PortToPowerSupply = True
                  GoTo SkipPSError
              Else
                  SerialComs.TransmitToPS ePS_CLEAR_ERROR
                  While g_ePS_COM_Status = eSTAT_RUNNING
                      DoEvents
                  Wend
              End If
            End If
        End If

        If Not InitCOM_PortToPowerSupply Then
            If Err.number <> 0 Then
                msgBoxMessage = "COM port " & iCOM_Port & " error!" & vbCrLf & vbCrLf & Err.Description
            Else
                msgBoxMessage = "Failed to connect to the Power Supply!" & vbCrLf & vbCrLf & _
                   "Ensure that the Power Supply's serial cable is" & vbCrLf & _
                   "connected to COM" & iCOM_Port & " on the tester PC and that it" & vbCrLf & _
                   "is turned on."
            End If
            If True = ALLOW_NON_STANDARD_PS Then
                msgBoxResponse = MsgBox(msgBoxMessage, vbAbortRetryIgnore + vbCritical)
            Else
                msgBoxResponse = MsgBox(msgBoxMessage, vbRetryCancel + vbCritical)
            End If
            If vbIgnore = msgBoxResponse Then
                If True = frmPassword.CheckPassword() Then
                    ' System administrator override
                    Err.Clear   ' Clear Err object fields
                    InitCOM_PortToPowerSupply = False
                    g_ePS_COM_Status = eSTAT_NOT_CONNECTED
                    g_ePS_Control = etPS_CTL.ePS_CTL_OFF
                    Exit Function
                End If
            End If
            If vbAbort = msgBoxResponse Or vbCancel = msgBoxResponse Then
                Err.Clear   ' Clear Err object fields
                g_bQuitProgram = True
            End If
        End If
        If 0 = Err.number Then
            GoTo SkipPS_COM_PortError
        End If
#End If
ComPSError:
        Select Case Err.number
            Case comPortAlreadyOpen
                MsgBox "COM port " & iCOM_Port & " is already open!", vbCritical
    
            Case comPortInvalid
                MsgBox "COM port " & iCOM_Port & " is invalid!" & vbCrLf & vbCrLf & "Try another port number.", vbCritical
    
            Case Else
                MsgBox "Error " & Err.number & " while attempting to open COM port " & iCOM_Port & " to Power Supply!", vbCritical
    
        End Select
        GoTo SkipPS_COM_PortError
SkipPSError:
        If Not InitCOM_PortToPowerSupply Then
            MsgBox "Failed to connect to the Power Supply!" & vbCrLf & vbCrLf & _
                   "Ensure that the Power Supply's serial cable is" & vbCrLf & _
                   "connected to COM" & iCOM_Port & " on the tester PC and that it" & vbCrLf & _
                   "is turned on.", vbCritical
            g_bQuitProgram = True
        End If

SkipPS_COM_PortError:
    Wend

End Function

' Configures the RS-232 COM port on the tester PC used to communicate with the device
Public Function InitCOM_PortToDevice(Optional bWarnOnError As Boolean = True) As Boolean

    Dim iCOM_Port As Integer

    On Error GoTo ConfigCOM_PortError
    
    iCOM_Port = CInt(Val(Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_PORT)))
    
    While Not InitCOM_PortToDevice And Not g_bQuitProgram And g_eElite_Status <> eSTAT_SHUTTING_DOWN
        '   We are stuck in this loop until the user either connects to the device
        '   successfully or they press the Quit button
        If _
        0 <> iCOM_Port And _
        "" <> Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_BAUD) And _
        "" <> Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_PARITY) And _
        "" <> Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_DATA_BITS) And _
        "" <> Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_STOP_BITS) Then

            ' Initialize the COM port settings then open it.
            ComPortToElite.CommPort = iCOM_Port
            ComPortToElite.RThreshold = 1
            ComPortToElite.SThreshold = 1
            ComPortToElite.Handshaking = comNone
            ComPortToElite.ParityReplace = Chr$(255)
            ComPortToElite.CTSTimeout = 0
            ComPortToElite.DSRTimeout = 0
            ComPortToElite.DTREnable = True
            ComPortToElite.RTSEnable = True
            ComPortToElite.NullDiscard = True
            ComPortToElite.InBufferSize = 8192
            ComPortToElite.InputLen = 1
            ComPortToElite.OutBufferSize = 512
            ComPortToElite.InBufferCount = 0
            ComPortToElite.OutBufferCount = 0
            ComPortToElite.Settings = _
                 Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_BAUD) & "," & _
                 Left(Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_PARITY), 1) & "," & _
                 Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_DATA_BITS) & "," & _
                 Config.GetComSettings(etCOM_DEV_DUT, etCOM_SETTING_STOP_BITS)
            
            InitCOM_PortToDevice = True
        Else
            InitCOM_PortToDevice = False
        End If
        GoTo SkipEliteCOMError

ConfigCOM_PortError:
        Select Case Err.number
            Case comPortInvalid
                MsgBox "COM port " & iCOM_Port & " is invalid!" & vbCrLf & vbCrLf & "Try another port number.", vbCritical
    
            Case Else
                MsgBox "Error " & Err.number & " while attempting to open COM port " & iCOM_Port, vbCritical
        End Select

SkipEliteCOMError:
        If Not InitCOM_PortToDevice Then
            MsgBox "Failed to connect to the device!" & vbCrLf & vbCrLf & _
                   "Ensure that the device's serial cable is" & vbCrLf & _
                   "connected to COM" & iCOM_Port & " on the tester PC.", vbCritical
            g_bQuitProgram = True
        End If
    Wend

End Function
'   This function connects a COM port to the Call Box and verifies that the
'   instrument is responding correctly.
Public Function InitCOM_PortToCallBox() As Boolean
    Dim iCOM_Port As Integer

    Dim CallBoxCfgCmd As etCB_STATES
    Dim numCallBoxCfgCmds As etCB_STATES
    Dim lastCallBoxCfgCmd As etCB_STATES
    
    iCOM_Port = CInt(Val(Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_PORT)))
    CallBoxCfgCmd = etCB_STATES.eCB_CONFIG_PRE_ATTEN
    lastCallBoxCfgCmd = etCB_STATES.eCB_CONFIG_GSM - 1
    numCallBoxCfgCmds = CallBoxCfgCmd - lastCallBoxCfgCmd
    

#If DISABLE_EQUIP_CHECKS = 0 Then
    On Error GoTo ComCBError
    '   We are stuck in this loop until the user either connects to the Power
    '   Supply successfully or they press the Quit button
    InitCOM_PortToCallBox = False
    While Not InitCOM_PortToCallBox And Not g_bQuitProgram And g_eCB_COM_Status <> eSTAT_SHUTTING_DOWN
        DoEvents
        If _
        0 <> iCOM_Port And _
        "" <> Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_BAUD) And _
        "" <> Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_PARITY) And _
        "" <> Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_DATA_BITS) And _
        "" <> Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_STOP_BITS) Then
            '   If the COM port to the Call Box is already open then close it.
            If ComPortToCallBox.PortOpen = True Then
                ComPortToCallBox.PortOpen = False
            End If
            ' Initialize the COM port settings then open it.
            ComPortToCallBox.CommPort = g_iComPortCallBox
            ComPortToCallBox.RThreshold = 1
            ComPortToCallBox.SThreshold = 1
            ComPortToCallBox.Handshaking = 1
            ComPortToCallBox.ParityReplace = Chr$(255)
            ComPortToCallBox.CTSTimeout = 0
            ComPortToCallBox.DSRTimeout = 0
            ComPortToCallBox.DTREnable = True
            ComPortToCallBox.NullDiscard = True
            ComPortToCallBox.InBufferSize = 1024
            ComPortToCallBox.InputLen = 1
            ComPortToCallBox.OutBufferSize = 512
            ComPortToCallBox.InBufferCount = 0
            ComPortToCallBox.OutBufferCount = 0
            ComPortToCallBox.Settings = _
                 Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_BAUD) & "," & _
                 Left(Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_PARITY), 1) & "," & _
                 Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_DATA_BITS) & "," & _
                 Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_STOP_BITS)
            ComPortToCallBox.PortOpen = True
            '   Attempt to communicate with the Power Supply and verify correct response.
            SerialComs.TransmitToCB etCB_STATES.eCB_RESET
            While g_eCB_COM_Status = eSTAT_RUNNING
                DoEvents
            Wend
            startup.Delay 1#
            SerialComs.TransmitToCB etCB_STATES.eCB_VERIFY
            While g_eCB_COM_Status = eSTAT_RUNNING
                DoEvents
            Wend
            
            Do Until 0 = callBoxCfgCmds Or eSTAT_SUCCESS <> g_eCB_COM_Status
                SerialComs.TransmitToCB callBoxCfgCmds
                While g_eCB_COM_Status = eSTAT_RUNNING
                    DoEvents
                Wend
                callBoxCfgCmds = callBoxCfgCmds - 1
            Loop

            If eSTAT_SUCCESS = g_eCB_COM_Status Then
                InitCOM_PortToCallBox = True
            Else
                SerialComs.TransmitToCB etCB_STATES.eCB_CLEAR_ERROR
                While g_eCB_COM_Status = eSTAT_RUNNING
                    DoEvents
                Wend
            End If
        End If
        GoTo SkipCBError
#Else
    Dim msgBoxMessage As String
    Dim msgBoxResponse As Integer
    On Error GoTo ComCBError
    ' Loop until the operator:
    '   powers up the power supply and, if DISABLE_EQUIP_CHECKS =  1, presses retry,
    '   presses ignore and enters the system administrator password, or
    '   presses abort
    InitCOM_PortToCallBox = False
    While Not InitCOM_PortToCallBox And Not g_bQuitProgram And g_eCB_COM_Status <> eSTAT_SHUTTING_DOWN
        DoEvents
        If _
        0 <> iCOM_Port And _
        "" <> Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_BAUD) And _
        "" <> Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_PARITY) And _
        "" <> Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_DATA_BITS) And _
        "" <> Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_STOP_BITS) Then
            '   If the COM port to the Power Supply is already open then close it.
               If True = ComPortToCallBox.PortOpen Then
                   ComPortToCallBox.PortOpen = False
               End If
               ' Initialize the COM port settings then open it.
               ComPortToCallBox.CommPort = iCOM_Port
               ComPortToCallBox.RThreshold = 1
               ComPortToCallBox.SThreshold = 1
               ComPortToCallBox.Handshaking = 1
               ComPortToCallBox.ParityReplace = Chr$(255)
               ComPortToCallBox.CTSTimeout = 0
               ComPortToCallBox.DSRTimeout = 0
               ComPortToCallBox.DTREnable = True
               ComPortToCallBox.NullDiscard = True
               ComPortToCallBox.InBufferSize = 1024
               ComPortToCallBox.InputLen = 1
               ComPortToCallBox.OutBufferSize = 512
               ComPortToCallBox.InBufferCount = 0
               ComPortToCallBox.OutBufferCount = 0
               ComPortToCallBox.Handshaking = comRTSXOnXOff
            ComPortToCallBox.Settings = _
                 Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_BAUD) & "," & _
                 Left(Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_PARITY), 1) & "," & _
                 Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_DATA_BITS) & "," & _
                 Config.GetComSettings(etCOM_DEV_GSM, etCOM_SETTING_STOP_BITS)
            
            If 0 = Err.number Then
                Dim callBoxCfgCmds As etCB_STATES
                callBoxCfgCmds = etCB_STATES.eCB_CONFIG_PRE_ATTEN - etCB_STATES.eCB_CONFIG_GSM
                
                ComPortToCallBox.PortOpen = True
                '   Attempt to communicate with the Power Supply and verify correct response.
                SerialComs.TransmitToCB etCB_STATES.eCB_RESET
                While g_eCB_COM_Status = eSTAT_RUNNING
                    DoEvents
                Wend
                ComPortToCallBox.PortOpen = False
                ComPortToCallBox.PortOpen = True
                startup.Delay 3#
                SerialComs.TransmitToCB eCB_VERIFY
                While g_eCB_COM_Status = eSTAT_RUNNING
                    DoEvents
                Wend
                
                Do Until lastCallBoxCfgCmd = CallBoxCfgCmd Or eSTAT_SUCCESS <> g_eCB_COM_Status
                    SerialComs.TransmitToCB CallBoxCfgCmd
                    While g_eCB_COM_Status = eSTAT_RUNNING
                        DoEvents
                    Wend
                    CallBoxCfgCmd = CallBoxCfgCmd - 1
                Loop
    
                If eSTAT_SUCCESS = g_eCB_COM_Status Then
                    InitCOM_PortToCallBox = True
                    'Set call box display back to local
                    ' < 10s delay will work but will produce status error
                    ' TransmitToCB will wait 2 sec. so set delay to 12 and
                    ' should come back right away success but will still have
                    ' 2s to parse response
                    startup.Delay 12#
                    SerialComs.TransmitToCB eCB_END_REMOTE
                    While g_eCB_COM_Status = eSTAT_RUNNING
                        DoEvents
                    Wend
                    If eSTAT_SUCCESS = g_eCB_COM_Status Then
                        GoTo SkipCBError
                    Else
                        SerialComs.TransmitToCB eCB_CLEAR_ERROR
                        While g_eCB_COM_Status = eSTAT_RUNNING
                            DoEvents
                        Wend
                    End If
                Else
                    SerialComs.TransmitToCB eCB_CLEAR_ERROR
                    While g_eCB_COM_Status = eSTAT_RUNNING
                        DoEvents
                    Wend
                End If
            End If
        End If

        If Not InitCOM_PortToCallBox Then
            If Err.number <> 0 Then
                msgBoxMessage = "COM port " & iCOM_Port & " error!" & vbCrLf & vbCrLf & Err.Description
            Else
                msgBoxMessage = "Failed to connect to the Call Box!" & vbCrLf & vbCrLf & _
                   "Ensure that the Call Box's serial cable is" & vbCrLf & _
                   "connected to COM" & iCOM_Port & " on the tester PC and that it" & vbCrLf & _
                   "is turned on."
            End If
            If True = ALLOW_NON_STANDARD_CB Then
                msgBoxResponse = MsgBox(msgBoxMessage, vbAbortRetryIgnore + vbCritical)
            Else
                msgBoxResponse = MsgBox(msgBoxMessage, vbRetryCancel + vbCritical)
            End If
            If vbIgnore = msgBoxResponse Then
                If True = frmPassword.CheckPassword() Then
                    ' System administrator override
                    Err.Clear   ' Clear Err object fields
                    InitCOM_PortToCallBox = False
                    g_eCB_COM_Status = eSTAT_NOT_CONNECTED
                    g_eCB_Control = etCB_CTL.eCB_CTL_OFF
                    ComPortToCallBox.PortOpen = False
                    Exit Function
                End If
            End If
            If vbAbort = msgBoxResponse Or vbCancel = msgBoxResponse Then
                Err.Clear   ' Clear Err object fields
                g_bQuitProgram = True
            End If
        End If
        If 0 = Err.number Then
            GoTo SkipCB_COM_PortError
        End If
#End If
ComCBError:
        Select Case Err.number
            Case comPortAlreadyOpen
                MsgBox "COM port " & iCOM_Port & " is already open!", vbCritical
    
            Case comPortInvalid
                MsgBox "COM port " & iCOM_Port & " is invalid!" & vbCrLf & vbCrLf & "Try another port number.", vbCritical
    
            Case Else
                MsgBox "Error " & Err.number & " while attempting to open COM port " & iCOM_Port & " to Power Supply!", vbCritical
    
        End Select
        GoTo SkipCB_COM_PortError
SkipCBError:
        If Not InitCOM_PortToCallBox Then
            MsgBox "Failed to connect to the Power Supply!" & vbCrLf & vbCrLf & _
                   "Ensure that the Power Supply's serial cable is" & vbCrLf & _
                   "connected to COM" & iCOM_Port & " on the tester PC and that it" & vbCrLf & _
                   "is turned on.", vbCritical
            g_bQuitProgram = True
        End If

SkipCB_COM_PortError:
        Err.number = 0
    Wend
    
    If True = ComPortToCallBox.PortOpen Then
        ComPortToCallBox.PortOpen = False
    End If
End Function
' Initialize the inerface to the device used for testing the device
Public Function ConfigureDeviceForTesting() As Boolean

    ' Because this command turns off unsolicited messages, it is typical that
    ' the tester receives unsolicited messages before the device can process
    ' the command. Run with minimal timeout and then reissue the command
    ' to verify communications with the device
  
    ConfigureDeviceForTesting = SerialComs.SendATCommand(eelite_verboseoff, 1)
    ' If the command did process then calling it a second time
    ' will either produce the expected response or fail, if the command actually
    ' does fail
    ConfigureDeviceForTesting = SerialComs.SendATCommand(eelite_verboseoff)
End Function
Public Function LogMessage(szMessage As String)

    If Not g_bDiagnosticActive And m_iCurrentTest >= 0 And m_iCurrentTest <= g_iNumOfTests Then
        m_aszMessages(m_iCurrentTest) = m_aszMessages(m_iCurrentTest) + vbCrLf & szMessage
    End If

End Function

Private Sub tmrAutoTest_Timer()

    Dim bTesterClosed As Boolean

    If g_bDiagnosticActive Then
        '   By setting m_eTestingState to eTEST_DIAGNOSTIC here means that a
        '   false auto-test start won't occur when the diagnostic is complete.
        m_eTestingState = eTEST_DIAGNOSTIC
        '   Don't do an autotest check while the wireless diagnostic is running
        Exit Sub
    End If
    
    bTesterClosed = g_fIO.IsTesterClosed
    
    ' Debounce Fixture Down and Fixture Gate Closed switches
    If m_bTesterClosed = bTesterClosed And 0 = m_iAutoTestDebounce Then
        ' No state change finished debouncing
        m_iAutoTestDebounce = AUTO_TEST_DEBOUNCE
        Exit Sub
    ElseIf m_bTesterClosed <> bTesterClosed And 0 = m_iAutoTestDebounce Then
        ' State change and finished debouncing. Continue processing
        m_iAutoTestDebounce = AUTO_TEST_DEBOUNCE
        m_bTesterClosed = bTesterClosed
    ElseIf m_bTesterClosed = bTesterClosed And AUTO_TEST_DEBOUNCE <> m_iAutoTestDebounce Then
        ' No state change and debouncing (detected a bounce)
        m_iAutoTestDebounce = AUTO_TEST_DEBOUNCE
        Exit Sub
    ElseIf m_bTesterClosed <> bTesterClosed Then
        ' State change and debouncing
        ' State change and need to start debounce checking
        m_iAutoTestDebounce = m_iAutoTestDebounce - 1
        Exit Sub
    End If
    
    ' Will only be here in one of two cases:
    '  No state change and not debouncing
    '  State change and finished debouncing

    If True = m_bTesterClosed Then
        If m_eTestingState = eTESTER_OPEN Or m_eTestingState = eTEST_APP_STARTED Or _
           m_eTestingState = eTEST_DIAGNOSTIC Then
            '   The tester has just been closed with a DUT inside of it or
            '   the application has just been started and there's a DUT
            '   already inside the tester.
            lblPassFail.Caption = "Ready: Test Fixture Closed"
            lblPassFail.BackColor = &HFF8080
            If g_bAutoTest And m_eTestingState = eTESTER_OPEN Then
                '   Since the Auto Test option has been selected
                '   automatically start testing on the inserted DUT.
                startup.Delay 1#
                cmdStartTest_Click
            Else
                '   Allow the user to start the test manually (the default
                '   state when the application first starts up, regardless
                '   of the auto-test setting.
                cmdStartTest.Enabled = True
            End If
        End If
    Else
        lblPassFail.Caption = "***Not Ready: Tester Fixture Open***"
        lblPassFail.BackColor = vbDarkGrey
        cmdStartTest.Enabled = False
        m_eTestingState = eTESTER_OPEN
    End If

End Sub

Private Sub tmrPSCom_Timer()

    '   Turn the timer off
    g_fMainTester.tmrPSCom.Enabled = False
    g_ePS_COM_Status = eSTAT_TIMEOUT

End Sub

Private Sub tmrCBCom_Timer()

    '   Turn the timer off
    g_fMainTester.tmrCBCom.Enabled = False
    g_eCB_COM_Status = eSTAT_TIMEOUT

End Sub


Private Sub tmrWaitPowerUp_Timer()

    g_bOpenATRunning = True
    tmrWaitPowerUp.Enabled = False

End Sub

Private Sub tmrWaitWireless_Timer()

    If Not m_bSuspendWirelessTimeout Then
        If m_iWirelessTimeout > CURRENT_TIMEOUT Then
            m_bCurrentTimeout = True
        End If
        If m_iWirelessTimeout > VOLTS_TIMEOUT Then
            m_bVoltsTimeout = True
        End If
        If m_iWirelessTimeout > GPRS_TIMEOUT Then
            m_bGPRSTimeout = True
        End If
        If m_iWirelessTimeout > g_iGPSTimeout Then
            m_bGPSTimeout = True
        Else
            m_iWirelessTimeout = m_iWirelessTimeout + 1
        End If
    End If

End Sub

Private Sub tmrEliteCom_Timer()

    '   Turn the timer off
    g_fMainTester.tmrEliteCom.Enabled = False
    If g_eElite_Status = eSTAT_RUNNING Then
        '   Flag a timeout, unless we're downloading
        g_eElite_Status = eSTAT_TIMEOUT
    End If

End Sub

Public Sub InventoryAddNewRemoveDupes(EntryArray As Variant)

    '
    '   We have a new entry to add to the log file, but it may already have
    '   one or more entries that have the same serial number as the one to
    '   be entered.  Remove all existing duplicate entries and add the new
    '   one.
    '
    Dim DupCount As Integer
    Dim i As Integer
    Dim iResult As Integer
    Dim NewKey As String
    
    On Error GoTo NewEntryIsADuplicate
    
AddNewEntryToCollection:
    
    NewKey = "K" & EntryArray(0)
    
    TestLogCollection.Add EntryArray, NewKey
    
    'iResult = MsgBox("Added new entry to collection for SN: " & EntryArray(0), vbOKOnly)
    '
    '   Call the routine to add the new entry to the text log file (actually update
    '   the WHOLE text file)
    '
    InventoryRegenerate
    
    Exit Sub
    
NewEntryIsADuplicate:
    
    'iResult = MsgBox("The new entry, SN: " & EntryArray(0) & " has a duplicate in the data file.  Deleting old entries.", vbOKOnly)
    Dim DuplicateEntry As Variant
    
    DuplicateEntry = TestLogCollection.Item(NewKey)
    DupCount = DuplicateEntry(8)
    
    For i = 0 To DupCount - 1
    
        TestLogCollection.Remove (NewKey)
        NewKey = "K" & NewKey
    Next
    NewKey = "K" & EntryArray(0)
    
    Resume AddNewEntryToCollection
    
End Sub


Public Sub InventoryRegenerate()

    Dim NewLogFileEntry As Variant
    Dim i As Integer
    Dim StringToWrite As String
    Dim iResult As Integer
    Dim fhNumber As Integer

    fhNumber = FreeFile
    
    On Error GoTo FileWriteError:
    
    Open strNewTestLog For Output As #fhNumber
    
    For i = 1 To TestLogCollection.Count
    
        NewLogFileEntry = TestLogCollection(i)
        '
        '   Format the text file correctly as comma delimited fields in qoutes
        '   (except for the first field -- serial number, which is not quoted)
        '
        StringToWrite = NewLogFileEntry(0)
        StringToWrite = StringToWrite & "," & Chr$(34) & NewLogFileEntry(1) & Chr$(34)
        StringToWrite = StringToWrite & "," & Chr$(34) & NewLogFileEntry(2) & Chr$(34)
        StringToWrite = StringToWrite & "," & Chr$(34) & NewLogFileEntry(3) & Chr$(34)
        StringToWrite = StringToWrite & "," & Chr$(34) & NewLogFileEntry(4) & Chr$(34)
        StringToWrite = StringToWrite & "," & Chr$(34) & NewLogFileEntry(5) & Chr$(34)
        StringToWrite = StringToWrite & "," & Chr$(34) & NewLogFileEntry(6) & Chr$(34)
        StringToWrite = StringToWrite & "," & Chr$(34) & NewLogFileEntry(7) & Chr$(34)
        StringToWrite = StringToWrite & "," & Chr$(34) & NewLogFileEntry(8) & Chr$(34)
        
        Print #fhNumber, StringToWrite
        
    Next i
    
    Close #fhNumber
    
    FileCopy strNewTestLog, g_szTestResultsFile
    
    Kill strNewTestLog
    
    Exit Sub
    
FileWriteError:
    iResult = MsgBox("Error writing to test log data file !!!", vbCritical)
    Resume Next
    
End Sub

Public Function ExecCmd(cmdline$)
   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   Dim ret&

   ' Initialize the STARTUPINFO structure:
   start.cb = Len(start)

   ' Start the shelled application:
   ret& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

   ' Wait for the shelled application to finish:
      ret& = WaitForSingleObject(proc.hProcess, INFINITE)
      Call GetExitCodeProcess(proc.hProcess, ret&)
      Call CloseHandle(proc.hThread)
      Call CloseHandle(proc.hProcess)
      ExecCmd = ret&
End Function

' Power device on/off and open/close the COM port
Public Function PowerAndCOM(lstRow As ListItem, bPwrOnComOpen As Boolean)
    Dim iWaitForInit As Integer

    PowerAndCOM = False
    ' TODO: Analyze why open COM port then apply power fails
    ' Ideally we would open the COM port then power on but
    ' in VB6 this causes the COM port to not respond
    SerialComs.PSOutput = bPwrOnComOpen
    
    If True = bPwrOnComOpen Then
        ' No delay should be needed because done by PSOutput but
        ' fails without at least a 500ms delay here
        For iWaitForInit = POST_WAIT To 1 Step -1
            UpdateResults lstRow, "Waiting for device to initialize..." & iWaitForInit, eRESULT_TESTING
            If m_eTestingState = eTESTER_OPEN Then
                Exit For
            Else
                startup.Delay 1#
            End If
        Next iWaitForInit
    End If
    PowerAndCOM = SerialComs.OpenEliteComPort(bPwrOnComOpen, False)
End Function


Private Function ConfigurePowerOn(lstRow As ListItem, bRunTest As Boolean) As etRESULTS

    ConfigurePowerOn = eRESULT_NOT_TESTED
    If Not bRunTest Then
        lstRow.ListSubItems(eITEM_DESCRIPTION).Text = "Turn Power On"
    ElseIf m_bDeviceFailure Then
        UpdateResults lstRow, "Skipped due to device failure", eRESULT_NOT_TESTED
        ConfigurePowerOn = eRESULT_NOT_TESTED
    ElseIf m_eTestingState = eTESTER_OPEN Then
        UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
    Else
        Dim iWaitForInit As Integer
        Dim oNode As IXMLDOMNode
        ConfigurePowerOn = eRESULT_TESTING
    
        If etPS_CTL.ePS_CTL_RS232 <> g_ePS_Control Then
            '   No communications to PS
            If vbOK = MsgBox("Apply power to DUT", vbOKCancel + vbCritical) Then
                UpdateResults lstRow, "Applying Power to the device...", eRESULT_NOT_TESTED
                ConfigurePowerOn = eRESULT_NOT_TESTED
                g_ePS_Control = etPS_CTL.ePS_CTL_ON
            Else
                UpdateResults lstRow, "Device not powered", eRESULT_FAILED
                ConfigurePowerOn = eRESULT_FAILED
                m_bDeviceFailure = True
            End If
        Else
            UpdateResults lstRow, "Applying Power to the device...", eRESULT_TESTING
            SerialComs.PSOutput = True
            startup.Delay 1#
        End If

        If eRESULT_FAILED <> ConfigurePowerOn Then
            Dim iPostWait
            Dim dDUTCurrentDraw As Double
            Dim iIdx As Integer
            Dim dMaxInRushCurrent As Double
            Dim dAvgCurrent As Double
            Dim dMaxCurrent As Double
            Dim dInRushDuration As Double
            
            Set oNode = g_oSelectedModel.selectSingleNode("Power").selectSingleNode("Nominal_Current")
            dAvgCurrent = CDbl(oNode.Attributes.getNamedItem("Avg").Text)
'            dMinCurrent = CDbl(oAttributes.getNamedItem("Min").Text)
            dMaxCurrent = CDbl(oNode.Attributes.getNamedItem("Max").Text)
            
            Set oNode = g_oSelectedModel.selectSingleNode("Power").selectSingleNode("Inrush_Current")
            dMaxInRushCurrent = CDbl(oNode.Attributes.getNamedItem("Max").Text)
            dInRushDuration = CDbl(oNode.Attributes.getNamedItem("Wait").Text)
            
            ' Check inrush current and shut off if too high
            ' get current reading from power suply
            iIdx = dInRushDuration
            Do
                dDUTCurrentDraw = SerialComs.PowerSupply1_Current
                If dDUTCurrentDraw > dMaxInRushCurrent Then
'                    SerialComs.PSOutput = False
                    LogMessage ("Exceeded Maximum Inrush Current")
                    TransmitToPS ePS_CLEAR_ERROR
                    While g_ePS_COM_Status = eSTAT_RUNNING
                        DoEvents
                    Wend
'                    UpdateResults lstRow, "Exceeded Maximum Inrush Current", eRESULT_FAILED
'                    ConfigurePowerOn = eRESULT_FAILED
'                    m_bDeviceFailure = True
'                    Exit Function
                End If
                UpdateResults lstRow, "Inrush current measurement " & iIdx & ": " & Format(dDUTCurrentDraw * 1000, "0.00" & " mA"), eRESULT_TESTING
                iIdx = iIdx - 1
                startup.Delay 1#
            Loop Until 0 = iIdx Or (dDUTCurrentDraw < dAvgCurrent And dDUTCurrentDraw > 0)
            
            If dDUTCurrentDraw > dMaxInRushCurrent Then
                SerialComs.PSOutput = False
                LogMessage ("Inrush current: " & Format(dDUTCurrentDraw * 1000, "0.00" & " mA"))
                UpdateResults lstRow, "Exceeded Maximum Inrush Current Duration", eRESULT_FAILED
                ConfigurePowerOn = eRESULT_FAILED
                m_bDeviceFailure = True
                Exit Function
            End If
            
            PowerSupply1_Amp_Limit "NOMINAL"

            If True = m_bDownloadFW Then
                iPostWait = POST_WAIT_FW_DNLD
            Else
                iPostWait = POST_WAIT
            End If
            For iWaitForInit = iPostWait To 1 Step -1
                UpdateResults lstRow, "Waiting for device to initialize..." & iWaitForInit, eRESULT_TESTING
                startup.Delay 1#
                If m_eTestingState = eTESTER_OPEN Then
                    UpdateResults lstRow, "JIG OPENED: Abort test", eRESULT_FAILED
                    ConfigurePowerOn = eRESULT_FAILED
                    m_bDeviceFailure = True
                    Exit For
                End If
            Next iWaitForInit
        End If
        
        If eRESULT_FAILED <> ConfigurePowerOn Then
            UpdateResults lstRow, "Applied power to the device", eRESULT_PASSED
            ConfigurePowerOn = eRESULT_PASSED
        End If
    End If
End Function

Private Function ConfigurePowerOff(lstRow As ListItem, bRunTest As Boolean) As etRESULTS
    Dim comPortClosed As Boolean
    
    m_bApplicationRunning = False
    If Not bRunTest Then
        lstRow.ListSubItems(eITEM_DESCRIPTION).Text = "Turn Power Off"
    Else
        comPortClosed = OpenEliteComPort(False, False)
#If DISABLE_EQUIP_CHECKS = 0 Then
        SerialComs.PSOutput = False
        UpdateResults lstRow, "Device powered down", eRESULT_PASSED
        ConfigurePowerOff = eRESULT_PASSED
#Else
        If etPS_CTL.ePS_CTL_RS232 <> g_ePS_Control Then
            If etPS_CTL.ePS_CTL_ON = g_ePS_Control Then
                '   No communications to PS
                MsgBox "Remove power from the DUT", vbOKOnly + vbCritical
            End If
            UpdateResults lstRow, "Device powered up", eRESULT_NOT_TESTED
            ConfigurePowerOff = eRESULT_NOT_TESTED
            g_ePS_Control = etPS_CTL.ePS_CTL_OFF
        Else
            SerialComs.PSOutput = False
            UpdateResults lstRow, "Device powered down", eRESULT_PASSED
            ConfigurePowerOff = eRESULT_PASSED
        End If
#End If
    End If
End Function

Private Function util_GetImeiCheckDigit(ByVal number As String) As String
    '   1. Starting from the right, double every other digit (e.g., 7 ? 14).
    '   2. Sum the digits (e.g., 14 ? 1 + 4).
    '   3. Check if the sum is divisible by 10.
    
    Dim numDigits, digit, sum, i As Integer
    Dim alternate As Boolean
    Dim checkDigit As Byte
    Dim debugOut As String
    
    util_GetImeiCheckDigit = ""
    
'    debugOut = ""
    
    sum = 0
    
    alternate = True
        
    numDigits = Len(number)
    
    If 14 <> numDigits Or False = IsNumeric(number) Then
        Return
    End If
        
    For i = Len(number) - 1 To 0 Step -1
        digit = CByte(Right(number, 1))
        number = Left(number, i)
        If True = alternate Then
            digit = digit * 2
            If 9 < digit Then
                digit = digit Mod 10
'                debugOut = debugOut & " + (1 + " & digit & ")"
                digit = digit + 1
'            Else
'                debugOut = debugOut & " + " & digit
            End If
            alternate = False
        Else
            alternate = True
'            debugOut = debugOut & " + " & digit
        End If
        sum = sum + digit
    Next i
    
    ' Check digit is the digit that when added to the sum
    ' makes the sum evenly divisible by 10
    checkDigit = (10 - (sum Mod 10)) Mod 10
    
    util_GetImeiCheckDigit = CStr(checkDigit)
    
End Function
