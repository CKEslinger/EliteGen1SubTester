VERSION 5.00
Begin VB.Form frmDiagGPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPS Diagnostic"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame_LongLat 
      Caption         =   "Longitude & Latiutude"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5160
      TabIndex        =   22
      Top             =   720
      Width           =   6495
      Begin VB.Label lblLatitude 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--° --' --.--"""
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblLatitudeTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Latitude:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblLongitude 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--° --' --.--"""
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblLongitudeTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Longitude:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame_GPS_In 
      Caption         =   "Device Antenna Select"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2640
      TabIndex        =   8
      Top             =   720
      Width           =   2295
      Begin VB.OptionButton Option_AntennaIn 
         Caption         =   "External Antenna"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option_AntennaIn 
         Caption         =   "On-Board Antenna"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame_GPS_Out 
      Caption         =   "GPS Simulator Output Select"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2295
      Begin VB.OptionButton Option_AntennaOut 
         Caption         =   "Device"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option_AntennaOut 
         Caption         =   "Test Fixture Antenna"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Timer tmrGPS 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7920
      Top             =   5280
   End
   Begin VB.Frame Frame_SigStr 
      Caption         =   "Satellite SNR"
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   11535
      Begin VB.OptionButton Option_ScanSatMode 
         Caption         =   "Scan for Satellite Vehicles"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton Option_ScanSatMode 
         Caption         =   "Specify Satellite Vehicle ID"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox Text_ID 
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
         Height          =   375
         Left            =   960
         TabIndex        =   15
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl_SNRsTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Max SNR:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lbl_SNRs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   6360
         TabIndex        =   33
         Top             =   2640
         Width           =   5055
      End
      Begin VB.Label lbl_SNRs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   32
         Top             =   2640
         Width           =   5055
      End
      Begin VB.Label lbl_SNRs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   6360
         TabIndex        =   31
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label lbl_SNRs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   30
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label lbl_SNRsTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Avg SNR:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lbl_SNRsTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Last SNR:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lbl_SatIDsTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "SV ID:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lbl_SNRs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   6360
         TabIndex        =   19
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label lbl_SatIDs 
         Caption         =   "17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   21
         Top             =   1440
         Width           =   5055
      End
      Begin VB.Label lbl_SNRs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   18
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label lbl_SatIDs 
         Caption         =   "01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   20
         Top             =   1440
         Width           =   5055
      End
      Begin VB.Label lblIdTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "SV ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblSNRTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "SNR:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblSNR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Label lblRunningTime 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9120
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblNumOfSatsTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Num of Satellites"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label lblNumOfSatellites 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label lblOperationalState 
      Alignment       =   2  'Center
      Caption         =   "Test jig open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmDiagGPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ANT_IN_ON_BOARD As Integer = 0
Private Const ANT_IN_EXTERNAL As Integer = 1
Private Const GPS_OUT_TO_TESTER_ANT As Integer = 0
Private Const GPS_OUT_TO_DEVICE As Integer = 1
Private Const PWR_ON_COM_DELAY As Integer = 1
Private Const PWR_ON_GPS_DELAY As Integer = 10
Private Const SPEC_SVID As Integer = 0
Private Const SCAN_FOR_SV As Integer = 1
Private Const LBLS_01_TO_16 As Byte = 0
Private Const LBLS_17_TO_32 As Byte = 1
Private Const SNR_STRINGS As String = "--,--,--,--,--,--,--,--," & _
                                      "--,--,--,--,--,--,--,--," & _
                                      "--,--,--,--,--,--,--,--," & _
                                      "--,--,--,--,--,--,--,--"
Private Const MAX_SAT_ID As Byte = 32
Private Const NUM_OF_SNR_STATS As Byte = 3
'Private Const SNR_STRINGS As String = "00,01,02,03,04,05,06,07," & _
'                                      "10,11,12,13,14,15,16,17," & _
'                                      "20,21,22,23,24,25,26,27," & _
'                                      "30,31,32,33,34,35,36,37"
' TODO: This averaging algoritm won't work unless we use floats
'       otherwise small changes will just get rounded out
' Need to average over a finite period.
' Desired time in minutes is D =>
'    MAX_SNR_AVG_CNT = D min * 60 sec/1 min * 2 tmrEvents/1 sec * 1 avgIncrement/MAX_SAT_ID tmrEvents
'                    =  (D * 60 * 2 * 1) / (1 * 1 * MAX_SAT_ID)
'                    =  D * 120 / MAX_SAT_ID
' If D = 30 min and MAX_SAT_ID = 32, MAX_SNR_AVG_CNT = 112.5
Private Const MAX_SNR_AVG_CNT As Long = 112

Private m_szArrSNR_Last() As String
Private m_szArrSNR_Avg() As String
Private m_szArrSNR_Max() As String
Private m_vStartTime
Private m_bKillTimer As Boolean
Private m_iStartGPSCommands As Integer
Private m_iStartCOMCommands As Integer
Private m_bWarnOnComFailure As Boolean
Private m_uSatScanID As Byte
Private m_lSatSNR_AvgCnt As Long


Private Sub cmdOK_Click()

    '   Stop the timer then turn the power off from the device
    tmrGPS.Enabled = False
    If SerialComs.PSOutput Then
        SerialComs.PSOutput = False
    End If
    g_bDiagnosticActive = False
    Me.Hide

End Sub

Public Sub DisplayGPSCoordinates(lLongitude As Long, lLatitude As Long, iNoSatellites As Integer)

    Dim bNorth As Boolean
    Dim bEast As Boolean
    Dim lDegrees As Long
    Dim lMinutes As Long
    Dim dSeconds As Double

    '   Brake down the Longitude and then reformat it into degrees, minutes and
    '   seconds then post it to the form.
    If lLongitude < 0 Then
        bEast = False
        lLongitude = -lLongitude
    Else
        bEast = True
    End If
    lDegrees = lLongitude \ 100000
    lLongitude = lLongitude - (lDegrees * 100000)
    lMinutes = lLongitude \ 1000
    dSeconds = lLongitude - (lMinutes * 1000)
    dSeconds = (dSeconds / 1000) * 60
    lblLongitude.Caption = CStr(lDegrees) & "° " & CStr(lMinutes) & "' " & CStr(dSeconds) & Chr(&H22)
    If bEast Then
        lblLongitude.Caption = lblLongitude.Caption + " E"
    Else
        lblLongitude.Caption = lblLongitude.Caption + " W"
    End If
    '   Brake down the Latitude and then reformat it into degrees, minutes and
    '   seconds then post it to the form.
    If lLatitude < 0 Then
        bNorth = False
        lLatitude = -lLatitude
    Else
        bNorth = True
    End If
    lDegrees = lLatitude \ 100000
    lLatitude = lLatitude - (lDegrees * 100000)
    lMinutes = lLatitude \ 1000
    dSeconds = lLatitude - (lMinutes * 1000)
    dSeconds = (dSeconds / 1000) * 60
    lblLatitude.Caption = CStr(lDegrees) & "° " & CStr(lMinutes) & "' " & CStr(dSeconds) & Chr(&H22)
    If bNorth Then
        lblLatitude.Caption = lblLatitude.Caption + " N"
    Else
        lblLatitude.Caption = lblLatitude.Caption + " S"
    End If
    '   Post the number of satellites to the form.
    lblNumOfSatellites.Caption = CStr(iNoSatellites)

End Sub

Private Sub Form_Activate()

    DiagGPS_Controls (False)
    tmrGPS.Enabled = True
    g_bDiagnosticActive = True
    m_uSatScanID = 0

End Sub

Private Sub Form_Load()
'    Dim szArrSNR_StringsTmp() As String
'    Dim bOuterIdx As Byte
'    Dim bInnerIdx As Byte

    m_bKillTimer = False
    
    If False = g_fIO.GPS_Ext Then
        ' GPS switch is set to use on-board antenna
        Option_AntennaOut(0).value = True
        Option_AntennaOut(1).value = False
    Else
        ' GPS switch is set to use external antenna
        Option_AntennaOut(0).value = False
        Option_AntennaOut(1).value = True
    End If
    
    Text_ID.Text = g_szGPSID
    
    ' Parse the SNR_STRINGS data
    m_szArrSNR_Last = Split(SNR_STRINGS, ",")
    m_szArrSNR_Avg = m_szArrSNR_Last
    m_szArrSNR_Max = m_szArrSNR_Last

    SetScannedSNR_FormLabels
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    m_bKillTimer = True
    tmrGPS.Enabled = False
    tmrGPS.Interval = 0

End Sub

Private Sub Option_AntennaIn_Click(Index As Integer)
'    If True = Option_AntennaIn(0).value Then
    If ANT_IN_ON_BOARD = Index Then
        ' Option_AntennaIn(ANT_IN_EXTERNAL).value = False
        ' Default, select on-board antenna
        If False = SerialComs.SetGpsAntenna(eElite_GPS_INT_ANT) Then
            Option_AntennaIn(ANT_IN_ON_BOARD).value = False
        End If
    Else
        ' External antenna selected
        ' Option_AntennaIn(ANT_IN_ON_BOARD).value = False
        If False = SerialComs.SetGpsAntenna(eElite_GPS_EXT_ANT) Then
            Option_AntennaIn(ANT_IN_EXTERNAL).value = False
        End If
    End If
End Sub

Private Sub Option_AntennaOut_Click(Index As Integer)
'    If True = Option_AntennaOut(GPS_OUT_TO_TESTER_ANT).value Then
    If GPS_OUT_TO_TESTER_ANT = Index Then
        ' Default, send GPS signal to test fixture antenna
        If True = g_fIO.GPS_Ext Then
            g_fIO.GPS_Ext = False
        End If
    Else
        ' Send GPS signal directly to device's external antenna connector
        If False = g_fIO.GPS_Ext Then
            g_fIO.GPS_Ext = True
        End If
    End If
End Sub
Private Sub SetScannedSNR_FormLabels()
    Dim szTmp As String
    Dim szTmpArr() As String
    
    szTmp = Join(m_szArrSNR_Last, " ")
    szTmpArr = Split(szTmp, " ", ((MAX_SAT_ID / 2) + 1))
    lbl_SNRs(1).Caption = szTmpArr((MAX_SAT_ID / 2))
    ReDim Preserve szTmpArr((MAX_SAT_ID / 2) - 1)
    lbl_SNRs(0).Caption = Join(szTmpArr, " ")
    
    szTmp = Join(m_szArrSNR_Avg, " ")
    szTmpArr = Split(szTmp, " ", ((MAX_SAT_ID / 2) + 1))
    lbl_SNRs(3).Caption = szTmpArr((MAX_SAT_ID / 2))
    ReDim Preserve szTmpArr((MAX_SAT_ID / 2) - 1)
    lbl_SNRs(2).Caption = Join(szTmpArr, " ")
    
    szTmp = Join(m_szArrSNR_Max, " ")
    szTmpArr = Split(szTmp, " ", ((MAX_SAT_ID / 2) + 1))
    lbl_SNRs(5).Caption = szTmpArr((MAX_SAT_ID / 2))
    ReDim Preserve szTmpArr((MAX_SAT_ID / 2) - 1)
    lbl_SNRs(4).Caption = Join(szTmpArr, " ")

End Sub
Private Sub DiagGPS_Controls(bEnable As Boolean)
        
        m_bWarnOnComFailure = bEnable
        Frame_GPS_In.Enabled = bEnable
        Option_AntennaIn(0).Enabled = bEnable
        Option_AntennaIn(1).Enabled = bEnable
' TODO: Update commands (if they exist) to get longitude, latitude and number of satellites
'        lblLongitudeTitle.Enabled = bEnable
'        lblLatitudeTitle.Enabled = bEnable
'        lblNumOfSatsTitle.Enabled = bEnable
'        lblLongitude.Enabled = bEnable
'        lblLatitude.Enabled = bEnable
'        lblNumOfSatellites.Enabled = bEnable
        lblSNRTitle.Enabled = bEnable
        lblSNR.Enabled = bEnable
        
        If True = bEnable Then
            '   Update the total running time
            lblRunningTime.Caption = Format(Time - m_vStartTime, "hh:mm:ss") & " "
            '   Don't start issuing GPS signals for at least 10 seconds after power
            '   has been applied to the DUT.
        Else
            ' If tmrGPS.Interval = 500ms, then m_iStartGPSCommands = 20
            ' and if m_iStartGPSCommands is used as a decrementing
            ' conditional, the conditional will reach zero 10 seconds
            ' after the timer event starts
            m_iStartGPSCommands = (PWR_ON_GPS_DELAY * 1000) / tmrGPS.Interval
            m_iStartCOMCommands = (PWR_ON_COM_DELAY * 1000) / tmrGPS.Interval

            lblLongitude.Caption = "--° --' --.--" & Chr(&H22)
            lblLatitude.Caption = "--° --' --.--" & Chr(&H22)
            lblRunningTime = "00:00:00 "
            lblNumOfSatellites.Caption = "-"
            lblSNR.Caption = ""
            m_vStartTime = Time
        End If
End Sub

Private Sub Option_ScanSatMode_Click(Index As Integer)
    Dim idx As Byte
    Dim bActionSpecSV As Boolean
    Dim bActionScan As Boolean
    
    bActionSpecSV = SPEC_SVID = Index
    bActionScan = Not bActionSpecSV
    
    ' Toggle enabled for the "Specified Satellite Vehicle" portion of the form
    lblSNRTitle.Enabled = bActionSpecSV
    lblIdTitle.Enabled = bActionSpecSV
    Text_ID.Enabled = bActionSpecSV
    lblSNR.Enabled = bActionSpecSV
    
    ' Toggle enabled for the "Scan for Satellite Vehicle" portion of the form
    If True = bActionScan Then
        ' Clear the displayed SNRs
        m_szArrSNR_Last = Split(SNR_STRINGS, ",")
        m_szArrSNR_Avg = m_szArrSNR_Last
        m_szArrSNR_Max = m_szArrSNR_Last
        SetScannedSNR_FormLabels
        m_uSatScanID = 0
    End If
    
    lbl_SatIDs(0).Enabled = bActionScan
    lbl_SatIDs(1).Enabled = bActionScan
    For idx = 0 To NUM_OF_SNR_STATS - 1
        lbl_SNRsTitle(idx).Enabled = bActionScan
        lbl_SNRs((idx * 2)).Enabled = bActionScan
        lbl_SNRs((idx * 2) + 1).Enabled = bActionScan
    Next idx
        
    lbl_SatIDsTitle.Enabled = bActionScan

End Sub

Private Sub tmrGPS_Timer()

    Dim i As Integer
    Dim szVersion As String
    Dim lLatitude As Long
    Dim lLongitude As Long
    Dim iNoSatellites As Integer
    Dim szSatID As String
    Dim uSatID As Byte
    Dim uMaxSatID As Byte
    Dim uSatSNR As Byte
    Dim lSatSNR_Aggregate As Long


    If m_bKillTimer Then
        tmrGPS.Enabled = False
        tmrGPS.Interval = 0
        Exit Sub
    End If
    
    If " " = Text_ID.Text Then
                Text_ID.Text = ""
        End If

    If "" = Text_ID.Text Then
        If "-" <> lblSNR.Caption Then
            DiagGPS_Controls (False)
            lblSNR.Caption = "-"
        End If
        Exit Sub
    ElseIf "-" = lblSNR.Caption Then
       lblSNR.Caption = ""
    End If

    If g_fIO.IsTesterClosed Then
        '   Update the total running time
        lblOperationalState.Caption = "Diagnostic Running"
        lblOperationalState.ForeColor = vbGreen
        lblRunningTime.Caption = Format(Time - m_vStartTime, "hh:mm:ss") & " "
        
        '   Check to see if the power has been turned on
        If Not SerialComs.PSOutput Then
            '   Power up the device
            SerialComs.PSOutput = True
            Exit Sub
        Else
            If 0 < m_iStartCOMCommands Then
                '   Wait a second before attempting to connect to the device.
                m_iStartCOMCommands = m_iStartCOMCommands - 1
                Exit Sub
            End If
        End If
        
        If False = g_fMainTester.ConfigureDeviceForTesting Then
            If g_bQuitProgram Then
                g_bQuitProgram = False
            End If
            If True = m_bWarnOnComFailure Then
                DiagGPS_Controls (False)
            End If
            Exit Sub
        ElseIf 0 = m_iStartCOMCommands Then
            ' Make sure that the DUT is using the antenna we are displaying
            If False = Option_AntennaIn(0).value And False = Option_AntennaIn(1).value Then
                MsgBox "DUT failed to switch antennas"
                cmdOK_Click
                Exit Sub
            End If
            m_iStartCOMCommands = -1
            Frame_GPS_In.Enabled = True
            Option_AntennaIn(0).Enabled = True
            Option_AntennaIn(1).Enabled = True
        End If
        
        '   Don't start issuing GPS signals for at least 10 seconds after power
        '   has been applied to the DUT.
        If m_iStartGPSCommands <= 0 Then
            '   Get the GPS coordinates
'            If SerialComs.GetGPSCoordinates(lLongitude, lLatitude, iNoSatellites) Then
'                DisplayGPSCoordinates lLongitude, lLatitude, iNoSatellites
'            End If
            '   Get the GPS signal strength
            ' TODO: Need better checking on the Text_ID.Text form field
            If "" = Text_ID.Text Then
                ' Stop reading the signal strength and free up the COM port so
                ' that we can use hyperterminal if we need to
                lblSNR.Caption = "-"
                DiagGPS_Controls (False)
            Else
                ' Read the GPS signal strength
                szSatID = Text_ID.Text
                If True = Option_ScanSatMode(SPEC_SVID).value Then
                    ' Setup to read GPS SNR from a single satellite
                    uMaxSatID = CByte(Val(Text_ID.Text))
                    uSatID = uMaxSatID - 1
                Else
                    ' Setup to read GPS SNR from incrementing satellite
                    uSatID = m_uSatScanID
                    uMaxSatID = MAX_SAT_ID
                End If
                    ' Convert zero based array index number to SV ID string
                    ' TODO: Create array for SV ID strings and don't rely on the index value
                    szSatID = CStr(uSatID + 1)
                    
                    ' Read GPS SNR for the specified satellite
                    If True = SerialComs.FindGPS(szSatID) Then
                        ' AT command was successful
                        ' Store formatted string in the "last GPS SNR read" array
                        m_szArrSNR_Last(uSatID) = Format(szSatID, "00")
                        
                        ' Check max SNR and set if necessary
                        If "--" = m_szArrSNR_Max(uSatID) Then
                            ' Initial read so max SNR is the last read SNR
                            m_szArrSNR_Max(uSatID) = m_szArrSNR_Last(uSatID)
                        Else
                            ' See if the newly read SNR is greater than the old "max" SNR
                            uSatSNR = CByte(Val(m_szArrSNR_Last(uSatID)))
                            If CByte(Val(m_szArrSNR_Max(uSatID))) < uSatSNR Then
                                ' Store new max SNR in the "max GPS SNR" array
                                m_szArrSNR_Max(uSatID) = uSatSNR
                            End If
                        End If
                        
                        ' Check avg SNR and set if necessary
                        If "--" = m_szArrSNR_Avg(uSatID) Then
                            ' Initial read so avg SNR is the last read SNR
                            m_szArrSNR_Avg(uSatID) = m_szArrSNR_Last(uSatID)
                            m_lSatSNR_AvgCnt = 1
                        Else
                            ' Calculate the new average
                            uSatSNR = CByte(Val(m_szArrSNR_Avg(uSatID)))
                            lSatSNR_Aggregate = ((uSatSNR * m_lSatSNR_AvgCnt) + m_szArrSNR_Last(uSatID)) / (m_lSatSNR_AvgCnt + 1)
                            
                            If (MAX_SAT_ID - 1) = uSatID Then
                                ' We've gotten the SNRs for all of the defined satellites
                                If MAX_SNR_AVG_CNT <> m_lSatSNR_AvgCnt Then
                                    ' Only increment for the first MAX_SNR_AVG_CNT times
                                    m_lSatSNR_AvgCnt = m_lSatSNR_AvgCnt + 1
                                End If
                            End If
                            m_szArrSNR_Avg(uSatID) = Format(CStr(lSatSNR_Aggregate), "00")
                        End If
                        
                        If True = Option_ScanSatMode(SPEC_SVID).value Then
                            lblSNR.Caption = szSatID
                        Else
                            m_szArrSNR_Last((uSatID + 1) Mod MAX_SAT_ID) = "--"
                        End If
                        SetScannedSNR_FormLabels
                    Else
                        lblSNR.Caption = ""
                        m_szArrSNR_Last(uSatID) = "--"
                    End If
                    m_uSatScanID = (m_uSatScanID + 1) Mod MAX_SAT_ID
'                Next uSatID
            End If
        Else
            m_iStartGPSCommands = m_iStartGPSCommands - 1
            If m_iStartGPSCommands = 1 Then
                If Not SerialComs.GetAppVer(szVersion) Then
                    '   Tell them the problem then exit the dialog
                    MsgBox "The inserted unit does not have a valid application downloaded into it yet!" & vbCrLf & vbCrLf & _
                           "You need to insert a fully downloaded unit in order to run this diagnostic", vbCritical
                    cmdOK_Click
                End If
            '   Ensure that all of the controls have been enabled
            DiagGPS_Controls (True)
            End If
        End If
    Else
        '   The test jig is open.
        '   Change the Operational State label and disable all controls
        '   except for the OK button
        lblOperationalState = "Test jig open"
        lblOperationalState.ForeColor = vbRed
        
        DiagGPS_Controls (False)
        
        '   Close connections to the device if necessary
        OpenEliteComPort False, False

        '   Check to see if the power has been turned off
        If SerialComs.PSOutput Then
            '   Power down the device
            SerialComs.PSOutput = False
        End If
    End If

End Sub

