VERSION 5.00
Begin VB.Form frmDiagDAQ 
   Caption         =   "Relay Diagnostics"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPwr 
      Caption         =   "RF Output"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox chkPwr 
      Caption         =   "Power Supply Output On"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   840
      Width           =   2295
   End
   Begin VB.CheckBox chkPwr 
      Caption         =   "Starter Voltage Connected"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CheckBox chkPwr 
      Caption         =   "Ignition Voltage Connected"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Timer tmrReadAnalog 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   3240
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
      Height          =   492
      Left            =   4680
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblAnalog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblAnalogTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Red LED Voltage:"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblAnalogTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Green LED Voltage:"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblAnalog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblAnalog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblAnalog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblAnalog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblAnalogTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Battery Voltage:"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblAnalogTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Ignition/Starter Voltage:"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblAnalogTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Relay Voltage:"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblRelayState 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Relay State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmDiagDAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bKillTimers As Boolean

Private Sub chkPwr_Click(Index As Integer)

    Select Case Index
        Case 0
            If chkPwr(0).value = vbChecked Then
                SerialComs.PSOutput = True
                chkPwr(1).Enabled = True
            ElseIf chkPwr(0).value = vbUnchecked Then
                chkPwr(1).value = vbUnchecked
                chkPwr(1).Enabled = False
                SerialComs.PSOutput = False
            End If

        Case 1
            If chkPwr(1).value = vbChecked Then
                g_fIO.RF_XmtChar = True
            ElseIf chkPwr(1).value = vbUnchecked Then
                g_fIO.RF_XmtChar = False
            End If

        Case 2
            If chkPwr(2).value = vbChecked Then
                g_fIO.StarterInput = True
            ElseIf chkPwr(2).value = vbUnchecked Then
                g_fIO.StarterInput = False
            End If

        Case 3
            If chkPwr(3).value = vbChecked Then
                g_fIO.IgnitionInput = True
            ElseIf chkPwr(3).value = vbUnchecked Then
                g_fIO.IgnitionInput = False
            End If
        
    End Select

End Sub

Private Sub cmdOK_Click()

    '   Stop the timer then turn the power off from the device
    tmrReadAnalog.Enabled = False
    SerialComs.PSOutput = False
    g_bDiagnosticActive = False
    Me.Hide

End Sub

Private Sub Form_Activate()

    Dim i As Integer

    g_bDiagnosticActive = True
    For i = 0 To 3
        If i < 3 Then
            lblAnalog(i).Caption = ""
        End If
        chkPwr(i).value = vbUnchecked
    Next i
    tmrReadAnalog.Enabled = True

End Sub


Private Sub Form_Load()

    m_bKillTimers = False
    tmrReadAnalog.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    m_bKillTimers = True
    tmrReadAnalog.Enabled = False
    tmrReadAnalog.Interval = 0

End Sub


Private Sub tmrReadAnalog_Timer()

    Dim i As Integer
    Dim dVoltage As Double

    If m_bKillTimers Then
        tmrReadAnalog.Enabled = False
        tmrReadAnalog.Interval = 0
        Exit Sub
    End If
    dVoltage = g_fIO.DUTRelayVoltage
    lblAnalog(0).Caption = FormatNumber(dVoltage, 3)
    If g_fIO.DUTRelayState Then
        lblRelayState.Caption = "Relay Closed"
    Else
        lblRelayState.Caption = "Relay Open"
    End If
    
    dVoltage = g_fIO.PowerSupply1_Voltage
    lblAnalog(1).Caption = FormatNumber(dVoltage, 3)
    
    dVoltage = g_fIO.PowerSupply2_Voltage
    lblAnalog(2).Caption = FormatNumber(dVoltage, 3)
    
    dVoltage = g_fIO.LedVoltage(0)
    lblAnalog(3).Caption = FormatNumber(dVoltage, 3)
    
    dVoltage = g_fIO.LedVoltage(1)
    lblAnalog(4).Caption = FormatNumber(dVoltage, 3)

End Sub

