VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTesterSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PassTime Tester Setup"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMinBuzzerAmplitude 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtGPSTimeout 
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CheckBox chkAutotest 
      Caption         =   "Enable Auto Test Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   2520
      Width           =   1332
   End
   Begin MSComDlg.CommonDialog cdApplicationFile 
      Left            =   3840
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Firmware files (*.hex)|*.hex"
      FilterIndex     =   1
      Flags           =   5
   End
   Begin VB.CommandButton cmdBrowseAppFile 
      Caption         =   "..."
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
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Width           =   372
   End
   Begin VB.Frame frLabelConfig 
      Caption         =   "Label Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   4095
      Begin VB.TextBox txtNumOfPrnLbls 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   372
      End
      Begin VB.TextBox txtNumOfScnLbls 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Text            =   "0"
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Labels to Print:"
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Labels to Scan:"
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   11520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   2520
      Width           =   1332
   End
   Begin VB.Label Label3 
      Caption         =   "Min. Buzzer Amplitude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "GPS Lock Wait Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblApplFile 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Application Firmware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmTesterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAutotest_Click()
    
    If chkAutotest.value = vbChecked Then
        g_bAutoTest = True
        '   Change the text of the start button in the main view to show that
        '   Autotest has been enabled.
        g_fMainTester.cmdStartTest.Caption = "Auto-Test Enabled"
        g_fMainTester.cmdStartTest.FontItalic = True
        g_fMainTester.cmdStartTest.FontBold = False
    Else
        g_bAutoTest = False
        '   Change the text of the start button in the main view to show that
        '   Manual test starting has been enabled.
        g_fMainTester.cmdStartTest.Caption = "START TEST"
        g_fMainTester.cmdStartTest.FontItalic = False
        g_fMainTester.cmdStartTest.FontBold = True
        g_fMainTester.cmdStartTest.Enabled = True
    End If

End Sub

Private Sub cmdBrowseAppFile_Click()
    Dim szFirmwarePathAndFN As String
    Dim szFirmwareFN As String

    On Error GoTo ErrHandler

    With cdApplicationFile
        .InitDir = g_szDataDirPath
        .Flags = cdlOFNHideReadOnly Or cdlOFNReadOnly Or cdlOFNNoChangeDir Or cdlOFNFileMustExist
        '   Go get the firmware file from the disk
        .ShowOpen
        szFirmwarePathAndFN = Left(.FileName, InStrRev(.FileName, ".") - 1)
        szFirmwareFN = Right(szFirmwarePathAndFN, Len(szFirmwarePathAndFN) - InStrRev(szFirmwarePathAndFN, "\"))
        If False = g_fMainTester.FirmwareVerFromFN(szFirmwareFN) Then
            Exit Sub
        End If
        g_szFirmwareFileName = szFirmwareFN
    End With
    lblApplFile.Caption = szFirmwarePathAndFN
    g_oSelectedModel.selectSingleNode("Firmware").Text = g_szFirmwareFileName

ErrHandler:

End Sub

Private Sub cmdCancel_Click()

    '   Restore everything on this form from the global settings
    Form_Activate
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    '   Update the global settings for the controls in this form.
    If chkAutotest.value = vbChecked Then
        g_bAutoTest = True
    Else
        g_bAutoTest = False
    End If
    g_iNumBarcodePrints = Val(txtNumOfPrnLbls.Text)
    g_iNumBarcodeScans = Val(txtNumOfScnLbls.Text)
    g_iGPSTimeout = Val(txtGPSTimeout.Text)
    g_iMinBuzzerAmplitude = Val(txtMinBuzzerAmplitude.Text)
    Me.Hide

End Sub

Private Sub Form_Activate()

    Dim szTest As String

    lblApplFile.Caption = g_szDataDirPath & g_szFirmwareFileName
    If g_bAutoTest Then
        chkAutotest.value = vbChecked
    Else
        chkAutotest.value = vbUnchecked
    End If
    txtNumOfPrnLbls.Text = Str(g_iNumBarcodePrints)
    txtNumOfScnLbls.Text = Str(g_iNumBarcodeScans)
    txtGPSTimeout.Text = Str(g_iGPSTimeout)
    txtMinBuzzerAmplitude.Text = Str(g_iMinBuzzerAmplitude)
    On Error GoTo ModelNotYetSelected
    szTest = g_oSelectedModel.baseName
    '   A Model has been selected.  Enable the file browser
    cmdBrowseAppFile.Enabled = True
    lblApplFile.Enabled = True
    Exit Sub

ModelNotYetSelected:
    '   A Model has not yet been selected.  Disable the file browser
    cmdBrowseAppFile.Enabled = False
    lblApplFile.Enabled = False
    Exit Sub

End Sub
