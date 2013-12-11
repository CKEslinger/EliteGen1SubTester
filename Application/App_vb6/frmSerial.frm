VERSION 5.00
Begin VB.Form frmSerialNumber 
   Caption         =   "Barcode Scan"
   ClientHeight    =   2790
   ClientLeft      =   3015
   ClientTop       =   2445
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReprint 
      Caption         =   "Reprint"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdRescan 
      Caption         =   "Rescan"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblFailure 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblScan 
      Caption         =   "Enter The Test Unit Serial Number"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmSerialNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnter_Click()

    If txtSerialNumber.Text = Format(g_lNextSerialNumber, "0000000") Then
        Globals.BarcodeVerify = True
        txtSerialNumber.Text = ""
        Me.Hide
    Else
        txtSerialNumber.Text = ""
        lblScan.Visible = False
        txtSerialNumber.Visible = False
        cmdEnter.Visible = False
        lblFailure.Caption = "The scanned barcode label is not equal to the DUT serial number.  Please click Rescan to rescan label or Reprint to print a new barcode label."
        lblFailure.Visible = True
        cmdRescan.Visible = True
        cmdReprint.Visible = True
    End If

End Sub

Private Sub cmdReprint_Click()

    g_fMainTester.BarcodeRePrint
    lblScan.Visible = True
    txtSerialNumber.Visible = True
    cmdEnter.Visible = True
    lblFailure.Visible = False
    cmdRescan.Visible = False
    cmdReprint.Visible = False
    txtSerialNumber.SetFocus

End Sub

Private Sub cmdRescan_Click()

    lblScan.Visible = True
    txtSerialNumber.Visible = True
    cmdEnter.Visible = True
    lblFailure.Visible = False
    cmdRescan.Visible = False
    cmdReprint.Visible = False
    txtSerialNumber.SetFocus

End Sub

Private Sub txtSerialNumber_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdEnter.value = 1
    End If

End Sub

Private Sub Form_Activate()

   lblFailure.Caption = "The scanned barcode label is not equal to the DUT serial number.  Please click Rescan to rescan label or Reprint to print a new barcode label."
   lblScan.Visible = True
   txtSerialNumber.Visible = True
   cmdEnter.Visible = True
   lblFailure.Visible = False
   cmdRescan.Visible = False
   cmdReprint.Visible = False
   txtSerialNumber.SetFocus

End Sub

