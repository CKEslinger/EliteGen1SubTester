VERSION 5.00
Begin VB.Form frmDiagPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Label Printing"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
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
      Left            =   3960
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtNumLblsScans 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtLptNum 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtNumLblsPrints 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "1"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtPrintNum 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Text            =   "0123456789ABCDEF0123"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtPrintNum 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Text            =   "0123456789ABCDE"
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtPrintNum 
      Height          =   285
      Index           =   0
      Left            =   3600
      TabIndex        =   0
      Text            =   "01234567"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "# of Lables to Scan"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Printer Port LPT#"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "# of Lables to Print"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ICCID"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "IMSI"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   375
      Width           =   495
   End
End
Attribute VB_Name = "frmDiagPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Me.Hide

End Sub

Private Sub cmdPrint_Click()

    Dim tempSN As Long
    Dim tempIMSI As String
    Dim tempICCID As String
    Dim tempNumPrints As Integer
    Dim tempNumScans As Integer

    If Len(txtPrintNum(0).Text) > 8 Or _
       Len(txtPrintNum(1).Text) > 15 Or _
       Len(txtPrintNum(2).Text) > 20 Then
        MsgBox "Barcode Diagnostic Error", vbCritical
    Else
        tempSN = g_lNextSerialNumber
        g_lNextSerialNumber = CLng(txtPrintNum(0).Text)
        g_szIMSI = txtPrintNum(1).Text
        g_szCCID = txtPrintNum(2).Text
        tempNumPrints = g_iNumBarcodePrints
        g_iNumBarcodePrints = CInt(txtNumLblsPrints.Text)
        tempNumScans = g_iNumBarcodeScans
        g_iNumBarcodeScans = CInt(txtNumLblsScans.Text)
        g_clsBarcode.PrintScan txtLptNum
        g_lNextSerialNumber = tempSN
        g_iNumBarcodePrints = tempNumPrints
        g_iNumBarcodeScans = tempNumScans
    End If

End Sub
