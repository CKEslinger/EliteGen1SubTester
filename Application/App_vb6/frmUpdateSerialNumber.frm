VERSION 5.00
Begin VB.Form frmUpdateSerialNumber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Serial Numbers"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStaticPost 
      Alignment       =   1  'Right Justify
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
      Left            =   4800
      MaxLength       =   8
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtStaticPre 
      Alignment       =   1  'Right Justify
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
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtNextNo 
      Alignment       =   1  'Right Justify
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
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtEndingNo 
      Alignment       =   1  'Right Justify
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
      Left            =   4800
      MaxLength       =   8
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtStartingNo 
      Alignment       =   1  'Right Justify
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
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   4
      Top             =   600
      Width           =   2175
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
      Left            =   4950
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblStaticPost 
      Alignment       =   1  'Right Justify
      Caption         =   "SV:"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label lblStaticPre 
      Alignment       =   1  'Right Justify
      Caption         =   "TAC:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Next available:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "End of range:"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Start of range:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   1095
   End
End
Attribute VB_Name = "frmUpdateSerialNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_szStaticPre As String
Public m_bUpdate As Boolean
Public m_szNextNo As String
Public m_szEndingNo As String
Public m_szStartingNo As String
Public m_szStaticPost As String

Private Sub cmdCancel_Click()

    m_bUpdate = False
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    If Val(txtStartingNo) >= Val(txtEndingNo) Then
        MsgBox "ERROR:  The start of the range must be less than the end of the range!"
    ElseIf Val(txtNextNo) < Val(txtStartingNo) Then
        MsgBox "ERROR:  The next serial number can not be less than the start of the range!"
    ElseIf Val(txtNextNo) >= Val(txtEndingNo) Then
        MsgBox "ERROR:  The next serial number must be less than the end of the range!"
    Else
        m_szStaticPre = txtStaticPre
        m_szNextNo = txtNextNo
        m_szEndingNo = txtEndingNo
        m_szStartingNo = txtStartingNo
        m_szStaticPost = txtStaticPost
        Me.Hide
    End If

End Sub

Private Sub Form_Activate()

    m_bUpdate = False

End Sub

Private Sub txtEndingNo_Change()

    m_bUpdate = True

End Sub

Private Sub txtNextNo_Change()

    m_bUpdate = True

End Sub

Private Sub txtStartingNo_Change()

    m_bUpdate = True

End Sub

Private Sub txtStaticPre_Change()

    m_bUpdate = True

End Sub

Private Sub txtStaticPost_Change()

    m_bUpdate = True

End Sub

