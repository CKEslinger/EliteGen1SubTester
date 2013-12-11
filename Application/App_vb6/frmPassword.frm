VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Window"
   ClientHeight    =   2685
   ClientLeft      =   2190
   ClientTop       =   2100
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Operation is Password Protected!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please Enter Password"
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
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_szPassword As String

Private Sub cmdCancel_Click()

    m_szPassword = "Cancel"
    Me.Hide

End Sub

Private Sub cmdEnter_Click()

    m_szPassword = txtPassword.Text
    Me.Hide

End Sub

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdEnter_Click
    End If

End Sub

Public Function CheckPassword() As Boolean

    m_szPassword = ""
    Me.Show vbModal
    If m_szPassword = g_oSettings.selectSingleNode("Company").Text Then
        CheckPassword = True
    Else
        CheckPassword = False
        If Not m_szPassword = "Cancel" Then
            MsgBox "The password you typed is incorrect", vbCritical, "Incorrect Password"
        End If
    End If
    txtPassword.Text = ""

End Function
