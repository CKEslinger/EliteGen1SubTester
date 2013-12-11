VERSION 5.00
Begin VB.Form frmContents 
   Caption         =   "Test Descriptions"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtContents 
      Height          =   5775
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Contents Help not yet available"
      Top             =   240
      Width           =   9375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   6240
      Width           =   1575
   End
End
Attribute VB_Name = "frmContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONTENTS_FILENAME As String = "contents.rtf"

Private Sub cmdOK_Click()

    Me.Hide

End Sub

Private Sub Form_Load()

'    rtbContents.LoadFile CONTENTS_FILENAME

End Sub
