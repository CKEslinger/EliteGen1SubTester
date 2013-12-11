VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3765
   ClientLeft      =   30
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1335
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   4995
      TabIndex        =   4
      Top             =   120
      Width           =   5052
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "OK"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Width           =   1452
   End
   Begin VB.Label lblConfiguration 
      Alignment       =   2  'Center
      Caption         =   "Configuration Profile: x.x"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   5052
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Application Version: x.x.x"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   5052
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PassTime Elite II Functional Tester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   299
      TabIndex        =   0
      Top             =   1800
      Width           =   4935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    Dim oNode As IXMLDOMNode
    
    lblVersion.Caption = "Application Version: " & App.Major & "." & App.Minor & "  Build: " & App.Revision
    
    Set oNode = g_oVersion.selectSingleNode("XML_File")
    lblConfiguration.Caption = "Configuration Profile: " & oNode.Attributes.getNamedItem("Major").Text & "." & oNode.Attributes.getNamedItem("Minor").Text

End Sub
