VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5535
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1335
      Left            =   241
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   120
      Width           =   5052
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Configuration Profile: x.x"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   241
      TabIndex        =   3
      Top             =   2760
      Width           =   5052
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   241
      TabIndex        =   2
      Top             =   1800
      Width           =   5052
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Application Version: x.x.x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   241
      TabIndex        =   1
      Top             =   2280
      Width           =   5052
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim oNode As IXMLDOMNode
    
    Label2.Caption = "Application Version: " & App.Major & "." & App.Minor & "  Build: " & App.Revision
    
    Set oNode = g_oVersion.selectSingleNode("XML_File")
    Label3.Caption = "Configuration Profile: " & oNode.Attributes.getNamedItem("Major").Text & "." & oNode.Attributes.getNamedItem("Minor").Text


End Sub
