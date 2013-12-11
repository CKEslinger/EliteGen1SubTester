VERSION 5.00
Begin VB.Form frmPrinterScripts 
   Caption         =   "Select directory for printer script files"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox dlbPrinterScripts 
      Height          =   3465
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrinterScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

    g_strPrintDir = dlbPrinterScripts.List(dlbPrinterScripts.ListIndex) & "\"
    Me.Hide

End Sub

