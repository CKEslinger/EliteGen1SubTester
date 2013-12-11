VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   4710
   ClientLeft      =   4470
   ClientTop       =   3600
   ClientWidth     =   8505
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Tag             =   "Options"
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Tag             =   "&Apply"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Tag             =   "Cancel"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Tag             =   "OK"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame fraFormat 
      Height          =   1785
      Index           =   1
      Left            =   4200
      TabIndex        =   8
      Tag             =   "DataFormat"
      Top             =   2160
      Width           =   3720
      Begin VB.OptionButton OptionDataFormat 
         Caption         =   "Text"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptionDataFormat 
         Caption         =   "CSV"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   855
         Width           =   1575
      End
   End
   Begin VB.Frame fraFormat 
      Height          =   1785
      Index           =   0
      Left            =   4200
      TabIndex        =   6
      Tag             =   "DataFormat"
      Top             =   120
      Width           =   3720
      Begin VB.ListBox ListBoxModel 
         Height          =   1230
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   5
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   4
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   2
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Models"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Display"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum etDISPLAY_FORMAT
    eTEXT
    eCSV
    eDEFAULT = eTEXT
End Enum

Private m_bDisplayFormat As etDISPLAY_FORMAT
Private m_bModelListIndex As Integer

Public Property Get DisplayFormat() As etDISPLAY_FORMAT
    DisplayFormat = m_bDisplayFormat
End Property

Private Property Let DisplayFormat(value As etDISPLAY_FORMAT)
    m_bDisplayFormat = value
End Property

Public Property Get ModelListIdx() As Integer
    If m_bModelListIndex > ListBoxModel.ListCount Then
        m_bModelListIndex = -1
    End If
    ModelListIdx = m_bModelListIndex
End Property

Private Property Let ModelListIdx(value As Integer)
    m_bModelListIndex = value
End Property

Private Sub cmdApply_Click()
    If True = OptionDataFormat(etDISPLAY_FORMAT.eTEXT).value Then
        DisplayFormat = etDISPLAY_FORMAT.eTEXT
    ElseIf True = OptionDataFormat(etDISPLAY_FORMAT.eCSV).value Then
        DisplayFormat = etDISPLAY_FORMAT.eCSV
    End If
End Sub

Private Sub cmdCancel_Click()
    ' Unload Me
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    ' Unload Me
    Me.Hide
End Sub

Private Sub Form_Activate()
    ' Set Models tab defaults
    If Not g_oModels Is Nothing Then
        Dim oModel As IXMLDOMNode
        Dim i As Integer
        
        ' Clear the list
        For i = ListBoxModel.ListCount - 1 To 0 Step -1
            ListBoxModel.RemoveItem (i)
        Next i

        For Each oModel In g_oModels.childNodes
            If oModel.nodeType = NODE_ELEMENT Then
                ListBoxModel.AddItem (oModel.baseName)
            End If
        Next oModel
    End If
    ListBoxModel.ListIndex = ModelListIdx
    

End Sub

Private Sub Form_Initialize()
    Dim i As Integer

    ' Size the form to fit the tabstrip and the buttons along the bottom
    Me.Height = tbsOptions.Height + cmdOK.Height + 800
    Me.Width = tbsOptions.Width + 350

    ' Make the tab frames fit in the tabstrip
    For i = 0 To tbsOptions.Tabs.Count - 1
        fraFormat(i).Height = tbsOptions.Height - 500
        fraFormat(i).Width = tbsOptions.Width - 200
    Next
    
    tbsOptions_Click
    
    ' Set Model tab defaults
    ModelListIdx = ListBoxModel.ListIndex
    
    ' Set Display tab defaults
    DisplayFormat = etDISPLAY_FORMAT.eDEFAULT
    OptionDataFormat(DisplayFormat).value = True
End Sub


Private Sub ListBoxModel_Click()
    ModelListIdx = ListBoxModel.ListIndex
End Sub

Private Sub tbsOptions_Click()
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            fraFormat(i).Top = tbsOptions.Top + 350
            fraFormat(i).Left = tbsOptions.Left + 100
            fraFormat(i).Enabled = True
        Else
            fraFormat(i).Left = -20000
            fraFormat(i).Enabled = False
        End If
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub
