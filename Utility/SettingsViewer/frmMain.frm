VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "SettingsViewer"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleMode       =   0  'User
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Text_Frame 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8415
      Begin VB.TextBox Main_Text 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   8175
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5895
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9604
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "7/9/2013"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:08 AM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9000
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   9480
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":005E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00BC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":011A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0178
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01D6
            Key             =   "Cut"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fOptionsForm As frmOptions

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)

    Set fOptionsForm = New frmOptions
End Sub


Private Sub Form_Resize()
    Text_Frame.Width = Width - 360
    Text_Frame.Height = Height - 1200
    Main_Text.Width = Text_Frame.Width - 240
    Main_Text.Height = Text_Frame.Height - 500
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Cut"
            mnuEditCut_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
    fOptionsForm.Show
    While fOptionsForm.Visible
        DoEvents
    Wend
        
    mnuViewRefresh_Click
End Sub

Private Sub mnuViewRefresh_Click()
    Main_Text.Text = ""
    
    CreateMainText
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    'ToDo: Add 'mnuEditPaste_Click' code.
    MsgBox "Add 'mnuEditPaste_Click' code."
End Sub

Private Sub mnuEditCopy_Click()
    'ToDo: Add 'mnuEditCopy_Click' code.
    MsgBox "Add 'mnuEditCopy_Click' code."
End Sub

Private Sub mnuEditCut_Click()
    'ToDo: Add 'mnuEditCut_Click' code.
    MsgBox "Add 'mnuEditCut_Click' code."
End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
    'ToDo: Add 'mnuFilePrint_Click' code.
    MsgBox "Add 'mnuFilePrint_Click' code."
End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileSaveAs_Click()
    'ToDo: Add 'mnuFileSaveAs_Click' code.
    MsgBox "Add 'mnuFileSaveAs_Click' code."
End Sub

Private Sub mnuFileSave_Click()
    'ToDo: Add 'mnuFileSave_Click' code.
    MsgBox "Add 'mnuFileSave_Click' code."
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFileName As String
    Dim sFileTitle As String
    Dim szError As String

    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFileName = .FileName
        sFileTitle = .FileTitle
    End With
    
    '   Open and validate the Model Type XML configuration file
    szError = OpenModelTypeFile(sFileTitle, sFileName)

    If "" <> szError Then
        '   There was something wrong with the ModelType.xml file.  Abort the
        '   app now before anyone gets hurt.
        MsgBox szError
        Exit Sub
    End If
    
End Sub

Private Sub CreateMainText()
    Dim oModel As IXMLDOMNode
    Dim oChild As IXMLDOMNode
    Dim oSIM As IXMLDOMNode
    Dim oTmpNode As IXMLDOMNode
    Dim szYesNo As String
    Dim szTmpStr As String
    Dim iNodeIdx As Integer
    Dim oNode As IXMLDOMNode
    
  
    If -1 = fOptionsForm.ModelListIdx Then
        Main_Text.Text = "No Model Selected." & vbCrLf & "You must first open the XML file and then select the model." & vbCrLf
        Main_Text.Text = Main_Text.Text & "   File->Open, Browse to the XML file to open and click 'Open'" & vbCrLf
        Main_Text.Text = Main_Text.Text & "   View->Options, Click the 'Models' tab, select a model and click 'OK'" & vbCrLf
    Else
        ' Get selected model type
        Main_Text.Text = Main_Text.Text & "Model Type: " & fOptionsForm.ListBoxModel.Text & vbCrLf
        
        ' Get XML version number
        Set oNode = g_oVersion.selectSingleNode("XML_File")
    
        Main_Text.Text = Main_Text.Text & "XML File Version: " & oNode.Attributes.getNamedItem("Major").Text & "." & oNode.Attributes.getNamedItem("Minor").Text & vbCrLf
        For Each oChild In g_oVersion.childNodes
            Dim major, minor, rev As String
            Select Case oChild.baseName
            Case "MinAppVer"
                major = oChild.Attributes.Item(0).firstChild.Text
                minor = oChild.Attributes.Item(1).firstChild.Text
                rev = oChild.Attributes.Item(2).firstChild.Text
                Main_Text.Text = Main_Text.Text & "XML File Compatibility" & vbCrLf & " Min. Version of Tester Application: " & major & "." & minor & "." & rev & vbCrLf
            Case "MaxAppVer"
                major = oChild.Attributes.Item(0).firstChild.Text
                minor = oChild.Attributes.Item(1).firstChild.Text
                rev = oChild.Attributes.Item(2).firstChild.Text
                Main_Text.Text = Main_Text.Text & " Max. Version of Tester Application: " & major & "." & minor & "." & rev & vbCrLf
            End Select
        Next oChild
        
        ' Get Global Settings
        For Each oChild In g_oSettings.childNodes
            Dim minSig, maxSig As String
            Select Case oChild.baseName
            Case "GPSSimulator"
                minSig = oChild.Attributes.Item(2).firstChild.Text
                maxSig = oChild.Attributes.Item(3).firstChild.Text
                Main_Text.Text = Main_Text.Text & "GPS P/F Criteria (dBm):" & vbCrLf & " Min.: " & minSig & vbCrLf & " Max.: " & maxSig & vbCrLf
            Case "CellSiteSimulator"
                minSig = oChild.Attributes.Item(1).firstChild.Text
                maxSig = oChild.Attributes.Item(2).firstChild.Text
                Main_Text.Text = Main_Text.Text & "GPRS P/F Criteria (RSSI):" & vbCrLf & " Min.: " & minSig & vbCrLf & " Max.: " & maxSig & vbCrLf
            Case "TestSystem"
                For iNodeIdx = 0 To oChild.childNodes.length - 1
                    Set oTmpNode = oChild.childNodes(iNodeIdx).selectSingleNode("Computer")
                    Main_Text.Text = Main_Text.Text & "IMEI info for test station " & oTmpNode.Attributes.Item(0).Text
                    Main_Text.Text = Main_Text.Text & ": " & vbCrLf
                    Set oTmpNode = oChild.childNodes(iNodeIdx).selectSingleNode("IMEI")
                    Main_Text.Text = Main_Text.Text & "   TAC: " & oTmpNode.selectSingleNode("TAC").Text
                    Main_Text.Text = Main_Text.Text & vbCrLf & "   S/N Range: "
                    Set oTmpNode = oTmpNode.selectSingleNode("SNR")
                    szTmpStr = oTmpNode.Attributes.getNamedItem("Start").Text & " to " & oTmpNode.Attributes.getNamedItem("End").Text & vbCrLf
                    Main_Text.Text = Main_Text.Text & szTmpStr
                    Next iNodeIdx
            End Select
        Next oChild
        
        ' Get SIM for selected model type
        For Each oModel In g_oModels.childNodes
            If oModel.baseName = fOptionsForm.ListBoxModel.Text Then
                For Each oChild In oModel.childNodes
                    Select Case oChild.baseName
                    Case "Firmware"
                        Main_Text.Text = Main_Text.Text & "FW Rev.: " & oChild.Text & vbCrLf
                    Case "SIMs"
                        For Each oSIM In g_oSIMs.childNodes
                            If oSIM.baseName = oChild.Text Then
'                                Main_Text.Text = "Product #: " & oModel.childNodes.Item(0).Text & vbCrLf & Main_Text.Text
                                Main_Text.Text = Main_Text.Text & "Product #: " & oModel.childNodes.Item(0).Text & vbCrLf
                                Exit For
                            End If
                        Next oSIM
                    Case "Current"
                        Main_Text.Text = Main_Text.Text & "Current: " & oChild.Attributes.Item(0).firstChild.Text & "A < I < " & oChild.Attributes.Item(1).firstChild.Text & "A" & vbCrLf
                    Case "GPS"
                        If oChild.Attributes.Item(0).firstChild.Text = "1" Then
                            szYesNo = "yes"
                        Else
                            szYesNo = "no"
                        End If
                        Main_Text.Text = Main_Text.Text & "GPS Always On: " & szYesNo & vbCrLf
                    Case "Security"
                        If oChild.Attributes.Item(0).firstChild.Text = "1" Then
                            szYesNo = "yes"
                        Else
                            szYesNo = "no"
                        End If
                        Main_Text.Text = Main_Text.Text & "Security Valet: " & szYesNo & vbCrLf
                    Case "SerialNo"
                        Main_Text.Text = Main_Text.Text & "Serial # Range: " & oChild.Attributes.Item(0).firstChild.Text & " - " & oChild.Attributes.Item(1).firstChild.Text & vbCrLf
                        Main_Text.Text = Main_Text.Text & "Next Serial #: " & oChild.Attributes.Item(2).firstChild.Text & vbCrLf
                    Case "IMEI"
                        Main_Text.Text = Main_Text.Text & "Modem Model #: " & oChild.firstChild.Text & vbCrLf
                        Main_Text.Text = Main_Text.Text & "IMEI TAC: " & oChild.childNodes.Item(1).Text & vbCrLf
                        Main_Text.Text = Main_Text.Text & "Next IMEI SN: " & oChild.childNodes.Item(3).Attributes(2).Text & vbCrLf
                    End Select
                Next oChild
                Exit For
            End If
        Next oModel
        
        Main_Text.Text = Main_Text.Text & "SIM: " & oSIM.baseName & vbCrLf
        
        Dim sSimInfo As String
        Dim oAddress As IXMLDOMNode
        For Each oChild In oSIM.childNodes
            Select Case oChild.baseName
            Case "APN", "Server", "SMS_Reply", "Reset"
                sSimInfo = sSimInfo & "   " & oChild.baseName & ":" & vbCrLf
                If "APN" = oChild.baseName Then
                    sSimInfo = sSimInfo & "      Password: " & oChild.selectSingleNode("PW").Attributes.getNamedItem("Pass").Text & vbCrLf
                End If
                For Each oAddress In oChild.childNodes
                    If Nothing Is oAddress.Attributes.getNamedItem("Index") Then
                    Else
                        If "Reset" = oAddress.baseName Then
                            sSimInfo = sSimInfo & "      " & oAddress.Attributes.getNamedItem("Name").Text & " " & oAddress.Attributes.Item(1).baseName & ": " & oAddress.Attributes.Item(1).Text & vbCrLf
                        Else
                            sSimInfo = sSimInfo & "      " & oAddress.baseName & oAddress.Attributes.getNamedItem("Index").Text & ": " & oAddress.Attributes.Item(0).Text & vbCrLf
                        End If
                    End If
                Next oAddress
            Case "SMS_Mode"
                sSimInfo = sSimInfo & "   SMS Mode: " & vbCrLf
                sSimInfo = sSimInfo & "      " & oChild.firstChild.Attributes.Item(0).baseName & "=" & oChild.firstChild.Attributes.Item(0).Text & vbCrLf
                sSimInfo = sSimInfo & "      " & oChild.firstChild.Attributes.Item(1).baseName & "=" & oChild.firstChild.Attributes.Item(1).Text & vbCrLf
            Case "SrvcCntrAddr"
                If Nothing Is oChild.selectSingleNode("SCA") Then
                Else
                sSimInfo = "   Service Center Address:" & vbCrLf & _
                "      " & oChild.selectSingleNode("SCA").Attributes.Item(0).baseName & ": " & oChild.selectSingleNode("SCA").Attributes.Item(0).Text & vbCrLf & _
                sSimInfo
                End If
            Case "Port"
                sSimInfo = "   Port: " & oChild.Text & vbCrLf & sSimInfo
            End Select
        Next oChild
        
        Main_Text.Text = Main_Text.Text & sSimInfo
        Main_Text.Text = Main_Text.Text & "Computer Name: " & MainModule.GetCompName()
    End If
End Sub
