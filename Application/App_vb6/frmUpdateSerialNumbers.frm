VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateSerialNumbers 
   Caption         =   "Update Serial Numbers"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
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
      Left            =   7800
      TabIndex        =   3
      Top             =   4320
      Width           =   1455
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
      Left            =   6120
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwSerialNumbers 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "WARNING!   Authorized Use of this Function ONLY!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "frmUpdateSerialNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum etITEMS
    eITEM_STATIC_PRE = 1
    eITEM_START
    eITEM_END
    eITEM_NEXT
    eITEM_STATIC_POST
    eITEM_NO_OF_ITEMS
End Enum

Public Enum etID_TYPE
    eSERIAL_NUM
    eSERIAL_NUM_READ_ONLY
    eIMEI_NUM
    eIMEI_NUM_READ_ONLY
End Enum

Private m_bRefresh As Boolean
Private m_bChangesMade As Boolean
Private m_fUpdateSerialNumber As frmUpdateSerialNumber
Private m_eIdType As etID_TYPE

Public Property Let UpdateType(eIdType As etID_TYPE)
    m_eIdType = eIdType
End Property

Private Sub cmdCancel_Click()

    '   Just blow all of these changes away and return back to the setup
    '   screen.
    m_bRefresh = True
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Dim lsiRow As ListItem
    Dim oNode As IXMLDOMNode
    Dim vbResult As VbMsgBoxResult
    Dim oAttributes As IXMLDOMNamedNodeMap

    If m_bChangesMade Then
        vbResult = MsgBox("Serial numbers WILL be updated back to the XML configuration file!" & vbCrLf & vbCrLf & _
                          "Are you certain you want to proceed?", vbYesNoCancel Or vbCritical)
    Else
        '   Since no changes have been made just drop through and exit the
        '   form.
        vbResult = vbCancel
    End If
    If vbResult = vbNo Then
        Exit Sub
    ElseIf vbResult = vbYes Then
        '   Update the changes back to the XML file
        For Each oNode In g_oModels.childNodes
            If oNode.nodeType = NODE_ELEMENT Then
                '   Find out which row this Model Type has been sorted into
                Set lsiRow = lvwSerialNumbers.FindItem(oNode.baseName)
                If etID_TYPE.eIMEI_NUM = m_eIdType Then
                    oNode.selectSingleNode("IMEI").selectSingleNode("TAC").Text = lsiRow.SubItems(eITEM_STATIC_PRE)
                    oNode.selectSingleNode("IMEI").selectSingleNode("SV").Text = lsiRow.SubItems(eITEM_STATIC_POST)
                    Set oAttributes = oNode.selectSingleNode("IMEI").selectSingleNode("SNR").Attributes
                Else
                    Set oAttributes = oNode.selectSingleNode("SerialNo").Attributes
                End If
                
                ' Update the incrementing serial number information
                oAttributes.getNamedItem("Start").Text = lsiRow.SubItems(eITEM_START)
                oAttributes.getNamedItem("End").Text = lsiRow.SubItems(eITEM_END)
                oAttributes.getNamedItem("Next").Text = lsiRow.SubItems(eITEM_NEXT)
            End If
        Next oNode
    End If
    '   vbResult = vbCancel will simply drop through to here without any of the
    '   above processing taking place.
    m_bRefresh = True
    Me.Hide

End Sub

Private Sub Form_Activate()

    Dim lsiRow As ListItem
    Dim oNode As IXMLDOMNode
    Dim oAttributes As IXMLDOMNamedNodeMap

    If m_bRefresh Then
        '   Create all of the columns for the ListView if necessary
        CreateIdEditor

        '   Populate the sub-items in the list view.  The first time this form is
        '   shown these items need to be populated with valid values.  Each
        '   subsequent these items need to be repopulated just in case the values
        '   have been changed.
        For Each oNode In g_oModels.childNodes
            If oNode.nodeType = NODE_ELEMENT Then
                '   Find out which row this Model Type has been sorted into
                Set lsiRow = lvwSerialNumbers.FindItem(oNode.baseName)
                
                If etID_TYPE.eSERIAL_NUM <> m_eIdType Then
                    ' Update non-incrementing portion of serial number
                    lsiRow.SubItems(eITEM_STATIC_PRE) = oNode.selectSingleNode("IMEI").selectSingleNode("TAC").firstChild.Text
                    lsiRow.SubItems(eITEM_STATIC_POST) = oNode.selectSingleNode("IMEI").selectSingleNode("SV").firstChild.Text
                    
                    Set oAttributes = oNode.selectSingleNode("IMEI").selectSingleNode("SNR").Attributes
                Else
                    Set oAttributes = oNode.selectSingleNode("SerialNo").Attributes
                End If
                
                ' Update the incrementing serial number information
                lsiRow.SubItems(eITEM_START) = oAttributes.getNamedItem("Start").Text
                lsiRow.ListSubItems(eITEM_START).ForeColor = vbBlack
                lsiRow.SubItems(eITEM_END) = oAttributes.getNamedItem("End").Text
                lsiRow.ListSubItems(eITEM_END).ForeColor = vbBlack
                lsiRow.SubItems(eITEM_NEXT) = oAttributes.getNamedItem("Next").Text
                lsiRow.ListSubItems(eITEM_NEXT).ForeColor = vbBlack
            End If
        Next oNode
        m_bChangesMade = False
        m_bRefresh = False
    End If

End Sub

Private Sub Form_Load()

    '   Create the serial number editor
    Set m_fUpdateSerialNumber = New frmUpdateSerialNumber
    Load m_fUpdateSerialNumber
    
    '   Create all of the columns for the ListView
    CreateIdEditor

    m_bRefresh = True

End Sub

Private Sub Form_Terminate()
    m_bRefresh = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload m_fUpdateSerialNumber

End Sub

Private Sub lvwSerialNumbers_Click()

    EditSerialNumbers

End Sub

Private Sub lvwSerialNumbers_KeyPress(KeyAscii As Integer)

    If Chr(KeyAscii) = vbCr Then
        EditSerialNumbers
    End If

End Sub

Public Sub EditSerialNumbers(Optional szIMEI_InputMethod As String = "BY_MODEL")

    Dim listItemRow As ListItem
    Dim selectedListItem As String
    Dim lIdLength As Long
    Dim szIdFormat As String
    Dim oNode As IXMLDOMNode
    Dim vbResult As VbMsgBoxResult
    
    If etID_TYPE.eSERIAL_NUM = m_eIdType Then
        szIdFormat = SERIAL_NUM_FORMAT
'        lIdLength = szIdFormat.length
        lIdLength = Len(SERIAL_NUM_FORMAT)
    Else
        ' IMEI serial number
        szIdFormat = IMEI_NUM_FORMAT
'        lIdLength = szIdFormat.length
        lIdLength = Len(IMEI_NUM_FORMAT)
    End If

    If etID_TYPE.eIMEI_NUM = m_eIdType And szIMEI_InputMethod <> "BY_MODEL" Then
        Set oNode = g_oSettings.selectSingleNode("TestSystem").childNodes(g_iTestStationID)
        Set oNode = oNode.selectSingleNode("IMEI")
        m_fUpdateSerialNumber.txtStaticPre = oNode.selectSingleNode("TAC").Text
        m_fUpdateSerialNumber.txtStaticPre.MaxLength = Len(IMEI_NUM_TAC_FORMAT)
        m_fUpdateSerialNumber.Caption = "Update IMEI Number"

        m_fUpdateSerialNumber.txtStaticPost = oNode.selectSingleNode("SV").Text
        m_fUpdateSerialNumber.txtStaticPost.MaxLength = Len(IMEI_NUM_SV_FORMAT)
        
        Set oNode = oNode.selectSingleNode("SNR")
        
        m_fUpdateSerialNumber.txtStartingNo = oNode.Attributes.getNamedItem("Start").Text
        m_fUpdateSerialNumber.txtStartingNo.MaxLength = lIdLength

        m_fUpdateSerialNumber.txtEndingNo = oNode.Attributes.getNamedItem("End").Text
        m_fUpdateSerialNumber.txtEndingNo.MaxLength = lIdLength

        m_fUpdateSerialNumber.txtNextNo = oNode.Attributes.getNamedItem("Next").Text
        m_fUpdateSerialNumber.txtNextNo.MaxLength = lIdLength

        m_fUpdateSerialNumber.txtStaticPre.Enabled = True
        m_fUpdateSerialNumber.txtStaticPre.Visible = True
        m_fUpdateSerialNumber.lblStaticPre.Visible = True

        m_fUpdateSerialNumber.txtStaticPost.Enabled = True
        m_fUpdateSerialNumber.txtStaticPost.Visible = True
        m_fUpdateSerialNumber.lblStaticPost.Visible = True
        
        m_fUpdateSerialNumber.Show vbModal

    
        If m_fUpdateSerialNumber.m_bUpdate Then
            vbResult = MsgBox("Serial numbers WILL be updated back to the XML configuration file!" & vbCrLf & vbCrLf & "Are you certain you want to proceed?", vbYesNoCancel Or vbCritical)
            
        End If
        If vbResult = vbYes Then
            oNode.Attributes.getNamedItem("Start").Text = m_fUpdateSerialNumber.txtStartingNo
            oNode.Attributes.getNamedItem("End").Text = m_fUpdateSerialNumber.txtEndingNo
            oNode.Attributes.getNamedItem("Next").Text = m_fUpdateSerialNumber.txtNextNo
            
            Set oNode = oNode.parentNode
            oNode.selectSingleNode("TAC").Text = m_fUpdateSerialNumber.txtStaticPre
            oNode.selectSingleNode("SV").Text = m_fUpdateSerialNumber.txtStaticPost
        End If
    Else
        For Each listItemRow In lvwSerialNumbers.ListItems
            If listItemRow.Selected Then
                '   Populate up the serial number editor dialog box before showing
                '   it.
                If etID_TYPE.eSERIAL_NUM <> m_eIdType Then
                    m_fUpdateSerialNumber.txtStaticPre = listItemRow.SubItems(eITEM_STATIC_PRE)
                    m_fUpdateSerialNumber.txtStaticPre.MaxLength = Len(IMEI_NUM_TAC_FORMAT)
                End If
                m_fUpdateSerialNumber.Caption = listItemRow.Text
    
                m_fUpdateSerialNumber.txtStartingNo = listItemRow.SubItems(eITEM_START)
                m_fUpdateSerialNumber.txtStartingNo.MaxLength = lIdLength
    
                m_fUpdateSerialNumber.txtEndingNo = listItemRow.SubItems(eITEM_END)
                m_fUpdateSerialNumber.txtEndingNo.MaxLength = lIdLength
    
                m_fUpdateSerialNumber.txtNextNo = listItemRow.SubItems(eITEM_NEXT)
                m_fUpdateSerialNumber.txtNextNo.MaxLength = lIdLength
    
                If etID_TYPE.eSERIAL_NUM <> m_eIdType Then
                    m_fUpdateSerialNumber.txtStaticPost = listItemRow.SubItems(eITEM_STATIC_POST)
                    m_fUpdateSerialNumber.txtStaticPost.MaxLength = Len(IMEI_NUM_SV_FORMAT)
                End If

            Dim b_ShowPrePost As Boolean
                If m_eIdType < etID_TYPE.eIMEI_NUM Then
                    b_ShowPrePost = False
                Else
                    b_ShowPrePost = True
                End If
    
                m_fUpdateSerialNumber.txtStaticPre.Enabled = b_ShowPrePost
                m_fUpdateSerialNumber.txtStaticPre.Visible = b_ShowPrePost
                m_fUpdateSerialNumber.lblStaticPre.Visible = b_ShowPrePost
    
                m_fUpdateSerialNumber.txtStaticPost.Enabled = b_ShowPrePost
                m_fUpdateSerialNumber.txtStaticPost.Visible = b_ShowPrePost
                m_fUpdateSerialNumber.lblStaticPost.Visible = b_ShowPrePost
    
                selectedListItem = listItemRow.Text
                DoEvents
                
                m_fUpdateSerialNumber.Show vbModal
    
                If m_fUpdateSerialNumber.m_bUpdate Then
                    m_bChangesMade = True
                    ' m_bRefresh = False
                    If listItemRow.SubItems(eITEM_STATIC_PRE) <> m_fUpdateSerialNumber.m_szStaticPre Then
                        listItemRow.SubItems(eITEM_STATIC_PRE) = m_fUpdateSerialNumber.m_szStaticPre
                        listItemRow.ListSubItems(eITEM_STATIC_PRE).ForeColor = vbRed
                    End If
    
                    If listItemRow.SubItems(eITEM_START) <> m_fUpdateSerialNumber.m_szStartingNo Then
                        listItemRow.SubItems(eITEM_START) = m_fUpdateSerialNumber.m_szStartingNo
                        listItemRow.ListSubItems(eITEM_START).ForeColor = vbRed
                    End If
    
                    If listItemRow.SubItems(eITEM_END) <> m_fUpdateSerialNumber.m_szEndingNo Then
                        listItemRow.SubItems(eITEM_END) = m_fUpdateSerialNumber.m_szEndingNo
                        listItemRow.ListSubItems(eITEM_END).ForeColor = vbRed
                    End If
    
                    If listItemRow.SubItems(eITEM_NEXT) <> m_fUpdateSerialNumber.m_szNextNo Then
                        listItemRow.SubItems(eITEM_NEXT) = m_fUpdateSerialNumber.m_szNextNo
                        listItemRow.ListSubItems(eITEM_NEXT).ForeColor = vbRed
                    End If
    
                    If listItemRow.SubItems(eITEM_STATIC_POST) <> m_fUpdateSerialNumber.m_szStaticPost Then
                        listItemRow.SubItems(eITEM_STATIC_POST) = m_fUpdateSerialNumber.m_szStaticPost
                        listItemRow.ListSubItems(eITEM_STATIC_POST).ForeColor = vbRed
                    End If
                End If
            Exit For
        End If
        Next listItemRow
    End If
End Sub

Private Sub CreateIdEditor()
    Dim chrHeader As ColumnHeader
    Dim i As Integer
    If lvwSerialNumbers.ListItems.Count > 0 Then
        For i = lvwSerialNumbers.ListItems.Count To 1 Step -1
            lvwSerialNumbers.ListItems.Remove (i)
            ' ltmRow.ListSubItems.Add , , ""
        Next i
        For i = lvwSerialNumbers.ColumnHeaders.Count To 1 Step -1
            lvwSerialNumbers.ColumnHeaders.Remove (i)
        Next i
    End If
    
        '   Create all of the columns for the ListView of Model Types.
        
        Set chrHeader = lvwSerialNumbers.ColumnHeaders.Add()
        chrHeader.Text = "Model Type"
        chrHeader.Width = lvwSerialNumbers.Width \ 5 - 20
        ' chrHeader.Alignment = lvwColumnCenter
        
        Set chrHeader = lvwSerialNumbers.ColumnHeaders.Add()
        If eIMEI_NUM <> m_eIdType Then
            chrHeader.Text = " "
            chrHeader.Width = 1
        Else
            chrHeader.Text = "TAC"
            chrHeader.Width = lvwSerialNumbers.Width \ 5 - 20
            ' Len(IMEI_NUM_TAC_FORMAT) + 2
        End If
        chrHeader.Alignment = lvwColumnCenter
        
        Set chrHeader = lvwSerialNumbers.ColumnHeaders.Add()
        chrHeader.Text = "Start"
        chrHeader.Width = lvwSerialNumbers.Width \ 5 - 20
        chrHeader.Alignment = lvwColumnCenter
        
        Set chrHeader = lvwSerialNumbers.ColumnHeaders.Add()
        chrHeader.Text = "End"
        chrHeader.Width = lvwSerialNumbers.Width \ 5 - 20
        chrHeader.Alignment = lvwColumnCenter
        
        Set chrHeader = lvwSerialNumbers.ColumnHeaders.Add()
        chrHeader.Text = "Next"
        chrHeader.Width = lvwSerialNumbers.Width \ 5 - 20
        chrHeader.Alignment = lvwColumnCenter
        
        Set chrHeader = lvwSerialNumbers.ColumnHeaders.Add()
        If eIMEI_NUM <> m_eIdType Then
            chrHeader.Text = " "
            chrHeader.Width = 1
        Else
            chrHeader.Text = "SV"
            chrHeader.Width = lvwSerialNumbers.Width \ 5 - 20
            ' Len(IMEI_NUM_SV_FORMAT) + 2
        End If
        chrHeader.Alignment = lvwColumnCenter
        
        Dim ltmRow As ListItem
        
        Dim oNode As IXMLDOMNode
        '   Create one row for each Model Type defined in the XML file.
        For Each oNode In g_oModels.childNodes
            If oNode.nodeType = NODE_ELEMENT Then
                Set ltmRow = lvwSerialNumbers.ListItems.Add
                ltmRow.Text = oNode.baseName
                '   Create blank spaces for all of the sub-items in this row.
                '   These blanks will get populated when the form is activated
                '   (i.e. when it is shown in modal form).
                For i = 1 To eITEM_NO_OF_ITEMS
                    ltmRow.ListSubItems.Add , , ""
                Next i
            End If
        Next oNode
End Sub
