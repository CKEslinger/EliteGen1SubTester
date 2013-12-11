Attribute VB_Name = "Config"
Option Explicit
'   These are the various status' of the COM links
Public Enum etCOM_DEVICES
    etCOM_DEV_PS
    etCOM_DEV_DUT
    etCOM_DEV_GSM
    etCOM_DEV_GPS
End Enum

Public Enum etCOM_SETTINGS
    etCOM_SETTING_PORT
    etCOM_SETTING_BAUD
    etCOM_SETTING_PARITY
    etCOM_SETTING_DATA_BITS
    etCOM_SETTING_STOP_BITS
    etCOM_SETTING_FLOW_CTRL
End Enum

Private Const XML_VERSION_MAJOR As Double = 2
Private Const XML_VERSION_MINOR As Double = 92

Private Function IsTagMissing(ByRef szErrLog As String, szTagLvl As Integer, szNode As String, Optional szTag As String) As Boolean
    ' g_oSettings
    Dim xmlNode As IXMLDOMNode
    Dim i As Integer
    
    IsTagMissing = True
    Set xmlNode = g_oEliteTester.documentElement.selectSingleNode(szNode)
    
    Select Case Val(szTagLvl)
    Case 2
        If xmlNode Is Nothing Then
            szErrLog = "XML file missing the <" & szNode & "> node"
        End If
    Case 3
        If xmlNode.selectSingleNode(szTag) Is Nothing Then
            szErrLog = "XML file missing the <" & xmlNode.baseName & "><" & szTag & "> tag"
        End If
    Case Else
        IsTagMissing = False
    End Select
    
End Function

Private Function SetNode(ByRef szErrLog As String, ByRef xmlNode As IXMLDOMNode, szNode() As String, szNodeText() As String) As Boolean
    SetNode = True
    Dim xmlDOMNode As IXMLDOMNode
    Dim cnt As Integer
    
    cnt = 0
    
    If xmlNode Is Nothing Then
        ' Setting the relative top level node
        Set xmlNode = g_oEliteTester.documentElement.selectSingleNode(szNode(0))
        If xmlNode Is Nothing Then
            szErrLog = "XML file missing the <" & szNode(0) & "> node"
            SetNode = False
        End If
    ElseIf False = xmlNode.hasChildNodes Then
        szErrLog = "XML file missing the <" & xmlNode.baseName & "><" & szNode(cnt) & "> node"
        SetNode = False
    Else
        ' Verify each child's base name matches the names listed in the array
        For Each xmlDOMNode In xmlNode.childNodes
            If cnt > UBound(szNode) Then
                szErrLog = "XML file contains unused node <" & xmlDOMNode.baseName & ">"
                SetNode = False
                Exit For
            ElseIf NODE_COMMENT = xmlDOMNode.nodeType Then
                'Skip
            ElseIf xmlDOMNode.baseName <> szNode(cnt) Then
                ' Technically the node could exist but is not in same indexed order of array
                szErrLog = "XML file missing the <" & xmlNode.baseName & "><" & szNode(cnt) & "> node"
                SetNode = False
                Exit For
            End If
            If NODE_COMMENT <> xmlDOMNode.nodeType Then
                If "" <> xmlDOMNode.Text Then
                    ' Node has child node text
                    szNodeText(cnt) = xmlDOMNode.Text
                Else
                    szNodeText(cnt) = ""
                End If
                cnt = cnt + 1
            End If
        Next xmlDOMNode
    End If
End Function

Private Function SetAttribute(ByRef szErrLog As String, xmlNode As IXMLDOMNode, szAttrib() As String, szAttribText() As String) As Boolean
    Dim xmlDOMAttrib As IXMLDOMAttribute
    Dim cnt As Integer
    SetAttribute = True
    
    cnt = 0
    
    If xmlNode Is Nothing Then
        ' Setting the relative top level node
        Set xmlNode = g_oEliteTester.documentElement.selectSingleNode(szAttrib(0))
        If xmlNode Is Nothing Then
            szErrLog = "XML file missing the <" & szAttrib(0) & "> node"
            SetAttribute = False
        End If
    ElseIf True = xmlNode.hasChildNodes Then
        szErrLog = "XML file: node <" & xmlNode.baseName & "> had child nodes prior to attribute <" & szAttrib(cnt) & ">"
        SetAttribute = False
    ElseIf 0 = xmlNode.Attributes.length Then
        szErrLog = "XML file missing the <" & xmlNode.baseName & "><" & szAttrib(cnt) & "> attribute"
        SetAttribute = False
    Else
        ' Verify each child's base name matches the names listed in the array
        For Each xmlDOMAttrib In xmlNode.Attributes
            If cnt > UBound(szAttrib) Then
                szErrLog = "XML file contains unused node <" & xmlDOMAttrib.baseName & ">"
                SetAttribute = False
                Exit For
            ElseIf xmlDOMAttrib.baseName <> szAttrib(cnt) Then
                ' Technically the node could exist but is not in same indexed order of array
                szErrLog = "XML file missing the <" & xmlNode.baseName & "><" & szAttrib(cnt) & "> node"
                SetAttribute = False
                Exit For
            Else
                szAttribText(cnt) = xmlDOMAttrib.Text
            End If
            cnt = cnt + 1
        Next xmlDOMAttrib
    End If
    If cnt <= UBound(szAttrib) Then
        ' Verify all specified attributes were accounted for
        If "" <> szAttrib(cnt) Then
            If xmlNode Is Nothing Then
                szErrLog = "Bad XML file processing <" & szAttrib(cnt) & "> node"
            Else
                szErrLog = "XML file missing the <" & xmlNode.baseName & "><" & szAttrib(cnt) & "> node"
            End If
            SetAttribute = False
        End If
    End If
End Function

Public Function OpenModelTypeFile(Optional sFileTitle As String = "ModelTypes.xml", Optional sFileName As String = ".\ModelTypes.xml") As String

    Dim oTag As IXMLDOMNode
    Dim oItem As IXMLDOMAttribute
    Dim bCOMPorts As Boolean
    Dim szErrLog As String
    Dim szErrorMsg As String
    Dim szElementName As String
    Dim oAttribute1 As IXMLDOMAttribute
    Dim oAttribute2 As IXMLDOMAttribute
    Dim oAttribute3 As IXMLDOMAttribute
    Dim oAttribute4 As IXMLDOMAttribute
    Dim oAttribute(0 To 3) As IXMLDOMAttribute
    Dim sAttribute(0 To 3) As String
    Dim iIndex As Integer
    Dim iCnt As Integer
    
    Dim oNode As IXMLDOMNode
    Dim szNode(15) As String
    Dim szNodeText(15) As String
    Dim szAttrib(15) As String
    Dim szAttribText(15) As String
    Dim oChildNode As IXMLDOMNode
    Dim szChildNode(15) As String
    Dim szChildNodeText(15) As String


    '   Open the Model Type XML configuration file
    g_oEliteTester.async = False
    If Not g_oEliteTester.Load(sFileName) Then
        OpenModelTypeFile = "The XML configuration file '" & sFileTitle & "' in folder '" & CurDir & "' is either missing or is corrupted!"
        Exit Function
    End If

    '   Validate its contents
    If g_oEliteTester.documentElement.tagName <> "Elite" Then
        OpenModelTypeFile = "XML file missing the root 'Elite' tag"
        Exit Function
    End If
    
    '
    '   Ensure that a <Version> tag has been defined
    '
    szNode(0) = "Version"
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    Set g_oVersion = g_oEliteTester.documentElement.selectSingleNode(szNode(0))
    
    ' Verify the version number nodes exist
    szNode(0) = "XML_File"
    szNode(1) = "MinAppVer"
    szNode(2) = "MaxAppVer"
    szNode(3) = ""
    
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
        
    Dim verString
    
    ' Verify the xml version number attibute exist
    Set oNode = g_oVersion.selectSingleNode(szNode(0))
    szAttrib(0) = "Major"
    szAttrib(1) = "Minor"
    szAttrib(2) = "Revision"
    szAttrib(3) = ""
    
    If False = SetAttribute(szErrLog, oNode, szAttrib, szAttribText()) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    ' Check that the XML version number in the code is greater than or equal to the actual XML file version
    If XML_VERSION_MAJOR >= CDbl(Val(szAttribText(0))) Then
        If XML_VERSION_MINOR >= CDbl(Val(szAttribText(1))) Then
            verString = "No Error"
        End If
    End If
    If "No Error" <> verString Then
        OpenModelTypeFile = "XML file mismatch.  Was expecting version " & _
                    szAttribText(0) & "." & szAttribText(1) & " but found '" & _
                    XML_VERSION_MAJOR & "." & XML_VERSION_MINOR & "'"
        Exit Function
    End If
    
    For iCnt = 1 To 2
        ' Verify the Min and Max attibutes exist
        Set oNode = g_oVersion.selectSingleNode(szNode(iCnt))
        If False = SetAttribute(szErrLog, oNode, szAttrib, szAttribText()) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
        
        ' Check that the XML is compatible with the application version
        verString = szAttribText(0) & "." & szAttribText(1) & "." & szAttribText(2)
        verString = verString & " but version is " & App.Major & "." & App.Minor & "." & App.Revision
        If "MinAppVer" = oNode.baseName Then
            ' Check if the application version is greater than the MinAppVer in XML
            If App.Major >= CDbl(Val(szAttribText(0))) Then
                If App.Minor >= CDbl(Val(szAttribText(1))) Then
                    If App.Revision >= CDbl(Val(szAttribText(2))) Then
                        verString = "No Error"
                    End If
                End If
            End If
            If "No Error" <> verString Then
                OpenModelTypeFile = "Application version mismatch.  Was expecting at least version " & verString
                Exit Function
            End If
        Else
            ' Check if the application version is less than the MaxAppVer in XML
            If App.Major <= CDbl(Val(szAttribText(0))) Then
                If App.Minor <= CDbl(Val(szAttribText(1))) Then
                    If App.Revision <= CDbl(Val(szAttribText(2))) Then
                        verString = "No Error"
                    End If
                End If
            End If
            If "No Error" <> verString Then
                OpenModelTypeFile = "Application version mismatch.  Was expecting at most version " & verString
                Exit Function
            End If
        End If
    Next iCnt
    
    '
    '   Ensure that a <Tests> tag has been defined
    '
    Set oNode = Nothing
    szNode(0) = "Tests"
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    Set g_oTests = g_oEliteTester.documentElement.selectSingleNode(szNode(0))
    
    '   Verify that the <Tests> tag has been configured correctly with all
    '   of its requisite elements.  Remove all white spaces from the text
    '   of all the elements in the process.
    
    ' Verify the "Tests" nodes exist
    szNode(0) = "NumOfTests"
    szNode(1) = "Download"
    szNode(2) = "Configure"
    szNode(3) = "Wireless"
    szNode(4) = "Test"
    szNode(5) = ""
    
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    g_iNumOfTests = Val(RemoveWhiteSpaces(szNodeText(0)))
    
    '
    '   Ensure that a <AddTests> tag has been defined
    '
    Set oNode = Nothing
    szNode(0) = "AddTests"
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    Set g_oAddTests = g_oEliteTester.documentElement.selectSingleNode(szNode(0))
    
    ' Verify the "add tests" attibutes exist
    szAttrib(0) = "Description"
    szAttrib(1) = "Command"
    szAttrib(2) = "Result"
    szAttrib(3) = "Time"
    szAttrib(4) = ""
    
    iCnt = 1
    For Each oNode In g_oAddTests.childNodes
        If oNode.baseName <> "ADD" & iCnt Then
            OpenModelTypeFile = "XML file <ADD" & iCnt & "> tag in wrong format"
            Exit Function
        End If
        
        ' Set oNode = g_oAddTests.selectSingleNode("ADD" & iChildCnt)
        If False = SetAttribute(szErrLog, oNode, szAttrib, szAttribText()) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
    Next oNode
        
    '
    '   Ensure that a <Settings> tag has been defined
    '
    Set oNode = Nothing
    szNode(0) = "Settings"
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    Set g_oSettings = g_oEliteTester.documentElement.selectSingleNode(szNode(0))
    
    ' Verify the "Settings" nodes exist
    szNode(0) = "Company"
    szNode(1) = "Main"
    szNode(2) = "Operations"
    szNode(3) = "DataPath"
    szNode(4) = "ResultsLog"
    szNode(5) = "PrinterDir"
    szNode(6) = "COMPorts"
    szNode(7) = "BarCodes"
    szNode(8) = "GPSSimulator"
    szNode(9) = "Sound"
    szNode(10) = "CellSiteSimulator"
    szNode(11) = "TestSystem"
    szNode(12) = ""
    
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    '   Check out the <COMPorts> tag
    szChildNode(0) = "PS"
    szChildNode(1) = "Elite"
    szChildNode(2) = "CellSiteSimulator"
    szChildNode(3) = ""
    
    Set oChildNode = oNode.selectSingleNode(szNode(6))
    If False = SetNode(szErrLog, oChildNode, szChildNode, szChildNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    szAttrib(0) = "Port"
    szAttrib(1) = "Baud"
    szAttrib(2) = "DataBits"
    szAttrib(3) = "Parity"
    szAttrib(4) = "StopBits"
    szAttrib(5) = "FlowCtrl"
    szAttrib(6) = ""
    
    ' Verify the COM port attibutes exist
    For iCnt = 0 To 2
        If False = SetAttribute(szErrLog, oChildNode.selectSingleNode(szChildNode(iCnt)), szAttrib, szAttribText()) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
    Next iCnt
    
    '   Check out the <BarCodes> tag
    szAttrib(0) = "NumToPrint"
    szAttrib(1) = "NumToScan"
    szAttrib(2) = ""
    
    ' Verify the barcode settings attibutes exist
    If False = SetAttribute(szErrLog, oNode.selectSingleNode(szNode(7)), szAttrib, szAttribText()) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    '   Check out the <GPSSimulator> tag
    szAttrib(0) = "SVID"
    szAttrib(1) = "Lock"
    szAttrib(2) = "Min"
    szAttrib(3) = "Max"
    szAttrib(4) = "xMin"
    szAttrib(5) = "xMax"
    szAttrib(6) = ""
    
    ' Verify the GPS Simulator settings attibutes exist
    If False = SetAttribute(szErrLog, oNode.selectSingleNode(szNode(8)), szAttrib, szAttribText()) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    ' Verify the GPS Simulator threshold settings make sense
    Dim dMin As Double
    Dim dMax As Double
    
    For iCnt = 2 To 5 Step 2
        dMin = CDbl(szAttribText(iCnt))
        dMax = CDbl(szAttribText(iCnt + 1))
        If dMin > dMax Then
            OpenModelTypeFile = "The <" & szNode(8) & "><" & szAttribText(iCnt) & "> attribute MUST be less than or equal to the <" & szAttribText(iCnt + 1) & "> attribute" & vbCrLf
            Exit Function
        End If
    Next iCnt
    
    '   Check out the <Sound> tag
    szAttrib(0) = "Min"
    szAttrib(1) = ""
    
    ' Verify the audio settings attibutes exist
    If False = SetAttribute(szErrLog, oNode.selectSingleNode(szNode(9)), szAttrib, szAttribText()) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    '   Check out the <CellSiteSimulator> tag
    szAttrib(0) = "Station"
    szAttrib(1) = "Min"
    szAttrib(2) = "Max"
    szAttrib(3) = "Reg_Wait"
    szAttrib(4) = "CSQ_Wait"
    szAttrib(5) = ""
    
    ' Verify the callbox settings attibutes exist
    If False = SetAttribute(szErrLog, oNode.selectSingleNode(szNode(10)), szAttrib, szAttribText()) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    g_iModemRegWait = CInt(szAttribText(3))
    g_iModemCSQ_Wait = CInt(szAttribText(4))
    
    '   Check out the <TestSystem><TestStation> tag
    szChildNode(0) = "Computer"
    szChildNode(1) = "IMEI"
    szChildNode(2) = "Attenuation"
    szChildNode(3) = ""
    
    For Each oChildNode In oNode.selectSingleNode(szNode(11)).childNodes
        If False = SetNode(szErrLog, oChildNode, szChildNode, szChildNodeText) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
        ' TODO: Check out all the childnoes and attributes of each child node
    Next oChildNode
    
    '
    '   Ensure that a <SIMs> tag has been defined
    '
    Set oNode = Nothing
    szNode(0) = "SIMs"
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    Set g_oSIMs = g_oEliteTester.documentElement.selectSingleNode(szNode(0))
    
    ' Below is a verification of the SIM cards that exist as of 6/8/2013.
    ' Verification is commented out so that new SIM cards can be added to the
    ' XML file without having to modify the application
    ' Uncomment to run unit test during development
    
    ' Verify the SIM nodes exist
'    szNode(0) = "T-Mobile"
'    szNode(1) = "Numerex"
'    szNode(2) = "Rogers"
'    szNode(3) = ""
'
'    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
'        OpenModelTypeFile = szErrLog
'        Exit Function
'    End If
    
    szNode(0) = "APN"
    szNode(1) = "Server"
    szNode(2) = "SMS_Reply"
    szNode(3) = "SMS_Mode"
    szNode(4) = "Reset"
    szNode(5) = "Port"
    szNode(6) = "SrvcCntrAddr"
    szNode(7) = ""
    
    For Each oNode In g_oSIMs.childNodes
        '   Verify that all requesite elements are present
        If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If

        ' Verify the APN child nodes exist
        szAttrib(0) = "URL"
        szAttrib(1) = "Index"
        szAttrib(2) = ""
        
        For Each oChildNode In oNode.selectSingleNode(szNode(0)).childNodes
            If "PW" = oChildNode.baseName Then
                If Nothing Is oChildNode.Attributes.getNamedItem("Pass") Then
                    OpenModelTypeFile = "In <" & oNode.baseName & "><APN><PW> the tag 'PASS' is missing"
                    Exit Function
                End If
            Else
                ' Verify the url attibute exist for each APN
                If False = SetAttribute(szErrLog, oChildNode, szAttrib, szAttribText()) Then
                    OpenModelTypeFile = szErrLog
                    Exit Function
                End If
            End If
        Next oChildNode
        
        ' Verify the Servers child nodes exist
        szAttrib(0) = "IP"
        szAttrib(1) = "Index"
        szAttrib(2) = ""
        For Each oChildNode In oNode.selectSingleNode(szNode(1)).childNodes
            ' Verify the IP attibutes exist for each server
            If False = SetAttribute(szErrLog, oChildNode, szAttrib, szAttribText()) Then
                OpenModelTypeFile = szErrLog
                Exit Function
            End If
        Next oChildNode

        ' Verify the SMS Reply Address child nodes exist
        szAttrib(0) = "ADDR"
        szAttrib(1) = "Index"
        szAttrib(2) = ""
        For Each oChildNode In oNode.selectSingleNode(szNode(2)).childNodes
            ' Verify the ADDR attibute exist for each node
            If False = SetAttribute(szErrLog, oChildNode, szAttrib, szAttribText()) Then
                OpenModelTypeFile = szErrLog
                Exit Function
            End If
        Next oChildNode

        ' Verify the SMS Mode attributes exist
        szAttrib(0) = "HelloPktInterval"
        szAttrib(1) = "State"
        szAttrib(2) = ""
        For Each oChildNode In oNode.selectSingleNode(szNode(3)).childNodes
            ' Verify the SMS Mode attibutes exist for each node
            If False = SetAttribute(szErrLog, oChildNode, szAttrib, szAttribText()) Then
                OpenModelTypeFile = szErrLog
                Exit Function
            End If
        Next oChildNode

        ' Verify the Reset child nodes exist
        szAttrib(0) = "Name"
        szAttrib(1) = "Minutes"
        szAttrib(2) = "Index"
        szAttrib(3) = ""
        For Each oChildNode In oNode.selectSingleNode(szNode(4)).childNodes
            ' Verify the Reset attibutes exist for each node
            If False = SetAttribute(szErrLog, oChildNode, szAttrib, szAttribText()) Then
                OpenModelTypeFile = szErrLog
                Exit Function
            End If
        Next oChildNode

        ' Verify the Port child node has an attribute
        If "" = szNodeText(5) Then
            OpenModelTypeFile = "XML file missing the <" & oNode.baseName & "><" & szNode(5) & "> attribute"
            Exit Function
        End If

        ' Verify the Service Center Address attributes exist
        szAttrib(0) = "ADDR"
        szAttrib(1) = ""
        For Each oChildNode In oNode.selectSingleNode(szNode(6)).childNodes
            ' Verify the Service Center Address attibutes exist for each node
            If False = SetAttribute(szErrLog, oChildNode, szAttrib, szAttribText()) Then
                OpenModelTypeFile = szErrLog
                Exit Function
            End If
        Next oChildNode

    Next oNode
    
    '
    '   Ensure that a <Models> tag has been defined
    '
    Set oNode = Nothing
    szNode(0) = "Models"
    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
        OpenModelTypeFile = szErrLog
        Exit Function
    End If
    
    Set g_oModels = g_oEliteTester.documentElement.selectSingleNode(szNode(0))
    
    ' Below is a verification of the models that exist as of 6/8/2013.
    ' Verification is commented out so that new models can be added to the
    ' XML file without having to modify the application
    ' Uncomment to run unit test during development
    
    ' Verify the Models exist
'    szNode(0) = "PTE-II.N"
'    szNode(1) = "PTE-II.R"
'    szNode(2) = "PTC-II.R"
'    szNode(3) = "TRAX-II.N"
'    szNode(4) = "TRAX-II.R"
'    szNode(5) = "TRAX-II"
'    szNode(6) = "TRAX-II.CR"
'    szNode(7) = "PTE-II.S"
'    szNode(8) = "PTE-II.NS"
'    szNode(9) = "PTC-II.RS"
'    szNode(10) = ""
'
'    If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
'        OpenModelTypeFile = szErrLog
'        Exit Function
'    End If
    
' Removed following node from models on 6/16/2013 to avoid confusion of having
' unused XML code in file. Code to use "BY_MODEL" method of IMEI umber tracking
' still exists and if the following node is added to each model type then
' could still track by model type
'
'            <IMEI>
'                <ModemModel>BGS2</ModemModel>
'                <TAC>01307900</TAC>
'                <SV>01</SV>
'                <SNR Start="000000" End="999999" Next="000100"></SNR>
'            </IMEI>

    szNode(0) = "ProdNo"
    szNode(1) = "Firmware"
    szNode(2) = "SIMs"
    szNode(3) = "Tests"
    szNode(4) = "GPS"
    szNode(5) = "Security"
    szNode(6) = "SerialNo"
    szNode(7) = "Power"
    szNode(8) = ""
    
    For Each oNode In g_oModels.childNodes
        '   Verify that all requesite elements are present
        If False = SetNode(szErrLog, oNode, szNode, szNodeText) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
    
        ' Verify the Prod Number, Firmware, SIM, and Tests child nodes have an attribute
        For iCnt = 0 To 3
            If "" = szNodeText(iCnt) Then
                OpenModelTypeFile = "XML file missing the <" & oNode.baseName & "><" & szNode(iCnt) & "> attribute"
                Exit Function
            End If
        Next iCnt
    
        Set oTag = g_oSIMs.selectSingleNode(szNodeText(2))
        If oTag Is Nothing Then
            OpenModelTypeFile = "In <" & oNode.baseName & "><SIMs> the specified SIM card '" & szNodeText(2) & "' is invalid"
        End If

        ' Verify the GPS status child node attributes exist
        szAttrib(0) = "AlwaysOn"
        szAttrib(1) = ""
        If False = SetAttribute(szErrLog, oNode.selectSingleNode(szNode(4)), szAttrib, szAttribText()) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
    
        ' Verify the Security child node attributes exist
        szAttrib(0) = "Valet"
        szAttrib(1) = ""
        If False = SetAttribute(szErrLog, oNode.selectSingleNode(szNode(5)), szAttrib, szAttribText()) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
    
        ' Verify the Serial Number child node attributes exist
        szAttrib(0) = "Start"
        szAttrib(1) = "End"
        szAttrib(2) = "Next"
        szAttrib(3) = ""
        If False = SetAttribute(szErrLog, oNode.selectSingleNode(szNode(6)), szAttrib, szAttribText()) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
    
        dMin = CDbl(szAttribText(0))
        dMax = CDbl(szAttribText(1))
        If dMin > dMax Then
            OpenModelTypeFile = "<" & oNode.baseName & "><" & szNode(6) & "><" & szAttrib(0) & "> = " & szAttribText(0) & " MUST be less than or equal to <" & szAttrib(1) & "> = " & szAttribText(1) & " attribute" & vbCrLf
            Exit Function
        End If
    
        dMin = CDbl(szAttribText(0))
        dMax = CDbl(szAttribText(2))
        If dMin > dMax Then
            OpenModelTypeFile = "<" & oNode.baseName & "><" & szNode(6) & "><" & szAttrib(0) & "> = " & szAttribText(0) & " MUST be less than or equal to <" & szAttrib(2) & "> = " & szAttribText(2) & " attribute" & vbCrLf
            Exit Function
        End If
    
        dMin = CDbl(szAttribText(2))
        dMax = CDbl(szAttribText(1))
        If dMin > dMax Then
            OpenModelTypeFile = "<" & oNode.baseName & "><" & szNode(6) & "><" & szAttrib(2) & "> = " & szAttribText(2) & " MUST be less than or equal to <" & szAttrib(1) & "> = " & szAttribText(1) & " attribute" & vbCrLf
            Exit Function
        End If
        
        ' Verify the IMEI child node's child nodes and attributes exist
'        szChildNode(0) = "ModemModel"
'        szChildNode(1) = "TAC"
'        szChildNode(2) = "SV"
'        szChildNode(3) = "SNR"
'        szChildNode(4) = ""
'
'        Set oChildNode = oNode.selectSingleNode(szNode(7))
'        If False = SetNode(szErrLog, oChildNode, szChildNode, szChildNodeText) Then
'            OpenModelTypeFile = szErrLog
'            Exit Function
'        End If
'
'        ' Verify the ModemModel, TAC, and SV child nodes have an attribute
'        For iCnt = 0 To 2
'            If "" = szChildNodeText(iCnt) Then
'                OpenModelTypeFile = "XML file missing the <" & oNode.baseName & "><" & szNode(iCnt) & "> attribute"
'                Exit Function
'            End If
'        Next iCnt
'
'        szAttrib(0) = "Start"
'        szAttrib(1) = "End"
'        szAttrib(2) = "Next"
'        szAttrib(3) = ""
'
'        ' Verify the IMEI child node attributes exist
'
'        If False = SetAttribute(szErrLog, oChildNode.selectSingleNode(szChildNode(3)), szAttrib, szAttribText()) Then
'            OpenModelTypeFile = szErrLog
'            Exit Function
'        End If
'
'        dMin = CDbl(szAttribText(0))
'        dMax = CDbl(szAttribText(1))
'        If dMin > dMax Then
'            OpenModelTypeFile = "<" & oNode.baseName & "><" & szNode(8) & "><" & szAttrib(0) & "> = " & szAttribText(0) & " MUST be less than or equal to <" & szAttrib(1) & "> = " & szAttribText(1) & " attribute" & vbCrLf
'            Exit Function
'        End If
'
'        dMin = CDbl(szAttribText(0))
'        dMax = CDbl(szAttribText(2))
'        If dMin > dMax Then
'            OpenModelTypeFile = "<" & oNode.baseName & "><" & szNode(8) & "><" & szAttrib(0) & "> = " & szAttribText(0) & " MUST be less than or equal to <" & szAttrib(2) & "> = " & szAttribText(2) & " attribute" & vbCrLf
'            Exit Function
'        End If
'
'        dMin = CDbl(szAttribText(2))
'        dMax = CDbl(szAttribText(1))
'        If dMin > dMax Then
'            OpenModelTypeFile = "<" & oNode.baseName & "><" & szNode(8) & "><" & szAttrib(2) & "> = " & szAttribText(2) & " MUST be less than or equal to <" & szAttrib(1) & "> = " & szAttribText(1) & " attribute" & vbCrLf
'            Exit Function
'        End If
        
        ' Verify the Power child node attributes exist
        szChildNode(0) = "Nominal_Current"
        szChildNode(1) = "Inrush_Current"
        szChildNode(2) = "Nominal_Voltage"
        szChildNode(3) = "PWR_Off_Voltage"
        szChildNode(4) = ""
        
        Set oChildNode = oNode.selectSingleNode(szNode(7))
        If False = SetNode(szErrLog, oChildNode, szChildNode, szChildNodeText) Then
            OpenModelTypeFile = szErrLog
            Exit Function
        End If
        
        szAttrib(0) = "Units"
        szAttrib(1) = "Avg"
        szAttrib(2) = "Min"
        szAttrib(3) = "Max"
        szAttrib(4) = "Wait"
        szAttrib(5) = ""
        
        For iCnt = 0 To 3
            If False = SetAttribute(szErrLog, oChildNode.selectSingleNode(szChildNode(iCnt)), szAttrib, szAttribText()) Then
                OpenModelTypeFile = szErrLog
                Exit Function
            End If
        
            dMin = CDbl(szAttribText(2))
            dMax = CDbl(szAttribText(3))
            If dMin > dMax Then
                OpenModelTypeFile = "The <" & szChildNode(iCnt) & "><" & szAttribText(3) & "> attribute MUST be less than or equal to the <" & szAttribText(2) & "> attribute" & vbCrLf
                Exit Function
            End If
        Next iCnt
        
    Next oNode
End Function

Private Function XMLTagError(szTag As String, szElement As String, Optional bElement As Boolean = True) As String

    XMLTagError = "The <" & szTag & "> tag is missing the <" & szElement
    If bElement Then
        XMLTagError = XMLTagError + "> element" & vbCrLf
    Else
        XMLTagError = XMLTagError + "> attribute" & vbCrLf
    End If

End Function

Private Sub XMLFormatError(szBasename As String, szElement As String)

    MsgBox "XML format error:" & vbCrLf & vbCrLf & _
           "The Model profile '" & szBasename & "' is missing the '" & szElement & "' element", vbCritical

End Sub

Private Function RemoveWhiteSpaces(szInput As String) As String

    Dim i As Integer
    Dim szChar As String

    RemoveWhiteSpaces = ""
    For i = 1 To Len(szInput)
        szChar = Mid(szInput, i, 1)
        If szChar <> vbCr And szChar <> vbLf And szChar <> vbTab And szChar <> " " Then
            RemoveWhiteSpaces = RemoveWhiteSpaces + szChar
        End If
    Next i

End Function

Public Function GetComSettings(device As etCOM_DEVICES, setting As etCOM_SETTINGS) As String
    Dim oDevSettings As IXMLDOMNode
    Dim oDevice As IXMLDOMNode
    Dim oTag As IXMLDOMNode

    Set oDevSettings = g_oEliteTester.documentElement.selectSingleNode("Settings")
    Set oTag = oDevSettings.selectSingleNode("COMPorts")
    
    Select Case device
    Case etCOM_DEVICES.etCOM_DEV_PS
        Set oDevice = oTag.selectSingleNode("PS")
    Case etCOM_DEVICES.etCOM_DEV_DUT
        Set oDevice = oTag.selectSingleNode("Elite")
    Case etCOM_DEVICES.etCOM_DEV_GSM
        Set oDevice = oTag.selectSingleNode("CellSiteSimulator")
    Case Else
    End Select

    Select Case setting
    Case etCOM_SETTINGS.etCOM_SETTING_PORT
        GetComSettings = oDevice.Attributes.getNamedItem("Port").Text
    Case etCOM_SETTINGS.etCOM_SETTING_BAUD
        GetComSettings = oDevice.Attributes.getNamedItem("Baud").Text
    Case etCOM_SETTINGS.etCOM_SETTING_PARITY
        GetComSettings = oDevice.Attributes.getNamedItem("Parity").Text
    Case etCOM_SETTINGS.etCOM_SETTING_DATA_BITS
        GetComSettings = oDevice.Attributes.getNamedItem("DataBits").Text
    Case etCOM_SETTINGS.etCOM_SETTING_STOP_BITS
        GetComSettings = oDevice.Attributes.getNamedItem("StopBits").Text
    Case etCOM_SETTINGS.etCOM_SETTING_FLOW_CTRL
        GetComSettings = oDevice.Attributes.getNamedItem("FlowCtrl").Text
    Case Else
    End Select

End Function

