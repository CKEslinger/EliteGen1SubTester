VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fso1 As FileSystemObject


Dim sSavCfg_ZPL_Cmd() As Variant
Dim sTmpCfg_ZPL_Cmd() As Variant
Dim sBar_ZPL_Cmd() As Variant
Dim sTxt_ZPL_Cmd() As Variant

Dim SavCfg(8) As String
Dim TmpCfg(5) As String
Dim Bar(2, 4) As String
Dim Txt(3, 4) As String
    
Private Sub InitZplCmds()

    sSavCfg_ZPL_Cmd() = Array("MM", "MN", "MT", "LH", "JM", "PR", "MD", "JU")
    sTmpCfg_ZPL_Cmd() = Array("CI", "LH", "LL", "PW", "PQ")
    sBar_ZPL_Cmd() = Array("BY", "FT", "B3", "FD")
    sTxt_ZPL_Cmd() = Array("FT", "A", "FH", "FD")

End Sub

Public Sub SetPrinter()
  Dim prnPrinter

  For Each prnPrinter In Printers
    If prnPrinter.DeviceName = "Zebra Text" Then
       Set Printer = prnPrinter
       Exit For
    End If
  Next
  
End Sub

Public Sub PrintScan(lptX As Integer)
    Dim i As Integer
   
    For i = 0 To (Globals.MAX_NUM_OF_LABELS - 1)
        If i < g_iNumBarcodeScans Then
            PrintBarcode lptX, 1
            g_fSerialNumber.Show vbModal

'            If Globals.BarcodeVerify Then
'                g_fMainTester.LogMessage ("Barcode and Serial Number Match. Apply Barcode To DUT.")
'            Else
'                g_fMainTester.LogMessage ("Barcode and Serial Number Do Not Match.")
'            End If
        Else
            If i < g_iNumBarcodePrints Then
                PrintBarcode lptX, g_iNumBarcodePrints - i
            End If
            Exit For
        End If
    Next i

End Sub

Public Sub PrintBarcode(lptX As Integer, numPrints As Integer)
    Dim i As Integer
    Dim j As Integer
    
    Dim ts As TextStream
    Dim strTemp As String
    
    Dim SerNum As String
    Dim ImsiNum As String
    Dim IccidNum As String
    
    Dim lptSettingsFN As String
    Dim lptOutputFN As String
    Dim lptBatchFN As String
    
    InitZplCmds
    
    SerNum = Format(g_lNextSerialNumber, "0000000")
    ImsiNum = Format(g_szIMSI, "000000000000000")
    IccidNum = Format(g_szCCID, "00000000000000000000")
    
    'ImsiNum = "000000000000000"
    'IccidNum = "00000000000000000000"

    lptSettingsFN = LABEL_SETTINGS_FN
    lptOutputFN = LPTX_OUTPUT_FN
    lptBatchFN = DOS_EXEC_BATCH_FN

'MsgBox "Checking if config file" & lptSettingsFN & " exists...", vbInformation, "Debug"
    If fso1.FileExists(lptSettingsFN) Then
    

        Set ts = fso1.OpenTextFile(lptSettingsFN)
'MsgBox "Opened config file " & lptSettingsFN, vbInformation, "Debug"
        
        'Read ZPL settings from text file
        strTemp = ts.ReadLine                'Tag "Save Config:"
        
        For i = 0 To (UBound(SavCfg) - 1)
            SavCfg(i) = ParseString(ts.ReadLine)
        Next i
        
        strTemp = ts.ReadLine                'Blank Line
        strTemp = ts.ReadLine                'Tag "Temp Config:"
        
        For i = 0 To (UBound(TmpCfg) - 1)
            TmpCfg(i) = ParseString(ts.ReadLine)
            If sTmpCfg_ZPL_Cmd(i) = "PQ" Then
                'strTemp = Left(TmpCfg(i), InStr(1, TmpCfg(i), ","))
                
                TmpCfg(i) = numPrints & Mid(TmpCfg(i), InStr(1, TmpCfg(i), ","))
            End If
        Next i
        
        For i = 0 To (UBound(Bar, 1) - 1)
        strTemp = ts.ReadLine                'Blank Line
            strTemp = ts.ReadLine            'Tag "BarcodeX:"
            For j = 0 To (UBound(Bar, 2) - 1)
                Bar(i, j) = ParseString(ts.ReadLine)
            Next j
        Next i
        
        For i = 0 To (UBound(Txt, 1) - 1)
        strTemp = ts.ReadLine                'Blank Line
            strTemp = ts.ReadLine            'Tag "TextX:"
            For j = 0 To (UBound(Txt, 2) - 1)
                Txt(i, j) = ParseString(ts.ReadLine)
            Next j
        Next i
    
        ts.Close
    Else
        MsgBox lptSettingsFN, vbCritical
        Exit Sub
    End If
    
    
'MsgBox "Checking if output file " & lptOutputFN & " exists...", vbInformation, "Debug"
    If fso1.FileExists(lptOutputFN) Then
        Set ts = fso1.OpenTextFile(lptOutputFN, ForWriting)
    Else
        Set ts = fso1.CreateTextFile(lptOutputFN)
    End If
'MsgBox "Opened output file " & lptOutputFN, vbInformation, "Debug"
    
    ' Build the Saved Configuration Commands
    strTemp = "^XA"
    
    'Save Config Commands
    For i = 0 To (UBound(SavCfg) - 1)
        strTemp = strTemp & "^" & sSavCfg_ZPL_Cmd(i) & SavCfg(i)
    Next i
    strTemp = strTemp & "^XZ"
    
    ts.WriteLine strTemp
    
    strTemp = "^XA"
    
    'Temp Config Commands
    For i = 0 To (UBound(TmpCfg) - 1)
        strTemp = strTemp & "^" & sTmpCfg_ZPL_Cmd(i) & TmpCfg(i)
    Next i
    ts.WriteLine strTemp
    
    'Barcode Commands
    For i = 0 To (UBound(Bar, 1) - 1)
        strTemp = ""
        For j = 0 To (UBound(Bar, 2) - 1)
            If sBar_ZPL_Cmd(j) = "FD" Then
                ts.WriteLine strTemp
                strTemp = "^FD"
                If i = 0 Then
                    strTemp = strTemp & IccidNum
                Else
                    strTemp = strTemp & SerNum
                End If
                strTemp = strTemp & "^FS"
    Else
                strTemp = strTemp & "^" & sBar_ZPL_Cmd(j) & Bar(i, j)
    End If
        Next j
        ts.WriteLine strTemp
    Next i
    
    'Text Commands
    For i = 0 To (UBound(Txt, 1) - 1)
        strTemp = ""
        For j = 0 To (UBound(Txt, 2) - 1)
            strTemp = strTemp & "^" & sTxt_ZPL_Cmd(j) & Txt(i, j)
            If sTxt_ZPL_Cmd(j) = "FD" Then
                If i = 0 Then
                    strTemp = strTemp & ImsiNum
                ElseIf i = 1 Then
                    strTemp = strTemp & SerNum
    Else
                    strTemp = strTemp & IccidNum
    End If
                strTemp = strTemp & "^FS"
            End If
        Next j
        ts.WriteLine strTemp
    Next i
    
    ts.WriteLine "^XZ"
    
    ts.Close
    
'MsgBox "Checking if batch file " & lptOutputFN & " exists...", vbInformation, "Debug"
    If fso1.FileExists(lptBatchFN) Then
        Shell (Chr(34) & lptBatchFN & Chr(34) & " COPY " & Chr(34) & lptOutputFN & Chr(34) & " lpt" & lptX & ":")
    Else
'MsgBox "Failed to find batch file " & lptOutputFN, vbInformation, "Debug"
    End If

End Sub

Private Function ParseString(strData As String) As String
    Dim pos As Integer
    
    pos = InStr(1, strData, ":")
    ParseString = Mid(strData, pos + 1)
End Function

Private Sub Class_Initialize()
    Set fso1 = New FileSystemObject   ''' Create the file handling object
End Sub

Private Sub Class_Terminate()
    Set fso1 = Nothing
End Sub

