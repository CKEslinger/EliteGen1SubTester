Attribute VB_Name = "Globals"
Public iFileNo As Integer
Public Const ALLDATA_LOG As String = "alldata_log.txt"
Public AllDataLogName As String
Option Explicit
Public Const DATA_COUNT = 8192
Public Const DEFAULT_DATA_DIR = "C:\PassTimeData\"
Public Const LABEL_SETTINGS_FN = "LabelConfig_ZPL.txt"
Public Const LPTX_OUTPUT_FN = "prnsend.txt"
Public Const TEST_LOG_FN = "test_log.txt"
Public Const DOS_EXEC_BATCH_FN = "DosExec.bat"
Public Const PIC_PROG_BATCH_FN = "ProgPic.bat"

Public Const MAX_NUM_OF_LABELS = 5

' DAQ Card Analog In Assignments
Public Const DUT_RELAY_VOLATAGE = 0
Public Const PS1_VOLTAGE = 1
Public Const PS2_VOLTAGE = 2
Public Const LED_GRN_VOLTAGE = 3
Public Const LED_RED_VOLTAGE = 4

' Data from XML file
Public g_oEliteTester As New DOMDocument
Public g_oModels As IXMLDOMNode
Public g_oVersion As IXMLDOMNode
Public g_oSettings As IXMLDOMNode
Public g_oSIMs As IXMLDOMNode
Public g_oTests As IXMLDOMNode
Public g_oSelectedSIM As IXMLDOMNode
Public g_oSelectedModel As IXMLDOMNode
Public g_oAddTests As IXMLDOMNode
Public g_iNumOfTests As Integer
Public g_iModemRegWait As Integer
Public g_iModemCSQ_Wait As Integer
Public g_iGSM_Attenuation As Integer
Public g_iTestStationID As Integer

Public Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

'Properties
Private mPassedParts As Integer
Private mFailedParts As Integer

'Test Result Properties
Private mScannedBarcode As Long
Private mBarcodeVerify As Boolean
Private mMicInput As Double

Private m_ComputerName As String

Private Function GetCompName() As String
    Dim retVal As Long

    'Create a string buffer for the computer name
    Dim strCompName As String
    strCompName = Space$(255)
    
    'Retrieve the Computer name
    retVal = GetComputerName(strCompName, 255)
    
    'Remove the trailing null character from the string
    GetCompName = Left$(strCompName, InStr(strCompName, vbNullChar) - 1)
End Function

Public Property Get ComputerName() As String
   If "" = m_ComputerName Then
        m_ComputerName = GetCompName()
   End If
   ComputerName = m_ComputerName
End Property

Public Property Get MicInput() As Double
    MicInput = mMicInput
End Property

Public Property Let MicInput(value As Double)
    mMicInput = value
End Property

Public Property Get PassedParts() As Integer

    PassedParts = mPassedParts

End Property

Public Property Let PassedParts(value As Integer)

    mPassedParts = value

End Property

Public Property Get FailedParts() As Integer

    FailedParts = mFailedParts

End Property

Public Property Let FailedParts(value As Integer)

    mFailedParts = value

End Property

Public Property Get ScannedBarcode() As Long
    ScannedBarcode = mScannedBarcode
End Property

Public Property Let ScannedBarcode(value As Long)
    mScannedBarcode = value
End Property

Public Property Get BarcodeVerify() As Boolean
    BarcodeVerify = mBarcodeVerify
End Property

Public Property Let BarcodeVerify(value As Boolean)
    mBarcodeVerify = value
End Property

