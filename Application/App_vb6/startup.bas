Attribute VB_Name = "startup"
Option Explicit

''' Constants
Public Const vbGrey As Long = &H808080
Public Const vbDarkGrey As Long = &H404040
Public Const vbLightGrey As Long = &HC0C0C0
Public Const vbDarkGreen As Long = &H1C001
Public Const SIGNAL_ON As Long = &HC000
Public Const SIGNAL_OFF As Long = &HC0FFFF
Public Const SIGNAL_DISABLED As Long = &H0
Public Const MIN_BUZZER_AMPLITUDE As Double = 20
Public Const RELAY_AFTER_RF As Boolean = True
Public Const RELAY_AFTER_STARTER As Boolean = False
Public Const RELAY_AFTER_IGNITION As Boolean = True
Public Const RELAY_AFTER_POWER_DOWN As Boolean = False
Public Const DELAY_AFTER_WOPEN_1 As Integer = 10
Public Const SERIAL_COMS_TIMEOUT As Integer = 3000
#If DISABLE_EQUIP_CHECKS = 1 Then
Public Const ALLOW_NON_STANDARD_PS As Boolean = True
Public Const ALLOW_NON_STANDARD_CB As Boolean = True
Public Const ALLOW_NO_DIO As Boolean = True
#End If
Public Const RF_XMT_CHAR As String = "7"
Public Const RF_ALT_RCV_CHAR As String = "8"
Public Const POST_OK_MESSAGE As String = "^SYSSTART"
Public Const POST_WAIT As Integer = 2
Public Const POST_WAIT_FW_DNLD As Integer = 2
Public Const GSM_VERIFY_CSQ_MAX_WAIT_NO_CALL_BOX_LOCK As Integer = 45
Public Const MODEM_INIT_MAX_WAIT As Integer = 60

Public Const SMS_MODE_STATE_DEFAULT As String = "3"
'Public Const SMS_MODE_STATE As String = "7"

Public Const SMS_MODE_HELLO_PKT_TMR_DEFAULT As String = "0"
'Public Const SMS_MODE_HELLO_PKT_TMR_TMOBILE As String = "10140"
'Public Const SMS_MODE_HELLO_PKT_TMR_ROGERS As String = "2940"

Public Const SMS_RESET_MIN_DEFAULT As String = "1440"
'Public Const SMS_RESET_MIN_TMOBILE As String = "10080"
'Public Const SMS_RESET_MIN_ROGERS As String = "1440"

'Public Const SMS_REPLY_ADDRESS_TMOBILE As String = "2239"
'Public Const SMS_REPLY_ADDRESS_ROGERS As String = "45000045"

' Currently using 15 digit IMEI number (14 decimal digits plus a "Luhn" check digit)
' not the 16 digit IMEISV (SV stands for software version)
Public Const IMEI_NUM_TAC_FORMAT As String = "00000000"
Public Const IMEI_NUM_SNR_FORMAT As String = "000000"
Public Const IMEI_NUM_CHK_FORMAT As String = "0"
Public Const IMEI_NUM_SV_FORMAT As String = "00"
Public Const IMEI_NUM_FORMAT As String = IMEI_NUM_TAC_FORMAT & IMEI_NUM_SNR_FORMAT
Public Const IMEI_NUM_FORMAT_CHK As String = IMEI_NUM_FORMAT & IMEI_NUM_CHK_FORMAT
Public Const IMEI_NUM_FORMAT_SV As String = IMEI_NUM_FORMAT & IMEI_NUM_SV_FORMAT

Public Const SERIAL_NUM_FORMAT As String = "00000000"

Type stIMEI_NUM
    TAC As Long
    SNR As Long
    SNR_Suffix As Byte
End Type

#If DISABLE_EQUIP_CHECKS = 1 Then
Public Enum etPS_CTL
    ePS_CTL_OFF
    ePS_CTL_ON
    ePS_CTL_RS232
End Enum

Public Enum etCB_CTL
    eCB_CTL_OFF
    eCB_CTL_ON
    eCB_CTL_RS232
End Enum
#End If

'   These are the states for communication with the Power Supply
Public Enum etPS_STATES
    ePS_NULL                        '   0
    ePS_VERIFY                      '   1
    ePS_CHECK_FOR_ERROR             '   2
    ePS_GET_ERROR                   '   3
    ePS_CLEAR_ERROR                 '   4
    ePS_INIT_PS                     '   5
    ePS_INIT_1 = ePS_INIT_PS        '   5
    ePS_INIT_2                      '   6
    ePS_INIT_3                      '   7
    ePS_INIT_4                      '   8
    ePS_INIT_5                      '   9
    ePS_INIT_6                      '   10
    ePS_INIT_7                      '   11
    ePS_INIT_8                      '   12
    ePS_TURN_ON                     '   13
    ePS_TURN_OFF                    '   14
    ePS_MEAS_CURRENT                '   15
    ePS_MEAS_VOLTS                  '   16
    ePS_SET_INRUSH_CURRENT_LIMIT    '   17
    ePS_SET_NOMINAL_CURRENT_LIMIT   '   18
    '   Any new PS state definitions must be added BEFORE this comment
    ePS_NUM_OF_STATES               '   19
End Enum

'   These are the states for communication with the device
Public Enum etELITE_STATES
    '   The assignment of eElite_NULL to ePS_NUM_OF_STATES ensures that there is
    '   no overlap between the values defined in these two steps (just in case
    '   one of the states gets passed off to the wrong serial link in the code.
    eElite_NULL = ePS_NUM_OF_STATES    '   19
    eElite_success                     '   20
    eElite_ERROR                       '   21
    eElite_verify                      '   22
    eElite_GET_IMEI                    '   23
    eElite_GET_IMSI                    '   24
    eElite_GET_CCID                    '   25
    eElite_LOCK_GPS                    '   26
    eElite_CHECK_IGN                   '   27
    eElite_CHECK_STARTER               '   28
    eElite_CHECK_BUZZER                '   29
    eElite_SET_MANUAL_REG_MODE_AND_SELECT_OPERATOR       '   30
    eElite_SET_STATION_ID_FORMAT       '   31
    eElite_GET_SIGNAL_STRENGTH         '   32
    eElite_SET_STATION_AUTOMATIC       '   33
    eElite_GET_GPS_LOCATION            '   34
    eElite_GET_SERIAL_NUM              '   35
    eElite_SET_SERIAL_NUM              '   36
    eElite_GET_APN                     '   37
    eElite_SET_APN                     '   38
    eElite_GET_SERVER_IP               '   39
    eElite_SET_SERVER_IP               '   40
    eElite_GET_PORT                    '   41
    eElite_SET_PORT                    '   42
    eElite_GET_VOLTS                   '   43
    eElite_SET_SMS_MODE                '   44
    eElite_SET_SMS_ADDRESS             '   45
    eElite_ABORT_SMS_SS_PLMN_SELECT    '   46
    eElite_GET_GSM_REG_STATUS          '   47
    eElite_SET_AUTO_RESET_MIN          '   48
    eElite_GET_SMS_MODE                '   49
    eelite_verboseoff                  '   50    AE added 8/27/09 to handle unsolicited messaging
    eElite_APPVER                      '   51    AE added 8/25/09 to handle getting app version
    eElite_Grn_LED                     '   52    AE added 8/26/09 to handle green LED check
    eElite_Red_LED                     '   53    AE added 8/26/09 to handle red LED check
    eElite_Grn_LED_Off                 '   54    AE added 8/26/09 to handle green LED check
    eElite_Red_LED_Off                 '   55    AE added 8/26/09 to handle red LED check
    eElite_Elite_Rly0                  '   56    AE added 8/27/09 to handle Elite relay
    eElite_Elite_Rly1                  '   57    AE added 8/27/09 to handle Elite relay
    eElite_Elite_Rly2                  '   58    AE added 8/27/09 to handle Elite relay
    eElite_RF_RCVR                     '   59    AE added 8/27/09 to handle RF Receiver
    eElite_GPS_INT_ANT                 '   60    JB added 8/30/09 to handle GPS Antenna Selection
    eElite_GPS_EXT_ANT                 '   61    JB added 8/30/09 to handle GPS Antenna Selection
    eElite_ADD                         '   62    AE added 8/28/09 to handle additional tests
    eElite_ADDX                        '   63    AE added 8/28/09 to handle OK result
    eElite_ANTITHEFT                   '   64    AE added 9/03/09 to handle AntiTheft
    eElite_SET_LOW_POWER               '   65    JB added 11/11/09 to handle Low Power Mode
    eElite_ENABLE_BYPASS               '   66    JB added 12/17/09
    eElite_DISABLE_BYPASS              '   67    JB added 12/17/09
    eElite_INIT_COMPLETE               '   68    Atypical state to catch unsollicited message
    eElite_GET_SMS_REPLY_ADDR          '   69
    eElite_MODEM_ID                    '   70
    eElite_MODEM_STATUS                '   71
    eElite_SAVE_FLASH_PARAMS           '   72
    eElite_SET_GSM_MODEM_BAUD_RATE     '   73
    eElite_CLR_GSM_MODEM_BAUD_RATE     '   74
    eElite_SET_GSM_MODEM_FLOW_CTRL     '   75
    eElite_RESET_GSM_MODULE            '   76
    eElite_DISABLE_AUTO_GSM_REG        '   77
    eElite_GET_SCA                     '   78
    eElite_SET_SCA                     '   79
    eElite_GET_GSM_COPS                '   80
    eElite_BGS2_MODEM_ID               '   81
    eElite_BGS2_MODEM_SVN              '   82
    eElite_GET_VOLTS_DUT               '   83
    eElite_GET_AUTO_RESET_MIN          '   84
    eElite_BGS2_SHUTOFF                '   85
'   Add any new Elite communication states before this line
    eElite_MAX_ELITE_STATES
End Enum

'   These are the states for communication with the Call Box
Public Enum etCB_STATES
    eCB_NULL                         '   0
    eCB_VERIFY                       '   1
    eCB_CHECK_FOR_ERROR              '   2
    eCB_GET_ERROR                    '   3
    eCB_CLEAR_ERROR                  '   4
    eCB_INIT_CB                      '   5
    eCB_RESET                        '   6
    eCB_LOC_UPD                      '   7
    eCB_ESC                          '   8
    eCB_END_REMOTE                   '   9
    eCB_CONFIG_GSM                   '   10
    eCB_CONFIG_BCCH                  '   11
    eCB_CONFIG_PCH                   '   12
    eCB_CONFIG_PWR_LVL_BCCH          '   13
    eCB_CONFIG_PWR_LVL_TCH           '   14
    eCB_CONFIG_PWR_LVL_RF            '   15
    eCB_CONFIG_PRE_ATTEN             '   16
'   Add any new Call Box communication states before this line
    eCB_NUM_OF_STATES                '
End Enum

'   These are the various status' of the COM links
Public Enum etCOM_STATUS
    eSTAT_SUCCESS
    eSTAT_RUNNING
    eSTAT_DOWNLOADING
    eSTAT_TIMEOUT
    eSTAT_NOT_CONNECTED
    eSTAT_ERROR
    eSTAT_SHUTTING_DOWN
    eSTAT_CME_3
    eSTAT_CME_515
End Enum

''' Classes
Public g_clsBarcode As clsPrintBarcode
'TODO: Remove clsOpenAT class from project
'Public g_clsOpenAT As clsOpenAT

''' Forms
Public g_fIO As frmIO
Public g_fMainTester As frmMainTester
'TODO: Remove frmDownloadProgress from project
'Public g_fDownloadProgress As frmDownloadProgress
Public g_fMicIn As frmMicIn
Public g_fSerialNumber As frmSerialNumber

''' Variables
' Why aren't these in globals???
Public g_iMainTop As Integer
Public g_szIMSI As String
Public g_szCCID As String
Public g_szIMEI As String
Public g_lNextIMEINumber As stIMEI_NUM
Public g_lEndingIMEINumber As stIMEI_NUM
Public g_iMainLeft As Integer
Public g_bAutoTest As Boolean
Public g_dSoundFreq As Double
Public g_szPrintDir As String
Public g_iMainWidth As Integer
Public g_dSoundLevel As Double
Public g_iGPSTimeout As Integer
Public g_szGPSID As String
Public g_iMainHeight As Integer
Public g_bQuitProgram As Boolean
Public g_szDataDirPath As String
Public g_szFirmwareFileName As String
Public g_szFirmwareVersion As String
Public g_ePS_State As etPS_STATES
Public g_eElite_State As etELITE_STATES
Public g_eCB_State As etCB_STATES
Public g_bOpenATRunning As Boolean
Public g_lNextSerialNumber As Long
Public g_ePS_COM_Status As etCOM_STATUS
Public g_eElite_Status As etCOM_STATUS
Public g_eCB_COM_Status As etCOM_STATUS
Public g_bStopApplication As Boolean
Public g_iNumBarcodeScans As Integer
Public g_szTestResultsFile As String
Public g_lEndingSerialNumber As Long
Public g_iNumBarcodePrints As Integer
Public g_iMinBuzzerAmplitude As Integer
Public g_iDownloadCheckSetting As Integer
Public g_bDiagnosticActive As Boolean
Public g_iTestFunctionsCheckSetting As Integer
Public g_iConfigureModemIDCheckSetting As Integer
Public g_szCellSiteSimulatorPLMN As String
Public m_szStateResponses(ePS_NUM_OF_STATES To eElite_MAX_ELITE_STATES) As String
#If DISABLE_EQUIP_CHECKS = 1 Then
Public g_ePS_Control As etPS_CTL
Public g_eCB_Control As etPS_CTL
Public g_bDIO_Installed As Boolean
#End If
'AE 10/17/09
Public g_szIntGPS As String
Public g_szExtGPS As String
Public g_szAPPVER As String
Public g_szAddResult As String
Public g_bStartupError As String
'@JMB
Public Const CME_ERR_STRINGS As String = ":?" & _
                                            ",3:Operation not allowed." & _
                                            ",32:Network not allowed - emergency calls only." & _
                                            ",515:Processing Command. Please Wait." & _
                                            ",527:RR or MM are busy. Please Wait." & _
                                            ",528:Location update failure. Emergency calls only." & _
                                            ",529:PLMN selection failure. Emergency calls only." & _
                                            ",547:Emergency call is allowed without SIM." & _
                                            ",548:No flash objects to delete."

'   These are the CME errors that we handle
Public Enum etCME_ERR
    eCME_ERR_FALSE = -1     ' Operation not allowed.
    eCME_ERR_UNKNOWN = 0    ' Operation not allowed.
    eCME_ERR_003 = 3        ' Operation not allowed.
    eCME_ERR_032 = 32       ' Network not allowed - emergency calls only.
    eCME_ERR_515 = 515      ' Processing Command. Please Wait.
    eCME_ERR_527 = 527      ' RR or MM are busy. Please Wait.
    eCME_ERR_528 = 528      ' Location update failure. Emergency calls only.
    eCME_ERR_529 = 529      ' PLMN selection failure. Emergency calls only.
    eCME_ERR_547 = 547      ' Emergency call is allowed without SIM.
    eCME_ERR_548 = 548      ' No flash objects to delete.
End Enum

Public g_sArrCME_ErrorNumStrings() As String
Public g_sArrCME_ErrorStrings() As String
'@JMB

Public g_bDebugOn As Boolean

'@RWL
Public strNewTestLog As String
Public TestLogCollection As New Collection
'@RWL
Private m_bSplashing As Boolean

Sub Main()
    Dim i As Integer
    Dim t As Double
    Dim szError As String
    Dim oAttributes As IXMLDOMNamedNodeMap
    Dim oNode As IXMLDOMNode
    Dim sArrCME_ErrorNumStringsTmp() As String
    Dim sArrCME_ErrorStringsTmp() As String

    '   Open and validate the Model Type XML configuration file
    szError = Config.OpenModelTypeFile
    If szError <> "" Then
        '   There was something wrong with the ModelType.xml file.  Abort the
        '   app now before anyone gets hurt.
        MsgBox szError
        Exit Sub
    End If
    
    '       Initialize the state response table
    m_szStateResponses(eElite_NULL) = "None"                                   '   19
    m_szStateResponses(eElite_success) = "None"                                '   20
    m_szStateResponses(eElite_ERROR) = "None"                                  '   21
    m_szStateResponses(eElite_verify) = "OK"                                   '   22
    m_szStateResponses(eElite_GET_IMEI) = "OK"                                 '   23
    m_szStateResponses(eElite_GET_IMSI) = "OK"                                 '   24
    m_szStateResponses(eElite_GET_CCID) = "OK"                                 '   25
    m_szStateResponses(eElite_LOCK_GPS) = "OK"                                 '   26
    m_szStateResponses(eElite_CHECK_IGN) = "OK"                                '   27
    m_szStateResponses(eElite_CHECK_STARTER) = "OK"                            '   28
    m_szStateResponses(eElite_CHECK_BUZZER) = "OK"                             '   29
    m_szStateResponses(eElite_SET_MANUAL_REG_MODE_AND_SELECT_OPERATOR) = "OK"  '   30
    m_szStateResponses(eElite_SET_STATION_ID_FORMAT) = "OK"                    '   31
    m_szStateResponses(eElite_GET_SIGNAL_STRENGTH) = "OK"                      '   32
    m_szStateResponses(eElite_SET_STATION_AUTOMATIC) = "OK"                    '   33
    m_szStateResponses(eElite_GET_GPS_LOCATION) = "OK"                         '   34
    m_szStateResponses(eElite_GET_SERIAL_NUM) = "OK"                           '   35
    m_szStateResponses(eElite_SET_SERIAL_NUM) = "OK"                           '   36
    m_szStateResponses(eElite_GET_APN) = "OK"                                  '   37
    m_szStateResponses(eElite_SET_APN) = "OK"                                  '   38
    m_szStateResponses(eElite_GET_SERVER_IP) = "OK"                            '   39
    m_szStateResponses(eElite_SET_SERVER_IP) = "OK"                            '   40
    m_szStateResponses(eElite_GET_PORT) = "OK"                                 '   41
    m_szStateResponses(eElite_SET_PORT) = "OK"                                 '   42
    m_szStateResponses(eElite_GET_VOLTS) = "OK"                                '   43
    m_szStateResponses(eElite_SET_SMS_MODE) = "OK"                             '   44
    m_szStateResponses(eElite_SET_SMS_ADDRESS) = "OK"                          '   45
    m_szStateResponses(eElite_ABORT_SMS_SS_PLMN_SELECT) = "OK"                 '   46
    m_szStateResponses(eElite_GET_GSM_REG_STATUS) = "+CREG: "                  '   47
    m_szStateResponses(eElite_SET_AUTO_RESET_MIN) = "OK"                       '   48
    m_szStateResponses(eElite_GET_SMS_MODE) = "OK"                             '   49
    m_szStateResponses(eelite_verboseoff) = "OK"                               '   50
    m_szStateResponses(eElite_APPVER) = "OK"                                   '   51
    m_szStateResponses(eElite_Grn_LED) = "OK"                                  '   52
    m_szStateResponses(eElite_Red_LED) = "OK"                                  '   53
    m_szStateResponses(eElite_Grn_LED_Off) = "OK"                              '   54
    m_szStateResponses(eElite_Red_LED_Off) = "OK"                              '   55
    m_szStateResponses(eElite_Elite_Rly0) = "OK"                               '   56
    m_szStateResponses(eElite_Elite_Rly1) = "OK"                               '   57
    m_szStateResponses(eElite_Elite_Rly2) = "OK"                               '   58
    m_szStateResponses(eElite_RF_RCVR) = "OK"                                  '   59
    m_szStateResponses(eElite_GPS_INT_ANT) = "OK"                              '   60
    m_szStateResponses(eElite_GPS_EXT_ANT) = "OK"                              '   61
    m_szStateResponses(eElite_ADD) = "OK"                                      '   62
    m_szStateResponses(eElite_ADDX) = "OK"                                     '   63
    m_szStateResponses(eElite_ANTITHEFT) = "OK"                                '   64
    m_szStateResponses(eElite_SET_LOW_POWER) = "OK"                            '   65
    m_szStateResponses(eElite_ENABLE_BYPASS) = "None"                          '   66
    m_szStateResponses(eElite_DISABLE_BYPASS) = "None"                         '   67
    m_szStateResponses(eElite_INIT_COMPLETE) = "Done initializing modem."      '   68
    m_szStateResponses(eElite_GET_SMS_REPLY_ADDR) = "OK"                       '   69
    m_szStateResponses(eElite_MODEM_ID) = "OK"                                 '   70
    m_szStateResponses(eElite_MODEM_STATUS) = "OK"                             '   71
    m_szStateResponses(eElite_SAVE_FLASH_PARAMS) = "OK"                        '   72
    m_szStateResponses(eElite_SET_GSM_MODEM_BAUD_RATE) = "OK"                  '   73
    m_szStateResponses(eElite_CLR_GSM_MODEM_BAUD_RATE) = "OK"                  '   74
    m_szStateResponses(eElite_SET_GSM_MODEM_FLOW_CTRL) = "OK"                  '   75
    m_szStateResponses(eElite_RESET_GSM_MODULE) = "OK"                         '   76
    m_szStateResponses(eElite_DISABLE_AUTO_GSM_REG) = "OK"                     '   77
    m_szStateResponses(eElite_GET_SCA) = "OK"                                  '   78
    m_szStateResponses(eElite_SET_SCA) = "OK"                                  '   79
    m_szStateResponses(eElite_GET_GSM_COPS) = "+COPS: "                        '   80
    m_szStateResponses(eElite_BGS2_MODEM_ID) = "OK"                            '   81
    m_szStateResponses(eElite_BGS2_MODEM_SVN) = "OK"                           '   82
    m_szStateResponses(eElite_GET_VOLTS_DUT) = "OK"                            '   83
    m_szStateResponses(eElite_GET_AUTO_RESET_MIN) = "OK"                       '   84
    m_szStateResponses(eElite_BGS2_SHUTOFF) = "^SHUTDOWN"                      '   85

    ' sorry about that, but VB6 doesn't seem to have a clean way to initialize an array
    
    ' Get the Cell Site Simulator PLMN from the xml file
    Set oAttributes = g_oSettings.selectSingleNode("CellSiteSimulator").Attributes
    g_szCellSiteSimulatorPLMN = oAttributes.getNamedItem("Station").Text
    
    ' Parse the CME_ERR_STRINGS data into two arrays
    sArrCME_ErrorStringsTmp = Split(CME_ERR_STRINGS, ",")
    
    ReDim g_sArrCME_ErrorNumStrings(UBound(sArrCME_ErrorStringsTmp))
    ReDim g_sArrCME_ErrorStrings(UBound(sArrCME_ErrorStringsTmp))
    
    For i = 0 To UBound(sArrCME_ErrorStringsTmp)
        sArrCME_ErrorNumStringsTmp = Split(sArrCME_ErrorStringsTmp(i), ":")
        g_sArrCME_ErrorNumStrings(i) = sArrCME_ErrorNumStringsTmp(0)
        g_sArrCME_ErrorStrings(i) = sArrCME_ErrorNumStringsTmp(1)
        Erase sArrCME_ErrorNumStringsTmp
    Next i
    
    Erase sArrCME_ErrorStringsTmp
    
    m_bSplashing = True
    t = Timer
    DoEvents
        ' JMB 11/22/2009 to facilitate testing
    If "-manual" = command Then
                ' Start tester in manual mode
        g_bAutoTest = False
        g_bDebugOn = True
    Else
                ' Start tester in auto mode
        g_bAutoTest = True
        g_bDebugOn = False
    End If
    g_ePS_State = ePS_NULL
    g_eElite_State = eElite_NULL
    g_eCB_State = eCB_NULL
    g_bQuitProgram = False
    g_bOpenATRunning = False
    g_bStopApplication = False
    g_bDiagnosticActive = False
#If DISABLE_EQUIP_CHECKS = 1 Then
    g_ePS_Control = etPS_CTL.ePS_CTL_RS232
    g_bDIO_Installed = True
#End If
    
    Set oAttributes = g_oSettings.selectSingleNode("Main").Attributes
    g_iMainTop = CInt(Val(oAttributes.getNamedItem("Top").Text))
    g_iMainLeft = CInt(Val(oAttributes.getNamedItem("Left").Text))
    g_iMainHeight = CInt(Val(oAttributes.getNamedItem("Height").Text))
    g_iMainWidth = CInt(Val(oAttributes.getNamedItem("Width").Text))
    Set oAttributes = g_oSettings.selectSingleNode("Operations").Attributes
    If oAttributes.getNamedItem("Download").Text = "1" Then
        g_iDownloadCheckSetting = vbChecked
    Else
        g_iDownloadCheckSetting = vbUnchecked
    End If
    If oAttributes.getNamedItem("Configure").Text = "1" Then
        g_iConfigureModemIDCheckSetting = vbChecked
    Else
        g_iConfigureModemIDCheckSetting = vbUnchecked
    End If
    If oAttributes.getNamedItem("Test").Text = "1" Then
        g_iTestFunctionsCheckSetting = vbChecked
    Else
        g_iTestFunctionsCheckSetting = vbUnchecked
    End If
    g_szDataDirPath = g_oSettings.selectSingleNode("DataPath").Text
    g_szTestResultsFile = g_szDataDirPath & g_oSettings.selectSingleNode("ResultsLog").Text
    
    '@RWL
    strNewTestLog = g_szDataDirPath & "New_ResultsLog.txt"
    '@RWL
    
    Set oAttributes = g_oSettings.selectSingleNode("BarCodes").Attributes
    g_iNumBarcodePrints = CInt(Val(oAttributes.getNamedItem("NumToPrint").Text))
    g_iNumBarcodeScans = CInt(Val(oAttributes.getNamedItem("NumToScan").Text))
    
    Set oAttributes = g_oSettings.selectSingleNode("GPSSimulator").Attributes
    g_szGPSID = oAttributes.getNamedItem("SVID").Text
    g_iGPSTimeout = CInt(Val(oAttributes.getNamedItem("Lock").Text))
    
    Set oAttributes = g_oSettings.selectSingleNode("Sound").Attributes
    g_iMinBuzzerAmplitude = CInt(Val(oAttributes.getNamedItem("Min").Text))

    i = 0
    For Each oNode In g_oSettings.selectSingleNode("TestSystem").childNodes
        If oNode.selectSingleNode("Computer").Attributes.getNamedItem("Name").Text = ComputerName Then
            Exit For
        End If
        i = i + 1
    Next oNode
    
    If oNode Is Nothing Then
        MsgBox "Could not find settings for tester computer named " & ComputerName
        Exit Sub
    End If
    
    Set oNode = oNode.selectSingleNode("IMEI")
    If oNode.selectSingleNode("Input_Method").Text <> "BY_COMPUTER_NAME" Then
        Debug.Print "Track IMEI by model was removed. Check config.bas for detailed comment"
        MsgBox "IMEI serial number tracking method not supported"
        Exit Sub
        If oNode.selectSingleNode("Input_Method").Text <> "BY_MODEL" Then
            MsgBox "IMEI serial number tracking method not supported"
            Exit Sub
        End If
    End If
    
    Set oNode = oNode.parentNode.selectSingleNode("Attenuation")
    g_iGSM_Attenuation = CInt(oNode.Attributes.getNamedItem("GSM").Text)
    
    g_iTestStationID = i

    '   Instantiate the printer class
    Set g_clsBarcode = New clsPrintBarcode
    
    SerialComs.Init (POST_OK_MESSAGE)
    
    '       Throw up the PassTime Splash screen while we're bringing up and initializing the application
    frmSplash.Show
    
    '   Create and then load the various forms in the application
    Set g_fIO = New frmIO
    Set g_fMainTester = New frmMainTester
    
    g_bStartupError = False
    Load g_fIO

#If DISABLE_EQUIP_CHECKS = 0 Then
    Load g_fMainTester
    If False = g_fMainTester.InitCOM_PortToPowerSupply Then
        g_bStartupError = True
    ElseIf False = SerialComs.InitPowerSupply Then
        g_bStartupError = True
    End If
    
    If 0 <> CInt(Val(Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PORT))) Then
        If False = g_fMainTester.InitCOM_PortToCallBox Then
            g_bStartupError = True
        End If
    End If
#Else
    If False = g_fIO.Init Then
        If True = g_bDIO_Installed Then
            g_bStartupError = True
        End If
    Else
        Load g_fMainTester
    End If
    If False = g_fMainTester.InitCOM_PortToPowerSupply Then
        ' RS232 Connection to PS failed
        If etPS_CTL.ePS_CTL_RS232 = g_ePS_Control Then
            ' RS232 communications could not be verified
            g_bStartupError = True
        End If
    ElseIf False = SerialComs.InitPowerSupply Then
            g_bStartupError = True
    End If
    
    If 0 <> CInt(Val(Config.GetComSettings(etCOM_DEV_PS, etCOM_SETTING_PORT))) Then
        If False = g_fMainTester.InitCOM_PortToCallBox Then
            ' RS232 Connection to Call Box failed
            If etCB_CTL.eCB_CTL_RS232 = g_ePS_Control Then
                ' RS232 communications could not be verified
                g_bStartupError = True
            End If
        End If
    End If
#End If

    If True = g_bStartupError Then
        Unload frmSplash
        Unload g_fMainTester
        Unload g_fIO
        Exit Sub
    End If
    
    '   Finally, start the various processes going
    If g_fMainTester.InitCOM_PortToDevice Then
        If g_fIO.Init Then
            '   Now wait until at least 3 seconds have passed before taking
            '   down the Splash screen
            Do Until (Timer - t) >= 3 Or t > Timer
                DoEvents
            Loop
            m_bSplashing = False
            Unload frmSplash
            g_fMainTester.Show
            On Error GoTo Fishnet
            While Not g_bStopApplication
                DoEvents
            Wend
            On Error GoTo 0
        End If
    End If
    EndProgram True
    Exit Sub

Fishnet:
    MsgBox "Error code: " & Err.number & vbCrLf & vbCrLf & _
           "Desc: " & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & Err.Source
    Resume Next
End Sub

Public Sub EndProgram(Optional bFull As Boolean = True)

    Dim oForm As Form
    Dim oAttributes As IXMLDOMNamedNodeMap

    If m_bSplashing Then
        Unload frmSplash
    End If
    SerialComs.Destroy
    If bFull Then
        Unload g_fMainTester
        Unload g_fIO
        Set g_fMainTester = Nothing
        Set g_fIO = Nothing
    End If
    Set g_clsBarcode = Nothing
    '   Update then write the contents of the XML object tree back out to the
    '   XML configuration file.
    Set oAttributes = g_oSettings.selectSingleNode("Operations").Attributes
    If g_iDownloadCheckSetting = vbChecked Then
        oAttributes.getNamedItem("Download").Text = "1"
    Else
        oAttributes.getNamedItem("Download").Text = "0"
    End If
    If g_iConfigureModemIDCheckSetting = vbChecked Then
        oAttributes.getNamedItem("Configure").Text = "1"
    Else
        oAttributes.getNamedItem("Configure").Text = "0"
    End If
    If g_iTestFunctionsCheckSetting = vbChecked Then
        oAttributes.getNamedItem("Test").Text = "1"
    Else
        oAttributes.getNamedItem("Test").Text = "0"
    End If
    Set oAttributes = g_oSettings.selectSingleNode("BarCodes").Attributes
    oAttributes.getNamedItem("NumToPrint").Text = Str(g_iNumBarcodePrints)
    oAttributes.getNamedItem("NumToScan").Text = Str(g_iNumBarcodeScans)
    
    Set oAttributes = g_oSettings.selectSingleNode("GPSSimulator").Attributes
    oAttributes.getNamedItem("SVID").Text = g_szGPSID
    oAttributes.getNamedItem("Lock").Text = Str(g_iGPSTimeout)
    
    Set oAttributes = g_oSettings.selectSingleNode("Sound").Attributes
    oAttributes.getNamedItem("Min").Text = Str(g_iMinBuzzerAmplitude)
    
    Set oAttributes = g_oSettings.selectSingleNode("Main").Attributes
    oAttributes.getNamedItem("Top").Text = Str(g_iMainTop)
    oAttributes.getNamedItem("Left").Text = Str(g_iMainLeft)
    oAttributes.getNamedItem("Height").Text = Str(g_iMainHeight)
    oAttributes.getNamedItem("Width").Text = Str(g_iMainWidth)
    g_oEliteTester.Save ("ModelTypes.xml")
    '   Final cleanup.  Some objects still remain executing after shutdown.
    For Each oForm In Forms
        ' MsgBox "'" & oForm.Caption & "' still open! Closing form."
        Unload oForm
    Next oForm

End Sub

Public Function Delay(Optional sec As Double = 0.1)
    Dim endTimer As Double
    Dim t As Double
    t = Timer
    endTimer = t + sec
    If endTimer >= 86400 Then
        t = endTimer - 86400
    End If
    Do Until (Timer - t) >= sec
        If t > Timer Then
            Debug.Print "Timer failure @" & Timer & " where t = " & t & " delay = " & sec & " seconds."
            Exit Do
        End If
        DoEvents
    Loop
End Function

'@RWL
Public Function LoadResultsLogIntoCollection() As Boolean

    Dim iResult As Integer
    Dim LineContents As Variant
    Dim FileLine As String
    Dim key As String
    Dim ResultsLogEntry(0 To 8) As Variant
    Dim FileLineCount As Long
    Dim fhNumber As Integer

    LoadResultsLogIntoCollection = False
    fhNumber = FreeFile
    '
    '   Check to see if the log file exists.  If not then create it. or else when we try
    '   to read from it we'll get an error
    '
    On Error GoTo FileDoesntExist
    Open g_szTestResultsFile For Input As #fhNumber
    GoTo FileExists:
    '
    '   Create the file.  Since it is new, there is no info in it to process.
    '   Just exit this subroutine.
    '
FileDoesntExist:
    If 55 <> Err.number Then
        ' Create an empty file
        Open g_szTestResultsFile For Output As #fhNumber
        Close #fhNumber
    End If
    LoadResultsLogIntoCollection = True
    Exit Function
    
FileExists:
    '
    '   The first thing that needs to be done is load the existing log file into the collection.
    '   Each item in the collection corresponds to a line in the text file.
    '   The key to each item in the collection is the serial number with "K" prepended to it,
    '   i.e, K6006895.  Duplicate entries found in the text file will alter the key by prepending
    '   additional "K"'s to avoid an error that will occur if an item with a duplicate key is
    '   added to the collection.  For example, the first duplicate entry will have a key of
    '   KK6006895 and the next duplicate entry will be KKK6006895, etc...
    '
    '******** 9/4/09  Are 'K's used anymore, when does DuplicateKey label get hit? ************
       
    On Error GoTo FileParseError
    Do While Not EOF(fhNumber)
    
        Line Input #fhNumber, FileLine
        LineContents = Split(FileLine, ",")
        
        ResultsLogEntry(0) = LineContents(0)
        ResultsLogEntry(1) = Mid(LineContents(1), 2, Len(LineContents(1)) - 2)
        ResultsLogEntry(2) = Mid(LineContents(2), 2, Len(LineContents(2)) - 2)
        ResultsLogEntry(3) = Mid(LineContents(3), 2, Len(LineContents(3)) - 2)
        ResultsLogEntry(4) = Mid(LineContents(4), 2, Len(LineContents(4)) - 2)
        ResultsLogEntry(5) = Mid(LineContents(5), 2, Len(LineContents(5)) - 2)
        ResultsLogEntry(6) = Mid(LineContents(6), 2, Len(LineContents(6)) - 2)
        ResultsLogEntry(7) = Mid(LineContents(7), 2, Len(LineContents(7)) - 2)
        ResultsLogEntry(8) = 1
        '
        '   Note:  Array member 8 of ResultsLogEntry is special.  It is used to indicate
        '   the number of entries having this same serial number (key).  A value of 1 means
        '   that there is only 1 item with the associated key and thus there are no
        '   duplicate entries with this same serial number
        '
        key = "K" & LineContents(0)
        On Error GoTo DuplicateKey:
        GoTo AddEntryToCollection

FileParseError:
        ' JMB 2009-11-14: Added error handling but not couldn't find root cause
        ' MsgBox "Error" & Err.Number & ": " & Err.Description, vbCritical
        Dim szErrHelp
        
        szErrHelp = "." & vbCr & vbCr & "Please make a copy of the " & vbCr & "file, then delete the file and restart the application"
        MsgBox "Error parsing results file, " & g_szTestResultsFile & szErrHelp, vbCritical, "Errror in function startup.LoadResultsLogIntoCollection"
        startup.g_bStartupError = True
        Close #fhNumber
        Exit Function
        
        
AddEntryToCollection:
        
        TestLogCollection.Add ResultsLogEntry, key
        FileLineCount = FileLineCount + 1
        
    Loop
    
    Close #fhNumber
    LoadResultsLogIntoCollection = True
    Exit Function
    
DuplicateKey:
    '
    '   The entry about to be added is a duplicate, but we cannot add the same key to the
    '   collection.  Change the key of this new entry (the duplicate) by prepending
    '   another "K" to it's key.  Also, increment array member 8 in the first occurrance
    '   of this duplicate in the collection to indicate during later processing that this
    '   serial number (key) has duplicates.
    '
    Dim DupEntry As Variant
    
    DupEntry = TestLogCollection.Item(key)
    
    DupEntry(8) = DupEntry(8) + 1
    
    TestLogCollection.Remove (key)
    TestLogCollection.Add DupEntry, key
    
    key = "K" & key
    
    'iResult = MsgBox("Duplicate entry found while loading log file", vbCritical)
    
    Resume AddEntryToCollection
    
End Function

Public Function ProcessCME_Error(szErrString As String) As etCME_ERR
    ' Match input string to list of CME errors that the tester handles.
    ' The CME error number strings and corresponding description strings
    ' are stored in two arrays: g_sArrCME_ErrorNumStrings and g_sArrCME_ErrorStrings
    ' Index zero of these arrays at initialization contain "", a null string, and
    ' "?" respectively.
    '
    ' Everytime the ProcessCME_Error function is called, the error number
    ' string array index zero entry is cleared (set to default null string value).
    ' Similarly the error description array index zero entry is reset to the
    ' default "?" string.
    '
    ' If a match to the input string is found, the error number string array index
    ' zero entry is updated to contain the matching error number string, a message
    ' is written to the tester log with the corresponding error description and
    ' the function returns the actual CME error number. Use the enumerated type
    ' etCME_ERR to handle these errors.
    '
    ' If the input string is not a CME error, -1 is returned.
    '
    ' If the input string is a CME error but not one that our tester handles,
    ' the ProcessCME_Error0 returs zero.

    Dim szSearchString As String
    Dim lNumOfDigits As Long
    Dim szErrorNumString As String
    Dim iCME_ErrorStringIdx As Integer
    Dim eCME_ErrorNumber As etCME_ERR
    Dim numOfErrorStrings As Integer
    
    szSearchString = "+CME ERROR: "
    
    ' Clear the current error number string stored in the
    ' g_sArrCME_ErrorNumStrings array at index 0
    g_sArrCME_ErrorNumStrings(0) = ""
    g_sArrCME_ErrorStrings(0) = "?"
    
    ' Always using zero based strings so LBound is always zero
    numOfErrorStrings = UBound(g_sArrCME_ErrorStrings) + 1

    ' Default return value: Not a CME error
    ProcessCME_Error = -1
    
    If Null = Len(szErrString) Then
        ' Nothing to search or do, just exit the function
        Exit Function
    End If

    ' See how many characters are left after we remove the text from the string
    lNumOfDigits = Len(szErrString) - Len(szSearchString)
    
    If lNumOfDigits > 0 And lNumOfDigits < 4 Then
        ' Error number string is 1, 2, or 3 characters long
        
        If 1 = InStr(1, szErrString, szSearchString) Then
            ' Verified that this is in fact a CME error string
            
            ' Parse out the error number string from the rest of the string
            szErrorNumString = Right(szErrString, lNumOfDigits)
            
            ' Convert error number string to an integer
            eCME_ErrorNumber = CInt(Val(szErrorNumString))
            
            ' Search the array of known error number strings for one that matches
            For iCME_ErrorStringIdx = 0 To (numOfErrorStrings - 1)
                If szErrorNumString = g_sArrCME_ErrorNumStrings(iCME_ErrorStringIdx) Then
                    ' Found a matching error number string
                    Exit For
                End If
            Next iCME_ErrorStringIdx
            
            If numOfErrorStrings = iCME_ErrorStringIdx Then
                ' This is a CME error that the tester does not handle
                
                ' Write the current unhandled error number string to the database at index 0
                g_sArrCME_ErrorNumStrings(0) = szErrorNumString
                
                ' Set the return value
                ProcessCME_Error = 0
            Else
                ' Update the current CME error info stored at index zero of the arrays
                g_sArrCME_ErrorNumStrings(0) = g_sArrCME_ErrorNumStrings(iCME_ErrorStringIdx)
                g_sArrCME_ErrorStrings(0) = g_sArrCME_ErrorStrings(iCME_ErrorStringIdx)
                
                ' Set the return value
                ProcessCME_Error = eCME_ErrorNumber
            End If

            ' Write the error number string and error description to the tester log file
            g_fMainTester.LogMessage (" +CME ERROR: " & g_sArrCME_ErrorNumStrings(0) & "--" & g_sArrCME_ErrorStrings(0))
            
        End If
    End If
End Function

