VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAS1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mBoardNumber As Integer
Private mError As Integer
Private mErrString As String
Private mAnalogIn(15) As Single
Private mPortA As Integer
Private mPortB As Integer
Private mPortCL As Integer
Private mPortCH As Integer
Private MemHandle As Long
Private mScanData() As Integer
Private mScanValues() As Single

'' Digital IO Ports
Public Enum cbDIOPORT
    CBDIOPORTA = FIRSTPORTA
    CBDIOPORTB = FIRSTPORTB
    CBDIOPORTCL = FIRSTPORTCL
    CBDIOPORTCH = FIRSTPORTCH
End Enum

''' These are the states which the ports can be in
Public Enum cbPORTSTATE
    cbPORTOUT = DIGITALOUT
    cbPORTIN = DIGITALIN
End Enum

'''These are the valid ranges for the DAS1200
Public Enum cbRange
    cbBP10V = BIP10VOLTS                   ' Bipolar Ranges (-10 to +10 Volts)
    cbBP5V = BIP5VOLTS
    cbBP2_5V = BIP2PT5VOLTS
    cbBP1_25 = BIP1PT25VOLTS
    cbUP10V = UNI10VOLTS                 ' Unipolar Ranges (0 to 10 Volts)
    cbUP5V = UNI5VOLTS
    cbUP2_5V = UNI2PT5VOLTS
    cbUP1_25 = UNI1PT25VOLTS
End Enum

Private mAnalogCH0 As Double
Private mAnalogCH1 As Double
Private mAnalogCH3 As Double
Private mAnalogCH4 As Double
Private mAnalogCH5 As Double
Private mAnalogCH6 As Double
Private mAnalogCH7 As Double
Private mAnalogCH8 As Double
Private mAnalogCH9 As Double
Private mAnalogCH10 As Double
Private mAnalogCH11 As Double
Private mAnalogCH12 As Double
Private mAnalogCH13 As Double
Private mAnalogCH14 As Double
Private mAnalogCH15 As Double

Public Event cbDAS1200Error(ErrorNumber As Integer, ErrorString As String)

Public Function Init(nBoardNum As Integer) As Boolean
    
    mBoardNumber = nBoardNum

End Function

Public Property Get BoardNumber() As Integer
    
    BoardNumber = mBoardNumber

End Property

Public Function AnalogOutSetup(nChannel As Integer, nGain As cbRange) As Boolean
    
    mError = cbAOut(mBoardNumber, nChannel, nGain, 0)
    If mError <> NOERRORS Then
        HandleError
        AnalogOutSetup = True
    End If

End Function

Public Function ConfigureDIOPort(nPort As cbDIOPORT, nPortState As cbPORTSTATE) As Boolean
    
#If DISABLE_EQUIP_CHECKS = 0 Then
    mError = cbDConfigPort(mBoardNumber, nPort, nPortState)
    
    If mError <> NOERRORS Then
        HandleError
        ConfigureDIOPort = True
    Else
        ConfigureDIOPort = False
    End If
#Else
    On Error Resume Next

    mError = cbDConfigPort(mBoardNumber, nPort, nPortState)
    
   If True = g_bDIO_Installed Then
        If mError <> NOERRORS Then
            HandleError
            ConfigureDIOPort = True
        Else
            ConfigureDIOPort = False
        End If
    End If
#End If

End Function

Public Function ReadAnalogIn(Channel As Integer, Range As cbRange) As Boolean
    
    Dim AnalogValue As Integer
    mError = cbAIn(mBoardNumber, Channel, Range, AnalogValue)
    If mError <> NOERRORS Then
        HandleError
        ReadAnalogIn = True
        Exit Function
    End If
    mError = cbToEngUnits(mBoardNumber, Range, AnalogValue, mAnalogIn(Channel))
    If mError <> NOERRORS Then
        HandleError
        ReadAnalogIn = True
    End If

End Function

Public Function ReadDIO(nPort As cbDIOPORT) As Boolean
    
    Select Case nPort
        Case CBDIOPORTA
            mError = cbDIn(mBoardNumber, nPort, mPortA)
        Case CBDIOPORTB
            mError = cbDIn(mBoardNumber, nPort, mPortB)
        Case CBDIOPORTCH
            mError = cbDIn(mBoardNumber, nPort, mPortCH)
        Case CBDIOPORTCL
            mError = cbDIn(mBoardNumber, nPort, mPortCL)
        Case Else
            MsgBox "Invalid Port", vbCritical
    End Select
    If mError <> NOERRORS Then
        ReadDIO = True
        HandleError
    Else
        ReadDIO = False
    End If

End Function

Public Function WriteDIO(nPort As cbDIOPORT) As Boolean
    
    Select Case nPort
        Case CBDIOPORTA
            mError = cbDOut(mBoardNumber, nPort, mPortA)
        Case CBDIOPORTB
            mError = cbDOut(mBoardNumber, nPort, mPortB)
        Case CBDIOPORTCH
            mError = cbDOut(mBoardNumber, nPort, mPortCH)
        Case CBDIOPORTCL
            mError = cbDOut(mBoardNumber, nPort, mPortCL)
    End Select
    If mError <> NOERRORS Then
        WriteDIO = True
        HandleError
    Else
        WriteDIO = False
    End If

End Function

Public Property Get PortA() As Integer
    
    PortA = mPortA

End Property

Public Property Let PortA(nValue As Integer)
    
    mPortA = nValue

End Property

Public Property Get PortB() As Integer

    ReadDIO CBDIOPORTB
    PortB = mPortB

End Property

Public Property Let PortB(nValue As Integer)

    '   I think that Port B is read only
    mPortB = nValue

End Property

Public Property Get PortCH() As Integer

    PortCH = mPortCH

End Property

Public Property Let PortCH(nValue As Integer)

    mPortCH = nValue

End Property

Public Property Get PortCL() As Integer

    PortCL = mPortCL

End Property

Public Property Let PortCL(nValue As Integer)

    mPortCL = nValue

End Property

Public Property Get AnalogIn(Index As Integer) As Single

    AnalogIn = mAnalogIn(Index)

End Property

Public Function SetAnalogScanSize(nSize As Long) As Boolean
    
#If DISABLE_EQUIP_CHECKS = 0 Then
    If MemHandle <> 0 Then  ''' in case we already allocated some
        mError = cbWinBufFree(MemHandle)
    End If
    MemHandle = cbWinBufAlloc(nSize)
    If MemHandle = 0 Then
        RaiseEvent cbDAS1200Error(0, "Failed To Allocate Memory For Scan Data")
    End If
    ReDim mScanData(0 To nSize - 1)
    ReDim mScanValues(0 To nSize - 1)

#Else
    SetAnalogScanSize = False
    On Error GoTo Err_SetAnalogScanSize
    If MemHandle <> 0 Then  ''' in case we already allocated some
        mError = cbWinBufFree(MemHandle)
    End If
    MemHandle = cbWinBufAlloc(nSize)
    If MemHandle = 0 Then
        RaiseEvent cbDAS1200Error(0, "Failed To Allocate Memory For Scan Data")
    End If
    ReDim mScanData(0 To nSize - 1)
    ReDim mScanValues(0 To nSize - 1)
    SetAnalogScanSize = True

    Exit Function

Err_SetAnalogScanSize:
    HandleError
    SetAnalogScanSize = False
#End If

End Function

Public Function ScanAnalog(nLoChan As Integer, nHiChan As Integer, nCount As Long, nDataRate As Long, nDataRange As cbRange) As Boolean
    
    Dim i As Long
    Dim NumPoints As Long
    
    NumPoints = nCount * (nHiChan - nLoChan + 1)  ''' figure out the number of points used
    ''' scan data in from card to memory
    mError = cbAInScan(mBoardNumber, nLoChan, nHiChan, nCount, nDataRate, nDataRange, MemHandle, 0)
    If mError <> NOERRORS Then
        HandleError
        ScanAnalog = True
        Exit Function
    End If
    ''''  Copy data in memory to array
    mError = cbWinBufToArray(MemHandle, mScanData(0), 0, NumPoints)
    If mError <> NOERRORS Then
        HandleError
        ScanAnalog = True
        Exit Function
    End If
    ''' convert array data to engineering units
    For i = 0 To NumPoints - 1
        mError = cbToEngUnits(mBoardNumber, nDataRange, mScanData(i), mScanValues(i))
        If mError <> NOERRORS Then
            HandleError
            ScanAnalog = True
            Exit Function
        End If
    Next
    
End Function

Public Function ScanDataArray() As Single()
    
    ScanDataArray = mScanValues

End Function

Private Sub HandleError()
#If DISABLE_EQUIP_CHECKS = 1 Then
    Dim msgBoxMessage As String
    Dim msgBoxResponse As Integer
    
    msgBoxResponse = vbAbort

    If 0 <> Err.Number Then
        ' Unhandled error
        If "" <> callFunction Then
            msgBoxMessage = callFunction & " exited abnormally" & vbCrLf & vbCrLf & Err.Description
        Else
            msgBoxMessage = Err.Description
        End If

        If 53 = Err.Number Then
            ' Reinstall driver/dll
            mError = BADFILENAME
            msgBoxResponse = vbIgnore
        Else
            mError = BADBOARD
        End If
    Else
        If NOERRORS = mError Then
            ' No errors
            msgBoxResponse = vbOK
        Else
            ' InstaCal error
            msgBoxMessage = "Error controlling digital IO." & vbCrLf & vbCrLf
            If CFGFILENOTFOUND = mError Then
                ' Run InstaCal to configure board
                msgBoxResponse = vbIgnore
            ElseIf BADBOARD = mError Then
                ' Board not installed
                msgBoxResponse = vbIgnore
            End If
        End If
    End If
    
    If vbOK <> msgBoxResponse Then
        msgBoxResponse = MsgBox(msgBoxMessage, vbAbortRetryIgnore)
    End If

    If vbIgnore = msgBoxResponse Then
        If True = frmPassword.CheckPassword() Then
            ' System administrator override
            If vbOK = MsgBox("Are you sure?", vbOKCancel) Then
                Err.Clear   ' Clear Err object fields
                mError = NOERRORS
                g_bDIO_Installed = False
            Else
                msgBoxResponse = vbAbort
            End If
        End If
    ElseIf vbAbort = msgBoxResponse Then
        Err.Clear   ' Clear Err object fields
        ' g_bQuitProgram = True
    ElseIf vbOK <> msgBoxResponse Then
        msgBoxResponse = vbAbort
    End If
    
    If Err.Number <> 0 Then
        ' Retry
        Err.Clear   ' Clear Err object fields
    End If
#End If
    RaiseEvent cbDAS1200Error(mError, GetErrorString(mError))

End Sub

Public Function GetError() As String
    
    GetError = GetErrorString(mError)

End Function

Private Function GetErrorString(Error As Integer) As String
    
    Select Case Error
        Case NOERRORS
            GetErrorString = "No Errors"
        Case BADBOARD
            GetErrorString = "Bad Board"
        Case DEADDIGITALDEV
            GetErrorString = "Dead Digital Device"
        Case DEADCOUNTERDEV
            GetErrorString = "Dead Counter Device"
        Case DEADDADEV
            GetErrorString = "Dead DA Device"
        Case DEADADDEV
            GetErrorString = "Dead AD Device"
        Case NOTDIGITALCONF
            GetErrorString = "Not Digital Conf"
        Case NOTCOUNTERCONF
            GetErrorString = "Not Counter Conf"
        Case NOTDACONF
            GetErrorString = "Not a Conf"
        Case NOTADCONF
            GetErrorString = "Not AD Conf"
        Case NOTMUXCONF
            GetErrorString = "Not Mux Conf"
        Case BADPORTNUM
            GetErrorString = "Bad Port Number"
        Case BADCOUNTERDEVNUM
            GetErrorString = "Get Error String"
        Case BADDADEVNUM
            GetErrorString = "Bad DA Device Number"
        Case BADSAMPLEMODE
            GetErrorString = "Bad Sample Mode"
        Case BADINT
            GetErrorString = "Bad Int"
        Case BADADCHAN
            GetErrorString = "Bad Channel"
        Case BADCOUNT
            GetErrorString = "Bad Count"
        Case BADCNTRCONFIG
            GetErrorString = "Bad Counter Config"
        Case BADDAVAL
            GetErrorString = "Bad DA Value"
        Case BADDACHAN
            GetErrorString = "Bad DA Channel"
        Case ALREADYACTIVE
            GetErrorString = "Already Active"
        Case PAGEOVERRUN
            GetErrorString = "Page Overrun"
        Case BADRATE
            GetErrorString = "Bad Rate"
        Case COMPATMODE
            GetErrorString = "Compatable Mode"
        Case TRIGSTATE
            GetErrorString = "Trigger State"
        Case ADSTATUSHUNG
            GetErrorString = "AD STATUS HUNG"
        Case TOOFEW
            GetErrorString = "Too Few"
        Case OVERRUN
            GetErrorString = "Overrun"
        Case BADRANGE
            GetErrorString = "Bad Range"
        Case NOPROGGAIN
            GetErrorString = "No Programable Gain"
        Case BADFILENAME
            GetErrorString = "Bad File Name"
        Case DISKISFULL
            GetErrorString = "Disk is full"
        Case COMPATWARN
            GetErrorString = "COMPATWARN"
        Case BADPOINTER
            GetErrorString = "Bad Pointer"
        Case TOOMANYGAINS
            GetErrorString = "Too many Gains"
        Case RATEWARNING
            GetErrorString = "Rate Warner"
        Case CONVERTDMA
            GetErrorString = "Convert DMA"
        Case DTCONNECTERR
            GetErrorString = "DT Connect Err"
        Case FORECONTINUOUS
            GetErrorString = "FOREContinuous"
        Case BADBOARDTYPE
            GetErrorString = "Bad Board Type"
        Case WRONGDIGCONFIG
            GetErrorString = "Wrong Digital Config"
        Case NOTCONFIGURABLE
            GetErrorString = "Not Configurable"
        Case BADPORTCONFIG
            GetErrorString = "Bad Port config"
        Case BADFIRSTPOINT
            GetErrorString = "Bad First Point"
        Case EndOfFile
            GetErrorString = "End of File"
        Case NOT8254CTR
            GetErrorString = "not 8254 Ctr"
        Case NOT9513CTR
            GetErrorString = "Not 9513 Ctr"
        Case BADTRIGTYPE
            GetErrorString = "Bad Trig Type"
        Case BADTRIGVALUE
            GetErrorString = "Bad Trig Value"
        Case BADOPTION
            GetErrorString = "Bad Option"
        Case BADPRETRIGCOUNT
            GetErrorString = "Bad Trig Count"
        Case BADDIVIDER
            GetErrorString = "Bad Divider"
        Case BADSOURCE
            GetErrorString = "Bad Source"
        Case BADCOMPARE
            GetErrorString = "Bad Compare"
        Case BADTIMEOFDAY
            GetErrorString = "Bad Time of Day"
        Case BADGATEINTERVAL
            GetErrorString = "Bad Gate Interval"
        Case BADGATECNTRL
            GetErrorString = "Bad Gate Control"
        Case BADCOUNTEREDGE
            GetErrorString = "Bad Counter Edge"
        Case BADSPCLGATE
            GetErrorString = "BAD SPCL Gate"
        Case BADRELOAD
            GetErrorString = "Bad Reload"
        Case BADRECYCLEFLAG
            GetErrorString = "Bad Recycle Flag"
        Case BADBCDFLAG
            GetErrorString = "Bad CD Flag"
        Case BADDIRECTION
            GetErrorString = "Bad Direction"
        Case BADOUTCONTROL
            GetErrorString = "Bad out Control"
        Case BADBITNUMBER
            GetErrorString = "Bad Bit Number"
        Case NONEENABLED
            GetErrorString = "None Enabled"
        Case BADCTRCONTROL
            GetErrorString = "Bad Ctr Control"
        Case BADEXPCHAN
            GetErrorString = "Bad Exp Chan"
        Case WRONGADRANGE
            GetErrorString = "Wrong AD Range"
        Case OUTOFRANGE
            GetErrorString = "Out of range"
        Case BADTEMPSCALE
            GetErrorString = "Bad Temp Scale"
        Case BADERRCODE
            GetErrorString = "Bad Err Code"
        Case NOQUEUE
            GetErrorString = "No Queue"
        Case CONTINUOUSCOUNT
            GetErrorString = "Continuous Count"
        Case UNDERRUN
            GetErrorString = "Underrun"
        Case BADMEMMODE
            GetErrorString = "Bad ME Mode"
        Case FREQOVERRUN
            GetErrorString = "Freq Overrun"
        Case NOCJCCHAN
            GetErrorString = "No CJC Chan"
        Case BADCHIPNUM
            GetErrorString = "Bad Chip Number"
        Case DIGNOTENABLED
            GetErrorString = "Digital Not Enabled"
        Case CONVERT16BITS
            GetErrorString = "Convert 16 Bits"
        Case NOMEMBOARD
            GetErrorString = "No Memory Board"
        Case DTACTIVE
            GetErrorString = "DT Active"
        Case NOTMEMCONF
            GetErrorString = "Not Mem Conf"
        Case ODDCHAN
            GetErrorString = "ODD Chan"
        Case CTRNOINIT
            GetErrorString = "Ctr No Init"
        Case NOT8536CTR
            GetErrorString = "Not 8536 Ctr"
        Case FREERUNNING
            GetErrorString = "Free Running"
        Case INTERRUPTED
            GetErrorString = "Interrupted"
        Case NOSELECTORS
            GetErrorString = "No Selectors"
        Case NOBURSTMODE
            GetErrorString = "No Burst Mode"
        Case NOTWINDOWSFUNC
            GetErrorString = "Not Windows Function"
        Case NOTSIMULCONF
            GetErrorString = "Not Simul Conf"
        Case EVENODDMISMATCH
            GetErrorString = "Even Odd Mismatch"
        Case M1RATEWARNING
            GetErrorString = "M1Rate Warning"
        Case NOTRS485
            GetErrorString = "Not RS485"
        Case NOTDOSFUNC
            GetErrorString = "Not Dos Func"
        Case RANGEMISMATCH
            GetErrorString = "Range Mismatch"
        Case CLOCKTOOSLOW
            GetErrorString = "Clock Too Slow"
        Case BADCALFACTORS
            GetErrorString = "Bad Cal Factors"
        Case BADCONFIGTYPE
            GetErrorString = "Bad Config Type"
        Case BADCONFIGITEM
            GetErrorString = "Bad Config Item"
        Case NOPCMCIABOARD
            GetErrorString = "No PCMCIA Board"
        Case NOBACKGROUND
            GetErrorString = "No Background"
        Case STRINGTOOSHORT
            GetErrorString = "String Too Short"
        Case CONVERTEXTMEM
            GetErrorString = "Conver Text Mem"
        Case BADEUADD
            GetErrorString = "Bad EU Add"
        Case DAS16JRRATEWARNING
            GetErrorString = "DAS16JR Rate Warning"
        Case DAS08TOOLOWRATE
            GetErrorString = "DAS08 Too Low Rate"
        Case NOCONVERSIONNEEDED    ' 12 bit board without chan tags - converted in ISR
            GetErrorString = "No Conversion Needed"
        Case NOEXTCONTINUOUS
            GetErrorString = "No ext Continuous"
        Case INVALIDPRETRIGCONVERT   ' cbConvertPretirg called after cbPretrigScan failed
            GetErrorString = "Invalid Pre Trig Convert"
        Case BADCTRREG            ' Bad arg to CLoad for 9513 }
            GetErrorString = "Bad arg to CLoad for 9513"
        Case BADTRIGTHRESHOLD      ' Invalid trigger threshold specified in cbSetTrigger }
            GetErrorString = "Invalid trigger threshold specified in cbSetTrigger"
        Case BADPCMSLOTREF      ' Invalid trigger threshold specified in cbSetTrigger
            GetErrorString = "Invalid trigger threshold specified in cbSetTrigger"
        Case AMBIGPCMSLOTREF   ' Ambiguous PCM Card slot reference
            GetErrorString = "Ambiguous PCM Card slot reference"
        Case BADSENSORTYPE  ' Bad sensor type selected in Instacal
            GetErrorString = "Bad sensor type selected in Instacal"
        Case DELBOARDNOTEXIST      ' tried to delete board number which doesn't exist
            GetErrorString = "tried to delete board number which doesn't exist"
        Case NOBOARDNAMEFILE       ' board name file not found
            GetErrorString = "board name file not found"
        Case CFGFILENOTFOUND       ' configuration file not found
            GetErrorString = "configuration file not found"
        Case NOVDDINSTALLED      ' CBUL.386 device driver not installed
            GetErrorString = "CBUL.386 device driver not installed"
        Case NOWINDOWSMEMORY     ' No Windows memory available
            GetErrorString = "No Windows memory available"
        Case OUTOFDOSMEMORY     ' No DOS memory available
            GetErrorString = "No DOS memory available"
        Case OBSOLETEOPTION       ' Option on longer supporeted in cbGetConfig/cbSetConfig
            GetErrorString = "Option on longer supporeted in cbGetConfig/cbSetConfig"
        Case NOPCMREGKEY      ' No registry entry for this PCMCIA board
            GetErrorString = "No registry entry for this PCMCIA board"
        Case NOCBUL32SYS         ' CBUL32.SYS device driver not installed
            GetErrorString = "CBUL32.SYS device driver not installed"
        Case NODMAMEMEMORY                 ' No memory for device driver's DMA buffer
            GetErrorString = "No memory for device driver's DMA buffer"
        Case IRQNOTAVAILABLE     ' IRQ in use by another device
                GetErrorString = "IRQ in use by another device"
        Case Else
            GetErrorString = "Unknown Error"
    End Select
    
End Function

Private Sub Class_Terminate()
    
    If MemHandle <> 0 Then
        mError = cbWinBufFree(MemHandle)
    End If

End Sub
