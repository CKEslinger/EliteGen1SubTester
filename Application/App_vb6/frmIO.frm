VERSION 5.00
Begin VB.Form frmIO 
   Caption         =   "I/O"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   2070
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cb As clsDAS1200
Private m_IgnitionInput As Boolean
Private m_StarterInput As Boolean
Private m_RF_XmtChar As Boolean
Private m_GPS_Ext As Boolean
Private m_RF_Output As Boolean
Private m_IsTesterClosed As Boolean

Public Property Get IsTesterClosed() As Boolean

    Dim i As Integer

    i = cb.PortB
'If MsgBox("Port B = 0x" & Hex(i), vbOKCancel) = vbCancel Then EndProgram True
    i = i And &H80
    If i = 0 Then
        IsTesterClosed = True
    Else
        IsTesterClosed = False
    End If

End Property

Private Property Let IsTesterClosed(state As Boolean)
    
    m_IsTesterClosed = state

End Property

Public Property Get IgnitionInput() As Boolean
    
    IgnitionInput = m_IgnitionInput

End Property

Public Property Let IgnitionInput(value As Boolean)
    
    m_IgnitionInput = value
    If value = 0 Then  'Relay Open
        cb.PortA = (cb.PortA And &HFD)
        cb.WriteDIO CBDIOPORTA
    Else                'Relay Closed
        cb.PortA = (cb.PortA And &HFD) Or &H2
        cb.WriteDIO CBDIOPORTA
    End If

End Property

Public Property Get StarterInput() As Boolean
    
    StarterInput = m_StarterInput

End Property

Public Property Let StarterInput(value As Boolean)
    
    m_StarterInput = value
    If value = 0 Then  'Relay Open
        cb.PortA = (cb.PortA And &HFE)
        cb.WriteDIO CBDIOPORTA
    Else                'Relay Closed
        cb.PortA = (cb.PortA And &HFE) Or &H1
        cb.WriteDIO CBDIOPORTA
    End If

End Property

Public Property Get GPS_Ext() As Boolean
    
    GPS_Ext = m_GPS_Ext

End Property

Public Property Let GPS_Ext(value As Boolean)
    
    m_GPS_Ext = value
    If value = 0 Then  'Relay Open
        cb.PortA = cb.PortA And (Not &H80)
        cb.WriteDIO CBDIOPORTA
    Else                'Relay Closed
        cb.PortA = cb.PortA Or &H80
        cb.WriteDIO CBDIOPORTA
   
    End If

End Property
Public Property Get RF_XmtChar() As Boolean
    
    RF_XmtChar = m_RF_XmtChar

End Property

Public Property Let RF_XmtChar(value As Boolean)
    
    m_RF_XmtChar = value
    If value = 0 Then  'Relay Open
        cb.PortA = cb.PortA And (Not &H70)
        cb.WriteDIO CBDIOPORTA
    Else                'Relay Closed
        cb.PortA = cb.PortA Or &H70
        cb.WriteDIO CBDIOPORTA
   
    End If

End Property

Public Property Get RF_Output() As Boolean
    
    RF_Output = m_RF_Output

End Property

Public Property Let RF_Output(value As Boolean)
    
    Dim i As Integer

    m_RF_Output = value
    If value = 0 Then
        'Normal Mode
'        IgnitionInput = False
'        StarterInput = False
'        SerialComs.PSOutput = False
'        startup.Delay 1#
'        SerialComs.PSOutput = True
    Else
        ' TODO: Call AT function to start RF test
        '       Function should look for RF input for 3 seconds and record the input value that it sees
        '       Optional could also record the number of inputs that it sees

            RF_XmtChar = True
            startup.Delay 0.25
            RF_XmtChar = False
         
    End If

End Property

Public Property Get DUTRelayState() As Boolean
    
    Dim i As Double
    
    i = DUTRelayVoltage
    If i > 1 And i < 3 Then
        DUTRelayState = False
    Else
        DUTRelayState = True
    End If

End Property

Public Property Get SwitchState() As Boolean
    
    Dim i As Integer
    
    i = cb.PortB And &H80
    If i = 0 Then
        SwitchState = True
    Else
        SwitchState = False
    End If

End Property

Private Function ReadChannelMean(Chan As Integer, Range As cbRange, Optional numRead As Integer = 5) As Double

    Dim r() As Double
    Dim i As Integer
    ReDim r(numRead - 1) As Double
    Dim total As Double

    For i = 0 To numRead - 1
        cb.ReadAnalogIn Chan, Range
        r(i) = cb.AnalogIn(Chan)
    Next

    total = 0
    For i = 0 To (numRead - 1)
        total = total + r(i)
    Next i

    ReadChannelMean = CDbl(total / numRead)

End Function

Public Function DUTRelayVoltage() As Double
    
    DUTRelayVoltage = ReadChannelMean(DUT_RELAY_VOLATAGE, cbUP10V)

End Function

Public Function PowerSupply1_Voltage() As Double

    '   The AtoD channel clips at 10 volts, but the battery voltage should be
    '   set at 12 volts.  The battery voltage should, therefore, go through a
    '   voltage divider before going to the channel else what will appear on
    '   the diagnostic display will be clipped at 10.  If the battery voltage
    '   is divided by 2 then multiply the results here by 2 in order to track
    '   the battery voltage accurately.
    '
    PowerSupply1_Voltage = ReadChannelMean(PS1_VOLTAGE, cbUP10V) * 2

End Function

Public Function LedVoltage(ledColor As Integer) As Double
    If ledColor = 0 Then
    LedVoltage = ReadChannelMean(LED_GRN_VOLTAGE, cbUP1_25)
    Else
    LedVoltage = ReadChannelMean(LED_RED_VOLTAGE, cbUP1_25)
    End If
End Function
Public Function PowerSupply2_Voltage() As Double
    
    PowerSupply2_Voltage = ReadChannelMean(PS2_VOLTAGE, cbUP10V)

End Function

Public Function Init() As Boolean

    Dim cbError As Boolean

    On Error GoTo DAQError
    cb.Init 1
    cbError = cb.ConfigureDIOPort(CBDIOPORTA, cbPORTOUT)      ''' Configure port a to digital output
    If cbError Then GoTo DAQError
    cbError = cb.ConfigureDIOPort(CBDIOPORTB, cbPORTIN)       ''' Configure port b to digital input
    If cbError Then GoTo DAQError
    cbError = cb.ConfigureDIOPort(CBDIOPORTCL, cbPORTOUT)     ''' Configure port c low to digital output
    If cbError Then GoTo DAQError
    cbError = cb.ConfigureDIOPort(CBDIOPORTCH, cbPORTOUT)     ''' Configure port c high to digital output
    If cbError Then GoTo DAQError
    cb.SetAnalogScanSize Globals.DATA_COUNT                   ''' Initialize test parameters
    If cbError Then GoTo DAQError
    
    startup.Delay 0.1
    cb.PortA = 0
    cb.PortCL = 3
    cb.PortCH = 8
    cbError = cb.WriteDIO(CBDIOPORTA)
    If cbError Then GoTo DAQError
    cbError = cb.WriteDIO(CBDIOPORTCL)
    If cbError Then GoTo DAQError
    cbError = cb.WriteDIO(CBDIOPORTCH)
    If cbError Then GoTo DAQError
    Init = True
    
    'Initialize GPS external antenna activate variable
    m_GPS_Ext = False
    Exit Function
    
DAQError:
    MsgBox "Failed to find/access the DAQ card", vbCritical
    Init = False

End Function

Private Sub Form_Load()
    
    Set cb = New clsDAS1200

End Sub

Private Sub Form_Unload(Cancel As Integer)

    cb.PortA = 0
    cb.PortCL = 3
    cb.PortCH = 8
    cb.WriteDIO CBDIOPORTA
    cb.WriteDIO CBDIOPORTCL
    cb.WriteDIO CBDIOPORTCH
    Set cb = Nothing

End Sub

Public Function PortA() As Byte

    PortA = cb.PortA

End Function

Public Function PortB() As Byte

    PortB = cb.PortB

End Function

