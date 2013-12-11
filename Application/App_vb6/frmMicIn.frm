VERSION 5.00
Begin VB.Form frmMicIn 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Spectrum Analyzer"
   ClientHeight    =   6225
   ClientLeft      =   4335
   ClientTop       =   2430
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   415
   ScaleMode       =   0  'User
   ScaleWidth      =   425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameStimulus 
      Caption         =   "Stimulus"
      Height          =   1095
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Width           =   5655
      Begin VB.Frame Frame_DeviceCOM 
         Caption         =   "Elite COMs"
         Enabled         =   0   'False
         Height          =   735
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   1935
         Begin VB.OptionButton Option_COM_Port 
            Caption         =   "COM Port Closed"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option_COM_Port 
            Caption         =   "COM Port Open"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame_Pwr 
         Caption         =   "Power Supply"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton Option_Pwr 
            Caption         =   "Power Off"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option_Pwr 
            Caption         =   "Power On"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check_Chirp 
         Caption         =   "Device Chirp"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "OK"
      Height          =   336
      Left            =   5016
      TabIndex        =   6
      Top             =   5760
      Width           =   984
   End
   Begin VB.PictureBox picScope 
      BackColor       =   &H80000009&
      ForeColor       =   &H80000002&
      Height          =   3690
      Left            =   360
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   372
      TabIndex        =   4
      Top             =   720
      Width           =   5640
      Begin VB.PictureBox picScopeBuf 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000002&
         Height          =   336
         Left            =   60
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   5
         Top             =   60
         Visible         =   0   'False
         Width           =   336
      End
   End
   Begin VB.TextBox txtStrength 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtFreq 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.ComboBox comboDevices 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3108
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "&Start"
      Height          =   336
      Left            =   360
      TabIndex        =   1
      Top             =   5760
      Width           =   984
   End
End
Attribute VB_Name = "frmMicIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---------------------------------------------------------------------------
'          Title | ctSpectrum Analyzer frmBase.Frm (formerly DeethSA)
'  Version / Rev | Version 2
'        Part Of | N/A
'  Original date | Unknown
'  Last Modified | January 18 / 2003
'         Status | Unknown
'         Author | Murphy McCauley
'   Author Email | MMcCauley@FullSpectrum.Com (Put "[ct]" in Subject)
'        Website | http://www.constantthought.com
'   Dependencies | None
' ---------------------------------------------------------------------------
' A simple audio spectrum analyzer.
' Opens a waveform audio device for 16-bit mono input, gets chunks of
' audio, runs the FFT on them, and displays the output in a little window.
' Demonstrates an easy and fairly fast way to do graphics double-buffering
' with a hidden picturebox, audio input, FFT usage, etc.
' ---------------------------------------------------------------------------
' I make no promises that this code will do what you want or even that it
' will do what I meant it to do.  You use it at your own risk.
' ---------------------------------------------------------------------------

' [January 18 / 2003]
' New version has better coding style, bugs fixed, output improved (true
' spectra power calculated, better scaling, Blackman window applied), is
' resizable, is probably faster (the FFT certainly is), uses a double
' buffer for the audio input (less waiting), has generally better code,
' and it's now very easy to adjust the sample chunk size and the sample
' rate (see below).


Option Explicit


' ---------------------------------------------------------------------------
' Adjustable stuff
' ---------------------------------------------------------------------------

' You can set the sample rate (pcSampleRate, below), as well as the number
' of samples to be analyzed and displayed at once (pNumBits in modAudioFFT).

' Must be one of 11025, 22050, 44100
Private Const pcSampleRate As Long = 44100

' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' Module level variables
' ---------------------------------------------------------------------------


' Set to true when the form should unload
Private pUnloading As Boolean

' Handle of the open audio device or 0 if no device is open
Private pDevHandle As Long

' True when the visualization is running
Private pVisualizing As Boolean

' Divisor for the scale of the analyzer's display
Private pDivisor As Long

' Current height of the scope (faster than accessing the property)
Private pScopeHeight As Long
                           
' Horizontal step size
Private pStepX As Double

' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' Declarations
' ---------------------------------------------------------------------------

Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Private Const pcWaveHdrLen As Long = 32

Private Type WaveInCaps
    ManufacturerID As Integer       ' wMid
    ProductID As Integer            ' wPid
    DriverVersion As Long           ' MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte    ' szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

' These are the only formats we can use
Private Const WAVE_FORMAT_1M16 = &H4&           ' 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2M16 = &H40&          ' 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4M16 = &H400&         ' 44.1   kHz, Mono,   16-bit

Private Const WAVE_FORMAT_PCM = 1
Private Const WHDR_DONE As Long = &H1& ' done bit

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, WaveHdrPointer As WaveHdr, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, WaveHdrPointer As WaveHdr, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, WaveHdrPointer As WaveHdr, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

' ---------------------------------------------------------------------------


Private Function InitDevices() As Long
    ' Fill the comboDevices box with all the compatible audio input devices.
    ' Returns the number of compatible devices.
    
    Dim Format As Long

    Select Case pcSampleRate
        Case 11025
            Format = WAVE_FORMAT_1M16
        Case 22050
            Format = WAVE_FORMAT_2M16
        Case 44100
            Format = WAVE_FORMAT_4M16
        Case Else
            Err.Raise 10101, , "Internal error #1"
    End Select
    
    Dim Caps As WaveInCaps, Which As Long
    comboDevices.Clear
    For Which = 0 To waveInGetNumDevs - 1
        waveInGetDevCaps Which, VarPtr(Caps), Len(Caps)
        If Caps.Formats And Format Then
            ' It supports our format
            comboDevices.AddItem StrConv(Caps.ProductName, vbUnicode)
            comboDevices.ItemData(comboDevices.ListCount - 1) = Which
            InitDevices = InitDevices + 1
        End If
    Next
    
    If InitDevices Then
        comboDevices.ListIndex = 0
    End If
End Function

Private Sub Check_Chirp_Click()
    Dim i As Integer
    Dim chkChirp As Boolean
    
    chkChirp = (vbChecked = Check_Chirp.value)
    
    ' Enable/Disable the controls
    Frame_Pwr.Enabled = chkChirp
    For i = 0 To 1
        Option_Pwr(i).Enabled = chkChirp
    Next i
    
    If True = chkChirp Then
        ' Set the Power Off/On control values
        If g_fIO.PowerSupply1_Voltage < 1 And g_fIO.PowerSupply2_Voltage < 1 Then
            If False = Option_Pwr(0).value Then
                Option_Pwr_Click (0)
            End If
        Else
            If False = Option_Pwr(1).value Then
                Option_Pwr_Click (1)
            End If
        End If
        
        ' Set the COM Port Closed/Open control values
        If False = g_fMainTester.ComPortToElite.PortOpen Then
            If False = Option_COM_Port(0).value Then
                Option_COM_Port_Click (0)
            End If
        Else
            If False = Option_COM_Port(1).value Then
                Option_COM_Port_Click (1)
            End If
        End If
    Else
    
        ' Set the COM Port Closed/Open control value to closed
        If True = Option_COM_Port(1).value Then
            Option_COM_Port_Click (0)
        End If
        If True = g_fMainTester.ComPortToElite.PortOpen Then
            g_fMainTester.ComPortToElite.PortOpen = False
        End If
    
        ' Set the Power Off/On control value to off
        If True = Option_Pwr(1).value Then
            Option_Pwr_Click (0)
        End If
        If g_fIO.PowerSupply1_Voltage > 1 Or g_fIO.PowerSupply2_Voltage > 1 Then
            SerialComs.PSOutput = False
        End If
    End If

End Sub

Private Sub Form_Activate()

    g_bDiagnosticActive = True

End Sub

Private Sub Form_Deactivate()
    g_bDiagnosticActive = False

    If vbChecked = Check_Chirp.value Then
        Check_Chirp.value = 0
    End If

End Sub

Private Sub Form_Load()

    If InitDevices = 0 Then
        MsgBox "You don't have any compatible audio input devices!", vbCritical
        
        ' Disable all controls
        Dim C As Control
        For Each C In Controls
            C.Enabled = False
        Next
    End If
    
    InitWavIn
    
    ' Initialize some stuff for the FFT
    InitFFT
    
    
    ' I leave double buffers a funky color (so I can see them).  Set it to the proper color.
    picScopeBuf.BackColor = picScope.BackColor
    
    
End Sub

Private Sub Form_Resize()
    ' Sort of a crazy function. =)
    Dim key As Single

    On Error Resume Next

    key = Me.ScaleHeight - FrameStimulus.Height - btnStart.Height - 15
    FrameStimulus.Top = key
    
    key = key + FrameStimulus.Height + 9
    btnStart.Top = key
    btnStop.Top = key
    txtFreq.Top = key
    txtStrength.Top = key

    comboDevices.Width = Me.ScaleWidth - comboDevices.Left * 2
    If Err Then
        Err.Clear
        comboDevices.Visible = False
    Else
        comboDevices.Visible = True
    End If
    
    picScope.Width = Me.ScaleWidth - picScope.Left * 2
    picScope.Height = key - picScope.Top - 10
    If Err Or picScope.Height < 4 Then
        picScope.Visible = False
    Else
        picScope.Visible = True
    End If

    If key < comboDevices.Top + comboDevices.Height Then
        FrameStimulus.Visible = False
        btnStart.Visible = False
        btnStop.Visible = False
    Else
        FrameStimulus.Visible = True
        btnStart.Visible = True
        btnStop.Visible = True
    End If
    
    picScopeBuf.Width = picScope.Width - (picScope.Width - picScope.ScaleWidth)
    picScopeBuf.Height = picScope.Height - (picScope.Height - picScope.ScaleHeight)
    
    pScopeHeight = picScopeBuf.Height
    pStepX = picScope.Width / (NumberOfSamples \ 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If pDevHandle <> 0 Then
        DoStop
        If pVisualizing = True Then
            Cancel = 1
            pUnloading = True
        End If
    End If
    
End Sub

Private Sub InitWavIn()
    
    Dim L As Long
    Dim bInit
    Dim WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 1
        .SamplesPerSec = pcSampleRate
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    On Error GoTo InitWavInError
    
    L = waveInOpen(pDevHandle, comboDevices.ItemData(comboDevices.ListIndex), VarPtr(WaveFormat), 0, 0, 0)
    If 0 <> pDevHandle Then
        Exit Sub
    End If

InitWavInError:
    Call MsgBox("Wave input device didn't open!  Please Restart Test Software.", vbExclamation, "Ack!")
End Sub

Public Function RecordWav() As Boolean

    Dim X As Long, ThisX As Double, NextX As Double
    Dim i As Integer

    '   Declaring as Static causes the compiler's optimizer to not go crazy
    '   'optimizing' it.  Sort of like a C "volatile" -- and for the same
    '   reason.  Since some of the values in WaveX change outside the normal
    '   flow of the program, the optimizations can cause problems.  For
    '   example, the loop on dwFlags will lock the program.
    Static Wave1 As WaveHdr
    Static Wave2 As WaveHdr

    Dim BufferOne As Boolean ' Used to flip back and forth between buffers
    Dim InData1(0 To NumberOfSamples - 1) As Integer
    Dim InData2(0 To NumberOfSamples - 1) As Integer

    Dim RealOut(0 To NumberOfSamples - 1) As Double
    Dim ImaginaryOut(0 To NumberOfSamples - 1) As Double
    Dim YValue(0 To NumberOfSamples / 2 - 1) As Double
    Dim ymax As Double, x_ymax As Double
    Dim xmax As Long

    RecordWav = False
    If 0 = pDevHandle Then
        Exit Function
    Else
        waveInStart pDevHandle
    End If

    btnStop.Enabled = True
    btnStart.Enabled = False
    comboDevices.Enabled = False
    BufferOne = True

    '   I now use two buffers.  This way one can be left to be filled while the
    '   other is being analyzed and displayed.  When it's time to analyze and
    '   display the next set of data, we don't have to wait for it as long
    '   (heck, it might even be full already!).

    pVisualizing = True

    Wave1.lpData = VarPtr(InData1(0))
    Wave2.lpData = VarPtr(InData2(0))
    Wave1.dwBufferLength = NumberOfSamples * 2
    Wave2.dwBufferLength = NumberOfSamples * 2
    Wave1.dwFlags = 0
    Wave2.dwFlags = 0

    waveInPrepareHeader pDevHandle, Wave1, pcWaveHdrLen
    waveInAddBuffer pDevHandle, Wave1, pcWaveHdrLen

    ymax = g_dSoundLevel
    xmax = g_dSoundFreq
    x_ymax = 0
    'Do
    For i = 1 To 1
        If BufferOne Then
            ' Wait for this buffer to be ready if necessary
            Do Until Wave1.dwFlags And WHDR_DONE
                ' Just wait for the block to be done
            Loop
            ' Leave the other buffer to be filled
            waveInPrepareHeader pDevHandle, Wave2, pcWaveHdrLen
            waveInAddBuffer pDevHandle, Wave2, pcWaveHdrLen

            waveInUnprepareHeader pDevHandle, Wave1, pcWaveHdrLen

            ' Analyze the data
            AudioFFT InData1, RealOut, ImaginaryOut
        Else
            ' Wait for this buffer to be ready if necessary
            Do Until Wave2.dwFlags And WHDR_DONE
                ' Just wait for the block to be done
            Loop

            ' Leave the other buffer to be filled
            waveInPrepareHeader pDevHandle, Wave1, pcWaveHdrLen
            waveInAddBuffer pDevHandle, Wave1, pcWaveHdrLen

            waveInUnprepareHeader pDevHandle, Wave2, pcWaveHdrLen

            ' Analyze the data
            AudioFFT InData2, RealOut, ImaginaryOut
        End If
        ' Switch buffer state
        BufferOne = Not BufferOne

        picScopeBuf.Cls
        ThisX = 0
        pDivisor = 1
        'For X = 0 To NumberOfSamples \ 2 - 1
        For X = 1200 To 2100
            NextX = ThisX + pStepX * 6
            YValue(X) = (Sqr(RealOut(X) * RealOut(X) + ImaginaryOut(X) * ImaginaryOut(X)) / pDivisor / NumberOfSamples)
            If YValue(X) > ymax Then
              ymax = YValue(X)
              xmax = X
              x_ymax = X * (pcSampleRate / NumberOfSamples)
            End If

            picScopeBuf.Line (ThisX, pScopeHeight)-(NextX, pScopeHeight - pScopeHeight / 40 * (Sqr(RealOut(X) * RealOut(X) + ImaginaryOut(X) * ImaginaryOut(X)) / pDivisor / NumberOfSamples)), , BF

            ThisX = NextX
        Next
        txtFreq.Text = Format(ymax, "0.000")
        txtStrength.Text = CStr(xmax)
        picScope.Picture = picScopeBuf.Image ' Display the double-buffer
    Next i
    DoEvents
    waveInStop pDevHandle

    g_dSoundLevel = ymax
    g_dSoundFreq = xmax

    ' Wait for and unprepare the buffer that we left to be filled
    If BufferOne Then
        Do Until Wave1.dwFlags And WHDR_DONE
            ' Just wait for the block to be done
        Loop
        waveInUnprepareHeader pDevHandle, Wave1, pcWaveHdrLen
    Else
        Do Until Wave2.dwFlags And WHDR_DONE
            ' Just wait for the block to be done
        Loop
        waveInUnprepareHeader pDevHandle, Wave2, pcWaveHdrLen
    End If

    pVisualizing = False

    btnStop.Enabled = True
    btnStart.Enabled = True
    comboDevices.Enabled = True

    If pUnloading Then
        Unload Me
    End If
    RecordWav = True

End Function

Private Sub btnStart_Click()

    g_dSoundLevel = 0
    g_dSoundFreq = 0
    If vbChecked = Check_Chirp.value Then
'        If True = g_fMainTester.ComPortToElite.PortOpen Then
        If True = Option_COM_Port(1).value Then
            If SendATCommand(etELITE_STATES.eElite_CHECK_BUZZER, 1) Then
                RecordWav
            End If
        End If
    End If
   
End Sub

Private Sub btnStop_Click()

    g_bDiagnosticActive = False
    Me.Hide

End Sub

Private Sub DoStop()

    waveInReset pDevHandle
    waveInClose pDevHandle
    pDevHandle = 0
    btnStop.Enabled = False
    btnStart.Enabled = True
    comboDevices.Enabled = True

End Sub

Private Sub Option_COM_Port_Click(Index As Integer)

    If False = Option_COM_Port(Index).value Then
        Option_COM_Port(Index).value = True
    End If
    
    If Index = 0 Then
        Option_COM_Port(0).value = SerialComs.OpenEliteComPort(False, True)
    Else
        Option_COM_Port(1).value = SerialComs.OpenEliteComPort(True, True)
    End If
End Sub

Private Sub Option_Pwr_Click(Index As Integer)
    Dim pwr As Boolean
    If g_fIO.PowerSupply1_Voltage < 1 And g_fIO.PowerSupply2_Voltage < 1 Then
        pwr = False
    Else
        pwr = True
    End If

    If False = Option_Pwr(Index).value Then
        Option_Pwr(Index).value = True
    End If

    If Index = 1 Then
        If False = pwr Then
            SerialComs.PSOutput = True
            pwr = True
        End If
    
    Else
        If True = pwr Then
            SerialComs.PSOutput = False
            pwr = False
        End If
    End If

    Dim i
    Frame_DeviceCOM.Enabled = pwr
    For i = 0 To 1
        Option_COM_Port(i).Enabled = pwr
    Next i
End Sub
