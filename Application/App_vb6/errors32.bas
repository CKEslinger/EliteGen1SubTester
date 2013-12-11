Attribute VB_Name = "ERRORS32"
'
' ERRORS32.BAS
'
 
Option Explicit

DefInt A-Z

Dim Text As String * 120

Function GetXMODEMErrorText(ByVal ErrorCode As Long) As String

    Dim Win32Err As Long

    Text = Space(120)
    Select Case ErrorCode
        Case IE_BADID
            GetXMODEMErrorText = "XMODEM ERROR: Invalid COM port name"

        Case IE_OPEN
            GetXMODEMErrorText = "XMODEM ERROR: COM port already open"

        Case IE_NOPEN
            GetXMODEMErrorText = "XMODEM ERROR: Cannot open COM port"

        Case IE_MEMORY
            GetXMODEMErrorText = "XMODEM ERROR: Unable to allocate memory"

        Case IE_DEFAULT
            GetXMODEMErrorText = "XMODEM ERROR: Error in default parameters"

        Case IE_HARDWARE
            GetXMODEMErrorText = "XMODEM ERROR: Hardware not present"

        Case IE_BYTESIZE
            GetXMODEMErrorText = "XMODEM ERROR: Unsupported byte size"

        Case IE_BAUDRATE
            GetXMODEMErrorText = "XMODEM ERROR: Unsupported baud rate"

        Case WSC_NO_DATA
            GetXMODEMErrorText = "XMODEM ERROR: No data"

        Case WSC_RANGE
            GetXMODEMErrorText = "XMODEM ERROR: Parameter out of range"

        Case WSC_ABORTED
            GetXMODEMErrorText = "XMODEM ERROR: Evaluation version corrupted"

        Case WSC_EXPIRED
            GetXMODEMErrorText = "XMODEM ERROR: Evaluation version expired or SioKeyCode not called"

        Case WSC_BUFFERS
            GetXMODEMErrorText = "XMODEM ERROR: Cannot allocate memory for buffers"

        Case WSC_THREAD
            GetXMODEMErrorText = "XMODEM ERROR: Cannot start thread"

        Case WSC_KEYCODE
            GetXMODEMErrorText = "XMODEM ERROR: Bad key code"

        Case WSC_WIN32ERR:
            Win32Err = SioWinError(Text, 120)
            GetXMODEMErrorText = Text

        Case Else
            GetXMODEMErrorText = "XMODEM ERROR: Error code " + Str$(ErrorCode)
      End Select

End Function

Sub SayError(F As Form, ByVal ErrorCode)

    F.Print GetXMODEMErrorText(ErrorCode)

End Sub
