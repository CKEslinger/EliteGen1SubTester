Attribute VB_Name = "modAudioFFT"
' ---------------------------------------------------------------------------
'          Title | modAudioFFT.Bas
'  Version / Rev | Rev 2
'        Part Of | N/A
'  Original date | August 14 / 1999
'  Last Modified | January 18 / 2003
'         Status | Unknown
'         Author | Murphy McCauley
'   Author Email | MMcCauley@FullSpectrum.Com (Put "[ct]" in Subject)
'        Website | http://www.constantthought.com
'   Dependencies | None
' ---------------------------------------------------------------------------
' Taken from my modFourierTransform.Bas, but modified to only work with
' fixed parameters (for speed).
' ---------------------------------------------------------------------------
' I make no promises that this code will do what you want or even that it
' will do what I meant it to do.  You use it at your own risk.
' ---------------------------------------------------------------------------

Option Explicit

' ---------------------------------------------------------------------------
' Adjustable stuff
' ---------------------------------------------------------------------------

' This will change the sample window size.  pNumBits = LogN(2, NumSamples)
'Private Const pNumBits As Long = 15
Private Const pNumBits As Long = 14

' ---------------------------------------------------------------------------


' ---------------------------------------------------------------------------
' Misc.
' ---------------------------------------------------------------------------

' Normally this is Private, but DeethSA's frmBase uses it.
Public Const NumberOfSamples As Long = 2 ^ pNumBits

Private Const PI As Double = 3.14159265358979

Private pReversals(0 To NumberOfSamples - 1) As Long
Private pWindow(0 To NumberOfSamples - 1) As Double

' ---------------------------------------------------------------------------


Sub InitFFT()
    Dim i As Long
    For i = 0 To NumberOfSamples - 1
        pReversals(i) = ReverseBits(i, pNumBits)
        pWindow(i) = BlackmanWindow(i, NumberOfSamples)
    Next
End Sub

Sub AudioFFT(AudioIn() As Integer, RealOut() As Double, ImaginaryOut() As Double)
    ' Performs the FFT on audio data.
    ' Original author: Don Cross
    
    ' AudioIn() is audio input (16 bit PCM samples).
    ' Use Power(X) = Sqr(RealOut(X) ^ 2 + ImaginaryOut(X) ^ 2)
    
    Dim i As Long, j As Long, K As Long, L As Long, N As Long, BlockSize As Long, BlockEnd As Long

    Dim DeltaAngle As Double, DeltaAr As Double
    Dim Alpha As Double, Beta As Double
    Dim TR As Double, TI As Double, AR As Double, AI As Double
    
    For i = 0 To NumberOfSamples - 1
        j = pReversals(i)
        RealOut(j) = AudioIn(i) * pWindow(i)
        ImaginaryOut(j) = 0 ' Faster to ZeroMemory() it
    Next
    'ZeroMemory ImaginaryOut(0), 8 * NumberOfSamples
    
    BlockEnd = 1
    BlockSize = 2
    
    For L = 0 To pNumBits - 1
        DeltaAngle = (-2 * PI) / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2# * Alpha * Alpha
        Beta = Sin(DeltaAngle)
        
        For i = 0 To NumberOfSamples - 1 Step BlockSize
            AR = 1#
            AI = 0#
            
            j = i
            For N = 0 To BlockEnd - 1
                K = j + BlockEnd
                TR = AR * RealOut(K) - AI * ImaginaryOut(K)
                TI = AI * RealOut(K) + AR * ImaginaryOut(K)
                RealOut(K) = RealOut(j) - TR
                RealOut(j) = RealOut(j) + TR
                ImaginaryOut(K) = ImaginaryOut(j) - TI
                ImaginaryOut(j) = ImaginaryOut(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - (Alpha * AI - Beta * AR)
                AR = AR - DeltaAr
                j = j + 1
            Next
        Next
        
        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Next
End Sub

Function BlackmanWindow(ByVal Index As Long, ByVal NumberOfSamples As Long) As Double
    ' Good for audio
    BlackmanWindow = 0.4 - 0.5 * Cos(2 * PI * Index / (NumberOfSamples - 1)) + 0.08 * Cos(4 * PI * Index / (NumberOfSamples - 1))
End Function

Function HannWindow(ByVal Index As Long, ByVal NumberOfSamples As Long) As Double
    ' Often referred to (wrongly) as the Hanning window.
    HannWindow = 0.5 - 0.5 * Cos(2 * PI * Index / (NumberOfSamples - 1))
End Function

Function BartlettHannWindow(ByVal Index As Long, ByVal NumberOfSamples As Long) As Double
    ' Modified Bartlett-Hann window
    BartlettHannWindow = 0.62 - 0.48 * Abs((Index / (NumberOfSamples - 1)) - 0.5) + 0.38 * Cos(2 * PI * ((Index / (NumberOfSamples - 1)) - 0.5))
End Function

Function HammingWindow(ByVal Index As Long, ByVal NumberOfSamples As Long) As Double
    HammingWindow = 0.54 - 0.46 * Cos(2 * PI * (Index / (NumberOfSamples - 1)))
End Function

Private Function ReverseBits(ByVal Index As Long, ByVal NumBits As Long) As Long
    ' Reverses a bit pattern.
    ' i.e. where NumBits is 5, decimal 3 is binary 00011.  ReverseBits(3, 5) returns 24.
    ' Decimal 24 is binary 11000.
    
    Dim i As Long, Rev As Long

    Rev = 0
    For i = 0 To NumBits - 1
        Rev = (Rev * 2) Or (Index And 1)
        Index = Index \ 2
    Next
    
    ReverseBits = Rev
End Function

