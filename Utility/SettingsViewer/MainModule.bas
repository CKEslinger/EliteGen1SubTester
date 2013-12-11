Attribute VB_Name = "MainModule"
Option Explicit

Public fMainForm As frmMain

Public Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetCompName() As String
    Dim retVal As Long

    'Create a string buffer for the computer name
    Dim strCompName As String
    strCompName = Space$(255)
    
    'Retrieve the Computer name
    retVal = GetComputerName(strCompName, 255)
    
    'Remove the trailing null character from the string
    GetCompName = Left$(strCompName, InStr(strCompName, vbNullChar) - 1)
End Function

Sub Main()
    Set fMainForm = New frmMain
   
    fMainForm.Show
   
End Sub

