Attribute VB_Name = "Globals"



Public Const LF = 10
Public Const CR = 13
Public Const NUL = 0
Public Const SP = 32 'space

Public IncomingString As String
Public FinishedString As String

Public LifeTest As Boolean

Public ActiveCommPort(1 To 16) As String

Public Const MAX_COMM_PORTS = 8

Public SetTopSpeed As Integer

'used for timing functions
Declare Function timeGetTime Lib "winmm.dll" () As Long '32 bit app.
Declare Function GetTickCount Lib "kernel32" () As Long '32 bit app.

'captures current time in millisec from computer BIOS
Function GetCurrentMSEC() As Long
 GetCurrentMSEC = timeGetTime
End Function

'Used for forcing delays within the program.
'msec delays
Public Function delay(ByVal MS As Long) As Integer
Dim j&, jj&
  j = GetCurrentMSEC
  Do Until jj > MS
    DoEvents
    jj = GetCurrentMSEC - j
  Loop
End Function
'Use following code usage format...
' starttimer = GetCurrentMSEC
'
'If GetCurrentMSEC - starttimer > Recipe(73).Val Then




