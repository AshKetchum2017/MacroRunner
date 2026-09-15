Option Explicit

' Class Module (Name) = MRStopwatch. Currency holds the 64-bit API values;
' its scale cancels when the counter difference is divided by frequency.
#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private pFrequency As Currency
Private pStart As Currency
Private pTimerStart As Double
Private pUseCounter As Boolean
Private pStopped As Boolean
Private pElapsed As Double

Public Sub Start()
    pTimerStart = Timer
    pStopped = False
    pElapsed = 0#
    pUseCounter = False
    On Error GoTo TimerOnly
    If QueryPerformanceFrequency(pFrequency) = 0 Then Exit Sub
    If pFrequency <= 0 Then Exit Sub
    If QueryPerformanceCounter(pStart) = 0 Then Exit Sub
    pUseCounter = True
TimerOnly:
End Sub

Public Function StopSeconds() As Double
    Dim finish As Currency
    If pStopped Then
        StopSeconds = pElapsed
        Exit Function
    End If
    On Error GoTo TimerOnly
    If pUseCounter Then
        If QueryPerformanceCounter(finish) <> 0 Then
            pElapsed = CDbl(finish - pStart) / CDbl(pFrequency)
            If pElapsed >= 0# Then GoTo Finished
        End If
    End If
TimerOnly:
    pElapsed = Timer - pTimerStart
    If pElapsed < 0# Then pElapsed = pElapsed + 86400#
Finished:
    pStopped = True
    StopSeconds = pElapsed
End Function
