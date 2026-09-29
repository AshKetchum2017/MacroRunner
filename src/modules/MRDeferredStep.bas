Option Explicit

' Standard Module (Name) = MRDeferredStep. One-shot dispatch after a form callback unwinds.
#If VBA7 Then
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal uIDEvent As LongPtr) As Long
Private pTimer As LongPtr
#Else
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal uIDEvent As Long) As Long
Private pTimer As Long
#End If
Private pOwner As MRPresenter

Public Function ScheduleMRStep(ByVal owner As MRPresenter) As Boolean
    If Not pOwner Is Nothing Then Exit Function
    Set pOwner = owner
    pTimer = SetTimer(0, 0, 15, AddressOf MRStepTimer)
    If pTimer = 0 Then
        Set pOwner = Nothing
    Else
        ScheduleMRStep = True
    End If
End Function

Public Sub CancelMRStep(ByVal owner As MRPresenter)
    If pOwner Is Nothing Then Exit Sub
    If Not pOwner Is owner Then Exit Sub
    If pTimer <> 0 Then KillTimer 0, pTimer
    pTimer = 0
    Set pOwner = Nothing
End Sub

#If VBA7 Then
Public Sub MRStepTimer(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
                        ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub MRStepTimer(ByVal hWnd As Long, ByVal uMsg As Long, _
                        ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    Dim owner As MRPresenter
    If idEvent <> pTimer Then Exit Sub
    KillTimer 0, pTimer
    pTimer = 0
    Set owner = pOwner
    Set pOwner = Nothing
    If owner Is Nothing Then Exit Sub
    On Error GoTo Failed
    owner.ResumeDeferredStep
    Exit Sub
Failed:
    owner.DeferredStepFailed Err.Number, Err.Description
End Sub
