Option Explicit

' Class Module (Name) = MRRunModel. No CorelDRAW / UserForm dependencies.
Private pNames As Variant
Private pStatus() As String
Private pSeconds() As Double
Private pCurrent As Long
Private pCount As Long
Private pState As String
Private pError As String

Public Sub Prepare(ByVal names As Variant)
    Dim i As Long
    pNames = names
    pCount = UBound(names) - LBound(names) + 1
    ReDim pStatus(0 To pCount - 1)
    ReDim pSeconds(0 To pCount - 1)
    For i = 0 To pCount - 1
        pStatus(i) = "PENDING"
    Next i
    pCurrent = -1
    pError = vbNullString
    pState = "ready"
End Sub

Public Sub BeginNext()
    If pState <> "ready" And pState <> "waiting" Then Err.Raise 5, "MacroRunner", "Antrean belum siap dilanjutkan."
    pCurrent = pCurrent + 1
    pStatus(pCurrent) = "RUNNING"
    pState = "running"
End Sub

Public Sub CompleteCurrent(ByVal seconds As Double, ByVal errorNumber As Long, ByVal description As String)
    Dim i As Long
    If pState <> "running" Then Exit Sub
    pSeconds(pCurrent) = seconds
    If errorNumber <> 0 Then
        pStatus(pCurrent) = "ERROR"
        pError = "Error " & CStr(errorNumber) & ": " & description
        For i = pCurrent + 1 To pCount - 1
            pStatus(i) = "SKIPPED"
        Next i
        pState = "finished"
    Else
        pStatus(pCurrent) = "OK"
        If pCurrent = pCount - 1 Then
            pState = "finished"
        Else
            pState = "waiting"
        End If
    End If
End Sub

Public Property Get State() As String
    State = pState
End Property

Public Property Get CurrentName() As String
    If pCurrent >= 0 And pCurrent < pCount Then CurrentName = CStr(pNames(pCurrent))
End Property

Public Property Get LastError() As String
    LastError = pError
End Property

Public Function StatisticsText() As String
    Dim i As Long, nameWidth As Long, prefixWidth As Long
    Dim text As String, duration As String, total As Double
    If pCount = 0 Then Exit Function
    nameWidth = 5
    For i = 0 To pCount - 1
        If Len(CStr(pNames(i))) > nameWidth Then nameWidth = Len(CStr(pNames(i)))
    Next i
    prefixWidth = Len(Format$(pCount, "00")) + 3 + nameWidth + 3 + 7 + 3
    For i = 0 To pCount - 1
        duration = "-"
        If pStatus(i) = "OK" Or pStatus(i) = "ERROR" Then
            duration = SecondsText(pSeconds(i)) & " s"
            total = total + pSeconds(i)
        End If
        If i > 0 Then text = text & vbCrLf
        text = text & Pad(Format$(i + 1, "00"), Len(Format$(pCount, "00"))) & " | " & _
            Pad(CStr(pNames(i)), nameWidth) & " | " & Pad(pStatus(i), 7) & " | " & duration
    Next i
    text = text & vbCrLf & String$(prefixWidth + 10, "-") & vbCrLf
    text = text & Pad("Total", prefixWidth) & SecondsText(total) & " s"
    StatisticsText = text
End Function

Private Function Pad(ByVal value As String, ByVal width As Long) As String
    Pad = value & Space$(width - Len(value))
End Function

Private Function SecondsText(ByVal value As Double) As String
    SecondsText = Replace(Format$(value, "0.000"), ",", ".")
End Function
