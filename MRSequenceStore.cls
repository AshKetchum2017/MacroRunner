Option Explicit

' Class Module (Name) = MRSequenceStore.
Private Const REG_APP As String = "RinCorelMacros"
Private Const REG_SECTION As String = "MacroRunner"
Private Const REG_KEY As String = "SequencesV1"

Public Function LoadSequences() As Collection
    Dim result As New Collection
    Dim parser As New MRSequenceParser
    Dim payload As String, lines As Variant, line As Variant
    payload = GetSetting(REG_APP, REG_SECTION, REG_KEY, vbNullString)
    If Len(payload) > 0 Then
        lines = Split(payload, vbLf)
        For Each line In lines
            result.Add parser.Canonical(CStr(line))
        Next line
    End If
    Set LoadSequences = result
End Function

Public Sub SaveSequences(ByVal sequences As Collection)
    Dim parser As New MRSequenceParser
    Dim payload As String, i As Long
    For i = 1 To sequences.Count
        If i > 1 Then payload = payload & vbLf
        payload = payload & parser.Canonical(CStr(sequences(i)))
    Next i
    ' One value: do not partially persist Count / Item1 / Item2 on failure.
    SaveSetting REG_APP, REG_SECTION, REG_KEY, payload
End Sub
