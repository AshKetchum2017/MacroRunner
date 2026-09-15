Option Explicit

' Class Module (Name) = MRTokenParser.
Public Function Parse(ByVal value As String) As Variant
    Dim parts As Variant
    Dim result() As String
    Dim i As Long, last As Long
    value = Trim$(value)
    If Len(value) = 0 Then Err.Raise 5, "MacroSelection", "Pilih atau ketik minimal satu nama macro."
    If Right$(value, 1) = ";" Then value = Left$(value, Len(value) - 1)
    If Len(Trim$(value)) = 0 Then Err.Raise 5, "MacroSelection", "Pilih atau ketik minimal satu nama macro."
    parts = Split(value, ";")
    last = UBound(parts)
    ReDim result(0 To last)
    For i = 0 To last
        result(i) = Trim$(CStr(parts(i)))
        If Len(result(i)) = 0 Then Err.Raise 5, "MacroSelection", "Nama macro kosong di antara pemisah titik koma."
        If InStr(result(i), vbCr) > 0 Or InStr(result(i), vbLf) > 0 Then _
            Err.Raise 5, "MacroSelection", "Nama macro tidak boleh berisi baris baru."
    Next i
    Parse = result
End Function

Public Function Canonical(ByVal value As String) As String
    Canonical = Join(Parse(value), "; ") & ";"
End Function

Public Function QueueLabel(ByVal value As String) As String
    QueueLabel = Join(Parse(value), " > ")
End Function

' A single deleted semicolon deletes its preceding token as well.
' Other text edits (typing, paste, selection deletion) remain ordinary edits.
' caret is zero based, as required by MSForms.TextBox.SelStart.
Public Function RewriteDeletedDelimiter(ByVal previous As String, ByVal current As String, _
                                        ByRef caret As Long) As String
    Dim position As Long, beforeToken As Long
    RewriteDeletedDelimiter = current
    If Len(previous) <> Len(current) + 1 Then Exit Function
    position = 1
    Do While position <= Len(current)
        If Mid$(previous, position, 1) <> Mid$(current, position, 1) Then Exit Do
        position = position + 1
    Loop
    If Mid$(previous, position, 1) <> ";" Then Exit Function
    If Mid$(previous, position + 1) <> Mid$(current, position) Then Exit Function
    beforeToken = InStrRev(Left$(previous, position - 1), ";")
    current = Left$(previous, beforeToken) & Mid$(previous, position + 1)
    If beforeToken = 0 Then current = LTrim$(current)
    RewriteDeletedDelimiter = RTrim$(current)
    caret = beforeToken
End Function
