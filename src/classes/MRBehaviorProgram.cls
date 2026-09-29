Option Explicit

' Class Module (Name) = MRBehaviorProgram. Runner-only outer grammar.
' Target GMS projects continue to receive and parse their original single block.
Private pText As String
Private pAt As Long
Private pActions As Collection
Private pBlockParser As MRBehaviorParser

Public Property Get Actions() As Collection
    Set Actions = pActions
End Property

Public Function Parse(ByVal script As String, ByVal names As Variant) As Collection
    Dim blocks As New Collection, actionsForBlock As Collection
    Dim parsed As Collection, block As MRBehaviorBlock
    Dim start As Long, blockText As String
    Dim name As String, i As Long, action As String, continued As Boolean

    Set pActions = New Collection
    Set pBlockParser = New MRBehaviorParser
    pText = script: pAt = 1
    SkipWhitespace
    If pAt > Len(pText) Then
        Set parsed = pBlockParser.Parse(script, names)
        For i = 1 To parsed.Count
            Set actionsForBlock = New Collection
            pActions.Add actionsForBlock
        Next i
        Set Parse = parsed
        Exit Function
    End If

    Do While pAt <= Len(pText)
        If Mid$(pText, pAt, 1) <> "{" Then Fail "Diharapkan blok macro '{...}'."
        i = LBound(names) + blocks.Count
        If i > UBound(names) Then Fail "Blok macro lebih banyak daripada txbSelectedMacro."
        start = pAt
        ReadBlock
        blockText = Mid$(pText, start, pAt - start)
        name = CStr(names(i))
        Set parsed = pBlockParser.Parse(blockText, Array(name))
        Set block = parsed(1)
        blocks.Add block

        Set actionsForBlock = New Collection
        pActions.Add actionsForBlock
        continued = False
        SkipWhitespace
        Do While Mid$(pText, pAt, 1) = "@"
            If continued Then Fail "@cmdContinue harus menjadi token terakhir sebelum blok berikutnya."
            action = ReadAction()
            actionsForBlock.Add action
            If StrComp(action, "cmdContinue", vbTextCompare) = 0 Then continued = True
            SkipWhitespace
        Loop
        If continued And pAt > Len(pText) Then Fail "@cmdContinue hanya berlaku di antara dua blok macro."
        If pAt <= Len(pText) And Mid$(pText, pAt, 1) <> "{" Then _
            Fail "Diharapkan token '@Nama;' atau blok macro berikutnya."
    Loop
    If blocks.Count <> UBound(names) - LBound(names) + 1 Then _
        Fail "Setiap macro terpilih harus mempunyai satu blok behavior."
    Set Parse = blocks
End Function

Private Sub ReadBlock()
    Dim quoted As Boolean, ch As String
    pAt = pAt + 1
    Do While pAt <= Len(pText)
        ch = Mid$(pText, pAt, 1)
        If ch = """" Then
            If quoted And Mid$(pText, pAt + 1, 1) = """" Then
                pAt = pAt + 2
                GoTo NextCharacter
            End If
            quoted = Not quoted
        ElseIf ch = "}" And Not quoted Then
            pAt = pAt + 1
            Exit Sub
        End If
        pAt = pAt + 1
NextCharacter:
    Loop
    Fail "Blok macro belum ditutup dengan '}'."
End Sub

Private Function ReadAction() As String
    Dim start As Long, ch As String
    pAt = pAt + 1: start = pAt
    Do While pAt <= Len(pText)
        ch = LCase$(Mid$(pText, pAt, 1))
        If Not ((ch >= "a" And ch <= "z") Or ch = "_" Or _
                (pAt > start And ch >= "0" And ch <= "9")) Then Exit Do
        pAt = pAt + 1
    Loop
    If pAt = start Then Fail "Nama token diperlukan setelah '@'."
    ReadAction = Mid$(pText, start, pAt - start)
    SkipWhitespace
    If Mid$(pText, pAt, 1) <> ";" Then Fail "Token luar blok harus diakhiri ';'."
    pAt = pAt + 1
End Function

Private Sub SkipWhitespace()
    Dim ch As String
    Do While pAt <= Len(pText)
        ch = Mid$(pText, pAt, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        pAt = pAt + 1
    Loop
End Sub

Private Sub Fail(ByVal message As String)
    Dim prefix As String, line As Long, column As Long
    prefix = Left$(pText, pAt - 1)
    line = UBound(Split(prefix, vbLf)) + 1
    column = pAt - InStrRev(prefix, vbLf)
    Err.Raise 5, "MRBehaviorProgram", _
        "Behavior baris " & CStr(line) & ", kolom " & CStr(column) & ": " & message
End Sub
