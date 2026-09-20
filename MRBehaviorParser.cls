Option Explicit

' Import this class plus MRBehaviorBlock/MRBehaviorInstruction into both
' MacroRunner and target GMS projects. Grammar only: no target semantics or side effects.
Private pText As String
Private pAt As Long

Public Function Parse(ByVal script As String, ByVal names As Variant) As Collection
    Dim blocks As New Collection, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction, form As String, start As Long, i As Long
    pText = script
    pAt = 1
    White
    If pAt > Len(pText) Then
        For i = LBound(names) To UBound(names)
            Set block = New MRBehaviorBlock
            block.MacroID = CStr(names(i))
            blocks.Add block
        Next i
        Set Parse = blocks
        Exit Function
    End If
    Do While pAt <= Len(pText)
        start = pAt
        Expect "{"
        Set block = New MRBehaviorBlock
        block.MacroID = MacroName()
        Expect ":"
        i = LBound(names) + blocks.Count
        If i > UBound(names) Then Fail "Blok macro lebih banyak daripada txbSelectedMacro."
        If StrComp(block.MacroID, CStr(names(i)), vbTextCompare) <> 0 Then _
            Fail "Urutan MacroID tidak cocok. Diharapkan: " & CStr(names(i))
        Do
            form = Identifier()
            Set ins = New MRBehaviorInstruction
            ins.FormName = form: ins.Kind = "stage": ins.Position = pAt
            block.Instructions.Add ins
            Expect "["
            Do While Peek() <> "]"
                If pAt > Len(pText) Then Fail "Behavior belum ditutup dengan ]."
                Set ins = New MRBehaviorInstruction
                ins.FormName = form: ins.Position = pAt
                If Peek() = "@" Then
                    Expect "@"
                    ins.Target = Identifier(): ins.Kind = "action"
                Else
                    ins.Target = Identifier()
                    If Peek() = "." Then
                        Expect "."
                        ins.Target = ins.Target & "." & Identifier()
                    End If
                    Expect "="
                    ReadValue ins
                End If
                block.Instructions.Add ins
                block.HasCommands = True
                If Peek() <> ";" Then Exit Do
                Expect ";"
            Loop
            Expect "]"
            If Peek() <> ";" Then Exit Do
            Expect ";"
            If Peek() = "}" Then Exit Do
        Loop
        Expect "}"
        block.SourceText = Mid$(pText, start, pAt - start)
        blocks.Add block
        White
    Loop
    If blocks.Count <> UBound(names) - LBound(names) + 1 Then Fail "Setiap macro terpilih harus mempunyai satu blok behavior."
    Set Parse = blocks
End Function

Private Sub ReadValue(ByVal ins As MRBehaviorInstruction)
    Dim raw As String, start As Long, ch As String, digits As Long, dots As Long
    If Peek() = """" Then
        ins.Kind = "string": ins.Value = Quoted(): Exit Sub
    End If
    start = pAt
    Do While pAt <= Len(pText)
        ch = Mid$(pText, pAt, 1)
        If ch = ";" Or ch = "]" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit Do
        pAt = pAt + 1
    Loop
    raw = Mid$(pText, start, pAt - start)
    Select Case LCase$(raw)
        Case "default": ins.Kind = "default"
        Case "nothing", "empty", "null": ins.Kind = "skip"
        Case "true", "false": ins.Kind = "boolean": ins.Value = (LCase$(raw) = "true")
        Case Else
            For start = 1 To Len(raw)
                ch = Mid$(raw, start, 1)
                If ch >= "0" And ch <= "9" Then
                    digits = digits + 1
                ElseIf ch = "." Then
                    dots = dots + 1
                ElseIf Not (start = 1 And (ch = "-" Or ch = "+")) Then
                    Fail "Nilai tidak dikenali; teks harus di dalam quote."
                End If
            Next start
            If digits = 0 Or dots > 1 Then Fail "Angka tidak valid."
            ins.Kind = "number": ins.Value = Val(raw)
    End Select
End Sub

Private Function Quoted() As String
    Dim result As String, ch As String
    Expect """"
    Do While pAt <= Len(pText)
        ch = Mid$(pText, pAt, 1): pAt = pAt + 1
        If ch = """" Then
            If Mid$(pText, pAt, 1) = """" Then
                pAt = pAt + 1
            Else
                Quoted = result: Exit Function
            End If
        End If
        result = result & ch
    Loop
    Fail "String belum ditutup dengan quote."
End Function

Private Function MacroName() As String
    Dim start As Long
    White: start = pAt
    Do While pAt <= Len(pText) And Mid$(pText, pAt, 1) <> ":"
        If InStr("{}[];""", Mid$(pText, pAt, 1)) > 0 Then Fail "MacroID tidak valid."
        pAt = pAt + 1
    Loop
    MacroName = Trim$(Replace(Replace(Replace(Mid$(pText, start, pAt - start), vbTab, " "), vbCr, " "), vbLf, " "))
    If Len(MacroName) = 0 Then Fail "MacroID kosong."
End Function

Private Function Identifier() As String
    Dim start As Long, ch As String
    White: start = pAt
    Do While pAt <= Len(pText)
        ch = LCase$(Mid$(pText, pAt, 1))
        If Not ((ch >= "a" And ch <= "z") Or ch = "_" Or (pAt > start And ch >= "0" And ch <= "9")) Then Exit Do
        pAt = pAt + 1
    Loop
    If pAt = start Then Fail "Nama form/target/action diperlukan."
    Identifier = Mid$(pText, start, pAt - start)
End Function

Private Sub White()
    Do While pAt <= Len(pText)
        If InStr(" " & vbTab & vbCr & vbLf, Mid$(pText, pAt, 1)) = 0 Then Exit Do
        pAt = pAt + 1
    Loop
End Sub

Private Function Peek() As String
    White
    Peek = Mid$(pText, pAt, 1)
End Function

Private Sub Expect(ByVal token As String)
    If Peek() <> token Then Fail "Diharapkan '" & token & "'."
    pAt = pAt + Len(token)
End Sub

Private Sub Fail(ByVal message As String)
    Dim prefix As String, line As Long, column As Long, lastBreak As Long
    prefix = Left$(pText, pAt - 1)
    line = UBound(Split(prefix, vbLf)) + 1
    lastBreak = InStrRev(prefix, vbLf)
    column = pAt - lastBreak
    Err.Raise 5, "MRBehaviorParser", "Behavior baris " & CStr(line) & ", kolom " & CStr(column) & ": " & message
End Sub
