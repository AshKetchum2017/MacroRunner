Option Explicit

' Import this class plus MRBehaviorBlock/MRBehaviorInstruction into both
' MacroRunner and ExportRelated. Parsing/validation has no UI or registry effects.
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
            If Not ValidForm(block.MacroID, form) Then Fail "UserForm tidak terdaftar: " & block.MacroID & ": " & form
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
                ValidateInstruction block.MacroID, ins
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
        If LCase$(block.MacroID) = "exportrelated" Then ValidateFlow block
        blocks.Add block
        White
    Loop
    If blocks.Count <> UBound(names) - LBound(names) + 1 Then Fail "Setiap macro terpilih harus mempunyai satu blok behavior."
    Set Parse = blocks
End Function

Private Function ValidForm(ByVal macro As String, ByVal form As String) As Boolean
    Dim expected As String
    Select Case LCase$(macro)
        Case "exportrelated"
            ValidForm = (LCase$(form) = "exportrelatedmenu" Or LCase$(form) = "exportrelatedsettings")
            Exit Function
        Case "autodistributeuf": expected = "autodistributemenu"
        Case "autodistribute": expected = "masterwizard"
        Case "autolabel": expected = "autolabelwizard"
        Case "autolabelbp": expected = "autolabelwizardbp"
        Case "autosavencreate": expected = "autosnc"
        Case "pagerelated": expected = "pagerelatedmenu"
        Case "pericount": expected = "pericountmenu"
        Case "startpointorganizer": expected = "startpointmenu"
        Case "simpletools": expected = "polakartunamalengkungwizard"
        Case "dashlinebreaker": expected = "dashlinebreakermenu"
        Case "objectcontainer": expected = "objectcontainermenu"
    End Select
    ValidForm = (Len(expected) > 0 And LCase$(form) = expected)
End Function

Private Sub ValidateInstruction(ByVal macro As String, ByVal ins As MRBehaviorInstruction)
    Dim expected As String, key As String
    If LCase$(macro) <> "exportrelated" Then Fail "Behavior aktif baru didukung untuk ExportRelated; gunakan [] untuk macro lain."
    key = LCase$(ins.Target)
    If ins.Kind = "action" Then
        If LCase$(ins.FormName) = "exportrelatedmenu" Then
            If key = "cmdaddsetting" Or key = "cmdexport" Or key = "cmdclose" Then Exit Sub
        Else
            If key = "cmdsave" Or key = "cmdcancel" Then Exit Sub
        End If
        Fail "Action tidak terdaftar: " & ins.FormName & ".@" & ins.Target
    End If
    If LCase$(ins.FormName) = "exportrelatedsettings" Then
        Select Case key
            Case "txbname", "txbpage", "txbdirectory", "cmbexformat": expected = "string"
            Case "chkparentdirectory": expected = "boolean"
        End Select
    Else
        Select Case key
            Case "lbxsettinglists.index": expected = "number"
            Case "chklayer1", "chklayer2", "chklayer3": expected = "boolean"
        End Select
    End If
    If Len(expected) = 0 Then Fail "Target tidak terdaftar: " & ins.FormName & "." & ins.Target
    If ins.Kind = "skip" Then Exit Sub
    If ins.Kind = "default" Then
        If Len(DefaultKey(ins.Target)) = 0 Or LCase$(ins.FormName) <> "exportrelatedsettings" Then _
            Fail "Target tidak mempunyai Default registry: " & ins.Target
        Exit Sub
    End If
    If ins.Kind <> expected Then Fail "Tipe nilai " & ins.Target & " harus " & expected & "."
    If key = "lbxsettinglists.index" Then
        If ins.Value < 1 Or ins.Value > 2147483647# Then Fail "Index harus bilangan bulat positif mulai dari 1."
        If Fix(ins.Value) <> ins.Value Then Fail "Index tidak boleh pecahan."
    End If
    If key = "cmbexformat" Then
        Select Case LCase$(CStr(ins.Value))
            Case ".pdf", ".dxf", ".png", ".jpg"
            Case Else: Fail "Format harus .pdf, .dxf, .png, atau .jpg."
        End Select
    End If
End Sub

Public Function DefaultKey(ByVal target As String) As String
    Select Case LCase$(target)
        Case "txbdirectory": DefaultKey = "LastDirectory"
        Case "cmbexformat": DefaultKey = "LastExportFormat"
        Case "chkparentdirectory": DefaultKey = "UseParentDirectory"
    End Select
End Function

Private Sub ValidateFlow(ByVal block As MRBehaviorBlock)
    Dim ins As MRBehaviorInstruction, editorOpen As Boolean, closed As Boolean
    Dim resumeAt As Long
    resumeAt = pAt
    For Each ins In block.Instructions
        pAt = ins.Position
        If closed Then Fail "Tidak boleh ada tahap/instruksi setelah @cmdClose."
        If LCase$(ins.FormName) = "exportrelatedsettings" Then
            If Not editorOpen Then Fail "ExportRelatedSettings belum dibuka; jalankan @cmdAddSetting dahulu."
        Else
            If editorOpen Then Fail "Selesaikan editor melalui @cmdSave atau @cmdCancel sebelum kembali ke menu."
        End If
        If ins.Kind = "action" Then
            Select Case LCase$(ins.Target)
                Case "cmdaddsetting": editorOpen = True
                Case "cmdsave", "cmdcancel": editorOpen = False
                Case "cmdclose": closed = True
            End Select
        End If
    Next ins
    pAt = resumeAt
End Sub

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
