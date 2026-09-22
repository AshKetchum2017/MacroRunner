Option Explicit

' Class Module (Name) = MRConfigJson. Strict versioned JSON; never evaluates code.
Private pText As String
Private pAt As Long

Public Function Encode(ByVal entries As Collection) As String
    Dim item As MRConfiguration, result As String
    result = "{""version"":2,""configurations"":["
    For Each item In entries
        If Right$(result, 1) <> "[" Then result = result & ","
        result = result & vbCrLf & _
            "  {""configId"":" & Quote(item.ConfigID) & _
            ",""{""sequence"":" & Quote(item.SequenceText) & _
            ",""behavior"":" & Quote(item.BehaviorText) & "}"
    Next item
    Encode = result & vbCrLf & "]}"
End Function

Public Function Decode(ByVal value As String) As Collection
    Dim entries As Collection, key As String, versionSeen As Boolean, entriesSeen As Boolean
    pText = value: pAt = 1
    Expect "{"
    Do
        key = ReadString(): Expect ":"
        Select Case key
            Case "version"
                If versionSeen Then Bad "version duplikat."
                versionSeen = True: Expect "2"
            Case "configurations"
                If entriesSeen Then Bad "configurations duplikat."
                entriesSeen = True: Set entries = ReadEntries()
            Case Else: Bad "Properti tidak dikenal: " & key
        End Select
        If Peek() <> "," Then Exit Do
        Expect ","
    Loop
    Expect "}"
    If Not versionSeen Or Not entriesSeen Then Bad "version/configurations tidak lengkap."
    If Len(Peek()) > 0 Then Bad "Teks setelah akhir JSON."
    Set Decode = entries
End Function

Private Function ReadEntries() As Collection
    Dim entries As New Collection, item As MRConfiguration, key As String
    Dim sequenceSeen As Boolean, behaviorSeen As Boolean, configIDSeen As Boolean
    Expect "["
    If Peek() <> "]" Then
        Do
            Set item = New MRConfiguration
            sequenceSeen = False: behaviorSeen = False: configIDSeen = False
            Expect "{"
            Do
                key = ReadString(): Expect ":"
                Select Case key
                    Case "configId"
                        If configIDSeen Then Bad "configId duplikat."
                        configIDSeen = True: item.ConfigID = ReadString()
                    Case "sequence"
                        If sequenceSeen Then Bad "sequence duplikat."
                        sequenceSeen = True: item.SequenceText = ReadString()
                    Case "behavior"
                        If behaviorSeen Then Bad "behavior duplikat."
                        behaviorSeen = True: item.BehaviorText = ReadString()
                    Case Else: Bad "Properti konfigurasi tidak dikenal: " & key
                End Select
                If Peek() <> "," Then Exit Do
                Expect ","
            Loop
            Expect "}"
            If Not sequenceSeen Or Not behaviorSeen Then Bad "sequence/behavior tidak lengkap."
            entries.Add item
            If Peek() <> "," Then Exit Do
            Expect ","
        Loop
    End If
    Expect "]"
    Set ReadEntries = entries
End Function

Private Function Quote(ByVal value As String) As String
    Dim result As String, i As Long, ch As String, code As Long
    result = """"
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1): code = AscW(ch)
        Select Case ch
            Case """", "\": result = result & "\" & ch
            Case vbCr: result = result & "\r"
            Case vbLf: result = result & "\n"
            Case vbTab: result = result & "\t"
            Case Else
                If code >= 0 And code < 32 Then
                    result = result & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next i
    Quote = result & """"
End Function

Private Function ReadString() As String
    Dim result As String, ch As String, hexText As String, i As Long
    Expect """"
    Do While pAt <= Len(pText)
        ch = Mid$(pText, pAt, 1): pAt = pAt + 1
        If ch = """" Then ReadString = result: Exit Function
        If ch = "\" Then
            ch = Mid$(pText, pAt, 1): pAt = pAt + 1
            Select Case ch
                Case """", "\", "/"
                Case "n": ch = vbLf
                Case "r": ch = vbCr
                Case "t": ch = vbTab
                Case "b": ch = ChrW(8)
                Case "f": ch = ChrW(12)
                Case "u"
                    hexText = Mid$(pText, pAt, 4)
                    If Len(hexText) <> 4 Then Bad "Escape Unicode tidak lengkap."
                    For i = 1 To 4
                        If InStr("0123456789abcdef", LCase$(Mid$(hexText, i, 1))) = 0 Then Bad "Escape Unicode tidak valid."
                    Next i
                    ch = ChrW(CLng("&H" & hexText))
                    pAt = pAt + 4
                Case Else: Bad "Escape string tidak valid."
            End Select
        ElseIf AscW(ch) >= 0 And AscW(ch) < 32 Then
            Bad "Karakter kontrol harus di-escape."
        End If
        result = result & ch
    Loop
    Bad "String belum ditutup."
End Function

Private Function Peek() As String
    Do While pAt <= Len(pText)
        If InStr(" " & vbTab & vbCr & vbLf, Mid$(pText, pAt, 1)) = 0 Then Exit Do
        pAt = pAt + 1
    Loop
    Peek = Mid$(pText, pAt, 1)
End Function

Private Sub Expect(ByVal token As String)
    If Peek() <> token Then Bad "Diharapkan '" & token & "'."
    pAt = pAt + 1
End Sub

Private Sub Bad(ByVal message As String)
    Err.Raise 5, "MRConfigJson", "JSON posisi " & CStr(pAt) & ": " & message
End Sub
