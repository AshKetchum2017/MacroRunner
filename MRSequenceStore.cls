Option Explicit

' Version 2 file storage. The old registry value is read only if no file exists.
#If VBA7 Then
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal source As LongPtr, ByVal target As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function MoveFileExW Lib "kernel32" (ByVal source As Long, ByVal target As Long, ByVal flags As Long) As Long
#End If

Public Property Get ConfigurationPath() As String
    Dim root As String
    root = Environ$("APPDATA")
    If Len(root) = 0 Then Err.Raise 76, "MRSequenceStore", "APPDATA tidak tersedia."
    ConfigurationPath = root & "\RinCorelMacros\MacroRunner\Configurations.json"
End Property

Public Function LoadSequences() As Collection
    Dim result As Collection, item As MRConfiguration, parser As New MRSequenceParser
    Dim codec As New MRConfigJson, fso As Object, stream As Object
    Dim path As String, payload As String, line As Variant
    Dim number As Long, description As String
    On Error GoTo Failed
    path = ConfigurationPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2: stream.Charset = "utf-8": stream.Open
        stream.LoadFromFile path
        payload = stream.ReadText
        stream.Close: Set stream = Nothing
        Set result = codec.Decode(payload)
    Else
        Set result = New Collection
        payload = GetSetting("RinCorelMacros", "MacroRunner", "SequencesV1", vbNullString)
        If Len(payload) > 0 Then
            For Each line In Split(payload, vbLf)
                Set item = New MRConfiguration
                item.SequenceText = parser.Canonical(CStr(line))
                result.Add item
            Next line
        End If
    End If
    For Each item In result
        item.SequenceText = parser.Canonical(item.SequenceText)
    Next item
    Set LoadSequences = result
    Exit Function
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise number, "MRSequenceStore.LoadSequences", path & vbCrLf & description
End Function

Public Sub SaveSequences(ByVal sequences As Collection)
    Dim codec As New MRConfigJson, parser As New MRSequenceParser
    Dim item As MRConfiguration, fso As Object, stream As Object
    Dim path As String, folder As String, parent As String, temp As String, payload As String
    Dim number As Long, description As String
    On Error GoTo Failed
    For Each item In sequences
        If parser.Canonical(item.SequenceText) <> item.SequenceText Then _
            Err.Raise 5, "MRSequenceStore", "Urutan macro belum dinormalisasi."
    Next item
    payload = codec.Encode(sequences)
    path = ConfigurationPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(path): parent = fso.GetParentFolderName(folder)
    If Not fso.FolderExists(fso.GetParentFolderName(parent)) Then Err.Raise 76, , "Folder APPDATA tidak ditemukan."
    If Not fso.FolderExists(parent) Then fso.CreateFolder parent
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    temp = fso.BuildPath(folder, fso.GetTempName())
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText payload
    stream.SaveToFile temp, 1
    stream.Close: Set stream = Nothing
    If MoveFileExW(StrPtr(temp), StrPtr(path), 9) = 0 Then _
        Err.Raise 75, "MRSequenceStore", "Gagal mengganti konfigurasi (Win32 " & CStr(Err.LastDllError) & "). File lama dipertahankan."
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Len(temp) > 0 Then
        If fso.FileExists(temp) Then fso.DeleteFile temp
    End If
    On Error GoTo 0
    Err.Raise number, "MRSequenceStore.SaveSequences", description
End Sub
