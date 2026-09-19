Option Explicit

' Class Module (Name) = MRPresenter. Coordinates views and models.
Private pView As Object
Private pSequences As Collection
Private pSelected As Long
Private pStore As MRSequenceStore
Private pParser As MRSequenceParser
Private pCatalog As MRCatalog
Private pRun As MRRunModel
Private pClock As MRStopwatch
Private pInvoking As Boolean
Private pClosedDuringCall As Boolean
Private pToken As String
Private pSerial As Long
Private pClosedSeconds As Double
Private pBehavior As MRBehaviorParser
Private pBlocks As Collection
Private pRunIndex As Long

Public Sub Initialize(ByVal view As Object)
    Set pView = view
    Set pStore = New MRSequenceStore
    Set pParser = New MRSequenceParser
    Set pCatalog = New MRCatalog
    Set pRun = New MRRunModel
    Set pClock = New MRStopwatch
    Set pBehavior = New MRBehaviorParser
    Set pSequences = pStore.LoadSequences()
    Render
End Sub

Public Sub Detach()
    Set pView = Nothing
End Sub

Public Property Get CanClose() As Boolean
    If pRun Is Nothing Then
        CanClose = True
    Else
        CanClose = Not pInvoking And pRun.State <> "running"
    End If
End Property

Private Function CanConfigure() As Boolean
    CanConfigure = Not pInvoking And (pRun.State = vbNullString Or pRun.State = "finished")
End Function

Public Sub OpenQueue()
    Dim view As MacroQueue

    If Not CanConfigure Then Exit Sub

    On Error GoTo Failed

    Set view = New MacroQueue

    view.Configure Me
    RefreshQueue view

    ' Blocking call.
    view.Show vbModal

    ' MacroQueue hanya Hide.
    ' Presenter adalah owner lifecycle dan melakukan Unload
    ' setelah modal call sudah kembali.
    Unload view
    Set view = Nothing

    ' Sekarang aman meng-update MacroRunnerMenu.
    Render

    Exit Sub

Failed:
    Dim errorNumber As Long
    Dim errorDescription As String

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    If Not view Is Nothing Then
        Unload view
        Set view = Nothing
    End If

    On Error GoTo 0

    ShowError "Membuka MacroQueue", _
              errorNumber, _
              errorDescription
End Sub

Public Sub OpenSelection(ByVal queueView As Object)
    Dim view As MacroSelection
    Dim operation As String, errorNumber As Long
    Dim errorSource As String, errorDescription As String
    If Not CanConfigure Then Exit Sub
    On Error GoTo Failed
    operation = "MRCatalog.Refresh (daftar file GMS)"
    pCatalog.Refresh
    operation = "New MacroSelection"
    Set view = New MacroSelection
    operation = "MacroSelection.Configure"
    view.Configure Me, pCatalog.Items
    operation = "MacroSelection.Show vbModal"
    view.Show vbModal
    operation = "Unload MacroSelection"
    Unload view
    operation = "Melepas MacroSelection"
    Set view = Nothing
    operation = "RefreshQueue setelah MacroSelection ditutup"
    RefreshQueue queueView
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If Not view Is Nothing Then Unload view
    Set view = Nothing
    On Error GoTo 0
    ShowError "Membuka MacroSelection / " & operation, errorNumber, _
        "Source: " & errorSource & vbCrLf & errorDescription
End Sub

Public Function AppendMacro(ByVal draft As String, ByVal name As String) As String
    Dim item As MRMacroDefinition
    Set item = pCatalog.Resolve(Trim$(name))
    If Len(Trim$(draft)) > 0 Then draft = pParser.Canonical(draft) & " "
    AppendMacro = draft & item.DisplayName & ";"
End Function

Public Function EditTokens(ByVal previous As String, ByVal current As String, ByRef caret As Long) As String
    EditTokens = pParser.RewriteDeletedDelimiter(previous, current, caret)
End Function

Public Function SaveSelection(ByVal draft As String, Optional ByVal behavior As String = "") As Boolean
    Dim names As Variant, name As Variant, result As String
    Dim item As MRMacroDefinition, updated As Collection
    Dim config As MRConfiguration, blocks As Collection
    On Error GoTo Failed
    pCatalog.Refresh
    names = pParser.Parse(draft)
    For Each name In names
        Set item = pCatalog.Resolve(CStr(name))
        If Len(result) > 0 Then result = result & " "
        result = result & item.DisplayName & ";"
    Next name
    Set blocks = pBehavior.Parse(behavior, names)
    Set config = New MRConfiguration
    config.SequenceText = result
    config.BehaviorText = behavior
    Set updated = CopySequences(0)
    updated.Add config
    pStore.SaveSequences updated
    Set pSequences = updated
    SaveSelection = True
    Exit Function
Failed:
    ShowError "Menyimpan daftar macro", Err.Number, Err.Description
End Function

Public Sub RemoveSequence(ByVal view As Object, ByVal index As Long)
    Dim updated As Collection
    If index < 1 Or index > pSequences.Count Then Exit Sub
    On Error GoTo Failed
    Set updated = CopySequences(index)
    pStore.SaveSequences updated
    Set pSequences = updated
    If pSelected = index Then
        pSelected = 0
    ElseIf pSelected > index Then
        pSelected = pSelected - 1
    End If
    RefreshQueue view
    Render
    Exit Sub
Failed:
    ShowError "Menghapus daftar macro", Err.Number, Err.Description
End Sub

Public Sub ClearSequences(ByVal view As Object)
    Dim updated As New Collection
    On Error GoTo Failed
    pStore.SaveSequences updated
    Set pSequences = updated
    pSelected = 0
    RefreshQueue view
    Render
    Exit Sub
Failed:
    ShowError "Mengosongkan daftar macro", Err.Number, Err.Description
End Sub

Public Function SelectSequence(ByVal index As Long) As Boolean
    If index < 1 Or index > pSequences.Count Then
        MsgBox "Pilih satu daftar macro terlebih dahulu.", vbExclamation, "Macro Queue"
        Exit Function
    End If
    pSelected = index
    SelectSequence = True
End Function

Private Function CopySequences(ByVal exclude As Long) As Collection
    Dim result As New Collection, i As Long
    For i = 1 To pSequences.Count
        If i <> exclude Then result.Add pSequences(i)
    Next i
    Set CopySequences = result
End Function

Private Sub RefreshQueue(ByVal view As Object)
    Dim labels As New Collection, i As Long
    For i = 1 To pSequences.Count
        labels.Add pParser.QueueLabel(pSequences(i).SequenceText)
    Next i
    view.DisplaySequences labels, pSelected
End Sub

Public Sub Process()
    Dim names As Variant, name As Variant, item As MRMacroDefinition
    If Not CanConfigure Then Exit Sub
    If pSelected < 1 Or pSelected > pSequences.Count Then
        MsgBox "Pilih daftar melalui MacroQueue terlebih dahulu.", vbExclamation, "Macro Runner"
        Exit Sub
    End If
    On Error GoTo Failed
    names = pParser.Parse(pSequences(pSelected).SequenceText)
    pCatalog.Refresh
    For Each name In names
        Set item = pCatalog.Resolve(CStr(name))
    Next name
    Set pBlocks = pBehavior.Parse(pSequences(pSelected).BehaviorText, names)
    pRunIndex = 0
    pRun.Prepare names
    RunNext
    Exit Sub
Failed:
    ShowError "Menyiapkan antrean", Err.Number, Err.Description
End Sub

Public Sub ContinueRun()
    If pInvoking Or pRun.State <> "waiting" Then Exit Sub
    RunNext
End Sub

Private Sub RunNext()
    Dim item As MRMacroDefinition, launched As Variant
    Dim block As MRBehaviorBlock
    Dim observerArgument As Variant
    Dim errorNumber As Long, errorDescription As String, errorSource As String
    Dim operation As String, errorDetails As String, macroName As String
    On Error GoTo Failed
    operation = "MRRunModel.BeginNext"
    pRun.BeginNext
    pRunIndex = pRunIndex + 1
    Set block = pBlocks(pRunIndex)
    macroName = pRun.CurrentName
    pSerial = pSerial + 1
    pToken = CStr(pSerial)
    pInvoking = True
    pClosedDuringCall = False
    operation = "MRStopwatch.Start sebelum pemeriksaan katalog"
    pClock.Start
    operation = "Menampilkan status RUNNING"
    Render
    operation = "MRCatalog.Refresh"
    pCatalog.Refresh
    operation = "MRCatalog.ResolveForRun: " & macroName
    Set item = pCatalog.ResolveForRun(pRun.CurrentName)
    operation = "Menyiapkan referensi callback MRPresenter"
    ' Preserve an object-valued Variant at the COM ParamArray boundary.
    Set observerArgument = Me
    ' Successful timing starts at the actual launch, excluding catalog lookup.
    operation = "MRStopwatch.Start sebelum membuka form"
    pClock.Start
    operation = "GMSManager.RunMacro / MRTargetBridge.OpenMacro"
    If block.HasCommands Then
        operation = "MRTargetBridge.RunBehavior: " & block.MacroID
        launched = Application.GMSManager.RunMacro(item.ProjectName, "MRTargetBridge.RunBehavior", _
            block.SourceText, observerArgument, pToken)
    Else
        launched = Application.GMSManager.RunMacro(item.ProjectName, "MRTargetBridge.OpenMacro", _
            item.FormName, item.Modal, observerArgument, pToken)
    End If
    Set observerArgument = Nothing
    operation = "Memvalidasi hasil MRTargetBridge.OpenMacro"
    If VarType(launched) <> vbBoolean Then Err.Raise 5, "MacroRunner", "MRTargetBridge tidak mengembalikan status pembukaan yang valid."
    If Not CBool(launched) Then Err.Raise 5, "MacroRunner", "Macro tujuan gagal dibuka."
    pInvoking = False
    operation = "Mencatat penutupan form modal"
    If pClosedDuringCall Then pRun.CompleteCurrent pClosedSeconds, 0, vbNullString
    operation = "Menampilkan status setelah pemanggilan bridge"
    Render
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    errorDetails = "Tahap: " & operation & vbCrLf & "Source: " & errorSource & vbCrLf & errorDescription
    If Not item Is Nothing Then
        errorDetails = errorDetails & vbCrLf & "Project: " & item.ProjectName & _
            vbCrLf & "Form: " & item.FormName
    End If
    pInvoking = False
    ' A secondary display/cleanup error must not replace the original failure.
    On Error Resume Next
    Set observerArgument = Nothing
    pRun.CompleteCurrent pClock.StopSeconds(), errorNumber, errorDetails
    Render
    On Error GoTo 0
    ShowError "Menjalankan " & macroName, errorNumber, errorDetails
End Sub

' A modal launch may still be on the stack. Never run the next macro here.
Public Sub MacroUnloaded(ByVal token As String)
    If token <> pToken Or pRun.State <> "running" Then Exit Sub
    pClosedSeconds = pClock.StopSeconds()
    If pInvoking Then
        pClosedDuringCall = True
    Else
        pRun.CompleteCurrent pClosedSeconds, 0, vbNullString
        Render
    End If
End Sub

Public Sub CopyStatistics()
    Dim clipboard As New MRClipboard
    Dim value As String
    On Error GoTo Failed
    value = pRun.StatisticsText()
    If Len(value) = 0 Then Exit Sub
    clipboard.CopyText value
    Exit Sub
Failed:
    ShowError "Menyalin statistik", Err.Number, Err.Description
End Sub

' Runtime failures after the script returned (for example a manual draft Save).
Public Sub BehaviorFailed(ByVal token As String, ByVal number As Long, ByVal description As String)
    If token <> pToken Or pRun.State <> "running" Then Exit Sub
    pRun.CompleteCurrent pClock.StopSeconds(), number, description
    Render
    ShowError "MacroBehavior " & pRun.CurrentName, number, description
End Sub

Private Sub Render()
    Dim caption As String
    If pView Is Nothing Then Exit Sub
    Select Case pRun.State
        Case "running": caption = "Berjalan: " & pRun.CurrentName
        Case "waiting": caption = "Macro ditutup. Tekan Continue untuk melanjutkan."
        Case "finished"
            caption = "Selesai"
            If Len(pRun.LastError) > 0 Then caption = "Dihentikan: " & pRun.LastError
        Case Else: caption = "Pilih daftar macro melalui Queue Settings."
    End Select
    If CanConfigure And pSelected > 0 Then caption = caption & " | " & pParser.QueueLabel(pSequences(pSelected).SequenceText)
    pView.DisplayState pRun.StatisticsText(), caption, CanConfigure And pSelected > 0, _
        Not pInvoking And pRun.State = "waiting", CanConfigure, CanClose
End Sub

Public Sub ShowError(ByVal operation As String, ByVal number As Long, ByVal description As String)
    MsgBox "Operasi: " & operation & vbCrLf & "Error " & CStr(number) & vbCrLf & _
        "Description: " & description, vbExclamation, "Macro Runner"
End Sub
