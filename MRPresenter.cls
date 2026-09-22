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
Private pBehaviorBridgeStage As String
Private pBehaviorBridgeError As Long
Private pBehaviorBridgeDescription As String
Private pBehaviorBridgeProcedure As String
Private pPreflighting As Boolean

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
    OpenSelectionEditor queueView, 0
End Sub

Public Sub ModifySequence(ByVal queueView As Object, ByVal index As Long)
    If index < 1 Or index > pSequences.Count Then
        MsgBox "Pilih satu daftar macro terlebih dahulu.", vbExclamation, "Macro Queue"
        Exit Sub
    End If

    OpenSelectionEditor queueView, index
End Sub

Private Sub OpenSelectionEditor(ByVal queueView As Object, ByVal editIndex As Long)
    Dim view As MacroSelection
    Dim availableItems As Collection
    Dim operation As String, errorNumber As Long
    Dim errorSource As String, errorDescription As String
    Dim sequenceText As String, behaviorText As String

    If Not CanConfigure Then Exit Sub
    On Error GoTo Failed

    If editIndex > 0 Then
        sequenceText = pSequences(editIndex).SequenceText
        behaviorText = pSequences(editIndex).BehaviorText
    End If

    operation = "MRCatalog.Refresh (daftar file GMS)"
    On Error Resume Next
    pCatalog.Refresh
    errorNumber = Err.Number: errorDescription = Err.Description
    On Error GoTo Failed
    If errorNumber = 0 Then
        Set availableItems = pCatalog.Items
    Else
        Set availableItems = New Collection
        MsgBox "Daftar GMS tidak tersedia; nama macro tetap dapat diketik manual." & vbCrLf & _
            errorDescription, vbInformation, "Macro Selection"
    End If

    operation = "New MacroSelection"
    Set view = New MacroSelection
    operation = "MacroSelection.Configure"
    view.Configure Me, availableItems, sequenceText, behaviorText, editIndex
    operation = "MacroSelection.Show vbModal"
    view.Show vbModal
    operation = "Unload MacroSelection"
    Unload view
    operation = "Melepas MacroSelection"
    Set view = Nothing
    operation = "RefreshQueue setelah MacroSelection ditutup"
    RefreshQueue queueView, editIndex
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

Public Function SaveSelection(ByVal draft As String, _
                              Optional ByVal behavior As String = "", _
                              Optional ByVal editIndex As Long = 0) As Boolean
    Dim names As Variant, result As String
    Dim updated As Collection
    Dim config As MRConfiguration, blocks As Collection
    On Error GoTo Failed

    If editIndex < 0 Or editIndex > pSequences.Count Then _
        Err.Raise 5, "MRPresenter.SaveSelection", "Index daftar macro tidak valid."

    names = pParser.Parse(draft)
    ' The selected sequence defines local MacroID references, even offline.
    result = pParser.Canonical(draft)
    Set blocks = pBehavior.Parse(behavior, names)

    Set config = New MRConfiguration
    config.SequenceText = result
    config.BehaviorText = behavior

    If editIndex = 0 Then
        Set updated = CopySequences(0)
        updated.Add config
    Else
        Set updated = ReplaceSequence(editIndex, config)
    End If

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

Private Function ReplaceSequence(ByVal index As Long, ByVal replacement As MRConfiguration) As Collection
    Dim result As New Collection, i As Long

    For i = 1 To pSequences.Count
        If i = index Then
            result.Add replacement
        Else
            result.Add pSequences(i)
        End If
    Next i

    Set ReplaceSequence = result
End Function

Private Sub RefreshQueue(ByVal view As Object, Optional ByVal highlighted As Long = 0)
    Dim labels As New Collection, i As Long

    For i = 1 To pSequences.Count
        labels.Add pParser.QueueLabel(pSequences(i).SequenceText)
    Next i

    If highlighted <= 0 Then highlighted = pSelected
    view.DisplaySequences labels, highlighted
End Sub

Public Sub Process()
    Dim names As Variant, number As Long, description As String
    If Not CanConfigure Then Exit Sub
    If pSelected < 1 Or pSelected > pSequences.Count Then
        MsgBox "Pilih daftar melalui MacroQueue terlebih dahulu.", vbExclamation, "Macro Runner"
        Exit Sub
    End If
    On Error GoTo Failed
    pBehaviorBridgeError = 0
    pBehaviorBridgeProcedure = "Grammar konfigurasi"
    names = pParser.Parse(pSequences(pSelected).SequenceText)
    Set pBlocks = pBehavior.Parse(pSequences(pSelected).BehaviorText, names)
    PreflightAll
    pRunIndex = 0
    pRun.Prepare names
    RunNext
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    pInvoking = False: pPreflighting = False
    On Error Resume Next
    Render
    On Error GoTo 0
    ShowError "Preflight antrean / " & pBehaviorBridgeProcedure, number, description
End Sub

Private Sub PreflightAll()
    Dim block As MRBehaviorBlock, item As MRMacroDefinition, result As Variant
    Dim observerArgument As Variant, index As Long
    Dim number As Long, description As String
    On Error GoTo Failed
    pInvoking = True: pPreflighting = True
    pBehaviorBridgeProcedure = "Discovery GMS"
    Render
    pCatalog.Refresh
    For Each block In pBlocks
        index = index + 1
        pBehaviorBridgeProcedure = "Resolusi langkah " & CStr(index) & ": " & block.MacroID
        Set item = pCatalog.ResolveForRun(block.MacroID, Len(block.SourceText) > 0)
        If Len(block.SourceText) > 0 Then
            pSerial = pSerial + 1
            pToken = "preflight:" & CStr(pSerial)
            pBehaviorBridgeStage = vbNullString
            pBehaviorBridgeError = 0
            pBehaviorBridgeDescription = vbNullString
            pBehaviorBridgeProcedure = "MRTargetBridge.ValidateBehavior"
            Set observerArgument = Me
            result = Application.GMSManager.RunMacro(item.ProjectName, pBehaviorBridgeProcedure, _
                block.SourceText, observerArgument, pToken)
            Set observerArgument = Nothing
            ValidateBehaviorBridge result
        End If
    Next block
    pPreflighting = False: pInvoking = False
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    If pBehaviorBridgeError <> 0 Then
        number = pBehaviorBridgeError: description = pBehaviorBridgeDescription
    End If
    Set observerArgument = Nothing
    If Not block Is Nothing Then description = "Langkah " & CStr(index) & " / " & block.MacroID & vbCrLf & description
    Err.Raise number, "MRPresenter.PreflightAll", description
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
    Dim operation As String, errorDetails As String, macroName As String, procedureName As String
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
    pBehaviorBridgeStage = vbNullString
    pBehaviorBridgeError = 0
    pBehaviorBridgeDescription = vbNullString
    operation = "MRStopwatch.Start sebelum pemeriksaan katalog"
    pClock.Start
    operation = "Menampilkan status RUNNING"
    Render
    operation = "MRCatalog.Refresh"
    pCatalog.Refresh
    operation = "MRCatalog.ResolveForRun: " & macroName
    Set item = pCatalog.ResolveForRun(pRun.CurrentName, Len(block.SourceText) > 0)
    operation = "Menyiapkan referensi callback MRPresenter"
    ' Preserve an object-valued Variant at the COM ParamArray boundary.
    Set observerArgument = Me
    ' Successful timing starts at the actual launch, excluding catalog lookup.
    operation = "MRStopwatch.Start sebelum membuka form"
    pClock.Start
    If Len(block.SourceText) > 0 Then
        procedureName = "MRTargetBridge.RunBehavior"
        pBehaviorBridgeProcedure = procedureName
        operation = "GMSManager.RunMacro / " & procedureName
        launched = Application.GMSManager.RunMacro(item.ProjectName, procedureName, _
            block.SourceText, observerArgument, pToken)
    Else
        procedureName = "MRTargetBridge.OpenMacro"
        operation = "GMSManager.RunMacro / " & procedureName
        launched = Application.GMSManager.RunMacro(item.ProjectName, procedureName, _
            item.FormName, item.Modal, observerArgument, pToken)
    End If
    Set observerArgument = Nothing
    operation = "Memvalidasi hasil " & procedureName
    If Len(block.SourceText) > 0 Then
        ValidateBehaviorBridge launched
    Else
        If VarType(launched) <> vbBoolean Then Err.Raise 5, "MacroRunner", _
            procedureName & " tidak mengembalikan Boolean. " & BridgeResultType(launched)
        If Not CBool(launched) Then Err.Raise 5, "MacroRunner", "Macro tujuan gagal dibuka."
    End If
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
    If pBehaviorBridgeError <> 0 Then
        errorNumber = pBehaviorBridgeError
        errorSource = "MRTargetBridge.RunBehavior"
        errorDescription = pBehaviorBridgeDescription
    End If
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

' These callbacks report dispatch/initialization, not completion of the macro step.
' Keep the error until RunMacro returns: hosts may not expose the function result.
Public Sub BehaviorBridgeEntered(ByVal token As String)
    If Not pInvoking Or token <> pToken Then Exit Sub
    If Len(pBehaviorBridgeStage) > 0 Then Exit Sub
    pBehaviorBridgeStage = "entered"
End Sub

Public Sub BehaviorBridgeFinished(ByVal token As String, ByVal number As Long, ByVal description As String)
    If Not pInvoking Or token <> pToken Then Exit Sub
    If pBehaviorBridgeStage <> "entered" Then Exit Sub
    pBehaviorBridgeError = number
    pBehaviorBridgeDescription = description
    pBehaviorBridgeStage = "returned"
End Sub

Private Sub ValidateBehaviorBridge(ByVal result As Variant)
    If pBehaviorBridgeError <> 0 Then _
        Err.Raise pBehaviorBridgeError, pBehaviorBridgeProcedure, pBehaviorBridgeDescription
    If pBehaviorBridgeStage = "returned" Then Exit Sub
    If Len(pBehaviorBridgeStage) = 0 Then
        Err.Raise 5, "MacroRunner", "Tidak menerima konfirmasi masuk " & pBehaviorBridgeProcedure & "." & vbCrLf & _
            BridgeResultType(result) & vbCrLf & _
            "Perbarui MRTargetBridge pada GMS tujuan dan compile kedua project. " & _
            "Pastikan prosedur tersebut tersedia dan memberi callback masuk/hasil."
    End If
    Err.Raise 5, "MacroRunner", pBehaviorBridgeProcedure & " sudah dimasuki, tetapi belum melaporkan hasil sesi." & vbCrLf & _
        BridgeResultType(result) & vbCrLf & "Periksa error/breakpoint pada validator atau session target."
End Sub

Private Function BridgeResultType(ByVal result As Variant) As String
    BridgeResultType = "Nilai balik RunMacro: " & TypeName(result) & " (VarType=" & CStr(VarType(result)) & ")."
End Function

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
    If pPreflighting Then caption = "Memvalidasi seluruh antrean sebelum eksekusi."
    If CanConfigure And pSelected > 0 Then caption = caption & " | " & pParser.QueueLabel(pSequences(pSelected).SequenceText)
    pView.DisplayState pRun.StatisticsText(), caption, CanConfigure And pSelected > 0, _
        Not pInvoking And pRun.State = "waiting", CanConfigure, CanClose
End Sub

Public Sub ShowError(ByVal operation As String, ByVal number As Long, ByVal description As String)
    MsgBox "Operasi: " & operation & vbCrLf & "Error " & CStr(number) & vbCrLf & _
        "Description: " & description, vbExclamation, "Macro Runner"
End Sub
