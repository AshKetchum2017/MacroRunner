Option Explicit

' UserForm (Name) = MacroSelection. txbMacroBehavior is reserved for next scope.
Private pPresenter As MRPresenter
Private pPrevious As String
Private pEditing As Boolean

Public Sub Configure(ByVal presenter As MRPresenter, ByVal items As Collection)
    Dim item As MRMacroDefinition
    Dim operation As String, errorNumber As Long
    Dim errorSource As String, errorDescription As String
    On Error GoTo Failed
    operation = "Menghubungkan Presenter"
    Set pPresenter = presenter
    operation = "cmbMacroLists.Clear"
    cmbMacroLists.Clear
    operation = "cmbMacroLists.Style"
    cmbMacroLists.Style = fmStyleDropDownList
    For Each item In items
        operation = "cmbMacroLists.AddItem: " & item.DisplayName
        cmbMacroLists.AddItem item.DisplayName
    Next item
    pPrevious = vbNullString
    operation = "txbSelectedMacro.Value"
    txbSelectedMacro.Value = vbNullString
    operation = "txbMacroBehavior.Enabled"
    txbMacroBehavior.Enabled = False
    operation = "cmdAdd.Enabled"
    cmdAdd.Enabled = False
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Raise errorNumber, "MacroSelection.Configure", "Operasi kontrol: " & operation & vbCrLf & _
        "Source: " & errorSource & vbCrLf & errorDescription
End Sub

Private Sub cmbMacroLists_Change()
    cmdAdd.Enabled = (cmbMacroLists.ListIndex >= 0)
End Sub

Private Sub cmdAdd_Click()
    Dim value As String
    On Error GoTo Failed
    If cmbMacroLists.ListIndex < 0 Then Exit Sub
    value = pPresenter.AppendMacro(CStr(txbSelectedMacro.Value), CStr(cmbMacroLists.Value))
    pEditing = True
    txbSelectedMacro.Value = value
    pPrevious = value
    pEditing = False
    txbSelectedMacro.SetFocus
    txbSelectedMacro.SelStart = Len(value)
    Exit Sub
Failed:
    pEditing = False
    pPresenter.ShowError "Menambahkan macro", Err.Number, Err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If pPresenter.SaveSelection(CStr(txbSelectedMacro.Value)) Then Unload Me
End Sub

Private Sub txbSelectedMacro_Change()
    Dim current As String, rewritten As String, caret As Long
    If pEditing Or pPresenter Is Nothing Then Exit Sub
    current = CStr(txbSelectedMacro.Value)
    caret = txbSelectedMacro.SelStart
    rewritten = pPresenter.EditTokens(pPrevious, current, caret)
    If rewritten <> current Then
        pEditing = True
        txbSelectedMacro.Value = rewritten
        txbSelectedMacro.SelStart = caret
        pEditing = False
    End If
    pPrevious = rewritten
End Sub

Private Sub txbMacroBehavior_Change()
End Sub
