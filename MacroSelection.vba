Option Explicit

' UserForm (Name) = MacroSelection. Draft is committed only on Save.
Private pPresenter As MRPresenter
Private pPrevious As String
Private pEditing As Boolean
Private pEditIndex As Long

Public Sub Configure(ByVal presenter As MRPresenter, ByVal items As Collection, _
                     Optional ByVal sequenceText As String = vbNullString, _
                     Optional ByVal behaviorText As String = vbNullString, _
                     Optional ByVal editIndex As Long = 0)
    Dim item As MRMacroDefinition
    Dim operation As String, errorNumber As Long
    Dim errorSource As String, errorDescription As String
    On Error GoTo Failed
    operation = "Menghubungkan Presenter"
    Set pPresenter = presenter
    pEditIndex = editIndex
    operation = "cmbMacroLists.Clear"
    cmbMacroLists.Clear
    operation = "cmbMacroLists.Style"
    cmbMacroLists.Style = fmStyleDropDownList
    For Each item In items
        operation = "cmbMacroLists.AddItem: " & item.DisplayName
        cmbMacroLists.AddItem item.DisplayName
    Next item
    operation = "txbSelectedMacro.Value"
    pEditing = True
    txbSelectedMacro.Value = sequenceText
    pPrevious = sequenceText
    pEditing = False
    operation = "txbMacroBehavior.Enabled"
    txbMacroBehavior.Enabled = True
    txbMacroBehavior.MultiLine = True
    txbMacroBehavior.EnterKeyBehavior = True
    txbMacroBehavior.TabKeyBehavior = True
    txbMacroBehavior.ScrollBars = fmScrollBarsBoth
    txbMacroBehavior.WordWrap = False
    txbMacroBehavior.Value = behaviorText
    operation = "cmdAdd.Enabled"
    cmdAdd.Enabled = False
    Exit Sub
Failed:
    pEditing = False
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
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    If pPresenter.SaveSelection( _
        CStr(txbSelectedMacro.Value), _
        CStr(txbMacroBehavior.Value), _
        pEditIndex _
    ) Then
        Me.Hide
    End If
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

Private Sub txbMacroBehavior_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Const INDENT_SPACES As Long = 5
    Dim caret As Long
    If KeyCode <> vbKeyTab Or Shift <> 0 Then Exit Sub
    caret = txbMacroBehavior.SelStart
    ' Plain Tab always inserts spaces, including while a quoted value is incomplete.
    KeyCode = 0
    txbMacroBehavior.SelText = Space$(INDENT_SPACES)
    txbMacroBehavior.SelStart = caret + INDENT_SPACES
    txbMacroBehavior.SelLength = 0
End Sub
