Option Explicit

' UserForm (Name) = MacroQueue.
Private pPresenter As MRPresenter

Public Sub Configure(ByVal presenter As MRPresenter)
    Set pPresenter = presenter
    lbxMacroQueue.MultiSelect = fmMultiSelectSingle
End Sub

Public Sub DisplaySequences(ByVal labels As Collection, ByVal selected As Long)
    Dim label As Variant
    lbxMacroQueue.Clear
    For Each label In labels
        lbxMacroQueue.AddItem CStr(label)
    Next label
    If selected > 0 And selected <= lbxMacroQueue.ListCount Then lbxMacroQueue.ListIndex = selected - 1
    RefreshButtons
End Sub

Private Sub cmdClear_Click()
    pPresenter.ClearSequences Me
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdRemove_Click()
    pPresenter.RemoveSequence Me, lbxMacroQueue.ListIndex + 1
End Sub

Private Sub cmdSelect_Click()
    If pPresenter Is Nothing Then Exit Sub

    If pPresenter.SelectSequence( _
        lbxMacroQueue.ListIndex + 1 _
    ) Then
        Me.Hide
    End If
End Sub

Private Sub cmdSet_Click()
    pPresenter.OpenSelection Me
End Sub

Private Sub lbxMacroQueue_Click()
    RefreshButtons
End Sub

Private Sub RefreshButtons()
    cmdRemove.Enabled = (lbxMacroQueue.ListIndex >= 0)
    cmdSelect.Enabled = (lbxMacroQueue.ListIndex >= 0)
    cmdClear.Enabled = (lbxMacroQueue.ListCount > 0)
End Sub
