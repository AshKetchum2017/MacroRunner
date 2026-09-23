Option Explicit

' UserForm (Name) = MacroQueue.
Private pPresenter As MRPresenter

Public Sub Configure(ByVal presenter As MRPresenter)
    Set pPresenter = presenter
    lbxMacroLists.MultiSelect = fmMultiSelectSingle
End Sub

Public Sub DisplaySequences(ByVal labels As Collection, ByVal selected As Long)
    Dim label As Variant
    lbxMacroLists.Clear
    For Each label In labels
        lbxMacroLists.AddItem CStr(label)
    Next label
    If selected > 0 And selected <= lbxMacroLists.ListCount Then lbxMacroLists.ListIndex = selected - 1
    RefreshButtons
End Sub

Private Sub cmdRemove_Click()

    Dim selectedIndex As Long
    Dim selectedLabel As String
    Dim response As VbMsgBoxResult

    If pPresenter Is Nothing Then Exit Sub

    selectedIndex = lbxMacroLists.ListIndex
    If selectedIndex < 0 Then Exit Sub

    selectedLabel = CStr(lbxMacroLists.List(selectedIndex))

    response = MsgBox( _
        "Hapus daftar macro berikut?" & vbCrLf & vbCrLf & _
        selectedLabel, _
        vbYesNo Or vbQuestion Or vbDefaultButton2, _
        "Macro Queue")

    If response <> vbYes Then Exit Sub

    pPresenter.RemoveSequence Me, selectedIndex + 1

End Sub

Private Sub cmdClear_Click()

    Dim response As VbMsgBoxResult

    If pPresenter Is Nothing Then Exit Sub
    If lbxMacroLists.ListCount = 0 Then Exit Sub

    response = MsgBox( _
        "Hapus seluruh daftar macro?" & vbCrLf & vbCrLf & _
        CStr(lbxMacroLists.ListCount) & _
        " konfigurasi akan dihapus." & vbCrLf & _
        "Tindakan ini tidak dapat dibatalkan.", _
        vbYesNo Or vbExclamation Or vbDefaultButton2, _
        "Macro Queue")

    If response <> vbYes Then Exit Sub

    pPresenter.ClearSequences Me

End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdModify_Click()
    If pPresenter Is Nothing Then Exit Sub

    pPresenter.ModifySequence Me, lbxMacroLists.ListIndex + 1
End Sub

Private Sub cmdSelect_Click()
    If pPresenter Is Nothing Then Exit Sub

    If pPresenter.SelectSequence( _
        lbxMacroLists.ListIndex + 1 _
    ) Then
        Me.Hide
    End If
End Sub

Private Sub cmdSet_Click()
    pPresenter.OpenSelection Me
End Sub

Private Sub lbxMacroLists_Click()
    RefreshButtons
End Sub

Private Sub RefreshButtons()
    cmdModify.Enabled = (lbxMacroLists.ListIndex >= 0)
    cmdRemove.Enabled = (lbxMacroLists.ListIndex >= 0)
    cmdSelect.Enabled = (lbxMacroLists.ListIndex >= 0)
    cmdClear.Enabled = (lbxMacroLists.ListCount > 0)
End Sub
