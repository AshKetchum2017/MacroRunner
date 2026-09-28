Option Explicit

' UserForm (Name) = MacroQueue. frmMacroLists is an MSForms.Frame.
Private pPresenter As MRPresenter
Private pLabels As Collection
Private pRows As Collection
Private pRowCount As Long
Private pSelectedIndex As Long

Public Sub Configure(ByVal presenter As MRPresenter)
    Set pPresenter = presenter
    Set pLabels = New Collection
    Set pRows = New Collection
    frmMacroLists.ScrollBars = fmScrollBarsVertical
End Sub

Public Sub DisplaySequences(ByVal labels As Collection, ByVal selected As Long)
    Dim title As MSForms.Label, sequence As MSForms.Label, separator As MSForms.Label
    Dim sink As MRQueueRow, i As Long, splitAt As Long
    Dim top As Single, titleHeight As Single, sequenceHeight As Single
    Dim rowWidth As Single, rowText As String
    Dim errorNumber As Long, errorDescription As String

    On Error GoTo DisplayFailed
    ClearRows
    Set pLabels = labels
    If selected > 0 And selected <= labels.Count Then
        pSelectedIndex = selected
    Else
        pSelectedIndex = 0
    End If

    rowWidth = frmMacroLists.Width - 18
    top = 6
    For i = 1 To labels.Count
        rowText = CStr(labels(i))
        splitAt = InStr(1, rowText, vbCrLf, vbBinaryCompare)

        Set title = frmMacroLists.Controls.Add("Forms.Label.1", "mrTitle" & CStr(i), True)
        title.Left = 6
        title.Top = top
        title.Width = rowWidth
        title.WordWrap = True
        title.Font.Name = "Tahoma"
        title.Font.Size = 10
        If splitAt > 0 Then title.Caption = Left$(rowText, splitAt - 1)
        title.AutoSize = True
        titleHeight = title.Height
        title.AutoSize = False
        title.Width = rowWidth
        If splitAt = 0 Then
            titleHeight = 0
            title.Height = 1
        Else
            title.Height = titleHeight
        End If
        title.Visible = (splitAt > 0)
        title.BackStyle = fmBackStyleOpaque

        Set sequence = frmMacroLists.Controls.Add("Forms.Label.1", "mrSequence" & CStr(i), True)
        sequence.Left = 6
        sequence.Top = top + titleHeight
        sequence.Width = rowWidth
        sequence.WordWrap = True
        sequence.Font.Name = "Tahoma"
        sequence.Font.Size = 8
        If splitAt > 0 Then
            sequence.Caption = Mid$(rowText, splitAt + Len(vbCrLf))
        Else
            sequence.Caption = rowText
        End If
        sequence.AutoSize = True
        sequenceHeight = sequence.Height
        sequence.AutoSize = False
        sequence.Width = rowWidth
        If titleHeight + sequenceHeight < 32 Then sequenceHeight = 32 - titleHeight
        sequence.Height = sequenceHeight + 6
        sequence.BackStyle = fmBackStyleOpaque

        Set sink = New MRQueueRow
        sink.Bind Me, title, sequence, i
        pRows.Add sink

        top = sequence.Top + sequence.Height + 3
        Set separator = frmMacroLists.Controls.Add("Forms.Label.1", "mrSeparator" & CStr(i), True)
        separator.Left = 6
        separator.Top = top
        separator.Width = rowWidth
        separator.Height = 1
        separator.BackStyle = fmBackStyleOpaque
        separator.BackColor = RGB(140, 140, 140)
        top = top + 9
        pRowCount = i
    Next i
    frmMacroLists.ScrollHeight = top + 6
    PaintSelection
    RefreshButtons
    If pSelectedIndex > 0 Then
        Set sequence = frmMacroLists.Controls("mrSequence" & CStr(pSelectedIndex))
        If sequence.Top + sequence.Height > frmMacroLists.Height Then _
            frmMacroLists.ScrollTop = frmMacroLists.Controls("mrTitle" & CStr(pSelectedIndex)).Top
    End If
    Exit Sub

DisplayFailed:
    errorNumber = Err.Number: errorDescription = Err.Description
    On Error Resume Next
    If i > pRowCount Then
        frmMacroLists.Controls.Remove "mrSeparator" & CStr(i)
        frmMacroLists.Controls.Remove "mrSequence" & CStr(i)
        frmMacroLists.Controls.Remove "mrTitle" & CStr(i)
    End If
    ClearRows
    On Error GoTo 0
    Err.Raise errorNumber, "MacroQueue.DisplaySequences", errorDescription
End Sub

Public Sub SelectRow(ByVal index As Long)
    If pLabels Is Nothing Then Exit Sub
    If index < 1 Or index > pLabels.Count Then Exit Sub
    pSelectedIndex = index
    PaintSelection
    RefreshButtons
End Sub

Private Sub ClearRows()
    Dim i As Long
    ReleaseRows
    Set pRows = New Collection
    For i = pRowCount To 1 Step -1
        frmMacroLists.Controls.Remove "mrSeparator" & CStr(i)
        frmMacroLists.Controls.Remove "mrSequence" & CStr(i)
        frmMacroLists.Controls.Remove "mrTitle" & CStr(i)
    Next i
    pRowCount = 0
    frmMacroLists.ScrollTop = 0
End Sub

Private Sub PaintSelection()
    Dim title As MSForms.Label, sequence As MSForms.Label, i As Long
    For i = 1 To pRowCount
        Set title = frmMacroLists.Controls("mrTitle" & CStr(i))
        Set sequence = frmMacroLists.Controls("mrSequence" & CStr(i))
        If i = pSelectedIndex Then
            title.BackColor = RGB(0, 120, 215)
            title.ForeColor = RGB(255, 255, 255)
            sequence.BackColor = RGB(0, 120, 215)
            sequence.ForeColor = RGB(255, 255, 255)
        Else
            title.BackColor = RGB(255, 255, 255)
            title.ForeColor = RGB(0, 0, 0)
            sequence.BackColor = RGB(255, 255, 255)
            sequence.ForeColor = RGB(0, 0, 0)
        End If
    Next i
End Sub

Private Sub cmdRemove_Click()
    Dim response As VbMsgBoxResult
    If pPresenter Is Nothing Or pSelectedIndex = 0 Then Exit Sub
    response = MsgBox( _
        "Hapus daftar macro berikut?" & vbCrLf & vbCrLf & _
        CStr(pLabels(pSelectedIndex)), _
        vbYesNo Or vbQuestion Or vbDefaultButton2, _
        "Macro Queue")
    If response <> vbYes Then Exit Sub
    pPresenter.RemoveSequence Me, pSelectedIndex
End Sub

Private Sub cmdClear_Click()
    Dim response As VbMsgBoxResult
    If pPresenter Is Nothing Then Exit Sub
    If pLabels.Count = 0 Then Exit Sub
    response = MsgBox( _
        "Hapus seluruh daftar macro?" & vbCrLf & vbCrLf & _
        CStr(pLabels.Count) & " konfigurasi akan dihapus." & vbCrLf & _
        "Tindakan ini tidak dapat dibatalkan.", _
        vbYesNo Or vbExclamation Or vbDefaultButton2, _
        "Macro Queue")
    If response <> vbYes Then Exit Sub
    pPresenter.ClearSequences Me
End Sub

Private Sub cmdClose_Click()
    ReleaseRows
    Me.Hide
    If Not pPresenter Is Nothing Then pPresenter.QueueClosed
End Sub

Private Sub cmdModify_Click()
    If pPresenter Is Nothing Then Exit Sub
    pPresenter.ModifySequence Me, pSelectedIndex
End Sub

Private Sub cmdSelect_Click()
    If pPresenter Is Nothing Then Exit Sub
    If pPresenter.SelectSequence(pSelectedIndex) Then
        ReleaseRows
        Me.Hide
        pPresenter.QueueClosed
    End If
End Sub

Private Sub cmdSet_Click()
    If pPresenter Is Nothing Then Exit Sub
    pPresenter.OpenSelection Me
End Sub

Private Sub cmdUp_Click()
    If pPresenter Is Nothing Then Exit Sub
    pPresenter.MoveSequence Me, pSelectedIndex, -1
End Sub

Private Sub cmdDown_Click()
    If pPresenter Is Nothing Then Exit Sub
    pPresenter.MoveSequence Me, pSelectedIndex, 1
End Sub

Private Sub RefreshButtons()
    Dim count As Long
    If Not pLabels Is Nothing Then count = pLabels.Count
    cmdModify.Enabled = (pSelectedIndex > 0)
    cmdRemove.Enabled = (pSelectedIndex > 0)
    cmdSelect.Enabled = (pSelectedIndex > 0)
    cmdClear.Enabled = (count > 0)
    cmdUp.Enabled = (pSelectedIndex > 1)
    cmdDown.Enabled = (pSelectedIndex > 0 And pSelectedIndex < count)
End Sub

Public Sub ReleaseRows()
    Dim sink As MRQueueRow
    If Not pRows Is Nothing Then
        For Each sink In pRows
            sink.Unbind
        Next sink
    End If
    Set pRows = Nothing
End Sub

Public Sub DetachPresenter()
    Set pPresenter = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = 1
        ReleaseRows
        Me.Hide
        If Not pPresenter Is Nothing Then pPresenter.QueueClosed
    Else
        ReleaseRows
    End If
End Sub

Private Sub UserForm_Terminate()
    ReleaseRows
    Set pPresenter = Nothing
End Sub
