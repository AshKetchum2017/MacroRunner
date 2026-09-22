Option Explicit

' UserForm (Name) = MacroRunnerMenu. Add CommandButton cmdContinue in VBE.
Private pPresenter As MRPresenter

Private Sub UserForm_Initialize()
    On Error GoTo Failed
    lbxStatistics.Font.Name = "Courier New"
    lbxStatistics.ColumnCount = 1
    Set pPresenter = New MRPresenter
    pPresenter.Initialize Me
    Exit Sub
Failed:
    cmdProcess.Enabled = False
    cmdContinue.Enabled = False
    cmdCopyStatistic.Enabled = False
    cmdQueueSettings.Enabled = False
    MsgBox "Gagal memuat Macro Runner (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not pPresenter Is Nothing Then
        If Not pPresenter.CanClose Then
            Cancel = 1
            MsgBox "Tutup macro yang sedang berjalan terlebih dahulu.", vbExclamation, "Macro Runner"
            Exit Sub
        End If
        pPresenter.Detach
        Set pPresenter = Nothing
    End If
End Sub

Private Sub cmdCopyStatistic_Click()
    If Not pPresenter Is Nothing Then pPresenter.CopyStatistics
End Sub

Private Sub cmdProcess_Click()
    If Not pPresenter Is Nothing Then pPresenter.Process
End Sub

Private Sub cmdContinue_Click()
    If Not pPresenter Is Nothing Then pPresenter.ContinueRun
End Sub

Private Sub cmdQueueSettings_Click()
    If Not pPresenter Is Nothing Then pPresenter.OpenQueue
End Sub

Public Sub DisplayState(ByVal statistics As String, ByVal status As String, ByVal canProcess As Boolean, _
                        ByVal canContinue As Boolean, ByVal canConfigure As Boolean, ByVal canClose As Boolean)
    Dim lines As Variant, line As Variant, width As Long
    lbxStatistics.Clear
    If Len(statistics) > 0 Then
        lines = Split(statistics, vbCrLf)
        For Each line In lines
            lbxStatistics.AddItem CStr(line)
            If Len(CStr(line)) > width Then width = Len(CStr(line))
        Next line
        lbxStatistics.ColumnWidths = CStr(width * CSng(lbxStatistics.Font.Size) * 0.65 + 12) & " pt"
    End If
    Me.Caption = "Macro Runner - " & status
    cmdProcess.Enabled = canProcess
    cmdContinue.Enabled = canContinue
    cmdQueueSettings.Enabled = canConfigure
    cmdClose.Enabled = canClose
    cmdCopyStatistic.Enabled = (Len(statistics) > 0)
End Sub

Private Sub lbxStatistics_Click()
End Sub
