Option Explicit

' Class Module (Name) = MRQueueRow. Keeps Click events for both row labels.
Private WithEvents pTitle As MSForms.Label
Private WithEvents pSequence As MSForms.Label
Private pOwner As Object
Private pIndex As Long

Public Sub Bind(ByVal owner As Object, ByVal title As MSForms.Label, _
                ByVal sequence As MSForms.Label, ByVal index As Long)
    Set pOwner = owner
    Set pTitle = title
    Set pSequence = sequence
    pIndex = index
End Sub

Public Sub Unbind()
    Set pTitle = Nothing
    Set pSequence = Nothing
    Set pOwner = Nothing
End Sub

Private Sub pTitle_Click()
    SelectRow
End Sub

Private Sub pSequence_Click()
    SelectRow
End Sub

Private Sub SelectRow()
    If Not pOwner Is Nothing Then CallByName pOwner, "SelectRow", VbMethod, pIndex
End Sub
