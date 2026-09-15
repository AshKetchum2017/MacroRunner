Option Explicit

' Class Module (Name) = MRClipboard. Copy full Unicode statistics in one block.
#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal window As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal format As Long, ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal memory As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
#Else
Private Declare Function OpenClipboard Lib "user32" (ByVal window As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal format As Long, ByVal memory As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal memory As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal memory As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal memory As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As Long, ByVal source As Long, ByVal bytes As Long)
#End If

Public Sub CopyText(ByVal value As String)
#If VBA7 Then
    Dim memory As LongPtr, pointer As LongPtr
#Else
    Dim memory As Long, pointer As Long
#End If
    Dim opened As Boolean, errorNumber As Long, errorDescription As String
    On Error GoTo Failed
    memory = GlobalAlloc(&H42, LenB(value) + 2)
    If memory = 0 Then Err.Raise 7, "MacroRunner", "Memori clipboard tidak tersedia."
    pointer = GlobalLock(memory)
    If pointer = 0 Then Err.Raise 7, "MacroRunner", "Gagal mengunci memori clipboard."
    CopyMemory pointer, StrPtr(value), LenB(value)
    GlobalUnlock memory
    pointer = 0
    If OpenClipboard(GetActiveWindow()) = 0 Then Err.Raise 5, "MacroRunner", "Clipboard sedang digunakan. Coba salin lagi."
    opened = True
    If EmptyClipboard() = 0 Then Err.Raise 5, "MacroRunner", "Gagal menyiapkan clipboard."
    If SetClipboardData(13, memory) = 0 Then Err.Raise 5, "MacroRunner", "Gagal menyalin statistik."
    memory = 0 ' Windows now owns the allocation.
    CloseClipboard
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    If pointer <> 0 Then GlobalUnlock memory
    If memory <> 0 Then GlobalFree memory
    If opened Then CloseClipboard
    Err.Raise errorNumber, "MRClipboard.CopyText", errorDescription
End Sub
