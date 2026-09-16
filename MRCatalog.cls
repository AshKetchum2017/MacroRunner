Option Explicit

' Class Module (Name) = MRCatalog.
Private pItems As Collection

Public Sub Refresh()
    Dim folder As String, fileName As String
    Dim item As MRMacroDefinition
    Dim items As New Collection
    Dim operation As String, errorNumber As Long
    Dim errorSource As String, errorDescription As String
    On Error GoTo Failed
    operation = "Application.UserDataPath / UserGMSFolder"
    folder = UserGMSFolder(Application.UserDataPath)
    operation = "FolderExists: " & folder
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(folder) Then _
        Err.Raise 76, "MacroSelection", "Folder GMS CorelDRAW aktif tidak ditemukan: " & folder
    operation = "Dir: " & folder & "*.gms"
    fileName = Dir$(folder & "*.gms")
    Do While Len(fileName) > 0
        operation = "Mencatat nama file: " & fileName
        Set item = New MRMacroDefinition
        item.DisplayName = Left$(fileName, Len(fileName) - 4)
        If InStr(item.DisplayName, ";") > 0 Then item.Problem = "Nama file GMS mengandung pemisah ';' dan tidak dapat dipakai dalam antrean."
        items.Add item
        operation = "Dir: file berikutnya di " & folder
        fileName = Dir$()
    Loop
    Set pItems = items
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Raise errorNumber, "MRCatalog.Refresh", "Operasi katalog: " & operation & vbCrLf & _
        "Source: " & errorSource & vbCrLf & errorDescription
End Sub

Public Function UserGMSFolder(ByVal userDataPath As String) As String
    Dim fso As Object, folder As String, leaf As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = Replace(Trim$(userDataPath), "/", "\")
    Do While Len(folder) > 0
        If Right$(folder, 1) <> "\" Then Exit Do
        folder = Left$(folder, Len(folder) - 1)
    Loop
    If Len(folder) = 0 Then Err.Raise 76, "MacroSelection", "UserDataPath CorelDRAW aktif kosong."
    leaf = LCase$(fso.GetFileName(folder))
    Select Case leaf
        Case "custom data"
            ' CorelDRAW 2024 reports <active suite>\Custom Data\, not Draw\.
            folder = fso.BuildPath(fso.GetParentFolderName(folder), "Draw")
        Case "draw"
            ' Keep support for hosts whose UserDataPath already ends in Draw\.
        Case Else
            Err.Raise 76, "MacroSelection", "Struktur UserDataPath CorelDRAW belum dikenali: " & userDataPath
    End Select
    UserGMSFolder = fso.BuildPath(folder, "GMS") & "\"
End Function

Public Property Get Items() As Collection
    Set Items = pItems
End Property

Public Function Resolve(ByVal name As String) As MRMacroDefinition
    Dim item As MRMacroDefinition
    If pItems Is Nothing Then Refresh
    For Each item In pItems
        If StrComp(item.DisplayName, name, vbTextCompare) = 0 Then
            If Len(item.Problem) > 0 Then Err.Raise 5, "MacroRunner", name & ": " & item.Problem
            Set Resolve = item
            Exit Function
        End If
    Next item
    Err.Raise 53, "MacroRunner", "GMS tidak ditemukan pada CorelDRAW aktif: " & name
End Function

' Only Process needs executable metadata. Selection/Add/Save use file names.
' Do not enumerate the procedures of unrelated GMS projects while opening a form.
Public Function ResolveForRun(ByVal name As String) As MRMacroDefinition
    Dim item As MRMacroDefinition
    Dim projects As GMSProjects, project As GMSProject, matched As GMSProject
    Dim macros As GMSMacros, macro As GMSMacro
    Dim targetPath As String, projectPath As String, entryName As String, formName As String
    Dim discoveredNames As String
    Dim modal As Boolean, matches As Long, projectIndex As Long, macroIndex As Long
    Dim operation As String, errorNumber As Long
    Dim errorSource As String, errorDescription As String
    On Error GoTo Failed
    operation = "Memeriksa nama GMS: " & name
    Set item = Resolve(name)
    item.ProjectName = vbNullString
    item.FormName = vbNullString
    operation = "Application.UserDataPath / UserGMSFolder"
    targetPath = UserGMSFolder(Application.UserDataPath) & item.DisplayName & ".gms"
    operation = "Application.GMSManager.Projects"
    Set projects = Application.GMSManager.Projects
    operation = "Enumerasi GMSProjects untuk " & targetPath
    For Each project In projects
        projectIndex = projectIndex + 1
        operation = "GMSProject.FullFileName proyek #" & CStr(projectIndex) & " (target: " & name & ")"
        projectPath = project.FullFileName
        If StrComp(projectPath, targetPath, vbTextCompare) = 0 Then
            Set matched = project
            Exit For
        End If
        operation = "Enumerasi proyek setelah #" & CStr(projectIndex) & ": " & projectPath
    Next project
    If matched Is Nothing Then Err.Raise 5, "MacroRunner", "GMS belum dimuat oleh CorelDRAW: " & name
    operation = "GMSProject.Name: " & targetPath
    item.ProjectName = matched.Name
    operation = "Membaca registrasi form: " & item.DisplayName
    item.FormName = RegisteredForm(item.DisplayName, modal)
    If Len(item.FormName) > 0 Then
        item.Modal = modal
        Set ResolveForRun = item
        Exit Function
    End If
    operation = "GMSProject.Macros: " & targetPath
    Set macros = matched.Macros
    operation = "Enumerasi prosedur: " & targetPath
    For Each macro In macros
        macroIndex = macroIndex + 1
        operation = "GMSMacro.Name #" & CStr(macroIndex) & ": " & targetPath
        entryName = macro.Name
        discoveredNames = discoveredNames & vbCrLf & "[" & entryName & "]"
        If InStrRev(entryName, ".") > 0 Then entryName = Mid$(entryName, InStrRev(entryName, ".") + 1)
        formName = DefaultForm(entryName, modal)
        If Len(formName) > 0 Then
            matches = matches + 1
            item.FormName = formName
            item.Modal = modal
        End If
        operation = "Enumerasi prosedur setelah " & entryName & ": " & targetPath
    Next macro
    operation = "Memvalidasi pembuka bawaan: " & name
    If macroIndex = 0 Then discoveredNames = vbCrLf & "(tidak ada prosedur yang dilaporkan)"
    If matches = 0 Then Err.Raise 5, "MacroRunner", "Tidak ada prosedur dari CorelDRAW yang cocok dengan pembuka terdaftar: " & name & vbCrLf & _
        "Project: " & item.ProjectName & vbCrLf & "Prosedur dari CorelDRAW (" & CStr(macroIndex) & "):" & discoveredNames
    If matches > 1 Then Err.Raise 5, "MacroRunner", "Ada beberapa pembuka bawaan dalam GMS " & name & _
        "; pemilihan prosedur belum tersedia." & vbCrLf & "Prosedur dari CorelDRAW:" & discoveredNames
    Set ResolveForRun = item
    Exit Function
Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Raise errorNumber, "MRCatalog.ResolveForRun", "Operasi katalog: " & operation & vbCrLf & _
        "Source: " & errorSource & vbCrLf & errorDescription
End Function

' Explicit file -> form registration based on source and user-confirmed VBE code.
' Opening a registered form through MRTargetBridge does not require its wrapper
' to appear in GMSProject.Macros. The bridge still validates the form and hook.
Private Function RegisteredForm(ByVal fileBaseName As String, ByRef modal As Boolean) As String
    modal = False
    Select Case LCase$(fileBaseName)
        Case "autodistributeuf": RegisteredForm = "AutoDistributeMenu"
    End Select
End Function

Private Function DefaultForm(ByVal entryName As String, ByRef modal As Boolean) As String
    modal = False
    Select Case LCase$(entryName)
        Case "mastermenu": DefaultForm = "MasterWizard": modal = True
        Case "autodistributemenu": DefaultForm = "AutoDistributeMenu"
        Case "autolabelmastermenu": DefaultForm = "AutoLabelWizard": modal = True
        Case "autolabelmasterbpmenu": DefaultForm = "AutoLabelWizardBP": modal = True
        Case "autosavencreate": DefaultForm = "AutoSNC": modal = True
        Case "exportrelatedwizard": DefaultForm = "ExportRelatedMenu"
        Case "objectcontainer": DefaultForm = "ObjectContainerMenu"
        Case "pagerelatedwizard": DefaultForm = "PageRelatedMenu"
        Case "pericountmastermenu": DefaultForm = "PeriCountMenu"
        Case "startpointmastermenu": DefaultForm = "StartPointMenu"
        Case "polakartunamalengkungmastermenu": DefaultForm = "PolaKartuNamaLengkungWizard": modal = True
        Case "dashlinebreaker": DefaultForm = "DashLineBreakerMenu"
    End Select
End Function
