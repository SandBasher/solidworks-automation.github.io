Attribute VB_Name = "RenameSheetsViewsFromModelProperties"
Option Explicit

' Rename Drawing Sheets and Views from Model Properties
' Free SOLIDWORKS VBA macro from sandbasher.github.io/solidworks-automation.github.io
' Provided as-is with no warranty. Test on copied files before production use.
' Assumptions: Open drawing with model views; Model custom properties

Const swDocPART As Long = 1
Const swDocASSEMBLY As Long = 2
Const swDocDRAWING As Long = 3
Const swOpenDocOptions_Silent As Long = 1
Const swSaveAsOptions_Silent As Long = 1

Dim swApp As SldWorks.SldWorks

Sub main()
    On Error GoTo Fail
    Set swApp = Application.SldWorks
    RunMacro
    Exit Sub
Fail:
    MsgBox "Macro failed: " & Err.Description, vbCritical, "Rename Drawing Sheets and Views from Model Properties"
End Sub

Private Function ActiveModel() As SldWorks.ModelDoc2
    Set ActiveModel = swApp.ActiveDoc
    If ActiveModel Is Nothing Then Err.Raise vbObjectError + 100, , "Open a SOLIDWORKS document first."
End Function

Private Function EnsureSaved(model As SldWorks.ModelDoc2) As String
    EnsureSaved = model.GetPathName
    If Len(EnsureSaved) = 0 Then Err.Raise vbObjectError + 101, , "Save the active document before running this macro."
End Function

Private Function FolderOf(filePath As String) As String
    FolderOf = Left$(filePath, InStrRev(filePath, "\") - 1)
End Function

Private Function BaseName(filePath As String) As String
    Dim nameOnly As String
    nameOnly = Mid$(filePath, InStrRev(filePath, "\") + 1)
    If InStrRev(nameOnly, ".") > 0 Then nameOnly = Left$(nameOnly, InStrRev(nameOnly, ".") - 1)
    BaseName = nameOnly
End Function

Private Function CleanFileName(value As String) As String
    Dim badChars As Variant
    Dim i As Long
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    CleanFileName = Trim$(value)
    For i = LBound(badChars) To UBound(badChars)
        CleanFileName = Replace(CleanFileName, badChars(i), "_")
    Next i
    Do While InStr(CleanFileName, "__") > 0
        CleanFileName = Replace(CleanFileName, "__", "_")
    Loop
    If Len(CleanFileName) = 0 Then CleanFileName = "export"
End Function

Private Function CombinePath(folderPath As String, fileName As String) As String
    If Right$(folderPath, 1) = "\" Then
        CombinePath = folderPath & fileName
    Else
        CombinePath = folderPath & "\" & fileName
    End If
End Function

Private Sub EnsureFolder(folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function GetProperty(model As SldWorks.ModelDoc2, propName As String) As String
    Dim mgr As SldWorks.CustomPropertyManager
    Dim rawValue As String
    Dim resolvedValue As String
    Set mgr = model.Extension.CustomPropertyManager(model.ConfigurationManager.ActiveConfiguration.Name)
    mgr.Get4 propName, False, rawValue, resolvedValue
    If Len(resolvedValue) = 0 Then
        Set mgr = model.Extension.CustomPropertyManager("")
        mgr.Get4 propName, False, rawValue, resolvedValue
    End If
    GetProperty = Trim$(resolvedValue)
End Function

Private Function PropertyFileStem(model As SldWorks.ModelDoc2, fallbackPath As String) As String
    Dim parts As Collection
    Dim value As Variant
    Dim result As String
    Set parts = New Collection
    AddIfPresent parts, GetProperty(model, "PartNo")
    AddIfPresent parts, GetProperty(model, "Part Number")
    AddIfPresent parts, GetProperty(model, "Revision")
    AddIfPresent parts, GetProperty(model, "Description")
    For Each value In parts
        If Len(result) > 0 Then result = result & "_"
        result = result & CStr(value)
    Next value
    If Len(result) = 0 Then result = BaseName(fallbackPath)
    PropertyFileStem = CleanFileName(result)
End Function

Private Sub AddIfPresent(items As Collection, value As String)
    If Len(Trim$(value)) > 0 Then items.Add Trim$(value)
End Sub

Private Function PickFolder(prompt As String) As String
    Dim shellApp As Object
    Dim folder As Object
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, prompt, 0)
    If folder Is Nothing Then Err.Raise vbObjectError + 102, , "No folder selected."
    PickFolder = folder.Self.Path
End Function

Private Function PickFile(prompt As String, filterDescription As String, extension As String) As String
    Dim dialog As Object
    Set dialog = CreateObject("MSComDlg.CommonDialog")
    dialog.DialogTitle = prompt
    dialog.Filter = filterDescription & "|*" & extension
    dialog.ShowOpen
    PickFile = dialog.FileName
    If Len(PickFile) = 0 Then Err.Raise vbObjectError + 103, , "No file selected."
End Function

Private Sub RunMacro()
    Dim drawing As SldWorks.ModelDoc2
    Dim view As SldWorks.View
    Dim refPath As String
    Dim refModel As SldWorks.ModelDoc2
    Dim errors As Long, warnings As Long, count As Long
    Set drawing = ActiveModel()
    If drawing.GetType <> swDocDRAWING Then Err.Raise vbObjectError + 900, , "Open a drawing."
    Set view = drawing.GetFirstView
    If Not view Is Nothing Then Set view = view.GetNextView
    Do While Not view Is Nothing
        refPath = view.GetReferencedModelName
        If Len(refPath) > 0 Then
            Set refModel = swApp.OpenDoc6(refPath, DocumentTypeFromPath(refPath), swOpenDocOptions_Silent, "", errors, warnings)
            If Not refModel Is Nothing Then
                If "rename-sheets-views-from-model-properties" = "title-block-sync" Then
                    CopyProperty refModel, drawing, "PartNo"
                    CopyProperty refModel, drawing, "Revision"
                    CopyProperty refModel, drawing, "Description"
                    CopyProperty refModel, drawing, "Material"
                Else
                    view.Name = CleanFileName(PropertyFileStem(refModel, refPath))
                End If
                swApp.CloseDoc refModel.GetTitle
                count = count + 1
            End If
        End If
        Set view = view.GetNextView
    Loop
    drawing.ForceRebuild3 False
    MsgBox "Updated " & count & " drawing references.", vbInformation, "Drawing update complete"
End Sub

Private Sub CopyProperty(sourceModel As SldWorks.ModelDoc2, targetDrawing As SldWorks.ModelDoc2, propName As String)
    Dim value As String
    value = GetProperty(sourceModel, propName)
    If Len(value) > 0 Then targetDrawing.Extension.CustomPropertyManager("").Add3 propName, 30, value, 2
End Sub

Private Function DocumentTypeFromPath(filePath As String) As Long
    If UCase$(Right$(filePath, 6)) = "SLDASM" Then DocumentTypeFromPath = swDocASSEMBLY Else DocumentTypeFromPath = swDocPART
End Function
