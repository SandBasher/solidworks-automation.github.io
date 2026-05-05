Attribute VB_Name = "BomTableToCsvExcel"
Option Explicit

' BOM and Table Export to CSV or Excel
' Free SOLIDWORKS VBA macro from sandbasher.github.io/solidworks-automation.github.io
' Provided as-is with no warranty. Test on copied files before production use.
' Assumptions: Open drawing or assembly containing BOM/table annotations

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
    MsgBox "Macro failed: " & Err.Description, vbCritical, "BOM and Table Export to CSV or Excel"
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
    Dim model As SldWorks.ModelDoc2
    Dim modelPath As String
    Dim outFolder As String
    Dim csvPath As String
    Set model = ActiveModel()
    modelPath = EnsureSaved(model)
    outFolder = CombinePath(FolderOf(modelPath), "BOM")
    EnsureFolder outFolder
    csvPath = CombinePath(outFolder, CleanFileName(BaseName(modelPath) & IIf("bom-export-excel-csv" = "bom-property-checker", "_bom_property_check", "_table_export")) & ".csv")
    ExportTables model, csvPath, "bom-export-excel-csv"
    MsgBox "CSV written to " & csvPath, vbInformation, "BOM export complete"
End Sub

Private Sub ExportTables(model As SldWorks.ModelDoc2, csvPath As String, mode As String)
    Dim outFile As Integer
    Dim view As SldWorks.View
    Dim tableAnn As SldWorks.TableAnnotation
    Dim row As Long, col As Long
    outFile = FreeFile
    Open csvPath For Output As #outFile
    If model.GetType = swDocDRAWING Then
        Set view = model.GetFirstView
        Do While Not view Is Nothing
            Set tableAnn = view.GetFirstTableAnnotation
            Do While Not tableAnn Is Nothing
                WriteTable tableAnn, outFile, mode
                Set tableAnn = tableAnn.GetNext
            Loop
            Set view = view.GetNextView
        Loop
    Else
        Err.Raise vbObjectError + 400, , "Open a drawing with BOM/table annotations for this v1 macro."
    End If
    Close #outFile
End Sub

Private Sub WriteTable(tableAnn As SldWorks.TableAnnotation, outFile As Integer, mode As String)
    Dim row As Long, col As Long, line As String, cellText As String
    If mode = "bom-property-checker" Then Print #outFile, "Row,Column,Issue,Value"
    For row = 0 To tableAnn.RowCount - 1
        line = ""
        For col = 0 To tableAnn.ColumnCount - 1
            cellText = tableAnn.Text(row, col)
            If mode = "bom-property-checker" Then
                If row > 0 And Len(Trim$(cellText)) = 0 Then Print #outFile, row & "," & col & ",""Missing visible BOM value"","""""
            Else
                If col > 0 Then line = line & ","
                line = line & CsvEscape(cellText)
            End If
        Next col
        If mode <> "bom-property-checker" Then Print #outFile, line
    Next row
End Sub

Private Function CsvEscape(value As String) As String
    CsvEscape = """" & Replace(value, """", """""") & """"
End Function
