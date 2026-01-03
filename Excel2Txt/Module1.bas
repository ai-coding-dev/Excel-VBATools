Attribute VB_Name = "Module1"
Option Explicit

Sub Excel2Txt_Click()

On Error GoTo ErrorHandle
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call Excel2Txt( _
        Sheets("Excel2Txt").Cells(4, 3), _
        Sheets("Excel2Txt").Cells(6, 3))
    
    MsgBox "OK", vbInformation
ExitHandle:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandle:
    MsgBox Err.Description
    Resume ExitHandle

End Sub

Sub Excel2Txt(inputFolder As String, outputFolder As String)
    Dim fileObj As Object: Set fileObj = CreateObject("Scripting.FileSystemObject")
    Dim fileList As Collection: Set fileList = New Collection
    Dim filePath As Variant
    
    Call GetExcelFilesRecursive(fileObj.GetFolder(inputFolder), fileList)
    
    For Each filePath In fileList
        Call ExportSheetsToText(CStr(filePath), inputFolder, outputFolder)
    Next filePath
End Sub

Sub GetExcelFilesRecursive(folder As Object, fileList As Collection)
    Dim file As Object
    Dim subFolder As Object
    
    For Each file In folder.Files
        If LCase(GetExtensionName(file.Name)) = "xlsx" Or _
           LCase(GetExtensionName(file.Name)) = "xlsm" Or _
           LCase(GetExtensionName(file.Name)) = "xls" Then
            fileList.Add file.Path
        End If
    Next file
    
    For Each subFolder In folder.SubFolders
        GetExcelFilesRecursive subFolder, fileList
    Next subFolder
End Sub

Function GetExtensionName(fileName As String) As String
    Dim dotRevPos As Long
    dotRevPos = InStrRev(fileName, ".")
    
    If 0 < dotRevPos And dotRevPos < Len(fileName) Then
        GetExtensionName = Mid(fileName, dotRevPos + 1)
    Else
        GetExtensionName = ""
    End If
End Function

Sub ExportSheetsToText(filePath As String, inputFolder As String, outputFolder As String)
    Dim excelWorkbook As Workbook: Set excelWorkbook = Workbooks.Open(fileName:=filePath, ReadOnly:=True)
    Dim excelWorksheet As Worksheet
    Dim excelCell As Range
    Dim excelRow As Range
    Dim excelLine As String
    
    Dim fileObj As Object: Set fileObj = CreateObject("Scripting.FileSystemObject")
    Dim textFile As Object
    
    Dim relativePath As String: relativePath = Replace(filePath, inputFolder, "")
    If Left(relativePath, 1) = "\" Then relativePath = Mid(relativePath, 2)
    relativePath = fileObj.GetParentFolderName(relativePath)
    
    For Each excelWorksheet In excelWorkbook.Worksheets
        If Not fileObj.FolderExists(outputFolder & "\" & relativePath) Then
            fileObj.CreateFolder outputFolder & "\" & relativePath
        End If
        
        Set textFile = fileObj.CreateTextFile( _
            outputFolder & "\" & _
            relativePath & "\" & _
            fileObj.GetBaseName(filePath) & "_" & _
            excelWorksheet.Name & ".txt", True, False)
        
        For Each excelRow In excelWorksheet.UsedRange.Rows
            excelLine = ""
            For Each excelCell In excelRow.Cells
                excelLine = excelLine & excelCell.Text & vbTab
            Next
            textFile.WriteLine Left(excelLine, Len(excelLine) - 1)
        Next
        
        textFile.Close
    Next
    
    excelWorkbook.Close SaveChanges:=False
End Sub
