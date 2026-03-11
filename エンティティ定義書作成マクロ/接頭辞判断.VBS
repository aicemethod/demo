Option Explicit

Const xlOpenXMLWorkbook = 51

Dim targetFolderPath
Dim outputFilePath
Dim fso
Dim excelApp
Dim outputBook
Dim folder
Dim file

targetFolderPath = ""
If WScript.Arguments.Count > 0 Then
    targetFolderPath = WScript.Arguments(0)
End If

If targetFolderPath = "" Then
    MsgBox "処理したいフォルダをこのVBSファイルにドラッグしてください。", vbExclamation
    WScript.Quit 1
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(targetFolderPath) Then
    MsgBox "指定されたパスはフォルダではありません。" & vbCrLf & targetFolderPath, vbExclamation
    WScript.Quit 1
End If

outputFilePath = fso.BuildPath(targetFolderPath, "全スキーマ名取得.xlsx")

On Error Resume Next
Set excelApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Excelを起動できませんでした。", vbCritical
    WScript.Quit 1
End If
On Error GoTo 0

excelApp.Visible = False
excelApp.DisplayAlerts = False

Set outputBook = OpenOrCreateOutputBook(excelApp, outputFilePath)
If outputBook Is Nothing Then
    Cleanup excelApp, Nothing
    MsgBox "出力ファイルを準備できませんでした。" & vbCrLf & outputFilePath, vbCritical
    WScript.Quit 1
End If

Set folder = fso.GetFolder(targetFolderPath)

For Each file In folder.Files
    If IsExcelFile(file.Name) Then
        If LCase(file.Path) <> LCase(outputFilePath) Then
            ProcessWorkbook excelApp, outputBook, file.Path
        End If
    End If
Next

outputBook.Save
Cleanup excelApp, outputBook

MsgBox "処理が完了しました。" & vbCrLf & outputFilePath, vbInformation

Function OpenOrCreateOutputBook(app, filePath)
    Dim book

    Set book = Nothing

    On Error Resume Next
    If fso.FileExists(filePath) Then
        Set book = app.Workbooks.Open(filePath)
    Else
        Set book = app.Workbooks.Add
        EnsureSecondSheet book
        book.Worksheets(1).Name = "スキーマ名"
        book.Worksheets(2).Name = "失敗ファイル"
    End If
    If Not book Is Nothing Then
        EnsureSecondSheet book
        If Not fso.FileExists(filePath) Then
            book.SaveAs filePath, xlOpenXMLWorkbook
        End If
    End If
    If Err.Number <> 0 Then
        Set book = Nothing
        Err.Clear
    End If
    On Error GoTo 0

    Set OpenOrCreateOutputBook = book
End Function

Sub EnsureSecondSheet(book)
    Do While book.Worksheets.Count < 2
        book.Worksheets.Add , book.Worksheets(book.Worksheets.Count)
    Loop
End Sub

Sub ProcessWorkbook(app, resultBook, sourceFilePath)
    Dim sourceBook
    Dim sourceSheet
    Dim resultSheet
    Dim lastRow
    Dim rowIndex
    Dim nextRow
    Dim cellValue
    Dim hasError

    Set sourceBook = Nothing
    hasError = False

    On Error Resume Next
    Set sourceBook = app.Workbooks.Open(sourceFilePath, False, True)
    If Err.Number <> 0 Then
        hasError = True
        Err.Clear
    End If

    If Not hasError Then
        If sourceBook.Worksheets.Count < 2 Then
            hasError = True
        Else
            Set sourceSheet = sourceBook.Worksheets(2)
            Set resultSheet = resultBook.Worksheets(1)

            lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(-4162).Row
            If lastRow >= 2 Then
                nextRow = GetNextRow(resultSheet, 1)
                For rowIndex = 2 To lastRow
                    cellValue = Trim(CStr(sourceSheet.Cells(rowIndex, 1).Value))
                    If cellValue = "" Then
                        Exit For
                    End If

                    resultSheet.Cells(nextRow, 1).Value = cellValue
                    nextRow = nextRow + 1
                Next
            End If
            If Err.Number <> 0 Then
                hasError = True
                Err.Clear
            End If
        End If
    End If

    If Not sourceBook Is Nothing Then
        sourceBook.Close False
    End If
    On Error GoTo 0

    If hasError Then
        AppendFailedFile resultBook.Worksheets(2), fso.GetFileName(sourceFilePath)
    End If
End Sub

Sub AppendFailedFile(failedSheet, fileName)
    Dim nextRow

    nextRow = GetNextRow(failedSheet, 1)
    failedSheet.Cells(nextRow, 1).Value = fileName
End Sub

Function GetNextRow(sheet, columnNumber)
    Dim lastRow

    lastRow = sheet.Cells(sheet.Rows.Count, columnNumber).End(-4162).Row
    If Trim(CStr(sheet.Cells(lastRow, columnNumber).Value)) = "" Then
        GetNextRow = lastRow
    Else
        GetNextRow = lastRow + 1
    End If
End Function

Function IsExcelFile(fileName)
    Dim extension

    extension = LCase(fso.GetExtensionName(fileName))
    Select Case extension
        Case "xls", "xlsx", "xlsm", "xlsb"
            IsExcelFile = True
        Case Else
            IsExcelFile = False
    End Select
End Function

Sub Cleanup(app, book)
    On Error Resume Next
    If Not book Is Nothing Then
        book.Close True
    End If
    If Not app Is Nothing Then
        app.Quit
    End If
    On Error GoTo 0
End Sub
