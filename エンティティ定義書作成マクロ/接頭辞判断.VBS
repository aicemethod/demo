Option Explicit

Const xlOpenXMLWorkbook = 51
Const xlCalculationManual = -4135
Const xlCalculationAutomatic = -4105

Dim targetFolderPath
Dim outputFilePath
Dim progressFilePath
Dim progressHtaPath
Dim fso
Dim excelApp
Dim outputBook
Dim folder
Dim file
Dim totalCount
Dim completedCount
Dim failedCount
Dim statusMessage
Dim shell

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
progressFilePath = fso.BuildPath(targetFolderPath, "全スキーマ名取得_進捗.txt")
progressHtaPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "進捗ダイアログ.hta")
Set shell = CreateObject("WScript.Shell")

On Error Resume Next
Set excelApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Excelを起動できませんでした。", vbCritical
    WScript.Quit 1
End If
On Error GoTo 0

excelApp.Visible = False
excelApp.DisplayAlerts = False
excelApp.ScreenUpdating = False
excelApp.EnableEvents = False
excelApp.AskToUpdateLinks = False
excelApp.DisplayStatusBar = True
excelApp.Calculation = xlCalculationManual

Set outputBook = OpenOrCreateOutputBook(excelApp, outputFilePath)
If outputBook Is Nothing Then
    Cleanup excelApp, Nothing
    MsgBox "出力ファイルを準備できませんでした。" & vbCrLf & outputFilePath, vbCritical
    WScript.Quit 1
End If

Set folder = fso.GetFolder(targetFolderPath)
totalCount = CountExcelFiles(folder, outputFilePath)
completedCount = 0
failedCount = 0

statusMessage = "開始しました。対象ファイル数: " & totalCount
UpdateProgress progressFilePath, "RUNNING", statusMessage
LaunchProgressDialog progressHtaPath, progressFilePath

For Each file In folder.Files
    If IsExcelFile(file.Name) Then
        If LCase(file.Path) <> LCase(outputFilePath) Then
            statusMessage = "処理中: " & (completedCount + 1) & "/" & totalCount & vbCrLf & fso.GetFileName(file.Path)
            excelApp.StatusBar = statusMessage
            If ProcessWorkbook(excelApp, outputBook, file.Path) Then
                failedCount = failedCount + 1
            End If
            completedCount = completedCount + 1
            statusMessage = "完了: " & completedCount & "/" & totalCount & " 失敗: " & failedCount
            UpdateProgress progressFilePath, "RUNNING", statusMessage
        End If
    End If
Next

outputBook.Save
statusMessage = "処理完了" & vbCrLf & _
                "完了: " & completedCount & "/" & totalCount & vbCrLf & _
                "失敗: " & failedCount
UpdateProgress progressFilePath, "DONE", statusMessage
Cleanup excelApp, outputBook

MsgBox "処理が完了しました。" & vbCrLf & _
       "成功: " & (completedCount - failedCount) & "件" & vbCrLf & _
       "失敗: " & failedCount & "件" & vbCrLf & _
       outputFilePath, vbInformation

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

Function ProcessWorkbook(app, resultBook, sourceFilePath)
    Dim sourceBook
    Dim sourceSheet
    Dim resultSheet
    Dim lastRow
    Dim nextRow
    Dim hasError
    Dim values
    Dim writeValues
    Dim itemCount

    Set sourceBook = Nothing
    hasError = False
    ProcessWorkbook = False

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

            itemCount = GetSchemaCount(sourceSheet)
            If itemCount > 0 Then
                values = sourceSheet.Range(sourceSheet.Cells(2, 1), sourceSheet.Cells(itemCount + 1, 1)).Value
                writeValues = NormalizeValues(values, itemCount)
                nextRow = GetNextRow(resultSheet, 1)
                resultSheet.Range(resultSheet.Cells(nextRow, 1), resultSheet.Cells(nextRow + itemCount - 1, 1)).Value = writeValues
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
        ProcessWorkbook = True
    End If
End Function

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

Function GetSchemaCount(sheet)
    Dim rowIndex
    Dim lastRow

    GetSchemaCount = 0
    lastRow = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
    If lastRow < 2 Then
        Exit Function
    End If

    For rowIndex = 2 To lastRow
        If Trim(CStr(sheet.Cells(rowIndex, 1).Value)) = "" Then
            Exit For
        End If
        GetSchemaCount = GetSchemaCount + 1
    Next
End Function

Function NormalizeValues(values, itemCount)
    Dim normalized()
    Dim i

    ReDim normalized(1 To itemCount, 1 To 1)

    If itemCount = 1 Then
        normalized(1, 1) = values
    Else
        For i = 1 To itemCount
            normalized(i, 1) = values(i, 1)
        Next
    End If

    NormalizeValues = normalized
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

Sub LaunchProgressDialog(htaPath, filePath)
    Dim command

    If Not fso.FileExists(htaPath) Then
        Exit Sub
    End If

    command = "mshta.exe " & Quote(htaPath & "?" & filePath)
    shell.Run command, 1, False
End Sub

Function Quote(text)
    Quote = Chr(34) & text & Chr(34)
End Function

Function CountExcelFiles(targetFolder, resultPath)
    Dim targetFile
    Dim count

    count = 0
    For Each targetFile In targetFolder.Files
        If IsExcelFile(targetFile.Name) Then
            If LCase(targetFile.Path) <> LCase(resultPath) Then
                count = count + 1
            End If
        End If
    Next

    CountExcelFiles = count
End Function

Sub UpdateProgress(filePath, status, message)
    Dim stream

    On Error Resume Next
    Set stream = fso.OpenTextFile(filePath, 2, True)
    If Err.Number = 0 Then
        stream.WriteLine status
        stream.WriteLine Now
        stream.WriteLine message
        stream.Close
    Else
        Err.Clear
    End If
    Set stream = Nothing
    On Error GoTo 0
End Sub

Sub Cleanup(app, book)
    On Error Resume Next
    If Not book Is Nothing Then
        book.Close True
    End If
    If Not app Is Nothing Then
        app.StatusBar = False
        app.Calculation = xlCalculationAutomatic
        app.EnableEvents = True
        app.ScreenUpdating = True
        app.Quit
    End If
    On Error GoTo 0
End Sub
