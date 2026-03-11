Option Explicit

Dim gFso
Dim gExcel

Main

Sub Main()
    Dim folderPath, message

    Set gFso = CreateObject("Scripting.FileSystemObject")

    If WScript.Arguments.Count = 0 Then
        MsgBox "フォルダをこのVBSにドラッグしてください。", vbExclamation, "確認"
        Exit Sub
    End If

    folderPath = WScript.Arguments(0)
    If Not gFso.FolderExists(folderPath) Then
        MsgBox "指定されたパスはフォルダではありません。" & vbCrLf & folderPath, vbCritical, "エラー"
        Exit Sub
    End If

    If Not StartExcel(message) Then
        Cleanup
        MsgBox message, vbCritical, "エラー"
        Exit Sub
    End If

    message = ProcessFolderRecursive(gFso.GetFolder(folderPath))

    Cleanup
    MsgBox message, vbInformation, "完了"
End Sub

Function ProcessFolderRecursive(targetFolder)
    Dim file
    Dim subFolder
    Dim fileCount, renameCount, cellUpdateCount, errorCount
    Dim details

    fileCount = 0
    renameCount = 0
    cellUpdateCount = 0
    errorCount = 0
    details = ""

    TraverseFolder targetFolder, fileCount, renameCount, cellUpdateCount, errorCount, details

    ProcessFolderRecursive = "処理が完了しました。" & vbCrLf & _
                             "対象ファイル: " & fileCount & "件" & vbCrLf & _
                             "ファイル名変更: " & renameCount & "件" & vbCrLf & _
                             "表紙B7変更: " & cellUpdateCount & "件" & vbCrLf & _
                             "エラー: " & errorCount & "件" & details
End Function

Sub TraverseFolder(targetFolder, ByRef fileCount, ByRef renameCount, ByRef cellUpdateCount, ByRef errorCount, ByRef details)
    Dim file
    Dim subFolder
    Dim currentPath
    Dim renameTargetPath
    Dim updatedCount
    Dim errMsg

    For Each file In targetFolder.Files
        fileCount = fileCount + 1
        currentPath = file.Path

        If IsExcelFile(currentPath) Then
            updatedCount = 0
            If UpdateCoverB7(currentPath, updatedCount, errMsg) Then
                cellUpdateCount = cellUpdateCount + updatedCount
            Else
                errorCount = errorCount + 1
                AppendDetail details, errMsg
            End If
        End If

        renameTargetPath = ""
        If RenameFileWithoutId(currentPath, renameTargetPath, errMsg) Then
            If renameTargetPath <> "" Then
                renameCount = renameCount + 1
            End If
        Else
            errorCount = errorCount + 1
            AppendDetail details, errMsg
        End If
    Next

    For Each subFolder In targetFolder.SubFolders
        TraverseFolder subFolder, fileCount, renameCount, cellUpdateCount, errorCount, details
    Next
End Sub

Function UpdateCoverB7(filePath, ByRef updatedCount, ByRef errMsg)
    Dim wb
    Dim ws
    Dim currentValue
    Dim newValue
    Dim closeErrMsg
    Dim hasError

    UpdateCoverB7 = False
    updatedCount = 0
    errMsg = ""
    Set wb = Nothing
    Set ws = Nothing
    closeErrMsg = ""
    hasError = False

    On Error Resume Next
    Set wb = gExcel.Workbooks.Open(filePath, 0, False)
    If Err.Number <> 0 Then
        errMsg = "Excelを開けません: " & filePath & " / " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Set ws = wb.Worksheets("表紙")
    If Err.Number <> 0 Then
        errMsg = "シート「表紙」がありません: " & filePath
        Err.Clear
        hasError = True
    Else
        currentValue = CStr(ws.Range("B7").Value)
        newValue = Replace(currentValue, "ID_", "")

        If newValue <> currentValue Then
            ws.Range("B7").Value = newValue
            wb.Save
            If Err.Number <> 0 Then
                errMsg = "Excel保存に失敗しました: " & filePath & " / " & Err.Description
                Err.Clear
                hasError = True
            Else
                updatedCount = 1
            End If
        End If
    End If

    If Not wb Is Nothing Then
        wb.Close False
        If Err.Number <> 0 Then
            closeErrMsg = "Excelクローズに失敗しました: " & filePath & " / " & Err.Description
            Err.Clear
            If errMsg = "" Then
                errMsg = closeErrMsg
            Else
                errMsg = errMsg & " / " & closeErrMsg
            End If
            hasError = True
        End If
    End If
    On Error GoTo 0

    UpdateCoverB7 = (Not hasError)
End Function

Function RenameFileWithoutId(filePath, ByRef renamedPath, ByRef errMsg)
    Dim folderPath
    Dim fileName
    Dim newFileName
    Dim targetPath

    RenameFileWithoutId = False
    renamedPath = ""
    errMsg = ""

    folderPath = gFso.GetParentFolderName(filePath)
    fileName = gFso.GetFileName(filePath)
    newFileName = Replace(fileName, "ID_", "")

    If newFileName = fileName Then
        RenameFileWithoutId = True
        Exit Function
    End If

    targetPath = gFso.BuildPath(folderPath, newFileName)
    If gFso.FileExists(targetPath) Then
        errMsg = "同名ファイルが存在するためリネームできません: " & targetPath
        Exit Function
    End If

    On Error Resume Next
    gFso.MoveFile filePath, targetPath
    If Err.Number <> 0 Then
        errMsg = "ファイル名変更に失敗しました: " & filePath & " / " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    renamedPath = targetPath
    RenameFileWithoutId = True
End Function

Function IsExcelFile(filePath)
    Dim ext

    ext = LCase(gFso.GetExtensionName(filePath))
    Select Case ext
        Case "xlsx", "xlsm", "xls", "xlsb"
            IsExcelFile = True
        Case Else
            IsExcelFile = False
    End Select
End Function

Function StartExcel(ByRef errMsg)
    StartExcel = False
    errMsg = ""

    On Error Resume Next
    Set gExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        errMsg = "Excelを起動できません: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    gExcel.Visible = False
    gExcel.DisplayAlerts = False
    StartExcel = True
End Function

Sub Cleanup()
    On Error Resume Next
    If Not gExcel Is Nothing Then
        gExcel.DisplayAlerts = True
        gExcel.Quit
    End If
    Set gExcel = Nothing
    Set gFso = Nothing
    On Error GoTo 0
End Sub

Sub AppendDetail(ByRef details, ByVal message)
    details = details & vbCrLf & " - " & message
End Sub
