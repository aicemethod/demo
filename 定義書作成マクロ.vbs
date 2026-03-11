Option Explicit

Const xlUp = -4162
Const xlToLeft = -4159
Const xlCalculationManual = -4135
Const xlCalculationAutomatic = -4105
Const xlOpenXMLWorkbook = 51
Const COLOR_BLACK = 0
Const COLOR_GRAY = 12632256

Dim gFso
Dim gExcel
Dim gTemplateWb
Dim gMemoMap

Main

Sub Main()
    Dim folderPath, templatePath, outputFolderPath, memoPath
    Dim files, message

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

    templatePath = FindTemplatePath(folderPath)
    If templatePath = "" Then
        MsgBox "同じ階層に「テンプレート.xlsx」が見つかりません。", vbCritical, "エラー"
        Exit Sub
    End If

    outputFolderPath = gFso.BuildPath(gFso.GetParentFolderName(folderPath), "20_作成済定義書")
    EnsureFolder outputFolderPath
    memoPath = FindMemoPath(folderPath)

    files = CollectExcelFiles(folderPath, templatePath)
    If IsEmpty(files) Then
        MsgBox "処理対象のExcelファイルがありません。", vbInformation, "完了"
        Exit Sub
    End If

    If Not StartExcel(templatePath, message) Then
        Cleanup
        MsgBox message, vbCritical, "エラー"
        Exit Sub
    End If

    If Not LoadMemoMap(memoPath, message) Then
        Cleanup
        MsgBox message, vbCritical, "エラー"
        Exit Sub
    End If

    message = ProcessAllFiles(files, outputFolderPath)
    Cleanup
    MsgBox message, vbInformation, "完了"
End Sub

Function ProcessAllFiles(files, outputFolderPath)
    Dim i, okCount, ngCount, detail

    okCount = 0
    ngCount = 0
    detail = ""

    For i = 0 To UBound(files)
        If ProcessOneFile(CStr(files(i)), outputFolderPath, detail) Then
            okCount = okCount + 1
        Else
            ngCount = ngCount + 1
        End If
    Next

    ProcessAllFiles = "処理が完了しました。" & vbCrLf & _
                      "成功: " & okCount & "件" & vbCrLf & _
                      "失敗: " & ngCount & "件" & detail
End Function

Function ProcessOneFile(filePath, outputFolderPath, ByRef detail)
    Dim srcWb, outWb
    Dim ws1, ws2, wsTable, wsField, wsCover
    Dim values1, fieldInfo, englishValues, displayName, outputPath
    Dim errMsg

    ProcessOneFile = False
    Set srcWb = Nothing
    Set outWb = Nothing

    On Error Resume Next
    Set srcWb = gExcel.Workbooks.Open(filePath, 0, True)
    If Err.Number <> 0 Then
        errMsg = "入力ファイルを開けません: " & gFso.GetFileName(filePath) & " / " & Err.Description
        Err.Clear
        On Error GoTo 0
        AppendDetail detail, errMsg
        Exit Function
    End If
    On Error GoTo 0

    If srcWb.Sheets.Count < 2 Then
        AppendDetail detail, "シート数不足: " & gFso.GetFileName(filePath)
        CloseWorkbookSafe srcWb, False
        Exit Function
    End If

    Set ws1 = srcWb.Sheets(1)
    Set ws2 = srcWb.Sheets(2)

    values1 = GetSheet1Values(ws1)
    displayName = CStr(values1(0))
    If Trim(displayName) = "" Then
        AppendDetail detail, "AJ4(Display Name)が空です: " & gFso.GetFileName(filePath)
        CloseWorkbookSafe srcWb, False
        Exit Function
    End If

    fieldInfo = ReadFieldSheet(ws2)
    If IsEmpty(fieldInfo) Then
        AppendDetail detail, "シート2の読み込みに失敗しました: " & gFso.GetFileName(filePath)
        CloseWorkbookSafe srcWb, False
        Exit Function
    End If

    englishValues = ReadEnglishSheetValues(filePath)

    On Error Resume Next
    gTemplateWb.Worksheets.Copy
    If Err.Number <> 0 Then
        errMsg = "テンプレート複製に失敗しました: " & gFso.GetFileName(filePath) & " / " & Err.Description
        Err.Clear
        On Error GoTo 0
        CloseWorkbookSafe srcWb, False
        AppendDetail detail, errMsg
        Exit Function
    End If
    Set outWb = gExcel.ActiveWorkbook
    On Error GoTo 0

    Set wsTable = Nothing
    Set wsField = Nothing
    Set wsCover = Nothing
    On Error Resume Next
    Set wsTable = outWb.Worksheets("テーブル")
    Set wsField = outWb.Worksheets("フィールド")
    Set wsCover = outWb.Worksheets("表紙")
    If Err.Number <> 0 Or wsTable Is Nothing Or wsField Is Nothing Or wsCover Is Nothing Then
        errMsg = "テンプレート内の必要シートが不足しています: " & gFso.GetFileName(filePath)
        Err.Clear
        On Error GoTo 0
        CloseWorkbookSafe outWb, False
        CloseWorkbookSafe srcWb, False
        AppendDetail detail, errMsg
        Exit Function
    End If
    On Error GoTo 0

    FillTableSheet wsTable, values1, fieldInfo
    FillFieldSheet wsField, fieldInfo, values1, englishValues
    wsCover.Range("B7").Value = "エンティティ定義書_ID_" & displayName & "_v0.2"

    outputPath = gFso.BuildPath(outputFolderPath, SanitizeFileName("エンティティ定義書_ID_" & displayName & "_v0.2.xlsx"))
    If SaveWorkbookAs(outWb, outputPath, errMsg) Then
        ProcessOneFile = True
    Else
        AppendDetail detail, errMsg
    End If

    CloseWorkbookSafe outWb, False
    CloseWorkbookSafe srcWb, False
End Function

Function GetSheet1Values(ws)
    Dim arr(35)

    arr(0) = Nz(ws.Range("AJ4").Value2)
    arr(1) = Nz(ws.Range("AI4").Value2)
    arr(2) = Nz(ws.Range("AH4").Value2)
    arr(3) = Nz(ws.Range("AW4").Value2)
    arr(4) = Nz(ws.Range("D4").Value2)
    arr(5) = Nz(ws.Range("E4").Value2)
    arr(6) = Nz(ws.Range("DU4").Value2)
    arr(7) = Nz(ws.Range("I4").Value2)
    arr(8) = Nz(ws.Range("BA4").Value2)
    arr(9) = Nz(ws.Range("AL4").Value2)
    arr(10) = Nz(ws.Range("DK4").Value2)
    arr(11) = Nz(ws.Range("DM4").Value2)
    arr(12) = Nz(ws.Range("BQ4").Value2)
    arr(13) = Nz(ws.Range("AA4").Value2)
    arr(14) = Nz(ws.Range("AN4").Value2)
    arr(15) = Nz(ws.Range("AM4").Value2)
    arr(16) = Nz(ws.Range("BG4").Value2)
    arr(17) = Nz(ws.Range("CH4").Value2)
    arr(18) = Nz(ws.Range("CL4").Value2)
    arr(19) = Nz(ws.Range("AS4").Value2)
    arr(20) = Nz(ws.Range("BZ4").Value2)
    arr(21) = Nz(ws.Range("BO4").Value2)
    arr(22) = Nz(ws.Range("BL4").Value2)
    arr(23) = Nz(ws.Range("AU4").Value2)
    arr(24) = Nz(ws.Range("L4").Value2)
    arr(25) = Nz(ws.Range("AV4").Value2)
    arr(26) = Nz(ws.Range("DT4").Value2)
    arr(27) = Nz(ws.Range("CD4").Value2)
    arr(28) = Nz(ws.Range("CS4").Value2)
    arr(29) = Nz(ws.Range("DJ4").Value2)

    GetSheet1Values = arr
End Function

Sub FillTableSheet(wsTable, values1, fieldInfo)
    Dim rowNo, tableArr
    Dim dmKey, foundData

    tableArr = wsTable.Range("E5:E41").Value

    tableArr(1, 1) = values1(0)
    tableArr(2, 1) = values1(1)
    tableArr(3, 1) = values1(2)
    tableArr(4, 1) = ConvertBooleanForTable(values1(3))
    tableArr(5, 1) = ConvertPrefixName(values1(4))
    tableArr(6, 1) = ConvertPrefixName(values1(5))
    tableArr(7, 1) = ConvertTableType(values1(6))
    tableArr(8, 1) = ConvertOwnershipType(values1(7))
    tableArr(9, 1) = DefaultText(values1(8), "なし")
    tableArr(10, 1) = values1(9)
    tableArr(11, 1) = ConvertPrefixName(values1(10))

    dmKey = values1(11)
    foundData = FindFieldRow(fieldInfo, dmKey)

    If IsArray(foundData) Then
        tableArr(12, 1) = ConvertBooleanForTable(foundData(2))
        tableArr(13, 1) = ConvertBooleanForTable(foundData(4))
        tableArr(14, 1) = ConvertPrefixName(foundData(1))
        tableArr(15, 1) = ConvertPrefixName(dmKey)
        tableArr(16, 1) = ConvertFieldRequirement(foundData(7))
        tableArr(17, 1) = ExtractTextMaxLength(foundData(12))
    Else
        tableArr(15, 1) = ConvertPrefixName(dmKey)
    End If

    tableArr(18, 1) = ConvertBooleanForTable(values1(12))
    tableArr(19, 1) = ConvertBooleanForTable(values1(13))
    tableArr(20, 1) = ConvertBooleanForTable(values1(14))
    tableArr(21, 1) = values1(15)
    tableArr(22, 1) = ConvertBooleanForTable(values1(16))
    tableArr(23, 1) = ConvertBooleanForTable(values1(17))
    tableArr(24, 1) = ConvertBooleanForTable(values1(18))
    tableArr(25, 1) = "-"
    tableArr(26, 1) = ConvertBooleanForTable(values1(19))
    tableArr(27, 1) = ConvertBooleanForTable(values1(20))
    tableArr(28, 1) = "-"
    tableArr(29, 1) = ConvertBooleanForTable(values1(21))
    tableArr(30, 1) = ConvertBooleanForTable(values1(22))
    tableArr(31, 1) = ConvertBooleanForTable(values1(23))
    tableArr(32, 1) = ConvertBooleanForTable(values1(24))
    tableArr(33, 1) = ConvertBooleanForTable(values1(25))
    tableArr(34, 1) = ConvertBooleanForTable(values1(26))
    tableArr(35, 1) = ConvertBooleanForTable(values1(27))
    tableArr(36, 1) = ConvertBooleanForTable(values1(28))
    tableArr(37, 1) = "-"

    wsTable.Range("E5:E41").Value = tableArr
    wsTable.Range("E5:E41").Font.Color = COLOR_BLACK
End Sub

Sub FillFieldSheet(wsField, fieldInfo, values1, englishValues)
    Dim dataArr, headerMap, rowCount
    Dim outArr, r, djKey, dmKey

    dataArr = fieldInfo(0)
    Set headerMap = fieldInfo(1)
    rowCount = UBound(dataArr, 1) - 1
    If rowCount <= 0 Then Exit Sub

    djKey = Nz(values1(29))
    dmKey = Nz(values1(11))

    outArr = wsField.Range("C7:AJ" & (6 + rowCount)).Value

    For r = 2 To UBound(dataArr, 1)
        FillFieldRow outArr, r - 1, dataArr, r, headerMap, djKey, dmKey
        If IsArray(englishValues) Then
            If (r - 2) <= UBound(englishValues) Then
                outArr(r - 1, 3) = englishValues(r - 2)
            End If
        End If
    Next

    SortFieldArrayByCustomAttribute outArr

    wsField.Range("C7:AJ" & (6 + rowCount)).Value = outArr
    wsField.Range("C7:AJ" & (6 + rowCount)).Font.Color = COLOR_BLACK

    wsField.Range("AG5").Value = "Schema Name"
    wsField.Range("AG7:AG" & (6 + rowCount)).Value = wsField.Range("AI7:AI" & (6 + rowCount)).Value
    wsField.Range("AI7:AI" & (6 + rowCount)).ClearContents
    wsField.Range("AL5").Value = "TempTarget"
    wsField.Range("AL7:AL" & (6 + rowCount)).Value = wsField.Range("AJ7:AJ" & (6 + rowCount)).Value
    wsField.Range("AJ7:AJ" & (6 + rowCount)).ClearContents

    ApplyMemoMapping wsField, rowCount
    wsField.Range("AH7:AH" & (6 + rowCount)).ClearContents
    wsField.Columns(38).Delete
    TrimEmptyFieldRows wsField, rowCount
End Sub

Sub SortFieldArrayByCustomAttribute(ByRef outArr)
    Dim i, j, maxRow, maxCol
    Dim rankI, rankJ, colNo, tempValue

    maxRow = UBound(outArr, 1)
    maxCol = UBound(outArr, 2)

    For i = 1 To maxRow - 1
        For j = i + 1 To maxRow
            rankI = GetCustomSortRank(outArr(i, 4))
            rankJ = GetCustomSortRank(outArr(j, 4))

            If rankJ < rankI Then
                For colNo = 1 To maxCol
                    tempValue = outArr(i, colNo)
                    outArr(i, colNo) = outArr(j, colNo)
                    outArr(j, colNo) = tempValue
                Next
            End If
        Next
    Next
End Sub

Function GetCustomSortRank(value)
    Dim textValue

    textValue = Trim(CStr(value))
    If textValue = "カスタム" Then
        GetCustomSortRank = 0
    ElseIf textValue = "標準" Then
        GetCustomSortRank = 1
    Else
        GetCustomSortRank = 2
    End If
End Function

Sub FillFieldRow(ByRef outArr, ByVal outRow, dataArr, ByVal srcRow, headerMap, ByVal djKey, ByVal dmKey)
    Dim schemaName, displayName, customAttr, attrType, typeValue, requiredLevel
    Dim auditEnabled, secured, advFind, description, additionalData
    Dim info, targetText, rawTargetText, defaultText, rowKey

    schemaName = GetCellByHeader(dataArr, srcRow, headerMap, "Schema Name")
    displayName = GetCellByHeader(dataArr, srcRow, headerMap, "Display Name")
    customAttr = ConvertCustomAttribute(GetCellByHeader(dataArr, srcRow, headerMap, "Custom Attribute"))
    attrType = GetCellByHeader(dataArr, srcRow, headerMap, "Attribute Type")
    typeValue = ConvertTypeLabel(GetCellByHeader(dataArr, srcRow, headerMap, "Type"))
    requiredLevel = ConvertRequiredLevel(GetCellByHeader(dataArr, srcRow, headerMap, "Required Level"))
    auditEnabled = ConvertBooleanForField(GetCellByHeader(dataArr, srcRow, headerMap, "Audit Enabled"))
    secured = ConvertBooleanForField(GetCellByHeader(dataArr, srcRow, headerMap, "Secured"))
    advFind = ConvertBooleanForField(GetCellByHeader(dataArr, srcRow, headerMap, "ValidFor AdvancedFind"))
    description = GetCellByHeader(dataArr, srcRow, headerMap, "Description")
    additionalData = CleanAdditionalData(GetCellByHeader(dataArr, srcRow, headerMap, "Additional data"))
    rowKey = Nz(dataArr(srcRow, 1))

    info = BuildAttributeTypeInfo(attrType, additionalData)
    rawTargetText = ExtractValueFromAdditionalData(additionalData, "Target:")
    If rawTargetText = "" Then
        rawTargetText = ExtractValueFromAdditionalData(additionalData, "targets:")
    End If
    targetText = ConvertPrefixList(rawTargetText)

    defaultText = info(4)
    If LCase(defaultText) = "n/a" Then
        defaultText = "なし"
    End If

    outArr(outRow, 1) = ConvertPrefixName(schemaName)
    outArr(outRow, 2) = displayName
    outArr(outRow, 4) = customAttr
    outArr(outRow, 7) = info(0)
    outArr(outRow, 8) = typeValue
    outArr(outRow, 9) = requiredLevel
    outArr(outRow, 12) = info(1)
    outArr(outRow, 13) = info(2)
    outArr(outRow, 17) = info(3)
    outArr(outRow, 18) = defaultText
    outArr(outRow, 19) = targetText
    outArr(outRow, 21) = auditEnabled
    outArr(outRow, 22) = secured
    outArr(outRow, 25) = advFind
    outArr(outRow, 26) = description
    outArr(outRow, 33) = schemaName
    outArr(outRow, 34) = rawTargetText

    If LCase(Trim(rowKey)) = LCase(Trim(djKey)) Then
        outArr(outRow, 5) = "○"
        outArr(outRow, 6) = "○"
    End If

    If LCase(Trim(rowKey)) = LCase(Trim(dmKey)) Then
        outArr(outRow, 6) = "○"
    End If

    If LCase(Trim(attrType)) = "text" Or LCase(Trim(attrType)) = "multiline text" Then
        outArr(outRow, 32) = additionalData
        outArr(outRow, 10) = ExtractMaxLengthValue(additionalData)
    End If
End Sub

Function BuildAttributeTypeInfo(attrType, additionalData)
    Dim info(5)
    Dim key, formatLabel

    key = LCase(Trim(CStr(attrType)))
    formatLabel = GetDateFormatLabel(additionalData)

    Select Case key
        Case "bigint"
            info(0) = "数値 - 整数(Int)"
            info(1) = ExtractValueFromAdditionalData(additionalData, "Maximum value:")
            info(2) = ExtractValueFromAdditionalData(additionalData, "Minimum value:")
        Case "choice"
            info(0) = "選択肢"
            info(3) = ExtractValueFromAdditionalData(additionalData, "Options:")
            info(4) = ExtractValueFromAdditionalData(additionalData, "Default:")
        Case "choices"
            info(0) = "選択肢(複数)"
            info(3) = ExtractValueFromAdditionalData(additionalData, "Options:")
            info(4) = ExtractValueFromAdditionalData(additionalData, "Default:")
        Case "currency"
            info(0) = "通貨"
            info(1) = ExtractValueFromAdditionalData(additionalData, "Maximum value:")
            info(2) = ExtractValueFromAdditionalData(additionalData, "Minimum value:")
        Case "decimal"
            info(0) = "数値 - 少数(10進数)"
            info(1) = ExtractValueFromAdditionalData(additionalData, "Maximum value:")
            info(2) = ExtractValueFromAdditionalData(additionalData, "Minimum value:")
        Case "double"
            info(0) = "数値 - 浮動小数点数"
            info(1) = ExtractValueFromAdditionalData(additionalData, "Maximum value:")
            info(2) = ExtractValueFromAdditionalData(additionalData, "Minimum value:")
        Case "multiline text"
            info(0) = "複数行テキスト - プレーン"
        Case "owner"
            info(0) = "所有者"
        Case "state"
            info(0) = "状態"
            info(3) = ExtractValueFromAdditionalData(additionalData, "States:")
        Case "status"
            info(0) = "ステータス"
            info(3) = ExtractValueFromAdditionalData(additionalData, "States:")
        Case "text"
            info(0) = "1行テキスト - プレーン"
        Case "two options"
            info(0) = "はい/いいえ"
            info(3) = ExtractValueFromAdditionalData(additionalData, "Options:")
            info(4) = ExtractValueFromAdditionalData(additionalData, "Default Value:")
            If info(4) = "" Then info(4) = ExtractValueFromAdditionalData(additionalData, "Default:")
        Case "uniqueidentifier"
            info(0) = "一意識別子"
        Case "whole number"
            info(0) = "数値 - 整数(Int)"
            info(1) = ExtractValueFromAdditionalData(additionalData, "Maximum value:")
            info(2) = ExtractValueFromAdditionalData(additionalData, "Minimum value:")
        Case "lookup"
            info(0) = "検索"
        Case "datetime", "dateandtime"
            If formatLabel <> "" Then
                info(0) = formatLabel
            Else
                info(0) = "日付と時刻"
            End If
        Case Else
            info(0) = attrType
    End Select

    info(5) = ConvertPrefixList(ExtractValueFromAdditionalData(additionalData, "targets:"))
    If info(5) = "" Then
        info(5) = ConvertPrefixList(ExtractValueFromAdditionalData(additionalData, "Target:"))
    End If

    BuildAttributeTypeInfo = info
End Function

Function ReadFieldSheet(ws)
    Dim lastRow, lastCol, dataArr, headerMap, col, headerText
    Dim result(1)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Or lastCol < 1 Then Exit Function

    dataArr = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2

    Set headerMap = CreateObject("Scripting.Dictionary")
    For col = 1 To lastCol
        headerText = LCase(Trim(CStr(dataArr(1, col))))
        If headerText <> "" Then
            headerMap(headerText) = col
        End If
    Next

    result(0) = dataArr
    Set result(1) = headerMap
    ReadFieldSheet = result
End Function

Function ReadEnglishSheetValues(sourceFilePath)
    Dim englishPath, englishWb, englishWs
    Dim lastRow, arr2d, values()
    Dim i, rowCount

    englishPath = FindEnglishFilePath(sourceFilePath)
    If englishPath = "" Then Exit Function

    Set englishWb = Nothing
    Set englishWs = Nothing

    On Error Resume Next
    Set englishWb = gExcel.Workbooks.Open(englishPath, 0, True)
    If Err.Number <> 0 Or englishWb Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If englishWb.Sheets.Count < 2 Then
        On Error GoTo 0
        CloseWorkbookSafe englishWb, False
        Exit Function
    End If

    Set englishWs = englishWb.Sheets(2)
    lastRow = englishWs.Cells(englishWs.Rows.Count, 3).End(xlUp).Row

    If lastRow < 2 Then
        On Error GoTo 0
        CloseWorkbookSafe englishWb, False
        Exit Function
    End If

    arr2d = englishWs.Range("C2:C" & lastRow).Value2
    rowCount = UBound(arr2d, 1)
    ReDim values(rowCount - 1)

    For i = 1 To rowCount
        values(i - 1) = Nz(arr2d(i, 1))
    Next

    On Error GoTo 0
    CloseWorkbookSafe englishWb, False
    ReadEnglishSheetValues = values
End Function

Function FindEnglishFilePath(sourceFilePath)
    Dim sourceFolderPath, baseParentPath, englishRootPath, englishPath

    sourceFolderPath = gFso.GetParentFolderName(sourceFilePath)
    baseParentPath = gFso.GetParentFolderName(sourceFolderPath)
    englishRootPath = gFso.BuildPath(baseParentPath, "30_英語ファイル")
    englishPath = gFso.BuildPath(englishRootPath, gFso.GetFileName(sourceFilePath))

    If gFso.FileExists(englishPath) Then
        FindEnglishFilePath = englishPath
    Else
        FindEnglishFilePath = ""
    End If
End Function

Function FindFieldRow(fieldInfo, logicalName)
    Dim dataArr, rowCount, r, value

    If Trim(CStr(logicalName)) = "" Then Exit Function

    dataArr = fieldInfo(0)
    rowCount = UBound(dataArr, 1)

    For r = 2 To rowCount
        value = Trim(CStr(dataArr(r, 1)))
        If LCase(value) = LCase(Trim(CStr(logicalName))) Then
            FindFieldRow = Array( _
                GetArrayCell(dataArr, r, 1), GetArrayCell(dataArr, r, 2), GetArrayCell(dataArr, r, 3), GetArrayCell(dataArr, r, 4), _
                GetArrayCell(dataArr, r, 5), GetArrayCell(dataArr, r, 6), GetArrayCell(dataArr, r, 7), GetArrayCell(dataArr, r, 8), _
                GetArrayCell(dataArr, r, 9), GetArrayCell(dataArr, r, 10), GetArrayCell(dataArr, r, 11), GetArrayCell(dataArr, r, 12), GetArrayCell(dataArr, r, 13) _
            )
            Exit Function
        End If
    Next
End Function

Function GetArrayCell(dataArr, rowNo, colNo)
    If colNo <= UBound(dataArr, 2) Then
        GetArrayCell = Nz(dataArr(rowNo, colNo))
    Else
        GetArrayCell = ""
    End If
End Function

Function GetCellByHeader(dataArr, rowNo, headerMap, headerName)
    Dim key

    key = LCase(headerName)
    If headerMap.Exists(key) Then
        GetCellByHeader = Nz(dataArr(rowNo, headerMap(key)))
    Else
        GetCellByHeader = ""
    End If
End Function

Function ConvertBooleanForTable(value)
    Dim lowerVal

    lowerVal = LCase(Trim(CStr(value)))
    If lowerVal = "true" Then
        ConvertBooleanForTable = ChrW(10003)
    ElseIf lowerVal = "false" Or lowerVal = "" Then
        ConvertBooleanForTable = "-"
    Else
        ConvertBooleanForTable = value
    End If
End Function

Function ConvertBooleanForField(value)
    Dim lowerVal

    lowerVal = LCase(Trim(CStr(value)))
    If lowerVal = "true" Then
        ConvertBooleanForField = "TRUE"
    ElseIf lowerVal = "false" Then
        ConvertBooleanForField = "FALSE"
    Else
        ConvertBooleanForField = value
    End If
End Function

Function ConvertCustomAttribute(value)
    Dim lowerVal

    lowerVal = LCase(Trim(CStr(value)))
    If lowerVal = "true" Then
        ConvertCustomAttribute = "カスタム"
    ElseIf lowerVal = "false" Then
        ConvertCustomAttribute = "標準"
    Else
        ConvertCustomAttribute = value
    End If
End Function

Function ConvertTypeLabel(value)
    Dim lowerVal

    lowerVal = LCase(Trim(CStr(value)))
    Select Case lowerVal
        Case "simple"
            ConvertTypeLabel = "シンプル"
        Case "calculated"
            ConvertTypeLabel = "計算"
        Case "rollup"
            ConvertTypeLabel = "ロールアップ"
        Case Else
            ConvertTypeLabel = value
    End Select
End Function

Function ConvertRequiredLevel(value)
    Dim lowerVal

    lowerVal = LCase(Trim(CStr(value)))
    Select Case lowerVal
        Case "none"
            ConvertRequiredLevel = "任意"
        Case "applicationrequired"
            ConvertRequiredLevel = "システム要求"
        Case "systemrequired"
            ConvertRequiredLevel = "必須項目"
        Case "recommended"
            ConvertRequiredLevel = "推奨項目"
        Case Else
            ConvertRequiredLevel = value
    End Select
End Function

Function ConvertTableType(value)
    Dim lowerVal

    lowerVal = LCase(Trim(CStr(value)))
    Select Case lowerVal
        Case "standard"
            ConvertTableType = "標準"
        Case "activity"
            ConvertTableType = "活動"
        Case "virtual"
            ConvertTableType = "仮想"
        Case Else
            ConvertTableType = value
    End Select
End Function

Function ConvertOwnershipType(value)
    Dim lowerVal

    lowerVal = LCase(Trim(CStr(value)))
    Select Case lowerVal
        Case "organizationowned"
            ConvertOwnershipType = "組織"
        Case "userowned"
            ConvertOwnershipType = "ユーザーまたはチーム"
        Case "businessowned"
            ConvertOwnershipType = ""
        Case Else
            ConvertOwnershipType = value
    End Select
End Function

Function ConvertFieldRequirement(value)
    Dim lowerVal

    lowerVal = LCase(Trim(CStr(value)))
    Select Case lowerVal
        Case "none"
            ConvertFieldRequirement = "任意"
        Case "applicationrequired"
            ConvertFieldRequirement = "必須"
        Case "systemrequired"
            ConvertFieldRequirement = "システム要求"
        Case Else
            ConvertFieldRequirement = value
    End Select
End Function

Function GetDateFormatLabel(additionalData)
    Dim formatValue

    formatValue = LCase(ExtractValueFromAdditionalData(additionalData, "Format:"))
    If InStr(formatValue, "dateandtime") > 0 Or InStr(formatValue, "datetime") > 0 Then
        GetDateFormatLabel = "日付と時刻 - 日時"
    ElseIf InStr(formatValue, "dateonly") > 0 Then
        GetDateFormatLabel = "日付と時刻 - 日付のみ"
    Else
        GetDateFormatLabel = ""
    End If
End Function

Function CleanAdditionalData(value)
    Dim p

    value = CStr(value)
    p = InStr(1, value, "Precision:", vbTextCompare)
    If p > 0 Then
        value = Left(value, p - 1)
    End If
    CleanAdditionalData = Trim(value)
End Function

Function ExtractValueFromAdditionalData(additionalData, keyword)
    Dim keywords, pos, valueStart, i, nextPos, nearestPos, tempValue

    keywords = Array("Minimum value:", "Maximum value:", "Options:", "Default:", "Default Value:", "Target:", "States:", "Format:", "targets:")
    pos = InStr(1, additionalData, keyword, vbTextCompare)
    If pos = 0 Then
        ExtractValueFromAdditionalData = ""
        Exit Function
    End If

    valueStart = pos + Len(keyword)
    nearestPos = 0
    For i = 0 To UBound(keywords)
        nextPos = InStr(valueStart, additionalData, keywords(i), vbTextCompare)
        If nextPos > 0 Then
            If nearestPos = 0 Or nextPos < nearestPos Then
                nearestPos = nextPos
            End If
        End If
    Next

    If nearestPos > 0 Then
        tempValue = Mid(additionalData, valueStart, nearestPos - valueStart)
    Else
        tempValue = Mid(additionalData, valueStart)
    End If

    tempValue = Replace(tempValue, vbCrLf, "")
    tempValue = Replace(tempValue, vbLf, "")
    tempValue = Replace(tempValue, vbCr, "")
    tempValue = Replace(tempValue, " ", "")
    tempValue = Trim(tempValue)
    ExtractValueFromAdditionalData = tempValue
End Function

Function ExtractTextMaxLength(value)
    ExtractTextMaxLength = ExtractMaxLengthByLabel(value, "TextMax Length:")
End Function

Function ExtractMaxLengthValue(value)
    Dim result

    result = ExtractMaxLengthByLabel(value, "Max length:")
    If result = "" Then
        result = ExtractMaxLengthByLabel(value, "TextMax Length:")
    End If
    ExtractMaxLengthValue = result
End Function

Function ExtractMaxLengthByLabel(value, labelText)
    Dim pos, tailText, i, ch, result

    pos = InStr(1, CStr(value), labelText, vbTextCompare)
    If pos = 0 Then
        ExtractMaxLengthByLabel = ""
        Exit Function
    End If

    tailText = Mid(CStr(value), pos + Len(labelText))
    result = ""
    For i = 1 To Len(tailText)
        ch = Mid(tailText, i, 1)
        If ch >= "0" And ch <= "9" Then
            result = result & ch
        ElseIf result <> "" Then
            Exit For
        End If
    Next
    ExtractMaxLengthByLabel = result
End Function

Function ConvertPrefixName(value)
    Dim mappings, i, sourcePrefix, destPrefix, textValue, lowerText

    textValue = Trim(CStr(value))
    If textValue = "" Then
        ConvertPrefixName = ""
        Exit Function
    End If

    mappings = Array( _
        Array("tel_wo_", "tel_wo_"), _
        Array("adx_", "tel_wo_adx_"), _
        Array("billto_", "tel_wo_billto_"), _
        Array("cnt_", "tel_wo_cnt_"), _
        Array("com_", "tel_wo_com_"), _
        Array("msa_", "tel_wo_msa_"), _
        Array("msdyn_", "tel_wo_"), _
        Array("mspp_", "tel_wo_"), _
        Array("parent_", "tel_wo_parent_"), _
        Array("resco_", "tel_wo_resco_"), _
        Array("shipto_", "tel_wo_shipto_"), _
        Array("tea_", "tel_wo_tea_"), _
        Array("tel_", "tel_wo_"), _
        Array("tsc_", "tel_wo_tsc_"), _
        Array("wo_", "tel_wo_") _
    )

    lowerText = LCase(textValue)
    For i = 0 To UBound(mappings)
        sourcePrefix = mappings(i)(0)
        destPrefix = mappings(i)(1)
        If Left(lowerText, Len(sourcePrefix)) = sourcePrefix Then
            ConvertPrefixName = destPrefix & Mid(textValue, Len(sourcePrefix) + 1)
            Exit Function
        End If
    Next

    ConvertPrefixName = textValue
End Function

Function ConvertPrefixList(value)
    Dim parts, i

    value = Trim(CStr(value))
    If value = "" Then
        ConvertPrefixList = ""
        Exit Function
    End If

    parts = Split(value, ",")
    For i = 0 To UBound(parts)
        parts(i) = ConvertPrefixName(Trim(parts(i)))
    Next
    ConvertPrefixList = Join(parts, ", ")
End Function

Function DefaultText(value, fallback)
    If Trim(CStr(value)) = "" Then
        DefaultText = fallback
    Else
        DefaultText = value
    End If
End Function

Function Nz(value)
    If IsNull(value) Or IsEmpty(value) Then
        Nz = ""
    Else
        Nz = value
    End If
End Function

Function FindTemplatePath(folderPath)
    Dim parentPath, templatePath

    parentPath = gFso.GetParentFolderName(folderPath)
    templatePath = gFso.BuildPath(parentPath, "テンプレート.xlsx")
    If gFso.FileExists(templatePath) Then
        FindTemplatePath = templatePath
        Exit Function
    End If

    templatePath = gFso.BuildPath(parentPath, "template.xlsx")
    If gFso.FileExists(templatePath) Then
        FindTemplatePath = templatePath
    Else
        FindTemplatePath = ""
    End If
End Function

Function FindMemoPath(folderPath)
    Dim parentPath, memoPath

    parentPath = gFso.GetParentFolderName(folderPath)
    memoPath = gFso.BuildPath(parentPath, "エンティティ定義書作成メモ.xlsx")
    If gFso.FileExists(memoPath) Then
        FindMemoPath = memoPath
    Else
        FindMemoPath = ""
    End If
End Function

Function CollectExcelFiles(folderPath, templatePath)
    Dim folder, file, count, items(), ext, templateName

    Set folder = gFso.GetFolder(folderPath)
    templateName = LCase(gFso.GetFileName(templatePath))
    count = -1

    For Each file In folder.Files
        ext = LCase(gFso.GetExtensionName(file.Name))
        If (ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "xlsb") Then
            If LCase(file.Name) <> templateName Then
                count = count + 1
                ReDim Preserve items(count)
                items(count) = file.Path
            End If
        End If
    Next

    If count = -1 Then Exit Function
    SortTextArray items
    CollectExcelFiles = items
End Function

Sub SortTextArray(ByRef items)
    Dim i, j, temp

    For i = 0 To UBound(items) - 1
        For j = i + 1 To UBound(items)
            If LCase(CStr(items(i))) > LCase(CStr(items(j))) Then
                temp = items(i)
                items(i) = items(j)
                items(j) = temp
            End If
        Next
    Next
End Sub

Sub EnsureFolder(folderPath)
    If Not gFso.FolderExists(folderPath) Then
        gFso.CreateFolder folderPath
    End If
End Sub

Function StartExcel(templatePath, ByRef message)
    StartExcel = False
    message = ""

    On Error Resume Next
    Set gExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        message = "Excelを起動できません: " & Err.Description
        Err.Clear
        Exit Function
    End If

    gExcel.Visible = False
    gExcel.DisplayAlerts = False
    gExcel.ScreenUpdating = False
    gExcel.EnableEvents = False
    gExcel.AskToUpdateLinks = False

    Err.Clear
    gExcel.Calculation = xlCalculationManual
    Err.Clear

    Set gTemplateWb = gExcel.Workbooks.Open(templatePath, 0, True)
    If Err.Number <> 0 Or gTemplateWb Is Nothing Then
        message = "テンプレートを開けません: " & templatePath & " / " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    StartExcel = True
End Function

Function LoadMemoMap(memoPath, ByRef message)
    Dim memoWb, memoWs, memoArr
    Dim rowNo, keyText, valueText

    LoadMemoMap = False
    message = ""

    If memoPath = "" Then
        message = "同じ階層に「エンティティ定義書作成メモ.xlsx」が見つかりません。"
        Exit Function
    End If

    Set gMemoMap = CreateObject("Scripting.Dictionary")
    Set memoWb = Nothing
    Set memoWs = Nothing

    On Error Resume Next
    Set memoWb = gExcel.Workbooks.Open(memoPath, 0, True)
    If Err.Number <> 0 Or memoWb Is Nothing Then
        message = "メモファイルを開けません: " & memoPath & " / " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Set memoWs = memoWb.Worksheets("作成対象一覧")
    If Err.Number <> 0 Or memoWs Is Nothing Then
        message = "メモファイルのシート「作成対象一覧」が見つかりません。"
        Err.Clear
        On Error GoTo 0
        CloseWorkbookSafe memoWb, False
        Exit Function
    End If

    memoArr = memoWs.Range("D3:F105").Value2
    On Error GoTo 0

    For rowNo = 1 To UBound(memoArr, 1)
        valueText = Trim(CStr(Nz(memoArr(rowNo, 1))))
        keyText = LCase(Trim(CStr(Nz(memoArr(rowNo, 3)))))
        If keyText <> "" And Not gMemoMap.Exists(keyText) Then
            gMemoMap.Add keyText, valueText
        End If
    Next

    CloseWorkbookSafe memoWb, False
    LoadMemoMap = True
End Function

Function SaveWorkbookAs(wb, outputPath, ByRef errMsg)
    SaveWorkbookAs = False
    errMsg = ""

    On Error Resume Next
    If gFso.FileExists(outputPath) Then
        gFso.DeleteFile outputPath, True
    End If
    wb.SaveAs outputPath, xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        errMsg = "保存に失敗しました: " & gFso.GetFileName(outputPath) & " / " & Err.Description
        Err.Clear
    Else
        SaveWorkbookAs = True
    End If
    On Error GoTo 0
End Function

Sub CloseWorkbookSafe(ByRef wb, ByVal saveChanges)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close saveChanges
        Set wb = Nothing
    End If
    On Error GoTo 0
End Sub

Sub Cleanup()
    On Error Resume Next

    CloseWorkbookSafe gTemplateWb, False

    If Not gExcel Is Nothing Then
        gExcel.Calculation = xlCalculationAutomatic
        gExcel.ScreenUpdating = True
        gExcel.EnableEvents = True
        gExcel.DisplayAlerts = True
        gExcel.Quit
        Set gExcel = Nothing
    End If

    Set gFso = Nothing
    Set gMemoMap = Nothing
    On Error GoTo 0
End Sub

Sub ApplyMemoMapping(wsField, rowCount)
    Dim rowNo, keyText, mappedText
    Dim lastRow

    If rowCount <= 0 Then Exit Sub

    lastRow = 6 + rowCount

    For rowNo = 7 To lastRow
        keyText = LCase(Trim(CStr(Nz(wsField.Cells(rowNo, 38).Value2))))

        If keyText <> "" Then
            If Not gMemoMap Is Nothing And gMemoMap.Exists(keyText) Then
                mappedText = gMemoMap(keyText)
                wsField.Cells(rowNo, 21).Value = mappedText
            Else
                wsField.Range(wsField.Cells(rowNo, 2), wsField.Cells(rowNo, 34)).Interior.Color = COLOR_GRAY
            End If
        End If
    Next
End Sub

Sub TrimEmptyFieldRows(wsField, rowCount)
    Dim lastRow, rowNo, lastValueRow

    If rowCount <= 0 Then Exit Sub

    lastRow = 6 + rowCount
    lastValueRow = 0

    For rowNo = lastRow To 7 Step -1
        If Trim(CStr(Nz(wsField.Cells(rowNo, 3).Value2))) <> "" Then
            lastValueRow = rowNo
            Exit For
        End If
    Next

    If lastValueRow = 0 Then
        wsField.Rows("7:" & wsField.Rows.Count).Delete
    ElseIf lastValueRow < wsField.Rows.Count Then
        wsField.Rows((lastValueRow + 1) & ":" & wsField.Rows.Count).Delete
    End If
End Sub

Sub AppendDetail(ByRef detail, ByVal message)
    If detail = "" Then
        detail = vbCrLf & vbCrLf & "詳細:" & vbCrLf & message
    Else
        detail = detail & vbCrLf & message
    End If
End Sub

Function SanitizeFileName(fileName)
    Dim chars, i

    chars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SanitizeFileName = fileName
    For i = 0 To UBound(chars)
        SanitizeFileName = Replace(SanitizeFileName, chars(i), "_")
    Next
End Function
