Option Explicit

'────────────────────────────────────────
'  処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルからDisplayNameを取得し、
'  template.xlsxをベースに出力ファイルを作成
'  ※高速化ポイント
'    ・template.xlsx は最初に1回だけ開き、以降は Worksheets.Copy で複製
'    ・シート2（フィールド）の読み取りは配列で一括取得してから参照
'    ・シート「フィールド」への書き込みも配列で一括書き込み
'────────────────────────────────────────

Dim fso, excel, wb, ws, wbTemplateMaster, wbOutput, wsTable, wsCover, wsField
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim templatePath, outputFolderPath, outputFilePath
Dim displayName
Dim lastCol, col, colName
Dim outputFileName
Dim mappingDict, colIndexDict
Dim fieldName, cellAddr, fieldValue
Dim rowNum, colNum
Dim convertedValue, lowerVal
Dim primaryNameAttribute, ws2, colIndexDict2
Dim logicalNameCol, lastRow2, row, foundRow
Dim attributeMappingDict, attrFieldName, attrCellAddr, attrValue
Dim rowLogicalName, pluralDisplayName
Dim maxLengthPos, maxLengthValue, afterMaxLength, i, char
Dim e13Value
Dim fieldMappingDict, outputRow, fieldValue2, outputCol
Dim additionalDataValue, targetsValue, formatValue, targetsPos, formatPos
Dim lowerFormatValue, attributeTypeValue
Dim attrTypeConverted, minValue, maxValue, optionsValue, defaultValue, targetValue, statesValue
Dim precisionPos
Dim formatLabelJP, lastOutputRow
Dim dataArr
Dim outArr, outRows, outCols, startRowField, startColField, arrRow, arrCol

' ▼ 引数チェック（ドラッグ&ドロップされたフォルダのパス）
If WScript.Arguments.Count = 0 Then
    MsgBox "フォルダをドラッグ&ドロップしてください。", vbCritical, "エラー"
    WScript.Quit
End If

folderPath = WScript.Arguments(0)

' ▼ FileSystemObject作成
Set fso = CreateObject("Scripting.FileSystemObject")

' ▼ フォルダの存在チェック
If Not fso.FolderExists(folderPath) Then
    MsgBox "指定されたパスはフォルダではありません: " & folderPath, vbCritical, "エラー"
    WScript.Quit
End If

Set folder = fso.GetFolder(folderPath)

' ▼ template.xlsxのパスを取得（同じ階層）
templatePath = fso.BuildPath(fso.GetParentFolderName(folderPath), "template.xlsx")

' ▼ template.xlsxの存在チェック
If Not fso.FileExists(templatePath) Then
    MsgBox "template.xlsxが見つかりません: " & templatePath, vbCritical, "エラー"
    WScript.Quit
End If

' ▼ 出力フォルダのパスを取得（同じ階層の「20_作成済定義書」）
outputFolderPath = fso.BuildPath(fso.GetParentFolderName(folderPath), "20_作成済定義書")

' ▼ 出力フォルダが存在しない場合は作成
If Not fso.FolderExists(outputFolderPath) Then
    fso.CreateFolder outputFolderPath
End If

' ▼ 共通マッピング（ループの外で一度だけ作成）
Set mappingDict = CreateObject("Scripting.Dictionary")
mappingDict.Add "Schema Name", "E9"
mappingDict.Add "Logical Name", "E10"
mappingDict.Add "Ownership Type", "E12"
mappingDict.Add "Change TrackingEnabled", "E23"
mappingDict.Add "Description", "E7"
mappingDict.Add "DisplayCollectionName", "E6"
mappingDict.Add "EntityColor", "E14"
mappingDict.Add "EntityHelpUrl", "E25"
mappingDict.Add "EntityHelpUrlEnabled", "E24"
mappingDict.Add "HasEmailAddresses", "E35"
mappingDict.Add "HasFeedback", "E37"
mappingDict.Add "HasNotes", "E8"
mappingDict.Add "IsActivity", "E30"
mappingDict.Add "IsAuditEnabled", "E26"
mappingDict.Add "IsAvailableOffline", "E39"
mappingDict.Add "IsConnectionsEnabled", "E34"
mappingDict.Add "IsDocumentManagementEnabled", "E33"
mappingDict.Add "IsDuplicateDetection Enabled", "E22"
mappingDict.Add "IsMailMergeEnabled", "E31"
mappingDict.Add "IsQuickCreateEnabled", "E27"
mappingDict.Add "IsRetentionEnabled", "E28"
mappingDict.Add "IsSLAEnabled", "E32"
mappingDict.Add "IsValidForAdvancedFind", "E38"
mappingDict.Add "IsValidForQueue", "E40"
mappingDict.Add "PrimarylmageAttribute", "E15"
mappingDict.Add "TableType", "E11"

' フィールドマッピング辞書を作成（フィールド名 → 列番号）
Set fieldMappingDict = CreateObject("Scripting.Dictionary")
fieldMappingDict.Add "Schema Name", 4    ' D列
fieldMappingDict.Add "Display Name", 5   ' E列
fieldMappingDict.Add "Custom Attribute", 7 ' G列
fieldMappingDict.Add "Attribute Type", 10  ' J列（後で個別処理）
fieldMappingDict.Add "Type", 11            ' K列
fieldMappingDict.Add "Required Level", 12  ' L列
fieldMappingDict.Add "Description", 31     ' AE列
fieldMappingDict.Add "Audit Enabled", 24   ' X列
fieldMappingDict.Add "Secured", 25         ' Y列
fieldMappingDict.Add "ValidFor AdvancedFind", 28 ' AB列

' ▼ Excel起動
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
' 高速化設定
excel.ScreenUpdating = False
excel.EnableEvents = False
' Calculation変更は環境によってエラーになるのでエラー無視で実行
On Error Resume Next
excel.Calculation = -4135   ' xlCalculationManual
On Error GoTo 0

' ▼ template.xlsx をマスタとして一度だけ開く（読み取り専用）
On Error Resume Next
Set wbTemplateMaster = excel.Workbooks.Open(templatePath, 0, True)
If Err.Number <> 0 Or wbTemplateMaster Is Nothing Then
    MsgBox "template.xlsxを開けませんでした: " & Err.Description, vbCritical, "エラー"
    Err.Clear
    On Error GoTo 0
    excel.Quit
    Set excel = Nothing
    WScript.Quit
End If
On Error GoTo 0

' ▼ フォルダ内のExcelファイルを順に処理
For Each file In folder.Files
    fileName = file.Name
    fileExt = LCase(fso.GetExtensionName(fileName))
    
    ' Excelファイルの拡張子をチェック
    If fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Or fileExt = "xlsb" Then
        filePath = file.Path
        
        ' 前のループで作成されたオブジェクトを解放（メモリリーク防止）
        On Error Resume Next
        If Not colIndexDict Is Nothing Then
            Set colIndexDict = Nothing
        End If
        On Error GoTo 0
        dataArr = Array() ' 配列のメモリ解放
        
        On Error Resume Next
        ' Excelファイルを開く（読み取り専用）
        Set wb = excel.Workbooks.Open(filePath, 0, True)
        
        If Err.Number = 0 Then
            On Error GoTo 0
            
            ' 1シート目を取得
            Set ws = wb.Sheets(1)
            
            ' 3行目（ヘッダー行）から列インデックスを作成
            Set colIndexDict = CreateObject("Scripting.Dictionary")
            lastCol = ws.Cells(3, ws.Columns.Count).End(-4159).Column ' xlToLeft
            
            For col = 1 To lastCol
                colName = Trim(CStr(ws.Cells(3, col).Value2))
                If colName <> "" Then
                    colIndexDict(LCase(colName)) = col
                End If
            Next
            
            ' DisplayNameを取得（ファイル名生成用）
            displayName = ""
            If colIndexDict.Exists("displayname") Then
                On Error Resume Next
                displayName = CStr(ws.Cells(4, colIndexDict("displayname")).Value2)
                If Err.Number <> 0 Then
                    displayName = ""
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            
            ' DisplayNameが見つかった場合のみ処理
            If displayName <> "" Then
                ' ▼ マスタ template から新しいブックをコピーして出力用ブックを作成
                On Error Resume Next
                wbTemplateMaster.Worksheets.Copy  ' 全シートを含む新規ブックが作成される
                If Err.Number = 0 Then
                    On Error GoTo 0
                    
                    Set wbOutput = excel.ActiveWorkbook
                    
                    ' シート「テーブル」「表紙」「フィールド」を取得
                    On Error Resume Next
                    Set wsTable = wbOutput.Sheets("テーブル")
                    Set wsCover = wbOutput.Sheets("表紙")
                    Set wsField = wbOutput.Sheets("フィールド")
                    
                    If Err.Number = 0 Then
                        On Error GoTo 0
                        
                        ' マッピングに従って値をセット（テーブルシート）
                        For Each fieldName In mappingDict.Keys
                            cellAddr = mappingDict(fieldName)
                            fieldValue = ""
                            
                            ' 列インデックスから値を取得
                            If colIndexDict.Exists(LCase(fieldName)) Then
                                On Error Resume Next
                                fieldValue = ws.Cells(4, colIndexDict(LCase(fieldName))).Value2
                                If Err.Number <> 0 Then
                                    fieldValue = ""
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            End If
                            
                            ' True/Falseを変換（True → ✓、False → -）
                            convertedValue = fieldValue
                            If IsNumeric(fieldValue) = False Then
                                lowerVal = LCase(Trim(CStr(fieldValue)))
                                If lowerVal = "true" Then
                                    convertedValue = ChrW(10003) ' ✓
                                ElseIf lowerVal = "false" Or lowerVal = "" Then
                                    convertedValue = "-"
                                End If
                            End If
                            
                            ' セルアドレスを解析（例：E5 → 行5、列5）
                            colNum = Asc(UCase(Left(cellAddr, 1))) - 64 ' A=1, B=2, ..., E=5
                            rowNum = CInt(Mid(cellAddr, 2))
                            
                            ' 値をセット（赤文字）
                            With wsTable.Cells(rowNum, colNum)
                                .Value = convertedValue
                                .Font.Color = RGB(255, 0, 0)
                            End With
                        Next
                        
                        ' ▼ Plural Display Nameの値をE5にセット
                        pluralDisplayName = ""
                        If colIndexDict.Exists("plural display name") Then
                            On Error Resume Next
                            pluralDisplayName = ws.Cells(4, colIndexDict("plural display name")).Value2
                            If Err.Number <> 0 Then
                                pluralDisplayName = ""
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                        
                        If pluralDisplayName <> "" Then
                            With wsTable.Cells(5, 5)
                                .Value = pluralDisplayName
                                .Font.Color = RGB(255, 0, 0)
                            End With
                        End If
                        
                        ' ▼ E13の値をチェック（空の場合は「なし」をセット）
                        e13Value = wsTable.Cells(13, 5).Value
                        If e13Value = "" Or IsEmpty(e13Value) Then
                            With wsTable.Cells(13, 5)
                                .Value = "なし"
                                .Font.Color = RGB(255, 0, 0)
                            End With
                        End If
                        
                        ' ▼ PrimaryNameAttributeの値を取得（シート1の4行目から）
                        primaryNameAttribute = ""
                        If colIndexDict.Exists("primarynameattribute") Then
                            On Error Resume Next
                            primaryNameAttribute = CStr(ws.Cells(4, colIndexDict("primarynameattribute")).Value2)
                            If Err.Number <> 0 Then
                                primaryNameAttribute = ""
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                        
                        ' ▼ シート2（フィールド元データ）を一括読み込み
                        If wb.Sheets.Count >= 2 Then
                            Set ws2 = wb.Sheets(2)
                            
                            ' シート2の1行目（ヘッダー行）から列インデックスを作成
                            Set colIndexDict2 = CreateObject("Scripting.Dictionary")
                            lastCol = ws2.Cells(1, ws2.Columns.Count).End(-4159).Column ' xlToLeft
                            
                            For col = 1 To lastCol
                                colName = Trim(CStr(ws2.Cells(1, col).Value2))
                                If colName <> "" Then
                                    colIndexDict2(LCase(colName)) = col
                                End If
                            Next
                            
                            ' シート2の最終行を取得
                            lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(-4162).Row ' xlUp
                            
                            ' 1行目～最終行を一括で配列に読み込む（読み取り高速化）
                            If lastRow2 >= 1 And lastCol >= 1 Then
                                dataArr = ws2.Range(ws2.Cells(1, 1), ws2.Cells(lastRow2, lastCol)).Value
                            Else
                                dataArr = Array()
                            End If
                            
                            ' ▼ PrimaryNameAttributeからテーブルシートに値をセット
                            If primaryNameAttribute <> "" Then
                                ' Logical Name列の位置を取得
                                logicalNameCol = 0
                                If colIndexDict2.Exists("logical name") Then
                                    logicalNameCol = colIndexDict2("logical name")
                                End If
                                
                                ' PrimaryNameAttributeと一致する行を検索（配列から検索）
                                foundRow = 0
                                If logicalNameCol > 0 And lastRow2 >= 2 Then
                                    For row = 2 To lastRow2
                                        On Error Resume Next
                                        rowLogicalName = Trim(CStr(dataArr(row, logicalNameCol)))
                                        If Err.Number <> 0 Then
                                            rowLogicalName = ""
                                            Err.Clear
                                        End If
                                        On Error GoTo 0
                                        
                                        If LCase(rowLogicalName) = LCase(primaryNameAttribute) Then
                                            foundRow = row
                                            Exit For
                                        End If
                                    Next
                                End If
                                
                                ' 該当行が見つかった場合、値をセット
                                If foundRow > 0 Then
                                    ' 属性マッピング辞書を作成
                                    Set attributeMappingDict = CreateObject("Scripting.Dictionary")
                                    attributeMappingDict.Add "Display Name", "E16"
                                    attributeMappingDict.Add "Description", "E17"
                                    attributeMappingDict.Add "Schema Name", "E18"
                                    attributeMappingDict.Add "Logical Name", "E19"
                                    attributeMappingDict.Add "Required Level", "E20"
                                    attributeMappingDict.Add "Additional data", "E21"
                                    
                                    ' 各属性の値を取得してセット
                                    For Each attrFieldName In attributeMappingDict.Keys
                                        attrCellAddr = attributeMappingDict(attrFieldName)
                                        attrValue = ""
                                        
                                        ' 列インデックスから値を取得（配列から）
                                        If colIndexDict2.Exists(LCase(attrFieldName)) Then
                                            On Error Resume Next
                                            attrValue = dataArr(foundRow, colIndexDict2(LCase(attrFieldName)))
                                            If Err.Number <> 0 Then
                                                attrValue = ""
                                                Err.Clear
                                            End If
                                            On Error GoTo 0
                                        End If
                                        
                                        ' E21（Additional data）の場合は「Max length:」の後の数値のみを抽出
                                        If attrCellAddr = "E21" Then
                                            maxLengthValue = ""
                                            maxLengthPos = InStr(1, CStr(attrValue), "Max length:", vbTextCompare)
                                            If maxLengthPos > 0 Then
                                                afterMaxLength = Mid(CStr(attrValue), maxLengthPos + Len("Max length:"))
                                                ' 数値部分を抽出
                                                For i = 1 To Len(afterMaxLength)
                                                    char = Mid(afterMaxLength, i, 1)
                                                    If IsNumeric(char) Then
                                                        maxLengthValue = maxLengthValue & char
                                                    ElseIf maxLengthValue <> "" Then
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                            convertedValue = maxLengthValue
                                        Else
                                            ' True/Falseを変換
                                            convertedValue = attrValue
                                            If IsNumeric(attrValue) = False Then
                                                lowerVal = LCase(Trim(CStr(attrValue)))
                                                If lowerVal = "true" Then
                                                    convertedValue = ChrW(10003) ' ✓
                                                ElseIf lowerVal = "false" Or lowerVal = "" Then
                                                    convertedValue = "-"
                                                End If
                                            End If
                                        End If
                                        
                                        ' セルアドレスを解析
                                        colNum = Asc(UCase(Left(attrCellAddr, 1))) - 64
                                        rowNum = CInt(Mid(attrCellAddr, 2))
                                        
                                        ' 値をセット（赤文字）
                                        With wsTable.Cells(rowNum, colNum)
                                            .Value = convertedValue
                                            .Font.Color = RGB(255, 0, 0)
                                        End With
                                    Next
                                    
                                    Set attributeMappingDict = Nothing
                                End If
                            End If
                            
                            ' ▼ シート2の2行目以降のデータをシート「フィールド」に出力（配列一括書き込み）
                            If lastRow2 >= 2 Then
                                startRowField = 7        ' 出力開始行
                                startColField = 4        ' D列
                                outRows = lastRow2 - 1   ' データ行数（2行目〜）
                                outCols = 37 - startColField + 1 ' D〜AK列
                                
                                ' VBScriptの配列は0ベース。Excelに書き込む際は1ベースとして解釈されるため、
                                ' インデックス1からoutRows、1からoutColsまで使うためにサイズをoutRows、outColsにする
                                ReDim outArr(outRows, outCols)
                                
                                arrRow = 0  ' 0ベースで開始（Excelに書き込む際は1ベースとして解釈される）
                                
                                For row = 2 To lastRow2
                                    ' 各フィールドの値を取得して outArr にセット（共通処理）
                                    For Each fieldName In fieldMappingDict.Keys
                                        outputCol = fieldMappingDict(fieldName)
                                        
                                        ' Attribute Type はここでは書き込まず、後続の専用処理で書く
                                        If fieldName <> "Attribute Type" Then
                                            fieldValue2 = ""
                                            
                                            ' 列インデックスから値を取得（配列から）
                                            If colIndexDict2.Exists(LCase(fieldName)) Then
                                                On Error Resume Next
                                                fieldValue2 = dataArr(row, colIndexDict2(LCase(fieldName)))
                                                If Err.Number <> 0 Then
                                                    fieldValue2 = ""
                                                    Err.Clear
                                                End If
                                                On Error GoTo 0
                                            End If
                                            
                                            ' フィールドごとの変換処理（空の値は空のまま）
                                            convertedValue = fieldValue2
                                            
                                            ' 空の値の場合は変換せずに空のまま
                                            If fieldValue2 = "" Or IsEmpty(fieldValue2) Then
                                                convertedValue = ""
                                            ElseIf IsNumeric(fieldValue2) = False Then
                                                lowerVal = LCase(Trim(CStr(fieldValue2)))
                                                
                                                ' Custom Attributeの変換（True → カスタム、False → 標準）
                                                If fieldName = "Custom Attribute" Then
                                                    If lowerVal = "true" Then
                                                        convertedValue = "カスタム"
                                                    ElseIf lowerVal = "false" Then
                                                        convertedValue = "標準"
                                                    End If
                                                ' Typeの変換（Simple → シンプル、Calculated → 計算、Rollup → ロールアップ）
                                                ElseIf fieldName = "Type" Then
                                                    Select Case lowerVal
                                                        Case "simple"
                                                            convertedValue = "シンプル"
                                                        Case "calculated"
                                                            convertedValue = "計算"
                                                        Case "rollup"
                                                            convertedValue = "ロールアップ"
                                                        Case Else
                                                            convertedValue = fieldValue2
                                                    End Select
                                                ' Required Levelの変換
                                                ElseIf fieldName = "Required Level" Then
                                                    Select Case lowerVal
                                                        Case "none"
                                                            convertedValue = "任意"
                                                        Case "applicationrequired"
                                                            convertedValue = "システム要求"
                                                        Case "systemrequired"
                                                            convertedValue = "必須項目"
                                                        Case "recommended"
                                                            convertedValue = "推奨項目"
                                                        Case Else
                                                            convertedValue = fieldValue2
                                                    End Select
                                                ' その他のTrue/FalseはTRUE/FALSEに変換
                                                Else
                                                    If lowerVal = "true" Then
                                                        convertedValue = "TRUE"
                                                    ElseIf lowerVal = "false" Then
                                                        convertedValue = "FALSE"
                                                    End If
                                                End If
                                            End If
                                            
                                            ' outArr にセット（0ベースのインデックス）
                                            arrCol = outputCol - startColField
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = convertedValue
                                            End If
                                        End If
                                    Next
                                    
                                    ' ▼ Additional data の取得（Precision: 以降は削除）
                                    additionalDataValue = ""
                                    If colIndexDict2.Exists("additional data") Then
                                        On Error Resume Next
                                        additionalDataValue = CStr(dataArr(row, colIndexDict2("additional data")))
                                        If Err.Number <> 0 Then
                                            additionalDataValue = ""
                                            Err.Clear
                                        End If
                                        On Error GoTo 0
                                        
                                        ' Precision: 以降を削除
                                        precisionPos = InStr(1, additionalDataValue, "Precision:", vbTextCompare)
                                        If precisionPos > 0 Then
                                            additionalDataValue = Left(additionalDataValue, precisionPos - 1)
                                        End If
                                    End If
                                    
                                    ' ▼ targets: の処理（V列 = 22列目）
                                    targetsValue = ""
                                    If additionalDataValue <> "" Then
                                        targetsPos = InStr(1, additionalDataValue, "targets:", vbTextCompare)
                                        If targetsPos > 0 Then
                                            targetsValue = Mid(additionalDataValue, targetsPos + Len("targets:"))
                                            ' 改行やスペースを取り除く
                                            targetsValue = Replace(targetsValue, vbCrLf, "")
                                            targetsValue = Replace(targetsValue, vbLf, "")
                                            targetsValue = Replace(targetsValue, vbCr, "")
                                            targetsValue = Replace(targetsValue, " ", "")
                                            targetsValue = Trim(targetsValue)
                                            
                                            arrCol = 22 - startColField ' V列（0ベース）
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = targetsValue
                                            End If
                                        End If
                                    End If
                                    
                                    ' ▼ Format: の処理（DateTime/DateOnly 判定用）
                                    formatValue = ""
                                    formatLabelJP = ""
                                    
                                    If additionalDataValue <> "" Then
                                        formatPos = InStr(1, additionalDataValue, "Format:", vbTextCompare)
                                        If formatPos > 0 Then
                                            formatValue = Mid(additionalDataValue, formatPos + Len("Format:"))
                                            ' 改行や余計な空白を取り除く
                                            formatValue = Replace(formatValue, vbCrLf, "")
                                            formatValue = Replace(formatValue, vbLf, "")
                                            formatValue = Replace(formatValue, vbCr, "")
                                            formatValue = Trim(formatValue)
                                            
                                            lowerFormatValue = LCase(formatValue)
                                            ' DateAndTime / DateTime → 日時
                                            If InStr(lowerFormatValue, "dateandtime") > 0 Or InStr(lowerFormatValue, "datetime") > 0 Then
                                                formatLabelJP = "日付と時刻 - 日時"
                                            ' DateOnly → 日付のみ
                                            ElseIf InStr(lowerFormatValue, "dateonly") > 0 Then
                                                formatLabelJP = "日付と時刻 - 日付のみ"
                                            Else
                                                ' その他はそのまま保持
                                                formatLabelJP = formatValue
                                            End If
                                        End If
                                    End If
                                    
                                    ' ▼ Attribute Type の変換と Additional data の反映
                                    attrTypeConverted = ""
                                    minValue = ""
                                    maxValue = ""
                                    optionsValue = ""
                                    defaultValue = ""
                                    targetValue = ""
                                    statesValue = ""
                                    
                                    If colIndexDict2.Exists("attribute type") Then
                                        On Error Resume Next
                                        attributeTypeValue = CStr(dataArr(row, colIndexDict2("attribute type")))
                                        If Err.Number <> 0 Then
                                            attributeTypeValue = ""
                                            Err.Clear
                                        End If
                                        On Error GoTo 0
                                        
                                        attributeTypeValue = Trim(CStr(attributeTypeValue))
                                        attributeTypeValue = Replace(attributeTypeValue, vbCrLf, "")
                                        attributeTypeValue = Replace(attributeTypeValue, vbLf, "")
                                        attributeTypeValue = Replace(attributeTypeValue, vbCr, "")
                                        
                                        lowerVal = LCase(Trim(attributeTypeValue))
                                        
                                        Select Case lowerVal
                                            Case "bigint"
                                                attrTypeConverted = "数値 - 整数(Int)"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "choice"
                                                attrTypeConverted = "選択肢"
                                                optionsValue = ExtractValueFromAdditionalData(additionalDataValue, "Options:")
                                                defaultValue = ExtractValueFromAdditionalData(additionalDataValue, "Default:")
                                                If LCase(Trim(defaultValue)) = "n/a" Then
                                                    defaultValue = "なし"
                                                End If
                                            Case "choices"
                                                attrTypeConverted = "選択肢(複数)"
                                                optionsValue = ExtractValueFromAdditionalData(additionalDataValue, "Options:")
                                                defaultValue = ExtractValueFromAdditionalData(additionalDataValue, "Default:")
                                                If LCase(Trim(defaultValue)) = "n/a" Then
                                                    defaultValue = "なし"
                                                End If
                                            Case "currency"
                                                attrTypeConverted = "通貨"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "decimal"
                                                attrTypeConverted = "数値 - 少数(10進数)"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "double"
                                                attrTypeConverted = "数値 - 浮動小数点数"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "multiline text"
                                                attrTypeConverted = "複数行テキスト - プレーン"
                                            Case "owner"
                                                attrTypeConverted = "所有者"
                                                targetValue = ExtractValueFromAdditionalData(additionalDataValue, "Target:")
                                            Case "state"
                                                attrTypeConverted = "状態"
                                                statesValue = ExtractValueFromAdditionalData(additionalDataValue, "States:")
                                            Case "status"
                                                attrTypeConverted = "ステータス"
                                                statesValue = ExtractValueFromAdditionalData(additionalDataValue, "States:")
                                            Case "text"
                                                attrTypeConverted = "1行テキスト - プレーン"
                                            Case "two options"
                                                attrTypeConverted = "はい/いいえ"
                                                optionsValue = ExtractValueFromAdditionalData(additionalDataValue, "Options:")
                                                defaultValue = ExtractValueFromAdditionalData(additionalDataValue, "Default Value:")
                                            Case "uniqueidentifier"
                                                attrTypeConverted = "一意識別子"
                                            Case "whole number"
                                                attrTypeConverted = "数値 - 整数(Int)"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "datetime", "dateandtime"
                                                ' 日付時刻型は Format: の情報を優先
                                                If formatLabelJP <> "" Then
                                                    attrTypeConverted = formatLabelJP
                                                Else
                                                    attrTypeConverted = "日付と時刻"
                                                End If
                                            Case Else
                                                ' Lookupの場合は検索に変換
                                                If lowerVal = "lookup" Then
                                                    attrTypeConverted = "検索"
                                                Else
                                                    attrTypeConverted = attributeTypeValue
                                                End If
                                        End Select
                                        
                                        ' Attribute Type を J列（10列目）にセット
                                        If attrTypeConverted <> "" Then
                                            arrCol = 10 - startColField
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = attrTypeConverted
                                            End If
                                        End If
                                        
                                        ' Additional data 由来の各種値をセット
                                        ' Minimum value → P列（16列目）
                                        If minValue <> "" Then
                                            arrCol = 16 - startColField
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = minValue
                                            End If
                                        End If
                                        
                                        ' Maximum value → O列（15列目）
                                        If maxValue <> "" Then
                                            arrCol = 15 - startColField
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = maxValue
                                            End If
                                        End If
                                        
                                        ' Options: → T列（20列目）
                                        If optionsValue <> "" Then
                                            arrCol = 20 - startColField
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = optionsValue
                                            End If
                                        End If
                                        
                                        ' Default: / Default Value: → U列（21列目）
                                        If defaultValue <> "" Then
                                            arrCol = 21 - startColField
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = defaultValue
                                            End If
                                        End If
                                        
                                        ' Target: → V列（22列目）
                                        If targetValue <> "" Then
                                            arrCol = 22 - startColField
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = targetValue
                                            End If
                                        End If
                                        
                                        ' States: → T列（20列目）
                                        If statesValue <> "" Then
                                            arrCol = 20 - startColField
                                            If arrCol >= 0 And arrCol < outCols Then
                                                outArr(arrRow, arrCol) = statesValue
                                            End If
                                        End If
                                        
                                        ' Multiline Text/Text の場合は Additional data 全体を AK列（37列目）にセット
                                        If lowerVal = "multiline text" Or lowerVal = "text" Then
                                            If additionalDataValue <> "" Then
                                                arrCol = 37 - startColField
                                                If arrCol >= 0 And arrCol < outCols Then
                                                    outArr(arrRow, arrCol) = additionalDataValue
                                                End If
                                            End If
                                        End If
                                    End If
                                    
                                    arrRow = arrRow + 1
                                Next
                                
                                ' ★フィールドシートの出力範囲を一括で書き込み＆赤文字にする
                                ' ExcelのRangeに配列を書き込む際、配列のインデックス0が1行目として解釈される
                                With wsField
                                    .Range(.Cells(startRowField, startColField), .Cells(startRowField + outRows - 1, startColField + outCols - 1)).Value = outArr
                                    .Range(.Cells(startRowField, startColField), .Cells(startRowField + outRows - 1, startColField + outCols - 1)).Font.Color = RGB(255, 0, 0)
                                End With
                            End If
                            
                            Set colIndexDict2 = Nothing
                            Set ws2 = Nothing
                            dataArr = Array()
                        End If
                        
                        ' シート「表紙」のB7に「エンティティ定義書_ID_<DisplayNameの値>_v0.1」をセット
                        wsCover.Cells(7, 2).Value = "エンティティ定義書_ID_" & displayName & "_v0.1"
                        
                        ' 出力ファイル名を生成
                        outputFileName = "エンティティ定義書_ID_" & displayName & "_v0.1.xlsx"
                        outputFilePath = fso.BuildPath(outputFolderPath, outputFileName)
                        
                        ' 既存ファイルがある場合は削除
                        If fso.FileExists(outputFilePath) Then
                            fso.DeleteFile outputFilePath, True
                        End If
                        
                        ' ファイルを保存
                        wbOutput.SaveAs outputFilePath
                        
                        ' 出力ブックを閉じる
                        wbOutput.Close False
                        Set wbOutput = Nothing
                        Set wsTable = Nothing
                        Set wsCover = Nothing
                        Set wsField = Nothing
                    Else
                        MsgBox "シート「テーブル」または「表紙」または「フィールド」が見つかりません: " & fileName, vbCritical, "エラー"
                        Err.Clear
                        On Error GoTo 0
                        If Not wbOutput Is Nothing Then
                            wbOutput.Close False
                            Set wbOutput = Nothing
                        End If
                    End If
                Else
                    MsgBox "templateコピー時にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
                    Err.Clear
                    On Error GoTo 0
                End If
            Else
                MsgBox "DisplayNameが見つかりませんでした: " & fileName, vbWarning, "警告"
            End If
            
            ' 元のファイルを閉じる
            wb.Close False
            Set wb = Nothing
            Set ws = Nothing
        Else
            MsgBox "ファイルを開けませんでした: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
            Err.Clear
            On Error GoTo 0
        End If
        
        ' 各ファイル処理後にcolIndexDictを解放（メモリリーク防止）
        If Not colIndexDict Is Nothing Then
            Set colIndexDict = Nothing
        End If
    End If
Next

' ▼ マスタ template.xlsx を閉じる
If Not wbTemplateMaster Is Nothing Then
    wbTemplateMaster.Close False
    Set wbTemplateMaster = Nothing
End If

' ▼ Excel終了（設定を戻してから Quit）
On Error Resume Next
excel.Calculation = -4105   ' xlCalculationAutomatic（失敗しても無視）
excel.ScreenUpdating = True
excel.EnableEvents = True
On Error GoTo 0

excel.Quit
Set excel = Nothing

MsgBox "処理が完了しました。", vbInformation, "完了"

'────────────────────────────────────────
'  関数：Additional dataから値を抽出
'────────────────────────────────────────
Function ExtractValueFromAdditionalData(additionalData, keyword)
    Dim pos, valueStart, valueEnd, nextKeywordPos
    Dim result, tempValue
    Dim keywords
    Dim i, keywordPos, minPos
    
    result = ""
    pos = InStr(1, additionalData, keyword, vbTextCompare)
    
    If pos > 0 Then
        valueStart = pos + Len(keyword)
        
        ' 次のキーワードの位置を探す
        keywords = Array("Minimum value:", "Maximum value:", "Options:", "Default:", "Default Value:", "Target:", "States:", "Format:", "targets:")
        minPos = 0
        
        For i = 0 To UBound(keywords)
            keywordPos = InStr(valueStart, additionalData, keywords(i), vbTextCompare)
            If keywordPos > 0 Then
                If minPos = 0 Or keywordPos < minPos Then
                    minPos = keywordPos
                End If
            End If
        Next
        
        If minPos > 0 Then
            tempValue = Mid(additionalData, valueStart, minPos - valueStart)
        Else
            tempValue = Mid(additionalData, valueStart)
        End If
        
        ' 改行やスペースを取り除く
        tempValue = Replace(tempValue, vbCrLf, "")
        tempValue = Replace(tempValue, vbLf, "")
        tempValue = Replace(tempValue, vbCr, "")
        tempValue = Replace(tempValue, " ", "")
        result = Trim(tempValue)
    End If
    
    ExtractValueFromAdditionalData = result
End Function
