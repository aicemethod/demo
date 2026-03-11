Option Explicit

Dim inputFolder
Dim outputFolder

inputFolder = SelectFolder("読み取り対象フォルダ（エクセルが格納されているフォルダ）を選択してください。")
If inputFolder = "" Then
    WScript.Quit 0
End If

outputFolder = SelectFolder("出力先フォルダを選択してください。")
If outputFolder = "" Then
    WScript.Quit 0
End If

WScript.Quit 0

Function SelectFolder(message)
    Dim shellApp
    Dim folder

    SelectFolder = ""

    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, message, 0)

    If Not folder Is Nothing Then
        SelectFolder = folder.Self.Path
    End If

    Set folder = Nothing
    Set shellApp = Nothing
End Function
