Option Explicit

'========================================================================
'
'   Function:   CSVの重複行を削除する
'   Usage:      このVBScriptファイルにCSVファイルを1つドラッグ＆ドロップする
'   Version:    2.0 (エラーになるフォルダを開く機能を無効化)
'
'========================================================================

' --- オブジェクトの準備 ---
Dim fso, dict, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set dict = CreateObject("Scripting.Dictionary")
Set shell = CreateObject("WScript.Shell")

' --- ドラッグ＆ドロップされたファイルのチェック ---
If WScript.Arguments.Count = 0 Then
    MsgBox "CSVファイルをこのアイコンにドラッグ＆ドロップしてください。", vbExclamation, "エラー"
    WScript.Quit
End If

If WScript.Arguments.Count > 1 Then
    MsgBox "一度に処理できるファイルは1つだけです。", vbExclamation, "エラー"
    WScript.Quit
End If

Dim inputFileFullPath
inputFileFullPath = WScript.Arguments(0)

' --- ファイル拡張子のチェック ---
If LCase(fso.GetExtensionName(inputFileFullPath)) <> "csv" Then
    MsgBox "これはCSVファイルではありません。" & vbCrLf & inputFileFullPath, vbExclamation, "エラー"
    WScript.Quit
End If

' --- 出力ファイル名の決定 ---
Dim inputParentFolder, inputBaseName, inputExt, outputFileFullPath
inputParentFolder = fso.GetParentFolderName(inputFileFullPath)
inputBaseName = fso.GetBaseName(inputFileFullPath)
inputExt = fso.GetExtensionName(inputFileFullPath)
outputFileFullPath = fso.BuildPath(inputParentFolder, inputBaseName & "_unique." & inputExt)

' --- 変数の宣言 ---
Dim objInFile, objOutFile
Dim line
Dim lineCountInput, lineCountOutput, lineCountRemoved
lineCountInput = 0
lineCountOutput = 0

' --- ファイル処理の開始 ---
On Error Resume Next

' 入力ファイルを開く (1 = 読み取り専用)
Set objInFile = fso.OpenTextFile(inputFileFullPath, 1)
If Err.Number <> 0 Then
    MsgBox "入力ファイルを開けません。" & vbCrLf & "ファイルが使用中でないか確認してください。" & vbCrLf & Err.Description, vbCritical, "ファイル エラー"
    WScript.Quit
End If

' 出力ファイルを作成 (True = 上書き許可)
Set objOutFile = fso.CreateTextFile(outputFileFullPath, True)
If Err.Number <> 0 Then
    MsgBox "出力ファイルを作成できません。" & vbCrLf & "書き込み権限があるか確認してください。" & vbCrLf & Err.Description, vbCritical, "ファイル エラー"
    objInFile.Close
    WScript.Quit
End If

On Error GoTo 0

' --- ヘッダー行の処理 ---
If Not objInFile.AtEndOfStream Then
    ' 最初の行 (ヘッダー) を無条件で読み書き
    line = objInFile.ReadLine()
    objOutFile.WriteLine(line)
    
    ' ヘッダー行をDictionaryに追加（ヘッダー自体が重複判定の基準になるため）
    dict.Add line, 1 ' 値(1)はダミー
    
    lineCountInput = 1
    lineCountOutput = 1
End If

' --- 2行目以降のデータ処理 ---
Do Until objInFile.AtEndOfStream
    line = objInFile.ReadLine()
    lineCountInput = lineCountInput + 1
    
    ' Dictionaryにその行(line)が存在するかチェック
    If Not dict.Exists(line) Then
        ' 存在しない場合 (ユニークな行)
        dict.Add line, 1       ' Dictionaryにキーとして追加
        objOutFile.WriteLine(line) ' 出力ファイルに書き込み
        lineCountOutput = lineCountOutput + 1
    Else
        ' 存在する場合 (重複行)
        ' 何もしない
    End If
Loop

' --- 後処理 ---
objInFile.Close
objOutFile.Close
Set fso = Nothing
Set dict = Nothing

' --- 完了メッセージ ---
lineCountRemoved = lineCountInput - lineCountOutput
MsgBox "処理が完了しました。" & vbCrLf & vbCrLf & _
       "入力ファイル: " & fso.GetFileName(inputFileFullPath) & vbCrLf & _
       "出力ファイル: " & fso.GetFileName(outputFileFullPath) & vbCrLf & _
       "----------------------------------" & vbCrLf & _
       "読み込んだ行数: " & lineCountInput & " 行" & vbCrLf & _
       "ユニークな行数: " & lineCountOutput & " 行" & vbCrLf & _
       "削除した重複行数: " & lineCountRemoved & " 行", _
       vbInformation, "重複削除完了"

' ▼▼▼ 修正箇所 ▼▼▼
' エラーの原因となるため、以下の行をコメントアウト（無効化）します。
' shell.Exec "explorer.exe /select,""" & outputFileFullPath & """"
' ▲▲▲ 修正箇所 ▲▲▲

Set shell = Nothing

WScript.Quit