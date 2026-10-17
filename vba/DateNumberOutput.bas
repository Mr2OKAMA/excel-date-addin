Attribute VB_Name = "DateNumberOutput"
Option Explicit

Private Const SETTINGS_APP As String = "DateNumberOutput"
Private Const SETTINGS_SECTION As String = "Settings"
Private Const SETTINGS_KEY As String = "Prefix"

Public Sub ConfigurePrefix()
    Dim prefix As Variant

    On Error GoTo Failed
    prefix = Application.InputBox( _
        Prompt:="任意の数字を入力してください（先頭の0も保存されます）。", _
        Title:="日付付き番号出力 - 設定を保存", _
        Default:=GetSetting(SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, ""), _
        Type:=2)

    If VarType(prefix) = vbBoolean Then Exit Sub
    If Not IsValidPrefix(CStr(prefix)) Then
        MsgBox "数字のみ入力してください", vbExclamation, "日付付き番号出力"
        Exit Sub
    End If

    SaveSetting SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, CStr(prefix)
    MsgBox "保存しました", vbInformation, "日付付き番号出力"
    Exit Sub

Failed:
    MsgBox "設定を保存できませんでした。" & vbCrLf & Err.Description, _
        vbExclamation, "日付付き番号出力"
End Sub

Public Sub OutputDateNumber()
    Dim prefix As String
    Dim cell As Excel.Range

    On Error GoTo Failed
    prefix = GetSetting(SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, "")
    If Not IsValidPrefix(prefix) Then
        MsgBox "数字のみ入力してください。設定を保存し直してください。", _
            vbExclamation, "日付付き番号出力"
        Exit Sub
    End If

    If Not TypeOf Application.ActiveSheet Is Excel.Worksheet Then
        MsgBox "出力先のワークシートのセルを選択してください。", _
            vbExclamation, "日付付き番号出力"
        Exit Sub
    End If

    Set cell = Application.ActiveCell
    cell.NumberFormat = "@"
    cell.Value2 = prefix & Format$(Date, "yymmdd")
    Exit Sub

Failed:
    MsgBox "出力できませんでした。セルの保護や編集状態を確認してください。" & _
        vbCrLf & Err.Description, vbExclamation, "日付付き番号出力"
End Sub

Private Function IsValidPrefix(ByVal prefix As String) As Boolean
    IsValidPrefix = Not (prefix Like "*[!0-9]*")
End Function
