Attribute VB_Name = "ErrorHandlerModule"
Option Explicit

'===============================================================================
' エラー情報を格納する構造体
'===============================================================================
Public Type ErrorInfo
    ErrorNumber As Long
    ErrorDescription As String
    ProcedureName As String
    moduleName As String
    additionalInfo As String
End Type

'===============================================================================
' 共通エラー表示関数
' 説明: ユーザーフレンドリーなエラーメッセージを表示
' ※VBAではユーザー定義型はByRefでのみ渡せる
'===============================================================================
Public Sub ShowError(ByRef errInfo As ErrorInfo)
    Dim msg As String
    
    msg = "【エラーが発生しました】" & vbCrLf & vbCrLf
    msg = msg & "■ 何が起きたか:" & vbCrLf
    msg = msg & "  " & GetFriendlyErrorMessage(errInfo.ErrorNumber, errInfo.ErrorDescription) & vbCrLf & vbCrLf
    
    If errInfo.additionalInfo <> "" Then
        msg = msg & "■ 詳細情報:" & vbCrLf
        msg = msg & "  " & errInfo.additionalInfo & vbCrLf & vbCrLf
    End If
    
    msg = msg & "■ 対処方法:" & vbCrLf
    msg = msg & "  " & GetRecoverySuggestion(errInfo.ErrorNumber) & vbCrLf & vbCrLf
    
    msg = msg & "━━━━━━━━━━━━━━━━━━━━" & vbCrLf
    msg = msg & "（技術情報: " & errInfo.moduleName & "." & errInfo.ProcedureName
    msg = msg & " / エラー" & errInfo.ErrorNumber & "）"
    
    MsgBox msg, vbCritical, "成績処理システム - エラー"
End Sub

'===============================================================================
' ユーザーフレンドリーなエラーメッセージを返す
'===============================================================================
Private Function GetFriendlyErrorMessage(ByVal errNum As Long, ByVal errDesc As String) As String
    Select Case errNum
        Case 6      ' オーバーフロー
            GetFriendlyErrorMessage = "数値が大きすぎます。入力した数値を確認してください。"
        Case 9      ' インデックスが有効範囲にありません
            GetFriendlyErrorMessage = "データが見つかりませんでした。シートの構成が変更されている可能性があります。"
        Case 11     ' 0で除算
            GetFriendlyErrorMessage = "計算でゼロ割りが発生しました。配点が0になっていないか確認してください。"
        Case 13     ' 型が一致しません
            GetFriendlyErrorMessage = "入力された値の形式が正しくありません。数値を入力すべき欄に文字が入っていないか確認してください。"
        Case 91     ' オブジェクト変数が設定されていません
            GetFriendlyErrorMessage = "必要なシートまたはセルが見つかりませんでした。"
        Case 1004   ' アプリケーション定義またはオブジェクト定義のエラー
            GetFriendlyErrorMessage = "Excelの操作でエラーが発生しました。シートが保護されていないか確認してください。"
        Case Else
            GetFriendlyErrorMessage = errDesc
    End Select
End Function

'===============================================================================
' 回復方法の提案を返す
'===============================================================================
Private Function GetRecoverySuggestion(ByVal errNum As Long) As String
    Select Case errNum
        Case 6      ' オーバーフロー
            GetRecoverySuggestion = "得点や配点に極端に大きな数値が入力されていないか確認してください。"
        Case 9      ' インデックスが有効範囲にありません
            GetRecoverySuggestion = "シート名が変更されていないか確認してください。解決しない場合は、管理者にご連絡ください。"
        Case 11     ' 0で除算
            GetRecoverySuggestion = "配点欄を確認し、0以外の値を入力してください。"
        Case 13     ' 型が一致しません
            GetRecoverySuggestion = "入力欄に正しい形式の値を入力してください。（例：点数欄には数値のみ）"
        Case 91     ' オブジェクト変数が設定されていません
            GetRecoverySuggestion = "ファイルを閉じて再度開いてみてください。解決しない場合は、管理者にご連絡ください。"
        Case 1004
            GetRecoverySuggestion = "シートの保護を解除するか、管理者にご連絡ください。"
        Case Else
            GetRecoverySuggestion = "操作をやり直してください。問題が続く場合は、管理者にご連絡ください。"
    End Select
End Function

'===============================================================================
' 入力検証エラーを表示（カスタムメッセージ用）
'===============================================================================
Public Sub ShowValidationError(ByVal message As String, Optional ByVal title As String = "入力エラー")
    Dim msg As String
    
    msg = "【入力内容に問題があります】" & vbCrLf & vbCrLf
    msg = msg & message & vbCrLf & vbCrLf
    msg = msg & "入力内容を確認して、もう一度お試しください。"
    
    MsgBox msg, vbExclamation, "成績処理システム - " & title
End Sub

'===============================================================================
' 確認ダイアログを表示
'===============================================================================
Public Function ShowConfirmation(ByVal message As String, Optional ByVal title As String = "確認") As Boolean
    Dim result As VbMsgBoxResult
    result = MsgBox(message, vbYesNo + vbQuestion, "成績処理システム - " & title)
    ShowConfirmation = (result = vbYes)
End Function

'===============================================================================
' 情報メッセージを表示
'===============================================================================
Public Sub ShowInfo(ByVal message As String, Optional ByVal title As String = "お知らせ")
    MsgBox message, vbInformation, "成績処理システム - " & title
End Sub

'===============================================================================
' 成功メッセージを表示
'===============================================================================
Public Sub ShowSuccess(ByVal message As String, Optional ByVal title As String = "完了")
    MsgBox message, vbInformation, "成績処理システム - " & title
End Sub

'===============================================================================
' エラー情報を作成するヘルパー関数
'===============================================================================
Public Function CreateErrorInfo(ByVal moduleName As String, ByVal procName As String, _
                                Optional ByVal additionalInfo As String = "") As ErrorInfo
    Dim info As ErrorInfo
    info.ErrorNumber = Err.Number
    info.ErrorDescription = Err.Description
    info.moduleName = moduleName
    info.ProcedureName = procName
    info.additionalInfo = additionalInfo
    CreateErrorInfo = info
End Function

'===============================================================================
' 処理開始時の共通設定
'===============================================================================
Public Sub BeginProcess()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

'===============================================================================
' 処理終了時の共通設定
'===============================================================================
Public Sub EndProcess()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

'===============================================================================
' エラー発生時のクリーンアップ（必ず呼び出す）
'===============================================================================
Public Sub CleanupOnError()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    On Error GoTo 0
End Sub


