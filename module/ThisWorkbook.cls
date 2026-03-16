VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' モジュール名: ThisWorkbook
' 用途: ワークブックレベルのイベント処理
' メンテナンス:
'   - コメントアウトされていたコードを整理
'   - エラーハンドリング追加
'   - 初期化処理の整理
'===============================================================================
Option Explicit

'===============================================================================
' ワークブックが開かれた時の処理
'===============================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler

    ' 入力フォームの日付を更新
    Call InitializeInputForm

    ' Subjectシートのチェックボックスを初期化
    Call InitializeSubjectCheckboxes

    ' Resultシートの列見出しを初期化する
    Call InitializeResultHeaders

    ' データシートの得点セルを保護
    Call InitializeDataSheetProtection

    ' 全シートのボタン・コントロールを保護（UserInterfaceOnly）
    Call InitializeSheetProtection

    ' 教科未設定の場合、初期設定ウィザードを案内
    Call CheckInitialSetup

    Exit Sub

ErrorHandler:
    ' 起動エラーはユーザーに通知
    MsgBox "初期処理にエラーが発生しました。" & vbCrLf & _
           "一部の機能が正常に動作しない可能性があります。" & vbCrLf & vbCrLf & _
           "エラー: " & Err.Description, vbExclamation, "成績処理システム"
End Sub

'===============================================================================
' 入力フォームの初期化
'===============================================================================
Private Sub InitializeInputForm()
    On Error Resume Next

    ' 日本に合った日付を設定
    sh_input.Range(RNG_INPUT_DATE).value = Date

    On Error GoTo 0
End Sub

'===============================================================================
' Subjectシートのチェックボックス初期化
'===============================================================================
Private Sub InitializeSubjectCheckboxes()
    On Error Resume Next

    Dim i As Long

    With sh_subject
        For i = 1 To MAX_PERSPECTIVES
            If Trim(sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).value & "") = "" Then
                ' 観点が設定されていない場合は非表示
                .CheckBoxes("perspective" & i).Visible = False
            Else
                ' 観点名を設定して表示
                .CheckBoxes("perspective" & i).Caption = _
                    sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).value
                .CheckBoxes("perspective" & i).Visible = True
            End If
        Next i
    End With

    On Error GoTo 0
End Sub

'===============================================================================
' Resultシートの列見出し初期化
'===============================================================================
Private Sub InitializeResultHeaders()
    On Error Resume Next

    ' Resultシートの列見出しが未生成なら生成
    Call ResultModule.GenerateResultHeaders

    On Error GoTo 0
End Sub

'===============================================================================
' データシートの得点セル保護
'===============================================================================
Private Sub InitializeDataSheetProtection()
    On Error Resume Next

    ' 得点セルのみを保護（ダブルクリックでフォームから修正可能）
    Call DataManagementModule.ProtectScoreCells

    On Error GoTo 0
End Sub

'===============================================================================
' シート保護の初期化（ボタン・コントロールの誤削除防止）
' 説明: ボタンのあるシートをUserInterfaceOnly:=Trueで保護する。
'       VBAからの操作は制限されず、ユーザーの手動操作のみをブロックする。
'       UserInterfaceOnlyはファイル保存時に保持されないため、毎回Open時に設定。
'       ※ Sh_dataはProtectScoreCellsで別途保護済み
'===============================================================================
Private Sub InitializeSheetProtection()
    On Error Resume Next

    Dim ws As Variant
    Dim sheets As Variant
    sheets = Array(sh_input, sh_subject, sh_MENU, sh_namelist, sh_result)

    Dim i As Long
    For i = LBound(sheets) To UBound(sheets)
        Set ws = sheets(i)
        ws.Unprotect Password:=SHEET_PROTECT_PASSWORD
        ws.Protect Password:=SHEET_PROTECT_PASSWORD, _
            DrawingObjects:=True, Contents:=True, Scenarios:=False, _
            UserInterfaceOnly:=True
    Next i

    ' Settingシートは教科列（A-C）のみロック、他は編集可能
    Call InitialSetupModule.SetupSettingSheetProtection

    On Error GoTo 0
End Sub

'===============================================================================
' 右クリック無効化（フォームボタンの誤削除防止）
'===============================================================================
'Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
'    Cancel = True
'End Sub

'===============================================================================
' ワークブック終了前の処理
'===============================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next

    ' 必要に応じて保存確認などを追加

    On Error GoTo 0
End Sub

'===============================================================================
' ワークブック保存前の処理
'===============================================================================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next

    ' 保存前にデータの整合性チェックなどを行う場合はここに追加

    On Error GoTo 0
End Sub

'===============================================================================
' 初期設定チェック
' 教科が未設定の場合、セットアップウィザードを起動する
'===============================================================================
Private Sub CheckInitialSetup()
    On Error Resume Next

    Dim subjectCount As Long
    subjectCount = Application.WorksheetFunction.CountA(sh_setting.Range("B3:B20"))

    If subjectCount = 0 Then
        If MsgBox("教科が設定されていません。" & vbCrLf & _
                  "初期設定ウィザードを起動しますか？", _
                  vbYesNo + vbQuestion, "成績処理システム") = vbYes Then
            Call InitialSetupModule.ShowSetupForm
        End If
    End If

    On Error GoTo 0
End Sub
