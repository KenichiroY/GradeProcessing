Attribute VB_Name = "MigrationModule"
'===============================================================================
' モジュール名: MigrationModule
' 機能: ActiveXコントロールをフォームコントロールに変換するマイグレーションマクロ
' 使用方法: MigrateAllActiveXToFormControls を1回実行する
' 注意: 実行後、このモジュールは削除して構いません
'===============================================================================
Option Explicit

'===============================================================================
' メイン: 全ActiveXコントロールをフォームコントロールに変換
'===============================================================================
Public Sub MigrateAllActiveXToFormControls()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' 確認ダイアログ
    If MsgBox("ActiveXコントロールをフォームコントロールに変換します。" & vbCrLf & vbCrLf & _
              "対象:" & vbCrLf & _
              "  - Subjectシート: チェックボックス(perspective1-5)、ボタン3個" & vbCrLf & _
              "  - 名簿シート: 登録ボタン(CommandButton1)" & vbCrLf & vbCrLf & _
              "実行前にファイルのバックアップを取ることを推奨します。" & vbCrLf & _
              "続行しますか？", _
              vbQuestion + vbYesNo, "ActiveX→フォームコントロール変換") = vbNo Then
        Exit Sub
    End If

    ' シート保護を一時解除
    On Error Resume Next
    sh_subject.Unprotect Password:=SHEET_PROTECT_PASSWORD
    sh_namelist.Unprotect Password:=SHEET_PROTECT_PASSWORD
    On Error GoTo ErrorHandler

    ' Subjectシートの変換
    Call MigrateSubjectCheckboxes
    Call MigrateSubjectButtons

    ' 名簿シートの変換
    Call MigrateNamelistButtons

    ' シート保護を再設定
    On Error Resume Next
    sh_subject.Protect Password:=SHEET_PROTECT_PASSWORD, _
        DrawingObjects:=True, Contents:=True, Scenarios:=False, _
        UserInterfaceOnly:=True
    sh_namelist.Protect Password:=SHEET_PROTECT_PASSWORD, _
        DrawingObjects:=True, Contents:=True, Scenarios:=False, _
        UserInterfaceOnly:=True
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = True

    MsgBox "変換が完了しました。" & vbCrLf & vbCrLf & _
           "ファイルを保存してから動作確認を行ってください。" & vbCrLf & _
           "問題なければ MigrationModule は削除して構いません。", _
           vbInformation, "変換完了"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "変換中にエラーが発生しました。" & vbCrLf & _
           "エラー: " & Err.Description & vbCrLf & vbCrLf & _
           "バックアップから復元してください。", _
           vbCritical, "エラー"
End Sub

'===============================================================================
' Subjectシートのチェックボックス変換 (perspective1-5)
'===============================================================================
Private Sub MigrateSubjectCheckboxes()
    Dim i As Long
    Dim oleObj As OLEObject
    Dim cb As CheckBox
    Dim l As Double, t As Double, w As Double, h As Double
    Dim cap As String
    Dim vis As Boolean

    For i = 1 To MAX_PERSPECTIVES
        On Error Resume Next
        Set oleObj = sh_subject.OLEObjects("perspective" & i)
        On Error GoTo 0

        If oleObj Is Nothing Then
            ' ActiveXが見つからない場合はスキップ（既に変換済みかも）
            GoTo NextCheckbox
        End If

        ' 位置・サイズ・状態を取得
        l = oleObj.Left
        t = oleObj.Top
        w = oleObj.Width
        h = oleObj.Height
        cap = Trim(sh_setting.Cells(i + 2, SETTING_PERSPECTIVE_COL).value & "")
        If cap = "" Then cap = "perspective" & i
        vis = oleObj.Visible

        ' ActiveXを削除
        oleObj.Delete
        Set oleObj = Nothing

        ' フォームチェックボックスを作成
        Set cb = sh_subject.CheckBoxes.Add(l, t, w, h)
        With cb
            .Name = "perspective" & i
            .Caption = cap
            .Visible = vis
            .value = xlOff
            ' 3D表示をオフ（フラットな見た目）
            .Display3DShading = False
        End With

        Set cb = Nothing

NextCheckbox:
        Set oleObj = Nothing
    Next i
End Sub

'===============================================================================
' Subjectシートのボタン変換 (Update, Ope_result, Delete_Sh_Subject)
'===============================================================================
Private Sub MigrateSubjectButtons()
    Dim btnDefs(1 To 3, 1 To 3) As String
    ' ボタン名, OnActionマクロ名
    btnDefs(1, 1) = "Update"
    btnDefs(1, 2) = "sh_subject.Update_Click"
    btnDefs(2, 1) = "Ope_result"
    btnDefs(2, 2) = "sh_subject.Ope_result_Click"
    btnDefs(3, 1) = "Delete_Sh_Subject"
    btnDefs(3, 2) = "sh_subject.Delete_Sh_Subject_Click"
    btnDefs(1, 3) = "追加/更新"
    btnDefs(2, 3) = "評価"
    btnDefs(3, 3) = "消去"

    Dim i As Long
    Dim oleObj As OLEObject
    Dim btn As Button
    Dim l As Double, t As Double, w As Double, h As Double
    Dim cap As String

    For i = 1 To 3
        On Error Resume Next
        Set oleObj = sh_subject.OLEObjects(btnDefs(i, 1))
        On Error GoTo 0

        If oleObj Is Nothing Then
            GoTo NextSubjectButton
        End If

        ' 位置・サイズ・キャプションを取得
        l = oleObj.Left
        t = oleObj.Top
        w = oleObj.Width
        h = oleObj.Height
        cap = btnDefs(i, 3)

        ' ActiveXを削除
        oleObj.Delete
        Set oleObj = Nothing

        ' フォームボタンを作成
        Set btn = sh_subject.Buttons.Add(l, t, w, h)
        With btn
            .Name = btnDefs(i, 1)
            .Caption = cap
            .OnAction = btnDefs(i, 2)
            .Font.Size = 9
        End With

        Set btn = Nothing

NextSubjectButton:
        Set oleObj = Nothing
    Next i
End Sub

'===============================================================================
' 名簿シートのボタン変換 (CommandButton1 → Btn_Posting)
'===============================================================================
Private Sub MigrateNamelistButtons()
    Dim oleObj As OLEObject
    Dim btn As Button
    Dim l As Double, t As Double, w As Double, h As Double
    Dim cap As String

    On Error Resume Next
    Set oleObj = sh_namelist.OLEObjects("CommandButton1")
    On Error GoTo 0

    If oleObj Is Nothing Then
        Exit Sub
    End If

    ' 位置・サイズ・キャプションを取得
    l = oleObj.Left
    t = oleObj.Top
    w = oleObj.Width
    h = oleObj.Height
        cap = "登録"

    ' ActiveXを削除
    oleObj.Delete
    Set oleObj = Nothing

    ' フォームボタンを作成
    Set btn = sh_namelist.Buttons.Add(l, t, w, h)
    With btn
        .Name = "Btn_Posting"
        .Caption = cap
        .OnAction = "sh_namelist.Btn_Posting_Click"
        .Font.Size = 9
    End With

    Set btn = Nothing
End Sub

