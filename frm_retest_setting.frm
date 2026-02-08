'===============================================================================
' フォーム名: frm_retest_setting
' 説明: 追試計算方法設定フォーム
'       追試シートの「最終得点計算」ボタンから呼び出される
'       算出方法を選択し、決定すると最終得点列に数式を設定する
'
' コントロール:
'   opbtn1 - OptionButton: 合格点
'   opbtn2 - OptionButton: 最大値
'   opbtn3 - OptionButton: 平均値
'   opbtn4 - OptionButton: 中央値
'   opbtn5 - OptionButton: 内分点
'   opbtn6 - OptionButton: 本試のみ
'   txtbox - TextBox: 内分比α値（opbtn5選択時のみ有効）
'   btn_ok - CommandButton: 決定
'   btn_cancel - CommandButton: キャンセル
'===============================================================================
Option Explicit

' フォームが正常に完了したかどうか
Private m_Cancelled As Boolean

'===============================================================================
' フォーム初期化
'===============================================================================
Private Sub UserForm_Initialize()
    m_Cancelled = True

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' txtboxの初期状態を無効化
    txtbox.Enabled = False
    txtbox.BackColor = &H80000004  ' グレーアウト

    ' 現在の算出方法を読み取り、該当するラジオボタンを選択
    Dim currentMethod As String
    currentMethod = Trim(ws.Range(RNG_RT_METHOD).Value & "")

    Select Case currentMethod
        Case RT_METHOD_PASS_SCORE
            opbtn1.Value = True
        Case RT_METHOD_MAX
            opbtn2.Value = True
        Case RT_METHOD_AVERAGE
            opbtn3.Value = True
        Case RT_METHOD_MEDIAN
            opbtn4.Value = True
        Case RT_METHOD_INTERPOLATION
            opbtn5.Value = True
            txtbox.Enabled = True
            txtbox.BackColor = &H80000005  ' 白
            ' 現在のα値を読み込み
            Dim paramVal As Variant
            paramVal = ws.Range(RNG_RT_PARAM).Value
            If Trim(paramVal & "") <> "" Then
                txtbox.Value = CStr(paramVal)
            End If
        Case RT_METHOD_ORIGINAL_ONLY
            opbtn6.Value = True
        Case Else
            ' 未設定の場合は何も選択しない
    End Select
End Sub

'===============================================================================
' 内分点ラジオボタン選択時 - txtboxを有効化
'===============================================================================
Private Sub opbtn5_Click()
    txtbox.Enabled = True
    txtbox.BackColor = &H80000005  ' 白
    txtbox.SetFocus
End Sub

'===============================================================================
' 内分点以外のラジオボタン選択時 - txtboxを無効化
'===============================================================================
Private Sub opbtn1_Click()
    Call DisableTextBox
End Sub

Private Sub opbtn2_Click()
    Call DisableTextBox
End Sub

Private Sub opbtn3_Click()
    Call DisableTextBox
End Sub

Private Sub opbtn4_Click()
    Call DisableTextBox
End Sub

Private Sub opbtn6_Click()
    Call DisableTextBox
End Sub

Private Sub DisableTextBox()
    txtbox.Enabled = False
    txtbox.BackColor = &H80000004  ' グレーアウト
End Sub

'===============================================================================
' 決定ボタン
'===============================================================================
Private Sub btn_ok_Click()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim selectedMethod As String
    Dim alphaValue As Double

    ' 選択された算出方法を判定
    If opbtn1.Value Then
        selectedMethod = RT_METHOD_PASS_SCORE
    ElseIf opbtn2.Value Then
        selectedMethod = RT_METHOD_MAX
    ElseIf opbtn3.Value Then
        selectedMethod = RT_METHOD_AVERAGE
    ElseIf opbtn4.Value Then
        selectedMethod = RT_METHOD_MEDIAN
    ElseIf opbtn5.Value Then
        selectedMethod = RT_METHOD_INTERPOLATION
    ElseIf opbtn6.Value Then
        selectedMethod = RT_METHOD_ORIGINAL_ONLY
    Else
        MsgBox "算出方法を選択してください。", vbExclamation, "未選択"
        Exit Sub
    End If

    ' 内分点の場合、α値のバリデーション
    If selectedMethod = RT_METHOD_INTERPOLATION Then
        Dim txtVal As String
        txtVal = Trim(txtbox.Value)

        If txtVal = "" Then
            MsgBox "α値を入力してください。", vbExclamation, "入力エラー"
            txtbox.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(txtVal) Then
            MsgBox "α値は数値で入力してください。", vbExclamation, "入力エラー"
            txtbox.SetFocus
            Exit Sub
        End If

        alphaValue = CDbl(txtVal)

        If alphaValue < 0 Or alphaValue > 1 Then
            MsgBox "α値は0～1の範囲で入力してください。" & vbCrLf & _
                   "（1に近いほど追試最高点寄り、0に近いほど本試寄り）", _
                   vbExclamation, "入力エラー"
            txtbox.SetFocus
            Exit Sub
        End If
    End If

    ' 合格点方式の場合、合格点が入力されているかチェック
    If selectedMethod = RT_METHOD_PASS_SCORE Then
        Dim passScore As Variant
        passScore = ws.Range(RNG_RT_PASS_SCORE).Value
        If Trim(passScore & "") = "" Or Not IsNumeric(passScore) Then
            MsgBox "合格点方式を使用するには、合格点（セル " & RNG_RT_PASS_SCORE & "）に" & vbCrLf & _
                   "数値を入力してください。", vbExclamation, "合格点未設定"
            Exit Sub
        End If
        If CDbl(passScore) <= 0 Then
            MsgBox "合格点は0より大きい値を入力してください。", vbExclamation, "入力エラー"
            Exit Sub
        End If
    End If

    ' 算出方法をシートに書き込み
    ws.Range(RNG_RT_METHOD).Value = selectedMethod

    ' 内分点の場合、α値をシートに書き込み
    If selectedMethod = RT_METHOD_INTERPOLATION Then
        ws.Range(RNG_RT_PARAM).Value = alphaValue
    Else
        ' 内分点以外の場合はパラメータをクリア
        ws.Range(RNG_RT_PARAM).Value = ""
    End If

    m_Cancelled = False
    Me.Hide
End Sub

'===============================================================================
' キャンセルボタン
'===============================================================================
Private Sub btn_cancel_Click()
    m_Cancelled = True
    Me.Hide
End Sub

'===============================================================================
' ×ボタンで閉じた場合
'===============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        m_Cancelled = True
        Me.Hide
    End If
End Sub

'===============================================================================
' キャンセルされたかどうかを返す
'===============================================================================
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property
