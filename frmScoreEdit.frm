VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScoreEdit 
   Caption         =   "得点修正フォーム"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8445.001
   OleObjectBlob   =   "frmScoreEdit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmScoreEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 編集対象のセル情報
Private mTargetRow As Long
Private mTargetCol As Long
Private mAllocateScore As Double

'===============================================================================
' フォーム初期化
'===============================================================================
Public Sub Initialize(ByVal targetRow As Long, ByVal targetCol As Long)
    mTargetRow = targetRow
    mTargetCol = targetCol

    ' テスト情報を取得して表示
    Dim testName As String
    Dim subjectName As String
    Dim perspectiveName As String
    Dim childName As String
    Dim currentScore As Variant

    With Sh_data
        testName = .Cells(eRowData.rowTestName, mTargetCol).value
        subjectName = .Cells(eRowData.rowSubject, mTargetCol).value
        perspectiveName = .Cells(eRowData.rowPerspective, mTargetCol).value
        mAllocateScore = .Cells(eRowData.rowAllocationScore, mTargetCol).value
        childName = .Cells(mTargetRow, eColData.colLastName).value & " " & _
                    .Cells(mTargetRow, eColData.colFirstName).value
        currentScore = .Cells(mTargetRow, mTargetCol).value
    End With

    ' ラベルに表示
    lblSubject.Caption = subjectName
    lblPerspective.Caption = perspectiveName
    lblTestname.Caption = testName
    lblChildName.Caption = childName
    lblAllocateScore.Caption = mAllocateScore & " 点"
    lblCurrentScore.Caption = IIf(IsEmpty(currentScore), "未入力", currentScore)

    ' 入力欄に現在の値をセット
    If Not IsEmpty(currentScore) Then
        txtNewScore.Text = CStr(currentScore)
    End If

    ' ヒントを更新
    lblHint.Caption = "「-」(免除)を選択した場合、成績算出時に計算から除外されます。"
    ' 入力欄にフォーカス
    txtNewScore.SetFocus
End Sub
'===============================================================================
' 更新ボタン
'===============================================================================
Private Sub btnUpdate_Click()
    Dim newScore As Variant
    Dim inputValue As String

    inputValue = Trim(txtNewScore.Text)

    ' 入力値の検証
    If inputValue = "" Then
        ' 空欄は許可（未入力として扱う）
        newScore = Empty
    ElseIf inputValue = "-" Then
        ' 欠席
        newScore = "-"
    ElseIf Not IsNumeric(inputValue) Then
        MsgBox "数値または「-」（免除）を入力してください。", vbExclamation, "入力エラー"
        txtNewScore.SetFocus
        Exit Sub
    Else
        newScore = CDbl(inputValue)

        ' 範囲チェック
        If newScore < 0 Then
            MsgBox "0以上の値を入力してください。", vbExclamation, "入力エラー"
            txtNewScore.SetFocus
            Exit Sub
        End If

        If newScore > mAllocateScore Then
            MsgBox "配点（" & mAllocateScore & "点）を超えています。", vbExclamation, "入力エラー"
            txtNewScore.SetFocus
            Exit Sub
        End If
    End If

    ' 保護を一時解除して更新
    On Error Resume Next
    Sh_data.Unprotect
    On Error GoTo 0

    Sh_data.Cells(mTargetRow, mTargetCol).value = newScore

    ' 保護を再設定
    Call ProtectScoreCells

    ' フォームを閉じる
    Unload Me
End Sub

'===============================================================================
' キャンセルボタン
'===============================================================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

'===============================================================================
' Escキーでキャンセル
'===============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 何もしない（通常のクローズを許可）
End Sub

Private Sub btn_Exempt_Click()
    txtNewScore.value = "-"
End Sub
'===============================================================================
' 得点セルの保護を設定
'===============================================================================
Private Sub ProtectScoreCells()
    Call DataManagementModule.ProtectScoreCells
End Sub


