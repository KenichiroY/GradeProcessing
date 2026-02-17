VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestEdit
   Caption         =   "テスト情報の編集"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11100
   OleObjectBlob   =   "frmTestEdit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTestEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' フォーム名: frmTestEdit
' 機能: テスト情報の編集 + 後出し追試設定
'       データシートのヘッダー行（4-22行）ダブルクリックで表示
'
' コントロール構成:
'   lblKeyValue       - Label: テストキー（読取専用）
'   lblSubjectValue   - Label: 教科（読取専用）
'   cmbCategoryValue  - ComboBox: カテゴリ（データシートの既存値から取得）
'   txtTestName       - TextBox: テスト名（編集可）
'   cmbPerspective    - ComboBox: 観点（sh_settingからリスト取得）
'   txtDetail         - TextBox: 詳細（編集可）
'   txtAllocateScore  - TextBox: 配点（編集可、数値検証）
'   cmbYear           - ComboBox: 実施日（年）
'   cmbMonth          - ComboBox: 実施日（月）
'   cmbDay            - ComboBox: 実施日（日）
'   btnUpdate         - CommandButton: 更新ボタン
'   btnCancel         - CommandButton: キャンセルボタン
'   btnRetest         - CommandButton: 追試を設定ボタン
'   btnDelete         - CommandButton: 削除ボタン
'===============================================================================
Option Explicit

' 編集対象の列番号
Private mTargetCol As Long

' 削除リクエストフラグ（フォームを閉じた後に呼び出し元で判定）
Private mDeleteRequested As Boolean

' 追試中テストの強制削除フラグ
Private mForceDeleteRetest As Boolean

'===============================================================================
' フォーム初期化
' 引数: targetCol - データシートの対象列番号
'===============================================================================
Public Sub Initialize(ByVal targetCol As Long)
    mTargetCol = targetCol
    mDeleteRequested = False
    mForceDeleteRetest = False

    ' === コンボボックスの選択肢を初期化 ===
    Call LoadCategoryList
    Call LoadPerspectiveList
    Call InitDatePickers

    With Sh_data
        ' テスト情報読み込み（読取専用フィールド）
        lblKeyValue.Caption = .Cells(eRowData.rowKey, mTargetCol).value
        lblSubjectValue.Caption = .Cells(eRowData.rowSubject, mTargetCol).value

        ' カテゴリ: コンボボックスに現在値をセット
        Dim currentCategory As String
        currentCategory = .Cells(eRowData.rowCategory, mTargetCol).value & ""
        SetComboValue cmbCategoryValue, currentCategory

        ' 編集可フィールド
        txtTestName.Text = .Cells(eRowData.rowTestName, mTargetCol).value & ""
        txtDetail.Text = .Cells(eRowData.rowDetail, mTargetCol).value & ""
        txtAllocateScore.Text = .Cells(eRowData.rowAllocationScore, mTargetCol).value & ""

        ' 観点: コンボボックスに現在値をセット
        Dim currentPerspective As String
        currentPerspective = .Cells(eRowData.rowPerspective, mTargetCol).value & ""
        SetComboValue cmbPerspective, currentPerspective

        ' 実施日: 年・月・日コンボボックスに現在値をセット
        Dim dateVal As Variant
        dateVal = .Cells(eRowData.rowTestDate, mTargetCol).value
        If IsDate(dateVal) Then
            Dim d As Date
            d = CDate(dateVal)
            SetComboValue cmbYear, CStr(Year(d))
            SetComboValue cmbMonth, CStr(Month(d))
            SetComboValue cmbDay, CStr(Day(d))
        End If

        ' 追試ボタンの有効/無効設定
        Dim hasRetestMarker As Boolean
        Dim firstChildVal As Variant
        firstChildVal = .Cells(eRowData.rowChildStart, mTargetCol).value
        hasRetestMarker = (CStr(firstChildVal) = RETEST_MARKER)

        If hasRetestMarker Then
            btnRetest.Caption = "追試設定済み"
            btnRetest.Enabled = False
        Else
            btnRetest.Caption = "追試を設定"
            btnRetest.Enabled = True
        End If
    End With

    ' テスト名にフォーカス
    txtTestName.SetFocus
End Sub

'===============================================================================
' カテゴリリストの読み込み（sh_settingのG列から取得）
'===============================================================================
Private Sub LoadCategoryList()
    cmbCategoryValue.Clear
    cmbCategoryValue.Style = fmStyleDropDownList

    Dim i As Long
    Dim val As String
    With sh_setting
        For i = SETTING_SUBJECT_START_ROW To SETTING_SUBJECT_START_ROW + 10
            val = Trim(.Cells(i, SETTING_CATEGORY_COL).value & "")
            If val = "" Then Exit For
            cmbCategoryValue.AddItem val
        Next i
    End With
End Sub

'===============================================================================
' 観点リストの読み込み（sh_settingから）
'===============================================================================
Private Sub LoadPerspectiveList()
    cmbPerspective.Clear
    cmbPerspective.Style = fmStyleDropDownList

    Dim i As Long
    Dim val As String
    With sh_setting
        For i = SETTING_SUBJECT_START_ROW To SETTING_SUBJECT_START_ROW + 10
            val = Trim(.Cells(i, SETTING_PERSPECTIVE_COL).value & "")
            If val = "" Then Exit For
            cmbPerspective.AddItem val
        Next i
    End With
End Sub

'===============================================================================
' 日付コンボボックスの初期化（年・月・日）
'===============================================================================
Private Sub InitDatePickers()
    Dim i As Long

    ' 年: 現在年の前後2年
    cmbYear.Clear
    cmbYear.Style = fmStyleDropDownList
    Dim currentYear As Long
    currentYear = Year(Date)
    For i = currentYear - 2 To currentYear + 1
        cmbYear.AddItem CStr(i)
    Next i

    ' 月: 1-12
    cmbMonth.Clear
    cmbMonth.Style = fmStyleDropDownList
    For i = 1 To 12
        cmbMonth.AddItem CStr(i)
    Next i

    ' 日: 1-31（月変更時に動的更新）
    Call UpdateDayList
End Sub

'===============================================================================
' 日リストの更新（年・月に応じて末日を変更）
'===============================================================================
Private Sub UpdateDayList()
    Dim selectedDay As String
    selectedDay = cmbDay.Text

    cmbDay.Clear
    cmbDay.Style = fmStyleDropDownList

    ' 年・月から末日を計算
    Dim maxDay As Long
    Dim y As Long, m As Long
    If IsNumeric(cmbYear.Text) And IsNumeric(cmbMonth.Text) Then
        y = CLng(cmbYear.Text)
        m = CLng(cmbMonth.Text)
        ' 翌月1日の前日 = 当月末日
        If m = 12 Then
            maxDay = 31
        Else
            maxDay = Day(DateSerial(y, m + 1, 0))
        End If
    Else
        maxDay = 31
    End If

    Dim i As Long
    For i = 1 To maxDay
        cmbDay.AddItem CStr(i)
    Next i

    ' 元の選択を復元（範囲内なら）
    If IsNumeric(selectedDay) Then
        If CLng(selectedDay) <= maxDay Then
            SetComboValue cmbDay, selectedDay
        Else
            SetComboValue cmbDay, CStr(maxDay)
        End If
    End If
End Sub

'===============================================================================
' 年・月変更時に日リストを更新
'===============================================================================
Private Sub cmbYear_Change()
    Call UpdateDayList
End Sub

Private Sub cmbMonth_Change()
    Call UpdateDayList
End Sub

'===============================================================================
' コンボボックスに値をセットするヘルパー
'===============================================================================
Private Sub SetComboValue(ByVal cmb As MSForms.ComboBox, ByVal val As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = val Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i
    ' リストにない場合は未選択のままにする
End Sub

'===============================================================================
' 日付コンボボックスからDate値を取得するヘルパー
' 戻り値: 有効ならDate、無効ならEmpty
'===============================================================================
Private Function GetDateFromCombos() As Variant
    If Not IsNumeric(cmbYear.Text) Or _
       Not IsNumeric(cmbMonth.Text) Or _
       Not IsNumeric(cmbDay.Text) Then
        GetDateFromCombos = Empty
        Exit Function
    End If

    Dim y As Long, m As Long, dayVal As Long
    y = CLng(cmbYear.Text)
    m = CLng(cmbMonth.Text)
    dayVal = CLng(cmbDay.Text)

    If m < 1 Or m > 12 Or dayVal < 1 Or dayVal > 31 Then
        GetDateFromCombos = Empty
        Exit Function
    End If

    On Error GoTo InvalidDate
    GetDateFromCombos = DateSerial(y, m, dayVal)
    Exit Function

InvalidDate:
    GetDateFromCombos = Empty
End Function

'===============================================================================
' 更新ボタン
'===============================================================================
Private Sub btnUpdate_Click()
    ' === 入力検証 ===

    ' テスト名: 必須
    If Trim(txtTestName.Text) = "" Then
        MsgBox "テスト名を入力してください。", vbExclamation, "入力エラー"
        txtTestName.SetFocus
        Exit Sub
    End If

    ' 観点: 必須
    If Trim(cmbPerspective.Text) = "" Then
        MsgBox "観点を選択または入力してください。", vbExclamation, "入力エラー"
        cmbPerspective.SetFocus
        Exit Sub
    End If

    ' 配点: 必須、数値、0より大きい
    If Not IsNumeric(txtAllocateScore.Text) Then
        MsgBox "配点には数値を入力してください。", vbExclamation, "入力エラー"
        txtAllocateScore.SetFocus
        Exit Sub
    End If
    If CDbl(txtAllocateScore.Text) <= 0 Then
        MsgBox "配点は0より大きい値を入力してください。", vbExclamation, "入力エラー"
        txtAllocateScore.SetFocus
        Exit Sub
    End If

    ' 実施日: 年・月・日がすべて選択されているか
    Dim newDate As Variant
    newDate = GetDateFromCombos()
    If IsEmpty(newDate) Then
        MsgBox "実施日を正しく選択してください。", vbExclamation, "入力エラー"
        cmbYear.SetFocus
        Exit Sub
    End If

    ' === データ更新 ===
    ' 注: UserInterfaceOnly:=Trueで保護済みのため、VBAからの書き込みに保護解除は不要
    With Sh_data
        .Cells(eRowData.rowCategory, mTargetCol).value = Trim(cmbCategoryValue.Text)
        .Cells(eRowData.rowTestName, mTargetCol).value = Trim(txtTestName.Text)
        .Cells(eRowData.rowPerspective, mTargetCol).value = Trim(cmbPerspective.Text)
        .Cells(eRowData.rowDetail, mTargetCol).value = Trim(txtDetail.Text)
        .Cells(eRowData.rowAllocationScore, mTargetCol).value = CDbl(txtAllocateScore.Text)
        .Cells(eRowData.rowTestDate, mTargetCol).value = CDate(newDate)
    End With

    MsgBox "テスト情報を更新しました。", vbInformation, "更新完了"
    Me.Hide
End Sub

'===============================================================================
' キャンセルボタン
'===============================================================================
Private Sub btnCancel_Click()
    Me.Hide
End Sub

'===============================================================================
' 追試設定ボタン
'===============================================================================
Private Sub btnRetest_Click()
    Dim testKey As String
    testKey = Sh_data.Cells(eRowData.rowKey, mTargetCol).value

    ' 確認ダイアログ
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox( _
        "テスト「" & testKey & "」に追試を設定します。" & vbCrLf & _
        "データシートの得点は追試中マーカー(N)に置き換わります。" & vbCrLf & vbCrLf & _
        "実行しますか？", _
        vbQuestion + vbYesNo, "追試設定の確認")
    If confirmResult <> vbYes Then Exit Sub

    ' 追試ファイルに同キーのシートが既にあるか確認
    Dim existingSheet As Boolean
    existingSheet = RetestModule.HasRetestSheetForKey(testKey)

    If existingSheet Then
        MsgBox "追試シート（キー: " & testKey & "）が既に存在します。" & vbCrLf & _
               "既存のシートはそのまま残ります。処理を中止します。", _
               vbExclamation, "追試シート重複"
        Exit Sub
    End If

    ' 追試シートを作成（Sh_dataから得点取得）
    Call RetestModule.CreateRetestSheetFromData(mTargetCol)

    ' 追試設定済みに更新
    btnRetest.Caption = "追試設定済み"
    btnRetest.Enabled = False

    MsgBox "追試シートを作成しました。" & vbCrLf & _
           "データシートの得点は追試中マーカー(N)に置き換わりました。", _
           vbInformation, "追試設定完了"
End Sub

'===============================================================================
' Escキーでキャンセル
'===============================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' ×ボタンやEscキーで閉じる場合もHideに統一
    ' （呼び出し元でプロパティ読み取り後にUnloadする）
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

'===============================================================================
' 削除ボタン
' 備考: 追試中のテストは強制削除の確認を行う。
'       実際の削除処理は呼び出し元（Sh_data.Worksheet_BeforeDoubleClick）で実行。
'===============================================================================
Private Sub btnDelete_Click()
    ' 追試中チェック
    Dim firstChildVal As Variant
    firstChildVal = Sh_data.Cells(eRowData.rowChildStart, mTargetCol).value
    If CStr(firstChildVal) = RETEST_MARKER Then
        Dim confirmResult As VbMsgBoxResult
        confirmResult = MsgBox( _
            "このテストは追試中です。" & vbCrLf & vbCrLf & _
            "強制削除すると、追試ファイル内の対応する追試シートも" & vbCrLf & _
            "同時に削除されます。" & vbCrLf & vbCrLf & _
            "強制削除しますか？", _
            vbExclamation + vbYesNo, "追試中テストの強制削除")
        If confirmResult <> vbYes Then Exit Sub
        mForceDeleteRetest = True
    End If

    ' フラグを立ててフォームを閉じる
    ' ※ 実際の削除確認はDataManagementModule.DeleteTestData内で行われる
    mDeleteRequested = True
    Me.Hide
End Sub

'===============================================================================
' 削除リクエストされたかどうかを返す
'===============================================================================
Public Property Get DeleteRequested() As Boolean
    DeleteRequested = mDeleteRequested
End Property

'===============================================================================
' 追試中テストの強制削除がリクエストされたかどうかを返す
'===============================================================================
Public Property Get ForceDeleteRetest() As Boolean
    ForceDeleteRetest = mForceDeleteRetest
End Property

'===============================================================================
' テストキーを返す
'===============================================================================
Public Property Get testKey() As String
    testKey = lblKeyValue.Caption
End Property
