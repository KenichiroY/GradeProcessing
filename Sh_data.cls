VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sh_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' シート名: Sh_data（データシート）
' 説明: テストデータ蓄積シートのイベント処理
' 機能:
'   - 得点セルのダブルクリックで修正フォームを表示
'===============================================================================
Option Explicit

'===============================================================================
' ダブルクリックイベント
' 説明: 得点セルをダブルクリックすると修正フォームを表示
'===============================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' 単一セルのみ対象
    If Target.Cells.Count <> 1 Then Exit Sub

    ' 得点セルの範囲かチェック（児童データ開始行以降、データ開始列以降）
    If Target.Row < eRowData.rowChildStart Then Exit Sub
    If Target.Column < eColData.colDataStart Then Exit Sub

    ' データがある列かチェック（キー行にデータがあるか）
    If IsEmpty(Me.Cells(eRowData.rowKey, Target.Column).Value) Then Exit Sub

    ' 児童データがある行かチェック（コード列にデータがあるか）
    If IsEmpty(Me.Cells(Target.Row, eColData.colCode).Value) Then Exit Sub

    ' デフォルトの編集動作をキャンセル
    Cancel = True

    ' 修正フォームを表示
    Dim frm As frmScoreEdit
    Set frm = New frmScoreEdit
    frm.Initialize Target.Row, Target.Column
    frm.Show vbModal
    Set frm = Nothing
End Sub
