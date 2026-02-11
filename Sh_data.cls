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
'   - 得点セル（23行目以降）のダブルクリックで修正フォーム(frmScoreEdit)を表示
'   - テスト情報ヘッダー（4-22行目）のダブルクリックで編集フォーム(frmTestEdit)を表示
'===============================================================================
Option Explicit

'===============================================================================
' ダブルクリックイベント
' 説明:
'   (A) 得点セル（23行目以降、D列以降）→ frmScoreEdit（得点修正）
'   (B) テスト情報ヘッダー（4-22行目、D列以降）→ frmTestEdit（テスト情報編集+追試設定）
'===============================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' 単一セルのみ対象
    If Target.Cells.Count <> 1 Then Exit Sub

    ' D列（データ開始列）以降でなければ何もしない
    If Target.Column < eColData.colDataStart Then Exit Sub

    ' データがある列かチェック（キー行にデータがあるか）
    If IsEmpty(Me.Cells(eRowData.rowKey, Target.Column).Value) Then Exit Sub

    ' === (A) 得点セル: 23行目以降 ===
    If Target.Row >= eRowData.rowChildStart Then
        ' 児童データがある行かチェック（コード列にデータがあるか）
        If IsEmpty(Me.Cells(Target.Row, eColData.colCode).Value) Then Exit Sub

        Cancel = True

        Dim frmScore As frmScoreEdit
        Set frmScore = New frmScoreEdit
        frmScore.Initialize Target.Row, Target.Column
        frmScore.Show vbModal
        Set frmScore = Nothing

    ' === (B) テスト情報ヘッダー: 4-22行目 ===
    ElseIf Target.Row >= eRowData.rowKey And Target.Row <= eRowData.rowCV Then
        Cancel = True

        Dim frmEdit As frmTestEdit
        Set frmEdit = New frmTestEdit
        frmEdit.Initialize Target.Column
        frmEdit.Show vbModal

        ' 削除リクエストの判定（Hide後にプロパティ読み取り→Unload）
        Dim deleteReq As Boolean
        Dim forceRetest As Boolean
        Dim delKey As String
        deleteReq = frmEdit.DeleteRequested
        forceRetest = frmEdit.ForceDeleteRetest
        delKey = frmEdit.TestKey
        Unload frmEdit
        Set frmEdit = Nothing

        If deleteReq Then
            Call DataManagementModule.DeleteTestData(delKey, forceRetest)
        End If
    End If
End Sub
