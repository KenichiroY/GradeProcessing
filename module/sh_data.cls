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
    If Target.Cells.count <> 1 Then Exit Sub

    ' D列（データ開始列）以降でなければ何もしない
    If Target.Column < eColData.colDataStart Then Exit Sub

    ' データがある列かチェック（キー行にデータがあるか）
    If IsEmpty(Me.Cells(eRowData.rowKey, Target.Column).value) Then Exit Sub

    ' === (A) 得点セル: 23行目以降 ===
    If Target.Row >= eRowData.rowChildStart Then
        ' 児童データがある行かチェック（コード列にデータがあるか）
        If IsEmpty(Me.Cells(Target.Row, eColData.colCode).value) Then Exit Sub

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
        delKey = frmEdit.testKey
        Unload frmEdit
        Set frmEdit = Nothing

        If deleteReq Then
            Call DataManagementModule.DeleteTestData(delKey, forceRetest)
        End If
    End If
End Sub

'===============================================================================
' 得点調整行の表示切替ボタン
'===============================================================================
Public Sub Btm_adjustscore_hide_reveal2()
    On Error Resume Next
    
    With Sh_data
        If .Range(RNG_SUBJECT_ADJSCORE_DISP).value = "表示" Then
            .Rows("12:16").Hidden = True
            .Range(RNG_SUBJECT_ADJSCORE_DISP).value = "非表示"
        Else
            .Rows("12:16").Hidden = False
            .Range(RNG_SUBJECT_ADJSCORE_DISP).value = "表示"
        End If
    End With
    
    On Error GoTo 0
End Sub
'===============================================================================
' 統計行の表示切替ボタン
' 説明: 19～22行（平均・中央値・標準偏差・変動係数）の表示/非表示を切り替える。
'       状態はB6セル（RNG_SUBJECT_STATS_DISP）で管理。
'===============================================================================
Public Sub Btn_stats_hide_reveal2()
    On Error Resume Next

    With Sh_data
        If .Range(RNG_SUBJECT_STATS_DISP).value = "表示" Then
            .Rows("19:22").Hidden = True
            .Range(RNG_SUBJECT_STATS_DISP).value = "非表示"
        Else
            .Rows("19:22").Hidden = False
            .Range(RNG_SUBJECT_STATS_DISP).value = "表示"
        End If
    End With

    On Error GoTo 0
End Sub

