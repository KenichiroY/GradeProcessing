VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sh_rt_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' モジュール名: sh_rt_menu (追試管理メニュー テンプレート)
' 説明: 追試ファイルのMENUシートのイベント処理
'       テンプレートコピー時にシートモジュールのコードも引き継がれるため、
'       追試ファイル内でもダブルクリックイベントが動作する。
'       ※ 追試ファイルでは本体の定数モジュールにアクセスできないため、
'         定数はリテラル値を使用する。
'===============================================================================
Option Explicit

'===============================================================================
' ダブルクリックイベント
' 説明: データ行（4行目以降）をダブルクリックすると、
'       G列のシート名に対応する追試シートを表示する。
'===============================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' 単一セルのみ対象
    If Target.Cells.count <> 1 Then Exit Sub

    ' データ行（4行目以降）でなければ何もしない
    If Target.Row < 4 Then Exit Sub

    ' G列（7列目）からシート名を取得
    Dim sheetName As String
    sheetName = Trim(Me.Cells(Target.Row, 7).value & "")
    If sheetName = "" Then Exit Sub

    Cancel = True

    ' 同じブック内のシートに移動
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Me.Parent.sheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Activate
    Else
        MsgBox "シート「" & sheetName & "」が見つかりません。", vbExclamation, "追試管理メニュー"
    End If
End Sub

