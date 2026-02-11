'===============================================================================
' モジュール名: UIFormatModule
' 説明: 全シートのUI書式を一括設定する
'       手動で一度だけ実行する想定（書式はファイル保存時に保持される）
'       メインプロシージャ: ApplyAllSheetFormats
'===============================================================================
Option Explicit

' ─── 共通カラー定数（青系） ───
Private Const CLR_HEADER_BG As Long = 3912473       ' 濃紺 RGB(25, 55, 59) → 実際は下で定義
Private Const CLR_HEADER_FONT As Long = 16777215    ' 白
Private Const CLR_INPUT_BG As Long = 15790320       ' 淡青 RGB(208, 228, 241) → 実際は下で定義
Private Const CLR_SECTION_BG As Long = 14408667     ' やや濃い青 RGB(171, 205, 219)
Private Const CLR_ALT_ROW As Long = 15921906        ' 交互行色 RGB(242, 242, 242) 薄灰
Private Const CLR_BORDER As Long = 10921638         ' 罫線色 RGB(166, 166, 166) 灰

'===============================================================================
' 共通カラー取得関数（RGB値を正確に返す）
'===============================================================================
Private Function HeaderBgColor() As Long
    HeaderBgColor = RGB(31, 73, 125)    ' 濃紺
End Function

Private Function HeaderFontColor() As Long
    HeaderFontColor = RGB(255, 255, 255) ' 白
End Function

Private Function InputBgColor() As Long
    InputBgColor = RGB(218, 232, 247)    ' 淡青（入力欄）
End Function

Private Function SectionBgColor() As Long
    SectionBgColor = RGB(180, 210, 235)  ' やや濃い青（セクション見出し）
End Function

Private Function SubHeaderBgColor() As Long
    SubHeaderBgColor = RGB(141, 180, 226) ' サブヘッダー
End Function

Private Function AltRowColor() As Long
    AltRowColor = RGB(242, 246, 252)     ' 交互行（うっすら青）
End Function

Private Function BorderColor() As Long
    BorderColor = RGB(166, 176, 192)     ' 罫線色（青灰）
End Function

Private Function WarningBgColor() As Long
    WarningBgColor = RGB(255, 235, 156)  ' 警告色（淡黄）
End Function

Private Function SuccessBgColor() As Long
    SuccessBgColor = RGB(198, 239, 206)  ' 成功色（淡緑）
End Function

'===============================================================================
' 全シートの書式を一括適用（マスタープロシージャ）
' 説明：手動で一度だけ実行する。Alt+F8 → ApplyAllSheetFormats
'===============================================================================
Public Sub ApplyAllSheetFormats()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call FormatMenuSheet
    Call FormatNamelistSheet
    Call FormatDataSheet
    Call FormatSubjectSheet
    Call FormatRetestTemplateSheet
    Call SetSheetTabColors

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "全シートの書式設定が完了しました。", vbInformation, "書式設定完了"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "書式設定中にエラーが発生しました。" & vbCrLf & _
           Err.Description, vbCritical, "エラー"
End Sub

'===============================================================================
' MENUシートの書式設定
' 改善内容：
'   - タイトルバンド化（濃紺背景+白文字）
'   - ヘッダー行の配色
'   - 点数入力欄（I列）に淡青背景
'   - 転記先行/列（K,L列）を非表示
'   - 未入力一覧のエリア分離
'===============================================================================
Public Sub FormatMenuSheet()
    Dim ws As Worksheet
    Set ws = sh_MENU

    With ws
        ' === 全体設定 ===
        .Cells.Font.Name = "游ゴシック"
        .Cells.Font.Size = 10

        ' === タイトル行（1行目）===
        .Range("A1:M1").Interior.Color = HeaderBgColor()
        .Range("A1").Font.Color = HeaderFontColor()
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True

        ' === ヘッダー行（10行目）===
        Dim headerRange As Range
        Set headerRange = .Range("B10:L10")
        With headerRange
            .Interior.Color = HeaderBgColor()
            .Font.Color = HeaderFontColor()
            .Font.Bold = True
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
        End With

        ' === ヘッダー罫線 ===
        Call SetThinBorders(headerRange, BorderColor())

        ' === 点数入力列（I列）のヘッダーを目立たせる ===
        .Cells(10, eColMenu.colScore).Interior.Color = WarningBgColor()
        .Cells(10, eColMenu.colScore).Font.Color = RGB(0, 0, 0)

        ' === データエリアの書式（動的部分は FormatMenuDataArea で適用） ===
        Call FormatMenuDataArea

        ' === 列幅調整 ===
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 12    ' コード
        .Columns("C").ColumnWidth = 8     ' 姓
        .Columns("D").ColumnWidth = 8     ' 名
        .Columns("E").ColumnWidth = 8     ' 教科
        .Columns("F").ColumnWidth = 10    ' 観点
        .Columns("G").ColumnWidth = 18    ' テスト名
        .Columns("H").ColumnWidth = 12    ' 詳細
        .Columns("I").ColumnWidth = 8     ' 点数
        .Columns("J").ColumnWidth = 8     ' 配点

        ' === 転記先行/列を非表示 ===
        .Columns("K:L").Hidden = True

        ' === 区切り線（ボタンエリアとデータエリアの間） ===
        .Range("A9:L9").Borders(xlEdgeBottom).Color = HeaderBgColor()
        .Range("A9:L9").Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

'===============================================================================
' MENUシートのデータエリア書式設定（動的）
' 説明：
'   SearchNotYetInput / SearchNotYetByTest でデータ生成後に呼び出す
'   実際のデータ行数に応じて罫線・入力欄色を設定する
'   ClearMenuData時にも呼ばれ、書式をリセットする
'===============================================================================
Public Sub FormatMenuDataArea()
    Dim ws As Worksheet
    Set ws = sh_MENU

    With ws
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, eColMenu.colCode).End(xlUp).Row

        ' データがない場合は11行目以降の書式をクリアして終了
        If lastRow < eRowMenu.rowStart Then
            Dim clearRange As Range
            Set clearRange = .Range(.Cells(eRowMenu.rowStart, 2), _
                                     .Cells(eRowMenu.rowStart + MAX_CHILDREN, 12))
            clearRange.Interior.ColorIndex = xlColorIndexNone
            clearRange.Borders.LineStyle = xlLineStyleNone
            Exit Sub
        End If

        ' データ範囲
        Dim dataRange As Range
        Set dataRange = .Range(.Cells(eRowMenu.rowStart, eColMenu.colCode), _
                                .Cells(lastRow, eColMenu.colToCol))

        ' 罫線
        Call SetThinBorders(dataRange, BorderColor())

        ' 点数入力列（I列）に入力欄色
        .Range(.Cells(eRowMenu.rowStart, eColMenu.colScore), _
               .Cells(lastRow, eColMenu.colScore)).Interior.Color = InputBgColor()

        ' データ行以降の余白をクリア（前回のデータが多かった場合）
        If lastRow + 1 <= lastRow + MAX_CHILDREN Then
            Dim excessRange As Range
            Set excessRange = .Range(.Cells(lastRow + 1, 2), _
                                      .Cells(lastRow + MAX_CHILDREN, 12))
            excessRange.Interior.ColorIndex = xlColorIndexNone
            excessRange.Borders.LineStyle = xlLineStyleNone
        End If
    End With
End Sub

'===============================================================================
' 名簿シートの書式設定
' 改善内容：
'   - タイトルバンド化（濃紺背景+白文字）
'   - 児童数表示（E8）のラベル強調
'   - ヘッダー行（10行目）の配色
'   - 児童データエリアの罫線・交互行色
'===============================================================================
Public Sub FormatNamelistSheet()
    Dim ws As Worksheet
    Set ws = sh_namelist

    With ws
        ' === 全体設定 ===
        .Cells.Font.Name = "游ゴシック"
        .Cells.Font.Size = 10

        ' === タイトル行（1行目）===
        .Range("A1:F1").Interior.Color = HeaderBgColor()
        .Range("A1").Font.Color = HeaderFontColor()
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True

        ' === 児童数エリア（E8付近）===
        .Range("D8").Font.Bold = True
        .Range("D8").Font.Size = 10
        .Range("E8").Interior.Color = InputBgColor()
        .Range("E8").Font.Size = 12
        .Range("E8").Font.Bold = True
        .Range("E8").HorizontalAlignment = xlCenter

        ' === ヘッダー行（10行目）===
        Dim headerRange As Range
        Set headerRange = .Range("A" & NAMELIST_HEADER_ROW & ":F" & NAMELIST_HEADER_ROW)
        With headerRange
            .Interior.Color = HeaderBgColor()
            .Font.Color = HeaderFontColor()
            .Font.Bold = True
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
        End With
        Call SetThinBorders(headerRange, BorderColor())

        ' === 児童データエリアの罫線（11行目～） ===
        Dim childCount As Long
        childCount = ws.Range(RNG_NAMELIST_CHILDCOUNT).Value
        If childCount > 0 Then
            Dim dataRange As Range
            Set dataRange = .Range(.Cells(NAMELIST_DATA_START_ROW, 1), _
                                    .Cells(NAMELIST_DATA_START_ROW + childCount - 1, 6))
            Call SetThinBorders(dataRange, BorderColor())

            ' 交互行色
            Dim r As Long
            For r = NAMELIST_DATA_START_ROW To NAMELIST_DATA_START_ROW + childCount - 1
                If (r - NAMELIST_DATA_START_ROW) Mod 2 = 1 Then
                    .Range(.Cells(r, 1), .Cells(r, 6)).Interior.Color = AltRowColor()
                End If
            Next r
        End If

        ' === 列幅 ===
        .Columns("A").ColumnWidth = 12   ' コード
        .Columns("B").ColumnWidth = 8    ' 姓
        .Columns("C").ColumnWidth = 8    ' 名
    End With
End Sub

'===============================================================================
' データシートの書式設定
' 改善内容：
'   - タイトル行を1行に圧縮（既存値保持）
'   - A列の重複ラベルを削除（C列のみに統一）
'   - 行カテゴリ別の色分け（基本情報/配点・調整/統計）
'   - エラー値（#DIV/0! #NUM!等）を条件付き書式で非表示
'   - データ列の最小列幅を確保（###表示防止）
'   - 児童データ開始行の上に太い境界線
'===============================================================================
Public Sub FormatDataSheet()
    Dim ws As Worksheet
    Set ws = Sh_data

    With ws
        ' === 全体設定 ===
        .Cells.Font.Name = "游ゴシック"
        .Cells.Font.Size = 10

        ' === タイトル行（1行目のみ）：値は書き換えず書式のみ ===
        .Range("A1:C1").Interior.Color = HeaderBgColor()
        .Range("A1").Font.Color = HeaderFontColor()
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        ' 2行目はバンドを外す（ボタン配置用の余白として残す）
        .Range("A2:C2").Interior.ColorIndex = xlColorIndexNone

        ' === 操作説明テキスト（3行目）===
        .Range("A3").Value = "※ 得点セルをダブルクリックで修正できます"
        .Range("A3").Font.Size = 9
        .Range("A3").Font.Color = RGB(100, 100, 100)
        .Range("A3").Font.Italic = True

        ' === A列の重複ラベルを削除（C列に既にラベルがあるため）===
        Dim r As Long
        For r = eRowData.rowKey To eRowData.rowCV
            .Cells(r, 1).Value = ""
        Next r

        ' === 行カテゴリ別の色分け ===
        ' (1) 基本情報（4-10行：キー～詳細）→ 濃い目の青
        .Range("A" & eRowData.rowKey & ":C" & eRowData.rowDetail).Interior.Color = SubHeaderBgColor()
        .Range("A" & eRowData.rowKey & ":C" & eRowData.rowDetail).Font.Bold = True
        .Range("A" & eRowData.rowKey & ":C" & eRowData.rowDetail).Font.Size = 9

        ' (2) 配点・調整（11-18行：配点～重み）→ やや薄い青
        .Range("A" & eRowData.rowAllocationScore & ":C" & eRowData.rowWeight).Interior.Color = SectionBgColor()
        .Range("A" & eRowData.rowAllocationScore & ":C" & eRowData.rowWeight).Font.Bold = True
        .Range("A" & eRowData.rowAllocationScore & ":C" & eRowData.rowWeight).Font.Size = 9

        ' (3) 統計値（19-22行：平均～変動係数）→ さらに薄い青
        .Range("A" & eRowData.rowAverage & ":C" & eRowData.rowCV).Interior.Color = InputBgColor()
        .Range("A" & eRowData.rowAverage & ":C" & eRowData.rowCV).Font.Bold = True
        .Range("A" & eRowData.rowAverage & ":C" & eRowData.rowCV).Font.Size = 9

        ' === 児童データエリアの上部境界線（太線）===
        .Range("A" & eRowData.rowCV & ":C" & eRowData.rowCV).Borders(xlEdgeBottom).Color = HeaderBgColor()
        .Range("A" & eRowData.rowCV & ":C" & eRowData.rowCV).Borders(xlEdgeBottom).Weight = xlMedium

        ' === 得点調整行（12-16行）のグループ化（折りたたみ可能に）===
        On Error Resume Next
        .Rows("12:16").Ungroup  ' 既存グループがあれば解除
        On Error GoTo 0
        .Rows("12:16").Group
        .Outline.ShowLevels RowLevels:=1  ' 初期状態は折りたたみ

        ' === エラー値の非表示（条件付き書式：エラーセルのフォントを白に）===
        Dim lastCol As Long
        lastCol = .Cells(eRowData.rowKey, .Columns.Count).End(xlToLeft).Column
        If lastCol >= eColData.colDataStart Then
            ' 統計行（平均～変動係数）のデータ列範囲に条件付き書式を設定
            Dim statsRange As Range
            Set statsRange = .Range( _
                .Cells(eRowData.rowAverage, eColData.colDataStart), _
                .Cells(eRowData.rowCV, lastCol))
            statsRange.FormatConditions.Delete
            With statsRange.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=ISERROR(" & .Cells(eRowData.rowAverage, eColData.colDataStart).Address(False, False) & ")")
                .Font.Color = RGB(255, 255, 255)  ' 白フォント
            End With
        End If

        ' === 列幅 ===
        .Columns("A").ColumnWidth = 12   ' コード
        .Columns("B").ColumnWidth = 8    ' 姓
        .Columns("C").ColumnWidth = 12   ' ラベル列（ラベルが切れないよう広めに）

        ' === データ列の最小列幅を確保（###表示防止）===
        If lastCol >= eColData.colDataStart Then
            Dim c As Long
            For c = eColData.colDataStart To lastCol
                If .Columns(c).ColumnWidth < 6 Then
                    .Columns(c).ColumnWidth = 6
                End If
            Next c
        End If

        ' === 追試中列のオレンジ色表示を再適用 ===
        If lastCol >= eColData.colDataStart Then
            Dim retestCol As Long
            For retestCol = eColData.colDataStart To lastCol
                If CStr(.Cells(eRowData.rowChildStart, retestCol).Value) = RETEST_MARKER Then
                    Call ApplyRetestColumnFormat(retestCol)
                End If
            Next retestCol
        End If
    End With
End Sub

'===============================================================================
' Subjectシートの書式設定
' 改善内容：
'   - ABC評価の●/★操作説明を追加
'   - 「最終決定」セルを目立たせる条件付き書式
'   - ボタンエリアとデータエリアの視覚的分離
'===============================================================================
Public Sub FormatSubjectSheet()
    Dim ws As Worksheet
    Set ws = sh_subject

    With ws
        ' === 全体設定 ===
        .Cells.Font.Name = "游ゴシック"
        .Cells.Font.Size = 10

        ' === タイトル（教科名表示エリア） ===
        .Range("A1:B1").Interior.Color = HeaderBgColor()
        .Range("A1").Font.Color = HeaderFontColor()
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True

        ' === 教科名セル（B2）===
        .Range("B2").Interior.Color = InputBgColor()
        .Range("B2").Font.Size = 12
        .Range("B2").Font.Bold = True

        ' === 設定表示エリア（B4-B6） ===
        .Range("A4:B6").Interior.Color = RGB(240, 240, 248)
        .Range("A4:B6").Font.Size = 9

        ' === ABC評価の操作説明（7行目付近、データの右側に） ===
        ' 操作説明はヘッダーエリアに追加
        .Range("A7").Value = "●をダブルクリック→★に採用"
        .Range("A7").Font.Size = 9
        .Range("A7").Font.Color = RGB(100, 100, 100)
        .Range("A7").Font.Italic = True

        ' === ヘッダーラベル列（A列、キー～変動係数） ===
        .Range("A" & eRowData.rowKey & ":C" & eRowData.rowCV).Interior.Color = SectionBgColor()
        .Range("A" & eRowData.rowKey & ":C" & eRowData.rowCV).Font.Size = 9
        .Range("A" & eRowData.rowKey & ":C" & eRowData.rowCV).Font.Bold = True

        ' === 児童データエリアの上部境界線 ===
        .Range("A22:C22").Borders(xlEdgeBottom).Color = HeaderBgColor()
        .Range("A22:C22").Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

'===============================================================================
' 追試テンプレートシートの書式設定
' 改善内容：
'   - 上部情報エリア（3-7行）のゾーニング
'   - 状態セルの条件付き書式（追試中=黄、反映済み=緑）
'   - ヘッダー行の配色
'===============================================================================
Public Sub FormatRetestTemplateSheet()
    Dim ws As Worksheet

    ' テンプレートシートをCodeNameで検索
    Dim wsTemplate As Worksheet
    For Each wsTemplate In ThisWorkbook.Worksheets
        If wsTemplate.CodeName = "sh_rt_template" Then
            Set ws = wsTemplate
            Exit For
        End If
    Next wsTemplate

    If ws Is Nothing Then
        ' テンプレートが見つからない場合はスキップ
        Exit Sub
    End If

    ' 一時的にVisibleにする
    Dim originalVisible As XlSheetVisibility
    originalVisible = ws.Visible
    ws.Visible = xlSheetVisible

    With ws
        ' === 全体設定 ===
        .Cells.Font.Name = "游ゴシック"
        .Cells.Font.Size = 10

        ' === タイトル行 ===
        .Range("A1:H1").Interior.Color = HeaderBgColor()
        .Range("A1").Font.Color = HeaderFontColor()
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True

        ' === 情報エリア（3-7行、A-B列）===
        .Range("A3:B7").Interior.Color = RGB(240, 240, 248)
        .Range("A3:A7").Font.Bold = True
        .Range("A3:A7").Font.Size = 10

        ' === 算出設定エリア（3-7行、D-E列）===
        .Range("D3:D7").Interior.Color = RGB(240, 240, 248)
        .Range("D3:D7").Font.Bold = True

        ' === 合格者数エリア（G3:H4）===
        .Range("G3:G4").Font.Size = 9
        .Range("H3:H4").Font.Bold = True

        ' === 状態セル（E7）の条件付き書式 ===
        .Range(RNG_RT_STATUS).FormatConditions.Delete
        ' 追試中 → 黄色背景
        With .Range(RNG_RT_STATUS).FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""追試中""")
            .Interior.Color = WarningBgColor()
            .Font.Bold = True
        End With
        ' 反映済み → 緑背景
        With .Range(RNG_RT_STATUS).FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""反映済み""")
            .Interior.Color = SuccessBgColor()
            .Font.Bold = True
        End With

        ' === ヘッダー行（10行目）===
        .Range("A" & RT_HEADER_ROW & ":H" & RT_HEADER_ROW).Interior.Color = HeaderBgColor()
        .Range("A" & RT_HEADER_ROW & ":H" & RT_HEADER_ROW).Font.Color = HeaderFontColor()
        .Range("A" & RT_HEADER_ROW & ":H" & RT_HEADER_ROW).Font.Bold = True
        .Range("A" & RT_HEADER_ROW & ":H" & RT_HEADER_ROW).HorizontalAlignment = xlCenter

        ' === 列幅 ===
        .Columns("A").ColumnWidth = 12   ' コード
        .Columns("B").ColumnWidth = 8    ' 姓
        .Columns("C").ColumnWidth = 8    ' 名
        .Columns("D").ColumnWidth = 10   ' 本試/算出方法ラベル
        .Columns("E").ColumnWidth = 12   ' 追試/算出設定値
    End With

    ' 元の表示状態に戻す
    ws.Visible = originalVisible
End Sub

'===============================================================================
' シートタブの色設定
' カテゴリ別に色分け：
'   入力系（MENU, テスト入力, 名簿）: 青系
'   データ系（データ, Subject, Result）: 緑系
'   設定系（Setting）: 灰色
'   テンプレート: 色なし（非表示のため）
'===============================================================================
Public Sub SetSheetTabColors()
    ' 入力・操作系 → 青
    sh_MENU.Tab.Color = RGB(68, 114, 196)
    sh_input.Tab.Color = RGB(68, 114, 196)
    sh_namelist.Tab.Color = RGB(68, 114, 196)

    ' データ・集計系 → 緑
    Sh_data.Tab.Color = RGB(84, 130, 53)
    sh_subject.Tab.Color = RGB(84, 130, 53)
    sh_result.Tab.Color = RGB(84, 130, 53)

    ' 設定系 → 灰
    sh_setting.Tab.Color = RGB(166, 166, 166)

    ' 個人分析（将来）→ 色なし
    On Error Resume Next
    sh_individual.Tab.ColorIndex = xlColorIndexNone
    On Error GoTo 0
End Sub

'===============================================================================
' 罫線設定ヘルパー
' 引数：rng - 対象範囲、clr - 罫線色
'===============================================================================
Private Sub SetThinBorders(ByVal rng As Range, ByVal clr As Long)
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        On Error Resume Next
        With rng.Borders(edge)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = clr
        End With
        On Error GoTo 0
    Next edge
End Sub

'===============================================================================
' 追試中列にオレンジ色のフォーマットを適用
' 説明: 追試中マーカー"N"が入った列のヘッダー行と得点セルを
'       オレンジ系の背景色で強調表示する
' 引数: targetCol - データシートの対象列番号
'===============================================================================
Public Sub ApplyRetestColumnFormat(ByVal targetCol As Long)
    Dim lastRow As Long

    With Sh_data
        lastRow = .Cells(.Rows.Count, eColData.colCode).End(xlUp).Row
        If lastRow < eRowData.rowChildStart Then lastRow = eRowData.rowChildStart

        ' ヘッダー行（4-22行）にオレンジ背景
        .Range(.Cells(eRowData.rowKey, targetCol), _
               .Cells(eRowData.rowCV, targetCol)).Interior.Color = COLOR_RETEST_HEADER

        ' 児童データ行（23行～最終行）に薄オレンジ背景
        .Range(.Cells(eRowData.rowChildStart, targetCol), _
               .Cells(lastRow, targetCol)).Interior.Color = COLOR_RETEST_CELL

        ' "N"セルのフォントを太字・濃いオレンジ色に
        Dim j As Long
        For j = eRowData.rowChildStart To lastRow
            If CStr(.Cells(j, targetCol).Value) = RETEST_MARKER Then
                .Cells(j, targetCol).Font.Bold = True
                .Cells(j, targetCol).Font.Color = RGB(200, 100, 0)  ' 濃いオレンジ
            End If
        Next j
    End With
End Sub

'===============================================================================
' 追試中列のオレンジ色フォーマットをクリアし通常色に戻す
' 説明: 追試完了時に呼び出し、ヘッダー行は行帯ごとの元の色に復元し、
'       得点セルの背景色とフォント装飾をリセットする
' 引数: targetCol - データシートの対象列番号
'===============================================================================
Public Sub ClearRetestColumnFormat(ByVal targetCol As Long)
    Dim lastRow As Long

    With Sh_data
        lastRow = .Cells(.Rows.Count, eColData.colCode).End(xlUp).Row
        If lastRow < eRowData.rowChildStart Then lastRow = eRowData.rowChildStart

        ' ヘッダー行の背景色を行帯ごとに復元
        ' (1) 基本情報（4-10行）→ SubHeaderBgColor
        .Range(.Cells(eRowData.rowKey, targetCol), _
               .Cells(eRowData.rowDetail, targetCol)).Interior.Color = SubHeaderBgColor()

        ' (2) 配点・調整（11-18行）→ SectionBgColor
        .Range(.Cells(eRowData.rowAllocationScore, targetCol), _
               .Cells(eRowData.rowWeight, targetCol)).Interior.Color = SectionBgColor()

        ' (3) 統計値（19-22行）→ InputBgColor
        .Range(.Cells(eRowData.rowAverage, targetCol), _
               .Cells(eRowData.rowCV, targetCol)).Interior.Color = InputBgColor()

        ' 児童データ行の背景色をクリア
        .Range(.Cells(eRowData.rowChildStart, targetCol), _
               .Cells(lastRow, targetCol)).Interior.ColorIndex = xlColorIndexNone

        ' フォント装飾をリセット
        Dim j As Long
        For j = eRowData.rowChildStart To lastRow
            .Cells(j, targetCol).Font.Bold = False
            .Cells(j, targetCol).Font.Color = RGB(0, 0, 0)  ' 黒に戻す
        Next j
    End With
End Sub
