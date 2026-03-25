Attribute VB_Name = "FormatResultModule"
'===============================================================================
' モジュール名: FormatResultModule
' 用途: Resultシートのデザイン整形（一回限り実行用）
' 実行後は本モジュールを削除してよい
'===============================================================================
Option Explicit

Private Const COLOR_HEADER_DARK As Long = 9917184
Private Const COLOR_LABEL_ROW As Long = 15921906
Private Const COLOR_NAME_COL As Long = 16316664

' 教科別カラーパレット（観点行用・最大10教科、出現順で自動割当）
Private mSubjectColors(0 To 9) As Long
Private mSubjectLightColors(0 To 9) As Long
Private mSubjectNames() As String
Private mSubjectCount As Long
Private mPaletteInitialized As Boolean

Public Sub FormatResultSheet()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim ws As Worksheet
    Set ws = sh_result

    Dim lastCol As Long, lastRow As Long
    Dim childCount As Long

    ' RESULT_SUBJECT_ROW（8行目）は結合セルのためEnd(xlToLeft)が左端を返す
    ' RESULT_LABEL_ROW（10行目）は個別セルなので正確な最終列が取得できる
    lastCol = ws.Cells(RESULT_LABEL_ROW, Columns.count).End(xlToLeft).Column
    childCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value
    lastRow = RESULT_DATA_START_ROW + childCount - 1

    If lastCol < RESULT_DATA_START_COL Then
        MsgBox "Resultシートにデータがありません。", vbInformation
        GoTo CleanExit
    End If

    ' カラーパレット初期化（教科の出現順で色を割り当て）
    mPaletteInitialized = False
    Call InitColorPalette(ws, lastCol)

    ' 0. シート全体リセット
    ws.Cells.Interior.ColorIndex = xlNone
    ws.Cells.Borders.LineStyle = xlNone
    ws.Cells.Font.Bold = False
    ws.Cells.Font.Color = vbBlack
    ws.Cells.HorizontalAlignment = xlGeneral

    ' 1. 行1-7を非表示
    ws.Rows("1:7").Hidden = True

    ' 2. 教科ヘッダー（8行目）
    Call FormatSubjectHeaders(ws, lastCol)

    ' 3. 観点行（9行目）
    Call FormatPerspectiveRow(ws, lastCol)

    ' 4. ラベル行（10行目）
    Call FormatLabelRow(ws, lastCol)

    ' 5. 児童名列（A-C列）
    Call FormatNameColumns(ws, lastRow)

    ' 6. データ領域
    Call FormatDataArea(ws, lastCol, lastRow)

    ' 7. 罫線
    Call ApplyBorders(ws, lastCol, lastRow)

    ' 8. 列幅調整
    Call AdjustColumnWidths(ws, lastCol)

    ' 9. ウィンドウ枠固定
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(RESULT_DATA_START_ROW, RESULT_DATA_START_COL).Select
    ActiveWindow.FreezePanes = True

    ' 10. 印刷設定
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintTitleRows = "$8:$10"
        .PrintTitleColumns = "$A:$C"
    End With

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Sub FormatSubjectHeaders(ByVal ws As Worksheet, ByVal lastCol As Long)
    Dim col As Long, startCol As Long
    Dim currentSubject As String, prevSubject As String

    ws.Rows(RESULT_SUBJECT_ROW).MergeCells = False

    prevSubject = ""
    startCol = RESULT_DATA_START_COL

    For col = RESULT_DATA_START_COL To lastCol + 1
        If col <= lastCol Then
            currentSubject = ws.Cells(RESULT_SUBJECT_ROW, col).value & ""
        Else
            currentSubject = ""
        End If

        If currentSubject <> prevSubject And prevSubject <> "" Then
            If col - 1 > startCol Then
                ws.Range(ws.Cells(RESULT_SUBJECT_ROW, startCol), _
                         ws.Cells(RESULT_SUBJECT_ROW, col - 1)).Merge
            End If

            With ws.Range(ws.Cells(RESULT_SUBJECT_ROW, startCol), _
                          ws.Cells(RESULT_SUBJECT_ROW, col - 1))
                .Interior.Color = COLOR_HEADER_DARK
                .Font.Color = vbWhite
                .Font.Bold = True
                .Font.Size = 11
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

            startCol = col
        ElseIf prevSubject = "" And currentSubject <> "" Then
            startCol = col
        End If

        prevSubject = currentSubject
    Next col

    ws.Rows(RESULT_SUBJECT_ROW).RowHeight = 22
End Sub

Private Sub FormatPerspectiveRow(ByVal ws As Worksheet, ByVal lastCol As Long)
    Dim col As Long
    Dim subjectName As String

    For col = RESULT_DATA_START_COL To lastCol
        subjectName = ws.Cells(RESULT_SUBJECT_ROW, col).MergeArea.Cells(1, 1).value & ""

        With ws.Cells(RESULT_PERSPECTIVE_ROW, col)
            .Interior.Color = GetSubjectColor(subjectName)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next col

    ws.Rows(RESULT_PERSPECTIVE_ROW).RowHeight = 18
End Sub

Private Sub FormatLabelRow(ByVal ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(RESULT_LABEL_ROW, RESULT_DATA_START_COL), _
                  ws.Cells(RESULT_LABEL_ROW, lastCol))
        .Interior.Color = COLOR_LABEL_ROW
        .Font.Size = 8
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Rows(RESULT_LABEL_ROW).RowHeight = 16
End Sub

Private Sub FormatNameColumns(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range(ws.Cells(RESULT_SUBJECT_ROW, 1), ws.Cells(RESULT_LABEL_ROW, 3))
        .Interior.Color = COLOR_HEADER_DARK
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(RESULT_SUBJECT_ROW, 1).value = ""
    ws.Cells(RESULT_SUBJECT_ROW, 2).value = ""
    ws.Cells(RESULT_SUBJECT_ROW, 3).value = ""
    ws.Cells(RESULT_PERSPECTIVE_ROW, 1).value = ""
    ws.Cells(RESULT_PERSPECTIVE_ROW, 2).value = ""
    ws.Cells(RESULT_PERSPECTIVE_ROW, 3).value = ""
    ws.Cells(RESULT_LABEL_ROW, 1).value = "コード"
    ws.Cells(RESULT_LABEL_ROW, 2).value = "姓"
    ws.Cells(RESULT_LABEL_ROW, 3).value = "名"

    ws.Range(ws.Cells(RESULT_SUBJECT_ROW, 1), ws.Cells(RESULT_PERSPECTIVE_ROW, 1)).Merge
    ws.Range(ws.Cells(RESULT_SUBJECT_ROW, 2), ws.Cells(RESULT_PERSPECTIVE_ROW, 2)).Merge
    ws.Range(ws.Cells(RESULT_SUBJECT_ROW, 3), ws.Cells(RESULT_PERSPECTIVE_ROW, 3)).Merge

    If lastRow >= RESULT_DATA_START_ROW Then
        With ws.Range(ws.Cells(RESULT_DATA_START_ROW, 1), ws.Cells(lastRow, 3))
            .Interior.Color = COLOR_NAME_COL
            .Font.Size = 10
        End With
        ws.Range(ws.Cells(RESULT_DATA_START_ROW, 1), ws.Cells(lastRow, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(RESULT_DATA_START_ROW, 2), ws.Cells(lastRow, 3)).HorizontalAlignment = xlLeft
    End If
End Sub

Private Sub FormatDataArea(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal lastRow As Long)
    Dim col As Long
    Dim labelValue As String
    Dim rng As Range
    Dim cell As Range
    Dim subjectName As String
    Dim bgColor As Long
    Dim r As Range

    If lastRow < RESULT_DATA_START_ROW Then Exit Sub

    For col = RESULT_DATA_START_COL To lastCol
        labelValue = ws.Cells(RESULT_LABEL_ROW, col).value & ""
        Set rng = ws.Range(ws.Cells(RESULT_DATA_START_ROW, col), ws.Cells(lastRow, col))

        If labelValue = "ABC" Then
            rng.HorizontalAlignment = xlCenter
            rng.Font.Bold = True
            rng.Font.Size = 11

            For Each cell In rng
                Select Case cell.value & ""
                    Case "A"
                        cell.Font.Color = RGB(0, 120, 60)
                    Case "B"
                        cell.Font.Color = RGB(50, 50, 50)
                    Case "C"
                        cell.Font.Color = RGB(200, 50, 50)
                End Select
            Next cell
        Else
            rng.HorizontalAlignment = xlCenter
            rng.NumberFormat = "0.0"
            rng.Font.Size = 9
            rng.Font.Color = RGB(80, 80, 80)
        End If

        subjectName = ws.Cells(RESULT_SUBJECT_ROW, col).MergeArea.Cells(1, 1).value & ""
        bgColor = GetSubjectLightColor(subjectName)
        If bgColor > 0 Then
            For Each r In rng
                If r.Row Mod 2 = 0 Then
                    r.Interior.Color = bgColor
                End If
            Next r
        End If
    Next col

    Dim i As Long
    For i = RESULT_DATA_START_ROW To lastRow
        ws.Rows(i).RowHeight = 18
    Next i
End Sub

Private Sub ApplyBorders(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal lastRow As Long)
    Dim fullRange As Range
    Dim col As Long
    Dim subjectName As String, prevSubject As String

    If lastRow < RESULT_DATA_START_ROW Then lastRow = RESULT_LABEL_ROW

    Set fullRange = ws.Range(ws.Cells(RESULT_SUBJECT_ROW, 1), ws.Cells(lastRow, lastCol))

    With fullRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(100, 100, 100)
    End With
    With fullRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(100, 100, 100)
    End With
    With fullRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(100, 100, 100)
    End With
    With fullRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(100, 100, 100)
    End With

    With ws.Range(ws.Cells(RESULT_LABEL_ROW, 1), ws.Cells(RESULT_LABEL_ROW, lastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(100, 100, 100)
    End With

    With ws.Range(ws.Cells(RESULT_SUBJECT_ROW, 3), ws.Cells(lastRow, 3)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(100, 100, 100)
    End With

    If lastRow >= RESULT_DATA_START_ROW Then
        Dim dataRange As Range
        Set dataRange = ws.Range(ws.Cells(RESULT_DATA_START_ROW, 1), ws.Cells(lastRow, lastCol))
        With dataRange.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = RGB(200, 200, 200)
        End With
    End If

    prevSubject = ""
    For col = RESULT_DATA_START_COL To lastCol
        subjectName = ws.Cells(RESULT_SUBJECT_ROW, col).MergeArea.Cells(1, 1).value & ""
        If subjectName <> prevSubject And prevSubject <> "" Then
            With ws.Range(ws.Cells(RESULT_SUBJECT_ROW, col), ws.Cells(lastRow, col)).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(150, 150, 150)
            End With
        End If
        prevSubject = subjectName
    Next col
End Sub

Private Sub AdjustColumnWidths(ByVal ws As Worksheet, ByVal lastCol As Long)
    Dim col As Long
    Dim labelValue As String

    ws.Columns(1).ColumnWidth = 10
    ws.Columns(2).ColumnWidth = 6
    ws.Columns(3).ColumnWidth = 6

    For col = RESULT_DATA_START_COL To lastCol
        labelValue = ws.Cells(RESULT_LABEL_ROW, col).value & ""
        If labelValue = "ABC" Then
            ws.Columns(col).ColumnWidth = 4.5
        Else
            ws.Columns(col).ColumnWidth = 5.5
        End If
    Next col
End Sub

'===============================================================================
' カラーパレット初期化（Resultシートの教科出現順で色を割り当て）
'===============================================================================
Private Sub InitColorPalette(ByVal ws As Worksheet, ByVal lastCol As Long)
    If mPaletteInitialized Then Exit Sub

    ' 観点行用パレット（10色）
    mSubjectColors(0) = RGB(205, 225, 238)   ' 青系
    mSubjectColors(1) = RGB(215, 238, 219)   ' 緑系
    mSubjectColors(2) = RGB(255, 225, 210)   ' オレンジ系
    mSubjectColors(3) = RGB(253, 228, 227)   ' 赤系
    mSubjectColors(4) = RGB(232, 222, 238)   ' 紫系
    mSubjectColors(5) = RGB(255, 240, 215)   ' 黄系
    mSubjectColors(6) = RGB(210, 235, 235)   ' シアン系
    mSubjectColors(7) = RGB(238, 225, 215)   ' ベージュ系
    mSubjectColors(8) = RGB(225, 235, 210)   ' ライム系
    mSubjectColors(9) = RGB(235, 220, 230)   ' ピンク系

    ' データ領域の偶数行用パレット（より薄い色）
    mSubjectLightColors(0) = RGB(235, 243, 250)
    mSubjectLightColors(1) = RGB(238, 248, 240)
    mSubjectLightColors(2) = RGB(255, 244, 238)
    mSubjectLightColors(3) = RGB(253, 243, 243)
    mSubjectLightColors(4) = RGB(245, 240, 248)
    mSubjectLightColors(5) = RGB(255, 250, 238)
    mSubjectLightColors(6) = RGB(238, 248, 248)
    mSubjectLightColors(7) = RGB(248, 243, 238)
    mSubjectLightColors(8) = RGB(243, 248, 235)
    mSubjectLightColors(9) = RGB(248, 240, 245)

    ' Resultシートの教科名を出現順で収集
    mSubjectCount = 0
    ReDim mSubjectNames(0 To 9)

    Dim col As Long
    Dim subj As String
    Dim found As Boolean
    Dim j As Long

    For col = RESULT_DATA_START_COL To lastCol
        subj = ws.Cells(RESULT_SUBJECT_ROW, col).value & ""
        If subj = "" Then GoTo NextCol

        ' 既に登録済みか確認
        found = False
        For j = 0 To mSubjectCount - 1
            If mSubjectNames(j) = subj Then
                found = True
                Exit For
            End If
        Next j

        If Not found And mSubjectCount <= 9 Then
            mSubjectNames(mSubjectCount) = subj
            mSubjectCount = mSubjectCount + 1
        End If
NextCol:
    Next col

    mPaletteInitialized = True
End Sub

'===============================================================================
' 教科名からインデックスを返す（出現順、見つからなければ-1）
'===============================================================================
Private Function GetSubjectIndex(ByVal subjectName As String) As Long
    Dim j As Long
    For j = 0 To mSubjectCount - 1
        If mSubjectNames(j) = subjectName Then
            GetSubjectIndex = j
            Exit Function
        End If
    Next j
    GetSubjectIndex = -1
End Function

'===============================================================================
' 教科名から観点行の背景色を返す（出現順で自動割当）
'===============================================================================
Private Function GetSubjectColor(ByVal subjectName As String) As Long
    Dim idx As Long
    idx = GetSubjectIndex(subjectName)
    If idx >= 0 Then
        GetSubjectColor = mSubjectColors(idx Mod 10)
    Else
        GetSubjectColor = RGB(240, 240, 240)
    End If
End Function

'===============================================================================
' 教科名からデータ領域の偶数行背景色を返す（出現順で自動割当）
'===============================================================================
Private Function GetSubjectLightColor(ByVal subjectName As String) As Long
    Dim idx As Long
    idx = GetSubjectIndex(subjectName)
    If idx >= 0 Then
        GetSubjectLightColor = mSubjectLightColors(idx Mod 10)
    Else
        GetSubjectLightColor = 0
    End If
End Function
