VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sh_namelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' 登録ボタンのクリックイベント
'===============================================================================
Public Sub Btn_Posting_Click()
    Call PostingModule.Posting
End Sub

'===============================================================================
' 未入力検索ボタンのクリックイベント
'===============================================================================
Public Sub SearchNotYetButton_Click()
    Call historycheckModule.SearchNotYetInput
End Sub

'===============================================================================
' 一括転記ボタンのクリックイベント
'===============================================================================
Public Sub TransferButton_Click()
    Call historycheckModule.TransferFromMenu
End Sub
'===============================================================================
' 児童コード自動採番
' 説明: 名簿の姓（C列）が入力されている行に対して、児童コードを自動採番する。
'       コード形式: "1001" & 4桁連番（10010001, 10010002, ...）
'===============================================================================
Public Sub Saiban()
    On Error GoTo ErrorHandler

    ' 既存コードの確認
    If sh_namelist.Cells(NAMELIST_DATA_START_ROW, 2) <> "" Then
        If MsgBox("既にコードが入っているようです。" & vbCrLf & _
                  "これらを消して自動的に採番しますか？", _
                  vbInformation + vbYesNo, "児童コード採番") = vbNo Then
            Exit Sub
        End If
    End If

    Dim i As Long
    i = 0
    Do While sh_namelist.Cells(i + NAMELIST_DATA_START_ROW, 3) <> "" _
             And i < MAX_CHILDREN
        sh_namelist.Cells(i + NAMELIST_DATA_START_ROW, 2) = "1001" & Format(i + 1, "0000")
        i = i + 1
    Loop

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("sh_namelist", "Saiban")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub
'===============================================================================
' 在籍開始日一括設定
' 説明: 名簿の姓（C列）が入力されている行に対して、在籍開始日（E列）に
'       今年度の4月1日を一括設定する。1～3月実行時は前年の4月1日となる。
'===============================================================================
Public Sub SetEnrollmentStartDate()
    On Error GoTo ErrorHandler

    ' 既存データの確認
    If sh_namelist.Cells(NAMELIST_DATA_START_ROW, 5) <> "" Then
        If MsgBox("既に在籍開始日が入っているようです。" & vbCrLf & _
                  "これらを消して自動的に今年度の4月1日を設定しますか？", _
                  vbInformation + vbYesNo, "在籍開始日設定") = vbNo Then
            Exit Sub
        End If
    End If

    ' 今年度の4月1日を算出（1～3月は前年度扱い）
    Dim fiscalYear As Long
    fiscalYear = Year(Date) - IIf(Month(Date) <= 3, 1, 0)
    Dim startDate As Date
    startDate = DateSerial(fiscalYear, 4, 1)

    Dim i As Long
    i = 0
    Do While sh_namelist.Cells(i + NAMELIST_DATA_START_ROW, 3) <> "" _
             And i < MAX_CHILDREN
        sh_namelist.Cells(i + NAMELIST_DATA_START_ROW, 5) = startDate
        i = i + 1
    Loop

    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("sh_namelist", "SetEnrollmentStartDate")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub


