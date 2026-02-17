VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sh_result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' シート名: sh_result (Resultシート)
' 機能: 評価結果の保存シート
' 備考: 列見出し生成はResultModule.GenerateResultHeadersで実行
'       （Workbook_Open時に自動生成、既存データがある場合はスキップ）
'===============================================================================
Option Explicit
