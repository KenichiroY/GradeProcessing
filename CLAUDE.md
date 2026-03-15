# 成績処理ソフト

## 概要

小学校向けテスト成績管理・ABC評価システム。Excel VBA (xlsm) で実装されており、テスト得点の登録、教科別集計、ABC評価の自動計算を行う。

## 主要ファイル

- `成績処理.xlsm` - メインのExcelマクロファイル（すべてのVBAコードを含む）
- `*.bas`, `*.cls`, `*.frm` - VBAモジュールのエクスポートファイル（xlsmと重複、参照用）

## シート構成

| シート名 | CodeName | 用途 | 主要セル範囲 |
|----------|----------|------|--------------|
| MENU | sh_MENU | 未入力データ一覧・一括入力 | 11行目～:未入力一覧 |
| 名簿 | sh_namelist | 児童名簿管理 | F8:児童数, 11行目～:児童データ |
| テスト入力 | sh_input | テスト得点入力フォーム | D4:教科, F4:カテゴリ, J4:日付, D6:テスト名, 31行目～:児童得点 |
| データ | Sh_data | テストデータ蓄積（自動更新） | 4行目:キー, 23行目～:児童得点 |
| Subject | sh_subject | 教科別集計・ABC評価 | B2:教科名, B4:得点調整有効/無効, B5:得点調整行表示状態, B6:統計行表示状態, B7:重み正規化状態 |
| Result | sh_result | 評価結果の保存 | 8行目:教科名, 9行目:観点, 10行目:ラベル, 11行目～:児童データ |
| IndividualAnalysis | sh_individual | テスト比較分析（箱ひげ図・散布図・T検定） | 9行目:キー, 10行目:日付, 11行目:教科, 12行目:テスト名, 13行目:観点, 14行目:配点, 15行目～:児童得点 |
| Setting | sh_setting | 教科・観点・カテゴリ・ABC閾値の設定 | A列:教科文字, B列:教科, C列:キー最終値, D列:観点, E列:学期, F列:カテゴリ, H-I列:ABC閾値 |
| RT_MENU（テンプレート） | sh_rt_menu | 追試ファイルMENUのテンプレート（VeryHidden） | - |
| RT_TEMPLATE（テンプレート） | sh_rt_template | 追試シートのテンプレート（VeryHidden） | - |

## VBAモジュール構成

### 標準モジュール

| モジュール | 用途 | 主要関数/プロシージャ |
|------------|------|----------------------|
| Module1 | ワークシート関数 ※名前変更禁止 | `調整後配点計算`, `調整後得点計算`, `調整後得点計算_shsubject`, `得点変換` |
| PublicConstListModule | 定数・列挙型定義 | MAX_CHILDREN, MAX_TESTS, eRowInput, eRowData, eColMenu等 |
| PostingModule | テストデータ登録処理 | `Posting`, `TransferData`, `ResetInputForm`, `AssignKey`, `ColumnIndexToLetter` |
| ValidationModule | 入力データ検証 | `ValidateRequiredFields`, `ValidateScoreData`, `ValidateClippingSettings`, `ValidateWeight` |
| SubjectModule | 教科別集計・ABC評価計算 | `CollectSubjectData`, `CalculateABCEvaluation`, `NormalizeWeightByAllocateScore` |
| DataManagementModule | データ修正・削除・エクスポート・保護・初期化 | `DeleteTestData`, `DeleteRetestSheetForKey`, `UpdateTestHeader`, `UpdateChildScore`, `ExportToCSV`, `ProtectScoreCells`, `UnprotectScoreCells`, `CompleteReset` |
| HistoryCheckModule | 未入力データ検索・一括転記 | `SearchNotYetInput`, `TransferFromMenu`, `SearchNotYetByTest` |
| ErrorHandlerModule | エラーハンドリング共通機能 | `ShowError`, `ShowValidationError`, `BeginProcess`, `EndProcess` |
| ScoreCalculationModule | 得点調整・変換計算（英語版） | `CalculateAdjustedAllocateScore`, `CalculateAdjustedScore` |
| ResultModule | Result転記・スナップショット保存 | `GenerateResultHeaders`, `TransferToResult`, `SaveSubjectSnapshot`, `FinalizeEvaluation`, `HasResultData`, `DeleteAllControls` |
| RetestModule | 追試機能（ファイル生成・シート作成・結果反映） | `CreateRetestSheet`, `CreateRetestSheetFromData`, `HasRetestSheetForKey`, `AddRetestRound(UI)`, `CompleteRetest(UI)`, `ApplyFinalScoreFormulas(UI)`, `OpenRetestFile`, `RefreshRetestMenu` |
| UIFormatModule | UI書式（本番使用のみ）+ 追試列視覚表示 | `FormatMenuDataArea`, `ApplyRetestColumnFormat`, `ClearRetestColumnFormat` |

### シートモジュール

| モジュール | 用途 | 主要イベント/プロシージャ |
|------------|------|--------------------------|
| ThisWorkbook | ワークブック開閉時の初期化 | `Workbook_Open` → 日付初期化、チェックボックス初期化、データシート保護 |
| Sh_data | データシートのイベント処理 | `Worksheet_BeforeDoubleClick` → 得点セル: frmScoreEdit / ヘッダー行: frmTestEdit（削除リクエスト判定含む） |
| sh_input | 入力シートのUI制御 | `Cb_clipping`, `Cb_convertScore`, `Cb_adjustScore`, `ClearInputForm`, `ToggleEnrollmentFilter` |
| sh_namelist | 名簿シートのボタンイベント | （ボタン割り当て用） |
| sh_subject | Subjectシートのボタン・イベント | `Update_Click`, `Ope_result_Click`, `Delete_Sh_Subject_Click`, `Btn_NormalizeWeight_Click`, `Worksheet_BeforeDoubleClick` |

### ユーザーフォーム

| フォーム | 用途 | 主要コントロール |
|----------|------|------------------|
| frmScoreEdit | 得点修正ダイアログ | `lblSubject`, `lblPerspective`, `lblTestname`, `lblChildName`, `lblAllocateScore`, `lblCurrentScore`, `txtNewScore`, `lblHint`, `btnUpdate`, `btnCancel`, `btn_Exempt` |
| frmTestEdit | テスト情報編集+後出し追試設定+削除ダイアログ | `lblKeyValue`, `lblSubjectValue`, `cmbCategoryValue`, `txtTestName`, `cmbPerspective`, `txtDetail`, `txtAllocateScore`, `cmbYear`, `cmbMonth`, `cmbDay`, `btnUpdate`, `btnCancel`, `btnRetest`, `btnDelete` |
| frm_retest_setting | 追試計算方法設定ダイアログ | `opbtn1`～`opbtn6`（ラジオボタン: 合格点/最大値/平均値/中央値/内分点/本試のみ）, `txtbox`（α値入力）, `btn_ok`, `btn_cancel` |
| Analysis | テスト比較分析フォーム | `Com_Subject`（教科選択）, `UnitList`（テスト一覧）, `Det1`/`Det2`（グループ1/2）, `btn_det1`/`btn_det2`（追加）, `Deldet1`/`Deldet2`（削除）, `CommandButton1`（実行） |

## 主要機能

### 1. テスト登録 (PostingModule.Posting)
- 入力シートから得点データをデータシートに転記
- テストキー自動採番（教科頭文字+連番: J001, S002等）
- 統計値（平均、中央値、標準偏差、変動係数）の自動計算
- 最大5列（5種類の観点）まで同時登録可能

### 2. 得点調整機能 (Module1, ScoreCalculationModule)
- **クリッピング**: 上限・下限で得点を制限（rowClippingSup, rowClippingInf）
- **得点変換**: 「平方根」または「対数」（底2）変換
- **範囲調整**: 調整後の得点範囲を指定（rowAdjScoreSup, rowAdjScoreInf）
- 調整後配点はワークシート数式で動的計算

### 3. 教科別集計 (SubjectModule.CollectSubjectData)
- 教科・観点でフィルタリングしてデータ収集
- 観点はチェックボックス（perspective1～5）で選択
- 重み付き/重みなし合計・配点の計算
- 得点調整の有効/無効切替（B4セル）

### 4. ABC評価 (SubjectModule.CalculateABCEvaluation)
- 達成率に基づくABC評価（加重達成率を使用）
- Settingシートから複数の閾値パターンを読込
- 7行目に●（候補）/★（採用）マークを表示
- ●をダブルクリックで★（採用）に変更、他の★は自動で●に戻る
- A計/B計/C計の自動集計
- 「最終決定」ダブルクリックでResultシートに転記＋スナップショット保存

### 5. 未入力管理 (HistoryCheckModule)
- 全テストの未入力データを検索してMENUシートに一覧表示
- MENUシートで点数を入力し、一括でデータシートに転記
- 特定テストの未入力者のみ検索も可能

### 6. 重み正規化 (SubjectModule.NormalizeWeightByAllocateScore)
- 配点の異なるテスト（50点満点と100点満点など）を同じ重要度として扱う
- 基準配点（100点）に対する比率で重みを自動調整
- 計算式: `新しい重み = (100 / 配点) × 現在の重み`
- 正規化実施後はB6セルに「実施済」と表示
- データ消去時に正規化状態もリセット

### 7. データ管理 (DataManagementModule)
- テストデータの削除（キー指定）
- ヘッダー情報の修正（テスト名、観点、詳細、配点、実施日）
- 児童の得点修正
- CSVエクスポート
- **得点セル保護**: 得点セルを誤編集から保護（ダブルクリックでフォームから安全に修正可能）

### 8. 得点修正フォーム (frmScoreEdit)
- データシートの得点セルをダブルクリックすると表示
- 教科名、観点名、テスト名、児童名、配点、現在の得点を表示
- 新しい得点を入力して更新、または「免除」ボタンで「-」を設定
- 入力値の検証（数値チェック、配点超過チェック）
- 保護されたセルでも安全に修正可能

### 9. 評価結果の保存 (ResultModule)
- **Result転記**: SubjectシートのABC評価をResultシートに転記
- **データ保護**: 既存データがある場合は列見出し再生成をスキップ（`HasResultData`関数）
- **スナップショット保存**: Subjectシートを統合ファイルにシートとして追加保存
  - 保存先: `./ファイル名_確定.xlsx`（単一ファイルに複数シートを追加）
  - シート名: `教科_観点_日付`（同名がある場合は連番追加 `_1`, `_2`...）
  - 数式は値に変換して保存（xlsxで保存可能）
  - フォームボタン・OLEオブジェクトを削除（マクロ実行防止）
  - シート保護（パスワードなし）で誤変更を防止
- **列見出し自動生成**: Workbook_Open時にSettingシートから見出しを生成（既存データがない場合のみ）

### 10. 追試機能 (RetestModule)
- **追試ファイル生成**: テスト登録時に追試フラグ（行28に"あり"）がある列について、別ファイル（`成績処理_追試.xlsm`）に追試シートを自動生成
- **テンプレート方式**: 本体ファイルのVeryHiddenシート（`sh_rt_menu`, `sh_rt_template`）をCodeNameで検索・コピーして追試シートを作成
- **追試回の追加**: 最終列の手前に列を挿入して追試回を追加（追試2, 追試3, ...）
- **算出方法**: frm_retest_settingフォームで選択、6方式対応
  - **合格点**: 本試≧合格点→本試得点、追試で合格→合格点、それ以外→最高点
  - **最大値**: MAX(本試, 追試1, 追試2, ...)
  - **平均値**: AVERAGE(本試, 追試全て)
  - **中央値**: MEDIAN(本試, 追試全て)
  - **内分点**: α × MAX(全回) + (1-α) × 本試（α値: 0～1）
  - **本試のみ**: 追試結果を無視し本試の得点をそのまま使用
- **合格者数・未合格者数**: 合格点(E4)が入力されている場合、最終列の得点から自動集計（H3/H4）
- **結果反映**: 追試完了で最終得点を本体ファイルのデータシートに反映（"N"マーカーを上書き）
- **ボタン**: 追試ファイルのボタンは`'本体ファイル名'!RetestModule.XXX`形式で本体マクロを呼び出す
- **後出し追試**: テスト登録後にデータシートのヘッダー行ダブルクリック→frmTestEditフォーム→「追試を設定」ボタンで追試シートを作成（`CreateRetestSheetFromData`）。既に同キーの追試シートがある場合は警告して中止。
- **追試中列の視覚表示**: 追試中の列はヘッダー行がオレンジ(`COLOR_RETEST_HEADER`)、得点セルが薄オレンジ(`COLOR_RETEST_CELL`)で着色され、"N"セルは太字・濃いオレンジフォントで強調表示。追試完了時に自動でクリアされる（`UIFormatModule.ApplyRetestColumnFormat` / `ClearRetestColumnFormat`）。

### 11. テスト情報編集・削除 (frmTestEdit)
- データシートのヘッダー行（4-22行）ダブルクリックで表示
- テスト名、カテゴリ、観点、詳細、配点、実施日の編集
- 追試設定ボタン（後出し追試の作成）
- **削除ボタン**: テストデータをフォームから直接削除可能。追試中のテストは強制削除確認後、追試ファイル内の対応シートとMENUエントリも同時に削除。削除はフォームを閉じた後に`DataManagementModule.DeleteTestData`経由で実行（`mDeleteRequested`/`mForceDeleteRetest`フラグ方式）。

### 12. 在籍フィルター (sh_input.ToggleEnrollmentFilter)
- 入力シートのボタン（`Btn_enrollment`）で在籍児童のみ/全員表示を切替
- 「在籍」ボタン押下: 名簿F列（在籍終了日）が実施日（J4）より前の児童行を非表示
- 「全員」ボタン押下: 全児童行を再表示
- ボタンのキャプション（「在籍」/「全員」）が状態を保持（セルに状態を保存しない）
- 実施日が未入力の場合は警告メッセージを表示して中止

### 13. 完全初期化 (DataManagementModule.CompleteReset)
- 設定シートのボタンから実行。新しい評価期間を開始するために全データを一括クリア
- 二重確認ダイアログ（2回のYes/No確認）で誤操作を防止
- **クリア対象**: データシート（テストデータ列）、Subjectシート（集計データ）、Resultシート（評価結果＋ヘッダー再生成）、MENUシート（未入力一覧＋書式）、入力シート（フォーム＋日付＋在籍ボタン）、IndividualAnalysisシート（分析データ）、Settingシートキーカウンター（C列を0にリセット）
- **保持対象**: 児童名簿（sh_namelist）、設定シート（教科/観点/カテゴリ/閾値）、スナップショットファイル（_確定.xlsx）
- 追試ファイルが存在する場合は警告メッセージを表示（自動削除はしない）

## 重要な定数・列挙型

```vba
' システム上限値 (PublicConstListModule)
MAX_CHILDREN = 40           ' 児童数上限
MAX_TESTS = 1000            ' テスト数上限
MAX_PERSPECTIVES = 5        ' 評価観点数上限
NORMALIZE_BASE_SCORE = 100  ' 重み正規化の基準配点

' 入力シート行番号 (eRowInput)
rowPerspective = 8          ' 評価観点
rowAllocateScore = 12       ' 配点
rowClippingSup = 14         ' クリッピング上限
rowClippingInf = 16         ' クリッピング下限
rowConvScore = 18           ' 得点変換方式
rowAdjScoreSup = 20         ' 調整後範囲上限
rowAdjScoreInf = 22         ' 調整後範囲下限
rowWeight = 26              ' 重み
rowChildStart = 31          ' 児童データ開始行

' データシート行番号 (eRowData)
rowKey = 4                  ' テストキー
rowTestDate = 5             ' 実施日
rowSubject = 6              ' 教科
rowCategory = 7             ' カテゴリ
rowTestName = 8             ' テスト名
rowPerspective = 9          ' 観点
rowDetail = 10              ' 詳細
rowAllocationScore = 11     ' 配点
rowAdjAllocateScore = 17    ' 調整後配点
rowWeight = 18              ' 重み
rowAverage = 19             ' 平均
rowChildStart = 23          ' 児童データ開始行

' Subjectシート行番号 (eRowSubject) - eRowDataと同じ構造
rowChildStart = 23          ' 児童データ開始行

' 重要なセル参照定数
RNG_NAMELIST_CHILDCOUNT = "F8"           ' 名簿シートの児童数
NAMELIST_COL_END_DATE = 6                ' 名簿シートの在籍終了日（F列）
RNG_INPUT_SUBJECT = "D4"                 ' 入力シートの教科
RNG_SUBJECT_SUBJECT = "B2"               ' Subjectシートの教科
RNG_SUBJECT_ISADJUST = "B4"              ' 得点調整有効/無効
RNG_SUBJECT_ADJSCORE_DISP = "B5"         ' 得点調整行表示状態
RNG_SUBJECT_STATS_DISP = "B6"            ' 統計行表示状態
RNG_SUBJECT_WEIGHT_NORMALIZED = "B7"     ' 重み正規化状態

' Resultシート定数
RESULT_SUBJECT_ROW = 8                   ' 教科名行
RESULT_PERSPECTIVE_ROW = 9               ' 観点行
RESULT_LABEL_ROW = 10                    ' ラベル行（達成率/ABC）
RESULT_DATA_START_ROW = 11               ' 児童データ開始行
RESULT_DATA_START_COL = 4                ' データ開始列（D列）

' 追試中列の視覚表示用色定数
COLOR_RETEST_HEADER = 7882751            ' RGB(255, 200, 120) ヘッダー行用オレンジ
COLOR_RETEST_CELL = 11854079             ' RGB(255, 230, 180) 得点セル用薄オレンジ

' 追試シート定数
ROW_INPUT_RETEST = 28                    ' 入力シートの追試有無行
RETEST_ENABLED_VALUE = "あり"            ' 追試有効の判定値

' 追試シートのセル位置（A-B列: テスト情報、D-E列: 算出設定）
RNG_RT_PARENT_KEY = "B3"                 ' 追試元キー
RNG_RT_SUBJECT = "B4"                    ' 教科
RNG_RT_TEST_NAME = "B5"                  ' テスト名
RNG_RT_PERSPECTIVE = "B6"                ' 観点
RNG_RT_DETAIL = "B7"                     ' 詳細
RNG_RT_ALLOCATE = "E3"                   ' 配点
RNG_RT_PASS_SCORE = "E4"                 ' 合格点（空欄可）
RNG_RT_METHOD = "E5"                     ' 算出方法
RNG_RT_PARAM = "E6"                      ' 内分比α値
RNG_RT_STATUS = "E7"                     ' 状態

' 追試シートのデータ領域
RT_HEADER_ROW = 10                       ' ヘッダー行
RT_DATA_START_ROW = 11                   ' 児童データ開始行
RT_COL_CODE = 1                          ' A列：コード
RT_COL_ORIGINAL = 4                      ' D列：本試
RT_COL_RETEST_START = 5                  ' E列～：追試1, 追試2, ...

' 算出方法の選択肢
RT_METHOD_PASS_SCORE = "合格点"
RT_METHOD_MAX = "最大値"
RT_METHOD_AVERAGE = "平均値"
RT_METHOD_MEDIAN = "中央値"
RT_METHOD_INTERPOLATION = "内分点"
RT_METHOD_ORIGINAL_ONLY = "本試のみ"
```

## データ構造

### データシートの列構造
```
列A: 児童コード (colCode = 1)
列B: 姓 (colLastName = 2)
列C: 名 (colFirstName = 3)
列D～: テストデータ (colDataStart = 4)
```

### MENUシートの列構造
```
列B: 児童コード (colCode = 2)
列C: 姓 (colLastName = 3)
列D: 名 (colFirstName = 4)
列E: 教科 (colSubject = 5)
列F: 観点 (colPerspective = 6)
列G: テスト名 (colTestName = 7)
列H: 詳細 (colDetail = 8)
列I: 点数（入力欄）(colScore = 9)
列J: 配点 (colAllocateScore = 10)
列K: 転記先行 (colToRow = 11)
列L: 転記先列 (colToCol = 12)
```

### Subjectシート集計列オフセット (eColShiftSubject)
```
+2: 重み無し合計 (colNoWeightSum)
+3: 重み無し配点 (colNoWeightAllocate)
+4: 加重合計 (colIncludeWeightSum)
+5: 加重配点 (colIncludeWeightAllocate)
+7: 重み無し達成率 (colNoWeightRatio)
+8: 加重達成率 (colIncludeWeightRatio)
+10: ABC閾値ヘッダー (colABCBorder)
```

## 開発時の注意事項

### 変更禁止事項
- `Module1`のモジュール名（ワークシート数式で参照）
- 日本語関数名（`調整後配点計算`, `調整後得点計算`, `得点変換`）
- シートのCodeName（sh_input, sh_namelist, Sh_data, sh_subject, sh_MENU, sh_setting, sh_result）

### コーディング規約
- 変数宣言は`Long`型を使用（`Integer`は非推奨）
- エラーハンドリングは`ErrorHandlerModule`を使用
- 処理開始時は`BeginProcess`、終了時は`EndProcess`を呼ぶ
- ユーザー向けメッセージは日本語で丁寧に
- 日本語関数名はModule1のみ、他モジュールは英語関数名

### 典型的なエラーハンドリングパターン
```vba
Public Sub SampleProcedure()
    On Error GoTo ErrorHandler

    Call ErrorHandlerModule.BeginProcess

    ' 処理本体

CleanExit:
    Call ErrorHandlerModule.EndProcess
    Exit Sub

ErrorHandler:
    Call ErrorHandlerModule.CleanupOnError
    Dim errInfo As ErrorInfo
    errInfo = ErrorHandlerModule.CreateErrorInfo("ModuleName", "ProcedureName")
    Call ErrorHandlerModule.ShowError(errInfo)
End Sub
```

### ValidationResult構造体の使用
```vba
Public Type ValidationResult
    IsValid As Boolean
    ErrorMessage As String
    ErrorRow As Long
    ErrorCol As Long
End Type
```

## ビルド・デバッグ

1. `成績処理.xlsm`をExcelで開く（マクロを有効化）
2. Alt+F11でVBAエディタを開く
3. VBEのデバッグ機能を使用

### VBAコードのエクスポート
```
ファイル > エクスポート でモジュールを.bas/.clsとして保存
```

## データフロー

```
┌─────────────────────────────────────────────────────────────┐
│ 入力シート                                                   │
│   教科・カテゴリ・日付・テスト名                              │
│   観点・配点・得点調整設定                                    │
│   児童別得点（最大5列）                                       │
└──────────────────────┬──────────────────────────────────────┘
                       │ PostingModule.Posting
                       ▼
┌─────────────────────────────────────────────────────────────┐
│ データシート                                                 │
│   テストキー(J001等)・ヘッダー情報・統計値                    │
│   児童別得点                                                 │
└──────────────────────┬──────────────────────────────────────┘
                       │ SubjectModule.CollectSubjectData
                       │ (教科・観点でフィルタ)
                       ▼
┌─────────────────────────────────────────────────────────────┐
│ Subjectシート                                                │
│   収集されたテストデータ                                     │
│   重み正規化（オプション）                                   │
└──────────────────────┬──────────────────────────────────────┘
                       │ SubjectModule.CalculateABCEvaluation
                       ▼
┌─────────────────────────────────────────────────────────────┐
│ ABC評価結果                                                  │
│   重み無し/加重の合計・配点・達成率                          │
│   7行目: ●（候補）/ ★（採用）                               │
│   ABC閾値候補 → ●ダブルクリックで★採用                      │
└──────────────────────┬──────────────────────────────────────┘
                       │ 「最終決定」ダブルクリック
                       │ ResultModule.FinalizeEvaluation
                       ▼
┌─────────────────────────────────────────────────────────────┐
│ Resultシート                                                 │
│   教科×観点ごとの達成率・ABC評価を保存                       │
│   （列見出しはWorkbook_Open時に自動生成）                    │
└─────────────────────────────────────────────────────────────┘
                       │
                       │ 同時にスナップショット保存
                       ▼
┌─────────────────────────────────────────────────────────────┐
│ スナップショットファイル                                     │
│   ./ファイル名_確定.xlsx（単一ファイル、シートとして追加）   │
│   Subjectシートの状態を保存（数式は値に変換）                │
│   フォームボタン削除、シート保護                             │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│ MENUシート（未入力管理）                                     │
│   HistoryCheckModule.SearchNotYetInput → 未入力一覧作成      │
│   点数入力 → TransferFromMenu → データシートに転記           │
└─────────────────────────────────────────────────────────────┘
```

## 検証ルール (ValidationModule)

| 検証対象 | ルール | 関数 |
|----------|--------|------|
| 必須項目 | 教科、カテゴリ、実施日（日付形式）、テスト名 | ValidateRequiredFields |
| 教科 | Settingシートに登録されている教科名 | ValidateSubjectExists |
| 配点 | 正の数値、0は不可（ゼロ除算防止） | ValidateScoreData |
| 得点 | 0以上、配点以下、または"-"（免除）、空欄可 | ValidateScoreData |
| 評価観点 | 得点データがある列は必須 | ValidateScoreData |
| クリッピング | 上限 >= 下限、上限 <= 配点、数値 | ValidateClippingSettings |
| 重み | 0以上の数値（空欄は1として扱う） | ValidateWeight |
| テスト数 | MAX_TESTS (1000) 以下 | ValidateTestCountLimit |

## 運用ルール

- 1ファイル = 1クラス × 1評価期間（1学期、前期など）
- 次の評価期間は新規ファイルを作成
- 縄跳び等の上限なしデータは、最大値または目標値を配点として入力
- 免除（評価対象外）は得点欄に「-」を入力（達成率算出時に計算から除外される）
- 空欄は未入力扱い（MENUシートの未入力検索で検出される）

## シート上のボタン・コントロール

### 入力シート (sh_input)
- チェックボックス: `Cb_clipping`, `Cb_convertScore`, `Cb_adjustScore`
- ボタン: 登録（PostingModule.Posting）、クリア（ClearInputForm）、在籍/全員切替（ToggleEnrollmentFilter、ボタン名: Btn_enrollment）

### Subjectシート (sh_subject)
- チェックボックス: `perspective1`～`perspective5`（観点選択）
- ボタン: 追加/更新（Update_Click）、評価（Ope_result_Click）、消去（Delete_Sh_Subject_Click）
- ボタン: 得点調整表示切替（Btm_adjustscore_hide_reveal）、得点調整有効/無効（Btn_IS_adj_score）
- ボタン: 重み正規化（Btn_NormalizeWeight_Click）

### MENUシート (sh_MENU)
- ボタン: 未入力検索（SearchNotYetInput）、転記（TransferFromMenu）

### Settingシート (sh_setting)
- ボタン: 完全初期化（DataManagementModule.CompleteReset）

## 用語

| 日本語 | 英語 | 説明 |
|--------|------|------|
| 免除 | Exempt | 評価対象外とする（得点欄に「-」を入力）。達成率算出時に計算から除外される |
| 配点 | AllocateScore | テストの満点（100点満点なら100） |
| 達成率 | Ratio | 得点÷配点×100（加重達成率は重みを考慮） |
| 観点 | Perspective | 評価の観点（例：知識・技能、思考・判断・表現） |

## 将来的な拡張候補

- **エクスポート機能の拡充**: 通知表用フォーマット出力、指導要録用フォーマット、成績分布グラフ等
