VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetup
   Caption         =   "初期設定"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "frmSetup.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' フォーム名: frmSetup
' 用途: 初期設定（名簿確認・教科選択・実行）
'
' ■ 手動配置が必要なコントロール一覧（全てフォーム上に直接配置）
'   名前                種類              ※位置・サイズはコードで自動設定
'   lblChildCount       Label
'   btnOpenNamelist      CommandButton
'   lstSubjects         ListBox           ★ ListStyle=1, MultiSelect=1
'   txtCustomSubject    TextBox
'   btnAddCustom        CommandButton
'   btnRemoveCustom     CommandButton
'   lblInfo             Label
'   btnExecute          CommandButton
'   btnCancel           CommandButton
'===============================================================================
Option Explicit

' プリセット教科
Private Const PRESET_COUNT As Long = 13

Private Type SubjectInfo
    Name As String
    KeyChar As String
    IsCustom As Boolean
End Type

Private mAllSubjects() As SubjectInfo
Private mSubjectCount As Long

' デザイン定数
Private Const CLR_BG As Long = 15790840         ' RGB(248, 249, 241)
Private Const CLR_BANNER As Long = 8345655       ' RGB(55, 90, 127)
Private Const CLR_WHITE As Long = 16777215       ' RGB(255, 255, 255)
Private Const CLR_SECTION As Long = 5592405      ' RGB(85, 85, 85)
Private Const CLR_PANEL_BG As Long = 16444400    ' RGB(240, 243, 250)
Private Const CLR_BORDER As Long = 14472894      ' RGB(190, 200, 220)
Private Const CLR_TEXT As Long = 5263440          ' RGB(80, 80, 80)
Private Const CLR_LINE As Long = 14803425        ' RGB(193, 200, 225)

'===============================================================================
' フォーム初期化
'===============================================================================
Private Sub UserForm_Initialize()
    mSubjectCount = PRESET_COUNT
    ReDim mAllSubjects(1 To PRESET_COUNT)

    ' プリセット教科
    ' プリセットキー文字（アルファベット固定）
    ' 使用済: D,E,G,J,K,L,M,O,R,S,T,X,Z
    ' カスタム用に空き: A,B,C,F,H,I,N,P,Q,U,V,W,Y
    Call SetPreset(1, "国語", "J")
    Call SetPreset(2, "社会", "S")
    Call SetPreset(3, "算数", "M")
    Call SetPreset(4, "理科", "R")
    Call SetPreset(5, "生活", "L")
    Call SetPreset(6, "音楽", "O")
    Call SetPreset(7, "図画工作", "Z")
    Call SetPreset(8, "家庭", "K")
    Call SetPreset(9, "体育", "T")
    Call SetPreset(10, "外国語", "E")
    Call SetPreset(11, "道徳", "D")
    Call SetPreset(12, "特別活動", "X")
    Call SetPreset(13, "総合的な学習の時間", "G")

    ' デザイン適用
    Call ApplyFormDesign

    Call RefreshSubjectList

    ' Settingシートの既存教科を読み込み、チェック状態を復元
    Call LoadExistingSubjects

    Call UpdateChildCount

    Me.Caption = "初期設定"
End Sub

Private Sub SetPreset(ByVal idx As Long, ByVal n As String, ByVal k As String)
    mAllSubjects(idx).Name = n
    mAllSubjects(idx).KeyChar = k
    mAllSubjects(idx).IsCustom = False
End Sub

'===============================================================================
' デザイン適用（コントロールの位置・サイズ・見た目を一括設定）
'===============================================================================
Private Sub ApplyFormDesign()
    ' --- フォームサイズを強制設定 ---
    Me.Width = 476
    Me.Height = 500

    ' --- フォーム背景 ---
    Me.BackColor = CLR_BG

    ' --- ヘッダーバナー ---
    Dim lblBannerBg As MSForms.Label
    Set lblBannerBg = Me.Controls.Add("Forms.Label.1", "lblBannerBg")
    With lblBannerBg
        .Left = 0: .Top = 0
        .Width = Me.InsideWidth: .Height = 36
        .BackColor = CLR_BANNER
        .BackStyle = fmBackStyleOpaque
        .Caption = ""
    End With

    Dim lblBannerTxt As MSForms.Label
    Set lblBannerTxt = Me.Controls.Add("Forms.Label.1", "lblBannerTxt")
    With lblBannerTxt
        .Left = 14: .Top = 8
        .Width = 300: .Height = 20
        .BackStyle = fmBackStyleTransparent
        .ForeColor = CLR_WHITE
        .Caption = "初期設定ウィザード"
        .Font.Size = 12
        .Font.Bold = True
    End With

    ' --- イントロ説明 ---
    Dim lblIntro As MSForms.Label
    Set lblIntro = Me.Controls.Add("Forms.Label.1", "lblIntro")
    With lblIntro
        .Left = 16: .Top = 44
        .Width = 435: .Height = 26
        .BackStyle = fmBackStyleTransparent
        .ForeColor = CLR_TEXT
        .Font.Size = 9
        .WordWrap = True
        .Caption = "成績処理に必要な基本設定を行います。" & _
                   "名簿の確認と、使用する教科の選択を行ってください。"
    End With

    ' ==============================
    ' Step 1: 名簿
    ' ==============================
    Dim lblStep1 As MSForms.Label
    Set lblStep1 = Me.Controls.Add("Forms.Label.1", "lblStep1")
    With lblStep1
        .Left = 14: .Top = 76
        .Width = 435: .Height = 14
        .Caption = "Step 1  名簿の確認"
        .ForeColor = CLR_BANNER
        .Font.Size = 9
        .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With

    Dim lblLine1 As MSForms.Label
    Set lblLine1 = Me.Controls.Add("Forms.Label.1", "lblLine1")
    With lblLine1
        .Left = 14: .Top = 90
        .Width = 435: .Height = 1
        .BackColor = CLR_LINE
        .BackStyle = fmBackStyleOpaque
    End With

    Dim lblStep1Desc As MSForms.Label
    Set lblStep1Desc = Me.Controls.Add("Forms.Label.1", "lblStep1Desc")
    With lblStep1Desc
        .Left = 16: .Top = 96
        .Width = 290: .Height = 14
        .BackStyle = fmBackStyleTransparent
        .ForeColor = CLR_TEXT
        .Font.Size = 9
        .Caption = "名簿シートに児童名を入力してから先に進んでください。"
    End With

    With lblChildCount
        .Left = 16: .Top = 114
        .Width = 280: .Height = 18
        .Font.Size = 10
        .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With

    With btnOpenNamelist
        .Left = 316: .Top = 112
        .Width = 130: .Height = 24
        .Caption = "名簿シートを開く"
        .Font.Size = 9
    End With

    ' ==============================
    ' Step 2: 教科選択
    ' ==============================
    Dim lblStep2 As MSForms.Label
    Set lblStep2 = Me.Controls.Add("Forms.Label.1", "lblStep2")
    With lblStep2
        .Left = 14: .Top = 144
        .Width = 435: .Height = 14
        .Caption = "Step 2  使用する教科を選択"
        .ForeColor = CLR_BANNER
        .Font.Size = 9
        .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With

    Dim lblLine2 As MSForms.Label
    Set lblLine2 = Me.Controls.Add("Forms.Label.1", "lblLine2")
    With lblLine2
        .Left = 14: .Top = 158
        .Width = 435: .Height = 1
        .BackColor = CLR_LINE
        .BackStyle = fmBackStyleOpaque
    End With

    Dim lblStep2Desc As MSForms.Label
    Set lblStep2Desc = Me.Controls.Add("Forms.Label.1", "lblStep2Desc")
    With lblStep2Desc
        .Left = 16: .Top = 164
        .Width = 435: .Height = 14
        .BackStyle = fmBackStyleTransparent
        .ForeColor = CLR_TEXT
        .Font.Size = 9
        .Caption = "チェックした教科がSettingシートに登録されます。後から変更も可能です。"
    End With

    ' --- 教科リスト ---
    With lstSubjects
        .Left = 16: .Top = 184
        .Width = 258: .Height = 178
        .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CLR_BORDER
    End With

    ' --- 情報パネル ---
    With lblInfo
        .Left = 286: .Top = 184
        .Width = 162: .Height = 178
        .BackColor = CLR_PANEL_BG
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CLR_BORDER
        .ForeColor = CLR_TEXT
        .Font.Size = 9
        .TextAlign = fmTextAlignLeft
        .WordWrap = True
        .Caption = BuildInfoCaption()
    End With

    ' --- カスタム教科の追加/削除 ---
    Dim lblCustom As MSForms.Label
    Set lblCustom = Me.Controls.Add("Forms.Label.1", "lblCustom")
    With lblCustom
        .Left = 16: .Top = 368
        .Width = 100: .Height = 14
        .BackStyle = fmBackStyleTransparent
        .ForeColor = CLR_TEXT
        .Font.Size = 8
        .Caption = "リストにない教科:"
    End With

    With txtCustomSubject
        .Left = 16: .Top = 384
        .Width = 178: .Height = 22
        .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CLR_BORDER
    End With

    With btnAddCustom
        .Left = 200: .Top = 384
        .Width = 55: .Height = 22
        .Caption = "追加"
        .Font.Size = 9
    End With

    With btnRemoveCustom
        .Left = 260: .Top = 384
        .Width = 55: .Height = 22
        .Caption = "削除"
        .Font.Size = 9
    End With

    ' --- 下部の区切り線 ---
    Dim lblBottomLine As MSForms.Label
    Set lblBottomLine = Me.Controls.Add("Forms.Label.1", "lblBottomLine")
    With lblBottomLine
        .Left = 14: .Top = 418
        .Width = 435: .Height = 1
        .BackColor = CLR_LINE
        .BackStyle = fmBackStyleOpaque
    End With

    ' --- ボタンエリア ---
    With btnCancel
        .Left = 220: .Top = 428
        .Width = 100: .Height = 28
        .Caption = "キャンセル"
        .Font.Size = 10
    End With

    With btnExecute
        .Left = 330: .Top = 428
        .Width = 120: .Height = 28
        .Caption = "実行"
        .Font.Size = 11
        .Font.Bold = True
    End With
End Sub

'===============================================================================
' Settingシートの既存教科を読み込み、チェック状態を復元
'===============================================================================
Private Sub LoadExistingSubjects()
    Dim r As Long
    Dim existingName As String
    Dim existingKey As String
    Dim found As Boolean
    Dim i As Long
    Dim needRefresh As Boolean

    ' 選択すべきインデックスを記録する配列（最大18教科分）
    Dim selectFlags() As Boolean
    ReDim selectFlags(0 To mSubjectCount + 17)

    needRefresh = False

    For r = SETTING_SUBJECT_START_ROW To 20
        existingName = Trim(sh_setting.Cells(r, SETTING_SUBJECT_COL).value & "")
        If existingName = "" Then Exit For

        existingKey = Trim(sh_setting.Cells(r, SETTING_KEY_CHAR_COL).value & "")

        ' プリセット一覧から一致を探す
        found = False
        For i = 1 To mSubjectCount
            If mAllSubjects(i).Name = existingName Then
                ' 見つかった → 選択フラグを記録
                selectFlags(i - 1) = True
                ' キー文字もSettingの値で上書き（ユーザーが変更済みの場合）
                If existingKey <> "" Then mAllSubjects(i).KeyChar = existingKey
                found = True
                Exit For
            End If
        Next i

        ' プリセットにない教科 → カスタムとして追加
        If Not found Then
            mSubjectCount = mSubjectCount + 1
            ReDim Preserve mAllSubjects(1 To mSubjectCount)
            mAllSubjects(mSubjectCount).Name = existingName
            If existingKey <> "" Then
                mAllSubjects(mSubjectCount).KeyChar = existingKey
            Else
                mAllSubjects(mSubjectCount).KeyChar = AssignKeyChar(existingName)
            End If
            mAllSubjects(mSubjectCount).IsCustom = True
            selectFlags(mSubjectCount - 1) = True
            needRefresh = True
        End If
    Next r

    ' カスタム教科が追加された場合はリストを再構築
    If needRefresh Then Call RefreshSubjectList

    ' まとめて選択状態を適用
    For i = 0 To lstSubjects.ListCount - 1
        If selectFlags(i) Then lstSubjects.Selected(i) = True
    Next i
End Sub

'===============================================================================
' リスト表示を更新
'===============================================================================
Private Sub RefreshSubjectList()
    Dim i As Long
    lstSubjects.Clear
    For i = 1 To mSubjectCount
        If mAllSubjects(i).IsCustom Then
            lstSubjects.AddItem mAllSubjects(i).Name & " （追加）"
        Else
            lstSubjects.AddItem mAllSubjects(i).Name
        End If
    Next i
End Sub

'===============================================================================
' 児童数を更新
'===============================================================================
Private Sub UpdateChildCount()
    Dim count As Long
    count = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value

    If count > 0 Then
        lblChildCount.Caption = "児童数: " & count & " 名"
        lblChildCount.ForeColor = RGB(0, 120, 60)
    Else
        lblChildCount.Caption = "児童が登録されていません"
        lblChildCount.ForeColor = RGB(200, 50, 50)
    End If
End Sub

'===============================================================================
' フォームがアクティブになった時（名簿編集後に戻った時に児童数を更新）
'===============================================================================
Private Sub UserForm_Activate()
    Call UpdateChildCount
End Sub

'===============================================================================
' フォームクリック時にも児童数を更新（モードレスでActivateが発火しない場合の補完）
'===============================================================================
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static lastCount As Long
    Dim currentCount As Long
    currentCount = sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value
    If currentCount <> lastCount Then
        Call UpdateChildCount
        lastCount = currentCount
    End If
End Sub

'===============================================================================
' 名簿シートを開くボタン（モードレスなのでそのまま名簿シートへ移動）
'===============================================================================
Private Sub btnOpenNamelist_Click()
    sh_namelist.Activate
End Sub

'===============================================================================
' カスタム教科の追加
'===============================================================================
Private Sub btnAddCustom_Click()
    Dim newName As String
    newName = Trim(txtCustomSubject.value)

    If newName = "" Then Exit Sub

    ' 重複チェック
    Dim i As Long
    For i = 1 To mSubjectCount
        If mAllSubjects(i).Name = newName Then
            MsgBox "この教科は既にリストにあります。", vbExclamation
            Exit Sub
        End If
    Next i

    ' 追加
    mSubjectCount = mSubjectCount + 1
    ReDim Preserve mAllSubjects(1 To mSubjectCount)
    mAllSubjects(mSubjectCount).Name = newName
    mAllSubjects(mSubjectCount).KeyChar = AssignKeyChar(newName)
    mAllSubjects(mSubjectCount).IsCustom = True

    Call RefreshSubjectList
    ' 新しく追加した項目を選択状態にし、フォーカスも移動する
    lstSubjects.ListIndex = mSubjectCount - 1
    lstSubjects.Selected(mSubjectCount - 1) = True
    txtCustomSubject.value = ""
End Sub

'===============================================================================
' カスタム教科の削除（カスタムのみ削除可能）
'===============================================================================
Private Sub btnRemoveCustom_Click()
    Dim idx As Long
    idx = lstSubjects.ListIndex

    If idx < 0 Then Exit Sub
    idx = idx + 1 ' 1-based

    If Not mAllSubjects(idx).IsCustom Then
        MsgBox "プリセット教科は削除できません。" & vbCrLf & _
               "チェックを外して除外してください。", vbInformation
        Exit Sub
    End If

    ' テストデータが存在する教科は削除不可
    If HasTestDataForSubject(mAllSubjects(idx).Name) Then
        MsgBox "「" & mAllSubjects(idx).Name & "」にはテストデータが登録されているため、" & vbCrLf & _
               "削除できません。" & vbCrLf & vbCrLf & _
               "教科を削除するには、先にデータシートから" & vbCrLf & _
               "該当教科のテストデータをすべて削除してください。", vbExclamation, "削除不可"
        Exit Sub
    End If

    ' 配列から削除（シフト）
    Dim i As Long
    For i = idx To mSubjectCount - 1
        mAllSubjects(i) = mAllSubjects(i + 1)
    Next i
    mSubjectCount = mSubjectCount - 1
    ReDim Preserve mAllSubjects(1 To mSubjectCount)

    Call RefreshSubjectList
End Sub

'===============================================================================
' 情報パネルのキャプション生成（Settingシートから動的取得）
'===============================================================================
Private Function BuildInfoCaption() As String
    Dim cap As String
    Dim pName As String
    Dim r As Long
    Dim hasPerspective As Boolean

    ' --- 観点 ---
    hasPerspective = False
    cap = " 【観点】" & vbCrLf

    For r = SETTING_SUBJECT_START_ROW To SETTING_SUBJECT_START_ROW + MAX_PERSPECTIVES - 1
        pName = Trim(sh_setting.Cells(r, SETTING_PERSPECTIVE_COL).value & "")
        If pName = "" Then Exit For
        cap = cap & "  ・" & pName & vbCrLf
        hasPerspective = True
    Next r

    If Not hasPerspective Then
        ' Settingに未設定の場合はデフォルトを表示
        cap = cap & "  ・知識・技能" & vbCrLf & _
                    "  ・思考・判断・表現" & vbCrLf & _
                    "  ・主体的に学習に" & vbCrLf & _
                    "    取り組む態度" & vbCrLf
    End If

    cap = cap & vbCrLf

    ' --- ABC閾値 ---
    Dim abVal As Variant, bcVal As Variant
    abVal = sh_setting.Cells(SETTING_SUBJECT_START_ROW, SETTING_AB_THRESHOLD_COL).value
    bcVal = sh_setting.Cells(SETTING_SUBJECT_START_ROW, SETTING_BC_THRESHOLD_COL).value

    cap = cap & " 【ABC閾値】" & vbCrLf
    If IsNumeric(abVal) And IsNumeric(bcVal) Then
        cap = cap & "   A >= " & abVal & "%" & vbCrLf & _
                    "   B >= " & bcVal & "%" & vbCrLf
    Else
        cap = cap & "   A >= 80%  B >= 50%" & vbCrLf
    End If

    cap = cap & vbCrLf & "  Settingシートで" & vbCrLf & _
                          "  変更できます。"

    BuildInfoCaption = cap
End Function

'===============================================================================
' キー文字の自動割当（アルファベット候補から空き文字を割当）
' プリセット使用済: D,E,G,J,K,L,M,O,R,S,T,X,Z
' カスタム用候補: A,B,C,F,H,I,N,P,Q,U,V,W,Y
'===============================================================================
Private Function AssignKeyChar(ByVal subjectName As String) As String
    Const CANDIDATES As String = "ABCFHINPQUVWY"
    Dim c As Long
    Dim i As Long
    Dim candidate As String
    Dim isDuplicate As Boolean

    ' 候補アルファベットから空いている文字を順に探す
    For c = 1 To Len(CANDIDATES)
        candidate = Mid(CANDIDATES, c, 1)

        isDuplicate = False
        For i = 1 To mSubjectCount
            If UCase(mAllSubjects(i).KeyChar) = candidate Then
                isDuplicate = True
                Exit For
            End If
        Next i

        If Not isDuplicate Then
            AssignKeyChar = candidate
            Exit Function
        End If
    Next c

    ' 全候補が使い切られた場合（26教科超は実用上ないが安全策）
    AssignKeyChar = "?"
End Function

'===============================================================================
' 指定教科のテストデータがデータシートに存在するか確認
'===============================================================================
Private Function HasTestDataForSubject(ByVal subjectName As String) As Boolean
    Dim col As Long
    Dim lastCol As Long

    HasTestDataForSubject = False

    With Sh_data
        lastCol = .Cells(eRowData.rowKey, Columns.count).End(xlToLeft).Column
        If lastCol < eColData.colDataStart Then Exit Function

        For col = eColData.colDataStart To lastCol
            If Trim(.Cells(eRowData.rowSubject, col).value & "") = subjectName Then
                HasTestDataForSubject = True
                Exit Function
            End If
        Next col
    End With
End Function

'===============================================================================
' 実行ボタン
'===============================================================================
Private Sub btnExecute_Click()
    ' 児童数チェック
    If sh_namelist.Range(RNG_NAMELIST_CHILDCOUNT).value = 0 Then
        MsgBox "名簿シートに児童を登録してください。", vbExclamation, "確認"
        Exit Sub
    End If

    ' 選択教科を収集
    Dim selectedCount As Long
    Dim i As Long
    selectedCount = 0

    For i = 0 To lstSubjects.ListCount - 1
        If lstSubjects.Selected(i) Then
            selectedCount = selectedCount + 1
        End If
    Next i

    If selectedCount = 0 Then
        MsgBox "教科を1つ以上選択してください。", vbExclamation, "確認"
        Exit Sub
    End If

    ' 既存教科との差分チェック（外された教科の検証）
    Dim removedList As String
    Dim removedCount As Long
    Dim blockedList As String
    Dim blockedCount As Long
    Dim r As Long
    Dim existName As String
    Dim stillSelected As Boolean
    Dim j As Long

    removedList = ""
    removedCount = 0
    blockedList = ""
    blockedCount = 0

    For r = SETTING_SUBJECT_START_ROW To 20
        existName = Trim(sh_setting.Cells(r, SETTING_SUBJECT_COL).value & "")
        If existName = "" Then Exit For

        ' 選択リストに含まれているかチェック
        stillSelected = False
        For j = 0 To lstSubjects.ListCount - 1
            If lstSubjects.Selected(j) Then
                If mAllSubjects(j + 1).Name = existName Then
                    stillSelected = True
                    Exit For
                End If
            End If
        Next j

        If Not stillSelected Then
            ' テストデータがある教科は除外不可
            If HasTestDataForSubject(existName) Then
                blockedCount = blockedCount + 1
                blockedList = blockedList & "  ・" & existName & vbCrLf
            Else
                removedCount = removedCount + 1
                removedList = removedList & "  ・" & existName & vbCrLf
            End If
        End If
    Next r

    ' テストデータがある教科の除外はブロック
    If blockedCount > 0 Then
        MsgBox "以下の教科にはテストデータが登録されているため、" & vbCrLf & _
               "除外できません:" & vbCrLf & vbCrLf & _
               blockedList & vbCrLf & _
               "チェックを入れ直すか、先にデータシートから" & vbCrLf & _
               "該当教科のテストデータを削除してください。", _
               vbExclamation, "教科除外不可"
        Exit Sub
    End If

    ' テストデータのない教科の除外は確認のみ
    If removedCount > 0 Then
        If MsgBox("以下の教科が除外されます:" & vbCrLf & vbCrLf & _
                  removedList & vbCrLf & _
                  "これらの教科はSettingシートから削除されます。" & vbCrLf & vbCrLf & _
                  "続行しますか？", vbYesNo + vbExclamation, "教科除外の確認") <> vbYes Then
            Exit Sub
        End If
    End If

    ' 最終確認
    If MsgBox("選択された " & selectedCount & " 教科で初期設定を実行します。" & vbCrLf & _
              "よろしいですか？", vbYesNo + vbQuestion, "確認") <> vbYes Then
        Exit Sub
    End If

    ' 配列に格納
    Dim subjects() As String
    Dim keyChars() As String
    ReDim subjects(1 To selectedCount)
    ReDim keyChars(1 To selectedCount)

    Dim idx As Long
    idx = 0
    For i = 0 To lstSubjects.ListCount - 1
        If lstSubjects.Selected(i) Then
            idx = idx + 1
            subjects(idx) = mAllSubjects(i + 1).Name
            keyChars(idx) = mAllSubjects(i + 1).KeyChar
        End If
    Next i

    ' セットアップ実行
    Call InitialSetupModule.ExecuteSetup(subjects, keyChars, selectedCount)

    Unload Me
End Sub

'===============================================================================
' キャンセルボタン
'===============================================================================
Private Sub btnCancel_Click()
    Unload Me
End Sub
