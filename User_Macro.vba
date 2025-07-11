'
' ユーザーマクロ
' 自由に記述、編集してください。
'

Sub 確認()
  '
  '【構文】  fn_確認("<メーッセージ>","<アイコンタイプ: i ? ! x >","<タイトル>")
  '【戻り値】［OK］［はい］は 1、［いいえ］は 2、［x］は 0
  '
  On Error GoTo subError
  
  NN = fn_確認("こんにちは！", "i", "確認")
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub メニュー()
  '
  '【構文】　fn_メニュー("<タイトル>","<メニュー1>,<メニュー2>,・・・","<メーッセージ>",メニュー初期値)
  '【戻り値】選択したメニュー番号。[中止]したときは 0 を返します。
  '
  On Error GoTo subError
  
  NN = fn_メニュー中("タイトル", "メニュー1,メニュー2", "選択してください", 0)
  
  '(結果表示)
  If NN > 0 Then
    NN = fn_確認(Str(n) & " 番を 選択しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub シートメニュー()
  '
  '【構文】　fs_シートメニュー("<クリック表示>","<非表示シート>",メニュー初期値, 表示オプション)
  '【戻り値】選択したシート名。[中止]したときは 空白 を返します。
  '
  '<クリック表示>:クリックしたときにシートを表示する
  '<非表示シート>:表示しているけども、メニューに表示しないシート
  'メニュー初期値:0～n
  '表示オプション:True:全て表示/False:非表示シートを除く
  '
  On Error GoTo subError
  
  SS = fs_シートメニュー("する", "サンプルマクロ,パーツ", 0, False)
  
  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " を 選択しました。", "i", "確認")
    NN = fn_シート表示("サンプルマクロ", "")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub 日付入力()
  '
  '【構文】　fs_条件日付("<日付初期値>","しない")
  '【戻り値】入力した日付。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError

  SS = fs_条件日付("", "しない", "")

  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " と 入力しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub 日付_月範囲()
  '
  '【構文】　fs_条件日付("<日付初期値>","する・月・週・しない")
  '【戻り値】入力した日付。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError

  SS = fs_条件日付(Date, "月", "")

  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " と 入力しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub 日付_週範囲()
  '
  '【構文】　fs_条件日付("<日付初期値>","する・月・週・しない")
  '【戻り値】入力した日付。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError

  SS = fs_条件日付(Date, "週", "")

  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " と 入力しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub 年月入力()
  '
  '【構文】　fs_条件年月("<年月初期値>","しない","メッセージ")
  '【戻り値】入力した年月(yyyy/mm)。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError

  SS = fs_条件年月("", "しない", "更新年月を選択してください")

  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " と 入力しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub 年月範囲()
  '
  '【構文】　fs_条件年月("<年月初期値>","<範囲指定 する・しない>","メッセージ")
  '【戻り値】入力した年月(yyyy/mm ～ yyyy/mm)。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError

  SS = fs_条件年月("2020/10", "する", "集計する期間を指定してください")

  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " と 入力しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub 条件入力()
  '
  '【構文】　fs_条件入力("<条件初期値>","しない","条件名","<メッセージ>")
  '【戻り値】入力した条件。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError

  SS = fs_条件入力("", "しない", "条件入力", "")

  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " と 入力しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub 条件範囲()
  '
  '【構文】　fs_条件入力("<条件初期値>","しない","条件名","<メッセージ>")
  '【戻り値】入力した条件。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError

  SS = fs_条件入力("", "する", "条件入力", "")

  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " と 入力しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub セル入力()
  '
  '【構文】　fs_セル入力("<入力初期値>", 入力数, 入力名List, メッセージ)
  '【戻り値】入力した値(ygSTR1～ygSTR5) "値1,値2,..."。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError

  SS = fs_セル入力("", 3, "(1),(2),(3)", "")

  '(結果表示)
  If SS <> "" Then
    NN = fn_確認(SS & " と 入力しました。", "i", "確認")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Sub VBA表示()
  '
  '[Alt] + [F11] → VBA表示
  '

  SendKeys "%{F11}"
  
End Sub

Public Sub 新規YUGEbook作成()
  '世代数:qCnt
  Dim wNewFile  As String
  Dim wMSG      As String
  Dim wNo       As Integer
  Dim wCnt      As Integer
  Dim wRunPath  As String
  Dim wRunFile  As String
  Dim wBook     As String
  Dim wNewBook  As String
  Dim WSH       As Worksheet
  Dim wSheet    As String
 
  wRunPath = ThisWorkbook.Path
  wRunFile = ThisWorkbook.Name
  
  'ファイル名
  wCnt = 1
  wNewBook = "YUGE_Book"
  wNewFile = wRunPath & "\" & wNewBook & Trim(Str(wCnt)) & ".xlsm"
  
  fyLng1 = PathFileExists(wNewFile)
  
  Do Until fyLng1 = 0
    wCnt = wCnt + 1
    If wCnt > 100 Then Exit Do
    
    wNewFile = wRunPath & "\" & wNewBook & Trim(Str(wCnt)) & ".xlsm"
    fyLng1 = PathFileExists(wNewFile)
  Loop
  
  'Error
  If wCnt > 100 Then
    wMSG = "新規 YUGE_Bookを作成できませんでした。"
    NN = fn_確認(wMSG, "!", "確認")
    Exit Sub
  End If
  
  'Save
  Workbooks(wRunFile).SaveAs Filename:=wNewFile
  
  'Sheet削除
  For Each WSH In Workbooks(gBookName).Worksheets
    wSheet = WSH.Name
    NN = fn_シート削除()
  Next
End Sub

Private Sub 新規YUGEbook作成_bak()
  'Backupファイル/数
  'wFiles = wDTFile 'エクセルカードDT
     
  '最終Bakを削除
  'xlsm
  wBakFile = wBakPath & "\Bak" & Trim(Str(qCnt)) & "_" & wDTFile & ".xlsm"
  gLng0 = PathFileExists(wBakFile)
    
  If gLng0 <> 0 Then
    Kill wBakFile
  End If
  
  'xlsx
  wBakFile = wBakPath & "\Bak" & Trim(Str(qCnt)) & "_" & wDTFile & ".xlsx"
  gLng0 = PathFileExists(wBakFile)
  
  If gLng0 <> 0 Then
    Kill wBakFile
  End If
  
  'BakファイルをRename
  For wNo = qCnt - 1 To 1 Step -1
    'xlsm
    wBakFile = wBakPath & "\Bak" & Trim(Str(wNo)) & "_" & wDTFile & ".xlsm"
    gLng0 = PathFileExists(wBakFile)
      
    If gLng0 <> 0 Then
      wNewFile = wBakPath & "\Bak" & Trim(Str(wNo + 1)) & "_" & wDTFile & ".xlsm"
      Name wBakFile As wNewFile
    End If
    
    'xlsx
    wBakFile = wBakPath & "\Bak" & Trim(Str(wNo)) & "_" & wDTFile & ".xlsx"
    gLng0 = PathFileExists(wBakFile)
      
    If gLng0 <> 0 Then
      wNewFile = wBakPath & "\Bak" & Trim(Str(wNo + 1)) & "_" & wDTFile & ".xlsx"
      Name wBakFile As wNewFile
    End If
  Next wNo

End Sub


Public Sub マニュアル()
  '
  '【構文】　fs_セル入力("<入力初期値>", 入力数, 入力名List, メッセージ)
  '【戻り値】入力した値(ygSTR1～ygSTR5) "値1,値2,..."。[中止]したときは 空白 を返します。
  '
  On Error GoTo subError
  
  NN = fyBrowserOpen("https://excel-databace.hatenablog.com/entry/yuge-help")
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

Public Sub タイマー停止()
  On Error GoTo subError
  Dim wMsgStr As String
  Dim wMsgInt As Integer
  
  NN = fn_タイマー停止()
  
  If n = -1 Then
     'MsgBox "couldn't kill the timer"
     wMsgStr = "プログラムを終了して、再起動してください。" & vbCrLf & vbCrLf & _
               "よろしいですか？"
     wMsgInt = MsgBox(wMsgStr, vbYesNo + vbQuestion, "確認")
     
     If wMsgInt = vbYes Then
       '終了
       'ThisWorkbook.Save
       Workbooks(ygStartBook).Save
       Application.Quit
     End If
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub
