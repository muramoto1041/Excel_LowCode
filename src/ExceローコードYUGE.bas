Attribute VB_Name = "ExcelローコードYUGE"
'
' ExcelローコードYUGE Ver1.33
' (注意)変更しないでください。
' https://excel-databace.hatenablog.com/entry/yuge-help
'
'Ver 1.17  2020/11/16  タイマー機能追加
'    1.18  2020/11/25  ユーザーマクロに説明追加/年月範囲追加/印刷ダイアログ
'    1.19  2020/12/25  テーブル作成/解除/ygSetTimerに変更/メニュー大小 > SetTimer
'    1.20  2021/04/28  fn_ファイル判定bugfix
'    1.20  2021/05/13  ファイルオープンでパスを判定/ADO追加
'    1.21  2021/05/28  fs_条件入力、fy_条件範囲入力追加
'    1.22  2021/09/18  fs_条件日付(日付名)
'    1.23  2021/11/24  Databaseコマンドの頭文字をfa_に変更
'    1.24  2022/02/23  fs_条件入力(条件名)拡張/Fyメニュー/Fyメニュー小/Fyメニュー大Bugfix
'    1.25  2022/03/19  囲み文字列関数
'    1.26  2022/07/05  fs_シートメニューを非表示シート対応にしました
'    1.27  2022/07/14  初期化領域1-5,初期化領域M1-M5
'    1.28  2022/08/10  セル入力(1-5)
'    1.29  2022/09/14  fa_転置Array/fyConvRange_Col を追加
'    1.30  2022/10/28  syADOAccdbOpenにMDB引数を追加/fn_処理中Open/fn_処理中Closeを追加
'    1.31  2023/07/07  Fy_印刷、fyTimeStampを追加
'    1.32  2024/04/11  fn_確認/fyTimeStamp bugfix/fy_→fn_fsに変更
'    1.33  2024/04/23  Fy_条件日付範囲_週/fyBrowserOpen/fyFileRunを追加/fyIsSheet・fyIsVisibleを変更

'[参照設定を有効にする]
' Microsoft Scripting Runtime（FileSystemObjectを使う）
' Microsoft ActiveX Data Object 2.7 Library（ADODB.Connectionを使う）
Option Explicit

'----- 戻り値取得変数 -----
Public NN As Long, SS As String

'----- YUGE共通変数(汎用) -----
Public ygInt1 As Integer, ygInt2  As Integer, ygInt3  As Integer, ygInt4  As Integer, ygInt5 As Integer
Public ygSTR1 As String, ygSTR2   As String, ygSTR3   As String, ygSTR4   As String, ygSTR5  As String
Public ygLBL1 As String, ygLBL2   As String, ygLBL3   As String, ygLBL4   As String, ygLBL5  As String
Public yg条件 As String, yg条件ST As String, yg条件ED As String, yg条件名 As String
Public ygLng1 As Long, ygLng2  As Long, ygLng3  As Long, ygLng4  As Long, ygLng5 As Long
Public ygMSG  As String
Public ygEnd  As Integer
Public ygBackForm  As String
Public ygStartBook As String
'----- YUGE共通変数(印刷) -----
Public ygCntPrt       As Integer
Public yg印刷Sheet    As String
Public yg初期化領域1  As String, yg初期化領域2  As String, yg初期化領域3  As String, yg初期化領域4  As String, yg初期化領域5  As String
Public yg初期化領域M1 As String, yg初期化領域M2 As String, yg初期化領域M3 As String, yg初期化領域M4 As String, yg初期化領域M5 As String
'----- YUGE共通変数(日付) -----
Public yg日付 As String, yg日付ST As String, yg日付ED As String, yg日付名 As String
Public yg年月 As String, yg年月ST As String, yg年月ED As String
Public yg年   As Integer, yg年ST  As Integer, yg年ED  As Integer
'----- YUGE共通変数(入力) -----
Public yg表題01 As String, yg表題02 As String, yg表題03 As String, yg表題04 As String, yg表題05 As String
'--- Timer ---
Public ygBlnTimer    As Boolean
Public ygLngTimerID  As Long
Public ygProcMacro   As String

'MkPassWord/RePassWord用(数字4桁を入力してください)
Private Const ygPassWord = "1234"

'----- YUGE Database -----
Public ygMDBPath As String
'ADO
Public ADB   As ADODB.Connection
Public ARST1 As ADODB.Recordset
Public ARST2 As ADODB.Recordset
Public ARST3 As ADODB.Recordset
Public ARST4 As ADODB.Recordset
Public ARST5 As ADODB.Recordset
'SQL
Public SQL1 As String
Public SQL2 As String
Public SQL3 As String
Public SQL4 As String
Public SQL5 As String

'----- DataBase -----
Public ygArrayDB As Variant

'2020/11/13 --------------------------------------------------*
Type SYSTEMTIME
    wYear As Integer         '現在の年
    wMonth As Integer        '月(1月=1, 2月=2)
    wDayOfWeek As Integer    '曜日(日曜=0, 月曜=1)
    wDay As Integer          '日
    wHour As Integer         '時
    wMinute As Integer       '分
    wSecond As Integer       '秒
    wMilliseconds As Integer 'ﾐﾘ秒
End Type

#If VBA7 Then
'(VBA7)
'Timer
Declare PtrSafe Function SetTimer Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As Long

Declare PtrSafe Function KillTimer Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal nIDEvent As Long) As Long
'GetLocalTime
Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'<PathFileExists> 指定ファイルの存在チェック
Declare PtrSafe Function PathFileExists Lib "SHLWAPI.DLL" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

#Else
'(Downlevel when using previous version of VBA7)
'Timer
Declare Function SetTimer Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

Declare Function KillTimer Lib "user32" _
       (ByVal hwnd As Long, _
        ByVal nIDEvent As Long) As Long
'GetLocalTime
Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'<PathFileExists> 指定ファイルの存在チェック
Declare Function PathFileExists Lib "SHLWAPI.DLL" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
#End If

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
  'タイマー 実行処理
  On Error Resume Next
  
  Application.Run ygProcMacro
End Sub

'2023/06/15 ----------------------------------------*
Public Function fyTimeStamp() As Currency
  Dim tm   As SYSTEMTIME
  Dim wSTR As String
  
  Call GetLocalTime(tm)
  
  wSTR = Format(tm.wYear Mod 100, "00") & Format(tm.wMonth, "00") & Format(tm.wDay, "00") & _
         Format(tm.wHour, "00") & Format(tm.wMinute, "00") & Format(tm.wSecond, "00") & Format(Int(tm.wMilliseconds / 10), "00")
  fyTimeStamp = Val(wSTR)
End Function

'2024/04/11 --------------------------------------------------*
Function fn_確認(qメッセージ As String, qアイコン As String, qタイトル As String) As Integer
  '
  '【構文】  Fn_確認("<メッセージ>","<アイコンタイプ>","<タイトル>")
  '【指定値】<アイコンタイプ> i ? ! x
  '【戻り値】［OK］［x］は 0、［はい］は 1、［いいえ］は 2
  '
  Dim wTitle  As String
  Dim wNN     As Integer
  
  If qタイトル = "" Then
    wTitle = "確認"
  Else
    wTitle = qタイトル
  End If
  
  Select Case qアイコン
    Case "i":  wNN = MsgBox(qメッセージ, vbOKOnly + vbInformation, wTitle)
    Case "?":  wNN = MsgBox(qメッセージ, vbYesNo + vbQuestion, wTitle)
    Case "!":  wNN = MsgBox(qメッセージ, vbOKOnly + vbExclamation, wTitle)
    Case "x":  wNN = MsgBox(qメッセージ, vbYesNo + vbCritical, wTitle)
    Case Else: wNN = MsgBox("アイコンタイプは、[i][?][!][x]で指定してください。", vbOKOnly + vbExclamation, "確認")
  End Select
  
  Select Case wNN
    Case 6: fn_確認 = 1
    Case 7: fn_確認 = 2
    Case Else: fn_確認 = 0
  End Select
End Function

'2020/11/26 --------------------------------------------------*
Function fn_シート印刷(qシート名 As String, qブック名 As String, qプレビュー処理 As String, q印刷枚数 As Integer) As Integer
  '
  '【構文】  fn_シート印刷("<シート名>","<ブック名>,"<プレビュー処理>", <印刷枚数>)
  '【戻り値】成功:1 失敗:0
  '
  Dim wFL     As Integer
  Dim wBook   As String
  Dim wIsPreview As String
  Dim wBLret  As Boolean
  
  fn_シート印刷 = 0
  
  'Book名省略
  If qブック名 = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qブック名
  End If
  
  'Bookを判定
  wFL = fn_ファイル判定(wBook, True)
  If wFL = 0 Then Exit Function
  
  'Sheetを判定
  wFL = fn_シート判定(qシート名, qブック名, False, True)
  If wFL = 0 Then Exit Function
  
  '印刷処理
  With Workbooks(wBook)
  
    '自動計算チェック
    'Application.Calculation = xlCalculationAutomatic
    .Worksheets(qシート名).Activate
    
    'プレビュー処理を省略
    If qプレビュー処理 = "" Then
      wIsPreview = "する"
    Else
      wIsPreview = qプレビュー処理
    End If
    
    Select Case wIsPreview
      Case "する"
        'プレビュー
        .Worksheets(qシート名).PrintOut preview:=True
        
      Case Else
        If ygCntPrt = 1 Then
          '(1枚目)印刷ダイアログ
          wBLret = Application.Dialogs(xlDialogPrint).Show(Arg4:=q印刷枚数)
          
          '中止
          If wBLret = False Then ygCntPrt = -1
        Else
          '(2枚目以降)印刷実行
          .Worksheets(qシート名).PrintOut Copies:=q印刷枚数
        End If
    End Select
  End With
  
  '初期状態 Close
  If wFL = 2 Then
    Workbooks(wBook).Close savechanges:=False
  End If
  
  fn_シート印刷 = 1
End Function

'2020/08/08 --------------------------------------------------*
Function fn_シート削除(qシート名 As String, qブック名 As String, q警告 As String) As Integer
  '
  '【構文】  fn_シート追加("<シート名>","<ブック名>")
  '【戻り値】成功:1 失敗:0
  '
  Dim wFL     As Integer
  Dim wBook   As String
  Dim WSH    As Worksheet

  fn_シート削除 = 0
  
  'Book名省略
  If qブック名 = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qブック名
  End If
  
  'Bookを判定
  wFL = fn_ファイル判定(wBook, True)
  If wFL = 0 Then Exit Function
  
  'オープン判定
  'シート名省略
  
  '警告表示
  If q警告 = "する" Or q警告 = "" Then
    Application.DisplayAlerts = True
  Else
    Application.DisplayAlerts = False
  End If
  
  'シート判定
  wFL = fn_シート判定(qシート名, qブック名, False, False)
  If wFL = 0 Then Exit Function
  
  'シート削除
  With Workbooks(wBook)
    .Worksheets(qシート名).Activate
    .Worksheets(qシート名).Delete
  End With

  Application.DisplayAlerts = True
  
  fn_シート削除 = 1
  
  For Each WSH In Workbooks(wBook).Worksheets
    If qシート名 = WSH.Name Then
      '(存在:キャンセル)
      fn_シート削除 = 0
      Exit For
    End If
  Next
End Function

'2020/06/03 --------------------------------------------------*
Function fn_シート判定(qシート名 As String, qブック名 As String, q閉じる As Boolean, q警告 As Boolean) As Integer
  '
  '【構文】  fn_シート判定("<シート名>","<ブック名>",<終了状態>,<エラー表示指定>)
  '【指定値】<終了状態>       True/False
  '          <エラー表示指定> True/False
  '【戻り値】成功(初期Open):1 成功(初期Close):2 失敗:0
  '
  Dim wFL     As Integer
  Dim wFL2    As Integer
  Dim wFLMsg  As Integer
  Dim wMSG    As String
  Dim wBook   As String
  Dim wPath   As String
  Dim wFile   As String
  Dim wExt    As String
  Dim wVar    As Variant

  fn_シート判定 = 0

  'シート名を省略
  If qシート名 = "" Then
    wMSG = "シート名を指定してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If

  'Book名を省略
  If qブック名 = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qブック名
  End If
  
  'パス指定がない
  If InStr(wBook, "\") = 0 Then
    wPath = ThisWorkbook.Path
    wFile = wPath & "\" & wBook
  Else
    wFile = wBook
  End If
  
  wFL = fn_ファイル判定(wFile, q警告)
  If wFL = 0 Then Exit Function
  
  'ExcelBookを判定(xlsx/xlsm)
  
  'Book Open を判定
  wBook = fyPickFile(wFile)
  
  wFL = fn_ブックオープン判定(wBook, False)
  If wFL = 0 Then
    wMSG = "ブック（" & wBook & "）が、開いていません。" & vbCrLf & _
           "オープンしますか？"
    wFLMsg = MsgBox(wMSG, vbYesNo + vbQuestion, "確認")
    If wFLMsg = vbYes Then
      'Workbooks.Open FileName:=wFile
      Set wVar = Workbooks.Open(wFile)
      
      If TypeName(wVar) <> "Workbook" Then
        wMSG = "オープンできません。"
        MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
        Exit Function
      End If
    Else
      Exit Function
    End If
  End If
  
  '警告
  If q警告 = True Then
    wFL2 = 1
  Else
    wFL2 = 0
  End If
  
  If fyIsSheet(qシート名, wBook, wFL2) = True Then
    If wFL = 0 Then
      '初期Close
      fn_シート判定 = 2
    Else
      '初期Open
      fn_シート判定 = 1
    End If
  End If
  
  'Bookを閉じる
  If wFL = 1 And q閉じる = True Then
    Workbooks(wBook).Close savechanges:=False
  End If
End Function

'2020/07/20 --------------------------------------------------*
Function fn_シート追加(qシート名 As String, qブック名 As String) As Integer
  '
  '【構文】  fn_シート追加("<シート名>","<ブック名>")
  '【戻り値】成功:1 失敗:0
  '
  Dim wFL     As Integer
  Dim wBook   As String
  
  
  fn_シート追加 = 0
  
  'Book名省略
  If qブック名 = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qブック名
  
    'Bookを判定
    wFL = fn_ファイル判定(wBook, True)
    If wFL = 0 Then Exit Function
  End If
  
  'シート追加
  With Workbooks(wBook)
    .Worksheets.Add
    
    'シート名指定
    If qシート名 <> "" Then
      If fyIsSheet(qシート名, "", 0) = False Then
        ActiveSheet.Name = qシート名
      End If
    End If
  End With

  fn_シート追加 = 1
End Function

'2024/04/24 --------------------------------------------------*
Function fn_シート表示(qシート名 As String, qブック名 As String, qエラー表示 As Boolean) As Integer
  '
  '【構文】  fn_シート表示("<シート名>","<ブック名>","<エラー表示>")
  '【指定値】<エラー表示> True/False
  '【戻り値】成功:1 失敗:0
  '
  Dim wFL     As Integer
  Dim wBook   As String
  
  fn_シート表示 = 0
  
  'Book名省略
  If qブック名 = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qブック名
  End If
  
  'Bookを判定
  wFL = fn_ファイル判定(wBook, True)
  If wFL = 0 Then Exit Function
  
  'Sheetを判定
  wFL = fn_シート判定(qシート名, qブック名, False, qエラー表示)
  If wFL = 0 Then Exit Function
  
  'シート表示
  With Workbooks(wBook)
    If fyIsSheet(qシート名, "", 0) = True Then
      .Worksheets(qシート名).Activate
    End If
  End With
  
  fn_シート表示 = 1
End Function

'2022/07/14 --------------------------------------------------*
Function fn_シート初期化(qシート名 As String, qブック名 As String) As Integer
  '
  '【構文】  fn_シート初期化("<シート名>","<ブック名>")
  '【戻り値】成功:1 失敗:0
  '
  Dim wFL     As Integer
  Dim wBook   As String
  
  fn_シート初期化 = 0
  
  'Book名省略
  If qブック名 = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qブック名
  End If

  'Bookを判定
  wFL = fn_ファイル判定(wBook, True)
  If wFL = 0 Then Exit Function

  'Sheetを判定
  wFL = fn_シート判定(qシート名, qブック名, False, True)
  If wFL = 0 Then Exit Function
  
  'シート初期化
  With Workbooks(wBook).Worksheets(qシート名)
    'シート初期化
    If yg初期化領域1 <> "" Then .Range(yg初期化領域1).ClearContents
    If yg初期化領域2 <> "" Then .Range(yg初期化領域2).ClearContents
    If yg初期化領域3 <> "" Then .Range(yg初期化領域3).ClearContents
    'シート初期化(結合セル)
    If yg初期化領域M1 <> "" Then .Range(yg初期化領域M1).MergeArea.ClearContents
    If yg初期化領域M2 <> "" Then .Range(yg初期化領域M2).MergeArea.ClearContents
    If yg初期化領域M3 <> "" Then .Range(yg初期化領域M3).MergeArea.ClearContents
  End With

  fn_シート初期化 = 1
End Function

'2022/07/14 --------------------------------------------------*
Function fn_初期化変数クリア() As Integer
  '
  '【構文】  fn_初期化変数クリア()
  '【戻り値】成功:1 失敗:0
  '
  yg初期化領域1 = "":  yg初期化領域2 = "":  yg初期化領域2 = ""
  yg初期化領域M1 = "": yg初期化領域M2 = "": yg初期化領域M2 = ""

  fn_初期化変数クリア = 1
End Function

'2020/07/26 --------------------------------------------------*
Function fs_シートメニュー(qクリック表示 As String, q非表示シート As String, qメニュー初期値 As Integer, qBL非表示 As Boolean) As String
  '
  '【構文】　fs_シートメニュー("<クリック表示>","<非表示シート>",メニュー初期値, True:全て表示/False:非表示シートを除く)
  '
  '【戻り値】選択したシート名。[中止]したときは 空白 を返します。
  '
  Dim WSH         As Worksheet
  Dim wMyBook     As String
  Dim wMySheet    As String
  Dim w非表示     As String
  Dim wSName      As String
  Dim wNo         As Integer
  Dim wMenuList   As String
  Dim wSheet(100) As String
  Dim wBLdisp     As Boolean
  
  wMyBook = ThisWorkbook.Name
  w非表示 = "," & q非表示シート & ",,,"
  wMenuList = ""
  wNo = 0
  
  For Each WSH In Workbooks(wMyBook).Worksheets
    wSName = WSH.Name
    wBLdisp = WSH.Visible
    
    'w非表示Sheetは、リストに追加しない
    If InStr(w非表示, "," & wSName & ",") = 0 Then
      Select Case qBL非表示
        'すべて表示する
        Case True
          wNo = wNo + 1
          wSheet(wNo) = wSName
          wMenuList = wMenuList & wSName & ","
        
        '非表示シートを除く
        Case False
          If wBLdisp = True Then
            wNo = wNo + 1
            wSheet(wNo) = wSName
            wMenuList = wMenuList & wSName & ","
          End If
        End Select
    End If
  Next
  
  If wMenuList <> "" Then
    wMenuList = Left(wMenuList, Len(wMenuList) - 1)
  End If
  
  '【Fy_メニュー】
  ygSTR1 = "$シートメニュー$"
  ygSTR2 = wMenuList
  ygSTR3 = "シートを選択してください"
  ygSTR4 = qクリック表示
  ygInt1 = qメニュー初期値
  Fy_メニューM.Show vbModal
  
  fs_シートメニュー = wSheet(ygInt1)

End Function

'2022/08/10 --------------------------------------------------*
Function fs_セル入力(q入力初期値List As String, q入力数 As Integer, q入力表題List As String, qメッセージ As String) As String
  '
  '【構文】　fs_セル入力("<入力初期値>", 入力数, 入力名List, 入力RangeList, メッセージ)
  '【戻り値】入力した値(ygSTR1～ygSTR5)。[OK]したときは "値1,値2,..." / [中止]したときは 空白 を返します。
  '
  Dim wMSG        As String
  Dim wArray表題 As Variant, wArraySTR  As Variant
  Dim w表題1 As String, w表題2 As String, w表題3 As String, w表題4 As String, w表題5 As String
  Dim wSTR1  As String, wSTR2  As String, wSTR3  As String, wSTR4  As String, wSTR5  As String
  Dim wNo    As Integer, wCnt表題 As Integer, wCntSTR As Integer
  
  fs_セル入力 = ""
  
  '(入力数Check)
  If q入力数 < 1 Or q入力数 > 5 Then
    wMSG = "入力数は、1 ～ 5 を指定してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  '初期化
  ygSTR1 = "": ygSTR2 = "": ygSTR3 = "": ygSTR4 = "": ygSTR5 = ""
  w表題1 = "": w表題2 = "": w表題3 = "": w表題4 = "": w表題5 = ""
  wSTR1 = "":  wSTR2 = "":  wSTR3 = "":  wSTR4 = "":  wSTR5 = ""
  
  '表題
  wArray表題 = Split(q入力表題List, ",")
  wCnt表題 = UBound(wArray表題) + 1
  
  For wNo = 1 To wCnt表題
    If wNo > 5 Then Exit For
    
    Select Case wNo
      Case 1: w表題1 = wArray表題(wNo - 1)
      Case 2: w表題2 = wArray表題(wNo - 1)
      Case 3: w表題3 = wArray表題(wNo - 1)
      Case 4: w表題4 = wArray表題(wNo - 1)
      Case 5: w表題5 = wArray表題(wNo - 1)
    End Select
  Next wNo
  
  '初期値
  wArraySTR = Split(q入力初期値List, ",")
  wCntSTR = UBound(wArraySTR) + 1
  
  For wNo = 0 To wCntSTR
    If wNo > 5 Then Exit For
    
    Select Case wNo
      Case 1: wSTR1 = wArraySTR(wNo - 1)
      Case 2: wSTR2 = wArraySTR(wNo - 1)
      Case 3: wSTR3 = wArraySTR(wNo - 1)
      Case 4: wSTR4 = wArraySTR(wNo - 1)
      Case 5: wSTR5 = wArraySTR(wNo - 1)
    End Select
  Next wNo
  
  '初期値/ラベル
  ygSTR1 = wSTR1: ygLBL1 = w表題1
  ygSTR2 = wSTR2: ygLBL2 = w表題2
  ygSTR3 = wSTR3: ygLBL3 = w表題3
  ygSTR4 = wSTR4: ygLBL4 = w表題4
  ygSTR5 = wSTR5: ygLBL5 = w表題5
  ygMSG = qメッセージ
  
  '入力フォーム
  Select Case q入力数
    Case 1: Fy_セル入力1.Show vbModal
    Case 2: Fy_セル入力2.Show vbModal
    Case 3: Fy_セル入力3.Show vbModal
    Case 4: Fy_セル入力4.Show vbModal
    Case 5: Fy_セル入力5.Show vbModal
  End Select
  
  '中止
  If ygEnd = 0 Then Exit Function
  
  '入力フォーム
  Select Case q入力数
    Case 1
      fs_セル入力 = Trim(ygSTR1)
    Case 2
      fs_セル入力 = Trim(ygSTR1) & "," & Trim(ygSTR2)
    Case 3
      fs_セル入力 = Trim(ygSTR1) & "," & Trim(ygSTR2) & "," & Trim(ygSTR3)
    Case 4
      fs_セル入力 = Trim(ygSTR1) & "," & Trim(ygSTR2) & "," & Trim(ygSTR3) & "," & Trim(ygSTR4)
    Case 5
      fs_セル入力 = Trim(ygSTR1) & "," & Trim(ygSTR2) & "," & Trim(ygSTR3) & "," & Trim(ygSTR4) & "," & Trim(ygSTR5)
  End Select
  
End Function

'2021/05/24 --------------------------------------------------*
Function fs_条件入力(q条件初期値 As String, q範囲指定 As String, q条件名 As String, qメッセージ As String) As String
  '
  '【構文】　fs_条件入力("<入力初期値>","範囲指定")
  '【戻り値】入力した値。[中止]したときは 空白 を返します。
  '
  Dim wMSG  As String
  
  fs_条件入力 = ""
  
  '初期化
  yg条件 = "": yg条件ST = "": yg条件ED = "": ygSTR1 = qメッセージ: yg条件名 = q条件名
  
  Select Case True
    'しない（入力入力）
    Case (q範囲指定 = "しない" Or q範囲指定 = "")
      yg条件 = q条件初期値
      Fy_条件入力_Bs.Show vbModal
      If ygEnd = 0 Then Exit Function
      
      fs_条件入力 = yg条件
    
    'する（条件範囲入力）
    Case (q範囲指定 = "する")
      yg条件 = q条件初期値
      Fy_条件入力範囲.Show vbModal
      If ygEnd = 0 Then Exit Function
      
      fs_条件入力 = yg条件ST & " ～ " & yg条件ED
    
    'エラー
    Case Else
      wMSG = "範囲指定は、'する' 'しない' を指定してください。（省略可）"
      MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
  End Select
End Function

'2024/04/16 --------------------------------------------------*
Function fs_条件日付(q日付初期値 As String, q範囲指定 As String, q日付名 As String) As String
  '
  '【構文】　fs_条件日付("<日付初期値>","範囲指定")
  '【戻り値】入力した日付。[中止]したときは 空白 を返します。
  '
  Dim wMSG  As String
  
  fs_条件日付 = ""
  
  '初期化
  yg日付 = "": yg日付ST = "": yg日付ED = "": yg日付名 = q日付名
  
  Select Case True
    'しない（日付入力）
    Case (q範囲指定 = "しない" Or q範囲指定 = "")
      yg日付 = q日付初期値
      Fy_条件日付_Bs.Show vbModal
      If ygEnd = 0 Then Exit Function
      
      fs_条件日付 = yg日付
    
    'する（日付範囲入力 月）
    Case (q範囲指定 = "する" Or q範囲指定 = "月")
      yg日付ST = q日付初期値
      Fy_条件日付範囲_月.Show vbModal
      If ygEnd = 0 Then Exit Function
      
      fs_条件日付 = yg日付ST & " ～ " & yg日付ED
    
    'する（日付範囲入力 週）
    Case (q範囲指定 = "週")
      yg日付ST = q日付初期値
      Fy_条件日付範囲_週.Show vbModal
      If ygEnd = 0 Then Exit Function
      
      fs_条件日付 = yg日付ST & " ～ " & yg日付ED
    
    'エラー
    Case Else
      wMSG = "範囲指定は、'する' '月' '週' 'しない' を指定してください。（省略可）"
      MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
  End Select
End Function

'2020/10/06 --------------------------------------------------*
Function fs_条件年月(q年月初期値 As String, q範囲指定 As String, qメッセージ As String) As String
  '
  '【構文】　fs_条件年月("<年月初期値>","範囲指定","メッセージ")
  '【戻り値】入力した年月(yyyy/mm)。[中止]したときは 空白 を返します。
  '
  Dim wMSG  As String
  
  fs_条件年月 = ""
  
  '初期化
  yg年月 = "": yg年月ST = "": yg年月ED = ""
  
  'メッセージ
  ygSTR1 = qメッセージ
  
  Select Case True
    'しない（年月入力）
    Case (q範囲指定 = "しない" Or q範囲指定 = "")
      yg年月 = q年月初期値
      Fy_条件年月_BS.Show vbModal
      If ygEnd = 0 Then Exit Function
      
      fs_条件年月 = yg年月
    
    'する（年月範囲入力）
    Case (q範囲指定 = "する")
      yg年月 = q年月初期値
      Fy_条件年月範囲.Show vbModal
      If ygEnd = 0 Then Exit Function
      
      fs_条件年月 = yg年月ST & " ～ " & yg年月ED
    
    'エラー
    Case Else
      wMSG = "範囲指定は、'する' 'しない' を指定してください。（省略可）"
      MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
  End Select
End Function

'2020/06/01 --------------------------------------------------*
Function fs_ダイヤログ(q取得属性 As String, q初期フォルダ As String) As String
  '
  '【構文】  fs_ダイヤログ("<取得属性>","<初期フォルダ>")
  '【戻り値】成功:ファイル名/フォルダ名 失敗:空白
  '
  Dim wIniFilename  As String
  
  fs_ダイヤログ = ""
  
  If Not (q取得属性 = "ファイル" Or q取得属性 = "フォルダ") Then Exit Function
  
  If q初期フォルダ = "" Then
    wIniFilename = ThisWorkbook.Path
  Else
    wIniFilename = q初期フォルダ
  End If
  
  'フォルダ
  If q取得属性 = "フォルダ" Then
    With Application.FileDialog(msoFileDialogFolderPicker)
      .InitialFileName = ""
      .Title = "フォルダ ダイイアログ"
      If .Show = True Then
        fs_ダイヤログ = .SelectedItems(1)
      End If
    End With
  End If
  
  'ファイル
  If q取得属性 = "ファイル" Then
    With Application.FileDialog(msoFileDialogOpen)
      .Title = "ファイル ダイイアログ"
      If .Show = True Then
          fs_ダイヤログ = .SelectedItems(1)
      End If
    End With
  End If
  
End Function

'2020/06/05 --------------------------------------------------*
Function fn_ファイルオープン(qファイル名 As String, qエラー表示指定 As Boolean) As Integer
  '
  '【構文】  fn_ファイルオープン("<ファイル名>",<エラー表示指定>)
  '【戻り値】成功:1 失敗:0 見つからない:-1
  '
  Dim wFL     As Integer
  Dim wMSG    As String
  Dim wPath   As String
  Dim wFile   As String
  Dim WSH
  
  fn_ファイルオープン = 0
  
  'ファイル名を省略
  If qファイル名 = "" Then
    wMSG = "ファイルを指定してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  '拡張子がない
  If InStr(qファイル名, ".") = 0 Then
    wMSG = "ファイル名に拡張子を記述してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  'パス指定がない
  If InStr(qファイル名, "\") = 0 Then
    wPath = ThisWorkbook.Path
    wFile = wPath & "\" & qファイル名
  Else
    wFile = qファイル名
  End If
  
  'ファイルがない
  If Dir(wFile) = "" Then
    If qエラー表示指定 = True Then
      wMSG = "フォルダ（" & fyPickFolder(wFile) & "）に" & vbCrLf & _
             "ファイル（" & fyPickFile(wFile) & "）が、見つかりません。"
      MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    End If
    Exit Function
  End If
  
  wFL = fn_ファイル判定(qファイル名, qエラー表示指定)
  If wFL = 0 Then
    fn_ファイルオープン = -1
    Exit Function
  End If
  
  Set WSH = CreateObject("Wscript.Shell")
  WSH.Run wFile, 3   '3:最大化された状態で起動します。選択状態になります。
  Set WSH = Nothing
  
  fn_ファイルオープン = 1
End Function

'2020/06/03 --------------------------------------------------*
Function fn_ファイル判定(qファイル名 As String, qエラー表示指定 As Boolean) As Integer
  '
  '【構文】  fn_ファイル判定("<ファイル名>",<エラー表示指定>)
  '【戻り値】成功:1 失敗:0
  '
  Dim wMSG    As String
  Dim wPath   As String
  Dim wFile   As String
  Dim wFolder As String
  
  fn_ファイル判定 = 0
  
  'ファイル名を省略
  If qファイル名 = "" Then
    wMSG = "ファイルを指定してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  '拡張子がない
  If InStr(qファイル名, ".") = 0 Then
    wMSG = "ファイル名に拡張子を記述してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  'パス指定がない
  If InStr(qファイル名, "\") = 0 Then
    wPath = ThisWorkbook.Path
    wFile = wPath & "\" & qファイル名
  Else
    wFile = qファイル名
  End If
  
  'ファイルがない
  If Dir(wFile) = "" Then
    If qエラー表示指定 = True Then
      wMSG = "フォルダ（" & fyPickFolder(wFile) & "）に" & vbCrLf & _
             "ファイル（" & fyPickFile(wFile) & "）が、見つかりません。"
      MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    End If
    Exit Function
  End If
  
  fn_ファイル判定 = 1
End Function

'2020/06/03 --------------------------------------------------*
Function fn_ブックオープン判定(qブック名 As String, qエラー表示指定 As Boolean) As Integer
  '
  '【構文】  fn_ブックオープン判定("<ブック名>",<エラー表示指定>)
  '【戻り値】成功:1 失敗:0
  '
  Dim wMSG    As String
  Dim wBook   As Workbook
  
  fn_ブックオープン判定 = 0
  
  'Book名を省略
  If qブック名 = "" Then
    wMSG = "ブック名を指定してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  'OpenBookをチェック
  For Each wBook In Workbooks
    If wBook.Name = qブック名 Then
      fn_ブックオープン判定 = 1
      Exit Function
    End If
  Next wBook
  
  'Not Open
  If qエラー表示指定 = True Then
    wMSG = "フォルダ（" & qブック名 & "）は、開いていません。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
  End If
End Function

'2020/07/20 --------------------------------------------------*
Function fn_メニュー中(qタイトル As String, qメニューリスト As String, qメッセージ As String, qメニュー初期値 As Integer) As Integer
  '
  '【構文】　fn_メニュー中("<タイトル>","<メニュー1>,<メニュー2>,・・・","<メーッセージ>",メニュー初期値)
  '【戻り値】選択したメニュー番号。[中止]したときは 0 を返します。
  '
  ygSTR1 = qタイトル
  ygSTR2 = qメニューリスト
  ygSTR3 = qメッセージ
  ygInt1 = qメニュー初期値
  Fy_メニューM.Show vbModal
  
  fn_メニュー中 = ygInt1
End Function

'2021/01/22 --------------------------------------------------*
Function fn_メニュー小(qタイトル As String, qメニューリスト As String, qメッセージ As String, qメニュー初期値 As Integer) As Integer
  '
  '【構文】　fn_メニュー("<タイトル>","<メニュー1>,<メニュー2>,・・・","<メーッセージ>",メニュー初期値)
  '【戻り値】選択したメニュー番号。[中止]したときは 0 を返します。
  '
  ygSTR1 = qタイトル
  ygSTR2 = qメニューリスト
  ygSTR3 = qメッセージ
  ygInt1 = qメニュー初期値
  Fy_メニューS.Show vbModal
  
  fn_メニュー小 = ygInt1
End Function

'2021/01/22 --------------------------------------------------*
Function fn_メニュー大(qタイトル As String, qメニューリスト As String, qメッセージ As String, qメニュー初期値 As Integer) As Integer
  '
  '【構文】　fn_メニュー("<タイトル>","<メニュー1>,<メニュー2>,・・・","<メーッセージ>",メニュー初期値)
  '【戻り値】選択したメニュー番号。[中止]したときは 0 を返します。
  '
  ygSTR1 = qタイトル
  ygSTR2 = qメニューリスト
  ygSTR3 = qメッセージ
  ygInt1 = qメニュー初期値
  Fy_メニューL.Show vbModal
  
  fn_メニュー大 = ygInt1
End Function

'2020/11/13 --------------------------------------------------*
Function fn_タイマー開始(q間隔秒 As Long, qマクロ名 As String) As Long
  '
  '【構文】　fn_タイマー開始(間隔(秒))
  '【戻り値】[失敗]したときは 0 を返します。成功した時は、0 以外
  '
  On Error GoTo subError
  Dim wMsgStr   As String
  Dim wElapse   As Long
  Dim wTimeST   As String
  Dim wTimeED   As String
  Dim w更新秒   As Long
  
  fn_タイマー開始 = 0
  
  '--- エラーチェック ---
  If q間隔秒 <= 0 Then
    wMsgStr = "タイマー間隔(秒)を設定してください。"
    MsgBox wMsgStr, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  If qマクロ名 = "" Then
    wMsgStr = "マクロ名を指定してください。"
    MsgBox wMsgStr, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  '更新時間をチェック
  wTimeST = Format(Time, "hh:mm:ss")
  
  '(マクロ実行)
  ygProcMacro = qマクロ名
  Application.Run ygProcMacro
  
  wTimeED = Format(Time, "hh:mm:ss")
  w更新秒 = (Hour(wTimeED) * 60 * 60 + Minute(wTimeED) * 60 + Second(wTimeED) + 5) - _
            (Hour(wTimeST) * 60 * 60 + Minute(wTimeST) * 60 + Second(wTimeST))
  
  If w更新秒 > q間隔秒 Then
    wMsgStr = "タイマーを停止します。" & vbCrLf & vbCrLf & _
              "更新間隔を " & Str(w更新秒) & " 秒 以上に、設定してください。"
    MsgBox wMsgStr, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  wElapse = q間隔秒 * 1000
  
  '(開始)
  ygLngTimerID = SetTimer(0, 0, wElapse, AddressOf TimerProc)
            
  If ygLngTimerID = 0 Then
    MsgBox "タイマーをセットできませんでした。プログラムを再起動してください。" & vbCrLf & vbCrLf & _
           "Timer not created. Ending Program"
           
    fn_タイマー開始 = -1
    GoTo subExit
  End If
  
  fn_タイマー開始 = ygLngTimerID
  ygBlnTimer = True
  
subExit:
  Exit Function

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : タイマー開始"
  Resume subExit
End Function

'2020/11/13 --------------------------------------------------*
Function fn_タイマー停止() As Long
  '
  '【構文】　fn_タイマー停止()
  '【戻り値】既に[停止]しているときは 0、成功した時は 1、[失敗]したときは -1 を返します。
  '
  On Error GoTo subError
  Dim wMsgStr As String
  
  fn_タイマー停止 = 0
  
  If ygBlnTimer = False Then
    wMsgStr = "タイマーは、停止しています。"
    MsgBox wMsgStr, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  '(停止)
  ygLngTimerID = KillTimer(0, ygLngTimerID)
  
  If ygLngTimerID = 0 Then
    fn_タイマー停止 = -1
  Else
    fn_タイマー停止 = ygLngTimerID
  End If
  
  ygBlnTimer = False
  
subExit:
  Exit Function

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : タイマー停止"
  Resume subExit
End Function

'2022/11/17 --------------------------------------------------*
Function fn_処理中Open() As Integer
  fn_処理中Open = 0
  Fy_処理中.Show:  Fy_処理中.Repaint
  fn_処理中Open = 1
End Function

'2022/11/17 --------------------------------------------------*
Function fn_処理中Close() As Integer
  fn_処理中Close = 0
  Unload Fy_処理中
  fn_処理中Close = 1
End Function

'************************** Database **************************
'2020/08/08 --------------------------------------------------*
Function fa_Accマクロ実行(qマクロ名 As String, qデータベース名 As String) As Integer
  '
  '【構文】  fa_Accマクロ実行("<マクロ名>","<データベース名>")
  '【戻り値】成功:1 失敗:0
  '
  Dim wQuery  As String
  Dim wAccdb  As String
  Dim wFilter As String
  Dim wMSG    As String
  'Open_Macro
  Dim wFile   As String
  Dim wFreeNO As Integer
  Dim wMacro  As String

  fa_Accマクロ実行 = 0
  
  'マクロ名がない
  If qマクロ名 = "" Then
    wMSG = "マクロ名を指定してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  'データベース名を省略
  If qデータベース名 = "" Then
    wAccdb = YUGE_DB
  Else
    wAccdb = qデータベース名
  End If
  
  'Open_Macro Save
  wFile = ThisWorkbook.Path & "\Open_Macro.txt"
  wFreeNO = FreeFile
  Open wFile For Output As wFreeNO
  Print #wFreeNO, qマクロ名
  Close wFreeNO
  
  ygEnd = fn_ファイルオープン(wAccdb, True)
  
  fa_Accマクロ実行 = 1
End Function

'2020/12/25 --------------------------------------------------*
Function fa_テーブル解除(qシート名 As String, qテーブル名 As String) As Integer
  '
  '【構文】  fa_テーブル解除("<シート名>","<テーブル名>",<セル範囲>)
  '【戻り値】成功:1 失敗:0 見つからない:-1
  '
  Dim WSH    As Worksheet
  Dim LST    As ListObject
  Dim wBook  As String
  Dim wMSG   As String
  Dim wFL    As Integer
  Dim wTblName As String
  Dim wRange   As String
  
  wFL = -1
  wBook = ThisWorkbook.Name
  
  'テーブル名
  If qテーブル名 = "" Then
    wTblName = "DataTableYUGE"
  Else
    wTblName = qテーブル名
  End If
  
  For Each WSH In Workbooks(wBook).Worksheets
    If qシート名 = WSH.Name Then
      wFL = 0
      
      For Each LST In WSH.ListObjects
        If wTblName = LST.Name Then
          '(解除)
          WSH.ListObjects(wTblName).Unlist
          wFL = 1
          Exit For
        End If
        
        If wFL = 0 Then Exit For
      Next LST
    End If
  Next WSH
  
  'シート名なし
  If wFL = -1 Then
    wMSG = "シート（" & qシート名 & "）が、見つかりません。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
  End If
  
  'テーブル名なし
  If wFL = 0 Then
    wMSG = "テーブル名（" & wTblName & "）が、見つかりません。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
  End If

  fa_テーブル解除 = wFL
  If wFL <> 1 Then Exit Function
  
  wRange = Sheets(qシート名).UsedRange.Address
  
  With Workbooks(wBook).Worksheets(qシート名).Range(wRange)
    '太字解除
    .Font.Bold = False
    
    '塗りつぶしなし
    .Interior.Pattern = xlNone
    .Interior.TintAndShade = 0
    .Interior.PatternTintAndShade = 0
    
    '罫線なし
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    .Borders(xlEdgeLeft).LineStyle = xlNone
    .Borders(xlEdgeTop).LineStyle = xlNone
    .Borders(xlEdgeBottom).LineStyle = xlNone
    .Borders(xlEdgeRight).LineStyle = xlNone
    .Borders(xlInsideVertical).LineStyle = xlNone
    .Borders(xlInsideHorizontal).LineStyle = xlNone
  End With
End Function

'2020/12/25 --------------------------------------------------*
Function fa_テーブル作成(qシート名 As String, qテーブル名 As String, qセル範囲 As String) As Integer
  '
  '【構文】  fa_テーブル作成("<シート名>","<テーブル名>",<セル範囲>)
  '【戻り値】成功:1 失敗:0 見つからない:-1
  '
  Dim wRange   As String
  Dim wTblName As String
  Dim wBook    As String
  Dim NM       As Name
  
  fa_テーブル作成 = 0
  
  Application.CutCopyMode = False
  'テーブル名
  If qテーブル名 = "" Then
    wTblName = "DataTableYUGE"
  Else
    wTblName = qテーブル名
  End If
  
  If fyIsTableName("", wTblName) = True Then
    ygInt1 = fa_テーブル解除(qシート名, wTblName)
  End If
  
  'セル範囲
  If qセル範囲 = "" Then
    wRange = Sheets(qシート名).UsedRange.Address
  Else
    wRange = qセル範囲
  End If
  
  Sheets(qシート名).ListObjects.Add(xlSrcRange, Range(wRange), , xlYes).Name = "DataTableYUGE"
  
  fa_テーブル作成 = 1
End Function

'2022/09/14 --------------------------------------------------*
Public Function fa_転置Array(qArray As Variant) As Variant
  Dim w縦 As Long
  Dim w横 As Long
  
  ReDim ygArrayDB(1 To UBound(qArray, 2) + 1, 1 To UBound(qArray, 1) + 1) As Variant

  For w縦 = 0 To UBound(qArray, 2)
    For w横 = 0 To UBound(qArray, 1)
      ygArrayDB(w縦 + 1, w横 + 1) = qArray(w横, w縦)
    Next w横
  Next w縦
  
  fa_転置Array = ygArrayDB
End Function

'*************************** EOE100 ***************************
'2020/07/23 --------------------------------------------------*
Function fy対応文字列(qSTR As String, qNUM As Integer) As String
  ' qSTR : ,区切り文字列 "abc,def,ghi"
  ' qNUM : 対応番号
  ' 見つからないときは ""を返す
  Dim wkST  As Integer
  Dim wkED  As Integer
  Dim wkNO  As Integer
  Dim wkMAX As Integer
  
  If fyNz(InStr(qSTR, ","), 0) > 1 And qNUM > 0 Then
    wkED = 0
    wkMAX = 0
    For wkNO = 1 To qNUM
      wkST = wkED + 1
      
      If Mid(qSTR, wkST, 1) = "," Then
        wkED = wkST
      Else
        wkED = fyNz(InStr(wkST + 1, qSTR, ","), 0)
      End If
      
      If wkED = 0 And wkMAX = 0 Then
        wkED = Len(qSTR) + 1
        wkMAX = wkNO
      End If
    
      If wkMAX > 0 And wkNO > wkMAX Then
        wkST = 0
        wkED = 0
        Exit For
      End If
    Next wkNO
    
    If wkST < wkED And wkED - wkST >= 1 Then
      fy対応文字列 = Trim(Mid(qSTR, wkST, wkED - wkST))
    Else
      fy対応文字列 = ""
    End If
  Else
    If qNUM = 1 Then
      fy対応文字列 = Trim(qSTR)
    Else
      fy対応文字列 = ""
    End If
  End If
End Function

'2020/07/23 --------------------------------------------------*
Function fyNz(qVari As Variant, qZero As Variant) As Variant
  ' 構文  : Nz(qVari, "") / Nz(qVari, 0)
  If IsNull(qVari) = True Then
    fyNz = qZero
  Else
    fyNz = qVari
  End If
End Function

'2007/10/14 --------------------------------------------------*
Function fyPickFolder(qFileName As String) As String
  ' フルパスのファイル名からフォルダ名を取り出す関数です。
  ' 例：PickFolder("C:\山田健一\Access\例題.MDB")は
  '     "C:\山田健一\Access" を返します。
  '     PickFolder("C:\例題.MDB")は "C:\" を返します。
    
    Dim wLen As Integer, wI As Integer, wJ As Integer
    wLen = Len(qFileName)
    For wI = wLen To 1 Step -1
                                     
        wJ = InStr(wI, qFileName, "\")
        If wJ <> 0 Then
            Exit For
        End If
    Next wI
    If wJ = 0 Then
        fyPickFolder = ""
    Else
        fyPickFolder = Mid$(qFileName, 1, wJ - 1)
    End If
End Function

Function fyPickFile(qFileName As String) As String
  ' フルパスのファイル名からファイル名を取り出す関数です。
  ' 例：PickFile("C:\山田健一\Access\例題.MDB")は "例題.MDB" を返します。
    
    Dim wLen As Integer, wI As Integer, wJ As Integer
    wLen = Len(qFileName)
    For wI = wLen To 1 Step -1
                                      
        wJ = InStr(wI, qFileName, "\")
        If wJ <> 0 Then
            Exit For
        End If
    Next wI
    If wJ = 0 Then
        fyPickFile = ""
    Else
        fyPickFile = Mid$(qFileName, wJ + 1, wLen - wJ + 1)
    End If
End Function

Public Function fyPickExtension(qFileName As String) As String
  ' ファイル名から拡張子を取り出す関数です。
  ' 例：PickFile("例題.MDB")は "MDB" を返します。
    
    If InStr(qFileName, ".") > 0 Then
      
      If InStr(qFileName, ".") < Len(qFileName) Then
        fyPickExtension = Mid(qFileName, InStr(qFileName, ".") + 1, Len(qFileName) - InStr(qFileName, "."))
      End If
    End If
End Function

Public Function fy囲み文字列(q文字列 As String, q開始文字 As String, q終了文字 As String, q番目 As Integer) As String
  '
  'fy囲み文字列( "<tr><td>タブレット</td><td>3</td><td>24800</td><td>74400</td></tr>" , "<td>", "</td>", 3 )
  '→ 24800
  '
  Dim wNo         As Integer
  Dim wSTPos      As Integer
  Dim wPos        As Integer
  '検索文字の長さ
  Dim w開始Len    As Integer
  Dim w終了Len    As Integer
  '切取位置
  Dim w開始Pos    As Integer
  Dim w終了Pos    As Integer
  '切取文字数
  Dim w文字Cnt    As Integer
  
  w開始Len = Len(q開始文字)
  w終了Len = Len(q終了文字)
  
  '検索開始位置を指定
  wSTPos = 1
  If q番目 > 0 Then wSTPos = q番目
  
  Select Case True
    Case wSTPos = 1
      w開始Pos = InStr(q文字列, q開始文字) + w開始Len
      w終了Pos = InStr(q文字列, q終了文字) - 1
      w文字Cnt = w終了Pos - w開始Pos + 1
      
    Case wSTPos < 1
      fy囲み文字列 = ""
      Exit Function
      
    Case Else
      '開始位置
      wPos = 1
      For wNo = 1 To q番目
        wPos = InStr(wPos, q文字列, q開始文字) + w開始Len
      Next wNo
      w開始Pos = wPos
      
      '終了位置
      wPos = 1
      For wNo = 1 To q番目
        wPos = InStr(wPos, q文字列, q終了文字) + w終了Len
      Next wNo
      w終了Pos = wPos - w終了Len - 1
      
      '文字数
      w文字Cnt = w終了Pos - w開始Pos + 1
     
  End Select
  
  '文字取得
  fy囲み文字列 = Mid(q文字列, w開始Pos, w文字Cnt)
End Function

'2010/06/16 -----------------------------------------------------*
' 機能      : RangeからColumn値を取得
' 引き数    : qRange "A1 B1"
' 備考      : Errorは 0 を返す
Public Function fyConvCol(qRange As String) As Integer
  Dim wST As Integer, wED As Integer
  Dim wNumST As Integer
  Dim wNum As String
  Dim wABC As String
  Dim wSTR As String
  Dim wCnt As Integer
  Dim wCol As Integer
  
  fyConvCol = 0
  
  If Len(qRange) = 0 Then
    Exit Function
  End If
  
  '--- Start ---
  wABC = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  wST = Len(qRange) + 1
  
  If InStr(wABC, Left(qRange, 1)) = 0 Then
    Exit Function
  End If

  For wCnt = 1 To 26
    wSTR = Mid(wABC, wCnt, 1)
    
    If InStr(qRange, wSTR) > 0 Then
      If InStr(qRange, wSTR) < wST Then wST = InStr(qRange, wSTR)
    End If
  Next wCnt

  '--- End ---
  wNum = "0123456789"
  wNumST = Len(qRange) + 1
  wED = 0

  For wCnt = 1 To 10
    wSTR = Mid(wNum, wCnt, 1)
    
    If InStr(qRange, wSTR) > 0 Then
      If InStr(qRange, wSTR) < wNumST Then wNumST = InStr(qRange, wSTR)
    End If
  Next wCnt
  
  If wNumST > 0 Then wED = wNumST - 1
  
  If (wST < Len(qRange) + 1) And (wED > 0) Then
    wCol = 0
    For wCnt = 1 To wED - wST + 1
      wCol = InStr(wABC, Mid(qRange, wCnt, 1)) * 26 ^ ((wED - wST + 1) - wCnt)
      fyConvCol = fyConvCol + wCol
    Next wCnt
  End If
End Function

'2016/05/23 -----------------------------------------------------*
' 機能      : Col からRange値を取得
' 引き数    : qCol
' 備考      : Errorは "" を返す
Public Function fyConvRange(qCol As Integer) As String
  Dim wABC As String
  Dim wCol As Integer
  Dim w商 As Integer, w余 As Integer
  
  fyConvRange = ""
  If qCol = 0 Then
    Exit Function
  End If
  
  wABC = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  w商 = Int((qCol - 1) / 26)
  w余 = ((qCol - 1) Mod 26) + 1
  
  fyConvRange = fyConvRange & Mid(wABC, w余, 1)
  
  If w商 > 0 Then
    fyConvRange = Mid(wABC, w商, 1) & fyConvRange
  End If
End Function

'2024/04/23 --------------------------------------------------*
' 機能      : URLを指定してブラウザーを起動
' 引き数    : qURL
' 備考      : Errorは 0 を返す
Public Function fyBrowserOpen(qURL As String) As Long
  On Error GoTo subError
  Dim WSH
  
  fyBrowserOpen = 0
  If qURL = "" Then Exit Function
  
  Set WSH = CreateObject("WScript.Shell")
  WSH.Run qURL, 3
  Set WSH = Nothing
  
  fyBrowserOpen = 1
  
subExit:
  Exit Function

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
End Function

'2024/04/23 --------------------------------------------------*
' 機能      : ファイルを指定して起動する
' 引き数    : qFile
' 備考      : Errorは 0 を返す
Public Function fyFileRun(qFile As String) As Long
  On Error GoTo subError
  Dim WSH
  
  fyFileRun = 0
  If qFile = "" Then Exit Function
  
  'Fileの存在を判定
  ygLng1 = PathFileExists(qFile)
  If ygLng1 = 0 Then
    MsgBox "ファイル(" & qFile & ")が、見つかりません。", vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  Set WSH = CreateObject("WScript.Shell")
  WSH.Run qFile, 3, False
  Set WSH = Nothing
  
  fyFileRun = 1
  
subExit:
  Exit Function

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
End Function

'*************************** EOE200 ***************************
'2020/05/13 --------------------------------------------------*
Public Function fyIsSheet(qSheet As String, qBook As String, qMode As Integer) As Boolean
  'Sheetの存在を判定
  'qMode : 0/警告しない 1/警告する
  '
  Dim WSH    As Worksheet
  Dim wMSG   As String
  Dim wBook  As String
  
  fyIsSheet = False
  
  'Book名を省略
  If qBook = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qBook
  End If
  
  For Each WSH In Workbooks(wBook).Worksheets
    If qSheet = WSH.Name Then
      '(存在)
      fyIsSheet = True
      Exit For
    End If
  Next
  
  If fyIsSheet = False Then
    If qMode = 1 Then
      wMSG = "シート（" & qSheet & "）が、見つかりません。"
      MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    End If
  End If
End Function

'2023/08/03 --------------------------------------------------*
Public Function fyIsVisible(qSheet As String, qBook As String) As Boolean
  'Sheetの表示/非表示を判定
  Dim WSH    As Worksheet
  Dim wMSG   As String
  Dim wBook  As String
  
  fyIsVisible = False
  
  'Book名を省略
  If qBook = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qBook
  End If
  
  For Each WSH In Workbooks(wBook).Worksheets
    If qSheet = WSH.Name And WSH.Visible = xlSheetVisible Then
      '(表示)
      fyIsVisible = True
      Exit For
    End If
  Next
End Function

'2022/11/08 --------------------------------------------------*
Public Function fyIsBook(qBook As String) As Boolean
  'BookのOpenを判定
  Dim wWBK  As Workbook
  
  fyIsBook = False
  
  For Each wWBK In Workbooks
    If wWBK.Name = qBook Then fyIsBook = True
  Next
End Function

'2022/11/12 --------------------------------------------------*
Public Function fyIsForm(qFormName As String) As Boolean
  Dim wForm As Object
  
  fyIsForm = False
  
  For Each wForm In UserForms
    If wForm.Name = qFormName Then
      fyIsForm = True
    End If
  Next
End Function

'2020/12/25 --------------------------------------------------*
Public Function fyIsTableName(qBook As String, qTableName As String) As Boolean
  'テーブル定義の存在を判定
  Dim WSH    As Worksheet
  Dim LST    As ListObject
  Dim wBook  As String
  
  fyIsTableName = False
  
  'Book名を省略
  If qBook = "" Then
    wBook = ThisWorkbook.Name
  Else
    wBook = qBook
  End If
  
  For Each WSH In Workbooks(wBook).Worksheets
    For Each LST In WSH.ListObjects
      If qTableName = LST.Name Then
        '(存在)
        fyIsTableName = True
        Exit For
      End If
      If fyIsTableName = True Then Exit For
    Next LST
  Next WSH
End Function

'*********************** Database Macro ***********************
'2022/10/28 --------------------------------------------------*
Public Sub syADOAccdbOpen(qMDB2 As String)
  On Error GoTo subError
  ygEnd = 0
  
  If qMDB2 = "" Then
    MsgBox "データベースが、設定されていません。", vbOKOnly + vbExclamation, "確認 : ADOAccDbOpen"
    Exit Sub
  End If
  
  'DB接続用SQL
  Set ADB = New ADODB.Connection
  ADB.Provider = "Microsoft.Ace.OLEDB.12.0; "
  ADB.ConnectionString = "Data Source=" & qMDB2 & ";"
  ADB.Open
  ygEnd = 1
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : syADOAccDbOpen"
  Resume subExit
End Sub

'2020/06/06 --------------------------------------------------*
Public Sub syADOAccDbClose()
  ADB.Close
  Set ADB = Nothing
End Sub

'2020/06/06 --------------------------------------------------*
Public Sub syADORecordsetOpen(qSQL As String, qNo As Integer)
  On Error GoTo subError
  
  ygEnd = 0
  If Not (qNo >= 1 And qNo <= 5) Then Exit Sub
  
  Select Case qNo
    Case 1
      Set ARST1 = New ADODB.Recordset
      ARST1.Open qSQL, ADB, adOpenKeyset, adLockOptimistic
    Case 2
      Set ARST2 = New ADODB.Recordset
      ARST2.Open qSQL, ADB, adOpenKeyset, adLockOptimistic
    Case 3
      Set ARST3 = New ADODB.Recordset
      ARST3.Open qSQL, ADB, adOpenKeyset, adLockOptimistic
    Case 4
      Set ARST4 = New ADODB.Recordset
      ARST4.Open qSQL, ADB, adOpenKeyset, adLockOptimistic
    Case 5
      Set ARST5 = New ADODB.Recordset
      ARST5.Open qSQL, ADB, adOpenKeyset, adLockOptimistic
  End Select
  
  ygEnd = 1
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : syADORecordsetOpen"
  Resume subExit
End Sub

'2020/06/06 --------------------------------------------------*
Public Sub syADORecordsetClose(qNo As Integer)
  If Not (qNo >= 1 And qNo <= 5) Then Exit Sub

  Select Case qNo
    Case 1
      ARST1.Close
      Set ARST1 = Nothing
    Case 2
      ARST2.Close
      Set ARST2 = Nothing
    Case 3
      ARST3.Close
      Set ARST3 = Nothing
    Case 4
      ARST4.Close
      Set ARST4 = Nothing
    Case 5
      ARST5.Close
      Set ARST5 = Nothing
  End Select
End Sub

