Attribute VB_Name = "データベースマクロ"
Option Explicit
'
'データベースマクロを使うには、このファイルと同じフォルダに下記のデータベース(YUGE_DB)をセットしてください。
'
'YUGEデータベース名
Public Const YUGE_DB = "YUGE_Database.accdb"
'

'2020/12/25 --------------------------------------------------*
Sub DBメニュー()
  On Error GoTo subError
  Dim wMenu  As String
  
LblStart:
  wMenu = "項目の追加,編集(フォーム画面),編集(一覧表画面)"
  
  NN = fn_メニュー小("タイトル", wMenu, "選択してください", 0)
  If NN = 0 Then Exit Sub
  '--- Exit ---
  
  '(メニュー実行)
  If NN = 1 Then Call 項目を作る
  If NN = 2 Then Call フォーム編集
  If NN = 3 Then Call テーブル編集
  'GoTo LblStart
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

'2020/12/25 --------------------------------------------------*
Sub テーブル編集()
  'テーブル編集
  Call 項目位置を保存
  NN = fa_Accマクロ実行("OpenTable", "")
End Sub

'2020/12/25 --------------------------------------------------*
Sub フォーム編集()
  'フォーム表示
  Call 項目位置を保存
  NN = fa_Accマクロ実行("OpenForm", "")
End Sub

'2020/12/25 --------------------------------------------------*
Sub 項目を作る()
  '項目を作る
  On Error GoTo subError
  Dim wNo   As Integer
  Dim wLeft As Single, wTop As Single, wWidth As Single, wHeight As Single
  Dim wSTR  As String
  Dim wTab  As Single
  Dim wCnt1 As Integer
  Dim wCnt2 As Integer
   
LblStart:
  SS = fs_条件入力("1,100", "する", "番号", "追加する項目番号を入力してください")
  If SS = "" Then Exit Sub
  '--- Exit ---
   
  wCnt1 = Val(yg条件ST)
  wCnt2 = Val(yg条件ED)
  
  If (wCnt1 < 1 Or wCnt1 > 100) And (wCnt2 < 1 Or wCnt2 > 100) Then
    MsgBox "1～100 の値を入力してください。", vbOKOnly + vbExclamation, "確認"
    GoTo LblStart
  End If
   
  If wCnt1 = 0 Then wCnt1 = 1
  If wCnt2 = 0 Then wCnt2 = 100
   
  Application.ScreenUpdating = False
   
  Call s_DeleteShape(wCnt1, wCnt2)
  
  For wNo = wCnt1 To wCnt2
    wTab = Int((wNo - 1) / 20) * 350
    
    '丸四角
    wTop = (((wNo - 1) Mod 20)) * 30 + 30
    wLeft = wTab + 10
    wWidth = 100
    wHeight = 25
    wSTR = "項目名" & Trim(Str(wNo))
    Call s_CreateMaruLabel(wLeft, wTop, wWidth, wHeight, wNo)
    
    '四角
    wLeft = wTab + 120
    wWidth = 200
    wSTR = "全角"
    Call s_CreateKakuBox(wLeft, wTop, wWidth, wHeight, wNo)
  Next wNo
  
  Application.ScreenUpdating = True
  Range("A1").Select
  
  NN = fn_確認("作成しました。", "i", "確認")
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_CreateMaruLabel(qLeft As Single, qTop As Single, qWidth As Single, qHeight As Single, qNo As Integer)
  '丸四角
  Dim wName As String
  Dim wRGB  As Long
  Dim wSTR  As String
  
  ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, qLeft, qTop, qWidth, qHeight).Select

  wName = "MaruLbl" & Format(qNo, "000"): Call s_ShapeName(wName)        'シェイプ名前
  wRGB = RGB(234, 234, 234):                Call s_BackColor(wRGB)          '背景色(RGB)
  wSTR = "項目名" & Trim(Str(qNo)):         Call s_ShapeCaption(wSTR, "R")  'テキスト編集
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_CreateKakuBox(qLeft As Single, qTop As Single, qWidth As Single, qHeight As Single, qNo As Integer)
  '四角
  Dim wName As String
  Dim wRGB  As Long
  Dim wSTR  As String
  
  ActiveSheet.Shapes.AddShape(msoShapeRectangle, qLeft, qTop, qWidth, qHeight).Select

  wName = "KakuTxt" & Format(qNo, "000"): Call s_ShapeName(wName)           'シェイプ名前
  wRGB = RGB(255, 255, 255):              Call s_BackColor(wRGB)            '背景色(RGB)
  wSTR = "半角/全角/指定なし":            Call s_ShapeCaption(wSTR, "L")    'テキスト編集
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_BackColor(qRGB As Long)
  '背景色
  With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .ForeColor.RGB = qRGB
    .Transparency = 0
    .Solid
  End With
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_ShapeCaption(qCaption As String, qAlign As String)
  'テキスト編集
  Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = qCaption
  
  'Alignment
  With Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat
    Select Case qAlign
      Case "R": .Alignment = msoAlignRight
      Case "L": .Alignment = msoAlignLeft
    End Select
  End With
  
  'Font
  With Selection.ShapeRange.TextFrame2.TextRange.Characters.Font
    .NameComplexScript = "+mn-cs"
    .NameFarEast = "+mn-ea"
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(0, 0, 0)
    .Fill.Transparency = 0
    .Fill.Solid
    .Size = 11
    .Name = "+mn-lt"
  End With
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_ShapeName(qName As String)
  'シェイプ名
  Selection.ShapeRange.Name = qName
  Selection.Name = qName
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_DeleteShape(qNoST As Integer, qNoED As Integer)
  '項目を消す
  Dim wObj   As Object
  Dim wName  As String
  Dim wNo    As Integer
  
  For Each wObj In ActiveSheet.Shapes
    wName = wObj.Name
    If (Left(wName, 4) = "Maru" Or Left(wName, 4) = "Kaku") And Len(wName) = 10 Then
      wNo = Val(Mid(wName, 8, 3))
      
      '範囲内のオブジェクトを消去
      If wNo >= qNoST And wNo <= qNoED Then
        ActiveSheet.Shapes(wName).Delete
      End If
    End If
  Next
End Sub

'2020/12/25 --------------------------------------------------*
Sub 項目位置を保存()
  '項目位置を保存
  On Error GoTo subError
  Dim wFile  As String
  
  Set ADB = New ADODB.Connection

  'DB接続用SQL
  wFile = ThisWorkbook.Path & "\" & YUGE_DB
  Call syADOAccdbOpen(wFile)
  
  '背景色による集計値
  ADB.Execute "DELETE FROM T_SwShapeInfo ;"

  SQL1 = "SELECT * FROM T_SwShapeInfo ;"
  Set ARST1 = New ADODB.Recordset
  ARST1.Open SQL1, ADB, adOpenKeyset, adLockOptimistic
  
  Call s_GetShapeInfo

  ARST1.Close
  ADB.Close
  Set ARST1 = Nothing
  Set ADB = Nothing

subExit:
  Exit Sub

subError:
  ARST1.Close
  ADB.Close
  Set ARST1 = Nothing
  Set ADB = Nothing
  
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認"
  Resume subExit
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_GetShapeInfo()
  Dim wObj    As Object
  Dim wLeft   As Single
  Dim wTop    As Single
  Dim wWidth  As Single
  Dim wHeight As Single
  Dim wSTR    As String
  Dim wIME    As String
  Dim wName   As String
  Dim wNo     As Integer
  
  For Each wObj In ActiveSheet.Shapes
    wName = wObj.Name
    wLeft = wObj.Left
    wTop = wObj.Top
    wWidth = wObj.Width
    wHeight = wObj.Height
    
    If Left(wName, 4) = "Maru" Or Left(wName, 4) = "Kaku" Then
      wNo = wNo + 1
      If Left(wName, 4) = "Maru" Then wSTR = wObj.TextFrame.Characters.Text
      If Left(wName, 4) = "Kaku" Then wSTR = wObj.TextFrame.Characters.Text
      
      ARST1.AddNew
      ARST1![IDno] = wNo
      ARST1![Shape] = wName
      ARST1![Caption] = wSTR
      ARST1![Left] = wLeft
      ARST1![Top] = wTop
      ARST1![Width] = wWidth
      ARST1![Height] = wHeight
      ARST1.Update
    End If
  Next
    
End Sub

