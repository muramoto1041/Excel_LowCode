Attribute VB_Name = "�f�[�^�x�[�X�}�N��"
Option Explicit
'
'�f�[�^�x�[�X�}�N�����g���ɂ́A���̃t�@�C���Ɠ����t�H���_�ɉ��L�̃f�[�^�x�[�X(YUGE_DB)���Z�b�g���Ă��������B
'
'YUGE�f�[�^�x�[�X��
Public Const YUGE_DB = "YUGE_Database.accdb"
'

'2020/12/25 --------------------------------------------------*
Sub DB���j���[()
  On Error GoTo subError
  Dim wMenu  As String
  
LblStart:
  wMenu = "���ڂ̒ǉ�,�ҏW(�t�H�[�����),�ҏW(�ꗗ�\���)"
  
  NN = fn_���j���[��("�^�C�g��", wMenu, "�I�����Ă�������", 0)
  If NN = 0 Then Exit Sub
  '--- Exit ---
  
  '(���j���[���s)
  If NN = 1 Then Call ���ڂ����
  If NN = 2 Then Call �t�H�[���ҏW
  If NN = 3 Then Call �e�[�u���ҏW
  'GoTo LblStart
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

'2020/12/25 --------------------------------------------------*
Sub �e�[�u���ҏW()
  '�e�[�u���ҏW
  Call ���ڈʒu��ۑ�
  NN = fa_Acc�}�N�����s("OpenTable", "")
End Sub

'2020/12/25 --------------------------------------------------*
Sub �t�H�[���ҏW()
  '�t�H�[���\��
  Call ���ڈʒu��ۑ�
  NN = fa_Acc�}�N�����s("OpenForm", "")
End Sub

'2020/12/25 --------------------------------------------------*
Sub ���ڂ����()
  '���ڂ����
  On Error GoTo subError
  Dim wNo   As Integer
  Dim wLeft As Single, wTop As Single, wWidth As Single, wHeight As Single
  Dim wSTR  As String
  Dim wTab  As Single
  Dim wCnt1 As Integer
  Dim wCnt2 As Integer
   
LblStart:
  SS = fs_��������("1,100", "����", "�ԍ�", "�ǉ����鍀�ڔԍ�����͂��Ă�������")
  If SS = "" Then Exit Sub
  '--- Exit ---
   
  wCnt1 = Val(yg����ST)
  wCnt2 = Val(yg����ED)
  
  If (wCnt1 < 1 Or wCnt1 > 100) And (wCnt2 < 1 Or wCnt2 > 100) Then
    MsgBox "1�`100 �̒l����͂��Ă��������B", vbOKOnly + vbExclamation, "�m�F"
    GoTo LblStart
  End If
   
  If wCnt1 = 0 Then wCnt1 = 1
  If wCnt2 = 0 Then wCnt2 = 100
   
  Application.ScreenUpdating = False
   
  Call s_DeleteShape(wCnt1, wCnt2)
  
  For wNo = wCnt1 To wCnt2
    wTab = Int((wNo - 1) / 20) * 350
    
    '�ێl�p
    wTop = (((wNo - 1) Mod 20)) * 30 + 30
    wLeft = wTab + 10
    wWidth = 100
    wHeight = 25
    wSTR = "���ږ�" & Trim(Str(wNo))
    Call s_CreateMaruLabel(wLeft, wTop, wWidth, wHeight, wNo)
    
    '�l�p
    wLeft = wTab + 120
    wWidth = 200
    wSTR = "�S�p"
    Call s_CreateKakuBox(wLeft, wTop, wWidth, wHeight, wNo)
  Next wNo
  
  Application.ScreenUpdating = True
  Range("A1").Select
  
  NN = fn_�m�F("�쐬���܂����B", "i", "�m�F")
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_CreateMaruLabel(qLeft As Single, qTop As Single, qWidth As Single, qHeight As Single, qNo As Integer)
  '�ێl�p
  Dim wName As String
  Dim wRGB  As Long
  Dim wSTR  As String
  
  ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, qLeft, qTop, qWidth, qHeight).Select

  wName = "MaruLbl" & Format(qNo, "000"): Call s_ShapeName(wName)        '�V�F�C�v���O
  wRGB = RGB(234, 234, 234):                Call s_BackColor(wRGB)          '�w�i�F(RGB)
  wSTR = "���ږ�" & Trim(Str(qNo)):         Call s_ShapeCaption(wSTR, "R")  '�e�L�X�g�ҏW
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_CreateKakuBox(qLeft As Single, qTop As Single, qWidth As Single, qHeight As Single, qNo As Integer)
  '�l�p
  Dim wName As String
  Dim wRGB  As Long
  Dim wSTR  As String
  
  ActiveSheet.Shapes.AddShape(msoShapeRectangle, qLeft, qTop, qWidth, qHeight).Select

  wName = "KakuTxt" & Format(qNo, "000"): Call s_ShapeName(wName)           '�V�F�C�v���O
  wRGB = RGB(255, 255, 255):              Call s_BackColor(wRGB)            '�w�i�F(RGB)
  wSTR = "���p/�S�p/�w��Ȃ�":            Call s_ShapeCaption(wSTR, "L")    '�e�L�X�g�ҏW
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_BackColor(qRGB As Long)
  '�w�i�F
  With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .ForeColor.RGB = qRGB
    .Transparency = 0
    .Solid
  End With
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_ShapeCaption(qCaption As String, qAlign As String)
  '�e�L�X�g�ҏW
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
  '�V�F�C�v��
  Selection.ShapeRange.Name = qName
  Selection.Name = qName
End Sub

'2020/12/25 --------------------------------------------------*
Private Sub s_DeleteShape(qNoST As Integer, qNoED As Integer)
  '���ڂ�����
  Dim wObj   As Object
  Dim wName  As String
  Dim wNo    As Integer
  
  For Each wObj In ActiveSheet.Shapes
    wName = wObj.Name
    If (Left(wName, 4) = "Maru" Or Left(wName, 4) = "Kaku") And Len(wName) = 10 Then
      wNo = Val(Mid(wName, 8, 3))
      
      '�͈͓��̃I�u�W�F�N�g������
      If wNo >= qNoST And wNo <= qNoED Then
        ActiveSheet.Shapes(wName).Delete
      End If
    End If
  Next
End Sub

'2020/12/25 --------------------------------------------------*
Sub ���ڈʒu��ۑ�()
  '���ڈʒu��ۑ�
  On Error GoTo subError
  Dim wFile  As String
  
  Set ADB = New ADODB.Connection

  'DB�ڑ��pSQL
  wFile = ThisWorkbook.Path & "\" & YUGE_DB
  Call syADOAccdbOpen(wFile)
  
  '�w�i�F�ɂ��W�v�l
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
  , "�m�F"
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

