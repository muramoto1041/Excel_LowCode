Attribute VB_Name = "���[�U�[�}�N��"
'
' ���[�U�[�}�N��
' ���R�ɋL�q�A�ҏW���Ă��������B
'

Sub �m�F()
  '
  '�y�\���z  fn_�m�F("<���b�Z�[�W>","<�A�C�R���^�C�v: i ? ! x >","<�^�C�g��>")
  '�y�߂�l�z�mOK�n�m�͂��n�� 1�A�m�������n�� 2�A�mx�n�� 0
  '
  On Error GoTo subError
  
  NN = fn_�m�F("����ɂ��́I", "i", "�m�F")
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub ���j���[()
  '
  '�y�\���z�@fn_���j���[("<�^�C�g��>","<���j���[1>,<���j���[2>,�E�E�E","<���[�b�Z�[�W>",���j���[�����l)
  '�y�߂�l�z�I���������j���[�ԍ��B[���~]�����Ƃ��� 0 ��Ԃ��܂��B
  '
  On Error GoTo subError
  
  NN = fn_���j���[��("�^�C�g��", "���j���[1,���j���[2", "�I�����Ă�������", 0)
  
  '(���ʕ\��)
  If NN > 0 Then
    NN = fn_�m�F(Str(n) & " �Ԃ� �I�����܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub �V�[�g���j���[()
  '
  '�y�\���z�@fs_�V�[�g���j���[("<�N���b�N�\��>","<��\���V�[�g>",���j���[�����l, �\���I�v�V����)
  '�y�߂�l�z�I�������V�[�g���B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  '<�N���b�N�\��>:�N���b�N�����Ƃ��ɃV�[�g��\������
  '<��\���V�[�g>:�\�����Ă��邯�ǂ��A���j���[�ɕ\�����Ȃ��V�[�g
  '���j���[�����l:0�`n
  '�\���I�v�V����:True:�S�ĕ\��/False:��\���V�[�g������
  '
  On Error GoTo subError
  
  SS = fs_�V�[�g���j���[("����", "�T���v���}�N��,�p�[�c", 0, False)
  
  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� �I�����܂����B", "i", "�m�F")
    NN = fn_�V�[�g�\��("�T���v���}�N��", "", True)
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub ���t����()
  '
  '�y�\���z�@fs_�������t("<���t�����l>","���Ȃ�")
  '�y�߂�l�z���͂������t�B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError

  SS = fs_�������t("", "���Ȃ�", "")

  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� ���͂��܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub ���t_���͈�()
  '
  '�y�\���z�@fs_�������t("<���t�����l>","����E���E�T�E���Ȃ�")
  '�y�߂�l�z���͂������t�B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError

  SS = fs_�������t(Date, "��", "")

  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� ���͂��܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub ���t_�T�͈�()
  '
  '�y�\���z�@fs_�������t("<���t�����l>","����E���E�T�E���Ȃ�")
  '�y�߂�l�z���͂������t�B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError

  SS = fs_�������t(Date, "�T", "")

  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� ���͂��܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub �N������()
  '
  '�y�\���z�@fs_�����N��("<�N�������l>","���Ȃ�","���b�Z�[�W")
  '�y�߂�l�z���͂����N��(yyyy/mm)�B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError

  SS = fs_�����N��("", "���Ȃ�", "�X�V�N����I�����Ă�������")

  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� ���͂��܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub �N���͈�()
  '
  '�y�\���z�@fs_�����N��("<�N�������l>","<�͈͎w�� ����E���Ȃ�>","���b�Z�[�W")
  '�y�߂�l�z���͂����N��(yyyy/mm �` yyyy/mm)�B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError

  SS = fs_�����N��("2020/10", "����", "�W�v������Ԃ��w�肵�Ă�������")

  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� ���͂��܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub ��������()
  '
  '�y�\���z�@fs_��������("<���������l>","���Ȃ�","������","<���b�Z�[�W>")
  '�y�߂�l�z���͂��������B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError

  SS = fs_��������("", "���Ȃ�", "��������", "")

  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� ���͂��܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub �����͈�()
  '
  '�y�\���z�@fs_��������("<���������l>","���Ȃ�","������","<���b�Z�[�W>")
  '�y�߂�l�z���͂��������B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError

  SS = fs_��������("", "����", "��������", "")

  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� ���͂��܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub �Z������()
  '
  '�y�\���z�@fs_�Z������("<���͏����l>", ���͐�, ���͖�List, ���b�Z�[�W)
  '�y�߂�l�z���͂����l(ygSTR1�`ygSTR5) "�l1,�l2,..."�B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError

  SS = fs_�Z������("", 3, "(1),(2),(3)", "")

  '(���ʕ\��)
  If SS <> "" Then
    NN = fn_�m�F(SS & " �� ���͂��܂����B", "i", "�m�F")
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Sub VBA�\��()
  '
  '[Alt] + [F11] �� VBA�\��
  '

  SendKeys "%{F11}"
  
End Sub

Public Sub �V�KYUGEbook�쐬()
  '���㐔:qCnt
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
  Dim w�폜SH   As String
 
  wRunPath = ThisWorkbook.Path
  wRunFile = ThisWorkbook.Name
  
  '�t�@�C����
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
    wMSG = "�V�K YUGE_Book���쐬�ł��܂���ł����B"
    NN = fn_�m�F(wMSG, "!", "�m�F")
    Exit Sub
  End If
  
  'Save1
  Workbooks(wRunFile).SaveAs Filename:=wNewFile
  
  Application.ScreenUpdating = False
  NN = fn_�V�[�g�ǉ�("", "")
  wSheet = ActiveSheet.Name
  
  'Sheet�폜
  For Each WSH In ThisWorkbook.Worksheets
    w�폜SH = WSH.Name
    If wSheet <> w�폜SH Then
      NN = fn_�V�[�g�폜(w�폜SH, "", 0)
    End If
  Next
  Application.ScreenUpdating = True
  
  'Save2
  ThisWorkbook.Save
End Sub

Private Sub �V�KYUGEbook�쐬_bak()
  'Backup�t�@�C��/��
  'wFiles = wDTFile '�G�N�Z���J�[�hDT
     
  '�ŏIBak���폜
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
  
  'Bak�t�@�C����Rename
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


Public Sub �}�j���A��()
  '
  '�y�\���z�@fs_�Z������("<���͏����l>", ���͐�, ���͖�List, ���b�Z�[�W)
  '�y�߂�l�z���͂����l(ygSTR1�`ygSTR5) "�l1,�l2,..."�B[���~]�����Ƃ��� �� ��Ԃ��܂��B
  '
  On Error GoTo subError
  
  NN = fyBrowserOpen("https://excel-databace.hatenablog.com/entry/yuge-help")
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

Public Sub �^�C�}�[��~()
  On Error GoTo subError
  Dim wMsgStr As String
  Dim wMsgInt As Integer
  
  NN = fn_�^�C�}�[��~()
  
  If n = -1 Then
     'MsgBox "couldn't kill the timer"
     wMsgStr = "�v���O�������I�����āA�ċN�����Ă��������B" & vbCrLf & vbCrLf & _
               "��낵���ł����H"
     wMsgInt = MsgBox(wMsgStr, vbYesNo + vbQuestion, "�m�F")
     
     If wMsgInt = vbYes Then
       '�I��
       'ThisWorkbook.Save
       Workbooks(ygStartBook).Save
       Application.Quit
     End If
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F"
  Resume subExit
End Sub

