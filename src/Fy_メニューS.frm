VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_���j���[S 
   Caption         =   "���j���["
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3555
   OleObjectBlob   =   "Fy_���j���[S.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_���j���[S"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'****************************************************************
'�쐬���F2020/06/01
'                             Fy_���j���[
'�쐬�ҁF���{�r�a
'****************************************************************
Option Explicit
'System
Dim sMsg   As Integer
Dim sWhere As String
'Procedure
Dim sTitle    As String
Dim sMenuList As String
Dim sStrMsg   As String
Dim sMenuNo   As Integer
'YUGE�R�}���h
Dim sCommand       As String
Dim s�N���b�N�\��  As String

'2020/05/10 --------------------------------------------------*
Private Sub UserForm_Initialize()
  On Error GoTo subError
      
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15
  
  ygEnd = 0
  
  sTitle = ygSTR1
  sMenuList = ygSTR2
  sStrMsg = ygSTR3
  sMenuNo = ygInt1
  
  'YUGE�R�}���h
  sCommand = ""
  
  Call s_MenuList(sMenuList)
  
  '�^�C�g��
  If sTitle <> "" Then
    'YUGE��Call
    If Left(sTitle, 1) = "$" And Right(sTitle, 1) = "$" And Len(sTitle) > 2 Then
      sTitle = Mid(sTitle, 2, Len(sTitle) - 2)
      sCommand = sTitle
      
      Select Case sCommand
        Case "�V�[�g���j���[": s�N���b�N�\�� = ygSTR4   '����A���Ȃ�
      End Select
    End If
    
    Me.Caption = sTitle
  End If
  
  '���b�Z�[�W
  Me.lblMsg.Caption = sStrMsg
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : FormIni"
  Resume Next
End Sub

'2020/06/01 --------------------------------------------------*
Private Sub s_MenuList(qMenuList As String)
  Dim wNo As Integer
  Dim wMenuStr As String
  
  If qMenuList = "" Then Exit Sub
  
  Do
    wNo = wNo + 1
    wMenuStr = fy�Ή�������(qMenuList, wNo)
    If wMenuStr = "" Then Exit Do
    
    Me.lstMenu.AddItem wMenuStr
  Loop
  
  If sMenuNo > 0 And sMenuNo < wNo Then
    Me.lstMenu.Selected(sMenuNo - 1) = True
  End If
End Sub

'2020/07/19 --------------------------------------------------*
Private Sub lstMenu_Click()
  '
  'YUGE�R�}���h�y�V�[�g���j���[�z
  '���j���[���N���b�N����ƃV�[�g��\������B
  '
  On Error GoTo subError
  Dim wMenu  As String
  Dim wNo    As Integer
  Dim wCnt   As Integer
  
  If Not (sCommand = "�V�[�g���j���[" And (s�N���b�N�\�� = "����" Or s�N���b�N�\�� = "")) Then Exit Sub
  
  wCnt = Me.lstMenu.ListCount
  
  For wNo = 0 To wCnt - 1
    If Me.lstMenu.Selected(wNo) = True Then
      wMenu = Me.lstMenu.List(wNo)
      Exit For
    End If
  Next wNo
  
  '�V�[�g�\��
  NN = fn_�V�[�g�\��(wMenu, "", True)
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : lstMenu"
  Resume subExit
End Sub

'2020/05/10 --------------------------------------------------*
Private Sub cmd�I��_Click()
  On Error GoTo subError
  Dim wMSG    As String
  Dim wMenu   As String
  Dim wNo     As Integer
  Dim wCnt    As Integer
  
  wCnt = Me.lstMenu.ListCount
  
  sMenuNo = 0
  
  For wNo = 0 To wCnt - 1
    If Me.lstMenu.Selected(wNo) = True Then
      wMenu = Me.lstMenu.List(wNo)
      sMenuNo = wNo + 1
      Exit For
    End If
  Next wNo
  
  If sMenuNo = 0 Then
    wMSG = "���j���[��I�����Ă��������B"
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Sub
  End If
  
  ygEnd = 1
  
  ygInt1 = sMenuNo
  Unload Me
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : �I��"
  Resume subExit
End Sub

'2020/07/20 --------------------------------------------------*
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  
End Sub

Public Sub Test3()
  MsgBox "Test3"
End Sub

