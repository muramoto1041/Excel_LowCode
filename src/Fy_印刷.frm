VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_��� 
   Caption         =   "���"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   OleObjectBlob   =   "Fy_���.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'�쐬���F2020/06/07
'                             Fy_���
'�쐬�ҁF���{�r�a
'****************************************************************
Option Explicit
'System
Dim sMsg   As Integer
Dim sWhere As String
'Procedure
Dim sFL�g��      As Integer    '0:�W�� 1:�g��
Dim s����NO      As Integer
Dim s���Sheet   As String
Dim s���ږ�(20)  As String
Dim s���ז�(20)  As String
Dim s���׍s��    As Integer
Dim sQueryName01 As String
Dim sQueryName02 As String
Dim sCardNO      As Integer
Dim sIDCard      As Long
Dim s�ꊇ����    As Long
Dim s�������    As Integer

'2020/05/13 --------------------------------------------------*
Private Sub UserForm_Initialize()
  On Error GoTo subError
  
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : FormIni"
  Resume Next
End Sub

'2020/06/08 --------------------------------------------------*
Private Sub cmdPreview_Click()
  On Error GoTo subError

  '--- �v���r���[ ---
  ygCntPrt = 1
  Me.Hide
  NN = fn_�V�[�g���(yg���Sheet, ygStartBook, "����", 1)
  Me.Show

subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : Preview"
  Resume subExit
End Sub

'2020/06/08 --------------------------------------------------*
Private Sub cmd���_Click()
  On Error GoTo subError

  '--- ��� ---
  ygCntPrt = 1
  Me.Hide
  NN = fn_�V�[�g���(yg���Sheet, ygStartBook, "���Ȃ�", 1)
  Me.Show

subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : Preview"
  Resume subExit
End Sub

'2022/07/20 --------------------------------------------------*
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  On Error GoTo subError
  
  If yg���Sheet <> "" Then
    NN = fn_�V�[�g������(yg���Sheet, "")
  End If

subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : QueryClose"
  Resume Next
End Sub
