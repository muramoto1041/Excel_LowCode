VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_�����N��_BS 
   Caption         =   "�N������"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "Fy_�����N��_BS.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_�����N��_BS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************
'�쐬���F2020/10/05
'                             Fy_�����N��
'�쐬�ҁF���{�r�a
'****************************************************************
Option Explicit
'System
Dim sMsg   As Integer
Dim sWhere As String
'Procedure

'2020/07/25 --------------------------------------------------*
Private Sub UserForm_Initialize()
  On Error GoTo subError
  
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15

  ygEnd = 0
  
  '���t�����l
  If yg�N�� = "" Then
    Me.cbo�N.Value = Null
    Me.cbo��.Value = Null
  Else
    Me.cbo�N.Value = Year(yg�N�� & "/01")
    Me.cbo��.Value = Month(yg�N�� & "/01")
  End If
  
  
  '���b�Z�[�W�\��
  If ygSTR1 <> "" Then
    Me.lblMsg.Caption = ygSTR1
  End If
  
  Call s_SetCombo�N��
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : FormIni"
  Resume Next
End Sub

'2020/10/05 --------------------------------------------------*
Private Sub s_SetCombo�N��()
  Dim wNo  As Integer
  
  '(�N)
  For wNo = 1 To 5
    Me.cbo�N.AddItem Year(Date) - 3 + wNo
  Next wNo
  
  '(��)
  For wNo = 1 To 12
    Me.cbo��.AddItem wNo
  Next wNo
End Sub

'2020/07/25 --------------------------------------------------*
Private Sub cmdOK_Click()
  On Error GoTo subError
  
  If fsInputCheck() = False Then Exit Sub
  
  ygEnd = 1
  Unload Me
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : cmdOK"
  Resume subExit
End Sub

'2020/07/25 --------------------------------------------------*
Private Function fsInputCheck() As Boolean
  Dim wMSG     As String
  Dim w�N      As String
  Dim w��      As String
  Dim w�N��    As String
  
  fsInputCheck = False
  
  w�N = Me.cbo�N.Text
  w�� = Me.cbo��.Text
  w�N�� = w�N & "/" & w�� & "/1"
  
  '--- Check ---
  If IsDate(w�N��) = False Then
    wMSG = wMSG & "�N���𐳂������͂��Ă��������B�i�N�F2020 etc. ���F1�`12�j" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Function
  End If
  
  yg�N�� = Format(w�N��, "yyyy/mm")
  
  fsInputCheck = True
End Function

