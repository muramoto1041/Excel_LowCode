VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_��������_BS 
   Caption         =   "��������"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   OleObjectBlob   =   "Fy_��������_BS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_��������_Bs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'�쐬���F2020/06/07
'                             Fy_��������
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
  
  '������
  If yg������ <> "" Then lbl������ = yg������
  
  '���������l
  If yg���� = "" Then
    Me.txt����.Value = Null
  Else
    Me.txt����.Text = yg����
  End If
  
  '���b�Z�[�W
  If ygSTR1 = "" Then
    Me.lblMsg.Caption = ""
  Else
    Me.lblMsg.Caption = ygSTR1
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : FormIni"
  Resume Next
End Sub

'2020/07/25 --------------------------------------------------*
Private Sub cmdOK_Click()
  
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
  Dim w����    As String
  
  fsInputCheck = False
  
  w���� = Me.txt����.Text
  
  '--- Check ---
  If w���� = "" Then
    wMSG = wMSG & "��������͂��Ă��������B" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Function
  End If
  
  yg���� = w����
  
  fsInputCheck = True
End Function

