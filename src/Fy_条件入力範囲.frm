VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_�������͔͈� 
   Caption         =   "�������́i�͈́j"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   OleObjectBlob   =   "Fy_�������͔͈�.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_�������͔͈�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************
'�쐬���F2020/06/07
'                             Fy_�������́i�͈́j
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
  Dim w����ST  As String
  Dim w����ED  As String
  
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15

  ygEnd = 0
  
  '������
  If yg������ <> "" Then lbl������ = yg������
  
  If InStr(yg����, ",") > 0 And Len(yg����) > 1 Then
    yg����ST = Left(yg����, InStr(yg����, ",") - 1)
    yg����ED = Mid(yg����, InStr(yg����, ",") + 1, Len(yg����) - InStr(yg����, ","))
  End If
  
  '���������l
  Me.txt����ST.Text = yg����ST
  Me.txt����ED.Text = yg����ED
  
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
  Dim w����ST  As String
  Dim w����ED  As String
  
  fsInputCheck = False
  
  w����ST = Me.txt����ST.Text
  w����ED = Me.txt����ED.Text
  
  '--- Check ---
  If w����ST = "" And w����ED = "" Then
    wMSG = wMSG & "�J�n�����܂��͏I����������͂��Ă��������B" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Function
  End If
  
  yg����ST = w����ST
  yg����ED = w����ED
  
  fsInputCheck = True
End Function



