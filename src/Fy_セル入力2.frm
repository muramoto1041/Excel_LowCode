VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_�Z������2 
   Caption         =   "�Z������2"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "Fy_�Z������2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_�Z������2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'�쐬���F2020/06/07
'                             Fy_�Z������
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
  
  '���x����
  If ygLBL1 <> "" Then Me.lbl���͖�01 = ygLBL1
  If ygLBL2 <> "" Then Me.lbl���͖�02 = ygLBL2
  
  '���͏����l
  '(1)
  If ygSTR1 = "" Then
    Me.txt����01.Value = Null
  Else
    Me.txt����01.Text = ygSTR1
  End If
  '(2)
  If ygSTR2 = "" Then
    Me.txt����02.Value = Null
  Else
    Me.txt����02.Text = ygSTR2
  End If
  
  '���b�Z�[�W
  If ygMSG = "" Then
    Me.lblMsg.Caption = ""
  Else
    Me.lblMsg.Caption = ygMSG
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
  Dim w����01  As String, w����02 As String
  
  fsInputCheck = False
  
  w����01 = Me.txt����01.Text
  w����02 = Me.txt����02.Text
  
  '--- Check ---
  If w����01 = "" Then
    wMSG = wMSG & "�l����͂��Ă��������B" & vbCrLf
  End If
  If w����02 = "" Then
    wMSG = wMSG & "�l����͂��Ă��������B" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Function
  End If
  
  ygSTR1 = w����01
  ygSTR2 = w����02
  
  fsInputCheck = True
End Function

