VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_�Z������5 
   Caption         =   "�Z������5"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4410
   OleObjectBlob   =   "Fy_�Z������5.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_�Z������5"
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
  If ygLBL3 <> "" Then Me.lbl���͖�03 = ygLBL3
  If ygLBL4 <> "" Then Me.lbl���͖�04 = ygLBL4
  If ygLBL5 <> "" Then Me.lbl���͖�05 = ygLBL5
  
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
  '(3)
  If ygSTR3 = "" Then
    Me.txt����03.Value = Null
  Else
    Me.txt����03.Text = ygSTR3
  End If
  '(4)
  If ygSTR4 = "" Then
    Me.txt����04.Value = Null
  Else
    Me.txt����04.Text = ygSTR4
  End If
  '(5)
  If ygSTR5 = "" Then
    Me.txt����05.Value = Null
  Else
    Me.txt����05.Text = ygSTR5
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
  Dim w����01  As String, w����02 As String, w����03 As String, w����04 As String, w����05 As String
  
  fsInputCheck = False
  
  w����01 = Me.txt����01.Text
  w����02 = Me.txt����02.Text
  w����03 = Me.txt����03.Text
  w����04 = Me.txt����04.Text
  w����05 = Me.txt����05.Text
  
  '--- Check ---
  If w����01 = "" Then
    wMSG = wMSG & "�l����͂��Ă��������B" & vbCrLf
  End If
  If w����02 = "" Then
    wMSG = wMSG & "�l����͂��Ă��������B" & vbCrLf
  End If
  If w����03 = "" Then
    wMSG = wMSG & "�l����͂��Ă��������B" & vbCrLf
  End If
  If w����04 = "" Then
    wMSG = wMSG & "�l����͂��Ă��������B" & vbCrLf
  End If
  If w����05 = "" Then
    wMSG = wMSG & "�l����͂��Ă��������B" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Function
  End If
  
  ygSTR1 = w����01
  ygSTR2 = w����02
  ygSTR3 = w����03
  ygSTR4 = w����04
  ygSTR5 = w����05
  
  fsInputCheck = True
End Function

