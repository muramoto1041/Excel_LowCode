VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_�����N���͈� 
   Caption         =   "�N�����́i�͈́j"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8985.001
   OleObjectBlob   =   "Fy_�����N���͈�.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_�����N���͈�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'�쐬���F2020/10/05
'                             Fy_�����N���i�͈́j
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
  Dim w�NST As Integer
  Dim w��ST As Integer
  Dim w�NED As Integer
  Dim w��ED As Integer
  
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15

  ygEnd = 0
  
  '�N�������l
  If yg�N�� = "" Then
    Me.cbo�NST.Value = Null: Me.cbo��ST.Value = Null
    Me.cbo�NED.Value = Null: Me.cbo��ED.Value = Null
  Else
    w�NST = Year(yg�N�� & "/01") - 1: w��ST = Month(yg�N�� & "/01") + 1
    If w��ST > 12 Then w��ST = 1
    w�NED = Year(yg�N�� & "/01"):     w��ED = Month(yg�N�� & "/01")
    
    Me.cbo�NST.Value = w�NST: Me.cbo��ST.Value = w��ST
    Me.cbo�NED.Value = w�NED: Me.cbo��ED.Value = w��ED
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
    Me.cbo�NST.AddItem Year(Date) - 3 + wNo
    Me.cbo�NED.AddItem Year(Date) - 3 + wNo
  Next wNo
  
  '(��)
  For wNo = 1 To 12
    Me.cbo��ST.AddItem wNo
    Me.cbo��ED.AddItem wNo
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
  Dim w�NST    As String
  Dim w��ST    As String
  Dim w�NED    As String
  Dim w��ED    As String
  Dim w�N��ST  As String
  Dim w�N��ED  As String
  
  fsInputCheck = False
  
  w�NST = Me.cbo�NST.Text: w��ST = Me.cbo��ST.Text
  w�NED = Me.cbo�NED.Text: w��ED = Me.cbo��ED.Text
  w�N��ST = w�NST & "/" & w��ST & "/1"
  w�N��ED = w�NED & "/" & w��ED & "/1"
  
  '--- Check ---
  If IsDate(w�N��ST) = False Then
    wMSG = wMSG & "�N��(�J�n)�𐳂������͂��Ă��������B�i�N�F2020 etc. ���F1�`12�j" & vbCrLf
  End If

  If IsDate(w�N��ED) = False Then
    wMSG = wMSG & "�N��(�I��)�𐳂������͂��Ă��������B�i�N�F2020 etc. ���F1�`12�j" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Function
  End If
  
  yg�N��ST = Format(w�N��ST, "yyyy/mm")
  yg�N��ED = Format(w�N��ED, "yyyy/mm")
  
  fsInputCheck = True
End Function

