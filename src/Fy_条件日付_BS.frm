VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_�������t_BS 
   Caption         =   "���t����"
   ClientHeight    =   2595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   OleObjectBlob   =   "Fy_�������t_BS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_�������t_Bs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************
'�쐬���F2020/06/07
'                             Fy_�������t
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
  
  '���t��
  If yg���t�� <> "" Then lbl���t = yg���t��
  
  '���t�����l
  If yg���t = "" Then
    Me.txt���t.Value = Null
  Else
    Me.txt���t.Text = yg���t
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : FormIni"
  Resume Next
End Sub

'2021/11/05 --------------------------------------------------*
Private Sub btn�O��_Click()
  Dim w���t As String

  w���t = Me.txt���t.Text
  If w���t = "" Then w���t = Date
  w���t = Format(DateSerial(Year(w���t), Month(w���t), Day(w���t) - 1), "yyyy/mm/dd")
  
  Me.txt���t.Value = w���t
End Sub

Private Sub btn����_Click()
  Dim w���t As String

  w���t = Format(Date, "yyyy/mm/dd")
  Me.txt���t.Value = w���t
End Sub

Private Sub btn����_Click()
  Dim w���t As String

  w���t = Me.txt���t.Text
  If w���t = "" Then w���t = Date
  w���t = Format(DateSerial(Year(w���t), Month(w���t), Day(w���t) + 1), "yyyy/mm/dd")
  
  Me.txt���t.Value = w���t
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
  Dim w���t    As String
  
  fsInputCheck = False
  
  w���t = Me.txt���t.Text
  
  '--- Check ---
  If IsDate(w���t) = False Then
    wMSG = wMSG & "���t�𐳂������͂��Ă��������B�i��:2020/4/20�j" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Function
  End If
  
  yg���t = w���t
  
  fsInputCheck = True
End Function

