VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_�������t�͈�_�T 
   Caption         =   "���t���́i�͈́j�T"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7155
   OleObjectBlob   =   "Fy_�������t�͈�_�T.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_�������t�͈�_�T"
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
  
  '���t��(Label)
  If yg���t�� <> "" Then lbl���t = yg���t��
  
  '���t�����l
  If yg���t = "" Then
    Call s_���TSet
  Else
    yg���tST = Format(DateSerial(Year(yg���t), Month(yg���t) - 1, Day(yg���t) + 1), "yyyy/mm/dd")
    yg���tED = yg���t
    Me.txt���tST.Text = yg���tST
    Me.txt���tED.Text = yg���tED
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "�m�F : FormIni"
  Resume Next
End Sub

'2024/04/15 --------------------------------------------------*
Private Sub s_���TSet()
  Dim w�j      As Integer
  Dim w���tST  As String
  Dim w���tED  As String
  
  w�j = Weekday(Date)
  w���tST = Format(DateSerial(Year(Date), Month(Date), Day(Date) - (w�j - 1)), "yyyy/mm/dd")
  w���tED = Format(DateSerial(Year(w���tST), Month(w���tST), Day(w���tST) + 6), "yyyy/mm/dd")

  Me.txt���tST.Value = w���tST
  Me.txt���tED.Value = w���tED
End Sub

'2024/04/15 --------------------------------------------------*
Private Sub btn�O�T_Click()
  Dim w�j      As Integer
  Dim w���tST  As String
  Dim w���tED  As String
  
  w���tST = Me.txt���tST.Text
  If w���tST = "" Then
    w�j = Weekday(Date)
    w���tST = Format(DateSerial(Year(Date), Month(Date), Day(Date) - (w�j - 1) - 7), "yyyy/mm/dd")
  Else
    w���tST = Format(DateSerial(Year(w���tST), Month(w���tST), Day(w���tST) - 7), "yyyy/mm/dd")
  End If
  w���tED = Format(DateSerial(Year(w���tST), Month(w���tST), Day(w���tST) + 6), "yyyy/mm/dd")
  
  Me.txt���tST.Value = w���tST
  Me.txt���tED.Value = w���tED
End Sub

'2024/04/15 --------------------------------------------------*
Private Sub btn���T_Click()
  Call s_���TSet
End Sub

'2024/04/15 --------------------------------------------------*
Private Sub btn���T_Click()
  Dim w�j      As Integer
  Dim w���tST  As String
  Dim w���tED  As String
  
  w���tST = Me.txt���tST.Text
  If w���tST = "" Then
    w�j = Weekday(Date)
    w���tST = Format(DateSerial(Year(Date), Month(Date), Day(Date) - (w�j - 1) + 7), "yyyy/mm/dd")
  Else
    w���tST = Format(DateSerial(Year(w���tST), Month(w���tST), Day(w���tST) + 7), "yyyy/mm/dd")
  End If
  w���tED = Format(DateSerial(Year(w���tST), Month(w���tST), Day(w���tST) + 6), "yyyy/mm/dd")
  
  Me.txt���tST.Value = w���tST
  Me.txt���tED.Value = w���tED
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
  Dim w���tST  As String
  Dim w���tED  As String
  
  fsInputCheck = False
  
  w���tST = Me.txt���tST.Text
  w���tED = Me.txt���tED.Text
  
  '--- Check ---
  If IsDate(w���tST) = False Then
    wMSG = wMSG & "���t�i�J�n�j�𐳂������͂��Ă��������B�i��:2020/4/20�j" & vbCrLf
  End If
  
  If IsDate(w���tED) = False Then
    wMSG = wMSG & "���t�i�I���j�𐳂������͂��Ă��������B�i��:2020/4/20�j" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "�m�F"
    Exit Function
  End If
  
  yg���tST = w���tST
  yg���tED = w���tED
  
  fsInputCheck = True
End Function



