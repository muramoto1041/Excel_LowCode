VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_�������t�͈�_�� 
   Caption         =   "���t���́i�͈́j��"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7155
   OleObjectBlob   =   "Fy_�������t�͈�_��.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Fy_�������t�͈�_��"
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
    Me.txt���tST.Value = Null
    Me.txt���tED.Value = Null
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

'2020/07/25 --------------------------------------------------*
Private Sub btn�O��_Click()
  Dim w���tST  As String
  Dim w���tED  As String
  
  w���tED = Me.txt���tED.Text
  If w���tED = "" Then w���tED = Date

  w���tST = Format(DateSerial(Year(w���tED), Month(w���tED) - 1, 1), "yyyy/mm/dd")
  w���tED = Format(DateSerial(Year(w���tST), Month(w���tST) + 1, 0), "yyyy/mm/dd")
  
  Me.txt���tST.Value = w���tST
  Me.txt���tED.Value = w���tED
End Sub

'2020/07/25 --------------------------------------------------*
Private Sub btn����_Click()
  Dim w���tST  As String
  Dim w���tED  As String

  w���tST = Format(DateSerial(Year(Date), Month(Date), 1), "yyyy/mm/dd")
  w���tED = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "yyyy/mm/dd")
  
  Me.txt���tST.Value = w���tST
  Me.txt���tED.Value = w���tED
End Sub

'2020/07/25 --------------------------------------------------*
Private Sub btn����_Click()
  Dim w���tST  As String
  Dim w���tED  As String
  
  w���tED = Me.txt���tED.Text
  If w���tED = "" Then w���tED = Date

  w���tST = Format(DateSerial(Year(w���tED), Month(w���tED) + 1, 1), "yyyy/mm/dd")
  w���tED = Format(DateSerial(Year(w���tST), Month(w���tST) + 1, 0), "yyyy/mm/dd")
  
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



