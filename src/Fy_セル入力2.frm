VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_セル入力2 
   Caption         =   "セル入力2"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "Fy_セル入力2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_セル入力2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'作成日：2020/06/07
'                             Fy_セル入力
'作成者：村本俊和
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
  
  'ラベル名
  If ygLBL1 <> "" Then Me.lbl入力名01 = ygLBL1
  If ygLBL2 <> "" Then Me.lbl入力名02 = ygLBL2
  
  '入力初期値
  '(1)
  If ygSTR1 = "" Then
    Me.txt入力01.Value = Null
  Else
    Me.txt入力01.Text = ygSTR1
  End If
  '(2)
  If ygSTR2 = "" Then
    Me.txt入力02.Value = Null
  Else
    Me.txt入力02.Text = ygSTR2
  End If
  
  'メッセージ
  If ygMSG = "" Then
    Me.lblMsg.Caption = ""
  Else
    Me.lblMsg.Caption = ygMSG
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : FormIni"
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
  , "確認 : cmdOK"
  Resume subExit
End Sub

'2020/07/25 --------------------------------------------------*
Private Function fsInputCheck() As Boolean
  Dim wMSG     As String
  Dim w入力01  As String, w入力02 As String
  
  fsInputCheck = False
  
  w入力01 = Me.txt入力01.Text
  w入力02 = Me.txt入力02.Text
  
  '--- Check ---
  If w入力01 = "" Then
    wMSG = wMSG & "値を入力してください。" & vbCrLf
  End If
  If w入力02 = "" Then
    wMSG = wMSG & "値を入力してください。" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  ygSTR1 = w入力01
  ygSTR2 = w入力02
  
  fsInputCheck = True
End Function

