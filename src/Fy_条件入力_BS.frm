VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_条件入力_BS 
   Caption         =   "条件入力"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   OleObjectBlob   =   "Fy_条件入力_BS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_条件入力_Bs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'作成日：2020/06/07
'                             Fy_条件入力
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
  
  '条件名
  If yg条件名 <> "" Then lbl条件名 = yg条件名
  
  '条件初期値
  If yg条件 = "" Then
    Me.txt条件.Value = Null
  Else
    Me.txt条件.Text = yg条件
  End If
  
  'メッセージ
  If ygSTR1 = "" Then
    Me.lblMsg.Caption = ""
  Else
    Me.lblMsg.Caption = ygSTR1
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
  Dim w条件    As String
  
  fsInputCheck = False
  
  w条件 = Me.txt条件.Text
  
  '--- Check ---
  If w条件 = "" Then
    wMSG = wMSG & "条件を入力してください。" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  yg条件 = w条件
  
  fsInputCheck = True
End Function

