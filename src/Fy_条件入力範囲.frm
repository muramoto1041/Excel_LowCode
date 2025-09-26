VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_条件入力範囲 
   Caption         =   "条件入力（範囲）"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   OleObjectBlob   =   "Fy_条件入力範囲.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_条件入力範囲"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************
'作成日：2020/06/07
'                             Fy_条件入力（範囲）
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
  Dim w条件ST  As String
  Dim w条件ED  As String
  
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15

  ygEnd = 0
  
  '条件名
  If yg条件名 <> "" Then lbl条件名 = yg条件名
  
  If InStr(yg条件, ",") > 0 And Len(yg条件) > 1 Then
    yg条件ST = Left(yg条件, InStr(yg条件, ",") - 1)
    yg条件ED = Mid(yg条件, InStr(yg条件, ",") + 1, Len(yg条件) - InStr(yg条件, ","))
  End If
  
  '条件初期値
  Me.txt条件ST.Text = yg条件ST
  Me.txt条件ED.Text = yg条件ED
  
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
  Dim w条件ST  As String
  Dim w条件ED  As String
  
  fsInputCheck = False
  
  w条件ST = Me.txt条件ST.Text
  w条件ED = Me.txt条件ED.Text
  
  '--- Check ---
  If w条件ST = "" And w条件ED = "" Then
    wMSG = wMSG & "開始条件または終了条件を入力してください。" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  yg条件ST = w条件ST
  yg条件ED = w条件ED
  
  fsInputCheck = True
End Function



