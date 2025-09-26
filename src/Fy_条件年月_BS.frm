VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_条件年月_BS 
   Caption         =   "年月入力"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "Fy_条件年月_BS.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_条件年月_BS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************
'作成日：2020/10/05
'                             Fy_条件年月
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
  
  '日付初期値
  If yg年月 = "" Then
    Me.cbo年.Value = Null
    Me.cbo月.Value = Null
  Else
    Me.cbo年.Value = Year(yg年月 & "/01")
    Me.cbo月.Value = Month(yg年月 & "/01")
  End If
  
  
  'メッセージ表示
  If ygSTR1 <> "" Then
    Me.lblMsg.Caption = ygSTR1
  End If
  
  Call s_SetCombo年月
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : FormIni"
  Resume Next
End Sub

'2020/10/05 --------------------------------------------------*
Private Sub s_SetCombo年月()
  Dim wNo  As Integer
  
  '(年)
  For wNo = 1 To 5
    Me.cbo年.AddItem Year(Date) - 3 + wNo
  Next wNo
  
  '(月)
  For wNo = 1 To 12
    Me.cbo月.AddItem wNo
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
  , "確認 : cmdOK"
  Resume subExit
End Sub

'2020/07/25 --------------------------------------------------*
Private Function fsInputCheck() As Boolean
  Dim wMSG     As String
  Dim w年      As String
  Dim w月      As String
  Dim w年月    As String
  
  fsInputCheck = False
  
  w年 = Me.cbo年.Text
  w月 = Me.cbo月.Text
  w年月 = w年 & "/" & w月 & "/1"
  
  '--- Check ---
  If IsDate(w年月) = False Then
    wMSG = wMSG & "年月を正しく入力してください。（年：2020 etc. 月：1～12）" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  yg年月 = Format(w年月, "yyyy/mm")
  
  fsInputCheck = True
End Function

