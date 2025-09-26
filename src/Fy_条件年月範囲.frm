VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_条件年月範囲 
   Caption         =   "年月入力（範囲）"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8985.001
   OleObjectBlob   =   "Fy_条件年月範囲.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_条件年月範囲"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'作成日：2020/10/05
'                             Fy_条件年月（範囲）
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
  Dim w年ST As Integer
  Dim w月ST As Integer
  Dim w年ED As Integer
  Dim w月ED As Integer
  
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15

  ygEnd = 0
  
  '年月初期値
  If yg年月 = "" Then
    Me.cbo年ST.Value = Null: Me.cbo月ST.Value = Null
    Me.cbo年ED.Value = Null: Me.cbo月ED.Value = Null
  Else
    w年ST = Year(yg年月 & "/01") - 1: w月ST = Month(yg年月 & "/01") + 1
    If w月ST > 12 Then w月ST = 1
    w年ED = Year(yg年月 & "/01"):     w月ED = Month(yg年月 & "/01")
    
    Me.cbo年ST.Value = w年ST: Me.cbo月ST.Value = w月ST
    Me.cbo年ED.Value = w年ED: Me.cbo月ED.Value = w月ED
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
    Me.cbo年ST.AddItem Year(Date) - 3 + wNo
    Me.cbo年ED.AddItem Year(Date) - 3 + wNo
  Next wNo
  
  '(月)
  For wNo = 1 To 12
    Me.cbo月ST.AddItem wNo
    Me.cbo月ED.AddItem wNo
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
  Dim w年ST    As String
  Dim w月ST    As String
  Dim w年ED    As String
  Dim w月ED    As String
  Dim w年月ST  As String
  Dim w年月ED  As String
  
  fsInputCheck = False
  
  w年ST = Me.cbo年ST.Text: w月ST = Me.cbo月ST.Text
  w年ED = Me.cbo年ED.Text: w月ED = Me.cbo月ED.Text
  w年月ST = w年ST & "/" & w月ST & "/1"
  w年月ED = w年ED & "/" & w月ED & "/1"
  
  '--- Check ---
  If IsDate(w年月ST) = False Then
    wMSG = wMSG & "年月(開始)を正しく入力してください。（年：2020 etc. 月：1～12）" & vbCrLf
  End If

  If IsDate(w年月ED) = False Then
    wMSG = wMSG & "年月(終了)を正しく入力してください。（年：2020 etc. 月：1～12）" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  yg年月ST = Format(w年月ST, "yyyy/mm")
  yg年月ED = Format(w年月ED, "yyyy/mm")
  
  fsInputCheck = True
End Function

