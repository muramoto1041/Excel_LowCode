VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_条件日付_BS 
   Caption         =   "日付入力"
   ClientHeight    =   2595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   OleObjectBlob   =   "Fy_条件日付_BS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_条件日付_Bs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************
'作成日：2020/06/07
'                             Fy_条件日付
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
  
  '日付名
  If yg日付名 <> "" Then lbl日付 = yg日付名
  
  '日付初期値
  If yg日付 = "" Then
    Me.txt日付.Value = Null
  Else
    Me.txt日付.Text = yg日付
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : FormIni"
  Resume Next
End Sub

'2021/11/05 --------------------------------------------------*
Private Sub btn前日_Click()
  Dim w日付 As String

  w日付 = Me.txt日付.Text
  If w日付 = "" Then w日付 = Date
  w日付 = Format(DateSerial(Year(w日付), Month(w日付), Day(w日付) - 1), "yyyy/mm/dd")
  
  Me.txt日付.Value = w日付
End Sub

Private Sub btn今日_Click()
  Dim w日付 As String

  w日付 = Format(Date, "yyyy/mm/dd")
  Me.txt日付.Value = w日付
End Sub

Private Sub btn次日_Click()
  Dim w日付 As String

  w日付 = Me.txt日付.Text
  If w日付 = "" Then w日付 = Date
  w日付 = Format(DateSerial(Year(w日付), Month(w日付), Day(w日付) + 1), "yyyy/mm/dd")
  
  Me.txt日付.Value = w日付
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
  Dim w日付    As String
  
  fsInputCheck = False
  
  w日付 = Me.txt日付.Text
  
  '--- Check ---
  If IsDate(w日付) = False Then
    wMSG = wMSG & "日付を正しく入力してください。（例:2020/4/20）" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  yg日付 = w日付
  
  fsInputCheck = True
End Function

