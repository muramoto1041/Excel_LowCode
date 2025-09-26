VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_メニューS 
   Caption         =   "メニュー"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3555
   OleObjectBlob   =   "Fy_メニューS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_メニューS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'****************************************************************
'作成日：2020/06/01
'                             Fy_メニュー
'作成者：村本俊和
'****************************************************************
Option Explicit
'System
Dim sMsg   As Integer
Dim sWhere As String
'Procedure
Dim sTitle    As String
Dim sMenuList As String
Dim sStrMsg   As String
Dim sMenuNo   As Integer
'YUGEコマンド
Dim sCommand       As String
Dim sクリック表示  As String

'2020/05/10 --------------------------------------------------*
Private Sub UserForm_Initialize()
  On Error GoTo subError
      
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15
  
  ygEnd = 0
  
  sTitle = ygSTR1
  sMenuList = ygSTR2
  sStrMsg = ygSTR3
  sMenuNo = ygInt1
  
  'YUGEコマンド
  sCommand = ""
  
  Call s_MenuList(sMenuList)
  
  'タイトル
  If sTitle <> "" Then
    'YUGE内Call
    If Left(sTitle, 1) = "$" And Right(sTitle, 1) = "$" And Len(sTitle) > 2 Then
      sTitle = Mid(sTitle, 2, Len(sTitle) - 2)
      sCommand = sTitle
      
      Select Case sCommand
        Case "シートメニュー": sクリック表示 = ygSTR4   'する、しない
      End Select
    End If
    
    Me.Caption = sTitle
  End If
  
  'メッセージ
  Me.lblMsg.Caption = sStrMsg
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : FormIni"
  Resume Next
End Sub

'2020/06/01 --------------------------------------------------*
Private Sub s_MenuList(qMenuList As String)
  Dim wNo As Integer
  Dim wMenuStr As String
  
  If qMenuList = "" Then Exit Sub
  
  Do
    wNo = wNo + 1
    wMenuStr = fy対応文字列(qMenuList, wNo)
    If wMenuStr = "" Then Exit Do
    
    Me.lstMenu.AddItem wMenuStr
  Loop
  
  If sMenuNo > 0 And sMenuNo < wNo Then
    Me.lstMenu.Selected(sMenuNo - 1) = True
  End If
End Sub

'2020/07/19 --------------------------------------------------*
Private Sub lstMenu_Click()
  '
  'YUGEコマンド【シートメニュー】
  'メニューをクリックするとシートを表示する。
  '
  On Error GoTo subError
  Dim wMenu  As String
  Dim wNo    As Integer
  Dim wCnt   As Integer
  
  If Not (sCommand = "シートメニュー" And (sクリック表示 = "する" Or sクリック表示 = "")) Then Exit Sub
  
  wCnt = Me.lstMenu.ListCount
  
  For wNo = 0 To wCnt - 1
    If Me.lstMenu.Selected(wNo) = True Then
      wMenu = Me.lstMenu.List(wNo)
      Exit For
    End If
  Next wNo
  
  'シート表示
  NN = fn_シート表示(wMenu, "", True)
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : lstMenu"
  Resume subExit
End Sub

'2020/05/10 --------------------------------------------------*
Private Sub cmd選択_Click()
  On Error GoTo subError
  Dim wMSG    As String
  Dim wMenu   As String
  Dim wNo     As Integer
  Dim wCnt    As Integer
  
  wCnt = Me.lstMenu.ListCount
  
  sMenuNo = 0
  
  For wNo = 0 To wCnt - 1
    If Me.lstMenu.Selected(wNo) = True Then
      wMenu = Me.lstMenu.List(wNo)
      sMenuNo = wNo + 1
      Exit For
    End If
  Next wNo
  
  If sMenuNo = 0 Then
    wMSG = "メニューを選択してください。"
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Sub
  End If
  
  ygEnd = 1
  
  ygInt1 = sMenuNo
  Unload Me
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : 選択"
  Resume subExit
End Sub

'2020/07/20 --------------------------------------------------*
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  
End Sub

Public Sub Test3()
  MsgBox "Test3"
End Sub

