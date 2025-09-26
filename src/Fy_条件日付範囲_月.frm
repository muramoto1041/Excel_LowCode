VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_条件日付範囲_月 
   Caption         =   "日付入力（範囲）月"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7155
   OleObjectBlob   =   "Fy_条件日付範囲_月.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_条件日付範囲_月"
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
    Me.txt日付ST.Value = Null
    Me.txt日付ED.Value = Null
  Else
    yg日付ST = Format(DateSerial(Year(yg日付), Month(yg日付) - 1, Day(yg日付) + 1), "yyyy/mm/dd")
    yg日付ED = yg日付
    Me.txt日付ST.Text = yg日付ST
    Me.txt日付ED.Text = yg日付ED
  End If
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : FormIni"
  Resume Next
End Sub

'2020/07/25 --------------------------------------------------*
Private Sub btn前月_Click()
  Dim w日付ST  As String
  Dim w日付ED  As String
  
  w日付ED = Me.txt日付ED.Text
  If w日付ED = "" Then w日付ED = Date

  w日付ST = Format(DateSerial(Year(w日付ED), Month(w日付ED) - 1, 1), "yyyy/mm/dd")
  w日付ED = Format(DateSerial(Year(w日付ST), Month(w日付ST) + 1, 0), "yyyy/mm/dd")
  
  Me.txt日付ST.Value = w日付ST
  Me.txt日付ED.Value = w日付ED
End Sub

'2020/07/25 --------------------------------------------------*
Private Sub btn今月_Click()
  Dim w日付ST  As String
  Dim w日付ED  As String

  w日付ST = Format(DateSerial(Year(Date), Month(Date), 1), "yyyy/mm/dd")
  w日付ED = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "yyyy/mm/dd")
  
  Me.txt日付ST.Value = w日付ST
  Me.txt日付ED.Value = w日付ED
End Sub

'2020/07/25 --------------------------------------------------*
Private Sub btn次月_Click()
  Dim w日付ST  As String
  Dim w日付ED  As String
  
  w日付ED = Me.txt日付ED.Text
  If w日付ED = "" Then w日付ED = Date

  w日付ST = Format(DateSerial(Year(w日付ED), Month(w日付ED) + 1, 1), "yyyy/mm/dd")
  w日付ED = Format(DateSerial(Year(w日付ST), Month(w日付ST) + 1, 0), "yyyy/mm/dd")
  
  Me.txt日付ST.Value = w日付ST
  Me.txt日付ED.Value = w日付ED
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
  Dim w日付ST  As String
  Dim w日付ED  As String
  
  fsInputCheck = False
  
  w日付ST = Me.txt日付ST.Text
  w日付ED = Me.txt日付ED.Text
  
  '--- Check ---
  If IsDate(w日付ST) = False Then
    wMSG = wMSG & "日付（開始）を正しく入力してください。（例:2020/4/20）" & vbCrLf
  End If
  
  If IsDate(w日付ED) = False Then
    wMSG = wMSG & "日付（終了）を正しく入力してください。（例:2020/4/20）" & vbCrLf
  End If

  If wMSG <> "" Then
    MsgBox wMSG, vbOKOnly + vbExclamation, "確認"
    Exit Function
  End If
  
  yg日付ST = w日付ST
  yg日付ED = w日付ED
  
  fsInputCheck = True
End Function



