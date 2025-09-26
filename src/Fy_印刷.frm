VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fy_印刷 
   Caption         =   "印刷"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   OleObjectBlob   =   "Fy_印刷.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Fy_印刷"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'作成日：2020/06/07
'                             Fy_印刷
'作成者：村本俊和
'****************************************************************
Option Explicit
'System
Dim sMsg   As Integer
Dim sWhere As String
'Procedure
Dim sFL拡張      As Integer    '0:標準 1:拡張
Dim s条件NO      As Integer
Dim s印刷Sheet   As String
Dim s項目名(20)  As String
Dim s明細名(20)  As String
Dim s明細行数    As Integer
Dim sQueryName01 As String
Dim sQueryName02 As String
Dim sCardNO      As Integer
Dim sIDCard      As Long
Dim s一括件数    As Long
Dim s印刷枚数    As Integer

'2020/05/13 --------------------------------------------------*
Private Sub UserForm_Initialize()
  On Error GoTo subError
  
  '--- FormSize ---
  Me.Height = Me.HRheight.Height + 30
  Me.Width = Me.HRwidth.Width + 15
  
subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : FormIni"
  Resume Next
End Sub

'2020/06/08 --------------------------------------------------*
Private Sub cmdPreview_Click()
  On Error GoTo subError

  '--- プレビュー ---
  ygCntPrt = 1
  Me.Hide
  NN = fn_シート印刷(yg印刷Sheet, ygStartBook, "する", 1)
  Me.Show

subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : Preview"
  Resume subExit
End Sub

'2020/06/08 --------------------------------------------------*
Private Sub cmd印刷_Click()
  On Error GoTo subError

  '--- 印刷 ---
  ygCntPrt = 1
  Me.Hide
  NN = fn_シート印刷(yg印刷Sheet, ygStartBook, "しない", 1)
  Me.Show

subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : Preview"
  Resume subExit
End Sub

'2022/07/20 --------------------------------------------------*
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  On Error GoTo subError
  
  If yg印刷Sheet <> "" Then
    NN = fn_シート初期化(yg印刷Sheet, "")
  End If

subExit:
  Exit Sub

subError:
  MsgBox Error$ & "(#" & Trim(Trim(Err.Number)) & ")", vbOKOnly + vbExclamation _
  , "確認 : QueryClose"
  Resume Next
End Sub
