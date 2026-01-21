Public Sub p_列表示切替()
  Dim wSheet  As String
  
  wSheet = ActiveSheet.Name

  With ThisWorkbook.Worksheets(wSheet)
    If .Columns("A:C").Hidden = False Then
      ' 現在表示されている場合は非表示にする
      .Columns("A:C").Hidden = True
        .Range("D1").Select
    Else
      ' 現在非表示の場合は表示する
      .Columns("A:C").Hidden = False
      .Range("A1").Select
    End If
  End With
End Sub
