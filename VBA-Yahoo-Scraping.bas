'Written by MICHAEL CORLEY (please leave in code)
'https://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.querytable_properties.aspx

Sub testrun()
'Worksheets("Input_Sheet").Range("A1:A303")
Call Yahoo_Data(Worksheets("Input_Sheet").Range("A1:A599"), 5, 1)
End Sub

Sub Yahoo_Data(arr_input, data_point, data_catagory)
lr = arr_input.Rows.Count

For i = 1 To lr
Call AddTable(arr_input.Cells(i, 1).value, data_point, data_catagory)
Next
End Sub


Sub AddTable(ByVal ticker As String, ByVal data_point As Long, ByVal data_catagory As Long)
'Analysis:
'2 = earnings estimates
'3 = revenue estimates
'4 = earnings history
'5 = EPS trend (aging schedule for estimates revisions)
'6 = EPS revisions in last x days
'7 = growth estimates
'there is nothing after 8 or before 2
'About
'2 = executives list

Select Case data_catagory

Case Is = 1 'Analyst estimates
catagory = "/analysis?p="

Case Is = 2 'About (Executives List)
catagory = "/profile?p="
End Select



lastrow = WorksheetFunction.CountA(Worksheets("Output_Sheet").Range("A:A"))

Worksheets("Output_Sheet").Cells(lastrow + 1, 1) = ticker

With Worksheets("Output_Sheet").QueryTables.Add(Connection:="URL;https://finance.yahoo.com/quote/" & ticker & catagory & ticker, _
            Destination:=Worksheets("Output_Sheet").Cells(lastrow + 2, 1))
    .WebTables = data_point
    .WebFormatting = xlWebFormattingNone
    .EnableRefresh = False
    .RefreshStyle = xlInsertEntireRows
    .BackgroundQuery = False
    .Refresh
End With

Call Delete_Connections

End Sub

Sub Delete_Connections()
Do While ActiveWorkbook.Connections.Count > 0
ActiveWorkbook.Connections.Item(ActiveWorkbook.Connections.Count).Delete
Loop
End Sub
