Attribute VB_Name = "Module6"
Sub US_Cleaning()
Attribute US_Cleaning.VB_ProcData.VB_Invoke_Func = " \n14"
'
' US_Cleaning Macro
'
    
'
    
    Workbooks.Open " T:\**Location_of_First_File.xlsx** "
    Workbooks.Open " T:\**Location_of_Second_File.xlsx** "
    
    Windows("First_File.xlsx").Activate
    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearContents
    
    Windows("Second_File.xlsx").Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("First_File.xlsx.xlsx").Activate
    Range("B2").Select
    ActiveSheet.Paste
    Range("A2").Select
    lr = Cells.Find("*", Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious, False).Row
    Application.CutCopyMode = False
    
    
    Selection.AutoFill Destination:=Range("A2:A" & lr)
    Range("A2:A3603").Select
    
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range( _
        "AI2:AI" & lr), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:AJ6002")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.Select
    ActiveSheet.Range("$A$1:$AJ" & lr).RemoveDuplicates Columns:=1, Header:= _
        xlYes
        
        
        
    Windows("First_File.xlsx").Activate
    ActiveWorkbook.Save
    
    
End Sub
