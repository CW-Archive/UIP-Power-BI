Attribute VB_Name = "CreateTradeExportTable_Module"
Sub CreateTradeExportTable_Sub()
    Dim SheetName As String
    Dim IsTradeSheet As Boolean
    Dim ExportTable_ColCount As Long, ac As Long, r As Long
    Dim Works_Sun As Boolean, Works_Mon As Boolean, Works_Tue As Boolean, Works_Wed As Boolean, Works_Thu As Boolean, Works_Fri As Boolean, Works_Sat As Boolean
    
    
    ' Check that active sheet is a trade sheet
    SheetName = ActiveSheet.Name
    IsTradeSheet = strOut(SheetName)
    
    If IsTradeSheet = False Then
        e = MsgBox("This isn't a trade sheet. Please select a trade sheet and try again.", vbExclamation, "Select Trade Sheet")
        Exit Sub
    End If
    
    ' check that the table hasn't already been made
    ExportTable_ColCount = Range("ExportTable_" & SheetName).ListObject.ListColumns.Count
    ExportTable_RowCount = Range("ExportTable_" & SheetName).ListObject.ListRows.Count
    
    If ExportTable_ColCount > 1 Or ExportTable_RowCount > 1 Then
        e = MsgBox("It looks like the Trade Export table has already been created. You can update the schedule by using the update schedule button but if you add areas you will need to create a new trade sheet and copy any previous production manually. This is all done to prevent data loss.", vbExclamation, "Table Already Exists!")
        Exit Sub
    End If
    
    ' import data / settings
    Works_Sun = IIf(Range("C11").Value = "YES", True, False)
    Works_Mon = IIf(Range("C12").Value = "YES", True, False)
    Works_Tue = IIf(Range("C13").Value = "YES", True, False)
    Works_Wed = IIf(Range("C14").Value = "YES", True, False)
    Works_Thu = IIf(Range("C15").Value = "YES", True, False)
    Works_Fri = IIf(Range("C16").Value = "YES", True, False)
    Works_Sat = IIf(Range("C17").Value = "YES", True, False)
    
    AreasTable = Range("AreasTable_" & SheetName).ListObject.DataBodyRange
    
    ' create table columns
    For ac = 0 To Range("AreasTable_" & SheetName).ListObject.ListRows.Count
        If ac = 0 Then
            Range("ExportTable_" & SheetName).ListObject.ListColumns(1).Name = "Date"
        Else
            Range("ExportTable_" & SheetName).ListObject.ListColumns.Add.Name = "PlanTotal_" & AreasTable(ac, 1)
            Range("ExportTable_" & SheetName).ListObject.ListColumns.Add.Name = "CompTotal_" & AreasTable(ac, 1)
        End If
    Next ac
    
    ' get first startdate and last finish date and create rows
    For r = 1 To Range("AreasTable_" & SheetName).ListObject.ListRows.Count
        If r = 1 Then
            StartDate_AllAreas = AreasTable(r, 4)
        Else
            If AreasTable(r, 4) < StartDate_AllAreas Then StartDate_AllAreas = AreasTable(r, 4)
        End If
    Next r
    
    For r = 1 To Range("AreasTable_" & SheetName).ListObject.ListRows.Count
        If r = 1 Then
            FinishDate_AllAreas = AreasTable(r, 5)
        Else
            If AreasTable(r, 5) > FinishDate_AllAreas Then FinishDate_AllAreas = AreasTable(r, 5)
        End If
    Next r
    
    ' count the number of working days in an area and total/working days for each area
    
    NumRowsToCreate = FinishDate_AllAreas - StartDate_AllAreas + 1
    
    For rc = 1 To NumRowsToCreate
        If rc = 1 Then
            Range("ExportTable_" & SheetName).ListObject.Range(rc + 1, 1) = StartDate_AllAreas
        Else
            Range("ExportTable_" & SheetName).ListObject.ListRows.Add
            Range("ExportTable_" & SheetName).ListObject.Range(rc + 1, 1) = StartDate_AllAreas + rc - 1
        End If
    Next rc
    
End Sub
