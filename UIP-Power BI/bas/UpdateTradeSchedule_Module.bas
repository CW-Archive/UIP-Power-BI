Attribute VB_Name = "UpdateTradeSchedule_Module"
Option Explicit
Sub UpdateTradeSchedule_Sub()
    Application.ScreenUpdating = False
    Dim SheetName As String
    Dim IsTradeSheet As Boolean
    Dim Works_Sun As Boolean, Works_Mon As Boolean, Works_Tue As Boolean, Works_Wed As Boolean, Works_Thu As Boolean, Works_Fri As Boolean, Works_Sat As Boolean
    Dim e
    Dim ExportTable_ColCount As Long, ExportTable_RowCount As Long, r As Long, AreasTable_ColCount As Long, AreasTable_RowCount As Long, i As Long, i_hacked As Long, NumRowsToCreate As Long, rc As Long, atr As Long, ExportTable_RowCountAfter As Long
    Dim AreasTable, ExportTable
    Dim StartDate_AllAreas As Date, FinishDate_AllAreas As Date
    Dim atr_days As Long, atr_workingdays As Long, atr_daysCounter As Long, atr_dayupdate As Long, x As Long, indexcol As Long, exr As Long
    Dim atr_dc_today As Date
    Dim atr_dc_isHoliday As Boolean, atr_dc_isWorkDay As Boolean
    Dim d
    Dim atr_production As Double
    Dim op_array As Long, op_table As Long, areas_count As Long, indexcol2 As Long
    
    ' check to see if you are on a trade sheet
    SheetName = ActiveSheet.Name
    IsTradeSheet = strOut(SheetName)
    
    If IsTradeSheet = False Then
        e = MsgBox("This isn't a trade sheet. Please select a trade sheet and try again.", vbExclamation, "Select Trade Sheet")
        Exit Sub
    End If
    ' check to see if the table has been created if not create it with CreateTradeExportTable
    ExportTable_ColCount = Range("ExportTable_" & SheetName).ListObject.ListColumns.Count
    ExportTable_RowCount = Range("ExportTable_" & SheetName).ListObject.ListRows.Count
    
    If ExportTable_ColCount <= 1 Or ExportTable_RowCount <= 1 Then
        CreateTradeExportTable_Sub

    End If
    
    ' Import Settings / Tables
    Works_Sun = IIf(Range("C11").Value = "YES", True, False)
    Works_Mon = IIf(Range("C12").Value = "YES", True, False)
    Works_Tue = IIf(Range("C13").Value = "YES", True, False)
    Works_Wed = IIf(Range("C14").Value = "YES", True, False)
    Works_Thu = IIf(Range("C15").Value = "YES", True, False)
    Works_Fri = IIf(Range("C16").Value = "YES", True, False)
    Works_Sat = IIf(Range("C17").Value = "YES", True, False)
    
    AreasTable = Range("AreasTable_" & SheetName).ListObject.DataBodyRange
    ExportTable = Range("ExportTable_" & SheetName).ListObject.DataBodyRange
    
    AreasTable_ColCount = Range("AreasTable_" & SheetName).ListObject.ListColumns.Count
    AreasTable_RowCount = Range("AreasTable_" & SheetName).ListObject.ListRows.Count
    
    ' check Start and Finish dates
    For atr = 1 To AreasTable_RowCount
        If AreasTable(atr, 5) - AreasTable(atr, 4) < 0 Then
            e = MsgBox(AreasTable(atr, 1) & " " & vbNewLine & vbNewLine & "Start Date (" & AreasTable(atr, 4) & ")" & vbNewLine & "Finish Date (" & AreasTable(atr, 5) & ")" & vbNewLine & vbNewLine & "Please update the schedule and try again.", vbExclamation, "Start / Finish Date Issue...")
            Exit Sub
        End If
    Next atr
    
    ' check to see if the export table is still correctly shaped (View Columns and Start / Finish Dates)
    ' Columns still correct?

    If AreasTable_RowCount = (ExportTable_ColCount - 1) / 2 Then
        If Range("Debug_Mode") = True Then Debug.Print "ExportTable is the correct size, " & ExportTable_ColCount & " columns."
    Else
        e = MsgBox("You have added or removed rows from the Areas Table. Rows cannot be added/removed automatically after creation. Please create a new Trade Sheet and copy any existing data to it. It must be done this way to prevent data loss.", vbExclamation, "Select Trade Sheet")
        Exit Sub
    
    End If
    
    i_hacked = 1
    For i = 2 To ExportTable_ColCount Step 2
        If Right(Range("ExportTable_" & SheetName).ListObject.ListColumns(i).Name, Len(Range("ExportTable_" & SheetName).ListObject.ListColumns(i).Name) - InStrRev(Range("ExportTable_" & SheetName).ListObject.ListColumns(i).Name, "_")) = AreasTable(i_hacked, 1) Then
        Else
            e = MsgBox("You have changed a row MID in the Areas Table. MID's cannot be added/removed/changed automatically after creation. Please create a new Trade Sheet and copy any existing data to it. It must be done this way to prevent data loss." & vbNewLine & vbNewLine & "You may have also changed some column names in the Trade Export table. You broke it. Create an new sheet and copy any data.", vbExclamation, "Select Trade Sheet")
            Exit Sub
        End If
        If Right(Range("ExportTable_" & SheetName).ListObject.ListColumns(i + 1).Name, Len(Range("ExportTable_" & SheetName).ListObject.ListColumns(i + 1).Name) - InStrRev(Range("ExportTable_" & SheetName).ListObject.ListColumns(i + 1).Name, "_")) = AreasTable(i_hacked, 1) Then
        Else
            e = MsgBox("You have changed a row MID in the Areas Table. MID's cannot be added/removed/changed automatically after creation. Please create a new Trade Sheet and copy any existing data to it. It must be done this way to prevent data loss." & vbNewLine & vbNewLine & "You may have also changed some column names in the Trade Export table. You broke it. Create an new sheet and copy any data.", vbExclamation, "Select Trade Sheet")
            Exit Sub
        End If
        
        i_hacked = i_hacked + 1
    Next i
    
    ' Rows still correct?
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

    NumRowsToCreate = (FinishDate_AllAreas - StartDate_AllAreas + 1) - ExportTable_RowCount
    
    ' Update row count
    If NumRowsToCreate > 0 Then
        For rc = 1 To NumRowsToCreate
            Range("ExportTable_" & SheetName).ListObject.ListRows.Add
        Next rc
    End If
    If NumRowsToCreate < 0 Then
        For rc = 1 To Abs(NumRowsToCreate)
            Range("ExportTable_" & SheetName).ListObject.ListRows(Range("ExportTable_" & SheetName).ListObject.ListRows.Count).Delete
        Next rc
    End If
    'update row dates
    
    For rc = 1 To NumRowsToCreate + ExportTable_RowCount
        If rc = 1 Then
            Range("ExportTable_" & SheetName).ListObject.Range(rc + 1, 1) = StartDate_AllAreas
        Else
            Range("ExportTable_" & SheetName).ListObject.Range(rc + 1, 1) = StartDate_AllAreas + rc - 1
        End If
    Next rc
        
    ' distribute production plan from each area to each day

    ' add production
    ExportTable_RowCountAfter = Range("ExportTable_" & SheetName).ListObject.ListRows.Count
    
    For atr = 1 To AreasTable_RowCount
        ' 0 things out
        For exr = 1 To ExportTable_RowCountAfter
            indexcol = Range("ExportTable_" & SheetName).ListObject.ListColumns("PlanTotal_" & AreasTable(atr, 1)).Index
            Range("ExportTable_" & SheetName).ListObject.Range(exr + 1, indexcol) = 0
        Next exr
        
        atr_days = AreasTable(atr, 5) - AreasTable(atr, 4)
        atr_workingdays = 0
        
        For atr_daysCounter = 0 To (atr_days)
            atr_dc_today = AreasTable(atr, 4) + atr_daysCounter
            
            ' Check if it's a holiday
            atr_dc_isHoliday = False
            For Each d In Range("Holidays_Table").ListObject.ListColumns("Date").Range
                If atr_dc_today = d Then atr_dc_isHoliday = True
            Next d
            
            'Check if it's a workday
            atr_dc_isWorkDay = False
            Select Case Weekday(atr_dc_today)
                Case Is = 1
                    If Works_Sun = True Then atr_dc_isWorkDay = True
                Case Is = 2
                    If Works_Mon = True Then atr_dc_isWorkDay = True
                Case Is = 3
                    If Works_Tue = True Then atr_dc_isWorkDay = True
                Case Is = 4
                    If Works_Wed = True Then atr_dc_isWorkDay = True
                Case Is = 5
                    If Works_Thu = True Then atr_dc_isWorkDay = True
                Case Is = 6
                    If Works_Fri = True Then atr_dc_isWorkDay = True
                Case Is = 7
                    If Works_Sat = True Then atr_dc_isWorkDay = True
            End Select
            
            If atr_dc_isHoliday = False And atr_dc_isWorkDay = True Then atr_workingdays = atr_workingdays + 1
        Next atr_daysCounter
        
        atr_production = AreasTable(atr, 6) / atr_workingdays
        
        For atr_dayupdate = 0 To (atr_days)
            atr_dc_today = AreasTable(atr, 4) + atr_dayupdate
            
            ' Check if it's a holiday
            atr_dc_isHoliday = False
            For Each d In Range("Holidays_Table").ListObject.ListColumns("Date").Range
                If atr_dc_today = d Then atr_dc_isHoliday = True
            Next d
            
            'Check if it's a workday
            atr_dc_isWorkDay = False
            Select Case Weekday(atr_dc_today)
                Case Is = 1
                    If Works_Sun = True Then atr_dc_isWorkDay = True
                Case Is = 2
                    If Works_Mon = True Then atr_dc_isWorkDay = True
                Case Is = 3
                    If Works_Tue = True Then atr_dc_isWorkDay = True
                Case Is = 4
                    If Works_Wed = True Then atr_dc_isWorkDay = True
                Case Is = 5
                    If Works_Thu = True Then atr_dc_isWorkDay = True
                Case Is = 6
                    If Works_Fri = True Then atr_dc_isWorkDay = True
                Case Is = 7
                    If Works_Sat = True Then atr_dc_isWorkDay = True
            End Select
            
            If atr_dc_isHoliday = False And atr_dc_isWorkDay = True Then
                For x = 1 To ExportTable_RowCountAfter + 1
                    If Range("ExportTable_" & SheetName).ListObject.Range(x, 1) = atr_dc_today Then
                        indexcol = Range("ExportTable_" & SheetName).ListObject.ListColumns("PlanTotal_" & AreasTable(atr, 1)).Index
                        Range("ExportTable_" & SheetName).ListObject.Range(x, indexcol) = atr_production
                    End If
                Next x
            Else
                For x = 1 To ExportTable_RowCountAfter + 1
                    If Range("ExportTable_" & SheetName).ListObject.Range(x, 1) = atr_dc_today Then
                        indexcol = Range("ExportTable_" & SheetName).ListObject.ListColumns("PlanTotal_" & AreasTable(atr, 1)).Index
                        Range("ExportTable_" & SheetName).ListObject.Range(x, indexcol) = 0
                    End If
                Next x
            End If
            
        Next atr_dayupdate
    Next atr
    ' copy old production from old ExportTable array NOT WORKING!
    For op_array = 1 To UBound(ExportTable, 1) - LBound(ExportTable, 1) + 1
        For op_table = 1 To ExportTable_RowCountAfter + 1
            If Range("ExportTable_" & SheetName).ListObject.Range(op_table, 1) = ExportTable(op_array, 1) Then
                'Debug.Print Range("ExportTable_" & SheetName).ListObject.Range(op_table, 1)
                For areas_count = 1 To AreasTable_RowCount
                    indexcol2 = Range("ExportTable_" & SheetName).ListObject.ListColumns("CompTotal_" & AreasTable(areas_count, 1)).Index
                    Range("ExportTable_" & SheetName).ListObject.Range(op_table, indexcol2) = ExportTable(op_array, indexcol2)
                Next areas_count
            Else

            End If
        Next op_table
    Next op_array
    
    ' check totals
    ' request if they would like to do an update if there is any production in the completed total
    
    Application.ScreenUpdating = True
End Sub
