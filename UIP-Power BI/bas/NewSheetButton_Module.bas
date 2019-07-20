Attribute VB_Name = "NewSheetButton_Module"
Option Explicit
Public NewTrade_ID As String

Sub NewSheetButton()
    ' Camron Walker 7/20/2019
    ' https://github.com/CamronWalker/UIP-Power-BI
start:
    Dim regCheck As Boolean
    Dim copiedSheet As Worksheet
    Dim tblCount As Long
    
    'Get NewTrade_ID from userform
    NewSheetForm.Show
    Debug.Print NewTrade_ID
    If NewTrade_ID = "Cancel" Then
        NewTrade_ID = ""
        Exit Sub
    End If
    regCheck = strOut(NewTrade_ID)
    If regCheck = False Then
        MsgBox ("Was that a 4 digit number? Try again.")
        NewTrade_ID = ""
        GoTo start
    End If
    
    If SheetExists(NewTrade_ID) Then
        MsgBox ("Error: Sheet (" & NewTrade_ID & ") already exists.  Please select a different Trade ID to create a sheet out of.")
        NewTrade_ID = ""
        GoTo start
    End If
    
    'Copy and Rename Sheet
    Sheets("Template").Copy After:=Sheets(Sheets.Count)
    Set copiedSheet = ActiveSheet
    copiedSheet.Name = NewTrade_ID
    
    'Add new sheet to Index
    Range("TradesTable").ListObject.ListRows.Add
    Range("TradesTable").Cells(Range("TradesTable").ListObject.ListRows.Count, 1) = NewTrade_ID
    
    'Table name changes
    For tblCount = 1 To ActiveSheet.ListObjects.Count
        Select Case Left(ActiveSheet.ListObjects(tblCount).Name, 19)
            Case Is = "AreasTable_Template"
                ActiveSheet.ListObjects(tblCount).Name = "AreasTable_" & NewTrade_ID
            Case Is = "ExportTable_Templat"
                ActiveSheet.ListObjects(tblCount).Name = "ExportTable_" & NewTrade_ID
        End Select
    Next tblCount
       
    NewTrade_ID = ""
End Sub
Function strOut(strIn As String) As String
    Dim objRegex As Object
    
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Pattern = "^[0-9]{4}$"
        strOut = .Test(strIn)
    End With
End Function
Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function

