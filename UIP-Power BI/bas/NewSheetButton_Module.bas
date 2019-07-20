Attribute VB_Name = "NewSheetButton_Module"
Public NewTrade_ID As String

Sub NewSheetButton()
    ' Camron Walker 7/20/2019
    ' https://github.com/CamronWalker/UIP-Power-BI
start:
    Dim regCheck As Boolean
    
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
