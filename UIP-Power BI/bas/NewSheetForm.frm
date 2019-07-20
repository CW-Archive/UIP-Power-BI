VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewSheetForm 
   Caption         =   "New Trade"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2520
   OleObjectBlob   =   "NewSheetForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewSheetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub NewTrade_Cancel_Click()
    NewTrade_ID = "Cancel"
    Unload Me
    
End Sub

Private Sub NewTrade_Create_Click()
    NewTrade_ID = NewTrade_IDBoxValue
    Unload Me
    
End Sub
