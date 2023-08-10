VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewEntry 
   Caption         =   "New Entry"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   OleObjectBlob   =   "NewEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload NewEntry
End Sub

Private Sub btnSave_Click()
    Dim month, fullDate As String
    Dim lastLine As Long
    Dim i As Integer
    
    Select Case cmbMonth.Value
        Case "January"
            month = "1"
        Case "February"
            month = "2"
        Case "March"
            month = "3"
        Case "April"
            month = "4"
        Case "May"
            month = "5"
        Case "June"
            month = "6"
        Case "July"
            month = "7"
        Case "August"
            month = "8"
        Case "September"
            month = "9"
        Case "October"
            month = "10"
        Case "November"
            month = "11"
        Case Else
            month = "12"
    End Select
    
    fullDate = month & "/" & cmbDay & "/" & cmbYear
    lastLine = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 4 To lastLine
        If IsEmpty(Cells(i, 1).Value) Then
            Cells(i, 1).Select
            Exit For
        End If
    Next
    
    Cells(i, 1).Value = fullDate
    Cells(i, 2).Value = txtClasses.Value
    Cells(i, 3).Value = txtAbs.Value
    Cells(i, 4).Value = txtContent.Value
    Cells(i, 5).Value = txtObs.Value
    
End Sub

Private Sub btnClear_Click()
    cmbDay.ListIndex = 0
    cmbMonth.ListIndex = 0
    cmbYear.ListIndex = 0
    
    txtClasses = ""
    txtAbs = ""
    txtContent = ""
    txtObs = ""
End Sub

Private Sub UserForm_Initialize()

    For i = 1 To 31
        cmbDay.AddItem CStr(i)
    Next
    
    cmbDay.ListIndex = 0
    
    cmbMonth.List = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    cmbMonth.ListIndex = 0
    
    cmbYear.List = Array("2022", "2023", "2024", "2025", "2026", "2027", "2028", "2029", "2030")
    cmbYear.ListIndex = 0
    
End Sub
