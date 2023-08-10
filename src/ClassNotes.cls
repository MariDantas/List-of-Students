VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnDeleteEntry_Click()
    Dim dateStr, message As String
    Dim lastLine As Long
    Dim i As Integer
    Dim found As Boolean
    
    found = False
    message = "Which day would you like to delete the entry from? Please provide the complete date"
    
    dateStr = InputBox(message, "Delete entry")
    
    If dateStr <> "" Then
       lastLine = Cells(Rows.Count, 1).End(xlUp).Row
       
       For i = 4 To lastLine
        If StrComp(Cells(i, 1).Value, dateStr, vbBinaryCompare) = 0 Then
            Cells(i, 1).Select
            found = True
            Exit For
        End If
       Next
       
       If found Then
        Cells(i, 1).Value = ""
        Cells(i, 2).Value = ""
        Cells(i, 3).Value = ""
        Cells(i, 4).Value = ""
        Cells(i, 5).Value = ""
        
        MsgBox "Entry successfully deleted", vbOKOnly, "Message"
      Else
        MsgBox "Entry not found", vbOKOnly, "Message"
      End If
    End If
End Sub

Private Sub btnNewEntry_Click()
    NewEntry.Show
End Sub
