VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Activities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnUpdate_Click()
    Dim cell As Range
    Dim obj As OLEObject
    Dim countActiv As Integer
    Dim line As Integer
    Dim what As Boolean
    Dim what2 As Boolean
    
    countActiv = 0
    line = 6
    
    Range("T6").Select
    
    For Each cell In Range("T6:T60")
        If cell.Value > 0 Then
            cell.Value = 0
        End If
        
        Set cell = cell.Offset(1, 0)
    Next cell
    
    For Each cell In Range("J1:S1")
        If (cell.Value <> "Class activity") And (Not (IsEmpty(cell.Value))) Then
            what = cell.Value <> "Class activity" Or IsEmpty(cell.Value)
            what2 = Not IsEmpty(cell.Value)
            countActiv = countActiv + 1
        End If
        
        Set cell = cell.Offset(1, 0)
    Next cell
    
    Range("C2").Select
    ActiveCell.FormulaR1C1 = countActiv
    
    With ActiveSheet
        For Each obj In .OLEObjects
            If TypeName(obj.Object) = "CheckBox" Then
                If obj.Object.Value = True Then
                    Range("T" + CStr(line)).Value = Range("T" + CStr(line)).Value + 1
                End If
                
                If line < 60 Then
                    line = line + 1
                Else
                    line = 6
                End If
            End If
        Next obj
    End With
End Sub
