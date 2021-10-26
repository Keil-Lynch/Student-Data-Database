# Student-Data-Database
In Excel, I built a database with code to store essential student information.
Blow is the code for the VBA in Excel.

Private Sub cmdAdd_Click()

Dim iRow As Long
Dim ws As Worksheet
Set ws = Worksheets("Student Data")

iRow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, _
    SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1

If Trim(Me.txtGPA.Value) = "" Then
  Me.txtGPA.SetFocus
  MsgBox "Please enter a part number"
  Exit Sub
End If

With ws
  .Cells(iRow, 1).Value = Me.txtStudentName.Value
  .Cells(iRow, 2).Value = Me.txtUniversity.Value
  .Cells(iRow, 3).Value = Me.txtState.Value
  .Cells(iRow, 4).Value = Me.txtMajor.Value
  .Cells(iRow, 5).Value = Me.txtGPA.Value
End With

Me.txtStudentName.Value = ""
Me.txtUniversity.Value = ""
Me.txtState.Value = ""
Me.txtMajor.Value = ""
Me.txtGPA.Value = ""
Me.txtGPA.SetFocus

End Sub
Private Sub cmdClose_Click()
  
  Unload Me

End Sub
