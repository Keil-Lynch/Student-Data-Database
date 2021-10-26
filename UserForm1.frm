VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Student Data                                                                 "
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5865
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
s e r s \ k e i l l _ 0 0 0 \ A p p D a t a \ L o c a l \ P r o g r a m s \ M i c r o s o f t   V S   C o d e \ b i n ; C : \ U s e r s \ k e i l l _ 0 0 0 \ A p p D a t a \ L o c a l \ G i t H u b D e s k t o p \ b i n ; C : \ P r o g r a m   F i l e s \ M i c r o s o f t   O f f i c e   1 5 \ r o o t \ C l i e n t   A x U  U[�  ��`q�q G f 8 C C b y v 6 1 X O z r O u Z x V 6 g P 6 8 o B Z x Z 9 n b C P u W a d 5 A p Z 2 p 6 b 4 k u + 2 / I N 6 s M O b n d m r h K k K k 4 7 + t s T k j D P G 7 H o x W B 