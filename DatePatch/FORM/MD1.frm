VERSION 5.00
Begin VB.MDIForm MD1 
   BackColor       =   &H8000000C&
   Caption         =   "PATCH"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13920
   Icon            =   "MD1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mP 
      Caption         =   "Patch"
   End
End
Attribute VB_Name = "MD1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private con As New Connection
Private rec As New ADODB.Recordset
Private rec1 As New ADODB.Recordset
Private str As String

Private Sub MDIForm_Load()

End Sub

Private Sub mP_Click()
str = InputBox("Please enter the Database Name")
con.Open str, "", "111"
With rec
If .State = 1 Then .Close
.Open "select * from Supervision", con, adOpenKeyset, adLockOptimistic
If .RecordCount Then
    .MoveFirst
    !CD = "2018/10/25"
    !CPW = "19800103"
    .Update
End If
End With
If con.State = 1 Then con.Close
MsgBox "Patch Updated"
End
End Sub
