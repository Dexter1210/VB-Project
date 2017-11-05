VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5040
   LinkTopic       =   "Form12"
   ScaleHeight     =   4620
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As Recordset
Dim rs1 As Recordset
Dim db As Database
Private Sub Combo1_Click()
Set rs1 = db.OpenRecordset("select * from Table1 ")
List1.Clear
List2.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
rs1.MoveFirst
While Not rs1.EOF
If Combo1.Text = rs1!customername Then
Text1.Text = rs1!customernumber
Text2.Text = rs1!Date
Text3.Text = rs1!Time
List1.AddItem rs1!itemnumber
List2.AddItem rs1!product
List4.AddItem rs1!Weight
List5.AddItem rs1!price
List6.AddItem rs1!quantity
List7.AddItem rs1!totalprice
End If
rs1.MoveNext
Wend
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set db = OpenDatabase(App.Path + "\bill.mdb")
Set rs = db.OpenRecordset("Select distinct customername from Table1")
rs.MoveFirst
While Not rs.EOF
Combo1.AddItem rs!customername
rs.MoveNext
Wend
End Sub

