VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Employee Pay Slip"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7200
      TabIndex        =   14
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   13
      Top             =   5640
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4080
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   10
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Calculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Net Salary"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Overtime Hours"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Enter Leaves"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Designation Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Pay Slip"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim db As Database
Dim db1 As Database
Private Sub Combo1_Click()
Set rs = db.OpenRecordset("Select * from Table1")
rs.MoveFirst
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
While Not rs.EOF
If Combo1.Text = rs!Name Then
Text1.Text = rs!designation
Text2.Text = rs!salary
End If
rs.MoveNext
Wend
End Sub
Private Sub Command1_Click()
MDIForm1.Enabled = True
Unload Me
End Sub
Private Sub Command2_Click()
rs1.AddNew
rs1!Name = Combo1.Text
rs1!designation = Text1.Text
rs1!salary = Text2.Text
rs1!leaves = Text3.Text
rs1!ot = Text4.Text
rs1!deductions = Text5.Text
rs1!netsalary = Text6.Text
rs1.Update

Beep
Beep
rs2.AddNew
rs2!Name = Combo1.Text
rs2!designation = Text1.Text
rs2!salary = Text2.Text
rs2!leaves = Text3.Text
rs2!ot = Text4.Text
rs2!deductions = Text5.Text
rs2!netsalary = Text6.Text
 rs2.Update
 CrystalReport1.Action = False
End Sub

Private Sub Command3_Click()


Dim a As Integer
Dim b As Integer
Dim ot As Integer
Dim net As Integer
Dim ded As Integer
a = Val(Text3.Text)
b = Val(Text4.Text)
ded = a * 10
Text5.Text = ded
ot = b * 5
Set rs = db.OpenRecordset("Select * from Table1")
rs.MoveFirst
While Not rs.EOF
If Combo1.Text = rs!Name Then
net = Val(rs!salary) + ot - ded
End If
rs.MoveNext
Wend
Text6.Text = net
Command2.Enabled = True
End Sub
Private Sub Form_Load()
Command2.Enabled = False
Command3.Enabled = False
Set db = OpenDatabase(App.Path + "\emp.mdb")
Set rs = db.OpenRecordset("Select name from Table1")
rs.MoveFirst
While Not rs.EOF
Combo1.AddItem rs!Name
rs.MoveNext
Wend
Set db1 = OpenDatabase(App.Path + "\payslip.mdb")
Set rs1 = db1.OpenRecordset("Table1")
Set rs2 = db1.OpenRecordset("Table2")
db1.Execute ("delete * from Table1")
End Sub
Private Sub Text4_Click()
Command3.Enabled = True
End Sub



