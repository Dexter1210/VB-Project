VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form6"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6180
   LinkTopic       =   "Form6"
   ScaleHeight     =   5985
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List4 
      Height          =   450
      Left            =   2520
      TabIndex        =   16
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   2520
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   2520
      TabIndex        =   14
      Top             =   1680
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   2520
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Time"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Date"
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
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Total Amount"
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
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity"
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
      Left            =   360
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Weight"
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
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Price"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Product Name"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Dealer Name"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Purchase Item"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim style As VbMsgBoxStyle
Dim result As VbMsgBoxResult
Dim db As Database
Dim rs As Recordset
Dim db1 As Database
Dim db4 As Database
Dim db2 As Database
Dim rs2 As Recordset
Dim rs1 As Recordset
Dim rs3 As Recordset
Dim rs4 As Recordset
Dim rs6 As Recordset
Private Sub Combo1_Click()
Set rs3 = db1.OpenRecordset("Table1")
rs3.MoveFirst
While Not rs3.EOF
If Combo1.Text = rs3!dealer Then
Combo2.AddItem rs3!product
Combo3.AddItem rs3!price
Combo4.AddItem rs3!Weight
End If
rs3.MoveNext
Wend
End Sub
Private Sub Combo2_Click()
Set rs3 = db1.OpenRecordset("Table1")
rs3.MoveFirst
While Not rs3.EOF
If Combo2.Text = rs3!product Then
Combo3.AddItem rs3!price
Combo4.AddItem rs3!Weight
End If
rs3.MoveNext
Wend
End Sub
Private Sub Command1_Click()
On Error Resume Next
rs.AddNew
rs!dealername = Combo1.Text
rs!itemname = Combo2.Text
rs!price = Combo3.Text
rs!quantity = Text3.Text
rs!amount = Text2.Text
rs!date1 = Text1.Text
rs!time1 = Text4.Text
rs!Weight = Combo4.Text
rs.Update result = MsgBox("Saved Successfully.", style, "Supermarket Billing 1.0")
Unload Me
Load Form7
Form7.Show
Form7.Move 0, 0
End Sub

Private Sub Command2_Click()
Unload Me
Load Form7
Form7.Show
Form7.Move 0, 0
End Sub
Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Command5_Click()

End Sub
Private Sub Command6_Click()
rs.Delete
End Sub
Private Sub Form_Load()
Command1.Enabled = False
Text1.Text = Date
Set db1 = OpenDatabase(App.Path + "\deal1.mdb")
Set rs1 = db1.OpenRecordset("Table1")
Set rs2 = db1.OpenRecordset("Select distinct dealer from Table1 ")
Set rs4 = db1.OpenRecordset("Table1")
Set db4 = apppath + OpenDatabase("c:\employee\transaction.mdb")
Set db2 = OpenDatabase(App.Path + "\save.mdb")
Set rs = db2.OpenRecordset("Table1")
rs2.MoveFirst
While Not rs2.EOF
Combo1.AddItem rs2!dealer
rs2.MoveNext
Wend
End Sub
Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()
Command1.Enabled = True
End Sub
Private Sub Text3_Change()
Text2.Text = Val(Combo3.Text) * Val(Text3.Text)
End Sub
Private Sub Timer1_Timer()
Text4.Text = Time
End Sub


