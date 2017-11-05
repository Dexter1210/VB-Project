VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form11"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   ScaleHeight     =   4980
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
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
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Balance Stock"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity Sold Out"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Quantity in Hand"
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
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Choose The Product"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "View Sold Stock"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim db1 As Database
Dim db2 As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim rs4 As Recordset
Dim style As VbMsgBoxStyle
Dim result As VbMsgBoxResult
Private Sub Combo1_Click()
Set rs1 = db.OpenRecordset("Select * from Table1 ")
Text2.Text = ""
rs1.MoveFirst
While Not rs1.EOF
If Combo1.Text = rs1!product Then
Text2.Text = Val(rs1!quantity) + Val(Text2.Text)
End If
rs1.MoveNext
Wend
Set rs2 = db1.OpenRecordset("Table1")
Set rs2 = db1.OpenRecordset("Select * from Table1 ")
Text1.Text = ""
rs2.MoveFirst
While Not rs2.EOF
If Combo1.Text = rs2!itemname Then
Text1.Text = Val(rs2!quantity) + Val(Text1.Text)
End If
rs2.MoveNext
Wend
Text3.Text = Val(Text1.Text) - Val(Text2.Text)
If Val(Text3.Text) <= 4 Then
result = MsgBox("WARNING STOCK LOW !!!.", style, "SupermarketBilling 1.0")

End If
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()
CrystalReport1.Action = False
End Sub
Private Sub Form_Load()
Set db = OpenDatabase(App.Path + "\bill.mdb")
Set rs = db.OpenRecordset("Select distinct product from Table1 ")
rs.MoveFirst
While Not rs.EOF
Combo1.AddItem rs!product
rs.MoveNext
Wend
Set db1 = OpenDatabase(App.Path + "\save.mdb")
End Sub

