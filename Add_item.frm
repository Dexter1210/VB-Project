VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   5535
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label9 
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
      Left            =   3000
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Save"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Add New"
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
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Enter Weight"
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
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Code"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Add New Items"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim i As Integer
Private Sub Command1_Click()
rs1!code = Text1.Text
rs1!dealer = Combo1.Text
rs1!product = Text5.Text
rs1!price = Text6.Text
rs1!Weight = Text2.Text
rs1.Update
Command1.Enabled = False
Command3.Enabled = True
Text1.Text = ""
Text5.Text = ""
Text6.Text = ""
Text2.Text = ""
Combo1.Clear
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
i = i + 1
Text1.Text = i
rs1.AddNew
Text1.Enabled = False
Combo1.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text2.Enabled = True
Command1.Enabled = True
End Sub
Private Sub Form_Load()
Command3.Enabled = True
Command1.Enabled = False
Text1.Enabled = False
'saving the current records
Set db1 = OpenDatabase(App.Path + "\deal1.mdb")
Set rs1 = db1.OpenRecordset("Table1")
'calling the dealer name field from other database
Set db = OpenDatabase(App.Path + "\deal.mdb")
Set rs2 = db.OpenRecordset("Select name from Table1")
rs2.MoveFirst
While Not rs2.EOF
Combo1.AddItem rs2!Name
rs2.MoveNext
Wend
rs1.MoveLast
Text1.Text = rs1!code
i = rs1!code
End Sub

