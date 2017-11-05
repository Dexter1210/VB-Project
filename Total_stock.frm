VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form5"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5535
   LinkTopic       =   "Form5"
   ScaleHeight     =   3945
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Total Quantity"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Total Stock"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Private Sub Combo1_Click()
Set rs1 = db.OpenRecordset("select * from Table1 ")
Text1.Text = ""
rs1.MoveFirst
While Not rs1.EOF
If Combo1.Text = rs1!itemname Then
Text1.Text = Val(rs1!quantity) + Val(Text1.Text)
End If
rs1.MoveNext
Wend
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set db = OpenDatabase(App.Path + "\save.mdb")
Set rs = db.OpenRecordset("Select distinct itemname from Table1 ")
rs.MoveFirst
While Not rs.EOF
Combo1.AddItem rs!itemname
rs.MoveNext
Wend
End Sub
