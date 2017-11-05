VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5190
   LinkTopic       =   "Form11"
   ScaleHeight     =   5550
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
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
      Left            =   840
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3240
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Enter Selling Price"
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
      TabIndex        =   6
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Stock in Hand"
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
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Dealer Price"
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
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
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
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Products In The Stock"
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
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Update"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim style As VbMsgBoxStyle
Dim result As VbMsgBoxResult
Dim db As Database
Dim db1 As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Private Sub Combo1_Click()
Set rs1 = db.OpenRecordset("select * from Table1 ")
Text3.Text = ""
rs1.MoveFirst
While Not rs1.EOF
If Combo1.Text = rs1!itemname Then
Text3.Text = Val(rs1!quantity) + Val(Text3.Text)
Text7.Text = rs1!itemname
Text1.Text = rs1!dealername
Text2.Text = rs1!price
Text4.Text = rs1!Weight
Text5.Text = ""
Text6.Text = ""
rs2.MoveFirst
While Not rs2.EOF
If Combo1.Text = rs2!product Then
Text5.Text = rs2!code
Text6.Text = rs2!sellingprice
Else
Text5.Text = ""
Text6.Text = ""
Text5.SetFocus
End If
rs2.MoveNext
Wend
End If
rs1.MoveNext
Wend
End Sub

Private Sub Command2_Click()
rs2.AddNew
rs2!product = Text7.Text
rs2!Dealer = Text1.Text
rs2!dealerprice = Text2.Text
rs2!quantity = Text3.Text
rs2!Weight = Text4.Text
rs2!sellingprice = Text6.Text
rs2.Update result = MsgBox("Saved Successfully.", style, "Supermarket Billing 1.0")
Unload Me
Load Form10
Form10.Show
Form10.Move 0, 0
End Sub
Private Sub Command3_Click()
Unload Me
MDIForm1.Enabled = True
End Sub
Private Sub Form_Load()
Command2.Enabled = False
Set db1 = OpenDatabase(App.Path + "\stock.mdb")
Set rs2 = db1.OpenRecordset("Table1")
Set db = OpenDatabase(App.Path + "\save.mdb")
Set rs = db.OpenRecordset("Select distinct itemname from Table1")
rs.MoveFirst
While Not rs.EOF
Combo1.AddItem rs!itemname
rs.MoveNext
Wend
End Sub
Private Sub Text6_GotFocus()
Command2.Enabled = True
End Sub
Private Sub Text7_Change()
End Sub

Private Sub Label5_Click()

End Sub
