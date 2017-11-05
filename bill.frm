VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7410
   LinkTopic       =   "Form13"
   ScaleHeight     =   5670
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label11 
      Caption         =   "Total Price"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label9 
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
      Left            =   3960
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label8 
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
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Select The Product"
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
      Left            =   3840
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Product Number"
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
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Number"
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
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Name"
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
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Create New Bill"
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
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim rs4 As Recordset
Dim rs5 As Recordset
Dim rs6 As Recordset
Dim db As Database
Dim db1 As Database
Dim db2 As Database
Dim i As Integer
Dim j As Integer
Private Sub Combo1_Click()
Set rs1 = db.OpenRecordset("Select * from Table1")
rs1.MoveFirst
Text4.Text = ""
Text5.Text = ""
While Not rs1.EOF
If Combo1.Text = rs1!product Then
Text3.Text = rs1!sellingprice
Text2.Text = rs1!Weight
End If
rs1.MoveNext
Wend
End Sub
Private Sub Command1_Click()
On Error Resume Next
rs2.AddNew rs2!customernumber = Text7.Text
rs2!customername = Text10.Text
rs2!itemnumber = Text6.Text
rs2!product = Combo1.Text
rs2!code = Text1.Text
rs2!Weight = Text2.Text
rs2!price = Text3.Text
rs2!quantity = Text4.Text
rs2!totalprice = Text5.Text
rs2!Date = Text8.Text
rs2!Time = Text9.Text
rs2.Update
Beep Beep
rs4.AddNew rs4!customernumber = Text7.Text
rs4!customername = Text10.Text
rs4!itemnumber = Text6.Text
rs4!product = Combo1.Text
rs4!code = Text1.Text
rs4!Weight = Text2.Text
rs4!price = Text3.Text
rs4!quantity = Text4.Text
rs4!totalprice = Text5.Text
rs4!Date = Text8.Text
rs4!Time = Text9.Text
rs4.Update
List1.AddItem Text6.Text
List2.AddItem Combo1.Text
List3.AddItem Text1.Text
List4.AddItem Text2.Text
List5.AddItem Text3.Text
List6.AddItem Text4.Text
List7.AddItem Text5.Text

b = 0
For a = 0 To List7.ListCount - 1
b = b + Val(List7.List(a))
Label19.Caption = b
Next a
i = i + 1
Text6.Text = i
End Sub
Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Command4_Click()
On Error Resume Next
List1.RemoveItem
List1.ListCount -1
List2.RemoveItem
List2.ListCount -1
List3.RemoveItem
List3.ListCount -1
List4.RemoveItem
List4.ListCount -1
List5.RemoveItem
List5.ListCount -1
List6.RemoveItem
List6.ListCount -1
List7.RemoveItem
List7.ListCount -1
Label19.Caption = ""
i = i - 1
Text6.Text = i
End Sub
Private Sub Command5_Click()
db2.Execute ("delete * from Table1")
End Sub
Private Sub Form_Load()
Text8.Text = Date
Set db = OpenDatabase(App.Path + "\stock.mdb")
Set rs = db.OpenRecordset("Select product from Table1")
rs.MoveFirst
While Not rs.EOF
Combo1.AddItem rs!product
rs.MoveNext
Wend
Set db1 = OpenDatabase(App.Path + "\bill.mdb")
Set rs2 = db1.OpenRecordset("Table1")
Set db2 = OpenDatabase(App.Path + "\temp.mdb")
Set rs4 = db2.OpenRecordset("Table1")
db2.Execute ("delete * from Table1")
i = 1
Text6.Text = i
j = 0
Text7.Text = j
End Sub
Private Sub Label5_Click()
End Sub
Private Sub Text4_Change()
Text5.Text = Val(Text3.Text) * Val(Text4.Text)
End Sub
Private Sub Timer1_Timer()
Text9.Text = Time
End Sub


Private Sub Label4_Click()

End Sub
