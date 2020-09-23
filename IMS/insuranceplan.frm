VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12195
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12195
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Text            =   " "
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "insuranceplan.frx":0000
      Height          =   4335
      Left            =   6720
      OleObjectBlob   =   "insuranceplan.frx":0014
      TabIndex        =   14
      Top             =   2280
      Width           =   5535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "insurance"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Text            =   " "
      Top             =   5040
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Text            =   " "
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Text            =   " "
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Text            =   " "
      Top             =   2040
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   480
      Picture         =   "insuranceplan.frx":09E7
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "% OF INTEREST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   16
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSURANCE-PLAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AMOUNT FOR THE YEARS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NUMBER OF YEARS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSURANCE-TYPE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSURANCE ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"insuranceplan.frx":797D
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = Combo1.Text
Data1.Recordset.Fields(4) = Text4.Text
Data1.Recordset.Fields(5) = Text5.Text
Data1.Recordset.Update
MsgBox ("Policy Updated")
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2 = ""
Text3 = ""
Text4 = ""

End Sub

Private Sub Command3_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Form_Activate()
Data1.Refresh
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
 If (Data1.Recordset.EOF) Then
   Text1 = Data1.Recordset.Fields(0).Value + 1
 End If
    Text1 = Data1.Recordset.Fields(0).Value + 1

Data1.Recordset.MoveNext

Loop
End Sub

Private Sub Form_Load()
 For i = 0 To 50 Step 5
  Combo1.AddItem (i)
 Next i
 Text1.Text = ""
Text2 = ""
Text3 = ""
Text4 = ""

End Sub
