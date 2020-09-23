VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM(COUNSELLING DETAILS)"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   240
      Picture         =   "counselling.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "OPTIONS"
      Height          =   2655
      Left            =   5760
      TabIndex        =   11
      Top             =   2760
      Width           =   3375
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\IMS\insurance.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "insur"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         Caption         =   "BACK TO MAIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "VIEW THE CUSTOMERS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "COUNSELL-UPDATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   " "
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Text            =   " "
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Text            =   " "
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Text            =   " "
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   2880
      TabIndex        =   6
      Text            =   " "
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Text            =   " "
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"counselling.frx":6F96
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
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   13695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
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
      Index           =   0
      Left            =   600
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DETAILS"
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
      Left            =   600
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE/CONTACT NUMBER"
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
      Left            =   600
      TabIndex        =   3
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF ENQUIRY"
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
      Left            =   600
      TabIndex        =   2
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "OCCUPATION"
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
      Left            =   600
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
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
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = Text4.Text
Data1.Recordset.Fields(4) = Text5.Text
Data1.Recordset.Fields(5) = Text6.Text
Data1.Recordset.Update
MsgBox "Customers-Counselled"
 
End Sub

Private Sub Command2_Click()
Me.Hide
Form5.Show
End Sub

Private Sub Command3_Click()
Me.Hide
MDIForm1.Show

End Sub

Private Sub Form_Load()
Text4.Text = Now
End Sub
