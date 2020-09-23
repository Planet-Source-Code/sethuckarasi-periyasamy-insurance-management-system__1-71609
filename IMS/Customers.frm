VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   240
      Picture         =   "Customers.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "register"
      Top             =   4080
      Visible         =   0   'False
      Width           =   6975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Customers.frx":6F96
      Height          =   5775
      Left            =   1800
      OleObjectBlob   =   "Customers.frx":6FAA
      TabIndex        =   0
      Top             =   1920
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"Customers.frx":797D
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
      TabIndex        =   1
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
MDIForm1.Show
End Sub
