VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bond"
      Top             =   7680
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RECIEPT"
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
      Height          =   4575
      Left            =   7680
      TabIndex        =   14
      Top             =   2880
      Width           =   4455
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Text            =   " "
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Text            =   " "
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Text            =   " "
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Text            =   " "
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RECIEPT NUMBER"
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
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTAL AMOUNT TO BE PAID"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NUMBER OF INSTALLMENTS"
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
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MATURITY-DATE"
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
         Left            =   240
         TabIndex        =   16
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MATURITY-AMT"
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
         Left            =   240
         TabIndex        =   15
         Top             =   3960
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   3240
      TabIndex        =   11
      Text            =   " "
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Text            =   " "
      Top             =   3120
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   600
      Picture         =   "reciept-bond.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   240
      Width           =   2055
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BILL-PRINT"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cust-insur"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame3 
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
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      Top             =   5400
      Width           =   6135
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Text            =   " "
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "INSURANCE NAME"
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
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PLAN TYPE"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"reciept-bond.frx":6F96
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
      TabIndex        =   9
      Top             =   0
      Width           =   13695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAME & ADDRESS"
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
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUSTOMER-ID"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
 If Combo1.Text = Data1.Recordset.Fields(1) Then
    Text1.Text = Data1.Recordset.Fields(0)
    Text2.Text = Data1.Recordset.Fields(4)
    Text3.Text = Data1.Recordset.Fields(2)
    Text4.Text = Data1.Recordset.Fields(3)
    Text6.Text = Data1.Recordset.Fields(11)
    Text9.Text = Data1.Recordset.Fields(6)
    Text10.Text = Data1.Recordset.Fields(7)
    Text12.Text = Data1.Recordset.Fields(8)
    Text13.Text = Data1.Recordset.Fields(9)
      GoTo ends
  Else
   If Data1.Recordset.EOF Then
   MsgBox "No Records"
   End If
  End If
 Data1.Recordset.MoveNext
Loop
ends:

End Sub

Private Sub Command1_Click()
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = Text1.Text
Data2.Recordset.Fields(1) = Text2.Text
Data2.Recordset.Fields(2) = Text3.Text
Data2.Recordset.Fields(3) = Text4.Text
Data2.Recordset.Fields(4) = Text6.Text
Data2.Recordset.Fields(5) = Text9.Text
Data2.Recordset.Fields(6) = Text10.Text
Data2.Recordset.Fields(7) = Text12.Text
Data2.Recordset.Fields(8) = Text13.Text
Data2.Recordset.Update
End Sub

Private Sub Command2_Click()
MsgBox "UNDER CONSTRUCTION"

End Sub

Private Sub Command3_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Form_Activate()
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
Combo1.AddItem (Data1.Recordset.Fields(1).Value)
Data1.Recordset.MoveNext
Loop
End Sub

