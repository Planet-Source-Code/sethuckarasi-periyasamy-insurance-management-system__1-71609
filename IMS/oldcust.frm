VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PRINT"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "oldpay"
      Top             =   10440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Text            =   " "
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   1575
      Left            =   3240
      TabIndex        =   17
      Text            =   " "
      Top             =   3480
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3240
      TabIndex        =   15
      Text            =   " "
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSURANCE-TYPE"
      Height          =   2655
      Left            =   600
      TabIndex        =   8
      Top             =   5400
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Text            =   " "
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Text            =   " "
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Width           =   2895
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NEXT-DUE"
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
         TabIndex        =   11
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   360
      Picture         =   "oldcust.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame6"
      Height          =   1455
      Left            =   600
      TabIndex        =   2
      Top             =   8040
      Width           =   6855
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   5160
         TabIndex        =   21
         Text            =   " "
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Text            =   " "
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Text            =   " "
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Text            =   " "
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AMOUNT"
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
         Left            =   3840
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BALANCE - INSTALLMENTS"
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
         Height          =   615
         Left            =   3720
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RECIEPT-NUMBER"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PAYMENT MODE"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PAYMENT"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   1095
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
      TabIndex        =   24
      Top             =   2400
      Width           =   2415
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
      TabIndex        =   18
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "                                                                                                RECIEPT-INSTALLMENTS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Data2.Refresh
Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
 If (Data2.Recordset.EOF) Then
   Text16 = Data2.Recordset.Fields(10).Value + 1
 End If
    Text16 = Data2.Recordset.Fields(10).Value + 1

Data2.Recordset.MoveNext
Loop
Data2.Refresh
Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
If Combo1.Text = Data2.Recordset.Fields(8).Value Then
Text3.Text = Data2.Recordset.Fields(0).Value
Text4.Text = Data2.Recordset.Fields(1).Value
Text19.Text = Data2.Recordset.Fields(2).Value
Text1.Text = Data2.Recordset.Fields(3).Value
Text17.Text = Data2.Recordset.Fields(4).Value
Text5.Text = Data2.Recordset.Fields(9).Value
Text2.Text = Data2.Recordset.Fields(6).Value
Text15.Text = Data2.Recordset.Fields(7).Value
GoTo ends
Else
Data2.Recordset.MoveNext
End If
Loop
ends:
If (Text2.Text = 0) Then
 Command2.Enabled = False
 Command3.Enabled = False
End If
Data2.Recordset.Close

End Sub

Private Sub Command1_Click()
MsgBox "Under construction"
'Me.PrintForm
End Sub

Private Sub Command2_Click()
If (Text2.Text > 0) Then
Data2.Refresh
Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
If Combo1.Text = Data2.Recordset.Fields(8).Value Then
 s = Year(Text15.Text) + 1
 s1 = Day(Text15)
 s2 = Month(Text15)
 Text15.Text = s1 & "/" & s2 & "/" & s
 a1 = Data2.Recordset.Fields(5).Value
 a = Data2.Recordset.Fields(9).Value
 Text5.Text = a1 - a
'MsgBox Data2.Recordset.Fields(5).Value - Val(Text5.Text)
 Text2.Text = Text2.Text - 1
 GoTo ends
Else
Data2.Recordset.MoveNext
End If
Loop
ends:
Data2.Recordset.Close
End If
End Sub

Private Sub Command3_Click()
Data2.Refresh
 
 Do Until Data2.Recordset.EOF
 If Combo1.Text = Data2.Recordset.Fields(8).Value Then
 Data2.Recordset.Edit
 Data2.Recordset.Fields(0) = Text3.Text
 Data2.Recordset.Fields(1) = Text4.Text
 Data2.Recordset.Fields(2) = Text19.Text
 Data2.Recordset.Fields(3) = Text1.Text
 Data2.Recordset.Fields(4) = Text17.Text
 Data2.Recordset.Fields(5) = Text5.Text
 Data2.Recordset.Fields(6) = Text2.Text
 Data2.Recordset.Fields(7) = Text15.Text
 Data2.Recordset.Fields(8) = Text16.Text
 Data2.Recordset.Update
 MsgBox "Payment Accepted"
 GoTo ends
 End If
 Data2.Recordset.MoveNext
Loop

ends:
 
 
End Sub

Private Sub Command4_Click()
Text3 = ""
Text4 = ""
Text5 = ""
Text19 = ""
Text2 = ""
Text17 = ""
Text1 = ""
Text15 = ""
Text16 = ""
End Sub

Private Sub Command5_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Activate()
Combo1.Clear
Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
Combo1.AddItem (Data2.Recordset.Fields(8).Value)
Data2.Recordset.MoveNext
Loop
Data2.Refresh
Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
 If (Data2.Recordset.EOF) Then
   Text16 = Data2.Recordset.Fields(10).Value + 1
 End If
    Text16 = Data2.Recordset.Fields(10).Value + 1

Data2.Recordset.MoveNext
Loop
End Sub

