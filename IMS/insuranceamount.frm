VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
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
      TabIndex        =   24
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "insurance"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   240
      Picture         =   "insuranceamount.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   23
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETAILS"
      Height          =   5295
      Left            =   6000
      TabIndex        =   9
      Top             =   2280
      Width           =   4575
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1920
         TabIndex        =   15
         Text            =   "Combo4"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text21 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Text            =   "Text21"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text22 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Text            =   "Text22"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Text            =   "Text25"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Text            =   "Text26"
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox Text27 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Text            =   "Text27"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label25 
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
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTAL AMOUNT"
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
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AMOUNT-PAID NOW"
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
         Top             =   3720
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
         Left            =   120
         TabIndex        =   16
         Top             =   4440
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SELECT THE INSURANCE"
      Height          =   3375
      Left            =   1200
      TabIndex        =   0
      Top             =   3360
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Text            =   "Combo2"
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Text            =   "Combo3"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Text            =   "Text18"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text24 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   "Text24"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label28 
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
         Left            =   240
         TabIndex        =   25
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label20 
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
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label21 
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
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTAL-NUMBER OF YEARS"
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
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label28 
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
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"insuranceamount.frx":6F96
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
      TabIndex        =   22
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
Combo3.Clear
Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
 If (Combo2.Text = (Data2.Recordset.Fields(2).Value)) Then
    Combo3.AddItem (Data2.Recordset.Fields(1).Value)
  End If
 Data2.Recordset.MoveNext
Loop
End Sub

Private Sub Combo3_Click()
Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
 If (Combo3.Text = (Data2.Recordset.Fields(1).Value)) Then
    Text18.Text = (Data2.Recordset.Fields(3).Value)
    Text24.Text = (Data2.Recordset.Fields(4).Value)
    Text1.Text = (Data2.Recordset.Fields(5).Value)
  End If
 Data2.Recordset.MoveNext
Loop
If (Text18.Text Mod 10 = 0) Then
For i = 0 To 50 Step 10
Combo4.AddItem (i)
Next i
Else
For i = 0 To 50 Step 5
Combo4.AddItem (i)
Next i
End If
End Sub

Private Sub Combo4_Click()
s = Year(Date) + 1
s3 = Year(Date) + Combo4.Text
s1 = Day(Date)
s2 = Month(Date)
 If (Val(Combo4.Text) < Val(Text18.Text)) Then
  MsgBox "Not-Possible"
 ElseIf Val(Combo4.Text) = Val(Text18.Text) Then
   Text21.Text = Text24.Text
   Text22.Text = Val(Combo4.Text) / Val(Text18.Text)
'   Text23.Text = s1 & "/" & s2 & "/" & s - 1
   Text25.Text = s1 & "/" & s2 & "/" & s3
   Text26.Text = Text24.Text
   Text27.Text = (Val(Text21.Text) * (Val(Text1.Text) / 100))
 Else
     
   Text22.Text = Val(Combo4.Text) / Val(Text18.Text)
   Text21.Text = Val(Text22.Text) * Text24.Text
  ' Text23.Text = s1 & "/" & s2 & "/" & s
   Text25.Text = s1 & "/" & s2 & "/" & s3
   Text26.Text = Text24.Text
   Text27.Text = (Val(Text21.Text) * (Val(Text1.Text) / 100))
 End If
End Sub

Private Sub Command1_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Form_Activate()

Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
Combo2.AddItem (Data2.Recordset.Fields(2).Value)
Data2.Recordset.MoveNext
Loop

End Sub

