VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM"
   ClientHeight    =   4380
   ClientLeft      =   9075
   ClientTop       =   6735
   ClientWidth     =   5925
   LinkTopic       =   "Form15"
   ScaleHeight     =   4380
   ScaleWidth      =   5925
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "user"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EXIT"
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LOG-IN"
      Height          =   255
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   1080
      Picture         =   "user.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   " "
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      TabIndex        =   1
      Text            =   " "
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                    LOG-IN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PASSOWORD"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "USER-ID"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
 If (Text1.Text = (Data1.Recordset.Fields(0).Value)) Then
    If (Text2.Text = (Data1.Recordset.Fields(1).Value)) Then
      Me.Hide
      MDIForm1.Show
     Else
       MsgBox "Check your password"
     End If
  Else
    MsgBox "Check the user id"
  End If
 Data1.Recordset.MoveNext
 If ((Text1.Text = admin) And (Text1.Text = admin)) Then
  Me.Hide
  MDIForm1.Show
 End If
Loop
End Sub

Private Sub Form_Load()
Text1 = ""
Text2 = ""
'Text1.SetFocus
End Sub
