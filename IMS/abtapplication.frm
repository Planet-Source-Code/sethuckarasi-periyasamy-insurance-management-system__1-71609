VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form14"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form14"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   10245
   Begin VB.TextBox Text1 
      Height          =   6375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "abtapplication.frx":0000
      Top             =   1680
      Width           =   10215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   8160
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   480
      Picture         =   "abtapplication.frx":0686
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"abtapplication.frx":761C
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
      TabIndex        =   0
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
MDIForm1.Show
End Sub
