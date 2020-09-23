VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Index           =   1
      Left            =   4800
      Picture         =   "author.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   22
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3240
      TabIndex        =   21
      Top             =   7800
      Width           =   6855
      Begin VB.Image Image9 
         Height          =   600
         Left            =   5640
         Picture         =   "author.frx":6F96
         Top             =   240
         Width           =   690
      End
      Begin VB.Image Image8 
         Height          =   600
         Left            =   4920
         Picture         =   "author.frx":90B4
         Top             =   240
         Width           =   570
      End
      Begin VB.Image Image7 
         Height          =   600
         Left            =   3000
         Picture         =   "author.frx":AE75
         Top             =   240
         Width           =   600
      End
      Begin VB.Image Image6 
         Height          =   600
         Left            =   3720
         Picture         =   "author.frx":CFC9
         Top             =   240
         Width           =   765
      End
      Begin VB.Image Image5 
         Height          =   600
         Left            =   2400
         Picture         =   "author.frx":FD79
         Top             =   240
         Width           =   570
      End
      Begin VB.Image Image4 
         Height          =   600
         Left            =   1680
         Picture         =   "author.frx":11B3A
         Top             =   240
         Width           =   555
      End
      Begin VB.Image Image3 
         Height          =   600
         Left            =   840
         Picture         =   "author.frx":13D0F
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   480
         Picture         =   "author.frx":14DD0
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   120
         Picture         =   "author.frx":15E91
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   9000
      Picture         =   "author.frx":16F52
      ScaleHeight     =   1335
      ScaleWidth      =   975
      TabIndex        =   20
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "BACK"
      Height          =   495
      Left            =   12120
      TabIndex        =   19
      Top             =   9480
      Width           =   1095
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00C0C000&
      Height          =   1215
      Left            =   12720
      Picture         =   "author.frx":1A898
      ScaleHeight     =   1155
      ScaleWidth      =   1035
      TabIndex        =   18
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   9960
      Picture         =   "author.frx":1DEB5
      ScaleHeight     =   1335
      ScaleWidth      =   975
      TabIndex        =   17
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   360
      Picture         =   "author.frx":21773
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   4
      Left            =   7680
      Picture         =   "author.frx":24FFF
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   15
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   3
      Left            =   6720
      Picture         =   "author.frx":280DC
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   14
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   2
      Left            =   5760
      Picture         =   "author.frx":2BBCF
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   13
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   1
      Left            =   4800
      Picture         =   "author.frx":2ECAC
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   12
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   3840
      Picture         =   "author.frx":325EE
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   11
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2880
      Picture         =   "author.frx":35E93
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   7440
      Picture         =   "author.frx":3983B
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   8520
      Picture         =   "author.frx":3D1E3
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   9600
      Picture         =   "author.frx":40A40
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   10680
      Picture         =   "author.frx":44376
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   6480
      Picture         =   "author.frx":47BD3
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   5160
      Picture         =   "author.frx":4B1F0
      ScaleHeight     =   1215
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4200
      Picture         =   "author.frx":4EA4D
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3240
      Picture         =   "author.frx":51F09
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2280
      Picture         =   "author.frx":54FE6
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1320
      Picture         =   "author.frx":5898E
      ScaleHeight     =   1215
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
MDIForm1.Show
End Sub

