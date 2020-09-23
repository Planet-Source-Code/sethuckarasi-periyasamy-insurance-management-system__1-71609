VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INSURANCE MANAGEMENT SYSTEM"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14190
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   14190
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000A&
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "oldpay"
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data5 
      Caption         =   "cust-pol"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cust-insur"
      Top             =   9360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data4 
      Caption         =   "NEWPOLCUST"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "newpolcust"
      Top             =   9360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "BACK-TO-MAIN"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\IMS\insurance.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "insurance"
      Top             =   9360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RECIEPT"
      Height          =   5535
      Left            =   8880
      TabIndex        =   44
      Top             =   5160
      Width           =   5055
      Begin VB.TextBox Text27 
         Height          =   375
         Left            =   4080
         TabIndex        =   69
         Text            =   " "
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   1920
         TabIndex        =   67
         Text            =   " "
         Top             =   4920
         Width           =   1455
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   1800
         TabIndex        =   65
         Text            =   " "
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   1800
         TabIndex        =   62
         Text            =   " "
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text22 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   59
         Text            =   " "
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text21 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   58
         Text            =   " "
         Top             =   2520
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1800
         TabIndex        =   57
         Text            =   "Combo4"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   1800
         TabIndex        =   54
         Text            =   " "
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text19 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   53
         Text            =   " "
         Top             =   360
         Width           =   1575
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
         Left            =   3600
         TabIndex        =   70
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   "+"
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
         Left            =   3600
         TabIndex        =   68
         Top             =   2520
         Width           =   255
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
         TabIndex        =   66
         Top             =   4920
         Width           =   1455
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
         TabIndex        =   64
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label29 
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
         Left            =   240
         TabIndex        =   63
         Top             =   3600
         Width           =   1335
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
         TabIndex        =   49
         Top             =   3000
         Width           =   1335
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
         TabIndex        =   48
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   47
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PAYMENT-MODE"
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
         TabIndex        =   46
         Top             =   1200
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
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SELECT THE INSURANCE"
      Height          =   3255
      Left            =   8760
      TabIndex        =   40
      Top             =   1800
      Width           =   4095
      Begin VB.TextBox Text28 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   73
         Text            =   " "
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text24 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   56
         Text            =   " "
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   1920
         TabIndex        =   52
         Text            =   " "
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1920
         TabIndex        =   51
         Text            =   "Combo3"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         TabIndex        =   50
         Text            =   "Combo2"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label35 
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
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   2880
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
         Left            =   240
         TabIndex        =   55
         Top             =   2400
         Width           =   1215
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
         TabIndex        =   43
         Top             =   1800
         Width           =   1455
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
         TabIndex        =   42
         Top             =   1200
         Width           =   1575
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
         TabIndex        =   41
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUSTOMERS"
      Height          =   7335
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   7935
      Begin VB.Data Data3 
         Caption         =   "NOMINEE"
         Connect         =   "Access"
         DatabaseName    =   "D:\IMS\insurance.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "nominnee"
         Top             =   6720
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data Data1 
         Caption         =   "REGISTER"
         Connect         =   "Access"
         DatabaseName    =   "D:\IMS\insurance.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "register"
         Top             =   6840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NOMINEE"
         Height          =   2415
         Left            =   3840
         TabIndex        =   33
         Top             =   3600
         Width           =   3735
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   2040
            TabIndex        =   39
            Text            =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2040
            TabIndex        =   38
            Text            =   "Combo1"
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox Text16 
            Height          =   375
            Left            =   2040
            TabIndex        =   35
            Text            =   " "
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NOMINEE-RELATION"
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
            TabIndex        =   37
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NOMINEE-GENDER"
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
            TabIndex        =   36
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NOMINEE-NAME"
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
            TabIndex        =   34
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox Text15 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   32
         Text            =   " "
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   6360
         TabIndex        =   31
         Text            =   " "
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   6360
         TabIndex        =   30
         Text            =   " "
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   6960
         TabIndex        =   29
         Text            =   " "
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   5280
         TabIndex        =   28
         Text            =   " "
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2160
         TabIndex        =   27
         Text            =   " "
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2160
         TabIndex        =   26
         Text            =   " "
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   2160
         TabIndex        =   25
         Text            =   " "
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Text            =   " "
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         Text            =   " "
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2160
         TabIndex        =   22
         Text            =   " "
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Text            =   " "
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Text            =   " "
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ANN-INCOME"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IDENTIFICATION MARKS"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WEIGHT"
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
         Left            =   6120
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HEIGHT"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BANK-ADDRESS"
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
         TabIndex        =   12
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BANK A/C NUMBER"
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
         TabIndex        =   11
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NATIONALITY"
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
         TabIndex        =   10
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MOBILE-NUMBER"
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
         TabIndex        =   9
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MARITAL STATUS"
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
         TabIndex        =   6
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DATE-OF-BIRTH (DD/MM/YYYY)"
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
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FATHER'S NAME"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      Picture         =   "customer1.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   $"customer1.frx":6F96
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
Attribute VB_Name = "Form2"
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
    Text28.Text = (Data2.Recordset.Fields(5).Value)
  End If
 Data2.Recordset.MoveNext
Loop
End Sub


Private Sub Combo4_Click()
s = Year(Date) + 1
s3 = Year(Date) + Val(Combo4.Text)
s1 = Day(Date)
s2 = Month(Date)
 If (Val(Combo4.Text) < Val(Text18.Text)) Then
  MsgBox "Not-Possible"
 ElseIf Val(Combo4.Text) = Val(Text18.Text) Then
   Text21.Text = Text24.Text
   Text22.Text = Val(Combo4.Text) / Val(Text18.Text)
   Text23.Text = s1 & "/" & s2 & "/" & s - 1
   Text25.Text = s1 & "/" & s2 & "/" & s3
   Text26.Text = Text24.Text
   Text27.Text = (Val(Text21.Text) * (Val(Text28.Text) / 100))
 Else
   Text22.Text = Val(Combo4.Text) / Val(Text18.Text)
   Text21.Text = Val(Text22.Text) * Text24.Text
   Text23.Text = s1 & "/" & s2 & "/" & s
   Text25.Text = s1 & "/" & s2 & "/" & s3
   Text26.Text = Text24.Text
   Text27.Text = (Val(Text21.Text) * (Val(Text28.Text) / 100))
 End If
End Sub

Private Sub Command1_Click()
 Data1.Recordset.AddNew
 Data1.Recordset.Fields(0) = Text1.Text
 Data1.Recordset.Fields(1) = Text2.Text
 Data1.Recordset.Fields(2) = Text3.Text
 Data1.Recordset.Fields(3) = Text4.Text
 Data1.Recordset.Fields(4) = Text5.Text
 Data1.Recordset.Fields(5) = Text6.Text
 Data1.Recordset.Fields(6) = Text7.Text
 Data1.Recordset.Fields(7) = Text8.Text
 Data1.Recordset.Fields(8) = Text9.Text
 Data1.Recordset.Fields(9) = Text10.Text
 Data1.Recordset.Fields(10) = Text11.Text
 Data1.Recordset.Fields(11) = Text12.Text
 Data1.Recordset.Fields(12) = Text13.Text
 Data1.Recordset.Fields(13) = Text14.Text
 Data1.Recordset.Fields(14) = Text15.Text
 Data1.Recordset.Update
 'data2=register
 Data3.Recordset.AddNew
 Data3.Recordset.Fields(0) = Text16.Text
 Data3.Recordset.Fields(1) = Text17.Text
 Data3.Recordset.Fields(2) = Combo1.Text
 Data3.Recordset.Fields(3) = Text1.Text
 Data3.Recordset.Fields(4) = Text15.Text
 Data3.Recordset.Update
 'data4-register
 Data4.Recordset.AddNew
 Data4.Recordset.Fields(0) = Text1.Text
 Data4.Recordset.Fields(1) = Text2.Text
 Data4.Recordset.Fields(2) = Text6.Text
 Data4.Recordset.Fields(3) = Text16.Text
 Data4.Recordset.Fields(4) = Text20.Text
 Data4.Recordset.Fields(5) = Text19.Text
 Data4.Recordset.Fields(6) = Combo2.Text
 Data4.Recordset.Fields(7) = Combo3.Text
 Data4.Recordset.Fields(8) = Combo4.Text
 Data4.Recordset.Fields(9) = Text23.Text
 Data4.Recordset.Fields(10) = Text15.Text
 Data4.Recordset.Update
 'data5
 Data5.Recordset.AddNew
 Data5.Recordset.Fields(0) = Text1.Text
 Data5.Recordset.Fields(1) = Text15.Text
 Data5.Recordset.Fields(2) = Combo2.Text
 Data5.Recordset.Fields(3) = Combo3.Text
 Data5.Recordset.Fields(4) = Text6.Text
 Data5.Recordset.Fields(5) = Text23.Text
 Data5.Recordset.Fields(6) = Text21.Text
 Data5.Recordset.Fields(7) = Text22.Text
  Data5.Recordset.Fields(8) = Text25.Text
 Data5.Recordset.Fields(9) = Text27.Text
 Data5.Recordset.Fields(10) = Text26.Text
 Data5.Recordset.Fields(11) = Text19.Text
 Data5.Recordset.Fields(12) = Val(Text21.Text) + Val(Text27.Text)
  Data5.Recordset.Update
  
 Data6.Recordset.AddNew
 Data6.Recordset.Fields(0) = Text1.Text
 Data6.Recordset.Fields(1) = Text6.Text
 Data6.Recordset.Fields(2) = Combo2.Text
 Data6.Recordset.Fields(3) = Combo3.Text
 Data6.Recordset.Fields(4) = Text20.Text
 Data6.Recordset.Fields(5) = Val(Text21.Text) - Val(Text24.Text)
 Data6.Recordset.Fields(6) = Val(Text22.Text) - 1
 Data6.Recordset.Fields(7) = Text23.Text
 Data6.Recordset.Fields(8) = Text15.Text
 Data6.Recordset.Fields(9) = Text26.Text
 Data6.Recordset.Fields(10) = Text19.Text
 Data6.Recordset.Update
 MsgBox "Record-Registered"
End Sub

Private Sub Command2_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text25.Text = ""
Text26.Text = ""
Text27.Text = ""

End Sub

Private Sub Form_Activate()
Data2.Recordset.MoveFirst
Do Until Data2.Recordset.EOF
Combo2.AddItem (Data2.Recordset.Fields(2).Value)
Data2.Recordset.MoveNext
Loop
Combo1.AddItem ("Father")
Combo1.AddItem ("Mother")
Combo1.AddItem ("Brother")
Combo1.AddItem ("Sister")
Combo1.AddItem ("Daughter")
Combo1.AddItem ("Son")
Combo1.AddItem ("Grand-Son")
Combo1.AddItem ("Grand-Daughter")
Combo1.AddItem ("Husband")
Combo1.AddItem ("Wife")
Combo1.AddItem ("Friend")

Data1.Refresh
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
 If (Data1.Recordset.EOF) Then
   Text15 = Data1.Recordset.Fields(14).Value + 1
 End If
    Text15 = Data1.Recordset.Fields(14).Value + 1

Data1.Recordset.MoveNext
Loop
Data4.Refresh
Data4.Recordset.MoveFirst
Do Until Data4.Recordset.EOF
 If (Data4.Recordset.EOF) Then
   Text19 = Data4.Recordset.Fields(5).Value + 1
 End If
    Text19 = Data4.Recordset.Fields(5).Value + 1

Data4.Recordset.MoveNext
Loop
End Sub

Private Sub Text18_Change()
Combo4.Clear
If (Val(Text18.Text) Mod 10 = 0) Then
 For i = 0 To 50 Step 10
 Combo4.AddItem (i)
 Next i
Else
 For i = 0 To 50 Step 5
 Combo4.AddItem (i)
 Next i
End If
End Sub
