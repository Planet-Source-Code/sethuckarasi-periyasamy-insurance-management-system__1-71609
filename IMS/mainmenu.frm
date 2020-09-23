VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "INSURANCE-MANAGEMENT SYSTEM"
   ClientHeight    =   9390
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11070
   LinkTopic       =   "MDIForm1"
   Picture         =   "mainmenu.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu inew 
      Caption         =   "NEW"
      Begin VB.Menu icouns 
         Caption         =   "COUNSELLING"
      End
      Begin VB.Menu iregister 
         Caption         =   "REGISTRATION"
      End
      Begin VB.Menu ipolicy 
         Caption         =   "POLICY"
      End
      Begin VB.Menu new 
         Caption         =   "NEW-USER"
      End
   End
   Begin VB.Menu ireport 
      Caption         =   "REPORTS"
      Begin VB.Menu pol 
         Caption         =   "POLICIES"
      End
      Begin VB.Menu xcust 
         Caption         =   "CUSTOMERS"
      End
      Begin VB.Menu icoun 
         Caption         =   "COUNSELLING"
      End
      Begin VB.Menu payment 
         Caption         =   "PAYMENT-RECORDS"
      End
      Begin VB.Menu report 
         Caption         =   "NOMINEES"
      End
   End
   Begin VB.Menu reci 
      Caption         =   "RECIEPTS"
      Begin VB.Menu inst 
         Caption         =   "INSTALLMENTS"
      End
      Begin VB.Menu bond 
         Caption         =   "BOND-DETAILS"
      End
      Begin VB.Menu prem 
         Caption         =   "PREMIUM-CALUCLATION"
      End
   End
   Begin VB.Menu utilit 
      Caption         =   "UTILITIES"
      Begin VB.Menu NOTE 
         Caption         =   "NOTEPAD"
      End
      Begin VB.Menu calcu 
         Caption         =   "CALCULATOR"
      End
   End
   Begin VB.Menu oth 
      Caption         =   "APPLICATION"
      Index           =   1
      Begin VB.Menu app 
         Caption         =   "ABOUT APPLICATION"
      End
      Begin VB.Menu abtinsur 
         Caption         =   "ABOUT INSURANCE"
      End
      Begin VB.Menu abt 
         Caption         =   "ABOUT OTHERS"
      End
      Begin VB.Menu auth 
         Caption         =   "ABOUT AUTHORS"
      End
   End
   Begin VB.Menu oth 
      Caption         =   "OTHERS"
      Index           =   2
      Begin VB.Menu brow 
         Caption         =   "BROWSER"
      End
      Begin VB.Menu exit 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu win 
      Caption         =   "WINDOW"
      Begin VB.Menu casc 
         Caption         =   "CASCADE"
      End
      Begin VB.Menu vert 
         Caption         =   "VERTICAL"
      End
      Begin VB.Menu horiz 
         Caption         =   "HORIZONTAL"
      End
      Begin VB.Menu ar 
         Caption         =   "ARRANGE ICONS"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abt_Click()
frmAbout.Show
End Sub

Private Sub abtinsur_Click()
Form13.Show
End Sub

Private Sub app_Click()
Form14.Show
End Sub

Private Sub ar_Click()
Me.Arrange vbArrangeIcons
End Sub

Private Sub auth_Click()
Form11.Show
End Sub

Private Sub bond_Click()
Form7.Show
End Sub

Private Sub brow_Click()
frmBrowser.Show
End Sub

Private Sub calcu_Click()
'Shell (calculator.exe)

MsgBox "Under construction"
End Sub

Private Sub casc_Click()
Me.Arrange vbCascade
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub horiz_Click()
Me.Arrange vbHorizontal
End Sub

Private Sub icoun_Click()
Form5.Show
End Sub

Private Sub icouns_Click()
Form1.Show
End Sub

Private Sub inst_Click()
Form4.Show
End Sub

Private Sub ipolicy_Click()
Form3.Show
End Sub

Private Sub iregister_Click()
Form2.Show
End Sub

Private Sub new_Click()
Form16.Show
End Sub

Private Sub NOTE_Click()
MsgBox "Under construction"

End Sub

Private Sub payment_Click()
Form8.Show
End Sub

Private Sub pol_Click()
Form12.Show
End Sub

Private Sub prem_Click()
Form9.Show
End Sub

Private Sub report_Click()
Form10.Show
End Sub

Private Sub vert_Click()
Me.Arrange vbVertical
End Sub

Private Sub xcust_Click()
Form6.Show
End Sub
