VERSION 5.00
Begin VB.MDIForm MDImain 
   BackColor       =   &H8000000C&
   Caption         =   "Library Management."
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7395
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu munDB 
      Caption         =   "DataBase"
      Begin VB.Menu munSubjects 
         Caption         =   "Subjects"
      End
      Begin VB.Menu munTitles 
         Caption         =   "Titles"
      End
      Begin VB.Menu munBooks 
         Caption         =   "Books"
      End
      Begin VB.Menu sdsd 
         Caption         =   "-"
      End
      Begin VB.Menu munMembers 
         Caption         =   "Members"
      End
      Begin VB.Menu munEmployees 
         Caption         =   "Employees"
      End
      Begin VB.Menu dd 
         Caption         =   "-"
      End
      Begin VB.Menu munOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu sdsdssd 
         Caption         =   "-"
      End
      Begin VB.Menu munExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu munTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu munIssue 
         Caption         =   "Issue"
      End
      Begin VB.Menu munRenewal 
         Caption         =   "Renewal"
      End
      Begin VB.Menu munReturn 
         Caption         =   "Return"
      End
      Begin VB.Menu kkj 
         Caption         =   "-"
      End
      Begin VB.Menu munRes 
         Caption         =   "Reserve Book"
      End
      Begin VB.Menu munMiss 
         Caption         =   "Missing Book"
      End
      Begin VB.Menu oo 
         Caption         =   "-"
      End
      Begin VB.Menu munPayfine 
         Caption         =   "Pay Fine"
      End
   End
   Begin VB.Menu munReports 
      Caption         =   "Reports"
      Begin VB.Menu munFineReport 
         Caption         =   "Fines Paid Report"
      End
      Begin VB.Menu munFineBal 
         Caption         =   "Fines Balance Report"
      End
      Begin VB.Menu munMissBook 
         Caption         =   "Missing Books Report"
      End
   End
   Begin VB.Menu munAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDImain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
M.FileName = App.Path & "\library.mdb"
M.LoadGlobalVariables
'MsgBox M.TotalIssueBook & M.RenewalCounter & M.MaxFineBal
End Sub

Private Sub munAbout_Click()
frmAbout.Show
End Sub

Private Sub munBooks_Click()
frmBooks.Show
End Sub

Private Sub munEmployees_Click()
frmEmployees.Show
End Sub

Private Sub munFineBal_Click()
 Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"
  '''''''''''''''''''''''''''''''''''''''
  
  Set adoprimaryrs = New Recordset
  adoprimaryrs.Open "select MemberId,BooksInHand,FineBal,Tel,Email,Address from Members where FineBal>0", db, adOpenStatic, adLockOptimistic
Set FineBalReport.DataSource = adoprimaryrs
FineBalReport.Show
End Sub

Private Sub munFineReport_Click()
 Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"

  Set adoprimaryrs = New Recordset
  adoprimaryrs.Open "select MemberId,FineAmount,PayDate from Fine", db, adOpenStatic, adLockOptimistic
Set FineReport.DataSource = adoprimaryrs

FineReport.Show

End Sub

Private Sub munIssue_Click()
'''''''
frmIssue.cmdIssue.Visible = True
frmIssue.cmdcharge.Visible = False
frmIssue.cmdrenewal.Visible = False
frmIssue.cmdreturn.Visible = False
'''''''
frmIssue.Caption = "Issue Book"
''''''
munIssue.Enabled = False
munRenewal.Enabled = True
munReturn.Enabled = True
'''''
frmIssue.Show
End Sub

Private Sub munMembers_Click()
frmMembers.Show
End Sub

Private Sub munMiss_Click()
frmIssue.cmdIssue.Visible = False
frmIssue.cmdcharge.Visible = False
frmIssue.cmdrenewal.Visible = False
frmIssue.cmdreturn.Visible = False
''''''''''''''
frmIssue.cmdMiss.Visible = True
'''''''
frmIssue.Caption = "Missing Book"
frmIssue.Show
End Sub

Private Sub munMissBook_Click()
Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"
  '''''''''''''''''''''''''''''''''''''''
  
  Set adoprimaryrs = New Recordset
   sSQL = "select Titles.TitleId,Titles.Subject,Titles.Title,Titles.Author,Books.BookId from Books,Titles where Titles.TitleId=Books.TitleID and books.condition='MISSING'"
  adoprimaryrs.Open sSQL, db, adOpenStatic, adLockOptimistic
Set MissReport.DataSource = adoprimaryrs
MissReport.Show

End Sub

Private Sub munOptions_Click()
frmOptions.Show
munOptions.Enabled = False
End Sub

Private Sub munPayfine_Click()
frmpayfine.Show
End Sub

Private Sub munRenewal_Click()
frmIssue.cmdrenewal.Visible = True
frmIssue.cmdcharge.Visible = False
frmIssue.cmdIssue.Visible = False
frmIssue.cmdreturn.Visible = False
frmIssue.Show
frmIssue.Caption = "Book Renewal"
munRenewal.Enabled = False
munIssue.Enabled = True
munReturn.Enabled = True
End Sub

Private Sub munRes_Click()
frmIssue.cmdIssue.Visible = False
frmIssue.cmdcharge.Visible = False
frmIssue.cmdrenewal.Visible = False
frmIssue.cmdreturn.Visible = False
''''''''''''''
frmIssue.cmdReserve.Visible = True
'''''''
frmIssue.Caption = "Reserve Book"
frmIssue.Show
End Sub

Private Sub munReturn_Click()
frmIssue.cmdreturn.Visible = True
frmIssue.cmdcharge.Visible = False
frmIssue.cmdIssue.Visible = False
frmIssue.cmdrenewal.Visible = False
frmIssue.Show
frmIssue.Caption = "Book Return"
munReturn.Enabled = False
munIssue.Enabled = True
munRenewal.Enabled = True
End Sub

Private Sub munSubjects_Click()
frmSubjects.Show
End Sub

Private Sub munTitles_Click()
frmTitles.Show
End Sub
