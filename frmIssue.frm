VERSION 5.00
Begin VB.Form frmIssue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Issue Book"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMiss 
      Caption         =   "Missing Book"
      Height          =   315
      Left            =   1050
      TabIndex        =   28
      Top             =   2175
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmdReserve 
      Caption         =   "Reserve Book"
      Height          =   315
      Left            =   1050
      TabIndex        =   27
      Top             =   1500
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enter.."
      Height          =   3135
      Left            =   75
      TabIndex        =   19
      Top             =   150
      Width           =   2535
      Begin VB.CommandButton cmdreturn 
         Caption         =   "&Return"
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdrenewal 
         Caption         =   "&Renewal"
         Height          =   315
         Left            =   1080
         TabIndex        =   23
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdcharge 
         Caption         =   "&Charge Fine"
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtBookId 
         Height          =   315
         Left            =   1095
         MaxLength       =   8
         TabIndex        =   1
         Top             =   735
         Width           =   1215
      End
      Begin VB.CommandButton cmdIssue 
         Caption         =   "&Issue Book"
         Height          =   315
         Left            =   1095
         TabIndex        =   2
         Top             =   1185
         Width           =   1215
      End
      Begin VB.TextBox txtMemId 
         Height          =   315
         Left            =   1095
         MaxLength       =   8
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Member ID"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   435
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Book ID"
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   810
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Book Information"
      Height          =   1815
      Left            =   2715
      TabIndex        =   4
      Top             =   1440
      Width           =   3375
      Begin VB.Label lblreturn 
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Return Date:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label dd 
         AutoSize        =   -1  'True
         Caption         =   "Is In:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label lblisin 
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblcondt 
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblres 
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbltitle 
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Condition:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Reserve ID:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Member Information"
      Height          =   1215
      Left            =   2715
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.Label lblbooks 
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblfinebal 
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblmemname 
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Books in hand:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fine Bal:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Member Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Connection
  
Private Sub cmdcharge_Click()
If Trim(txtMemId) = "" Or Trim(txtBookId) = "" Then
MsgBox "Please enter the fields.."
Else
On Error GoTo aderr
Dim adoprimaryrs0 As Recordset
Set adoprimaryrs0 = New Recordset
adoprimaryrs0.Open "select price,titleid from Titles where titleId = '" & lbltitle.Caption & "'", db, adOpenStatic, adLockOptimistic
Charge = adoprimaryrs0.Fields(0)
On Error GoTo aderr
Dim adoprimaryrs10 As Recordset
Set adoprimaryrs10 = New Recordset
adoprimaryrs10.Open "select memberid,finebal from members where memberId = '" & Trim(txtMemId.Text) & "'", db, adOpenStatic, adLockOptimistic
Charge = Charge + adoprimaryrs10.Fields(1)
adoprimaryrs10.Fields(1) = Charge
adoprimaryrs10.Update
Txtmemid_LostFocus
Exit Sub
aderr:
MsgBox Err.Description
End If
End Sub

Private Sub cmdIssue_Click()
If Trim(txtMemId) = "" Or Trim(txtBookId) = "" Then
MsgBox "Please enter the fields.."
Else
'''''''''''''''''''''''''''''
'make a module variable of max fine allowed to check here
If Val(lblfinebal.Caption) > M.MaxFineBal Then
MsgBox "Member should clear the Fines before issue"
txtBookId.Text = ""
txtMemId.SetFocus
Exit Sub
End If
'make a module variable of max books allowed to check here
If Val(lblbooks.Caption) >= M.TotalIssueBook Then
MsgBox " Memeber already has the maximum number of books"
txtBookId.Text = ""
txtMemId.SetFocus
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''
If UCase(lblisin.Caption) = "FALSE" Then
MsgBox "Book is not in the library"
txtBookId.Text = ""
txtMemId.SetFocus
Exit Sub
End If
If UCase(lblcondt.Caption) = "MISSING" Then
MsgBox "Book is Missing"
txtBookId.Text = ""
txtMemId.SetFocus
Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''
If Not lblres.Caption = "0" And Trim(lblres.Caption) <> Trim(txtMemId.Text) Then
MsgBox "This books is reserved by " & lblres.Caption
txtBookId.Text = ""
txtMemId.SetFocus
Exit Sub
End If


Dim adoprimaryrs1 As Recordset
Set adoprimaryrs1 = New Recordset
adoprimaryrs1.Open "select Subject from Titles where titleid ='" & lbltitle.Caption & "'", db, adOpenStatic, adLockOptimistic
subj = adoprimaryrs1.Fields(0)
''''''''''''''''''''
Dim adoprimaryrs2 As Recordset
Set adoprimaryrs2 = New Recordset
adoprimaryrs2.Open "select IssueDays,Noofbooks,FineCharge,ReserveCharge,Issuedays from Subjects where Subject='" & subj & "'", db, adOpenStatic, adLockOptimistic
noofdaysx = adoprimaryrs2.Fields(4)

'MsgBox adoprimaryrs2.Fields(0) & adoprimaryrs2.Fields(1) & adoprimaryrs2.Fields(2) & adoprimaryrs2.Fields(3)
''''''''''''''''''''''
Dim adoprimaryrs3 As Recordset
Set adoprimaryrs3 = New Recordset
adoprimaryrs3.Open "select titleid from books where memberid='" & Trim(txtMemId.Text) & "'", db, adOpenStatic, adLockOptimistic
Dim adoprimaryrs4 As Recordset
Set adoprimaryrs4 = New Recordset
While Not adoprimaryrs3.EOF
adoprimaryrs4.Open "select Subject from Titles where titleid ='" & adoprimaryrs3.Fields(0) & "'", db, adOpenStatic, adLockOptimistic
'MsgBox adoprimaryrs3.Fields(0) & adoprimaryrs4.Fields(0)
If subj = adoprimaryrs4.Fields(0) Then
Counter = Counter + 1
End If
adoprimaryrs3.MoveNext
adoprimaryrs4.Close
Wend
'''''''''''''''''''''''''''''''''''
If Counter >= adoprimaryrs2.Fields(1) Then
MsgBox "Member has taken maximum number of books in the Subject: " & subj
txtBookId.Text = ""
txtMemId.SetFocus
Exit Sub
End If
Dim adoprimaryrs5 As Recordset
Set adoprimaryrs5 = New Recordset
adoprimaryrs5.Open "select memberid,BOOKID,ISIN,returndate from books ", db, adOpenStatic, adLockOptimistic
adoprimaryrs5.Find ("BookId='" & txtBookId.Text & "'")
adoprimaryrs5.Fields(0) = Trim(txtMemId)
adoprimaryrs5.Fields(2) = False
adoprimaryrs5.Fields(3) = DateAdd("d", noofdaysx, Date)
adoprimaryrs5.Update
lblreturn.Caption = DateAdd("d", noofdaysx, Date)
Dim adoprimaryrs6 As Recordset
Set adoprimaryrs6 = New Recordset
adoprimaryrs6.Open "select MemberId,BooksInHand,FineBal from members where memberid= '" & Trim(txtMemId.Text) & "'", db, adOpenStatic, adLockOptimistic
''''''''''''''''''''''''''''''''''''''''''''''''
'specal code for check ing the reserve charge '''
If lblres.Caption = Trim(txtMemId.Text) Then
adoprimaryrs6.Fields(2) = adoprimaryrs6.Fields(1) + adoprimaryrs2.Fields(4)
End If
'''''''''''''''''''''''''''''''''''''''''''''''''
adoprimaryrs6.Fields(1) = adoprimaryrs6.Fields(1) + 1
adoprimaryrs6.Update
txtBookId_LostFocus
Txtmemid_LostFocus
txtMemId.SetFocus
End If
End Sub



Private Sub cmdrenewal_Click()
Dim adoprimaryrs1 As Recordset
Set adoprimaryrs1 = New Recordset
adoprimaryrs1.Open "select memberid,BOOKID,ISIN,returndate,issuecounter from books where memberid='" & Trim(txtMemId.Text) & "' and bookid='" & Trim(txtBookId.Text) & "'", db, adOpenStatic, adLockOptimistic
If adoprimaryrs1.RecordCount = 0 Then
MsgBox "Member:" & txtMemId & " Doesn't have the Book:" & txtBookId
Exit Sub
End If
''''''''''''''''''''''''''''''''''''
returndate = adoprimaryrs1.Fields(3)
If returndate < Date Then
MsgBox "You can't renewal this Book:" & txtBookId + vbCrLf + vbCrLf + "Please Return the Book and pay the Fine"
Exit Sub
End If
''''''''''''''''''''''''''''''''''''
'''change the number to global issue counter variable'''''''''''
IssueCounter = adoprimaryrs1.Fields(4)
If IssueCounter > M.RenewalCounter Then
MsgBox "You can't renewal this Book:" + txtBookId + vbCrLf + " Member:" & txtMemId & " have crossed the Renewal Limit"
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''
Dim adoprimaryrs11 As Recordset
Set adoprimaryrs11 = New Recordset
adoprimaryrs11.Open "select Subject from Titles where titleid ='" & lbltitle.Caption & "'", db, adOpenStatic, adLockOptimistic
subj = adoprimaryrs11.Fields(0)
'MsgBox subj
Dim adoprimaryrs12 As Recordset
Set adoprimaryrs12 = New Recordset
adoprimaryrs12.Open "select issuedays,subject from Subjects where Subject='" & subj & "'", db, adOpenStatic, adLockOptimistic
SubjectReturnDate = adoprimaryrs12.Fields(0)
adoprimaryrs1.Fields(3) = DateAdd("d", SubjectReturnDate, returndate)
adoprimaryrs1.Fields(4) = adoprimaryrs1.Fields(4) + 1
'MsgBox SubjectReturnDate
adoprimaryrs1.Update
lblreturn.Caption = adoprimaryrs1(3)
End Sub

Private Sub cmdReserve_Click()
'''''''''''''''''''''''''''''''''''''''''
If UCase(lblisin.Caption) = "TRUE" Then
MsgBox "Book is in the library"
txtBookId.Text = ""
txtMemId.SetFocus
Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''
If Not lblres.Caption = "0" And Trim(lblres.Caption) <> Trim(txtMemId.Text) Then
MsgBox "This books is reserved by " & lblres.Caption
txtBookId.Text = ""
txtMemId.SetFocus
Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''
Dim adoprimaryrs As Recordset
Set adoprimaryrs = New Recordset
adoprimaryrs.Open "select memberid,reserveid,bookid from books ", db, adOpenStatic, adLockOptimistic
adoprimaryrs.Find ("BookId='" & txtBookId.Text & "'")
adoprimaryrs.Fields(1) = Trim(txtMemId)
adoprimaryrs.Update
'''''''''''''''''''''''''''''''''''''''''''''''
Txtmemid_LostFocus
txtBookId_LostFocus
End Sub

Private Sub cmdreturn_Click()
Dim adoprimaryrs1 As Recordset
Set adoprimaryrs1 = New Recordset
adoprimaryrs1.Open "select MemberId,BookId,IsIn,ReturnDate,IssueCounter from Books where MemberId='" & Trim(txtMemId.Text) & "' and bookid='" & Trim(txtBookId.Text) & "'", db, adOpenStatic, adLockOptimistic

'''''''''''''''''''''''''''''''''''''''''''''''''''
If adoprimaryrs1.RecordCount = 0 Then
MsgBox "Member:" & txtMemId & " Doesn't have the Book:" & txtBookId
Exit Sub
End If
If adoprimaryrs1.Fields(2) = True Then
MsgBox "The book:" & txtBookId & " is already in the Library"
Exit Sub
End If
''''''''''''''''''''''''''''''''''''
returndate = adoprimaryrs1.Fields(3)
If returndate < Date Then
 FineDays = DateDiff("d", returndate, Date)
    ''''''''''''''''''''''''
    Dim adoprimaryrs11 As Recordset
    Set adoprimaryrs11 = New Recordset
    adoprimaryrs11.Open "select Subject from Titles where titleid ='" & lbltitle.Caption & "'", db, adOpenStatic, adLockOptimistic
    subj = adoprimaryrs11.Fields(0)
    ''''''''''''''''''
    Dim adoprimaryrs12 As Recordset
    Set adoprimaryrs12 = New Recordset
    adoprimaryrs12.Open "select issuedays,subject,finecharge from Subjects where Subject='" & subj & "'", db, adOpenStatic, adLockOptimistic
 FineCharge = adoprimaryrs12.Fields(2)
MsgBox "The Member has Kept the Book:" & txtBookId & " For " & FineDays & " days extra," & " and Must pay Rs." & FineDays * FineCharge & "/-", vbInformation, "Fine Charged.."
    Dim adoprimaryrs13 As Recordset
    Set adoprimaryrs13 = New Recordset
    adoprimaryrs13.Open "select MemberID,BooksInHand,FineBal from Members where MemberId = '" & Trim(txtMemId) & "'", db, adOpenStatic, adLockOptimistic
    adoprimaryrs13.Fields(1) = adoprimaryrs13.Fields(1) - 1
    adoprimaryrs13.Fields(2) = adoprimaryrs13.Fields(2) + (FineDays * FineCharge)
    adoprimaryrs1.Fields(2) = "True"
    adoprimaryrs1.Fields(0) = "0"
    adoprimaryrs13.Update
    adoprimaryrs1.Update
End If
End Sub

Private Sub cmdMiss_Click()
Dim adoprimaryrs1 As Recordset
Set adoprimaryrs1 = New Recordset
adoprimaryrs1.Open "select MemberId,BOOKID,ISIN,returndate,issuecounter,condition from books where memberid='" & Trim(txtMemId.Text) & "' and bookid='" & Trim(txtBookId.Text) & "'", db, adOpenStatic, adLockOptimistic
If adoprimaryrs1.RecordCount = 0 Then
MsgBox "Member:" & txtMemId & " Doesn't have the Book:" & txtBookId
Exit Sub
Else
Dim adoprimaryrs2 As Recordset
Set adoprimaryrs2 = New Recordset
adoprimaryrs2.Open "select price from titles,books where titles.titleid = books.titleid and books.bookid='" & Trim(txtBookId.Text) & "'", db, adOpenStatic, adLockOptimistic
'MsgBox adoprimaryrs2.Fields(0)
Dim adoprimaryrs3 As Recordset
Set adoprimaryrs3 = New Recordset
adoprimaryrs3.Open "select FineBal,memberid from Members where MemberId = '" & Trim(txtMemId) & "'", db, adOpenStatic, adLockOptimistic
adoprimaryrs3.Fields(0) = adoprimaryrs3.Fields(0) + adoprimaryrs2.Fields(0)
adoprimaryrs1.Fields(5) = "MISSING"
adoprimaryrs1.Update
adoprimaryrs3.Update
MsgBox "This is now marked as MISSING and its cost is added to Members Fine Balance"
End If
''''''''''''''''''''''''''''''''''''

End Sub

Private Sub Form_Load()
Set db = New Connection
  db.CursorLocation = adUseClient
 db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"
 End Sub

Private Sub Form_Unload(Cancel As Integer)
MDImain.munIssue.Enabled = True
MDImain.munRenewal.Enabled = True
MDImain.munReturn.Enabled = True
End Sub

Private Sub txtBookId_LostFocus()
Dim adoprimaryrs As Recordset
If Trim(txtBookId) = "" Then
MsgBox "Please Enter the Book ID"
Else
txtBookId.Text = UCase(txtBookId)
  Set adoprimaryrs = New Recordset
  adoprimaryrs.Open "select titleid,reserveid,condition,isin from Books where BookId = '" & Trim(txtBookId) & "'", db, adOpenStatic, adLockOptimistic
On Error GoTo oerr1:
lbltitle.Caption = adoprimaryrs.Fields(0)
lblres.Caption = adoprimaryrs.Fields(1)
lblcondt.Caption = adoprimaryrs.Fields(2)
lblisin.Caption = adoprimaryrs.Fields(3)
End If
Exit Sub
oerr1:
MsgBox "Book ID Not found ..Try again", vbInformation + vbOKOnly, "No Member ID"
txtBookId.Text = ""
txtBookId.SetFocus
End Sub

Private Sub Txtmemid_LostFocus()
Dim adoprimaryrs As Recordset
If Trim(txtMemId) = "" Then
MsgBox "Please Enter the member ID"
Else
txtMemId.Text = UCase(txtMemId)
Set adoprimaryrs = New Recordset
adoprimaryrs.Open "select FirstName,LastName,BooksInHand,FineBal from Members where MemberId = '" & Trim(txtMemId) & "'", db, adOpenStatic, adLockOptimistic
On Error GoTo oerr
lblmemname.Caption = adoprimaryrs.Fields(0) & " " & adoprimaryrs.Fields(1)
lblfinebal.Caption = adoprimaryrs.Fields(3)
lblbooks.Caption = adoprimaryrs.Fields(2)
End If
Exit Sub
oerr:
MsgBox "Member ID Not found ..Try again", vbInformation + vbOKOnly, "No Member ID"
txtMemId.Text = ""
txtMemId.SetFocus
End Sub
