VERSION 5.00
Begin VB.Form frmpayfine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pay Fine"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Member Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Member Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fine Bal:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Books in hand:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblmemname 
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblfinebal 
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblbooks 
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdpay 
      Caption         =   "Pay Fine"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtmemid 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Member ID"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmpayfine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Connection


Private Sub cmdpay_Click()
Dim adoprimaryrs1 As Recordset
Set adoprimaryrs1 = New Recordset
If Val(lblfinebal.Caption) = 0 Then
MsgBox "The Member has no fine balance"
txtMemId.Text = ""
txtMemId.SetFocus
lblbooks.Caption = ""
lblfinebal.Caption = ""
lblmemname.Caption = ""
Exit Sub
Else
On Error GoTo oerr:
adoprimaryrs1.Open "select FirstName,LastName,BooksInHand,FineBal from Members where MemberId = '" & Trim(txtMemId) & "'", db, adOpenStatic, adLockOptimistic
adoprimaryrs1.Fields(3) = 0
adoprimaryrs1.Update
Dim adoprimaryrs2 As Recordset
Set adoprimaryrs2 = New Recordset
On Error GoTo oerr:
adoprimaryrs2.Open "select Memberid,fineamount,paydate from fine where MemberId = '" & Trim(txtMemId) & "'", db, adOpenStatic, adLockOptimistic
adoprimaryrs2.AddNew
adoprimaryrs2.Fields(0) = Trim(txtMemId)
adoprimaryrs2.Fields(1) = Val(lblfinebal.Caption)
adoprimaryrs2.Fields(2) = Date
adoprimaryrs2.Update
txtMemId.Text = ""
txtMemId.SetFocus
lblbooks.Caption = ""
lblfinebal.Caption = ""
lblmemname.Caption = ""
End If
Exit Sub
oerr:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
Set db = New Connection
 db.CursorLocation = adUseClient
 db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"
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

