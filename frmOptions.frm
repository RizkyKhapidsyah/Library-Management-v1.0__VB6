VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Library Options"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRenewal 
      Height          =   285
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   12
      Top             =   2700
      Width           =   735
   End
   Begin VB.TextBox txtFees 
      Height          =   285
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   10
      Top             =   2250
      Width           =   735
   End
   Begin VB.TextBox txtDuration 
      Height          =   285
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2325
      TabIndex        =   7
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1125
      TabIndex        =   6
      Top             =   3300
      Width           =   1095
   End
   Begin VB.TextBox txtMaxFine 
      Height          =   285
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1350
      Width           =   735
   End
   Begin VB.TextBox txtRenualCounter 
      Height          =   285
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   4
      Top             =   825
      Width           =   735
   End
   Begin VB.TextBox txtTotalIssue 
      Height          =   285
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   3
      Top             =   300
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   -7200
      X2              =   -1636
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   3
      X1              =   -7200
      X2              =   -1636
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   -7200
      X2              =   -1636
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   -7200
      X2              =   -1636
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -7200
      X2              =   -1636
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Membership Renewal Fees"
      Height          =   195
      Left            =   600
      TabIndex        =   13
      Top             =   2700
      Width           =   1920
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Membership Fees"
      Height          =   195
      Left            =   1275
      TabIndex        =   11
      Top             =   2250
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Membership Duration(In Months)"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Maximum Fine Balance allowed "
      Height          =   195
      Left            =   255
      TabIndex        =   2
      Top             =   1350
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum Number of Renuals a Member can make"
      Height          =   495
      Left            =   150
      TabIndex        =   1
      Top             =   750
      Width           =   2370
   End
   Begin VB.Label Label1 
      Caption         =   "Maximun Number of Book Issued to a Member"
      Height          =   420
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   2370
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
MDImain.munOptions.Enabled = True
Unload Me
End Sub

Private Sub cmdOK_Click()
M.TotalIssueBook = txtTotalIssue.Text
M.MaxFineBal = txtMaxFine.Text
M.RenewalCounter = txtRenualCounter

M.MembershipDuration = txtDuration
M.MembershipFee = txtFees
M.RenewalFees = txtRenewal
''''''''''''''''''''''''''''''''''''
Dim db As Connection, adoprimaryrs As Recordset
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"
'''''''''''''''''''''''''''''''''''''
Set adoprimaryrs = New Recordset
adoprimaryrs.Open "select TotalIssueBooks,RenewalCounter,MaxFineBal,MembershipDuration, MembershipFee, RenewalFees from GlobalVariables", db, adOpenStatic, adLockOptimistic
adoprimaryrs.Fields(0) = txtTotalIssue.Text
adoprimaryrs.Fields(1) = txtRenualCounter
adoprimaryrs.Fields(2) = txtMaxFine.Text

adoprimaryrs.Fields(3) = txtDuration
adoprimaryrs.Fields(4) = txtFees
adoprimaryrs.Fields(5) = txtRenewal

adoprimaryrs.Update
''''''''''''''''''''''''''''''''''''''
db.Close
''''''''''''''''''''''''''''''''''''''
MDImain.munOptions.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
txtTotalIssue.Text = M.TotalIssueBook
txtMaxFine.Text = M.MaxFineBal
txtRenualCounter = M.RenewalCounter

txtDuration = M.MembershipDuration
txtFees = M.MembershipFee
txtRenewal = M.RenewalFees
End Sub






Private Sub txtDuration_LostFocus()
With txtDuration
If Not IsNumeric(.Text) Then
MsgBox "Enter a Number!!!"
.Text = ""
.SetFocus
End If
End With

End Sub



Private Sub txtFees_LostFocus()
With txtFees
If Not IsNumeric(.Text) Then
MsgBox "Enter a Number!!!"
.Text = ""
.SetFocus
End If
End With

End Sub

Private Sub txtMaxFine_LostFocus()
With txtMaxFine
If Not IsNumeric(.Text) Then
MsgBox "Enter a Number!!!"
.Text = ""
.SetFocus
End If
End With

End Sub

Private Sub txtRenewal_LostFocus()
With txtRenewal
If Not IsNumeric(.Text) Then
MsgBox "Enter a Number!!!"
.Text = ""
.SetFocus
End If
End With

End Sub

Private Sub txtRenualCounter_LostFocus()
With txtRenualCounter
If Not IsNumeric(.Text) Then
MsgBox "Enter a Number!!!"
.Text = ""
.SetFocus
End If
End With

End Sub


Private Sub txtTotalIssue_LostFocus()
With txtTotalIssue
If Not IsNumeric(.Text) Then
MsgBox "Enter a Number!!!"
.Text = ""
.SetFocus
End If
End With
End Sub
