VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books"
   ClientHeight    =   5685
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6015
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2025
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   375
      Visible         =   0   'False
      Width           =   3390
   End
   Begin MSDataListLib.DataCombo DComTitleId 
      Height          =   315
      Left            =   2025
      TabIndex        =   37
      Top             =   75
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2025
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   675
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      Height          =   990
      Left            =   150
      TabIndex        =   32
      Top             =   4500
      Width           =   5565
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   150
         TabIndex        =   35
         Top             =   375
         Width           =   1065
      End
      Begin VB.ComboBox comSearch 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   375
         Width           =   1590
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   3075
         TabIndex        =   33
         Top             =   375
         Width           =   2340
      End
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   24
      Top             =   3300
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1200
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   75
      ScaleHeight     =   300
      ScaleWidth      =   7950
      TabIndex        =   18
      Top             =   3750
      Width           =   7950
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   5325
         Picture         =   "frmBooks.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4950
         Picture         =   "frmBooks.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmBooks.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmBooks.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   675
         TabIndex        =   23
         Top             =   0
         Width           =   4260
      End
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "TypeIssue"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ReserveId"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   2300
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "IssueCounter"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ReturnDate"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MemberId"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "IsIn"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Condition"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BookId"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TitleId"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "TypeIssue:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ReserveId:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "IssueCounter:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ReturnDate:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MemberId:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "IsIn:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Condition:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BookId:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TitleId:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoprimaryrs2 As Recordset
Attribute adoprimaryrs2.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub DComTitleId_Change()
Dim qunt As Integer
'Set adoPrimaryRS2 = New Recordset
adoprimaryrs2.MoveFirst
adoprimaryrs2.Find ("TitleId" & "='" & DComTitleId & "'")
qunt = adoprimaryrs2.Fields(1)
Combo2.Clear
For i = 1 To qunt
Combo2.AddItem (DComTitleId & "/" & i)
Next i
Combo2.ListIndex = 0
txtFields(1).Text = Combo2.Text
End Sub

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select TitleId,BookId,Condition,IsIn,MemberId,ReturnDate,IssueCounter,ReserveId,TypeIssue from Books", db, adOpenStatic, adLockOptimistic

  Set adoprimaryrs2 = New Recordset
  adoprimaryrs2.Open "select TitleId,Quantity from Titles", db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Dim oCheck As CheckBox
  'Bind the check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
  '''''''''''''''''''''''''''''''''''''
  Combo1.AddItem ("EXCELLENT")
  Combo1.AddItem ("GOOD")
  Combo1.AddItem ("POOR")
  Combo1.AddItem ("WORST")
  Combo1.AddItem ("MISSING")
  Combo1.ListIndex = 0
  comSearch.AddItem ("BookId")
  comSearch.AddItem ("TitleId")
  comSearch.AddItem ("MemberId")
  comSearch.AddItem ("ReserveId")
  comSearch.ListIndex = 0
  ''''''''''''''''''''''''''''''''''''''
txtFields(0).Locked = True
txtFields(1).Locked = True
''''''''''''''''''''''''''''''''''''''''''
Set DComTitleId.DataSource = adoprimaryrs2
Set DComTitleId.RowSource = adoprimaryrs2
 DComTitleId.ListField = "TitleId"
''''''''''''''''''''''''''''''''''''''''''
 If M.BooksCallByTitle Then
 cmdAdd_Click
 End If
 
  End Sub




Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  '''''''''''''''''''
  DComTitleId.Refresh
  '''''''''''''''''''
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With
txtFields(6).Text = "0"
txtFields(4) = "0"
txtFields(7) = "0"
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  Dim q As Integer
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  '''''''''''''''''''''''''''''''''''
'update the title table '''''''''''''
adoprimaryrs2.MoveFirst
adoprimaryrs2.Find ("TitleId" & "='" & txtFields(0).Text & "'")
q = adoprimaryrs2.Fields(1)
q = q - 1
adoprimaryrs2.Fields(1) = q
adoprimaryrs2.Update
''''''''''''''''''''''''''''''''''''
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
'''''''''''''''''''''''''''''
'hide the combos cause no edit allowed to title id and bookid
Combo2.Visible = False
DComTitleId.Visible = False
'''''''''''''''''''''''''''''
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  ''''''''''''''''''''''''''''''''
txtFields(2).Text = Combo1.Text
txtFields(0).Text = DComTitleId.Text
txtFields(1).Text = Combo2.Text

''''''''''''''''''''''''''''''''''
  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoPrimaryRS.CancelUpdate

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoPrimaryRS.CancelUpdate

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoPrimaryRS.CancelUpdate

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoPrimaryRS.CancelUpdate

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  ''extra code for our frame
  Frame1.Enabled = bVal
  Combo1.Visible = Not bVal
  DComTitleId.Visible = Not bVal
  Combo2.Visible = Not bVal


End Sub

Private Sub txtFields_Change(Index As Integer)
txtFields(Index).Text = UCase(Trim(txtFields(Index).Text))
End Sub
