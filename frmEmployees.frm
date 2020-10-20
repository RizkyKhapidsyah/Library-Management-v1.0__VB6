VERSION 5.00
Begin VB.Form frmEmployees 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employees"
   ClientHeight    =   4965
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5910
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      Height          =   990
      Left            =   150
      TabIndex        =   28
      Top             =   3825
      Width           =   5565
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   150
         TabIndex        =   31
         Top             =   375
         Width           =   1065
      End
      Begin VB.ComboBox comSearch 
         Height          =   315
         ItemData        =   "frmEmployees.frx":0000
         Left            =   1350
         List            =   "frmEmployees.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   375
         Width           =   1590
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   3075
         TabIndex        =   29
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
      TabIndex        =   20
      Top             =   2700
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   59
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   21
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
      ScaleWidth      =   5985
      TabIndex        =   14
      Top             =   3075
      Width           =   5985
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   5370
         Picture         =   "frmEmployees.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   5025
         Picture         =   "frmEmployees.frx":0346
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmEmployees.frx":0688
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmEmployees.frx":09CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   19
         Top             =   0
         Width           =   4335
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tel"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Email"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LastName"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FirstName"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EmployeeId"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Password:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tel:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Email:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Address:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LastName:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "FirstName:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EmployeeId:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoprimaryrs As Recordset
Attribute adoprimaryrs.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean



Private Sub cmdSearch_Click()
txtSearch = Trim(txtSearch)
adoprimaryrs.MoveFirst
adoprimaryrs.Find (comSearch.Text & "='" & txtSearch & "'")
If adoprimaryrs.AbsolutePosition < 0 Then
MsgBox comSearch & "  Not Found!!!"
adoprimaryrs.MoveFirst
End If
End Sub

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"

  Set adoprimaryrs = New Recordset
  adoprimaryrs.Open "select EmployeeId,FirstName,LastName,Address,Email,Tel,Password from Employees", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoprimaryrs
  Next

  mbDataChanged = False
  
  comSearch.AddItem ("EmployeeId")
  comSearch.AddItem ("FirstName")
  comSearch.AddItem ("LastName")
  comSearch.AddItem ("Tel")
  comSearch.ListIndex = 0
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoprimaryrs.AbsolutePosition)
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
  answer = MsgBox("Are you sure of the changes made!!", vbYesNo)
  If answer = vbNo Then
  adoprimaryrs.CancelUpdate
  End If
  
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoprimaryrs
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description & "[cmdAdd_Click]"
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoprimaryrs
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description & "[cmdDelete_Click]"
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoprimaryrs.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description & "[cmdRefresh_Click]"
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description & "[cmdEdit_Click]"
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoprimaryrs.CancelUpdate
  If mvBookMark > 0 Then
    adoprimaryrs.Bookmark = mvBookMark
  Else
    adoprimaryrs.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoprimaryrs.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoprimaryrs.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description & "[cmdUpdate_Click]"
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

'i dont want to make canges while moving << or >>
'just use the add update buttons
adoprimaryrs.CancelUpdate

  adoprimaryrs.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description & "[cmdFirst_Click]"
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

'i dont want to make canges while moving << or >>
'just use the add update buttons
adoprimaryrs.CancelUpdate

  adoprimaryrs.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description & "[cmdLast_Click]"
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
  
'i dont want to make canges while moving << or >>
'just use the add update buttons
adoprimaryrs.CancelUpdate

  If Not adoprimaryrs.EOF Then adoprimaryrs.MoveNext
  If adoprimaryrs.EOF And adoprimaryrs.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoprimaryrs.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description & "[cmdNext_Click]"
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

'i dont want to make canges while moving << or >>
'just use the add update buttons
adoprimaryrs.CancelUpdate

  If Not adoprimaryrs.BOF Then adoprimaryrs.MovePrevious
  If adoprimaryrs.BOF And adoprimaryrs.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoprimaryrs.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description & "[cmdPrevious_Click]"
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
End Sub

Private Sub txtFields_LostFocus(Index As Integer)

'remove the extra spaces
txtFields(Index).Text = Trim(txtFields(Index).Text)
'MsgBox Index
'check for tel number
If Index = 5 Then

If Not IsNumeric(txtFields(Index).Text) Then
MsgBox "Enter a Telephone number!!!"
txtFields(Index).Text = ""
txtFields(Index).SetFocus
End If

End If

End Sub
