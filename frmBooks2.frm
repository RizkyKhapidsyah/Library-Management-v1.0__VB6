VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBooks2 
   Caption         =   "Books"
   ClientHeight    =   5880
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   9720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9720
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   9660
      TabIndex        =   2
      Top             =   0
      Width           =   9720
      Begin VB.Frame Frame1 
         Caption         =   "Search Options"
         Height          =   990
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   5565
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   315
            Left            =   150
            TabIndex        =   6
            Top             =   375
            Width           =   1065
         End
         Begin VB.ComboBox comSearch 
            Height          =   315
            ItemData        =   "frmBooks2.frx":0000
            Left            =   1350
            List            =   "frmBooks2.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   375
            Width           =   1590
         End
         Begin VB.TextBox txtSearch 
            Height          =   315
            Left            =   3075
            TabIndex        =   4
            Top             =   375
            Width           =   2340
         End
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   300
      Left            =   3840
      TabIndex        =   0
      Top             =   5400
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4080
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   7197
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   4
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      GridColor       =   12632256
      GridColorFixed  =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "TitleId|BookId|TypeIssue|IsIn|Condition|ReserveId|ReturnDate"
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
End
Attribute VB_Name = "frmBooks2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MARGIN_SIZE = 60      ' in Twips
' variables for data binding
 Private dfwConn As ADODB.Connection
Private datPrimaryRS As ADODB.Recordset

Private Sub cmdSearch_Click()
Dim SQLx As String
SQLx = "select Titles.TitleId,Titles.Subject,Titles.Title,Titles.Author,Books.BookId,Books.TypeIssue,Books.Condition,Books.IsIn,Books.ReserveId from Books,Titles where Titles.TitleId=Books.TitleID  and Titles." & comSearch & "='" & Trim(txtSearch.Text) & "'"
'MsgBox SQLx
datPrimaryRS.Close
datPrimaryRS.Open SQLx, dfwConn, adOpenForwardOnly, adLockReadOnly
MsgBox datPrimaryRS.AbsolutePage
MSHFlexGrid1.Refresh
End Sub

Private Sub Form_Load()

''''''''''''''''''''''''''''''''''''''''
  comSearch.AddItem ("Title")
  comSearch.AddItem ("Author")
  comSearch.AddItem ("Subject")
  comSearch.AddItem ("TitleId")
  comSearch.ListIndex = 0
  ''''''''''''''''''''''''''''''''''
    Dim sConnect As String
    Dim sSQL As String
   

    ' set strings
    sConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;User ID=Admin;Data Source=C:\WINDOWS\Desktop\library\Library.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Locale Identifier=1033;Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='';Jet OLEDB:Global Partial Bulk Ops=2"
    sSQL = "select Titles.TitleId,Titles.Subject,Titles.Title,Titles.Author,Books.BookId,Books.TypeIssue,Books.Condition,Books.IsIn,Books.ReserveId from Books,Titles where Titles.TitleId=Books.TitleID"

    ' open connection
    Set dfwConn = New Connection
    dfwConn.Open sConnect

    ' create a recordset using the provided collection
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, dfwConn, adOpenForwardOnly, adLockReadOnly

    Set MSHFlexGrid1.DataSource = datPrimaryRS

    With MSHFlexGrid1

        .Redraw = False
        ' set grid's column widths
        .ColWidth(0) = -1
        .ColWidth(1) = -1
        .ColWidth(2) = -1
        .ColWidth(3) = -1
        .ColWidth(4) = -1
        .ColWidth(5) = -1
        .ColWidth(6) = -1

        ' set grid's style
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' make header bold
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With

End Sub



Private Sub cmdClose_Click()

    Unload Me

End Sub


