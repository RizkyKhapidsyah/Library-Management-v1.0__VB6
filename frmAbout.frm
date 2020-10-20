VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3990
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6345
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2753.968
   ScaleMode       =   0  'User
   ScaleWidth      =   5958.283
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   5250
      Top             =   3975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2505
      TabIndex        =   0
      Top             =   2175
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   690
      Left            =   135
      Top             =   2775
      Width           =   6015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Now get back to Planet-Source-Code.com and Vote me 5 globes!!!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   210
      TabIndex        =   4
      Top             =   2775
      Width           =   5865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail : rafaymansoor@yahoo.com"
      Height          =   195
      Left            =   1905
      TabIndex        =   3
      ToolTipText     =   ".   windows_me@rediffmail.com   ."
      Top             =   495
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Author :Abdul  Rafay  Mansoor"
      Height          =   195
      Left            =   2055
      TabIndex        =   2
      ToolTipText     =   ".   Abdul Rafay Mansoor   ."
      Top             =   225
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   281.716
      X2              =   5606.139
      Y1              =   507.31
      Y2              =   507.31
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":0442
      ForeColor       =   &H00000000&
      Height          =   1125
      Left            =   315
      TabIndex        =   1
      ToolTipText     =   ".   Just Kidding !!!!!!!   ."
      Top             =   840
      Width           =   5655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Shape1.BorderColor = vbBlack

End Sub

Private Sub Timer1_Timer()
If Shape1.BorderColor = vbBlack Then
Shape1.BorderColor = vbRed
Else
Shape1.BorderColor = vbBlack
End If
End Sub
