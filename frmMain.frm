VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1110
      TabIndex        =   3
      Text            =   "50"
      Top             =   2550
      Width           =   495
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   1635
      Left            =   240
      TabIndex        =   2
      Top             =   2100
      Width           =   5475
      Begin VB.CommandButton Command2 
         Caption         =   "Clear All"
         Height          =   435
         Left            =   3180
         TabIndex        =   12
         Top             =   720
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add Item"
         Height          =   435
         Left            =   3180
         TabIndex        =   10
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   "Knicks"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   2160
         TabIndex        =   7
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1230
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Value:                %"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   480
         Width           =   1290
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5460
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.jChart UserControl11 
      Height          =   1695
      Left            =   435
      Top             =   360
      Width           =   5295
      _ExtentX        =   9551
      _ExtentY        =   2990
      BackColor       =   16761087
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Stinkage Percent of NBA Teams."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   11
      Top             =   60
      Width           =   3405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "100"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   360
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   285
      TabIndex        =   0
      Top             =   1860
      Width           =   90
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    UserControl11.AddItem Text2, Val(Text1), Label5(1).BackColor, Label5(0).BackColor
End Sub

Private Sub Command2_Click()
    UserControl11.Clear
End Sub

Private Sub Label5_Click(Index As Integer)
    On Error GoTo Err
    With CommonDialog1
        .CancelError = True
        .ShowColor
        Label5(Index).BackColor = .Color
    End With
Err:
End Sub

Private Sub Label6_Click()
    On Error GoTo Err
    With CommonDialog1
        .CancelError = True
        .ShowColor
        Label6.BackColor = .Color
    End With
Err:
End Sub

Private Sub Label9_Click()
    On Error GoTo Err
    With CommonDialog1
        .CancelError = True
        .ShowColor
        Label9.BackColor = .Color
    End With
Err:
End Sub
