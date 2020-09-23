VERSION 5.00
Begin VB.Form FrmNetStat 
   Caption         =   "NetStat"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmNetStat.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   4050
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton Option1 
         Caption         =   "Use NetStat With API"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   3120
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Use NetStat Built Into Windows"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   3120
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   780
         TabIndex        =   0
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1980
         TabIndex        =   3
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3240
         Picture         =   "FrmNetStat.frx":030A
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3240
         Picture         =   "FrmNetStat.frx":0614
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "When using this option you will use Netstat used with direct API calls to Windows."
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   180
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"FrmNetStat.frx":1F96
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   3615
      End
   End
End
Attribute VB_Name = "FrmNetStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Option1.Value = True Then
FrmNetStat1.Show
FrmNetStat1.SetFocus
End If

If Option2.Value = True Then
FrmNetStat2.Show
FrmNetStat2.SetFocus
End If

Unload Me

End Sub

Private Sub Command2_Click()


Unload Me


End Sub

Private Sub Form_Load()

Me.Height = 5130
Me.Width = 4140

End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Me.ScaleHeight / 2 - Frame1.Height / 2

End Sub
