VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3480
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "- Port Scan Options -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "Disable Port Scan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Port Scan Selected Ports"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Text            =   "21"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>>"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<<"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   420
         Left            =   1800
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Port Scan Selected Range"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtLowerBound 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "1"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtUpperBound 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "65535"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTo 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   2280
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "- Advanced Port Scan Options -"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   3255
      Begin VB.OptionButton Option4 
         Caption         =   "Do Port Scan After Each Valid Found IP Address."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Do Port Scan After IP Scan Has Finished."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtMaxConnections 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Text            =   "99"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Max Connections"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   2535
      Left            =   3345
      TabIndex        =   16
      ToolTipText     =   "Port Scan Progress"
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   4471
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Port Scan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3345
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
If txtLowerBound.Text = "" Then
MsgBox "Please Enter A Begining Port", vbCritical
txtLowerBound.SetFocus
Exit Sub
End If
If txtUpperBound.Text = "" Then
MsgBox "Please Enter A Ending Port", vbCritical
txtUpperBound.SetFocus
Exit Sub
End If
If txtMaxConnections.Text = "" Then
MsgBox "Please Enter A Max Connection", vbCritical
txtMaxConnections.SetFocus
Exit Sub
End If
End Sub

Private Sub Option1_Click()
On Error Resume Next
txtLowerBound.Enabled = False
txtUpperBound.Enabled = False
lblTo.Enabled = False
Text1.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
List1.Enabled = False
Frame3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Label2.Enabled = False
txtMaxConnections.Enabled = False
End Sub

Private Sub Option2_Click()
On Error Resume Next
txtLowerBound.Enabled = False
txtUpperBound.Enabled = False
lblTo.Enabled = False
Text1.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
List1.Enabled = True
Frame3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Label2.Enabled = True
txtMaxConnections.Enabled = True

End Sub

Private Sub Option3_Click()
On Error Resume Next
txtLowerBound.Enabled = True
txtUpperBound.Enabled = True
lblTo.Enabled = True
Text1.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
List1.Enabled = False
Frame3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Label2.Enabled = True
txtMaxConnections.Enabled = True
End Sub
