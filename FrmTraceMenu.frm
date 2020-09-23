VERSION 5.00
Begin VB.Form FrmTraceMenu 
   Caption         =   "Trace Route"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTraceMenu.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   4320
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2115
         TabIndex        =   4
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   915
         TabIndex        =   3
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Use Trace Route Built Into Windows"
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
         Width           =   3360
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Use Trace Route With API"
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
         Width           =   3360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"FrmTraceMenu.frx":1982
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "When using this option you will use Trace Route used with direct API calls to Windows."
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3855
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3480
         Picture         =   "FrmTraceMenu.frx":1A16
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3480
         Picture         =   "FrmTraceMenu.frx":3398
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmTraceMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command1_Click

If Option1.Value = True Then
FrmTrace.Show
FrmTrace.SetFocus
End If

If Option2.Value = True Then
FrmTrace2.Show
FrmTrace2.SetFocus
End If

Unload Me

EXIT_Command1_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Command1_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Command1_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Command1_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Command1_Click

End Sub

Private Sub Command2_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command2_Click

Unload Me

EXIT_Command2_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Command2_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Command2_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Command2_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Command2_Click

End Sub

Private Sub Form_Load()
Me.Height = 5175
Me.Width = 4440
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Me.ScaleHeight / 2 - Frame1.Height / 2

End Sub
