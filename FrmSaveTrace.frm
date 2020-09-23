VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSaveTrace 
   Caption         =   "Save Trace Route"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSaveTrace.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   7425
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   3
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save To File"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   2
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   6240
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6240
         Picture         =   "FrmSaveTrace.frx":1982
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmSaveTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub List_Add(List As ListBox, txt As String)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_List_Add

On Error Resume Next
    List1.AddItem txt

EXIT_List_Add:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_List_Add:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in List_Add" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_List_Add
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_List_Add

End Sub

Public Sub List_Load(thelist As ListBox, Filename As String)
    'Loads a file to a list box

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_List_Load

    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open Filename For Input As fFile
    Do
        Line Input #fFile, TheContents$
        Call List_Add(List1, TheContents$)
    Loop Until EOF(fFile)
    Close fFile

EXIT_List_Load:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_List_Load:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in List_Load" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_List_Load
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_List_Load

End Sub

Public Sub List_Save(thelist As ListBox, Filename As String)
    'Save a listbox as FileName

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_List_Save

    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open Filename For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List1.List(Save)
    Next Save
    Close fFile

EXIT_List_Save:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_List_Save:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in List_Save" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_List_Save
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_List_Save

End Sub

Private Sub Command1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command1_Click

On Error GoTo exitme
Dim Filename As String
CD1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
CD1.DefaultExt = "txt"
CD1.DialogTitle = "Select the destination file"
CD1.Filename = "Traced_" & FrmTrace.Host.text & ".txt"
CD1.CancelError = True
CD1.ShowSave
Filename = CD1.Filename

Call List_Save(List1, Filename)
exitme:

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

Me.Height = 4320
Me.Width = 7545

End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Me.ScaleHeight / 2 - Frame1.Height / 2

End Sub
