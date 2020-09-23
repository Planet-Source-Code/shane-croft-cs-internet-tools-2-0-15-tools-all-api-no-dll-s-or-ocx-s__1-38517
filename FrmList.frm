VERSION 5.00
Begin VB.Form FrmList 
   Caption         =   "List"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "FrmList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   4650
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4455
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
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Selected"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3720
         Width           =   2000
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save && Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3240
         Top             =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Total Entries In List:"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3120
         Picture         =   "FrmList.frx":1982
         Top             =   2040
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub List_Add(List As ListBox, txt As String)

On Error Resume Next
    List1.AddItem txt

End Sub

Public Sub List_Load(thelist As ListBox, Filename As String)
    'Loads a file to a list box

    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open Filename For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call List_Add(List1, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub

Public Sub List_Save(thelist As ListBox, Filename As String)

    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open Filename For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List1.List(Save)
    Next Save
    Close fFile

End Sub
Private Sub Command1_Click()

On Error Resume Next
If Text1.text = "" Then
MsgBox "Please enter a Computer Name,Web Site Address, or a IP"
FrmList.Text1.SetFocus
Exit Sub
End If
List1.AddItem Text1.text
DoEvents
Text1.text = ""
FrmList.Text1.SetFocus

End Sub

Private Sub Command2_Click()

On Error Resume Next
List1.RemoveItem List1.ListIndex

End Sub

Private Sub Command3_Click()

On Error Resume Next
Call List_Save(List1, App.Path & "\List.ini")
DoEvents
FrmOnline.Form_Load
DoEvents
Unload Me

End Sub

Private Sub Command4_Click()

Unload Me

End Sub

Private Sub Form_Load()

On Error Resume Next
Me.Height = 4770
Me.Width = 4770

Call List_Load(List1, App.Path & "\List.ini")
DoEvents
FrmList.Text1.SetFocus

End Sub

Private Sub Form_Resize()

On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Me.ScaleHeight / 2 - Frame1.Height / 2

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If

End Sub

Private Sub Text1_LostFocus()

On Error Resume Next
Text1.text = Replace(Text1.text, " ", "", 1, , vbTextCompare)

End Sub

Private Sub Timer1_Timer()

On Error Resume Next
Label2.Caption = List1.ListCount

End Sub
