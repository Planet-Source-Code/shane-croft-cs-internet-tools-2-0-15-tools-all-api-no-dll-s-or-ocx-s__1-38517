VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmNetStat1 
   Caption         =   "NetStat With API"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmNetStat1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   7710
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6480
      Top             =   4080
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   5280
      TabIndex        =   2
      Text            =   "60"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Update Every"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update List"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   4440
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   135
      TabIndex        =   3
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7011
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Local IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Remote IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Sec."
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Total In List:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   4080
      Width           =   2760
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   135
      TabIndex        =   4
      Top             =   4560
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FrmNetStat1.frx":030A
      Top             =   4080
      Width           =   480
   End
End
Attribute VB_Name = "FrmNetStat1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Long
Private Sub Command1_Click()
'On Error Resume Next

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command1_Click

  Dim pTcpTable As MIB_TCPTABLE
  Dim pdwSize As Long
  Dim bOrder As Long
  Dim nRet As Long
  Dim i As Integer, s As String
  ListView1.ListItems.Clear
  DoEvents
  nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
  nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
  For i = 0 To pTcpTable.dwNumEntries - 1
    If pTcpTable.table(i).dwState - 1 <> MIB_TCP_STATE_LISTEN Then
    Set Item = ListView1.ListItems.Add(, , c_ip(pTcpTable.table(i).dwLocalAddr))
    Item.SubItems(1) = c_port(pTcpTable.table(i).dwLocalPort)
    Item.SubItems(2) = c_ip(pTcpTable.table(i).dwRemoteAddr)
    Item.SubItems(3) = c_port(pTcpTable.table(i).dwRemotePort)
    Item.SubItems(4) = c_state(pTcpTable.table(i).dwState - 1)
    'Item.EnsureVisible
    Else
    Set Item = ListView1.ListItems.Add(, , c_ip(pTcpTable.table(i).dwLocalAddr))
    Item.SubItems(1) = c_port(pTcpTable.table(i).dwLocalPort)
    Item.SubItems(2) = c_ip(pTcpTable.table(i).dwRemoteAddr)
    Item.SubItems(3) = "0"
    Item.SubItems(4) = c_state(pTcpTable.table(i).dwState - 1)
    'Item.EnsureVisible
    End If
  Next
  DoEvents
    Me.MousePointer = vbNormal
    Label2.Caption = "Netstat status as of: " & Date & " " & Time
    Command1.Enabled = True

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

Private Sub Form_Load()

X = 0
Call Command1_Click

End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 1000
'DoEvents
ListView1.ColumnHeaders(1).Width = Me.ListView1.Width / 5 - 63
ListView1.ColumnHeaders(2).Width = Me.ListView1.Width / 5 - 63
ListView1.ColumnHeaders(3).Width = Me.ListView1.Width / 5 - 63
ListView1.ColumnHeaders(4).Width = Me.ListView1.Width / 5 - 63
ListView1.ColumnHeaders(5).Width = Me.ListView1.Width / 5 - 63
'DoEvents
Command1.Move Me.ScaleWidth - 1350, ListView1.Height + 150, Command1.Width, Command1.Height
Label3.Move Me.ScaleWidth - 1965, ListView1.Height + 150, Label3.Width, Label3.Height
Text1.Move Me.ScaleWidth - 2400, ListView1.Height + 150, Text1.Width, Text1.Height
Check1.Move Me.ScaleWidth - 3855, ListView1.Height + 150, Check1.Width, Check1.Height
Image1.Move Image1.Left, ListView1.Height + 150, Image1.Width, Image1.Height
Label1.Move 650, ListView1.Height + 150, Label1.Width, Label1.Height
Label2.Move 0, Me.ScaleHeight - 300, Me.ScaleWidth, Label2.Height

End Sub

Private Sub Timer1_Timer()

Label1.Caption = "Total In List: " & ListView1.ListItems.Count

End Sub


Private Sub Timer2_Timer()

On Error Resume Next

If Check1.Value = 1 Then

X = X + 1

If X >= Text1.text Then
Command1_Click
X = 0
End If

Else

X = 0
End If


End Sub
