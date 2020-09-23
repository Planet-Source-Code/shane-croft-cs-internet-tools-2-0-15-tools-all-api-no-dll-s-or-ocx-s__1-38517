VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmIpScan 
   Caption         =   "IP Address Scanner"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3570
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmIpScan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   3570
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   29
         Text            =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "- IP Information -"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox Check1 
            Caption         =   "Resolve IP's To Their Host Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.PictureBox picIP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            MousePointer    =   3  'I-Beam
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   105
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   360
            Width           =   1575
            Begin VB.TextBox txtip 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   1
               Left            =   480
               MaxLength       =   3
               TabIndex        =   2
               Text            =   "168"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtip 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   0
               Left            =   135
               MaxLength       =   3
               TabIndex        =   1
               Text            =   "192"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtip 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   2
               Left            =   825
               MaxLength       =   3
               TabIndex        =   3
               Text            =   "0"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtip 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   3
               Left            =   1170
               MaxLength       =   3
               TabIndex        =   4
               Text            =   "1"
               Top             =   0
               Width           =   285
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   0
               Left            =   420
               TabIndex        =   26
               Top             =   0
               Width           =   60
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   2
               Left            =   765
               TabIndex        =   25
               Top             =   0
               Width           =   60
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   3
               Left            =   1110
               TabIndex        =   24
               Top             =   0
               Width           =   60
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            MousePointer    =   3  'I-Beam
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   105
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   720
            Width           =   1575
            Begin VB.TextBox txtip2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00C00000&
               Height          =   300
               Index           =   3
               Left            =   1170
               MaxLength       =   3
               TabIndex        =   5
               Text            =   "255"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtip2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   825
               Locked          =   -1  'True
               MaxLength       =   3
               TabIndex        =   19
               Text            =   "0"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtip2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   3
               TabIndex        =   18
               Text            =   "192"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtip2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   480
               Locked          =   -1  'True
               MaxLength       =   3
               TabIndex        =   17
               Text            =   "168"
               Top             =   0
               Width           =   285
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   5
               Left            =   1110
               TabIndex        =   22
               Top             =   0
               Width           =   60
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   4
               Left            =   765
               TabIndex        =   21
               Top             =   0
               Width           =   60
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   1
               Left            =   420
               TabIndex        =   20
               Top             =   0
               Width           =   60
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Start Address:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "End Address:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Begin Scan"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Close"
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         Caption         =   "- Results -"
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
         Height          =   4935
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   3255
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0.0.0.0"
            Top             =   480
            Width           =   3015
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   3375
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Found IP Addresses"
            Top             =   1200
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   5953
            _Version        =   393217
            LineStyle       =   1
            Style           =   6
            FullRowSelect   =   -1  'True
            Appearance      =   0
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "IP Scan Progress"
            Top             =   840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Total Found:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "0"
            Height          =   255
            Left            =   1320
            TabIndex        =   13
            Top             =   4560
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Current IP"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3015
         End
      End
   End
End
Attribute VB_Name = "FrmIpScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Type HOSTENT2
   hName     As Long
   hAliases  As Long
   hAddrType As Integer
   hLength   As Integer
   hAddrList As Long
End Type

Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type
'
Private Type WSADATA
    wversion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Private Const ERROR_BUFFER_OVERFLOW = 111&
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_NO_DATA = 232&
Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_SUCCESS = 0&
'
Private Const MIB_TCP_STATE_CLOSED = 1
Private Const MIB_TCP_STATE_LISTEN = 2
Private Const MIB_TCP_STATE_SYN_SENT = 3
Private Const MIB_TCP_STATE_SYN_RCVD = 4
Private Const MIB_TCP_STATE_ESTAB = 5
Private Const MIB_TCP_STATE_FIN_WAIT1 = 6
Private Const MIB_TCP_STATE_FIN_WAIT2 = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT = 8
Private Const MIB_TCP_STATE_CLOSING = 9
Private Const MIB_TCP_STATE_LAST_ACK = 10
Private Const MIB_TCP_STATE_TIME_WAIT = 11
Private Const MIB_TCP_STATE_DELETE_TCB = 12

Private mWSData As WSADataType ' this will hold the wsadata we need
'
'
'
Dim X As Long
Dim xx As Long
Dim PingTimes As Long
Dim Speed As Long
Dim IP As String
Dim KeepGoing As Long
Dim TotalNum As Long
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim HOSTENT As HOSTENT, PointerToPointer As Long, ListAddress As Long
Dim WSADATA As WSADATA, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Long
Dim Description As String, Status As String
Dim ExitTheFor As Long

Dim IPAddress_Number As String
Dim StopIt As Boolean
Dim lngNextPort As Long
Dim intI As Long
Dim intI2 As Long

' Ping Variables
Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim addr As Long
Dim RCode As String
Dim RespondingHost As String
' TRACERT Variables
Dim TraceRT As Boolean
Dim TTL As Integer
' WSock32 Constants
Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0
Public Sub Ping_IP()

On Error Resume Next
Dim Times As Integer
IPAddress_Number = Text3.text
    Speed = 0
    PingTimes = 0

    szBuffer = Space(Val("32"))
    DoEvents
    vbWSAStartup
    DoEvents
    If Len(IPAddress_Number) = 0 Then
        vbGetHostName
    End If
    DoEvents
    vbGetHostByName
    vbIcmpCreateFile
    DoEvents
    pIPo2.TTL = Trim$(255)
    '
    'For Times = 1 To "1"
    'If ExitTheFor = 1 Then ExitTheFor = 0: Exit For
    vbIcmpSendEcho
    DoEvents
    'Next
    DoEvents
    vbIcmpCloseHandle
    vbWSACleanup

    'Speed = Speed / PingTimes

End Sub
Public Sub Resolve_IP()

On Error Resume Next
' The inet_addr function returns a long value
    Dim lInteAdd As Long
' pointer to the HOSTENT
    Dim lPointtoHost As Long
' host name we are looking for
    Dim sHost As String
' Hostent
    Dim mHost As HOSTENT2
' IP Address
    Dim sIP As String

    sIP = Trim$(Text3.text)

' Convert the IP address
    lInteAdd = inet_addr(sIP)

' if the wrong IP format was entered there is an err generated
    If lInteAdd = INADDR_NONE Then

        'WSErrHandle (Err.LastDllError)
        TreeView1.Nodes.Add IPAddress_Number, tvwChild, , "Unable To Resolve"
    Else

' pointer to the Host
        lPointtoHost = gethostbyaddr(lInteAdd, 4, PF_INET)

' if zero is returned then there was an error
        If lPointtoHost = 0 Then

            'WSErrHandle (Err.LastDllError)
            TreeView1.Nodes.Add IPAddress_Number, tvwChild, , "Unable To Resolve"

        Else

            RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

            sHost = String(256, 0)

' Copy the host name
            RtlMoveMemory ByVal sHost, ByVal mHost.hName, 256

' Cut the chr(0) character off
            sHost = Left(sHost, InStr(1, sHost, Chr(0)) - 1)

' Return the host name
            TreeView1.Nodes.Add IPAddress_Number, tvwChild, , sHost

        End If

    End If
TreeView1.Nodes.Item(TreeView1.Nodes.Count - 1).Expanded = True
TreeView1.Nodes.Item(TreeView1.Nodes.Count - 1).EnsureVisible

End Sub

Private Sub Command4_Click()

On Error Resume Next

Dim YYY As Long
Dim XXX As Long

YYY = txtip(3).text
XXX = txtip2(3).text
DoEvents
If YYY > XXX Then
MsgBox "Starting IP Address Can't Be Greater Than Ending Address"
txtip2(3).SetFocus
DoEvents
SendKeys "{HOME}+{END}"
Exit Sub
End If
DoEvents
ProgressBar1.Max = XXX - YYY + 1
ProgressBar1.Min = 0
ProgressBar1.Value = 0
DoEvents

If Command4.Caption = "Stop Scan" Then
StopIt = True
Command4.Caption = "Begin Scan"
Command5.Enabled = True
Exit Sub
End If
Command5.Enabled = False
Command4.Caption = "Stop Scan"
DoEvents
TreeView1.Nodes.Clear
DoEvents
Label8.Caption = "0"
DoEvents
DoEvents
Text4.text = txtip(3).text
DoEvents
Text3.text = txtip(0).text & "." & txtip(1).text & "." & txtip(2).text & "." & Text4.text

Do Until Text4.text = txtip2(3).text + 1

If StopIt = True Then
StopIt = False
Exit Do
End If
Call Ping_IP
DoEvents
If Check1.Value = 1 Then
Call Resolve_IP
DoEvents
End If
DoEvents
ProgressBar1.Value = ProgressBar1.Value + 1
Text4.text = Text4.text + 1
DoEvents
Text3.text = txtip(0).text & "." & txtip(1).text & "." & txtip(2).text & "." & Text4.text
DoEvents
Loop


Text3.text = "Done"
Command4.Caption = "Begin Scan"
Command5.Enabled = True

End Sub

Private Sub Command5_Click()

On Error Resume Next
DoEvents
Unload Me

End Sub

Private Sub Form_Load()

On Error Resume Next

Me.Height = 7665
Me.Width = 3690

Dim Y As String
StopIt = False
Dim mWSD As WSADataType
Call WSAStartup(&H202, mWSD)
vbWSAStartup
vbWSACleanup


txtip(0).text = Split(FrmMain.StatusBar1.Panels(3).text, ".")(0)
txtip(1).text = Split(FrmMain.StatusBar1.Panels(3).text, ".")(1)
txtip(2).text = Split(FrmMain.StatusBar1.Panels(3).text, ".")(2)
txtip(3).text = Split(FrmMain.StatusBar1.Panels(3).text, ".")(3)
DoEvents
txtip2(0).text = Split(FrmMain.StatusBar1.Panels(3).text, ".")(0)
txtip2(1).text = Split(FrmMain.StatusBar1.Panels(3).text, ".")(1)
txtip2(2).text = Split(FrmMain.StatusBar1.Panels(3).text, ".")(2)
txtip2(3).text = "255"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next
If Command4.Caption = "Stop Scan" Then
Cancel = True
Exit Sub
End If

End Sub

Private Sub Form_Resize()

On Error Resume Next
Frame2.Move Me.ScaleWidth / 2 - Frame2.Width / 2, Me.ScaleHeight / 2 - Frame2.Height / 2

End Sub

Private Sub txtip2_Change(Index As Integer)

On Error Resume Next
  'On Error Resume Next
  'If the section = "" we need to put a value there
  'If txtip2(Index) = "" Then txtip2(Index) = "0": SendKeys "{HOME}+{END}"
  'Now we need to set a range of numbers allowed.
  'If CInt(txtip2(Index).Text) > 255 Then
  '  MsgBox "Number must be between 0 - 255." & Chr(13) & "Please re-enter number.", vbApplicationModal + vbDefaultButton1 + vbInformation, "Error"
  '  SendKeys "{HOME}+{END}"
  'End If
  'If Len(txtip2(Index).Text) = 3 Then
  '  If Index = txtip2.Count - 1 Then
  '    txtip2(0).SetFocus
  '  Else
  '    txtip2(Index + 1).SetFocus
  '  End If
  'End If

End Sub
Private Sub txtip_Change(Index As Integer)

  On Error Resume Next
  'If the section = "" we need to put a value there
  If txtip(Index) = "" Then txtip(Index) = "0": SendKeys "{HOME}+{END}"
  'Now we need to set a range of numbers allowed.
  If CInt(txtip(Index).text) > 255 Then
    MsgBox "Number must be between 0 - 255." & Chr(13) & "Please re-enter number.", vbApplicationModal + vbDefaultButton1 + vbInformation, "Error"
    SendKeys "{HOME}+{END}"
  End If
  If Len(txtip(Index).text) = 3 Then
    If Index = txtip.Count - 1 Then
      txtip(0).SetFocus
    Else
      txtip(Index + 1).SetFocus
    End If
  End If

  txtip2(0).text = txtip(0).text
  txtip2(1).text = txtip(1).text
  txtip2(2).text = txtip(2).text

End Sub
Private Sub txtip_Click(Index As Integer)

On Error Resume Next
  'select the section
  SendKeys "{HOME}+{END}"

End Sub

Private Sub txtip_GotFocus(Index As Integer)

On Error Resume Next
  'Select the section
  SendKeys "{HOME}+{END}"

End Sub

Private Sub txtip_KeyPress(Index As Integer, KeyAscii As Integer)

  On Error Resume Next
  Dim tindex As Integer
  'If the '.' or the 'Enter' Is pressed then goto the next section
  If KeyAscii = Asc(".") Or KeyAscii = 13 Then
    If Index = txtip.Count - 1 Then
      tindex = 0
      txtip(tindex).SetFocus
    Else
      tindex = Index + 1
      txtip(tindex).SetFocus
    End If
  End If

End Sub
Public Sub GetRCode()

On Error Resume Next
RCode = ""
    If pIPe.Status = 0 Then RCode = "Success"
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Destination Unreahable"
    If pIPe.Status = 11003 Then RCode = "Dest Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Dest Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Dest Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Reqested Timed Out"
    If pIPe.Status = 11011 Then RCode = "Bad Request"
    If pIPe.Status = 11012 Then RCode = "Bad Route"
    If pIPe.Status = 11014 Then RCode = "TTL Exprd Reassemb"
    If pIPe.Status = 11015 Then RCode = "Parameter Problem"
    If pIPe.Status = 11016 Then RCode = "Source Quench"
    If pIPe.Status = 11017 Then RCode = "Option too Big"
    If pIPe.Status = 11018 Then RCode = "Bad Destination"
    If pIPe.Status = 11019 Then RCode = "Address Deleted"
    If pIPe.Status = 11020 Then RCode = "Spec MTU Change"
    If pIPe.Status = 11021 Then RCode = "MTU Change"
    If pIPe.Status = 11022 Then RCode = "Unload"
    If pIPe.Status = 11050 Then RCode = "General Failure"

    DoEvents
DoEvents
        If RCode <> "" Then
        DoEvents
            If RCode = "Success" Then
                'Speed = Speed wct+ Val(Trim$(CStr(pIPe2.RoundTripTime)))
                DoEvents
                TreeView1.Nodes.Add , , IPAddress_Number, IPAddress_Number & " - (" & pIPe2.RoundTripTime & " ms)"
            Label8.Caption = Label8.Caption + 1
            DoEvents
            Exit Sub
            Else
            vbWSACleanup
            End If
            DoEvents
            KeepGoing = 1
            'RCode
            DoEvents
        Else
        DoEvents
            KeepGoing = 1
            ' RCode
            DoEvents
        End If

    End Sub


Public Sub vbGetHostByName()

On Error Resume Next
    Dim szString As String
    Dim Host As String
    Host = Trim$(IPAddress_Number)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))
    DoEvents

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        MsgBox sMsg
        ExitTheFor = 1
        DoEvents
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory HOSTENT.h_name, ByVal _
        PointerToPointer, Len(HOSTENT) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = HOSTENT.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong2, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong2.Byte4)) + "." + CStr(Asc(IPLong2.Byte3)) _
        + "." + CStr(Asc(IPLong2.Byte2)) + "." + CStr(Asc(IPLong2.Byte1)))
        DoEvents
    End If

End Sub


Public Sub vbGetHostName()

On Error Resume Next
    Host = String(64, &H0)
    


    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        MsgBox sMsg
        ExitTheFor = 1
        DoEvents
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        ' Host
        DoEvents
    End If

End Sub


Public Sub vbIcmpSendEcho()

On Error Resume Next
    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)

        DoEvents
            bReturn = IcmpSendEcho(hIP, addr, szBuffer, Len(szBuffer), pIPo2, pIPe2, Len(pIPe2) + 8, 2700)
           DoEvents
            If bReturn Then
                If KeepGoing = 1 Then KeepGoing = 0: Exit For
                PingTimes = PingTimes + 1
                DoEvents
                RespondingHost = CStr(pIPe2.Address(0)) + "." + CStr(pIPe2.Address(1)) + "." + CStr(pIPe2.Address(2)) + "." + CStr(pIPe2.Address(3))
                DoEvents
                GetRCode
                DoEvents
            Else
               ' "Offline"
            End If
        Next NbrOfPkts

    End Sub


Sub vbWSAStartup()

On Error Resume Next
Dim wsAdata2 As WSADataType
    iReturn = WSAStartup(&H101, wsAdata2)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        MsgBox "WSock32.dll is Not responding!"
        ExitTheFor = 1
    End If


    If LoByte(wsAdata2.wversion) < WS_VERSION_MAJOR Or (LoByte(wsAdata2.wversion) = WS_VERSION_MAJOR And HiByte(wsAdata2.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(wsAdata2.wversion)))
        sLowByte = Trim$(Str$(LoByte(wsAdata2.wversion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        MsgBox sMsg
        ExitTheFor = 1
        End
    End If
DoEvents

    If wsAdata2.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            MsgBox sMsg
            ExitTheFor = 1
        End
    End If
    
    MaxSockets = wsAdata2.iMaxSockets
DoEvents

    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    MaxUDP = wsAdata2.iMaxUdpDg

DoEvents
    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    
    Description = ""


    'For i = 0 To WSADESCRIPTION_LEN
        'If wsadata2.szDescription(i) = 0 Then Exit For
        'Description = Description + Chr$(wsadata2.szDescription(i))
    'Next i
    Status = ""


    'For i = 0 To WSASYS_STATUS_LEN
        'If wsadata2.szSystemStatus(i) = 0 Then Exit For
        'Status = Status + Chr$(wsadata2.szSystemStatus(i))
    'Next i

End Sub


Public Function HiByte(ByVal wParam As Integer)

On Error Resume Next
    HiByte = wParam \ &H100 And &HFF&

End Function


Public Function LoByte(ByVal wParam As Integer)

On Error Resume Next
    LoByte = wParam And &HFF&

End Function


Public Sub vbWSACleanup()

On Error Resume Next
    iReturn = WSACleanup()
End Sub


Public Sub vbIcmpCloseHandle()

On Error Resume Next
    bReturn = IcmpCloseHandle(hIP)

End Sub


Public Sub vbIcmpCreateFile()

On Error Resume Next
    hIP = IcmpCreateFile()

End Sub
