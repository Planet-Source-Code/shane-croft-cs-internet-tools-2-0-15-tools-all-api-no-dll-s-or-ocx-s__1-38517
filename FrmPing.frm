VERSION 5.00
Begin VB.Form FrmPing 
   Caption         =   "Ping"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "FrmPing.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   5985
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   960
      TabIndex        =   6
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdClose 
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
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   330
         Width           =   2055
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "Ping"
         Default         =   -1  'True
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
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox lblPingTimes 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         TabIndex        =   2
         Text            =   "4"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox lblPacketSize 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         TabIndex        =   3
         Text            =   "32"
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3600
         Picture         =   "FrmPing.frx":1982
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblPacket 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Packet:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   9
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblPings 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Ping(s):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblIpHost 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Ip/Host:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   960
         TabIndex        =   7
         Top             =   120
         Width           =   660
      End
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   5775
   End
End
Attribute VB_Name = "FrmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PingTimes As Integer
Dim Speed As Long
Dim IP As String
Dim KeepGoing As Integer
Dim TotalNum As Long
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim HOSTENT As HOSTENT, PointerToPointer As Long, ListAddress As Long
Dim WSADATA As WSADATA, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
Dim ExitTheFor As Integer
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
Dim StopMe As Boolean


Private Sub cmdClose_Click()

StopMe = True
DoEvents
DoEvents
Unload Me

End Sub

Private Sub cmdPing_Click()

On Error Resume Next
If Host.text = "" Then
MsgBox "Please Enter A Ip/Host To Ping.", vbInformation
Exit Sub
End If

If cmdPing.Caption = "Stop" Then
StopMe = True
cmdPing.Caption = "Ping"
cmdClose.Enabled = True
Exit Sub
End If

cmdPing.Caption = "Stop"
cmdClose.Enabled = False

If gethostbyname(Host.text) = 0 Then
txtStatus.text = "Unable To Resolve Host"
cmdPing.Caption = "Ping"
cmdClose.Enabled = True
Exit Sub
End If
    Speed = 0
    PingTimes = 0
    txtStatus = ""
    szBuffer = Space(Val(lblPacketSize))
    vbWSAStartup
    If Len(Host.text) = 0 Then
        vbGetHostName
    End If
    vbGetHostByName
    vbIcmpCreateFile
    pIPo2.TTL = Trim$(255)
    '
    For Times = 1 To lblPingTimes
    If ExitTheFor = 1 Then ExitTheFor = 0: Exit For
    If StopMe = True Then
    StopMe = False
    Exit Sub
    End If
    vbIcmpSendEcho
    Next
    vbIcmpCloseHandle
    vbWSACleanup

    Speed = Speed / PingTimes
    txtStatus = txtStatus & vbCrLf & " Average Speed: " & Speed & "."
    txtStatus.SelStart = Len(txtStatus)
cmdPing.Caption = "Ping"
cmdClose.Enabled = True
    Exit Sub

End Sub

Public Sub GetRCode()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetRCode

RCode = ""
    If pIPe.Status = 0 Then RCode = "Success"
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Destination Unreahable"
    If pIPe.Status = 11003 Then RCode = "Destination Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Destination Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Destination Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Requested Timed Out"
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

        If RCode <> "" Then
            If RCode = "Success" Then
                Speed = Speed + Val(Trim$(CStr(pIPe2.RoundTripTime)))
                txtStatus.text = txtStatus.text + " Reply from " + RespondingHost + ": Bytes = " + Trim$(CStr(pIPe2.DataSize)) + " RTT = " + Trim$(CStr(pIPe2.RoundTripTime)) + "ms TTL = " + Trim$(CStr(pIPe2.Options.TTL)) + vbCrLf
                txtStatus.SelStart = Len(txtStatus)
            Exit Sub
            End If
            KeepGoing = 1
            txtStatus.text = txtStatus.text & RCode
        Else
            KeepGoing = 1
            txtStatus.text = txtStatus.text & RCode
        End If
        txtStatus.SelStart = Len(txtStatus)

EXIT_GetRCode:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GetRCode:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in GetRCode" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_GetRCode
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_GetRCode

    End Sub


Public Sub vbGetHostByName()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbGetHostByName

    Dim szString As String
    Host = Trim$(Host.text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        txtStatus = sMsg
        ExitTheFor = 1
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
    End If

EXIT_vbGetHostByName:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbGetHostByName:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbGetHostByName" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbGetHostByName
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbGetHostByName

End Sub


Public Sub vbGetHostName()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbGetHostName
    
    Host = String(64, &H0)
    


    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        txtStatus = sMsg
        ExitTheFor = 1
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        Host.text = Host
    End If

EXIT_vbGetHostName:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbGetHostName:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbGetHostName" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbGetHostName
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbGetHostName

End Sub


Public Sub vbIcmpSendEcho()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbIcmpSendEcho

    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)

        DoEvents
            bReturn = IcmpSendEcho(hIP, addr, szBuffer, Len(szBuffer), pIPo2, pIPe2, Len(pIPe2) + 8, 2700)
            If bReturn Then
                If KeepGoing = 1 Then KeepGoing = 0: Exit For
                PingTimes = PingTimes + 1
                RespondingHost = CStr(pIPe2.Address(0)) + "." + CStr(pIPe2.Address(1)) + "." + CStr(pIPe2.Address(2)) + "." + CStr(pIPe2.Address(3))
                GetRCode
            Else
                txtStatus.text = txtStatus.text + " Request Timeout" + vbCrLf
                txtStatus.SelStart = Len(txtStatus)
            End If
        Next NbrOfPkts

EXIT_vbIcmpSendEcho:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbIcmpSendEcho:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbIcmpSendEcho" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbIcmpSendEcho
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbIcmpSendEcho

    End Sub


Sub vbWSAStartup()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbWSAStartup

Dim wsAdata2 As WSADataType
    iReturn = WSAStartup(&H101, wsAdata2)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        txtStatus = "WSock32.dll is Not responding!"
        ExitTheFor = 1
    End If


    If LoByte(wsAdata2.wversion) < WS_VERSION_MAJOR Or (LoByte(wsAdata2.wversion) = WS_VERSION_MAJOR And HiByte(wsAdata2.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(wsAdata2.wversion)))
        sLowByte = Trim$(Str$(LoByte(wsAdata2.wversion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        txtStatus = sMsg
        ExitTheFor = 1
        End
    End If


    If wsAdata2.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            txtStatus = sMsg
            ExitTheFor = 1
        End
    End If
    
    MaxSockets = wsAdata2.iMaxSockets


    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    MaxUDP = wsAdata2.iMaxUdpDg


    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    
    Description = ""


    'For i = 0 To WSADESCRIPTION_LEN
    '    If Wsadata2.szDescription(i) = 0 Then Exit For
    '    Description = Description + Chr$(Wsadata2.szDescription(i))
    'Next i
    Status = ""


    'For i = 0 To WSASYS_STATUS_LEN
    '    If Wsadata2.szSystemStatus(i) = 0 Then Exit For
    '    Status = Status + Chr$(Wsadata2.szSystemStatus(i))
    'Next i

EXIT_vbWSAStartup:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbWSAStartup:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbWSAStartup" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbWSAStartup
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbWSAStartup

End Sub


Public Function HiByte(ByVal wParam As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_HiByte

    HiByte = wParam \ &H100 And &HFF&

EXIT_HiByte:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_HiByte:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in HiByte" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_HiByte
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_HiByte

End Function


Public Function LoByte(ByVal wParam As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_LoByte

    LoByte = wParam And &HFF&

EXIT_LoByte:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_LoByte:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in LoByte" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_LoByte
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_LoByte

End Function


Public Sub vbWSACleanup()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbWSACleanup

    iReturn = WSACleanup()

EXIT_vbWSACleanup:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbWSACleanup:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbWSACleanup" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbWSACleanup
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbWSACleanup

End Sub


Public Sub vbIcmpCloseHandle()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbIcmpCloseHandle

    bReturn = IcmpCloseHandle(hIP)

EXIT_vbIcmpCloseHandle:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbIcmpCloseHandle:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbIcmpCloseHandle" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbIcmpCloseHandle
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbIcmpCloseHandle

End Sub


Public Sub vbIcmpCreateFile()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_vbIcmpCreateFile

    hIP = IcmpCreateFile()

EXIT_vbIcmpCreateFile:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_vbIcmpCreateFile:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in vbIcmpCreateFile" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_vbIcmpCreateFile
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_vbIcmpCreateFile

End Sub


Private Sub Form_Load()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

Me.Height = 3870
Me.Width = 6105

StopMe = False
Dim mWSD As WSADataType
lV = WSAStartup(&H202, mWSD)
vbWSAStartup
vbWSACleanup

EXIT_Form_Load:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Form_Load:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Form_Load" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Form_Load
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Form_Load

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_QueryUnload

If cmdPing.Caption = "Stop" Then
Cancel = True
Exit Sub
End If

EXIT_Form_QueryUnload:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Form_QueryUnload:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Form_QueryUnload" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Form_QueryUnload
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Form_QueryUnload

End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Frame1.Top
txtStatus.Move txtStatus.Left, txtStatus.Top, Me.ScaleWidth - 200, Me.ScaleHeight - 1300

End Sub

Private Sub Host_KeyDown(KeyCode As Integer, Shift As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Host_KeyDown

If KeyCode = vbKeyReturn Then
 Call cmdPing_Click
 DoEvents
 End If

EXIT_Host_KeyDown:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Host_KeyDown:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Host_KeyDown" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Host_KeyDown
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Host_KeyDown

End Sub

Private Sub Host_LostFocus()

On Error Resume Next
Host.text = Replace(Host.text, " ", "", 1, , vbTextCompare)

End Sub

Private Sub lblPacketSize_LostFocus()

On Error Resume Next
lblPacketSize.text = Replace(lblPacketSize.text, " ", "", 1, , vbTextCompare)

End Sub

Private Sub lblPingTimes_LostFocus()


On Error Resume Next
lblPingTimes.text = Replace(lblPingTimes.text, " ", "", 1, , vbTextCompare)

End Sub
