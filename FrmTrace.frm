VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmTrace 
   Caption         =   "Trace Route With API"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTrace.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   7815
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7575
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
         Height          =   300
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton Close 
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
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton TraceRT2 
         Caption         =   "Trace Route"
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
         Left            =   3495
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Resolve Ip To Host When Finished."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4215
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Resolve"
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
         Left            =   3495
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Left            =   4815
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblIPHost 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "IP/Host:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
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
         Left            =   600
         TabIndex        =   10
         Top             =   600
         Width           =   210
      End
      Begin VB.Label IP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   2475
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6615
         Picture         =   "FrmTrace.frx":1982
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   960
         Width           =   3975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6960
      Top             =   120
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Double Click To Port Scan Selected IP"
      Top             =   1440
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hop"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Host"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time (ms)"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "FrmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TotalNum As Long
Dim KeepGoing As Integer
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim HOSTENT As HOSTENT, PointerToPointer As Long, ListAddress As Long
Dim WSADATA As WSADATA, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
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
Private mWSData As WSADataType

Public Sub GetRCode()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetRCode

If Me.Host.text = "" Then
Me.Host.SetFocus
Exit Sub
End If
RCode = ""
DoEvents
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
'    If pIPe.Status = 11013 Then RCode = "TTL Exprd In Transit"
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
'    RCode = RCode + " (" + CStr(pIPe.Status) + ")"


    DoEvents

        If RCode <> "" Then
            If RCode = "Reqested Timed Out" Then
                'vbWSAStartup
                vbWSACleanup
                Set Item = ListView1.ListItems.Add(, , " # " & Format(TotalNum, "00"))
                Item.SubItems(1) = RCode
                Item.SubItems(3) = pIPe.RoundTripTime
                Item.EnsureVisible
            Exit Sub
            End If
            If RCode = "Success" Then
                'vbWSAStartup
                vbWSACleanup
                Set Item = ListView1.ListItems.Add(, , " # " & Format(TotalNum, "00"))
                Item.SubItems(1) = IP
                Item.SubItems(3) = pIPe.RoundTripTime
                Item.EnsureVisible
            Exit Sub
            End If
            KeepGoing = 1
            Set Item = ListView1.ListItems.Add(, , " # " & Format(TotalNum, "00"))
            Item.SubItems(1) = RCode
            Item.SubItems(3) = pIPe.RoundTripTime
            Item.EnsureVisible
        Else
            If TTL - 1 < 10 Then
            Set Item = ListView1.ListItems.Add(, , " # " & Format(TotalNum, "00"))
            Item.SubItems(1) = RespondingHost
            Item.SubItems(3) = pIPe.RoundTripTime
            Item.EnsureVisible
            Else
            Set Item = ListView1.ListItems.Add(, , " # " & Format(TotalNum, "00"))
            Item.SubItems(1) = RespondingHost
            Item.SubItems(3) = pIPe.RoundTripTime
            Item.EnsureVisible
            End If
        End If

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
        MsgBox sMsg, 0, ""
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory HOSTENT.h_name, ByVal _
        PointerToPointer, Len(HOSTENT) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = HOSTENT.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory addr, ByVal ListAddr, 4
        IP.Caption = Trim$(CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) _
        + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1)))
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
        MsgBox sMsg, 0, ""
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

    'vbWSACleanup
    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)


        DoEvents
        'vbWSACleanup
            bReturn = IcmpSendEcho(hIP, addr, szBuffer, Len(szBuffer), pIPo, pIPe, Len(pIPe) + 8, 2700)
            If bReturn Then
                TotalNum = TotalNum + 1
                RespondingHost = CStr(pIPe.Address(0)) + "." + CStr(pIPe.Address(1)) + "." + CStr(pIPe.Address(2)) + "." + CStr(pIPe.Address(3))
                GetRCode
            Else
                TotalNum = TotalNum + 1
                    GetRCode
                    TTL = TTL + 1
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

Private Sub Close_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Close_Click

Unload Me

EXIT_Close_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Close_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Close_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Close_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Close_Click

End Sub

Private Sub Command1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command1_Click

On Error Resume Next
Command2.Enabled = False
vbWSAStartup
Do Until ListView1.SelectedItem.Index = ListView1.ListItems.Count
If Me.Host.text = "" Then
Me.Host.SetFocus
Exit Sub
End If
ListView1.SelectedItem.EnsureVisible
DoEvents
ListView1.SelectedItem.SubItems(2) = "Working..."
DoEvents
' The inet_addr function returns a long value
    Dim lInteAdd As Long
' pointer to the HOSTENT
    Dim lPointtoHost As Long
' host name we are looking for
    Dim sHost As String
' Hostent
    Dim mHost As HOSTENT
' IP Address
    Dim sIP As String

    sIP = Trim$(ListView1.SelectedItem.SubItems(1))
Label1.Caption = "Resolving " & ListView1.SelectedItem.SubItems(1) & " To Host"
DoEvents
' Convert the IP address
    lInteAdd = inet_addr(sIP)

' if the wrong IP format was entered there is an err generated
    If lInteAdd = INADDR_NONE Then

        'WSErrHandle (Err.LastDllError)
ListView1.SelectedItem.SubItems(2) = "Unable To Resolve"
DoEvents
    Else

' pointer to the Host
        lPointtoHost = gethostbyaddr(lInteAdd, 4, PF_INET)

' if zero is returned then there was an error
        If lPointtoHost = 0 Then

            'WSErrHandle (Err.LastDllError)
ListView1.SelectedItem.SubItems(2) = "Unable To Resolve"
DoEvents
        Else

            RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

            sHost = String(256, 0)

' Copy the host name
            RtlMoveMemory ByVal sHost, ByVal mHost.h_name, 256

' Cut the chr(0) character off
            sHost = Left(sHost, InStr(1, sHost, Chr(0)) - 1)

' Return the host name
            ListView1.SelectedItem.SubItems(2) = sHost
            DoEvents

        End If

    End If
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
Loop
If Me.Host.text = "" Then
Me.Host.SetFocus
Exit Sub
End If
ListView1.SelectedItem.EnsureVisible
DoEvents
ListView1.SelectedItem.SubItems(2) = "Working..."
DoEvents
    sIP = Trim$(ListView1.SelectedItem.SubItems(1))
Label1.Caption = "Resolving " & ListView1.SelectedItem.SubItems(1) & " To Host"
DoEvents
' Convert the IP address
    lInteAdd = inet_addr(sIP)

' if the wrong IP format was entered there is an err generated
    If lInteAdd = INADDR_NONE Then

        'WSErrHandle (Err.LastDllError)
ListView1.SelectedItem.SubItems(2) = "Unable To Resolve"
DoEvents
    Else

' pointer to the Host
        lPointtoHost = gethostbyaddr(lInteAdd, 4, PF_INET)

' if zero is returned then there was an error
        If lPointtoHost = 0 Then

            'WSErrHandle (Err.LastDllError)
ListView1.SelectedItem.SubItems(2) = "Unable To Resolve"
DoEvents
        Else

            RtlMoveMemory mHost, ByVal lPointtoHost, LenB(mHost)

            sHost = String(256, 0)

' Copy the host name
            RtlMoveMemory ByVal sHost, ByVal mHost.h_name, 256

' Cut the chr(0) character off
            sHost = Left(sHost, InStr(1, sHost, Chr(0)) - 1)

' Return the host name
            ListView1.SelectedItem.SubItems(2) = sHost
            DoEvents

        End If

    End If
Label1.Caption = "Resolving IP To Host Is Complete."
Command2.Enabled = True
vbWSACleanup

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

On Error Resume Next
Dim X As Long
X = Me.ListView1.SelectedItem.Index - 1
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index - X)
DoEvents
FrmSaveTrace.Show
FrmSaveTrace.SetFocus
DoEvents
FrmSaveTrace.List1.Clear
DoEvents
FrmSaveTrace.List1.AddItem "Address Traced: " & FrmTrace.Host.text
FrmSaveTrace.List1.AddItem ""
FrmSaveTrace.List1.AddItem "Total Hops: " & FrmTrace.ListView1.ListItems.Count
FrmSaveTrace.List1.AddItem ""
DoEvents
Do Until ListView1.SelectedItem.Index = ListView1.ListItems.Count
FrmSaveTrace.List1.AddItem "Hop: " & FrmTrace.ListView1.SelectedItem.text
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  IP: " & FrmTrace.ListView1.SelectedItem.SubItems(1)
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  Host: " & FrmTrace.ListView1.SelectedItem.SubItems(2)
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  Time (ms): " & FrmTrace.ListView1.SelectedItem.SubItems(3)
DoEvents
ListView1.SelectedItem = ListView1.ListItems(ListView1.SelectedItem.Index + 1)
DoEvents
Loop
FrmSaveTrace.List1.AddItem "Hop: " & FrmTrace.ListView1.SelectedItem.text
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  IP: " & FrmTrace.ListView1.SelectedItem.SubItems(1)
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  Host: " & FrmTrace.ListView1.SelectedItem.SubItems(2)
DoEvents
FrmSaveTrace.List1.AddItem vbTab & "  Time (ms): " & FrmTrace.ListView1.SelectedItem.SubItems(3)
DoEvents

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

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

Me.Height = 4110
Me.Width = 7935

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

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Frame1.Top
ListView1.Move ListView1.Left, ListView1.Top, Me.ScaleWidth - 200, Me.ScaleHeight - 1500
ListView1.ColumnHeaders(3).Width = Me.ListView1.Width - ListView1.ColumnHeaders(1).Width - ListView1.ColumnHeaders(2).Width - ListView1.ColumnHeaders(4).Width - 350

End Sub

Private Sub Form_Unload(Cancel As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Unload

vbWSACleanup

EXIT_Form_Unload:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Form_Unload:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Form_Unload" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Form_Unload
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Form_Unload

End Sub

Private Sub Host_LostFocus()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Host_LostFocus

On Error Resume Next
Host.text = Replace(Host.text, " ", "", 1, , vbTextCompare)

EXIT_Host_LostFocus:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Host_LostFocus:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Host_LostFocus" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Host_LostFocus
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Host_LostFocus

End Sub

Private Sub ListView1_DblClick()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ListView1_DblClick

On Error Resume Next
If ListView1.ListItems.Count = 0 Then
Exit Sub
End If

FrmPortScanner.Show
FrmPortScanner.SetFocus
FrmPortScanner.cmdStop_Click
DoEvents
FrmPortScanner.cmdClearList_Click
DoEvents
FrmPortScanner.txtip = Me.ListView1.SelectedItem.SubItems(1)
DoEvents
FrmPortScanner.SetFocus

EXIT_ListView1_DblClick:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ListView1_DblClick:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in ListView1_DblClick" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_ListView1_DblClick
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_ListView1_DblClick

End Sub

Private Sub Timer1_Timer()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Timer1_Timer

If Host.text = "" Then
TraceRT2.Enabled = False
Else
TraceRT2.Enabled = True
End If

If ListView1.ListItems.Count = 0 Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If

EXIT_Timer1_Timer:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Timer1_Timer:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Timer1_Timer" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Timer1_Timer
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Timer1_Timer

End Sub

Private Sub TraceRT2_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_TraceRT2_Click

If Me.Host.text = "" Then
MsgBox "Please enter a domain name.", vbCritical
Me.Host.SetFocus
Exit Sub
End If

Command1.Enabled = False
Command2.Enabled = False
TotalNum = 0
    szBuffer = Space(32)
    ListView1.ListItems.Clear
vbWSAStartup
vbWSACleanup

    If Len(Host.text) = 0 Then
        vbGetHostName
    End If
    vbGetHostByName
    vbIcmpCreateFile
    ' The following determines the TTL of th
    '     e ICMPEcho for TRACE function
    TraceRT = True
    Label1.Caption = "Tracing Route To " + IP.Caption

    For TTL = 1 To 255
If Me.Host.text = "" Then
Me.Host.SetFocus
Exit Sub
End If
        If KeepGoing = 1 Then
        KeepGoing = 0
        Exit For
        End If
        pIPo.TTL = TTL
        DoEvents
        vbIcmpSendEcho


        DoEvents

            If RespondingHost = IP.Caption Then
                Label1.Caption = "Trace Route has Completed"
                Exit For ' Stop TraceRT
            End If
        Next TTL
        TraceRT = False
        vbIcmpCloseHandle
        vbWSAStartup
        vbWSACleanup
DoEvents
DoEvents
Command1.Enabled = True
Command2.Enabled = True
If Check1.Value = 1 Then
Command1_Click
End If

EXIT_TraceRT2_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_TraceRT2_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in TraceRT2_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_TraceRT2_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_TraceRT2_Click

End Sub

