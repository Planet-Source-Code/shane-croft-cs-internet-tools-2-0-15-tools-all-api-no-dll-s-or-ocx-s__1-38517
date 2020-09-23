VERSION 5.00
Begin VB.Form FrmIpConfig 
   Caption         =   "TCP/IP Configuration"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmIpConfig.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   6105
   Begin VB.Frame Frame3 
      Height          =   7815
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame2 
         Caption         =   "- Fixed Information -"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2655
         Left            =   120
         TabIndex        =   33
         Top             =   5040
         Width           =   5790
         Begin VB.TextBox Text17 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox Text16 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox Text15 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1080
            Width           =   3135
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            TabIndex        =   13
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   3135
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "FrmIpConfig.frx":030A
            Left            =   2160
            List            =   "FrmIpConfig.frx":030C
            TabIndex        =   15
            Top             =   1440
            Width           =   3465
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "NetBIOS Resolution Uses DNS?"
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
            Left            =   1080
            TabIndex        =   40
            Tag             =   "NOCLEAR"
            Top             =   2280
            Width           =   2730
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "WINS Proxy Enabled?"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3000
            TabIndex        =   39
            Tag             =   "NOCLEAR"
            Top             =   1920
            Width           =   1890
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "IP Routing Enabled?"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   38
            Tag             =   "NOCLEAR"
            Top             =   1920
            Width           =   1890
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "NetBIOS Scope ID"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   600
            TabIndex        =   37
            Tag             =   "NOCLEAR"
            Top             =   1080
            Width           =   1530
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Node Type"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            TabIndex        =   36
            Tag             =   "NOCLEAR"
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "DNS Servers"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            TabIndex        =   35
            Tag             =   "NOCLEAR"
            Top             =   1440
            Width           =   1290
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Host Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   825
            TabIndex        =   34
            Tag             =   "NOCLEAR"
            Top             =   375
            Width           =   1290
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "- Adapter Information -"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4740
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   5790
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   2160
            Width           =   3135
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1440
            Width           =   3135
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   4320
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "FrmIpConfig.frx":030E
            Left            =   2175
            List            =   "FrmIpConfig.frx":0310
            TabIndex        =   0
            Top             =   300
            Width           =   3465
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1080
            Width           =   3135
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2520
            Width           =   3135
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   2880
            Width           =   3135
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   3240
            Width           =   3135
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   3600
            Width           =   3135
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   3960
            Width           =   3135
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Adapter Type"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Tag             =   "NOCLEAR"
            Top             =   750
            Width           =   1965
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "DHCP Enabled?"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   720
            TabIndex        =   31
            Tag             =   "NOCLEAR"
            Top             =   4320
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Adapter"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   30
            Tag             =   "NOCLEAR"
            Top             =   360
            Width           =   1965
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Adapter Address"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   29
            Tag             =   "NOCLEAR"
            Top             =   1110
            Width           =   1965
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Lease Expires"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   450
            TabIndex        =   28
            Tag             =   "NOCLEAR"
            Top             =   4005
            Width           =   1665
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Lease Obtained"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   450
            TabIndex        =   27
            Tag             =   "NOCLEAR"
            Top             =   3630
            Width           =   1665
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Secondary WINS Server"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   75
            TabIndex        =   26
            Tag             =   "NOCLEAR"
            Top             =   3270
            Width           =   2040
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Primary WINS Server"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   75
            TabIndex        =   25
            Tag             =   "NOCLEAR"
            Top             =   2910
            Width           =   2040
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "DHCP Server"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   825
            TabIndex        =   24
            Tag             =   "NOCLEAR"
            Top             =   2550
            Width           =   1290
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Default Gateway"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   225
            TabIndex        =   23
            Tag             =   "NOCLEAR"
            Top             =   2190
            Width           =   1890
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Subnet Mask"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   825
            TabIndex        =   22
            Tag             =   "NOCLEAR"
            Top             =   1830
            Width           =   1290
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "IP Address"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   21
            Tag             =   "NOCLEAR"
            Top             =   1470
            Width           =   1965
         End
      End
   End
End
Attribute VB_Name = "FrmIpConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim error As Long
    Dim FixedInfoSize As Long
    Dim AdapterInfoSize As Long
    Dim i As Integer
    Dim PhysicalAddress  As String
    Dim NewTime As Date
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim Adapt As IP_ADAPTER_INFO
    Dim AddrStr As IP_ADDR_STRING
    Dim FixedInfo As FIXED_INFO
    Dim buffer As IP_ADDR_STRING
    Dim pAddrStr As Long
    Dim pAdapt As Long
    Dim Buffer1 As IP_ADAPTER_INFO
    Dim Buffer2 As IP_ADAPTER_INFO
    Dim FixedInfoBuffer() As Byte
    Dim AdapterInfoBuffer() As Byte
Private Sub Cleanup()

On Error Resume Next
Err.Clear 'the error object was the largest culpret  :)
'then clean up the other stuff just for shits and giggles
error = 0
FixedInfoSize = 0
AdapterInfoSize = 0
i = 0
PhysicalAddress = vbNullString
pAddrStr = 0
pAdapt = 0

End Sub
Private Sub Combo1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Combo1_Click

    AdapterInfoSize = 0
    error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           Exit Sub
        End If
    End If
   ReDim AdapterInfoBuffer(AdapterInfoSize - 1)
 
 ' Get actual adapter information
   error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
   If error <> 0 Then
      Exit Sub
   End If


CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)
AdapterInfo.Index = 1
pAdapt = AdapterInfo.Next
   
                For X = 1 To Combo1.ListIndex
                If pAdapt <> 0 Then
                CopyMemory AdapterInfo, ByVal pAdapt, Len(AdapterInfo)
                End If
                Next
                
                CopyMemory Buffer2, AdapterInfo, Len(Buffer2)
                
          Select Case Buffer2.Type
                Case MIB_IF_TYPE_ETHERNET
                    Text12.text = "Ethernet adapter "
                Case MIB_IF_TYPE_TOKENRING
                    Text12.text = "Token Ring adapter "
                Case MIB_IF_TYPE_FDDI
                    Text12.text = "FDDI adapter "
                Case MIB_IF_TYPE_PPP
                    Text12.text = "PPP adapter"
                Case MIB_IF_TYPE_LOOPBACK
                    Text12.text = "Loopback adapter "
                Case MIB_IF_TYPE_SLIP
                    Text12.text = "Slip adapter "
                Case Else
                    Text12.text = "Other adapter "
        End Select

    For i = 0 To Buffer2.AddressLength - 1
           PhysicalAddress = PhysicalAddress & Hex(Buffer2.Address(i))
            If i < Buffer2.AddressLength - 1 Then
             PhysicalAddress = PhysicalAddress & "-"
            End If

    Next
    Text1.text = PhysicalAddress
    If Buffer2.DhcpEnabled Then
            Text11.text = "Yes" 'Enabled
    Else
            Text11.text = "No" 'Disabled
    End If

    pAddrStr = Buffer2.IpAddressList.Next
    Do While pAddrStr >= 0
           CopyMemory buffer, Buffer2.IpAddressList, LenB(buffer)
           Text2.text = buffer.IpAddress
           Text3.text = buffer.IpMask
           pAddrStr = buffer.Next
           If pAddrStr <> 0 Then
            CopyMemory Buffer2.IpAddressList, ByVal pAddrStr, Len(Buffer2.IpAddressList)
           End If
           If pAddrStr = 0 Then
           Exit Do
           End If
   Loop
    Text4.text = Buffer2.GatewayList.IpAddress
    pAddrStr = Buffer2.GatewayList.Next
    
    'Do While pAddrStr <> 0
    '        CopyMemory Buffer, Buffer2.GatewayList, Len(Buffer)
    '        Form1.List1.AddItem "IP Address: " & Buffer.IpAddress
    '        pAddrStr = Buffer.Next
    '        If pAddrStr <> 0 Then
    '        CopyMemory Buffer2.GatewayList, ByVal pAddrStr, Len(Buffer2.GatewayList)
    '        End If
    'Loop

    Text5.text = Buffer2.DhcpServer.IpAddress
    Text6.text = Buffer2.PrimaryWinsServer.IpAddress
    Text7.text = Buffer2.SecondaryWinsServer.IpAddress

    ' Display time
    'this is the tricky part, the .leaseobtaine and expires returns a Long number
    'which is in sec. you then need to add those sec. with a dateadd function.
    ' I think I got the date and time right that the sec needed to be added to.
    ' check your built in ip configuration and compare it. adjust the time as needed.
    
       If Buffer2.LeaseObtained = 0 Then
       Text8.text = ""
       Text9.text = ""
       Else
       NewTime = DateAdd("s", Buffer2.LeaseObtained, #12/31/1969 4:00:00 PM#)
       Text8.text = CStr(Format(NewTime, "ddd, mmm dd, yyyy hh:mm:ss am/pm"))
     
       NewTime = DateAdd("s", Buffer2.LeaseExpires, #12/31/1969 4:00:00 PM#)
       Text9.text = CStr(Format(NewTime, "ddd, mmm dd, yyyy hh:mm:ss am/pm"))
       End If
    'pAdapt = Buffer2.Next
    'If pAdapt <> 0 Then
    '    CopyMemory AdapterInfo, ByVal pAdapt, Len(AdapterInfo)
    'End If
    Cleanup

EXIT_Combo1_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Combo1_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Combo1_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Combo1_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Combo1_Click

End Sub
Private Sub Form_Load()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

Me.Height = 8415
Me.Width = 6225
DoEvents
    'Get the main IP configuration information for this machine using a FIXED_INFO structure
    FixedInfoSize = 0
    error = GetNetworkParams(ByVal 0&, FixedInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           Exit Sub
        End If
    End If
    ReDim FixedInfoBuffer(FixedInfoSize - 1)
    

    error = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
    If error = 0 Then
            CopyMemory FixedInfo, FixedInfoBuffer(0), Len(FixedInfo)
            Text10.text = FixedInfo.HostName
            Combo2.AddItem FixedInfo.DnsServerList.IpAddress
            pAddrStr = FixedInfo.DnsServerList.Next
            Do While pAddrStr <> 0
                  CopyMemory buffer, ByVal pAddrStr, Len(buffer)
                  Combo2.AddItem buffer.IpAddress
                  pAddrStr = buffer.Next
                  Combo2.ListIndex = 0
            Loop
            Select Case FixedInfo.NodeType
                       Case 1
                                  Text13.text = "Broadcast"
                       Case 2
                                  Text13.text = "Peer to peer"
                       Case 4
                                  Text13.text = "Mixed"
                       Case 8
                                  Text13.text = "Hybrid"
                       Case Else
                                  Text13.text = "Unknown node type"
            End Select
            
            Text14.text = FixedInfo.ScopeId
            
            If FixedInfo.EnableRouting Then
                       Text15.text = "Yes"
            Else
                       Text15.text = "No"
            End If
            If FixedInfo.EnableProxy Then
                       Text16.text = "Yes"
            Else
                       Text16.text = "No"
            End If
            If FixedInfo.EnableDns Then
                      Text17.text = "Yes"
            Else
                      Text17.text = "No"
            End If
End If

    AdapterInfoSize = 0
    error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           Exit Sub
        End If
    End If
   ReDim AdapterInfoBuffer(AdapterInfoSize - 1)
 
 ' Get actual adapter information
   error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
   If error <> 0 Then
      Exit Sub
   End If
   CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)
   pAdapt = AdapterInfo.Next

            Do While pAdapt >= 0
                  CopyMemory Buffer1, AdapterInfo, Len(Buffer1)
                  Combo1.AddItem Buffer1.Description
                  pAdapt = Buffer1.Next
                  
                If pAdapt <> 0 Then
                CopyMemory AdapterInfo, ByVal pAdapt, Len(AdapterInfo)
                End If
                If pAdapt = 0 Then
                Exit Do
                End If
            Loop
    Combo1.ListIndex = 0
    Cleanup

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
Frame3.Move Me.ScaleWidth / 2 - Frame3.Width / 2, Me.ScaleHeight / 2 - Frame3.Height / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
'make it a habbit to always manually Clean up your App objects when Closing the last form.
Cleanup
Unload Me

End Sub

