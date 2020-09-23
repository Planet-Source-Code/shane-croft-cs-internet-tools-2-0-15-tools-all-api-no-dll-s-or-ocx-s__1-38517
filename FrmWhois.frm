VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmWhois 
   Caption         =   "Whois & MX Lookup"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmWhois.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   6345
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtSearch 
         BackColor       =   &H80000014&
         Height          =   270
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "lookup"
         Height          =   285
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbDNS 
         Height          =   345
         Left            =   3600
         TabIndex        =   2
         Text            =   "0.0.0.0"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "www."
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " DNS Server For MX Lookup To Use"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.ListBox lstMX 
      Height          =   2310
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox txtResponse 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8705
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Whois"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MX Record"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmWhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DNS_RECURSION As Byte = 1

Private Type DNS_HEADER
    qryID As Integer
    Options As Byte
    response As Byte
    qdcount As Integer
    ancount As Integer
    nscount As Integer
    arcount As Integer
End Type

' Registry data types
Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

' Registry access types
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

' Registry keys
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

' The only registry error that I care about =)
Const ERROR_SUCCESS = 0&

' Registry access functions
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

' Variant (string array) that holds all the DNS servers found in the registry
Dim sDNS As Variant
Private Sub Command4_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Command4_Click

If txtSearch.text = "" Then
MsgBox "Please Enter A Domain", vbCritical
Exit Sub
End If
   MousePointer = vbHourglass
   txtResponse = ""
   lstMX.Clear
   Winsock1.Close
   Winsock1.LocalPort = 0
   DoEvents
   If Right(txtSearch, 3) = ".tr" Then
      Winsock1.connect "whois.metu.edu.tr", 43
   Else
      Winsock1.connect "whois.networksolutions.com", 43
   End If

    Dim sMX As String
    
    If (cmbDNS.text <> "") Then
        lstMX.AddItem "Mail routing information for " & txtSearch
        lstMX.AddItem "     using DNS server of " & cmbDNS.text
        sMX = MX_Query
        If (Len(sMX) > 0) Then
            lstMX.AddItem "Best MX record to send through: " & sMX
        Else
            lstMX.AddItem "No mail routing information found"
        End If
    Else
        MsgBox "ERROR: Can not find DNS information for MX Lookup"
    End If

EXIT_Command4_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Command4_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Command4_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Command4_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Command4_Click
    
End Sub

Private Sub Form_Load()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

Me.Height = 6615
Me.Width = 6465

    GetDNSInfo
    
    If (cmbDNS.ListCount > 0) Then cmbDNS.ListIndex = 0
 Exit Sub

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
TabStrip1.Move 100, TabStrip1.Top, Me.ScaleWidth - 200, Me.ScaleHeight - 1275
txtResponse.Move 200, txtResponse.Top, Me.ScaleWidth - 400, Me.ScaleHeight - 1850
lstMX.Move 200, lstMX.Top, Me.ScaleWidth - 400, Me.ScaleHeight - 1850

End Sub

Private Sub TabStrip1_Click()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_TabStrip1_Click

If TabStrip1.Tabs(2).Selected = True Then
Me.txtResponse.Visible = False
Me.lstMX.Visible = True
End If

If TabStrip1.Tabs(1).Selected = True Then
Me.txtResponse.Visible = True
Me.lstMX.Visible = False
End If

EXIT_TabStrip1_Click:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_TabStrip1_Click:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in TabStrip1_Click" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_TabStrip1_Click
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_TabStrip1_Click

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtSearch_KeyDown

If KeyCode = vbKeyReturn Then
 Call Command4_Click
 DoEvents
 End If

EXIT_txtSearch_KeyDown:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtSearch_KeyDown:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtSearch_KeyDown" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtSearch_KeyDown
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtSearch_KeyDown

End Sub

Private Sub txtSearch_LostFocus()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_txtSearch_LostFocus

On Error Resume Next
txtSearch.text = Replace(txtSearch.text, " ", "", 1, , vbTextCompare)

EXIT_txtSearch_LostFocus:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_txtSearch_LostFocus:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in txtSearch_LostFocus" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_txtSearch_LostFocus
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_txtSearch_LostFocus

End Sub

Private Sub Winsock1_Connect()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Winsock1_Connect

    Winsock1.SendData txtSearch & vbCrLf

EXIT_Winsock1_Connect:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Winsock1_Connect:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Winsock1_Connect" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Winsock1_Connect
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Winsock1_Connect

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Winsock1_DataArrival

    Dim strdata As String

    On Error Resume Next

    Winsock1.GetData strdata
    strdata = Replace(strdata, Chr$(10), vbCrLf)
    txtResponse = txtResponse & strdata
    MousePointer = vbDefault

EXIT_Winsock1_DataArrival:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_Winsock1_DataArrival:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Winsock1_DataArrival" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_Winsock1_DataArrival
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_Winsock1_DataArrival

End Sub
Private Sub Form_Unload(Cancel As Integer)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Unload

   Unload Me

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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  StripTerminator
'''''''''''
''' Remove the NULL character from the end of a string
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function StripTerminator(ByVal strString As String) As String

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_StripTerminator

    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If

EXIT_StripTerminator:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_StripTerminator:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in StripTerminator" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_StripTerminator
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_StripTerminator

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  GetDNSInfo
'''''''''''
''' Read the registry to find all the DNS servers (DHCP and configured)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub GetDNSInfo()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetDNSInfo

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
    Dim Buffer2 As IP_ADAPTER_INFO
    Dim FixedInfoBuffer() As Byte
    Dim AdapterInfoBuffer() As Byte
    
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
            Me.cmbDNS.AddItem FixedInfo.DnsServerList.IpAddress
            pAddrStr = FixedInfo.DnsServerList.Next
            Do While pAddrStr <> 0
                  CopyMemory buffer, ByVal pAddrStr, Len(buffer)
                  Me.cmbDNS.AddItem buffer.IpAddress
                  pAddrStr = buffer.Next
            Loop
            

    Else
            Exit Sub
    End If

EXIT_GetDNSInfo:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_GetDNSInfo:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in GetDNSInfo" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_GetDNSInfo
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_GetDNSInfo

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  ParseName
'''''''''''
''' Parse the server name out of the MX record, returns it in variable sName, iNdx is also
''' modified to point to the end of the parsed structure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, sName As String)

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_ParseName

    Dim iCompress As Integer        ' Compression index (index into original buffer)
    Dim iChCount As Integer         ' Character count (number of chars to read from buffer)
        
    ' While we didn't encounter a null char (end-of-string specifier)
    While (dnsReply(iNdx) <> 0)
        ' Read the next character in the stream (length specifier)
        iChCount = dnsReply(iNdx)
        ' If our length specifier is 192 (0xc0) we have a compressed string
        If (iChCount = 192) Then
            ' Read the location of the rest of the string (offset into buffer)
            iCompress = dnsReply(iNdx + 1)
            ' Call ourself again, this time with the offset of the compressed string
            ParseName dnsReply(), iCompress, sName
            ' Step over the compression indicator and compression index
            iNdx = iNdx + 2
            ' After a compressed string, we are done
            Exit Sub
        End If
        
        ' Move to next char
        iNdx = iNdx + 1
        ' While we should still be reading chars
        While (iChCount)
            ' add the char to our string
            sName = sName + Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
        Wend
        ' If the next char isn't null then the string continues, so add the dot
        If (dnsReply(iNdx) <> 0) Then sName = sName + "."
    Wend

EXIT_ParseName:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_ParseName:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in ParseName" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_ParseName
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_ParseName

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  GetMXName
'''''''''''
''' Parses the buffer returned by the DNS server, returns the best MX server (lowest preference
''' number), iNdx is modified to point to current buffer position (should be the end of buffer
''' by the end, unless a record other than MX is found)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As String

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetMXName

    Dim iChCount As Integer     ' Character counter
    Dim sTemp As String         ' Holds original query string
    
    Dim iBestPref As Integer    ' Holds the "best" preference number (lowest)
    Dim sBestMX As String       ' Holds the "best" MX record (the one with the lowest preference)
    
    iBestPref = -1
    
    ParseName dnsReply(), iNdx, sTemp
    ' Step over null
    iNdx = iNdx + 2
    
    ' Step over 6 bytes (not sure what the 6 bytes are, but all other
    '   documentation shows steping over these 6 bytes)
    iNdx = iNdx + 6
    
    While (iAnCount)
        ' Check to make sure we received an MX record
        If (dnsReply(iNdx) = 15) Then
            Dim sName As String
            Dim iPref As Integer
            
            sName = ""
            
            ' Step over the last half of the integer that specifies the record type (1 byte)
            ' Step over the RR Type, RR Class, TTL (3 integers - 6 bytes)
            iNdx = iNdx + 1 + 6
            
            ' Read the MX data length specifier
            '              (not needed, hence why it's commented out)
            'MemCopy iMXLen, dnsReply(iNdx), 2
            'iMXLen = ntohs(iMXLen)
            
            ' Step over the MX data length specifier (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            MemCopy iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            ' Step over the MX preference value (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            ' Have to step through the byte-stream, looking for 0xc0 or 192 (compression char)
            ParseName dnsReply(), iNdx, sName
            lstMX.AddItem "[Preference = " & iPref & "] " & sName
            
            If (iBestPref = -1 Or iPref < iBestPref) Then
                iBestPref = iPref
                sBestMX = sName
            End If
            
            ' Step over 3 useless bytes
            iNdx = iNdx + 3
        Else
            GetMXName = sBestMX
            Exit Function
        End If
        iAnCount = iAnCount - 1
    Wend
    
    GetMXName = sBestMX

EXIT_GetMXName:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_GetMXName:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in GetMXName" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_GetMXName
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_GetMXName

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  MakeQName
'''''''''''
''' Takes sDomain and converts it to the QNAME-type string, returns that. QNAME is how a
''' DNS server expects the string.
'''
'''    Ex...    Pass -        mail.com
'''             Returns -     &H4mail&H3com
'''                            ^      ^
'''                            |______|____ These two are character counters, they count the
'''                                         number of characters appearing after them
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function MakeQName(sDomain As String) As String

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_MakeQName

    Dim iQCount As Integer      ' Character count (between dots)
    Dim iNdx As Integer         ' Index into sDomain string
    Dim iCount As Integer       ' Total chars in sDomain string
    Dim sQName As String        ' QNAME string
    Dim sDotName As String      ' Temp string for chars between dots
    Dim sChar As String         ' Single char from sDomain string
    
    iNdx = 1
    iQCount = 0
    iCount = Len(sDomain)
    ' While we haven't hit end-of-string
    While (iNdx <= iCount)
        ' Read a single char from our domain
        sChar = Mid(sDomain, iNdx, 1)
        ' If the char is a dot, then put our character count and the part of the string
        If (sChar = ".") Then
            sQName = sQName & Chr(iQCount) & sDotName
            iQCount = 0
            sDotName = ""
        Else
            sDotName = sDotName + sChar
            iQCount = iQCount + 1
        End If
        iNdx = iNdx + 1
    Wend
    
    sQName = sQName & Chr(iQCount) & sDotName
    
    MakeQName = sQName

EXIT_MakeQName:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_MakeQName:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MakeQName" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_MakeQName
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_MakeQName

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''  MakeQName
'''''''''''
''' Performs the actual IP work to contact the DNS server, calls the other functions to parse
''' and return the best server to send email through
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function MX_Query() As String

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_MX_Query

    Dim StartupData As WSADataType
    Dim SocketBuffer As sockaddr
    Dim IpAddr As Long
    Dim iRC As Integer
    Dim dnsHead As DNS_HEADER
    Dim iSock As Integer

    ' Initialize the Winsocket
    iRC = WSAStartup(&H101, StartupData)
    iRC = WSAStartup(&H101, StartupData)
    If iRC = SOCKET_ERROR Then Exit Function
    
    ' Create a socket
    iSock = socket(AF_INET, SOCK_DGRAM, 0)
    If iSock = SOCKET_ERROR Then Exit Function
    
    IpAddr = GetHostByNameAlias(cmbDNS.text)
    If IpAddr = -1 Then Exit Function
    
    ' Setup the connnection parameters
    SocketBuffer.sin_family = AF_INET
    SocketBuffer.sin_port = htons(53)
    SocketBuffer.sin_addr = IpAddr
    SocketBuffer.sin_zero = String$(8, 0)
    
    ' Set the DNS parameters
    dnsHead.qryID = htons(&H11DF)
    dnsHead.Options = DNS_RECURSION
    dnsHead.qdcount = htons(1)
    dnsHead.ancount = 0
    dnsHead.nscount = 0
    dnsHead.arcount = 0
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    '' Variables
    ''''''''''''''''''''''''''''''''''''''''''''''
    Dim dnsQuery() As Byte
    Dim sQName As String
    Dim dnsQueryNdx As Integer
    Dim itemp As Integer
    Dim iNdx As Integer
    dnsQueryNdx = 0
    
    ReDim dnsQuery(4000)
    
    ' Setup the dns structure to send the query in
    
    
    ' First goes the DNS header information
    MemCopy dnsQuery(dnsQueryNdx), dnsHead, 12
    dnsQueryNdx = dnsQueryNdx + 12
    
    ' Then the domain name (as a QNAME)
    sQName = MakeQName(txtSearch)
    iNdx = 0
    While (iNdx < Len(sQName))
        dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
        iNdx = iNdx + 1
    Wend

    dnsQueryNdx = dnsQueryNdx + Len(sQName)
    
    ' Null terminate the string
    dnsQuery(dnsQueryNdx) = &H0
    dnsQueryNdx = dnsQueryNdx + 1
    
    ' The type of query (15 means MX query)
    itemp = htons(15)
    MemCopy dnsQuery(dnsQueryNdx), itemp, Len(itemp)
    dnsQueryNdx = dnsQueryNdx + Len(itemp)
    
    ' The class of query (1 means INET)
    itemp = htons(1)
    MemCopy dnsQuery(dnsQueryNdx), itemp, Len(itemp)
    dnsQueryNdx = dnsQueryNdx + Len(itemp)
    
    ReDim Preserve dnsQuery(dnsQueryNdx - 1)
    ' Send the query to the DNS server
    iRC = sendto(iSock, dnsQuery(0), dnsQueryNdx + 1, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Then
        MsgBox "Problem sending"
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    '' Variables
    ''''''''''''''''''''''''''''''''''''''''''''''
    Dim dnsReply(2048) As Byte
    ' Wait for answer from the DNS server
    iRC = recvfrom(iSock, dnsReply(0), 2048, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Then
        MsgBox "Problem receiving"
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    '' Variables
    ''''''''''''''''''''''''''''''''''''''''''''''
    Dim iAnCount As Integer
    ' Get the number of answers
    MemCopy iAnCount, dnsReply(6), 2
    iAnCount = ntohs(iAnCount)
    ' Parse the answer buffer
    MX_Query = GetMXName(dnsReply(), 12, iAnCount)

EXIT_MX_Query:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_MX_Query:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MX_Query" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_MX_Query
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_MX_Query

End Function
