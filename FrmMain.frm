VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "CS Internet Tools"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10800
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   720
      Top             =   600
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   2400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1982
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":362E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":6952
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":8674
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A006
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B998
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":D32A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":ECBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1064E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":11FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":13972
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":15304
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":16C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":18628
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1947A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":19794
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1B126
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1BA00
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1C2DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1DA6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1DD86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1F518
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2431A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2911C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2DF1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":32D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":38512
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3DD04
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3E716
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3F128
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3FB3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4054C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   4  'Align Right
      Height          =   6540
      Left            =   10230
      TabIndex        =   1
      Top             =   0
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   11536
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "TCP/IP Configuration"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Time Sync"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Trace Route"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Whois && MX Lookup"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stats"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Check For Update"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Web Page"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About CS Internet Tools"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   6540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   11536
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bandwidth Monitor"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "IP Address Scanner"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "IP Calculator"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "IP Converter"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "NetStat"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Online - Offline Checker"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ping"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Port Listener"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Port Scanner"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Resolve Host && IP"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   6540
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
            MinWidth        =   9948
            Text            =   "© Crofts Software - Networking Software & More"
            TextSave        =   "© Crofts Software - Networking Software & More"
            Object.ToolTipText     =   "© Crofts Software - Networking Software && More"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   609
            MinWidth        =   617
            Object.ToolTipText     =   "Bandwidth Meter"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuView 
      Caption         =   "&View"
      Begin VB.Menu MenuToolsbars 
         Caption         =   "Toolbars"
         Begin VB.Menu menuviewright 
            Caption         =   "Right ToolBar"
            Checked         =   -1  'True
         End
         Begin VB.Menu MenuViewLeft 
            Caption         =   "Left Toolbar"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSettings 
         Caption         =   "Settings..."
      End
   End
   Begin VB.Menu MenuTools 
      Caption         =   "&Tools"
      Begin VB.Menu MenuBandwidth 
         Caption         =   "Bandwidth Monitor"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenIpScan 
         Caption         =   "IP Address Scanner"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu MenuIP 
         Caption         =   "IP Calculator"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu MenuIPConvert 
         Caption         =   "IP Converter"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu MenuNetstat 
         Caption         =   "NetStat"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuOnOffLine 
         Caption         =   "Online - Offline Checker"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuPing 
         Caption         =   "Ping"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MeunPortListener 
         Caption         =   "Port Listener"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuPortScan 
         Caption         =   "Port Scanner"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuResolve 
         Caption         =   "Resolve Host && IP"
         Shortcut        =   {F7}
      End
      Begin VB.Menu menuIPConfig 
         Caption         =   "TCP/IP Configuration"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu MenuTime 
         Caption         =   "Time Sync"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MenuTrace 
         Caption         =   "Trace Route"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MwnuWhois 
         Caption         =   "Whois && MX Lookup"
         Shortcut        =   {F11}
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuStats 
         Caption         =   "Stats..."
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MenuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCheck 
         Caption         =   "Check For Update..."
      End
      Begin VB.Menu linea 
         Caption         =   "-"
      End
      Begin VB.Menu MenuWeb 
         Caption         =   "Web Page..."
      End
      Begin VB.Menu MenuBug 
         Caption         =   "Bug Report..."
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private FilePathName As String

Private m_objIpHelper As CIpHelper

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub MDIForm_Load()

On Error Resume Next


FrmSettings.Show
DoEvents
FrmSettings.Hide
DoEvents

AppDir = App.Path

FilePathName = AppDir + "\ViewSettings.inf"
maintoolbar1 = GetPrivateProfileString("settings", "maintoolbar1", "", FilePathName)
maintoolbar2 = GetPrivateProfileString("settings", "maintoolbar2", "", FilePathName)

DoEvents
Me.menuviewright.Checked = maintoolbar1
Me.MenuViewLeft.Checked = maintoolbar2
DoEvents

Set m_objIpHelper = New CIpHelper

StatusBar1.Panels(2).text = Me.Winsock1.LocalHostName
StatusBar1.Panels(3).text = Me.Winsock1.LocalIP
StatusBar1.Panels(2).ToolTipText = "Current Local Computer Name"
StatusBar1.Panels(3).ToolTipText = "Current Local IP Address"
StatusBar1.Panels(4).Picture = ImageList1.ListImages(4).Picture
DoEvents

If FrmSettings.Check1.Value = 1 Then
Me.WindowState = 2
End If
If FrmSettings.Check4.Value = 0 Then
Me.Timer1.Enabled = False
End If

If Me.menuviewright.Checked = True Then
Me.Toolbar1.Visible = True
Else
Me.Toolbar1.Visible = False
End If

If Me.MenuViewLeft.Checked = True Then
Me.Toolbar2.Visible = True
Else
Me.Toolbar2.Visible = False
End If

 Dim wavSetup As String
 If FrmSettings.Check2.Value = 1 Then
 wavSetup = NoiseGet(App.Path & "\Sounds\" & "startup.wav")
 NoisePlay wavSetup, SND_SYNC
 End If

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

Dim fFile As Integer
fFile = FreeFile
 
Open App.Path & "\ViewSettings.inf" For Output As fFile
Print #fFile, "[settings]"
Print #fFile, "maintoolbar1=" & Me.menuviewright.Checked
Print #fFile, "maintoolbar2=" & Me.MenuViewLeft.Checked
Close fFile
DoEvents
DoEvents

 Dim wavSetup As String
 If FrmSettings.Check3.Value = 1 Then
 wavSetup = NoiseGet(App.Path & "\Sounds\" & "shutdown.wav")
 NoisePlay wavSetup, SND_SYNC
  DoEvents
  DoEvents
End
Else
End
End If

End Sub

Private Sub MenIpScan_Click()

On Error Resume Next
FrmIpScan.Show
FrmIpScan.SetFocus

End Sub

Private Sub MenuAbout_Click()
On Error Resume Next
frmAbout.Show
frmAbout.SetFocus

End Sub

Private Sub MenuBandwidth_Click()

On Error Resume Next
FrmBandwidth.Show
FrmBandwidth.SetFocus

End Sub

Private Sub MenuBug_Click()

On Error Resume Next
Call ShellExecute(hwnd, "Open", "mailto:webmaster@croftssoftware.com", "", App.Path, 1)

End Sub

Private Sub MenuCheck_Click()

On Error Resume Next
FrmUpdate.Show
FrmUpdate.SetFocus

End Sub

Private Sub MenuExit_Click()

On Error Resume Next
Unload Me

End Sub

Private Sub menuip_Click()


On Error Resume Next
FrmIpCalc.Show
FrmIpCalc.SetFocus

End Sub

Private Sub menuIPConfig_Click()

On Error Resume Next
FrmIpConfig.Show
FrmIpConfig.SetFocus

End Sub

Private Sub MenuIPConvert_Click()

On Error Resume Next
FrmIpConvert.Show
FrmIpConvert.SetFocus

End Sub

Private Sub MenuNetstat_Click()

On Error Resume Next
FrmNetStat.Show
FrmNetStat.SetFocus

End Sub

Private Sub MenuOnOffLine_Click()

On Error Resume Next
FrmOnline.Show
FrmOnline.SetFocus

End Sub

Private Sub MenuPing_Click()

On Error Resume Next
FrmPing.Show
FrmPing.SetFocus

End Sub

Private Sub MenuPortScan_Click()
On Error Resume Next
FrmPortScanner.Show
FrmPortScanner.SetFocus

End Sub

Private Sub MenuResolve_Click()

On Error Resume Next
FrmResolve.Show
FrmResolve.SetFocus

End Sub

Private Sub MenuSettings_Click()
On Error Resume Next
FrmSettings.Show
FrmSettings.SetFocus

End Sub

Private Sub MenuStats_Click()

On Error Resume Next
FrmStats.Show
FrmStats.SetFocus

End Sub

Private Sub MenuTime_Click()

FrmTime.Show
FrmTime.SetFocus

End Sub

Private Sub MenuTrace_Click()

FrmTraceMenu.Show
FrmTraceMenu.SetFocus

End Sub

Private Sub MenuViewLeft_Click()

If Me.MenuViewLeft.Checked = True Then
Me.MenuViewLeft.Checked = False
Me.Toolbar2.Visible = False
Else
Me.MenuViewLeft.Checked = True
Me.Toolbar2.Visible = True
End If

End Sub

Private Sub menuviewright_Click()

If Me.menuviewright.Checked = True Then
Me.menuviewright.Checked = False
Me.Toolbar1.Visible = False
Else
Me.menuviewright.Checked = True
Me.Toolbar1.Visible = True
End If

End Sub

Private Sub MenuWeb_Click()

On Error Resume Next
Call ShellExecute(hwnd, "Open", "http://www.croftssoftware.com", "", App.Path, 1)

End Sub

Private Sub MeunPortListener_Click()

FrmPortListen.Show
FrmPortListen.SetFocus

End Sub

Private Sub MwnuWhois_Click()

FrmWhois.Show
FrmWhois.SetFocus

End Sub

Private Sub Timer1_Timer()

Call UpdateInterfaceInfo

End Sub
Private Sub UpdateInterfaceInfo()

On Error Resume Next
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)

Set st_objInterface = objInterface
'---------------
blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
If blnIsRecv And blnIsSent Then
StatusBar1.Panels(4).Picture = ImageList1.ListImages(1).Picture
ElseIf (Not blnIsRecv) And blnIsSent Then
StatusBar1.Panels(4).Picture = ImageList1.ListImages(3).Picture
ElseIf blnIsRecv And (Not blnIsSent) Then
StatusBar1.Panels(4).Picture = ImageList1.ListImages(2).Picture
ElseIf Not (blnIsRecv And blnIsSent) Then
StatusBar1.Panels(4).Picture = ImageList1.ListImages(4).Picture
End If
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent
'StatusBar1.Panels(4).ToolTipText = "Bytes received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) & "  Bytes sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))

End Sub
Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetPrivateProfileString

   Dim szTmp                     As String
   Dim nRet                      As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)

EXIT_GetPrivateProfileString:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_GetPrivateProfileString:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in GetPrivateProfileString" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_GetPrivateProfileString
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_GetPrivateProfileString

End Function
Private Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   ' *** Get an entry in the WIN inifile ***

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_GetProfileString

   Dim szTmp                    As String
   Dim nRet                     As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)

EXIT_GetProfileString:
   Exit Function

   ' #VBIDEUtilsERROR#
ERROR_GetProfileString:
   Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in GetProfileString" & vbCrLf & "The error occured at line: " & Erl, vbAbortRetryIgnore + vbCritical, "Error")
      Case vbAbort
         Screen.MousePointer = vbDefault
         Resume EXIT_GetProfileString
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
   Resume EXIT_GetProfileString

End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next

Select Case Button.Index

Case 1
FrmBandwidth.Show
FrmBandwidth.SetFocus
Case 2
FrmIpScan.Show
FrmIpScan.SetFocus
Case 3
FrmIpCalc.Show
FrmIpCalc.SetFocus
Case 4
FrmIpConvert.Show
FrmIpConvert.SetFocus
Case 5
FrmNetStat.Show
FrmNetStat.SetFocus
Case 6
FrmOnline.Show
FrmOnline.SetFocus
Case 7
FrmPing.Show
FrmPing.SetFocus
Case 8
FrmPortListen.Show
FrmPortListen.SetFocus
Case 9
FrmPortScanner.Show
FrmPortScanner.SetFocus
Case 10
FrmResolve.Show
FrmResolve.SetFocus
End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next

Select Case Button.Index

Case 1
FrmIpConfig.Show
FrmIpConfig.SetFocus
Case 2
FrmTime.Show
FrmTime.SetFocus
Case 3
FrmTraceMenu.Show
FrmTraceMenu.SetFocus
Case 4
FrmWhois.Show
FrmWhois.SetFocus
Case 5
FrmStats.Show
FrmStats.SetFocus
Case 6
FrmSettings.Show
FrmSettings.SetFocus
Case 7
FrmUpdate.Show
FrmUpdate.SetFocus
Case 8
Call ShellExecute(hwnd, "Open", "http://www.croftssoftware.com", "", App.Path, 1)
Case 9
frmAbout.Show
frmAbout.SetFocus
Case 10
Unload Me
End Select

End Sub

