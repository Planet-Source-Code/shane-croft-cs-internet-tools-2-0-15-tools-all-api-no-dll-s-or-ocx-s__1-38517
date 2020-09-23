VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBandwidth 
   Caption         =   "Bandwidth Monitor"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBandwidth.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   5040
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "0 KB"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "0 KB"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Estimated Upload Speed"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Estimated Download Speed"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblRecv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblSent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   195
         TabIndex        =   5
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sent"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Received"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   165
         X2              =   2160
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         X1              =   165
         X2              =   2160
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   2160
         X2              =   2160
         Y1              =   240
         Y2              =   480
      End
      Begin VB.Line Line5 
         X1              =   2160
         X2              =   3960
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line7 
         X1              =   165
         X2              =   3960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line8 
         X1              =   165
         X2              =   3960
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line9 
         X1              =   960
         X2              =   960
         Y1              =   480
         Y2              =   960
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Kb"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Kb"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4080
         Picture         =   "FrmBandwidth.frx":1982
         Top             =   615
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4200
         Picture         =   "FrmBandwidth.frx":3304
         Top             =   975
         Width           =   480
      End
      Begin VB.Line Line6 
         X1              =   3960
         X2              =   3960
         Y1              =   480
         Y2              =   960
      End
      Begin VB.Line Line1 
         X1              =   165
         X2              =   165
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1440
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   960
      Top             =   1680
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBandwidth.frx":4C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBandwidth.frx":9A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBandwidth.frx":E88A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBandwidth.frx":1368C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmBandwidth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_objIpHelper As CIpHelper
Private TransferRate                    As Single
Private TransferRate2                   As Single

Private Sub Form_Load()

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_Form_Load

'On Error Resume Next
Set m_objIpHelper = New CIpHelper
Me.Height = 2430
Me.Width = 5130

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
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Me.ScaleHeight / 2 - Frame1.Height / 2

End Sub

Private Sub Timer1_Timer()

On Error Resume Next
Call UpdateInterfaceInfo
End Sub
Private Sub UpdateInterfaceInfo()

On Error Resume Next
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Double
Static lngBytesSent     As Double
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
Case MIB_IF_TYPE_ETHERNET: lblType.Caption = "Ethernet"
Case MIB_IF_TYPE_FDDI: lblType.Caption = "FDDI"
Case MIB_IF_TYPE_LOOPBACK: lblType.Caption = "Loopback"
Case MIB_IF_TYPE_OTHER: lblType.Caption = "Other"
Case MIB_IF_TYPE_PPP: lblType.Caption = "PPP"
Case MIB_IF_TYPE_SLIP: lblType.Caption = "SLIP"
Case MIB_IF_TYPE_TOKENRING: lblType.Caption = "TokenRing"
End Select
lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived / 1024, "###,###,###,###"))
lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent / 1024, "###,###,###,###"))
Set st_objInterface = objInterface
'---------------
blnIsRecv = (m_objIpHelper.BytesReceived / 1024 > lngBytesRecv / 1024)
blnIsSent = (m_objIpHelper.BytesSent / 1024 > lngBytesSent / 1024)
If blnIsRecv And blnIsSent Then
Image1.Picture = ImageList1.ListImages(4).Picture
ElseIf (Not blnIsRecv) And blnIsSent Then
Image1.Picture = ImageList1.ListImages(2).Picture
ElseIf blnIsRecv And (Not blnIsSent) Then
Image1.Picture = ImageList1.ListImages(3).Picture
ElseIf Not (blnIsRecv And blnIsSent) Then
Image1.Picture = ImageList1.ListImages(1).Picture
End If
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent
DoEvents

End Sub

Private Sub Timer2_Timer()

On Error Resume Next
DoEvents
Dim xx As Long
Dim YY As Long
Dim XXX As Long
Dim YYY As Long
YYY = Label6.Caption
YY = Label5.Caption
DoEvents
xx = Me.lblRecv.Caption - YY
XXX = Me.lblSent.Caption - YYY
DoEvents
TransferRate = Format(Int(xx) / 1024, "00.00")
DoEvents
TransferRate2 = Format(Int(XXX) / 1024, "00.00")
DoEvents

                Label10.Caption = TransferRate2 & " Kb"
                DoEvents

                Label9.Caption = TransferRate & " Kb"
                DoEvents
    DoEvents
    Label5.Caption = Me.lblRecv.Caption
    Label6.Caption = Me.lblSent.Caption
    DoEvents

End Sub
