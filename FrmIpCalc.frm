VERSION 5.00
Begin VB.Form FrmIpCalc 
   Caption         =   "IP Calculator"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmIpCalc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   9540
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.Frame Frame5 
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
         Height          =   1695
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   4815
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   12
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Reset"
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Default Mask"
            Height          =   255
            Left            =   3240
            TabIndex        =   11
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Compute &Now"
            Height          =   255
            Left            =   3240
            TabIndex        =   9
            Top             =   1200
            Width           =   1455
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
            ScaleWidth      =   113
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   360
            Width           =   1695
            Begin VB.TextBox txtip 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
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
               TabIndex        =   49
               Top             =   0
               Width           =   60
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   2
               Left            =   765
               TabIndex        =   48
               Top             =   0
               Width           =   60
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   3
               Left            =   1110
               TabIndex        =   47
               Top             =   0
               Width           =   60
            End
         End
         Begin VB.PictureBox picSM 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            MousePointer    =   3  'I-Beam
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   113
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   720
            Width           =   1695
            Begin VB.TextBox txtsm 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   3
               Left            =   1170
               MaxLength       =   3
               TabIndex        =   8
               Text            =   "0"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtsm 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   825
               MaxLength       =   3
               TabIndex        =   7
               Text            =   "255"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtsm 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   120
               MaxLength       =   3
               TabIndex        =   5
               Text            =   "255"
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtsm 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   465
               MaxLength       =   3
               TabIndex        =   6
               Text            =   "255"
               Top             =   0
               Width           =   285
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   4
               Left            =   1110
               TabIndex        =   45
               Top             =   0
               Width           =   60
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   5
               Left            =   750
               TabIndex        =   44
               Top             =   0
               Width           =   60
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "."
               Height          =   300
               Index           =   6
               Left            =   405
               TabIndex        =   43
               Top             =   0
               Width           =   60
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address:"
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
            TabIndex        =   52
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Subnet Mask:"
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
            TabIndex        =   51
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Network ID:"
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
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "- Binary Information -"
         Enabled         =   0   'False
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
         Height          =   1575
         Left            =   120
         TabIndex        =   37
         Top             =   2040
         Width           =   4815
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address:"
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mask:"
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Network ID:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "- Network Information -"
         Enabled         =   0   'False
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
         Height          =   1695
         Left            =   5040
         TabIndex        =   31
         Top             =   240
         Width           =   4335
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Good IP For Host:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Type:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address Class:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "Yes/No"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   33
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Reason"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2400
            TabIndex        =   32
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "- Subnetting Information -"
         Enabled         =   0   'False
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
         Height          =   3135
         Left            =   5040
         TabIndex        =   25
         Top             =   2040
         Width           =   4335
         Begin VB.ListBox List1 
            Enabled         =   0   'False
            Height          =   1185
            Left            =   120
            TabIndex        =   26
            Top             =   1800
            Width           =   4095
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Range:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   960
            TabIndex        =   30
            Top             =   1080
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "# of Hosts:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   960
            TabIndex        =   29
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "# of Subnetworks:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   960
            TabIndex        =   28
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Network ID's                                         Broadcast ID's"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   120
            TabIndex        =   27
            Top             =   1560
            Width           =   4155
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save To File"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   4815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   4800
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2280
         Picture         =   "FrmIpCalc.frx":0E42
         Top             =   3720
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmIpCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BA(7) As Integer
Dim IPType As String
Dim iRange As Integer
Dim iMaskEndPosition As Integer
Dim spacelen As String
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Private Const SHACF_DEFAULT  As Long = &H0

Const GWL_EXSTYLE = (-20)
Private Function ConvertBin(iValue As Integer) As String

  On Error Resume Next
  Dim c As Integer
  Dim TempValue As Integer
  Dim tempdata As String
  If iValue = Null Then ConvertBin = "00000000": Exit Function
  TempValue = 0
  For c = 7 To 0 Step -1
    If TempValue + BA(c) <= iValue Then
      tempdata = tempdata & "1"
      TempValue = TempValue + BA(c)
    Else
      tempdata = tempdata & "0"
    End If
  Next c
  ConvertBin = tempdata

End Function

Private Function GetBinNetID(strip As String, StrSubnetMask As String) As String

  On Error Resume Next
  Dim pos As Integer, tempnetid As String, X As String, Y As String, z As String
  pos = 1
  Do While pos <> Len(strip) + 1
    If Mid(strip, pos, 1) <> "." Then
      X = Mid(strip, pos, 1)
      Y = Mid(StrSubnetMask, pos, 1)
      z = (CInt(X) * CInt(Y))
      tempnetid = tempnetid & z
    Else
      tempnetid = tempnetid & "."
    End If
    pos = pos + 1
  Loop
  GetBinNetID = tempnetid

End Function

Private Function ConvertBinToIP(strBin As String) As String

  On Error Resume Next
  Dim pos As Integer, binarray, tempnetid As String, ix As Integer, X As Integer, Y As Integer, z As String
  strBin = strBin & "."
  binarray = Split(strBin, ".")
  For ix = 0 To UBound(binarray) - 1
    X = 0
    For Y = 7 To 0 Step -1
      If Mid(StrReverse(binarray(ix)), Y + 1, 1) = "1" Then
        X = X + BA(Y)
      Else
        X = X
      End If
    Next Y
    z = z & CStr(X) & "."
  Next ix
  ConvertBinToIP = Left(z, Len(z) - 1)

End Function

Private Function GetIPClass(strip As String) As String

  On Error Resume Next
  Dim tempip, X As Integer
  strip = strip & "."
  tempip = Split(strip, ".")

  Select Case tempip(0)
    Case 0 To 127
      GetIPClass = "A"
      If tempip(0) = 10 Then
        IPType = "Reserved"
        Exit Function
      ElseIf tempip(0) = 127 Then
        IPType = "Loopback"
        GetIPClass = "Loopback"
        Exit Function
      End If
      IPType = "Public"
    Case 128 To 191
      GetIPClass = "B"
      If tempip(0) = 172 Then
        Select Case tempip(1)
          Case 16 To 31
            IPType = "Resreved"
        End Select
        Exit Function
      End If
      IPType = "Public"
    Case 192 To 223
      GetIPClass = "C"
      If tempip(0) = 192 And tempip(1) = 168 Then
        IPType = "Reserved"
        Exit Function
      End If
      IPType = "Public"
    Case 224 To 239
      GetIPClass = "D"
      IPType = "Multicast(RFC 1112)"
    Case 240 To 255
      GetIPClass = "E"
      IPType = "Experemential"
  End Select

End Function
Private Function GetBits(strmask As String) As Single

  On Error Resume Next
  Dim tempdata, ix As Integer, pos As Integer, itemp As Single
  strmask = strmask & "."
  tempdata = Split(strmask, ".")
  For ix = 0 To UBound(tempdata) - 1
    Select Case tempdata(ix)
      Case "255"
        itemp = itemp + 8
      Case "128"
        itemp = itemp + 1
      Case "192"
        itemp = itemp + 2
      Case "224"
        itemp = itemp + 3
      Case "240"
        itemp = itemp + 4
      Case "248"
        itemp = itemp + 5
      Case "252"
        itemp = itemp + 6
      Case "254"
        itemp = itemp + 7
    End Select
  Next ix
  GetBits = itemp

End Function

Private Function GetRange(strmask As String) As Integer

On Error Resume Next
  'uses the couchie method
  Dim tempdata, ix As Integer, itemp As Integer, ipclass As String
  strmask = strmask & "."
  tempdata = Split(strmask, ".")
  For ix = 0 To UBound(tempdata) - 1
    Select Case tempdata(ix)
      Case "128"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "192"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "224"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "240"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "248"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "252"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case "254"
        iRange = 256 - CInt(tempdata(ix))
        Exit For
      Case Else
        iRange = 256
    End Select
  Next ix
  GetRange = iRange

End Function
Private Function GetPosNetworks(strmask As String) As Integer

On Error Resume Next
  Dim tempdata, ix As Integer, itemp As Integer, ipclass As String
  strmask = strmask & "."
  tempdata = Split(strmask, ".")
  For ix = 0 To UBound(tempdata) - 1
    Select Case tempdata(ix)
      Case "128"
        GetPosNetworks = 2 ^ 1
      Case "192"
        GetPosNetworks = 2 ^ 2
      Case "224"
        GetPosNetworks = 2 ^ 3
      Case "240"
        GetPosNetworks = 2 ^ 4
      Case "248"
        GetPosNetworks = 2 ^ 5
      Case "252"
        GetPosNetworks = 2 ^ 6
      Case "254"
        GetPosNetworks = 2 ^ 7
        'Case Else
        'GetPosNetworks = 1
    End Select
  Next ix

End Function
Private Function GetPosHosts(strBinaryMask As String) As String

On Error Resume Next
  Dim tempdata, ix As Integer, itemp As Integer, ipclass As String, pos As Integer, BitCount As Integer
  strBinaryMask = strBinaryMask & "."
  tempdata = Split(strBinaryMask, ".")
  For ix = 0 To UBound(tempdata) - 1
    pos = 1
    Do While pos <= Len(tempdata(ix)) + 1
      If Mid(tempdata(ix), pos, 1) = "0" Then
        BitCount = BitCount + 1
      End If
      pos = pos + 1
    Loop
  Next ix
  'Takes 2 to the BitCount power that is not used by NetID - 2 (for netid and broadcast)
  GetPosHosts = Format((2 ^ BitCount) - 2, "###,###,###,###")

End Function
Private Function GetMaskEndPosition(strmask As String) As Integer

On Error Resume Next
  Dim tmpmask, X As Integer
  strmask = strmask & "."
  tmpmask = Split(strmask, ".")
  For X = 0 To UBound(tmpmask) - 1
    Select Case tmpmask(X)
      Case "255"
        GetMaskEndPosition = X + 1
        iMaskEndPosition = X + 1
      Case Else
        Exit For
    End Select
  Next X

End Function
Private Sub LoadNetID(strNetID As String, strmask As String, ipRange As Integer)

On Error Resume Next
  Dim inet As Integer, ibroad As Integer, X As Integer, ipleft As String, iptemp, imaskend As Integer
  imaskend = GetMaskEndPosition(strmask)
  strNetID = Mid(strNetID, 1, InStrRev(strNetID, "/", Len(strNetID)) - 1)
  strNetID = strNetID & "."
  iptemp = Split(strNetID, ".")
  If ipRange = 0 Then ipRange = 256
  For X = 0 To imaskend - 1
    ipleft = ipleft & iptemp(X) & "."
  Next X
List1.Clear
  If ipRange <> 256 Then
    For X = 0 To 255 Step ipRange
      With Me.List1
        iptemp = Split(ipleft & "x", ".")
        Select Case UBound(iptemp) + 1
          Case 1
            .AddItem ipleft & X & ".0.0.0" & spacelen & ipleft & X + (ipRange - 1) & ".255.255.255"
          Case 2
            .AddItem ipleft & X & ".0.0" & spacelen & ipleft & X + (ipRange - 1) & ".255.255"
          Case 3
            .AddItem ipleft & X & ".0" & spacelen & ipleft & X + (ipRange - 1) & ".255"
          Case 4
            .AddItem ipleft & X & spacelen & ipleft & X + (ipRange - 1)
        End Select
      End With
      DoEvents
    Next X
  Else
    With Me.List1
      iptemp = Split(ipleft & "x", ".")
      Select Case UBound(iptemp) + 1
        Case 1
          .AddItem ipleft & "0.0.0.0" & spacelen & ipleft & (ipRange - 1) & ".255.255.255"
        Case 2
          .AddItem ipleft & "0.0.0" & spacelen & ipleft & (ipRange - 1) & ".255.255"
        Case 3
          .AddItem ipleft & "0.0" & spacelen & ipleft & (ipRange - 1) & ".255"
        Case 4
          .AddItem ipleft & "0" & spacelen & ipleft & (ipRange - 1)
      End Select
    End With
    DoEvents
  End If

End Sub


Private Function IsGoodIP(strip As String) As String

On Error Resume Next
  Dim tempdata, tempinfo As String, X As Integer, Y As Integer, tempgood As String, temptype As String
  Dim tempclass As String
  
  tempclass = GetIPClass(strip)
  strip = Left$(strip, Len(strip) - 1)
  tempgood = ""
  
  If tempclass = "D" Or tempclass = "E" Then
  Text9.text = "No"
  Text10.text = "(Invalid Class)"
  Exit Function
  End If
  
  For X = 0 To Me.List1.ListCount - 1
    tempinfo = Me.List1.List(X) & spacelen
    tempdata = Split(tempinfo, spacelen)
    
    For Y = 0 To UBound(tempdata) - 1
      If strip = tempdata(Y) Then
        tempgood = "No"
        If Y = 0 Then
          temptype = "Network ID"
        Else
          temptype = "Broadcast ID"
        End If
      End If
    Next Y
    If tempgood = "No" Then Exit For
  Next X
  If tempgood = "No" Then
    Text9.text = tempgood
    Text10.text = temptype
  Else
    Text9.text = "Yes"
  End If

End Function
Private Sub HighlightNetworkID(strNetID As String)

On Error Resume Next
  Dim X As Integer, tempdata
  For X = 0 To Me.List1.ListCount - 1
    tempdata = Split(Me.List1.List(X), spacelen)
    If strNetID = tempdata(0) Then
      Me.List1.ListIndex = X
      Exit For
    End If
  Next X

End Sub

Private Sub Command1_Click()

On Error Resume Next
Dim X As Integer
  With FrmIpCalc
    .Frame2.Enabled = False
    .Frame3.Enabled = False
    .Frame4.Enabled = False
    For X = 2 To .lbl.Count - 1
      .lbl(X).Enabled = False
    Next X
    Label1(1).Enabled = False
    Label2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
List1.Enabled = False
  End With

  txtip(0).text = "192"
  txtip(1).text = "168"
  txtip(2).text = "0"
  txtip(3).text = "1"
  txtsm(0).text = "255"
  txtsm(1).text = "255"
  txtsm(2).text = "255"
  txtsm(3).text = "0"
Text3.text = ""
Text4.text = ""
Text5.text = ""
Text6.text = ""
Text7.text = ""
Text8.text = ""
Text9.text = ""
Text10.text = ""
Text11.text = ""
Text12.text = ""
Text13.text = ""
List1.Clear

End Sub

Private Sub Command2_Click()

On Error Resume Next
  txtsm(0).text = "255"
  txtsm(1).text = "255"
  txtsm(2).text = "255"
  txtsm(3).text = "0"

End Sub

Private Sub Command3_Click()

On Error Resume Next
Dim xx As Integer
  With FrmIpCalc
    .Frame2.Enabled = True
    .Frame3.Enabled = True
    .Frame4.Enabled = True
    For xx = 2 To .lbl.Count - 1
      .lbl(xx).Enabled = True
    Next xx
    Label1(1).Enabled = True
    Label2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
List1.Enabled = True
  End With
Text3.text = ""
Text4.text = ""
Text5.text = ""
Text6.text = ""
Text7.text = ""
Text8.text = ""
Text9.text = ""
Text10.text = ""
Text11.text = ""
Text12.text = ""
Text13.text = ""
List1.Clear
  Dim X As Integer, tempsm As String, temprange As Integer, tempip As String
  Dim tindex As Integer
  'Hide Our Tip If its Visible

  'hold our ip and mask for later use
  tempsm = txtsm(0).text & "." & txtsm(1).text & "." & txtsm(2).text & "." & txtsm(3).text
  tempip = txtip(0).text & "." & txtip(1).text & "." & txtip(2).text & "." & txtip(3).text
  'Check the mask incase it wasnt checked before
  For tindex = 0 To txtsm.Count - 1
    If checkmask(tindex) = False Then
      MsgBox "Number for mask must be:" & Chr(13) & "0, 128, 192, 224, 240, 248, 252, 254, or 255" & Chr(13) & "Please reenter number.", vbApplicationModal + vbDefaultButton1 + vbInformation, "Error"
      txtsm(tindex).SetFocus
      'highlight the section
      SendKeys "{HOME}+{END}"
      Exit Sub
    End If
  Next tindex
  'Set the binary IP And Mask Labels
  For X = 0 To txtip.Count - 1
    Text4.text = Text4.text & "." & ConvertBin(CInt(txtip(X).text))
    Text5.text = Text5.text & "." & ConvertBin(CInt(txtsm(X).text))
  Next X
  Text4.text = Mid(Text4.text, 2)
  Text5.text = Mid(Text5.text, 2)
  'Set the binary Network ID Label
  Text6.text = GetBinNetID(Text4.text, Text5.text)
  'Set the Network ID by converting the Binary Network ID
  Text3.text = ConvertBinToIP(Text6.text)
  'Set the IP Class Label
  Text7.text = GetIPClass(Text3.text)
  'Set the type of IP Label (Public, Private, Loopback, Multicast, or Experimental)
  Text8.text = IPType
  'Get the range
  temprange = GetRange(tempsm)
  'If the range = 256 then the range = 1
  If temprange = 256 Then temprange = 1
  'Set the Range Label
  Text13.text = temprange
  'Add the CIDR to the Network ID
  Text3.text = Text3.text & "/" & GetBits(tempsm)
  'Set the Total Number of networks label
  Text11.text = GetPosNetworks(tempsm)
  'Set the total number of hosts allowed per network label
  Text12.text = GetPosHosts(Text5.text)
  'Load all Network ID's and Broadcast ID's to the list
  LoadNetID Text3.text, tempsm, iRange
  'Set If IP Can be assigned to a host
  Call IsGoodIP(tempip)
  
  'Heighlight the network in the listbox
  HighlightNetworkID (Mid(Text3.text, 1, InStr(1, Text3.text, "/") - 1))
  'Return Focus to The IP Text Box
  txtip(0).SetFocus

End Sub

Private Sub Command4_Click()

On Error Resume Next
List1.ListIndex = 0
DoEvents
FrmSaveIpCalc.Show
FrmSaveIpCalc.SetFocus
DoEvents
FrmSaveIpCalc.List1.Clear
DoEvents
FrmSaveIpCalc.List1.AddItem "- IP Information -"
FrmSaveIpCalc.List1.AddItem vbTab & "IP Address: " & txtip(0).text & "." & txtip(1).text & "." & txtip(2).text & "." & txtip(3).text
FrmSaveIpCalc.List1.AddItem vbTab & "Subnet Mask: " & txtsm(0).text & "." & txtsm(1).text & "." & txtsm(2).text & "." & txtsm(3).text
FrmSaveIpCalc.List1.AddItem vbTab & "Network ID: " & Text3.text
FrmSaveIpCalc.List1.AddItem ""
FrmSaveIpCalc.List1.AddItem "- Binary Information -"
FrmSaveIpCalc.List1.AddItem vbTab & "IP Address: " & Text4.text
FrmSaveIpCalc.List1.AddItem vbTab & "Subnet Mask: " & Text5.text
FrmSaveIpCalc.List1.AddItem vbTab & "Network ID: " & Text6.text
FrmSaveIpCalc.List1.AddItem ""
FrmSaveIpCalc.List1.AddItem "- Network Information -"
FrmSaveIpCalc.List1.AddItem vbTab & "IP Address Class: " & Text7.text
FrmSaveIpCalc.List1.AddItem vbTab & "Address Type: " & Text8.text
FrmSaveIpCalc.List1.AddItem vbTab & "Good IP For Host: " & Text9.text & " " & Text10.text
FrmSaveIpCalc.List1.AddItem ""
FrmSaveIpCalc.List1.AddItem "- Subnetting Information -"
FrmSaveIpCalc.List1.AddItem vbTab & "# of Subnetworks: " & Text11.text
FrmSaveIpCalc.List1.AddItem vbTab & "# of Hosts: " & Text12.text
FrmSaveIpCalc.List1.AddItem vbTab & "Range: " & Text13.text
FrmSaveIpCalc.List1.AddItem ""
FrmSaveIpCalc.List1.AddItem vbTab & "Network ID's " & vbTab & vbTab & "Broadcast ID's"
FrmSaveIpCalc.List1.AddItem ""
Do Until Me.List1.ListIndex = Me.List1.ListCount - 1
FrmSaveIpCalc.List1.AddItem vbTab & List1.text
List1.ListIndex = List1.ListIndex + 1
Loop
FrmSaveIpCalc.List1.AddItem vbTab & List1.text
FrmSaveIpCalc.List1.AddItem ""
FrmSaveIpCalc.List1.AddItem ""
DoEvents

End Sub

Private Sub Command5_Click()

Unload Me

End Sub

Private Sub Form_Load()

  On Error Resume Next
  Dim c As Integer
  spacelen = vbTab & vbTab
  BA(0) = 1
  For c = 1 To UBound(BA)
    BA(c) = 2 ^ c
  Next c
  
Me.Height = 5715
Me.Width = 9630

End Sub

Private Function checkmask(Index As Integer) As Boolean

On Error Resume Next
  'this returns true if the Mask is a valid mask
  If CInt(txtsm(Index).text) <> 255 And CInt(txtsm(Index).text) <> 0 And CInt(txtsm(Index).text) <> 128 And _
      CInt(txtsm(Index).text) <> 224 And CInt(txtsm(Index).text) <> 240 And CInt(txtsm(Index).text) <> 248 And _
      CInt(txtsm(Index).text) <> 252 And CInt(txtsm(Index).text) <> 254 Then
    checkmask = False
  Else
    checkmask = True
  End If

End Function

Private Sub Form_Resize()
On Error Resume Next
Frame1.Move Me.ScaleWidth / 2 - Frame1.Width / 2, Me.ScaleHeight / 2 - Frame1.Height / 2
End Sub

Private Sub txtip_Change(Index As Integer)

  On Error Resume Next
  'If the section = "" we need to put a value there
  If txtip(Index) = "" Then txtip(Index) = "0": SendKeys "{HOME}+{END}"
  'Now we need to set a range of numbers allowed.
  If CInt(txtip(Index).text) > 255 Then
    MsgBox "Number must be between 0 - 255." & Chr(13) & "Please reenter number.", vbApplicationModal + vbDefaultButton1 + vbInformation, "Error"
    SendKeys "{HOME}+{END}"
  End If
  If Len(txtip(Index).text) = 3 Then
    If Index = txtip.Count - 1 Then
      txtsm(0).SetFocus
    Else
      txtip(Index + 1).SetFocus
    End If
  End If

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
      txtsm(tindex).SetFocus
    Else
      tindex = Index + 1
      txtip(tindex).SetFocus
    End If
  End If

End Sub
Private Sub txtsm_Change(Index As Integer)
  On Error Resume Next
  If txtsm(Index) = "" Then txtsm(Index) = "0": SendKeys "{HOME}+{END}"
  If Len(txtsm(Index).text) = 3 Then
    If Index = txtsm.Count - 1 Then
      txtip(0).SetFocus
    Else
      txtsm(Index + 1).SetFocus
    End If
  End If

End Sub

Private Sub txtsm_Click(Index As Integer)

On Error Resume Next
  SendKeys "{HOME}+{END}"

End Sub

Private Sub txtsm_GotFocus(Index As Integer)

On Error Resume Next
  SendKeys "{HOME}+{END}"

End Sub

Private Sub txtsm_KeyPress(Index As Integer, KeyAscii As Integer)

  On Error Resume Next
  Dim tindex As Integer
  If KeyAscii = Asc(".") Or KeyAscii = 13 Or KeyAscii = Asc(vbTab) Then
    If checkmask(Index) = True Then
      If Index = txtsm.Count - 1 Then
        tindex = 0
      Else
        tindex = Index + 1
      End If
    Else
      tindex = Index
      MsgBox "Number for mask must be:" & Chr(13) & "0, 128, 192, 224, 240, 248, 252, 254, or 255" & Chr(13) & "Please reenter number.", vbApplicationModal + vbDefaultButton1 + vbInformation, "Error"
    End If
    txtsm(tindex).SetFocus
    SendKeys "{HOME}+{END}"
  End If

End Sub
