VERSION 5.00
Begin VB.Form FrmIpConvert 
   Caption         =   "IP Converter - Dot IP to Long / Long IP To Dot"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmIpConvert.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   10110
   Begin VB.Frame Frame4 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame1 
         Caption         =   " - Binary -"
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
         Height          =   2295
         Left            =   5160
         TabIndex        =   20
         Top             =   240
         Width           =   4815
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   480
            Width           =   4590
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1560
            Width           =   4590
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "- Dot IP -"
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
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "MSByte           To               LSByte"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   150
            TabIndex        =   23
            Top             =   840
            Width           =   4575
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "- Long IP -"
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
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   4575
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "MSByte           To               LSByte"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   90
            TabIndex        =   21
            Top             =   1920
            Width           =   4575
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   2295
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4935
         Begin VB.CommandButton Command2 
            Caption         =   "Dot IP Address To ""Long"""
            Height          =   315
            Left            =   2160
            TabIndex        =   2
            Top             =   480
            Width           =   2400
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   360
            TabIndex        =   1
            Top             =   480
            Width           =   1440
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   3
            Top             =   1440
            Width           =   1440
         End
         Begin VB.CommandButton Command1 
            Caption         =   """Long"" To Dot IP Address"
            Height          =   315
            Left            =   2160
            TabIndex        =   4
            Top             =   1440
            Width           =   2400
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Example - 192.168.0.1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   9
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   1965
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Example - 16,820,416"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   8
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "- Network Information -"
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
         TabIndex        =   11
         Top             =   2640
         Width           =   4935
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   3225
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   480
            Width           =   1440
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   480
            Width           =   1245
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Of /"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   3270
            TabIndex        =   16
            Top             =   1095
            Width           =   1395
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Of /"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   1485
            TabIndex        =   15
            Top             =   1095
            Width           =   1260
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Local Host Number"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   3240
            TabIndex        =   14
            Top             =   840
            Width           =   1620
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Network Number"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   1485
            TabIndex        =   13
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Network Class"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1200
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   3720
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   7320
         Picture         =   "FrmIpConvert.frx":08CA
         Top             =   2880
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmIpConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Nuttin
Public Function NetIpLong_ToDot(ByVal InAddr As String) As String

''Input InAddr, An unsigned Long INet Address String
''Uses Currency Data Type To Hack Out The Four Byte Vals
''Returns Byte.Byte.Byte.Byte "Dotted Decimal" Notation, IP Address String

On Error GoTo BadIn

StarUp1@ = CCur(InAddr)              '''convert to currency

'''4,294,967,263                  '''Largest Posible Valid address

 If StarUp1@ > 4294967263# Then '''''''Added MsgBox For This Prog Only  REM REM
  Mnsg$ = "4,294,967,263 Is The Largest Posible Valid Address"
  MsgBox Mnsg$, 0
 End If ''''''''''''''''''''''''''''''''''''''''''''''''''''''          REM REM

If StarUp1@ < 0 Then                 '''convert to pos. 32 bit unsigned
StarUp1@ = StarUp1@ + 4294967296#
End If

CkekkerCurre@ = CCur(4294967295#)    ''range check val

If StarUp1@ > CkekkerCurre@ Then     ''check range, within 32 bit
StarUp1@ = 0                         ''if bad, set = 0
End If

Rett1@ = CCur(Fix(StarUp1@ / 16777216))                  '''Get High Order Byte

StarUp2@ = CCur(Round(StarUp1@ - (Rett1@ * 16777216)))
Rett2@ = CCur(Fix(StarUp2@ / 65536))                     '''Get Next Byte


StarUp3@ = CCur(Round(StarUp2@ - (Rett2@ * 65536)))
Rett3@ = CCur(Fix(StarUp3@ / 256))                       '''Get Next Byte


Rett4@ = CCur(Round(StarUp3@ - (Rett3@ * 256)))          '''Remainder, Low Order Byte


NetIpLong_ToDot = Trim$(Str$(Rett4@)) & "." & Trim$(Str$(Rett3@)) & "." & Trim$(Str$(Rett2@)) & "." & Trim$(Str$(Rett1@))


GoTo GoodIn
BadIn:
Resume 10
10
NetIpLong_ToDot = "0.0.0.0"          '''If Error, Returns This

GoodIn:

'''''PrntNetSpecs, Below, Added For This Prog Only'''''''''''

PrntNetSpecs Rett1@, Rett2@, Rett3@, Rett4@

End Function
Private Sub Command1_Click()

On Error Resume Next
Text2.text = NetIpLong_ToDot(Trim$(Text1.text))

End Sub


Private Sub Command2_Click()

On Error Resume Next
Ding$ = DotTo_NetIpLong(Text2.text)

Text1.text = Format$(Ding$, "###,###,###,###,##0")

End Sub



Public Function DotTo_NetIpLong(ByVal InDot As String) As String

On Error Resume Next
''' Input "Dotted Decimal" Notation IP Address String
''' Returns 32 Bit Unsigned Decimal Address As A String Of Digit Characters

Ased1& = InStr(1, InDot, ".", vbTextCompare)

Byt1Lo@ = CCur(Val(Mid$(InDot, 1, Ased1& - 1)))

If Byt1Lo@ > 255 Then
Byt1Lo@ = 255
End If

 If Byt1Lo@ > 223 Then '''''''Added Msg PopFor This Prog Only      REM REM
  Mnsg$ = "223 Is The Largest Posible LSByte In A Valid Address"
  MsgBox Mnsg$, 0
 End If ''''''''''''''''''''''''''''''''''''''''''''''''''''''     REM REM

If Byt1Lo@ < 0 Then
Byt1Lo@ = 0
End If


Ased2& = InStr(Ased1& + 1, InDot, ".", vbTextCompare)

Byt2@ = CCur(Val(Mid$(InDot, Ased1& + 1, (Ased2& - Ased1&) - 1)))

If Byt2@ > 255 Then
Byt2@ = 255
End If

If Byt2@ < 0 Then
Byt2@ = 0
End If

Byt2a@ = Byt2@               '''REM REM this Line Used With PrntNetSpecs Sub
Byt2@ = Byt2@ * 256


Ased3& = InStr(Ased2& + 1, InDot, ".", vbTextCompare)

Byt3@ = CCur(Val(Mid$(InDot, Ased2& + 1, (Ased3& - Ased2&) - 1)))

If Byt3@ > 255 Then
Byt3@ = 255
End If

If Byt3@ < 0 Then
Byt3@ = 0
End If

Byt3a@ = Byt3@               '''REM REM this Line Used With PrntNetSpecs Sub
Byt3@ = Byt3@ * 65536


Byt4Hi@ = CCur(Val(Mid$(InDot, Ased3& + 1)))

If Byt4Hi@ > 255 Then
Byt4Hi@ = 255
End If

If Byt4Hi@ < 0 Then
Byt4Hi@ = 0
End If

Byt4aHi@ = Byt4Hi@               '''REM REM this Line Used With PrntNetSpecs Sub
Byt4Hi@ = Byt4Hi@ * 16777216


AddrssLon@ = Byt1Lo@ + Byt2@ + Byt3@ + Byt4Hi@   ''The Value To Be Returned

DotTo_NetIpLong = Trim$(Str$(AddrssLon@))       ''Converted To A String And Returned


'''''PrntNetSpecs, Below, Added For This Prog Only

PrntNetSpecs Byt4aHi@, Byt3a@, Byt2a@, Byt1Lo@


End Function


Public Function ByteToBinary(ByVal Byytte As String) As String

On Error Resume Next
'''0000 0001

Byt22& = CLng(Byytte)

NxBit8& = (&H80 And Byt22&) / 128

NxBit7& = (&H40 And Byt22&) / 64

NxBit6& = (&H20 And Byt22&) / 32

NxBit5& = (&H10 And Byt22&) / 16

NxBit4& = (&H8 And Byt22&) / 8

NxBit3& = (&H4 And Byt22&) / 4

NxBit2& = (&H2 And Byt22&) / 2

NxBit1& = (&H1 And Byt22&)


ByteToBinary = Trim$(Str$(NxBit8&)) & Trim$(Str$(NxBit7&)) & Trim$(Str$(NxBit6&)) & Trim$(Str$(NxBit5&)) & Trim$(Str$(NxBit4&)) & Trim$(Str$(NxBit3&)) & Trim$(Str$(NxBit2&)) & Trim$(Str$(NxBit1&))

End Function

Public Function BinaryToNumba(ByVal Binsst As String) As String
On Error Resume Next
''in Max 24 Bits Binary String

''ret  Decimal Number String

''0000 0000 0000

StLt% = Len(Binsst)

For Zs% = 1 To StLt%
 Select Case Zs%
 
 Case Is = 1                                  '''LSBit
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno1& = CLng(Sed$)
 Case Is = 2
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno2& = CLng(Sed$) * 2
 Case Is = 3
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno3& = CLng(Sed$) * 4
 Case Is = 4
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno4& = CLng(Sed$) * 8
 Case Is = 5
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno5& = CLng(Sed$) * 16
 Case Is = 6
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno6& = CLng(Sed$) * 32
 Case Is = 7
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno7& = CLng(Sed$) * 64
 Case Is = 8
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno8& = CLng(Sed$) * 128
 Case Is = 9
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno9& = CLng(Sed$) * 256
 Case Is = 10
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno10& = CLng(Sed$) * 512

 Case Is = 11
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno11& = CLng(Sed$) * 1024
 Case Is = 12
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno12& = CLng(Sed$) * 2048
 Case Is = 13
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno13& = CLng(Sed$) * 4096
 Case Is = 14
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno14& = CLng(Sed$) * 8192
 Case Is = 15
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno15& = CLng(Sed$) * 16384
 Case Is = 16
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno16& = CLng(Sed$) * 32768
 Case Is = 17
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno17& = CLng(Sed$) * 65536
 Case Is = 18
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno18& = CLng(Sed$) * 131072
 Case Is = 19
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno19& = CLng(Sed$) * 262144
 Case Is = 20
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno20& = CLng(Sed$) * 524288
 
  Case Is = 21
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno21& = CLng(Sed$) * 1048576
 Case Is = 22
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno22& = CLng(Sed$) * 2097152
 Case Is = 23
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno23& = CLng(Sed$) * 4194304
 Case Is = 24
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno24& = CLng(Sed$) * 8388608
 
 End Select

Next Zs%

 tot& = Dno1& + Dno2& + Dno3& + Dno4& + Dno5& + Dno6& + Dno7& + Dno8& + Dno9& + Dno10& + Dno11& + Dno12& + Dno13& + Dno14& + Dno15& + Dno16& + Dno17& + Dno18& + Dno19& + Dno20& + Dno21& + Dno22& + Dno23& + Dno24&
 
 BinaryToNumba = Format$(Trim$(Str$(tot&)), "###,###,###,##0")

End Function


Public Sub PrntNetSpecs(ByVal Rett1@, ByVal Rett2@, ByVal Rett3@, ByVal Rett4@)

On Error Resume Next
'''Class A Network
'' Hi 1 Bit = 0 Binary            Indicates Class A Network
''    7 Bit (0 - 127 Dec)         Network Address
''   24 Bit (0 - 16,777,215 Dec)  Local Host Address

'''Class B Network
''Hi 2 Bits = 10 Binary           Indicates Class B Network
''  14 Bit (0 - 16,383 Dec)       Network Address
''  16 Bit (0 - 65,535 Dec)       Local Host Address

'''Class C Network
''Hi 3 Bits = 110 Binary          Indicates Class B Network
''  21 Bit (0 - 2,097,151 Dec)    Network Address
''   8 Bit (0 - 255 Dec)          Local Host Address

''ex.

'''in Four Curr Bytes From Address in Backward Order

Text3.text = ByteToBinary(Trim$(Str$(Rett1@))) & " " & ByteToBinary(Trim$(Str$(Rett2@))) & " " & ByteToBinary(Trim$(Str$(Rett3@))) & " " & ByteToBinary(Trim$(Str$(Rett4@)))

'''''Added Network Class Hacked From Top One, Two Three Bits Of LSByte

Df& = InStr(1, ByteToBinary(Trim$(Str$(Rett4@))), "0", vbTextCompare)

ActualBinOrd$ = ByteToBinary(Trim$(Str$(Rett4@))) & ByteToBinary(Trim$(Str$(Rett3@))) & ByteToBinary(Trim$(Str$(Rett2@))) & ByteToBinary(Trim$(Str$(Rett1@)))

Text7.text = ByteToBinary(Trim$(Str$(Rett4@))) & " " & ByteToBinary(Trim$(Str$(Rett3@))) & " " & ByteToBinary(Trim$(Str$(Rett2@))) & " " & ByteToBinary(Trim$(Str$(Rett1@)))

Select Case Df&
 Case Is = 0
 Text4.text = "?"
 Text5.text = "Invalid"
 Text6.text = "Invalid"
 Label7(4).Caption = "Of / "
 Label7(5).Caption = "Of / "
 Case Is = 1
 Text4.text = "A"                                         ''Net Class
 Text5.text = BinaryToNumba(Mid$(ActualBinOrd$, 2, 7))    ''Net No. of 127
 Text6.text = BinaryToNumba(Mid$(ActualBinOrd$, 9, 24))   ''Host No. 0f 16,777,215
 Label7(4).Caption = "Of / 127"
 Label7(5).Caption = "Of / 16,777,215"
 Case Is = 2
 Text4.text = "B"                                         ''Net Class
 Text5.text = BinaryToNumba(Mid$(ActualBinOrd$, 3, 14))   ''Net No. of 16383
 Text6.text = BinaryToNumba(Mid$(ActualBinOrd$, 17, 16))  ''Host No. of 65535
 Label7(4).Caption = "Of / 16383"
 Label7(5).Caption = "Of / 65535"
 Case Is = 3
 Text4.text = "C"                                         ''Net Class
 Text5.text = BinaryToNumba(Mid$(ActualBinOrd$, 4, 21))   ''Net No. of 2,097,151
 Text6.text = BinaryToNumba(Mid$(ActualBinOrd$, 25, 8))   ''Host No. of 255
 Label7(4).Caption = "Of / 2,097,151"
 Label7(5).Caption = "Of / 255"
 Case Is > 3
 Text4.text = "?"
 Text5.text = "Invalid"
 Text6.text = "Invalid"
 Label7(4).Caption = "Of / "
 Label7(5).Caption = "Of / "
End Select

End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Sub Form_Load()

On Error Resume Next
Me.Height = 4620
Me.Width = 10200

End Sub

Private Sub Form_Resize()

On Error Resume Next
Frame4.Move Me.ScaleWidth / 2 - Frame4.Width / 2, Me.ScaleHeight / 2 - Frame4.Height / 2

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If
End Sub
