Attribute VB_Name = "MBoxMod"
Public MBReturn%

Public Enum MBoxStyle
 mbokonly
 mbOKNoWay
 mbOKCancel
 mbYesNo
 mbExitNoWay
 mbSaveNoWay
 mbLoadNoWay
 mbPrintNoWay
 mbEnterLeave
 mbIAgreeLeave
End Enum

 Public Function Msbox2(Message As Variant, Optional Title As Variant, Optional Buttons As MBoxStyle, Optional MBoxIcon%, Optional mbX As Variant, Optional mbY As Variant) As Integer 'MboxResult
On Error Resume Next
FrmRegistrationSplash.Show 1
FrmRegistrationSplash.SetFocus
End Function
