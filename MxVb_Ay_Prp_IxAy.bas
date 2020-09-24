Attribute VB_Name = "MxVb_Ay_Prp_IxAy"
Option Compare Text
Const CMod$ = "MxVb_Ay_Ix."
Option Explicit

Function NotPfxIx&(Ly$(), Pfx$, Optional Bix = 0)
If Bix < 0 Then NotPfxIx = -1: Exit Function
Dim O&: For O = Bix To UB(Ly)
   If Not HasPfx(Ly(O), Pfx) Then NotPfxIx = O: Exit Function
Next
NotPfxIx = -1
End Function
Function PfxIx&(Ly$(), Pfx$, Optional Bix = 0)
If Bix < 0 Then PfxIx = -1: Exit Function
Dim O&: For O = Bix To UB(Ly)
   If HasPfx(Ly(O), Pfx) Then PfxIx = O: Exit Function
Next
PfxIx = -1
End Function

Function IxMay&(Ay, Ele, Optional Bix& = 0)
Dim J&: For J = Bix To UB(Ay)
    If Ay(J) = Ele Then IxMay = J: Exit Function
Next
End Function
Function IxEle&(Ay, Ele, Optional Bix& = 0)
For IxEle = Bix To UB(Ay)
    If Ay(IxEle) = Ele Then Exit Function
Next
IxEle = -1
End Function
Function IxMust&(Ay, Ele, Optional Bix& = 0)
IxMust = IxEle(Ay, Ele, Bix)
If IxMust = -1 Then Thw CSub, "@Ele not found in @Ay from @Bix", "@Ele @Bix @Ay-Tyn @Ay", Ele, Bix, TypeName(Ay), Ay
End Function

Function IxiMay&(Ay, Ele, Optional Bi& = 0)
Dim J&: For J = Bi To UB(Ay)
    If Ay(J) = Ele Then IxiMay = J: Exit Function
Next
End Function
Function IxiEle&(Ay, Ele, Optional Bi& = 0)
Dim J&: For J = Bi To UB(Ay)
    If Ay(J) = Ele Then IxiEle = J: Exit Function
Next
IxiEle = -1
End Function
Function IxiMust&(Ay, Ele, Optional Bi& = 0)
IxiMust = IxiEle(Ay, Ele, Bi)
If IxiMust = -1 Then Thw CSub, "@Ele not found in @Ay from @Bi", "@Ele @Bi @Ay-Tyn @Ay", Ele, Bi, TypeName(Ay), Ay
End Function

Function Cix%(Fny$(), C$)
Const CSub$ = CMod & "Cix"
Cix = IxEle(Fny, C): If Cix = -1 Then Stop 'Thw CSub, "Given @C not found in @Fny", "C Fny", C, TmlAy(Fny)
End Function

