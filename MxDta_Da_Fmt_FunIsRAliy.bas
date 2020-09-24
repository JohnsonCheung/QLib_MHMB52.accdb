Attribute VB_Name = "MxDta_Da_Fmt_FunIsRAliy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Fmt_FunIsRAliy."
Function BoolyAliR(AlirCii$, UCol%) As Boolean(): BoolyAliR = BoolyAliRCiy(IntySs(AlirCii), UCol): End Function
Function BoolyAliRCiy(CiyAliR%(), UCol%) As Boolean()
Dim O() As Boolean: ReDim O(UCol)
Dim J%: For J = 0 To UB(CiyAliR)
    O(CiyAliR(J)) = True
Next
BoolyAliRCiy = O
End Function
