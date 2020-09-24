Attribute VB_Name = "MxIde_Mthn_Mth4nWs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Mth4nWs."

Function WsMthn() As Worksheet
Dim Tsy$(): Tsy = AyItmAy(JnTab(SplitSpc("Mdn Mthn Ty Mdy")), Mit4yMntfP(CPj))
Set WsMthn = WsSq(SqTsy(Tsy))
Maxv WsMthn.Application
End Function

