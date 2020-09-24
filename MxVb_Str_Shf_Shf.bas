Attribute VB_Name = "MxVb_Str_Shf_Shf"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Shf_Shf."
Function ShfBefSS$(OLn$, BefSS$)
Const CSub$ = CMod & "ShfBefSS"
Dim Bef: For Each Bef In SplitSpc(BefSS)
    ShfBefSS = ShfBefIf(OLn, Bef)
    If ShfBefSS <> "" Then Exit Function
Next
Thw CSub, "No BefSS in OLn", "BefSS OLn", BefSS, OLn
End Function

Function ShfDotn$(OLn$)
Stop '
End Function

Function ShfLHS$(OLn$)
Dim L$:                   L = OLn
Dim IsSet As Boolean: IsSet = IsShfTm(L, "Set")
Dim S$:                       If IsSet Then S = "Set "
Dim Lhs$:               Lhs = ShfDotn(L)
If ChrFst(L) = "(" Then
    Lhs = Lhs & QuoBkt(BetBkt(L))
    L = AftBkt(L)
End If
If Not IsShfPfx(L, " = ") Then Exit Function
ShfLHS = S & Lhs & " = "
OLn = L
End Function

Function ShfRHS(OLn$) As Variant()
Dim L$:     L = OLn
Dim Lhs$: Lhs = ShfLHS(L)
With Brk1(L, "'")
    Dim RHS$:  RHS = .S1
              OLn = "   ' " & .S2
End With
ShfRHS = Array(Lhs, RHS)
End Function
