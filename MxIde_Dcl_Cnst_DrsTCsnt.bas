Attribute VB_Name = "MxIde_Dcl_Cnst_DrsTCsnt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Cnst_DrsTCsnt."
Public Const FFCnst$ = "Mdn Mdy Cnstn Tycn AftEq"
Private Sub B_DrsTCnstPC():                BrwDrs DrsTCnstPC: End Sub
Function DrsTCnstPC() As Drs: DrsTCnstPC = DrsTCnstP(CPj):    End Function
Function DrsTCnstP(P As VBProject) As Drs
Dim O As Drs
Dim C As VBComponent: For Each C In P.VBComponents
    O = DrsAdd(O, DrsTCnstM(C.CodeModule))
Next
DrsTCnstP = O
End Function

Function DrsTCnstDcl(Dcl$(), Mdn$) As Drs: DrsTCnstDcl = DrsFf(FFCnst, WDy(Dcl, Mdn)): End Function
Private Function WDy(Dcl$(), Mdn$) As Variant()
Dim L: For Each L In Itr(Contlny(Dcl))
    PushSomSi WDy, WDr(L, Mdn)
Next
End Function
Private Function WDr(Contln, Optional Mdn$) As Variant()
Const CSub$ = CMod & "WDr"
Dim S$: S = Contln
Dim ShtMdy$: ShtMdy = ShfShtMdy(S)               '<-- 1 Mdy
If Not IsShfCnst(S) Then Exit Function
Dim CnstnL$: CnstnL = ShfNm(S)                '<-- 2 Nm
If CnstnL = "" Then Thw CSub, "@Contln has [ShtMdy] & [Const], but no name", "L ShtMdy", Contln, ShtMdy
Dim Tyc$: Tyc = ShfTyc(S)             '<-- 3 Tyc
If Not IsShfPfx(S, " = ") Then If CnstnL = "" Then Thw CSub, "@L has [ShtMdy], [Const], [Constn] and [Tyc] , but [=] is missing", "L ShtMdy Constn Tyc", Contln, ShtMdy, CnstnL, Tyc
WDr = Array(Mdn, ShtMdy, CnstnL, Tyc, S)
End Function

Function DrsTCnstM(M As CodeModule) As Drs: DrsTCnstM = DrsTCnstDcl(DclM(M), Mdn(M)): End Function
