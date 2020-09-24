Attribute VB_Name = "MxIde_Src_Shf_IsShf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Shf_A_IsShf."

Function IsShfCnst(OLn$) As Boolean: IsShfCnst = IsShfKw(OLn, "Const"):         End Function
Function IsShfDim(OLn$) As Boolean:   IsShfDim = IsShfKw(OLn, "Dim"):           End Function
Function IsShfSub(OLn$) As Boolean:   IsShfSub = IsShfKw(OLn, "Sub"):           End Function
Function IsShfKw(OLn$, Kw$):           IsShfKw = IsShfPfxSpc(OLn, Kw, eCasSen): End Function
Function IsShfPrv(OLn$) As Boolean:   IsShfPrv = IsShfKw(OLn, "Private"):       End Function
Function IsShfFrd(OLn$) As Boolean:   IsShfFrd = IsShfKw(OLn, "Friend"):        End Function
Function IsShfPub(OLn$) As Boolean
If IsShfPrv(OLn) Then Exit Function
If IsShfFrd(OLn) Then Exit Function
IsShfPub = True
End Function
