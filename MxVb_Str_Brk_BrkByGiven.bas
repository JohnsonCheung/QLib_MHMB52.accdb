Attribute VB_Name = "MxVb_Str_Brk_BrkByGiven"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Brk_BrkByGiven."
Function Brk1Spc(S) As S12:                                 Brk1Spc = Brk1(Trim(S), " "):   End Function
Function Brk1Colon(S, Optional NoTrim As Boolean) As S12: Brk1Colon = Brk1(S, ":", NoTrim): End Function
Function Brk1Cma(S, Optional NoTrim As Boolean) As S12:     Brk1Cma = Brk1(S, ",", NoTrim): End Function
Function Brk1Dot(S, Optional NoTrim As Boolean) As S12:     Brk1Dot = Brk1(S, ".", NoTrim): End Function
Function Brk2Dot(S, Optional NoTrim As Boolean) As S12:     Brk2Dot = Brk2(S, ".", NoTrim): End Function
Function BrkDot(S, Optional NoTrim As Boolean) As S12:       BrkDot = Brk(S, ".", NoTrim):  End Function
Function BrkSpc(S) As S12:                                   BrkSpc = Brk(Trim(S), " "):    End Function
Function Brk1Eq(S, Optional NoTrim As Boolean) As S12:       Brk1Eq = Brk1(S, "=", NoTrim): End Function
Function Brk2Eq(S, Optional NoTrim As Boolean) As S12:       Brk2Eq = Brk2(S, "=", NoTrim): End Function
Function BrkEq(S, Optional NoTrim As Boolean) As S12:         BrkEq = Brk1(S, "=", NoTrim): End Function
