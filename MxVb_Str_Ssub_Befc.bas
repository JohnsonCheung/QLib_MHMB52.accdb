Attribute VB_Name = "MxVb_Str_Ssub_Befc"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Ssub_Befc."

Function BefcDash$(S): BefcDash = Bef(S, "_"):               End Function
Function BefcHH$(S):     BefcHH = RTrim(BefOrAll(S, "--")):  End Function
Function BefcDDD$(S):   BefcDDD = RTrim(BefOrAll(S, "---")): End Function
