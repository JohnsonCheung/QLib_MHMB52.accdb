Attribute VB_Name = "MxVb_Str_Shf_IsShf"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Shf_IsShf."
Private Sub B_IsShfPfx()
Dim O$: O = "AA{|}BB "
Ass IsShfPfx(O, "{|}") = "AA"
Ass O = "BB "
End Sub
Function IsShfBkt(OLn$, Optional BktOpn$ = vbBktOpn) As Boolean
Dim L%: L = Len(OLn)
ShfPfxSpc OLn, BktOpn & BktCls(BktOpn)
IsShfBkt = L > Len(OLn)
End Function
Function IsShfPfxSpc(OLn$, Pfx$, Optional C As eCas) As Boolean
Dim L%: L = Len(OLn)
ShfPfxSpc OLn, Pfx, C
IsShfPfxSpc = L > Len(OLn)
End Function
Function IsShfPfx(OLn$, Pfx$, Optional C As eCas) As Boolean
Dim L%: L = Len(OLn)
ShfPfx OLn, Pfx, C
IsShfPfx = L > Len(OLn)
End Function
