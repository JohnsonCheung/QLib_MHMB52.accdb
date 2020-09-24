Attribute VB_Name = "MxVb_Str_Jn"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Jn."
Function Jn$(Ay, Optional Sep$ = ""): Jn = Join(SyAy(Ay), Sep): End Function
Function JnAp$(ParamArray Sap())
Dim Av(): Av = Sap
JnAp = Jn(Av)
End Function
Function JnApDot$(ParamArray Ap())
Dim Av(): If UBound(Ap) > 0 Then Av = Ap: JnApDot = JnDot(Av)
End Function
Function JnApDotNB$(ParamArray Ap())
Dim Av(): If UBound(Ap) > 0 Then Av = Ap: JnApDotNB = JnDotNB(Av)
End Function
Function JnApSep$(Sep$, ParamArray Sap())
Dim Av(): Av = Sap
JnApSep = Jn(Av, Sep)
End Function
Function JnBq$(Ay):         JnBq = Jn(Ay, "`"):     End Function ':Bq: :Chr ! #Back-quo#
Function JnCma$(Ay):       JnCma = Jn(Ay, ","):     End Function
Function JnCmaSpc$(Ay): JnCmaSpc = Jn(Ay, ", "):    End Function
Function JnCrLf$(Ay):     JnCrLf = Jn(Ay, vbCrLf):  End Function
Function JnLf$(Ay):         JnLf = Jn(Ay, vbLf):    End Function
Function JnCrLf2$(Ay):   JnCrLf2 = Jn(Ay, vbCrLf2): End Function
Function JnCrLfAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnCrLfAp = Jn(Av, vbCrLf)
End Function
Function JnDblDollar$(Ay): JnDblDollar = Jn(Ay, "$$"):   End Function
Function JnDot$(Ay):             JnDot = Jn(Ay, "."):    End Function
Function JnDotNB$(Ay):         JnDotNB = JnNB(Ay, "."):  End Function
Function JnOr$(Ay):               JnOr = Jn(Ay, " or "): End Function
Function JnPfxSno$(Pfx$, Eno%, Optional Bno% = 1, Optional NDig% = 2, Optional Sep$ = vbCmaSpc)
Dim F$: F = String(NDig, "0")
Dim O$()
    Dim J%: For J = Bno To Eno
        PushI O, Pfx & Format(J, F)
    Next
JnPfxSno = Jn(O, Sep)
End Function
Function JnPthSep$(Ay): JnPthSep = Jn(Ay, SepPth): End Function
Function JnPthSepAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnPthSepAp = JnPthSep(Av)
End Function
Function JnSemi$(Ay):           JnSemi = Jn(Ay, ";"):       End Function
Function JnSpc$(Ay):             JnSpc = Jn(Ay, " "):       End Function
Function JnTab$(Ay):             JnTab = Join(Ay, vbTab):   End Function
Function JnVBar$(Ay):           JnVBar = Jn(Ay, "|"):       End Function
Function JnVbarSpc$(Ay):     JnVbarSpc = Jn(Ay, " | "):     End Function
Function QuoBktJnCma$(Ay): QuoBktJnCma = QuoBkt(JnCma(Ay)): End Function
