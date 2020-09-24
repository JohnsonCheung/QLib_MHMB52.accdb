Attribute VB_Name = "MxVb_Str_Hit"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Hit."
Function HasPfxySpc(S, Pfxy$(), Optional C As eCas) As Boolean
Dim P: For Each P In Pfxy
    If HasPfxSpc(S, P, C) Then HasPfxySpc = True: Exit Function
Next
End Function

Function HitKss(S, Kss) As Boolean
HitKss = HitLiky(S, SySs(Kss))
End Function

Function HitT1y(S, T1y$(), Optional C As eCas) As Boolean: HitT1y = HasEleStr(T1y, Tm1(S), C): End Function
Function HitAp(V, ParamArray Ap()) As Boolean
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
HitAp = HasEle(Av, V)
End Function
Private Sub B_HitPatn()
Dim S$, Patn$
Ept = True: S = "AA": Patn = "AA": GoSub Tst
Ept = True: S = "AA": Patn = "^AA$": GoSub Tst
Exit Sub
Tst:
    Act = HitPatn(S, Patn)
    C
    Return
End Sub

Function HitPatn(S, Patn$) As Boolean: HitPatn = Rx(Patn).Test(S): End Function
Function HitRx(S, Rx As RegExp) As Boolean
If S = "" Then Exit Function
If IsNothing(Rx) Then Exit Function
HitRx = Rx.Test(S)
End Function
