Attribute VB_Name = "MxVb_Dta_T1likk"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_T1likk."
Function HitLikk(S, Likk) As Boolean: HitLikk = HitLiky(S, SySs(Likk)): End Function
Function HitLiky(S, Liky$()) As Boolean
Dim Lik: For Each Lik In Itr(Liky)
    If S Like Lik Then HitLiky = True: Exit Function
Next
End Function

Function HitSsyLik(S, SsyLik$()) As Boolean
Dim Likk: For Each Likk In SsyLik
    If HitLikk(S, Likk) Then HitSsyLik = True: Exit Function
Next
End Function

Private Sub B_T1T1likssy()
Dim A$(), S$
GoSub T1
GoSub T2
Exit Sub
T1:
    A = SplitVBar("a bb* *dd | c x y")
    S = "x"
    Ept = "c"
    GoTo Tst
T2:
    A = SplitVBar("a bb* *dd | c x y")
    S = "bb1"
    Ept = "a"
    GoTo Tst
Tst:
    Act = T1T1likssy(A, S)
    C
    Return
End Sub

Function T1T1likssy$(T1likssy$(), S)
Dim T1likss: For Each T1likss In Itr(T1likssy)
    If HitLikk(S, T1likss) Then
        T1T1likssy = Tm1(T1likss)
        Exit Function
    End If
Next
End Function
