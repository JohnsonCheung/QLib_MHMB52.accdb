Attribute VB_Name = "MxVb_Fs_Ffn_FfnNxt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_Nxt."

Private Sub B_FfnNxt()
Dim Ffn$
'GoSub T0
GoSub T1
Exit Sub
T1: Ffn = "AA(000).xls"
    Ept = "AA(001).xls"
    GoTo Tst
T0:
    Ffn = "AA.xls"
    Ept = "AA(001).xls"
    GoTo Tst
Tst:
    Act = FfnNxt(Ffn)
    C
    Return
End Sub

Function FfnNxt$(Ffn)
Dim N%: N = WNbrNxt(Ffn)
Dim F$: F = FfnRmvNxtNbr(Ffn)
FfnNxt = FfnAddFnsfx(F, "(" & Pad0(N + 1, 3) & ")")
End Function

Private Function WNbrNxt%(Ffn)
Dim A$: A = Right(Ffnn(Ffn), 5)
If ChrFst(A) <> "(" Then Exit Function
If ChrLas(A) <> ")" Then Exit Function
Dim M$: M = Mid(A, 2, 3)
If Not IsStrDig(M) Then Exit Function
WNbrNxt = M
End Function
Function FfnRmvNxtNbr$(Ffn)
If IsFfnNxt(Ffn) Then
    Dim A$: A = Ffnn(Ffn)
    FfnRmvNxtNbr = RmvLasN(A, 5) & Ext(Ffn)
Else
    FfnRmvNxtNbr = Ffn
End If
End Function
