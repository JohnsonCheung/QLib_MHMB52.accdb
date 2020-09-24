Attribute VB_Name = "MxVb_Str_Fmt_FmtParcc"
Option Compare Text
Option Explicit
Function Tm2yPChdS1(SS2y$()) As String()
'Each S1 will be one line and all S2 of same will be put in that line.
If Si(SS2y) <= 0 Then Tm2yPChdS1 = SS2y: Exit Function
Dim A$(): A = SySrtQ(CvSy(AwDis(SS2y)))
Dim M$(), S1Las$: S1Las = BefSpc(SS2y(0))
Dim J&: For J = 1 To UB(A)
    Dim S12Cur As S12: S12Cur = BrkSpc(A(J))
    Dim S1Cur$: S1Cur = S12Cur.S1
    If S1Cur = S1Las Then
        PushI M, S12Cur.S2
    Else
        If Si(M) > 0 Then PushI Tm2yPChdS1, S1Las & " " & JnSpc(M): Erase M
        PushI M, S12Cur.S2
        S1Las = S12Cur.S1
    End If
Next
If Si(M) > 0 Then PushI Tm2yPChdS1, S1Las & " " & JnSpc(M): Erase M
End Function
Function Tm2yPChdS2(SS2y$()): Tm2yPChdS2 = Tm2yPChdS1(SS2ySwap(SS2y)): End Function
Function SS2ySwap(SS2y$()) As String()
Dim SS2: For Each SS2 In Itr(SS2y)
    PushI SS2ySwap, SS2Swap(SS2)
Next
End Function
Function SS2Swap$(SS2)
With BrkSpc(SS2)
    If HasSpc(.S2) Then Thw CSub, "SS2.S2 cannot have space after BrkSpc"
    SS2Swap = .S2 & " " & .S1
End With
End Function
Function FmtParcc(Tm2yPChd$(), Optional H12$ = "S1 S2ss", Optional Wdt% = 120): FmtParcc = FmtS12y(S12yTm2yPChd(Tm2yPChd), H12, Wdt): End Function
Function FmtChdpp(Tm2yPChd$(), Optional H12$ = "S1 S2ss", Optional Wdt% = 120)
: FmtChdpp = FmtS12y(S12yTm2yCPar(Tm2yPChd), H12, Wdt): End Function
Private Function S12yTm2yPChd(Tm2yPChd$()) As S12()
Dim SSPChd: For Each SSPChd In Itr(Tm2yPChd)
    PushS12 S12yTm2yPChd, Brk1Spc(SSPChd)
Next
End Function
Private Function S12yTm2yCPar(Tm2yPChd$()) As S12()
Dim SSPChd: For Each SSPChd In Itr(Tm2yPChd)
    PushS12 S12yTm2yCPar, S12Swap(Brk1Spc(SSPChd))
Next
End Function

