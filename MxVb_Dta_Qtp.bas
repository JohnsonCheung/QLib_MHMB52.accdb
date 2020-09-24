Attribute VB_Name = "MxVb_Dta_Qtp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Qtp."
Function NyQss(Qtp_or_Ss$) As String()
Dim Qss$: Qss = Qtp_or_Ss$
If HasSsub(Qss, "?") Then
    NyQss = NyQtp(Qss)
Else
    NyQss = SySs(Qss)
End If
End Function
Function NyQtp(Qtp) As String(): NyQtp = NyQtp2(Tm1(Qtp), RmvA1T(Qtp)): End Function
Function NnQtp$(Qtp):            NnQtp = JnSpc(NyQtp(Qtp)):             End Function
Function LyQtpy(Qtpy$()) As String
Dim Qtp: For Each Qtp In Itr(Qtpy)
    PushI LyQtpy, NnQtp(Qtp)
Next
End Function
Function NnQtp2$(Qtp1$, Rst$): NnQtp2 = JnSpc(NyQtp2(Qtp1, Rst)): End Function
Function NyQtp2(Qtp1$, Rst$) As String()
Dim IRst: For Each IRst In ItrSS(Rst)
    PushI NyQtp2, RplQ(Qtp1, IRst)
Next
End Function
Private Sub B_LyQtpy()
GoSub T1
Exit Sub
Dim Qtpy$()
T1:
    Qtpy = Sy("AAA? X Y", "B?BB 1 2 3")
    Ept = Sy("AAAX AAAY", "B1BB B2BB B3BB")
    GoTo Tst
Tst:
    Act = LyQtpy(Qtpy)
    C
    Return
End Sub
