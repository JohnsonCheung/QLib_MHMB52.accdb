Attribute VB_Name = "MxVb_Str_Term_AsgTml"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Term_AsgTml."

Sub AsgTAp(Tml$, ParamArray OApTm())
Dim T$(): T = Tmy(Tml)
Dim Av(): Av = OApTm

Dim U1%, U2%: U1 = UB(T): U2 = UB(Av)
Dim J%
For J = 0 To U2: OApTm(J) = Empty: Next
For J = 0 To Min(U1, U2)
    OApTm(J) = T(J)
Next
End Sub
Sub AsgN12(N12$, ON1$, ON2$):                     AsgT1r N12, ON1, ON2:                              End Sub
Sub AsgN123(N123$, ON1$, ON2$, ON3$):             AsgT2r N123, ON1, ON2, ON3:                        End Sub
Sub AsgT2(Tmo2, OT1, OT2):                        AsgAy Tmy2(Tmo2), OT1, OT2:                        End Sub
Sub AsgT3(Tmo3, OT1, OT2, OT3):                   AsgAy Tmy3(Tmo3), OT1, OT2, OT3:                   End Sub
Sub AsgT4(Tmo4, OT1, OT2, OT3, OT4):              AsgAy Tmy4(Tmo4), OT1, OT2, OT3, OT4:              End Sub
Sub AsgT5(Tmo5, OT1, OT2, OT3, OT4, OT5):         AsgAy Tmy5(Tmo5), OT1, OT2, OT3, OT4, OT5:         End Sub
Sub AsgT1r(Tmo1r, OT1, ORst):                     AsgAy Tmy1r(Tmo1r), OT1, ORst:                     End Sub
Sub AsgT2r(Tmo2r, OT1, OT2, ORst):                AsgAy Tmy2r(Tmo2r), OT1, OT2, ORst:                End Sub
Sub AsgT3r(Tmo3r, OT1, OT2, OT3, ORst):           AsgAy Tmy3r(Tmo3r), OT1, OT2, OT3, ORst:           End Sub
Sub AsgT4r(Tmo4r, OT1, OT2, OT3, OT4, ORst):      AsgAy Tmy4r(Tmo4r), OT1, OT2, OT3, OT4, ORst:      End Sub
Sub AsgT5r(Tmo5r, OT1, OT2, OT3, OT4, OT5, ORst): AsgAy Tmy5r(Tmo5r), OT1, OT2, OT3, OT4, OT5, ORst: End Sub

Sub AsgT1ry(Tmo1ry$(), OT1y, ORsty)
Erase OT1y, ORsty
Dim L: For Each L In Itr(Tmo1ry)
    PushI OT1y, Tm1(L)
    PushI ORsty, RmvTm1(L)
Next
End Sub
Sub AsgT2ry(Tmo2ry$(), OT1y, OT2y, ORsty)
Erase OT1y, OT2y, ORsty
Dim T1$, T2$, Rst$
Dim L: For Each L In Itr(Tmo2ry)
    AsgT2r L, T1, T2, Rst
    PushI OT1y, T1
    PushI OT2y, T2
    PushI ORsty, Rst
Next
End Sub
Sub AsgT3ry(Tmo3ry$(), OT1y, OT2y, OT3y, ORsty)
Erase OT1y, OT2y, OT3y, ORsty
Dim T1$, T2$, T3$, Rst$
Dim L: For Each L In Itr(Tmo3ry)
    AsgT3r L, T1, T2, T3, Rst
    PushI OT1y, T1
    PushI OT2y, T2
    PushI OT3y, T3
    PushI ORsty, Rst
Next
End Sub
