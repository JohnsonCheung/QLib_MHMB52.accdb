Attribute VB_Name = "MxDta_Da_Wh_Dw"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Wh_Dw."

Function DeInVy(D As Drs, N, VyIn) As Drs
Const CSub$ = CMod & "DeInVy"
Dim Ix&: Ix = IxEle(D.Fny, N)
If Not IsArray(VyIn) Then Thw CSub, "Given VyIn is not an array", "Ty-VyIn", TypeName(VyIn)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    If Not HasEle(VyIn, Dr(Ix)) Then
        PushI Dy, Dr
    End If
Next
DeInVy = Drs(D.Fny, Dy)
End Function

Function DeRxy(D As Drs, Rxy&()) As Drs
Dim Dy(): Dy = D.Dy
Dim ODy()
Dim RxyInl&(): RxyInl = AyMinus(IxyU(UB(D.Dy)), Rxy)
Dim Rix: For Each Rix In Itr(RxyInl)
    PushI ODy, Dy(Rix)
Next
DeRxy = Drs(D.Fny, ODy)
End Function

Function DeVap(D As Drs, NN$, ParamArray Vap()) As Drs
Dim Vy(): Vy = Vap
DeVap = DeVy(D, NN, Vy)
End Function

Function DeVy(D As Drs, FF$, Vy) As Drs
Dim DyKey(): DyKey = DrsSelFf(D, FF).Dy
Dim Rxy&(): Rxy = RxyeDyVy(DyKey, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DeVy = Drs(D.Fny, ODy)
End Function
Function Dw2Eq(D As Drs, N12$, V1, V2) As Drs
Dim A$, B$: AsgN12 N12, A, B
Dw2Eq = DwEq(DwEq(D, A, V1), B, V2)
End Function
Function Dw3Eq(D As Drs, N123$, V1, V2, V3) As Drs
Dim A$, N12$: AsgN12 N123, A, N12
Dw3Eq = Dw2Eq(DwEq(D, A, V1), N12, V2, V3)
End Function

Function DwNBlnk(D As Drs, N$) As Drs:             DwNBlnk = DwNe(D, N, ""):                       End Function
Function Dw2EqDrp(D As Drs, N12$, V1, V2) As Drs: Dw2EqDrp = DrsDrpDc(Dw2Eq(D, N12, V1, V2), N12): End Function
Function Dw2Patn(D As Drs, N12$, Patn1$, Patn2$) As Drs
Dim A$, B$: AsgN12 N12, A, B
Dw2Patn = DwPatn(DwPatn(D, A, Patn1), B, Patn2)
End Function
Function Dw3EqDrp(D As Drs, N123$, V1, V2, V3) As Drs: Dw3EqDrp = DrsDrpDc(Dw3Eq(D, N123, V1, V2, V3), N123):  End Function
Function DwGt(D As Drs, N$, V) As Drs:                     DwGt = Drs(D.Fny, DyWhGt(D.Dy, CixDrs(D, N), V)):   End Function
Function DwNe(D As Drs, N$, V) As Drs:                     DwNe = Drs(D.Fny, DyWhNe(D.Dy, CixDrs(D, N), V)):   End Function
Function DwBlnk(D As Drs, N$) As Drs:                    DwBlnk = DwEq(D, N, ""):                              End Function
Function DwNB(D As Drs, N$) As Drs:                        DwNB = DwNe(D, N, ""):                              End Function
Function DwEq(D As Drs, N$, Eq) As Drs:                    DwEq = Drs(D.Fny, DyWhEq(D.Dy, Cix(D.Fny, N), Eq)): End Function
Function DwEqStr(D As Drs, N$, Str$) As Drs
If Str = "" Then DwEqStr = D: Exit Function
DwEqStr = DwEq(D, N, Str)
End Function

Function DwSsub(D As Drs, N$, Ssub) As Drs
Dim Ix&, Fny$()
Fny = D.Fny
Ix = CixDrs(D, N)
DwSsub = Drs(Fny, DyWhSsub(D.Dy, CixDrs(D, N), Ssub))
End Function

Function DwLik(D As Drs, N$, Lik) As Drs:                DwLik = DrsDy(D, DyWhLik(D.Dy, CixDrs(D, N), Lik)):      End Function
Function DwFalse(D As Drs, N$) As Drs:                 DwFalse = DwEq(D, N, False):                               End Function
Function DwFalseDrp(D As Drs, N$) As Drs:           DwFalseDrp = DrsDrpDc(DwFalse(D, N), N):                      End Function
Function DwEqDrp(D As Drs, N$, V) As Drs:              DwEqDrp = DrsDrpDc(DwEq(D, N, V), N):                      End Function
Function DwEqSelFf(D As Drs, N$, V, FfSel$) As Drs:  DwEqSelFf = DrsSelFf(DwEq(D, N, V), FfSel):                  End Function
Function DwNeSelFf(D As Drs, N$, V, FfSel$) As Drs:  DwNeSelFf = DrsSelFf(DwNe(D, N, V), FfSel):                  End Function
Function DwIn(D As Drs, N, VyIn) As Drs:                  DwIn = Drs(D.Fny, DyWhIn(D.Dy, IxEle(D.Fny, N), VyIn)): End Function
Function DwRxy(D As Drs, Rxy&()) As Drs:                 DwRxy = Drs(D.Fny, CvAv(AwIxy(D.Dy, Rxy))):              End Function

Function DwPatn(D As Drs, N$, Patn$) As Drs
If Patn = "" Then DwPatn = D: Exit Function
Dim R As RegExp: Set R = Rx(Patn)
Dim Cix%: Cix = CixDrs(D, N)
Dim Dy(), Dr: For Each Dr In Itr(D.Dy)
    If HitRx(Dr(Cix), R) Then PushI Dy, Dr
Next
DwPatn = Drs(D.Fny, Dy)
End Function

Function DwPfx(D As Drs, N$, Pfx) As Drs:             DwPfx = Drs(D.Fny, DyWhPfx(D.Dy, CixDrs(D, N), Pfx)): End Function
Function DwTop(D As Drs, Optional NTop& = 50) As Drs: DwTop = Drs(D.Fny, CvAv(AwFstN(D.Dy, NTop))):         End Function

Function DwVap(D As Drs, CC$, ParamArray Vap()) As Drs
'Fm D : ..@CC..            ! to be selected.  It has col-@CC
'Fm Vap : #Val-Ay-of-Param ! to select what rec in @D to be returned
'Ret    : ..@D..           ! sam stru as @D.  Subset of @D.  @@
Dim Vy(): Vy = Vap
Dim KeyDy(): KeyDy = DrsSelFf(D, CC).Dy
Dim Rxy&(): Rxy = RxywDyVy(KeyDy, Vy)
Dim ODy(): ODy = AwIxy(D.Dy, Rxy)
DwVap = Drs(D.Fny, ODy)
End Function

Function DwEqFny(D As Drs, N$, V, SelFny$()) As Drs: DwEqFny = DrsSelFny(DwEq(D, N, V), SelFny): End Function
Function DwInSel(D As Drs, N, VyIn, Sel$) As Drs:    DwInSel = DrsSelFf(DwIn(D, N, VyIn), Sel):  End Function
