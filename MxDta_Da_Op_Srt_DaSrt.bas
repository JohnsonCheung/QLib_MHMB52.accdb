Attribute VB_Name = "MxDta_Da_Op_Srt_DaSrt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Op_Srt_DaSrt."
Private Sub B_DrsSrt()
Dim Drs As Drs, Act As Drs, Ept As Drs, SrtByFf$
GoSub T0
Exit Sub
T0:
    SrtByFf = "A B"
    Drs = DrsFf("A B C", DySSVBL("4 5 6|1 2 3|2 3 4"))
    Ept = DrsFf("A B C", DySSVBL("1 2 3|2 3 4|4 5 6"))
    GoTo Tst
Tst:
    Act = DrsSrt(Drs, SrtByFf)
    If Not IsEqDrs(Act, Ept) Then Stop
    Return
End Sub

Function DrsSrt(D As Drs, Optional TmlColnHyp$ = "") As Drs
If TmlColnHyp = "" Then DrsSrt = D: Exit Function
If NoRecDrs(D) Then DrsSrt = D: Exit Function
Dim K() As Srkey:      K = WSrkeyyTmlColnHyp(TmlColnHyp, D.Fny)
Dim Dy():               Dy = DySrtKeyy(D.Dy, K)
                    DrsSrt = Drs(D.Fny, Dy)
End Function
Private Function WSrkeyyTmlColnHyp(TmlColnHyp$, Fny$()) As Srkey()
Dim ColnHyp: For Each ColnHyp In Tmy(TmlColnHyp$)
    PushSrkey WSrkeyyTmlColnHyp, WSrkeyColnHyp(ColnHyp, Fny)
Next
End Function
Private Function WSrkeyColnHyp(ColnHyp, Fny$()) As Srkey
With WSrkeyColnHyp
If HasSfx(ColnHyp, "-") Then
    .Ci = IxEle(Fny, RmvLas(ColnHyp))
    .IsDes = True
ElseIf HasPfx(ColnHyp, "-") Then
    .Ci = IxEle(Fny, RmvFst(ColnHyp))
    .IsDes = True
Else
    .Ci = IxEle(Fny, ColnHyp)
    .IsDes = False
End If
End With
End Function
Private Function W2Srkey(CiiHypSrt$) As Srkey()
Dim S: For Each S In Itr(SplitSpc(CiiHypSrt))
    PushSrkey W2Srkey, WSrkeyCiHyp(S)
Next
End Function
Private Sub B_WSrkeyCiHyp()
GoSub T1
GoSub T2
Exit Sub
Dim CiHyp$, Ept As Srkey, Act As Srkey
T1: Ept = Srkey(1, False):  CiHyp = "1":  GoTo Tst
T2: Ept = Srkey(1, True): CiHyp = "1-": GoTo Tst
T3: Ept = Srkey(1, True): CiHyp = "-1": GoTo Tst
Tst:
    Act = WSrkeyCiHyp(CiHyp)
    Ass IsEqSrkey(Act, Ept)
    Return
With WSrkeyCiHyp("-0")
End With
End Sub
Private Function WSrkeyCiHyp(CiHyp) As Srkey
If ChrFst(CiHyp) = "-" Then
    WSrkeyCiHyp = Srkey(RmvFst(CiHyp), True)
ElseIf ChrLas(CiHyp) = "-" Then
    WSrkeyCiHyp = Srkey(RmvLas(CiHyp), True)
Else
    WSrkeyCiHyp = Srkey(CiHyp, False)
End If
End Function

Function DtSrt(A As Dt, Optional TmlColnHyp$ = "") As Dt: DtSrt = DtDrs(DrsSrt(DrsDt(A), TmlColnHyp), A.Dtn): End Function

Private Function SrkeyyTmlColnHyp(TmlColnHyp$, Fny$()) As Srkey()
Const CSub$ = CMod & "SrkeyyTmlColnHyp"
If TmlColnHyp = "" Then Exit Function
Dim F: For Each F In ItrTml(TmlColnHyp)
    Dim IsDes As Boolean
    If HasPfx(F, "-") Then
        F = RmvPfx(F, "-")
        IsDes = True
    Else
        IsDes = False
    End If
    Dim Ix%: Ix = IxEle(Fny, F)
    If Ix = -1 Then Thw CSub, "@TmlColnHyp has a *er-field not in @Fny", "@TmlColnHyp *er-fldn @", TmlColnHyp, F, Fny
    PushSrkey SrkeyyTmlColnHyp, Srkey(Ix, IsDes) '<===
Next
End Function
