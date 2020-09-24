Attribute VB_Name = "MxDao_Dta_TDtaSrc_ChkEr"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dta_TDtaSrc_ChkEry."
Private Type FldMis: Tbn As String: EptFny() As String: ActFny() As String: End Type 'Deriving(Ctor Ay Opt)
Private Type OptFldMis: Som As Boolean: FldMis As FldMis: End Type
Private Type TEr
    EptFm As TDtaSrcFm: ActFm As TDtaSrcFm
    EptTny() As String: ActTny() As String
    FldMis() As FldMis
    End Type 'Deriving(Ctor Opt)
Private Type OptTEr: Som As Boolean: TEr As TEr: End Type
Private Function TEr(EptFm As TDtaSrcFm, ActFm As TDtaSrcFm, EptTny$(), ActTny$(), FldMis() As FldMis) As TEr
With TEr
    .EptFm = EptFm
    .ActFm = ActFm
    .EptTny = EptTny
    .ActTny = ActTny
    .FldMis = FldMis
End With
End Function
Private Function OptTEr(Som, A As TEr) As OptTEr: With OptTEr: .Som = Som: .TEr = A: End With: End Function
Private Function SomTEr(A As TEr) As OptTEr: SomTEr.Som = True: SomTEr.TEr = A: End Function
Private Sub PushTErDtaSrcOpt(A() As TEr, M As OptTEr)
With M
   If .Som Then PushTErDtaSrc A, .TEr
End With
End Sub
Private Function FldMis(Tbn, EptFny$(), ActFny$()) As FldMis
With FldMis
    .Tbn = Tbn
    .EptFny = EptFny
    .ActFny = ActFny
End With
End Function
Private Function AddFldMis(A As FldMis, B As FldMis) As FldMis(): PushFldMis AddFldMis, A: PushFldMis AddFldMis, B: End Function
Private Sub PushFldMisAy(O() As FldMis, A() As FldMis): Dim J&: For J = 0 To FldMisUB(A): PushFldMis O, A(J): Next: End Sub
Private Sub PushFldMis(O() As FldMis, M As FldMis): Dim N&: N = FldMisSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function FldMisSI&(A() As FldMis): On Error Resume Next: FldMisSI = UBound(A) + 1: End Function
Private Function FldMisUB&(A() As FldMis): FldMisUB = FldMisSI(A) - 1: End Function
Private Function OptFldMis(Som, A As FldMis) As OptFldMis: With OptFldMis: .Som = Som: .FldMis = A: End With: End Function
Private Function SomFldMis(A As FldMis) As OptFldMis: SomFldMis.Som = True: SomFldMis.FldMis = A: End Function
Private Sub PushFldMisOpt(A() As FldMis, M As OptFldMis)
With M
   If .Som Then PushFldMis A, .FldMis
End With
End Sub
Private Function AddTErDtaSrc(A As TEr, B As TEr) As TEr(): PushTErDtaSrc AddTErDtaSrc, A: PushTErDtaSrc AddTErDtaSrc, B: End Function
Private Sub PushTErDtaSrcAy(O() As TEr, A() As TEr): Dim J&: For J = 0 To TErDtaSrcUB(A): PushTErDtaSrc O, A(J): Next: End Sub
Private Sub PushTErDtaSrc(O() As TEr, M As TEr): Dim N&: N = TErDtaSrcSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function TErDtaSrcSI&(A() As TEr): On Error Resume Next: TErDtaSrcSI = UBound(A) + 1: End Function
Private Function TErDtaSrcUB&(A() As TEr): TErDtaSrcUB = TErDtaSrcSI(A) - 1: End Function

Sub ChkTDtaSrcNoEr(Ept As TDtaSrc, Act As TDtaSrc): ChkEry VVFmt(VVErOpt(Ept, Act)): End Sub

Private Sub B_VVErOpt()
GoSub T1
Exit Sub
Dim E As TDtaSrc, A As TDtaSrc, Act As OptTEr, Ept As OptTEr
    Dim Fm As TDtaSrcFm
    Dim TFy() As TF
T1:
    Fm = TDtaSrcFm("EptTDtaSrcFfn", "EptTDtaSrcn")
    PushTF TFy, TF("Tbn1", SplitSpc("A B C"))
    PushTF TFy, TF("Tbn2", SplitSpc("X Y Z"))
    E = TDtaSrc(Fm, TFy)
    
    Fm = TDtaSrcFm("ActTDtaSrcFfn", "ActTDtaSrcn")
    Erase TFy
    PushTF TFy, TF("Tbn1", SplitSpc("A B X"))
    A = TDtaSrc(Fm, TFy)
    GoTo Tst
Tst:
    Act = VVErOpt(E, A)
    Ass W3IsEq(Ept, Act)
    Return
End Sub
Private Function W3IsEq(A As OptTEr, B As OptTEr) As Boolean
Stop
End Function
Private Function VVErOpt(E As TDtaSrc, A As TDtaSrc) As OptTEr
Dim ETny$(): ETny = TnyTDtaSrc(E)
Dim ATny$(): ATny = TnyTDtaSrc(A)
Dim Intersect$(): Intersect = AyIntersect(ETny, ATny)
Dim FldMis() As FldMis
    Dim OkEpt() As TF: OkEpt = TFwTny(E.TF, Intersect)
    Dim OkAct() As TF: OkAct = TFwTny(A.TF, Intersect)
    FldMis = WFldMis(OkEpt, OkAct)
If Si(Intersect) = 0 Then
    If FldMisSI(FldMis) = 0 Then Exit Function
End If
VVErOpt = SomTEr(TEr(E.Fm, A.Fm, ETny, ATny, FldMis))
End Function
Private Function WFldMis(OkE() As TF, A() As TF) As FldMis() '@OkE is All Tbl should be found in @A
Dim J%: For J = 0 To UbTF(OkE)
    Dim Ept As TF: Ept = OkE(J)
    Dim Act As TF: Act = WFndAct(Ept.Tbn, A)
    Dim Opt As OptFldMis: Opt = WFldMisOpt(Ept, Act)
    PushFldMisOpt WFldMis, Opt
Next
End Function
Private Function WFndAct(EptTbn$, A() As TF) As TF
Const CSub$ = CMod & "WFndAct"
Dim J%: For J = 0 To UbTF(A)
    If A(J).Tbn = EptTbn Then WFndAct = A(J): Exit Function
Next
ThwLgc CSub, "EptTbn should be found in ActTFy", "EptTbn ActTny", EptTbn, TnyTF(A)
End Function
Private Function WFldMisOpt(E As TF, A As TF) As OptFldMis
Dim EptFny$(): EptFny = E.Fny
Dim ActFny$(): ActFny = A.Fny
Dim MisFny$(): MisFny = SyMinus(EptFny, ActFny)
If Si(MisFny) > 0 Then
    WFldMisOpt = SomFldMis(FldMis(E.Tbn, EptFny, ActFny))
End If
End Function

Private Sub B_VVFmt()
'GoSub T1
GoSub T2
Exit Sub
Dim Er As OptTEr
Dim FldMisy() As FldMis, ActSrcn$, ActFfn$, EptSrcn$, EptFfn$, ActTTT$, EptTTT$
T1:
    Erase FldMisy
    PushFldMis FldMisy, FldMis(Tbn:="AA", EptFny:=SplitSpc("A B C"), ActFny:=SplitSpc("A B X"))
    ActSrcn = "ActSrcn"
    ActFfn = "sdfsdf\ActFfn"
    EptSrcn = "ptSrcn"
    EptFfn = "sdfsdf\EptFfn"
    ActTTT = "TbA TbB TbC"
    EptTTT = "TbA TbB TbX"
    Er = SomTEr(TEr( _
        TDtaSrcFm(ActFfn, ActSrcn), _
        TDtaSrcFm(EptFfn, EptSrcn), _
        SplitSpc(EptTTT), _
        SplitSpc(ActTTT), _
        FldMisy))
        
    GoTo Tst
T2:
    Erase FldMisy
    'PushFldMis FldMisy, FldMis(Tbn:="AA", EptFny:=SplitSpc("A B C"), ActFny:=SplitSpc("A B X"))
    ActSrcn = "ActSrcn"
    ActFfn = "sdfsdf\ActFfn"
    EptSrcn = "sdfsdf\EptSrcn"
    EptFfn = "EptFfn"
    ActTTT = "TbA TbB TbC"
    EptTTT = "TbA TbB TbX"
    Er = SomTEr(TEr( _
        TDtaSrcFm(ActFfn, ActSrcn), _
        TDtaSrcFm(EptFfn, EptSrcn), _
        SplitSpc(EptTTT), _
        SplitSpc(ActTTT), _
        FldMisy))
        
    GoTo Tst
Tst:
    Brw VVFmt(Er)
    Return
End Sub
Private Function VVFmt(E As OptTEr) As String()
If Not E.Som Then Exit Function
Dim Nv() As S12
    With E.TEr
    Dim Mis$(): Mis = SyMinus(.ActTny, .EptTny)
    Dim Exc$(): Exc = SyMinus(.EptTny, .ActTny)
    Dim Msg$: Msg = FmtQQ("#Tables: Mis[?] Exccess[?] Act[?] Ept[?].  #Tables-with-MisFld[?]", _
        Si(Mis), Si(Exc), Si(.ActTny), Si(.EptTny), FldMisSI(.FldMis))
    PushS12 Nv, S12("Message", Msg)
    PushS12 Nv, S12("Missing table names", TmlAy(Mis))
    PushS12 Nv, S12("Excess table names", TmlAy(Exc))
    PushS12 Nv, S12("Expected table names", TmlAy(.EptTny))
    PushS12 Nv, S12("Actual table names", TmlAy(.ActTny))
    PushS12 Nv, S12("Expected data source", W2FmtFm(.EptFm))
    PushS12 Nv, S12("Actual data source", W2FmtFm(.ActFm))
    PushS12 Nv, S12("Table(s) with missing fields", JnCrLf(W2FmtFldMisy(.FldMis)))
    End With
Dim B$(): B = Box("Actual TDtaSrc is not as Expected")
VVFmt = Sy(B, FmtS12y(Nv))
End Function
Private Function W2FmtBox(TnyMis$(), FldMis() As FldMis) As String()
Dim NTblMis%: NTblMis = Si(TnyMis)
Dim NFldMis%: NFldMis = FldMisSI(FldMis)
W2FmtBox = Box(FmtQQ("[?] missing tables and [?] tables with missing fields", NTblMis, NFldMis))
End Function
Private Function W2FmtFldMisy(M() As FldMis) As String()
Const LyUL$ = "================================="
If FldMisSI(M) = 0 Then PushI W2FmtFldMisy, "N/A": Exit Function
PushI W2FmtFldMisy, LyUL
Dim J%: For J = 0 To FldMisUB(M)
    PushIAy W2FmtFldMisy, W2FmtFldMis(M(J))
    PushI W2FmtFldMisy, LyUL
Next
End Function
Private Function W2FmtFldMis(M As FldMis) As String()
Const NN$ = "[Tbn with missing fields] [Missing fields] [Expected fields] [Actual fields] [Excess fields in actual]"
With M
    Dim Mis$: Mis = TmlAy(SyMinus(.EptFny, .ActFny))
    Dim Ept$: Ept = TmlAy(.EptFny)
    Dim Act$: Act = TmlAy(.ActFny)
    Dim Exc$: Exc = TmlAy(SyMinus(.ActFny, .EptFny))
    W2FmtFldMis = MsgyNNAp(NN, .Tbn, Mis, Ept, Act, Exc)
End With
End Function
Private Function W2FmtFm$(Fm As TDtaSrcFm)
With Fm
W2FmtFm = FmtQQ("Name[?] Path[?] File[?]", .TDtaSrcn, Pth(.TDtaSrcFfn), Fn(.TDtaSrcFfn))
End With
End Function
