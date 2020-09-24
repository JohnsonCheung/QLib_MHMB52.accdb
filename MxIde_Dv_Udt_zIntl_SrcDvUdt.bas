Attribute VB_Name = "MxIde_Dv_Udt_zIntl_SrcDvUdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Deri_Udt_DvuSrc."
Private Sub B_SrcDvUdt()
GoSub ZZ
Exit Sub
Dim Src$()
ZZ:
    Dim C As VBComponent: For Each C In CPj.VBComponents
        SrcDvUdt SrcCmp(C)
    Next
    Return
End Sub
Private Function SrcDvUdt(Src$()) As String()
Dim U() As TUdt: U = TUdtyDcl(DclSrc(Src))
Dim O$(): O = Src
Dim J%: For J = 0 To UbTUdt(U)
    O = SrcDvUdtTUdt(O, U(J))
Next
SrcDvUdt = O
End Function
Private Function SrcDvUdtTUdt(Src$(), U As TUdt) As String()
Const Tp$ = "Type ?opt: Som As Boolean: ? As ?: End Type"

Dim MthnyDlt$(): MthnyDlt = MthnyDvUdtToDlt(U)
Dim UdtnDlt$: UdtnDlt = StrTrue(Not U.GenOpt, U.Udtn & "opt")
Dim CdlMth$: CdlMth = gen_mth_Cdl(U)
Dim CdlUdt$: If U.GenOpt Then CdlUdt = MdyPrv(U.IsPrv) & RplQ(Tp, U.Udtn)
Dim O$()
    O = SrcEnsUdt(Src, CdlUdt, UdtnDlt, U.Udtn)
SrcDvUdtTUdt = SrcEnsMth(O, CdlMth, MthnyDlt)
End Function

Private Function gen_mth_Cdl$(U As TUdt) '#Cd-For-Cur-Udt#
      Const TpAdd$ = "Function ?Add(A As ?, B As ?) As ?(): ?Push ?Add, A: ?Push ?Add, B: End Function"
   Const TpPushAy$ = "Sub Push?y(O() As ?, A() As ?): Dim J&: For J = 0 To Ub?(A): Push? O, A(J): Next: End Sub"
     Const TpPush$ = "Sub Push?(O() As ?, M As ?): Dim N&: N = Si?(O): ReDim Preserve O(N): O(N) = M: End Sub"
       Const TpSi$ = "Function Si?&(A() As ?): On Error Resume Next: Si? = UBound(A) + 1: End Function"
       Const TpUB$ = "Function Ub?&(A() As ?): Ub? = Si?(A) - 1: End Function"
  Const TpOptCtor$ = "Function ?opt(Som, A As ?) As ?Opt: With ?Opt: .Som = Som: .? = A: End With: End Function"
   Const TpOptSom$ = "Function Som?(A As ?) As ?opt: Som?.Som = True: Som?.? = A: End Function"
  Const TpOptPush$ = "Sub Push?opt(A() As ?, M As ?opt)|With M|   If .Som Then Push? A, .?|End With|End Sub"
Dim O$()
With U
    Dim PfxPrv$: PfxPrv = MdyPrv(.IsPrv) ' PrvPfx
    Dim N$: N = .Udtn
    PushIAy O, ctor__Cdy(PfxPrv, N, .GenCtor, U.Mbr)
    PushIAy O, CdyIf(PfxPrv, N, .GenAy, Sy(TpPush, TpSi, TpUB))
    PushIAy O, CdyIf(PfxPrv, N, .GenAdd, Sy(TpAdd))
    PushIAy O, CdyIf(PfxPrv, N, .GenPushAy, Sy(TpPushAy))
End With
gen_mth_Cdl = LinesLyNB(O)
End Function
Private Function ctor__Cdy(PfxPrv$, Udtn$, IsGen As Boolean, Mbr() As TUmb) As String()
If Not IsGen Then Exit Function
Dim O$()
    PushI O, ctor_Mthln(PfxPrv, Udtn, Mbr)
    PushI O, "With " & Udtn
    Dim J%: For J = 0 To UbTUmb(Mbr)
        PushI O, "    " & ctor_Mbln(Mbr(J))
    Next
    PushI O, "End With"
    PushI O, "End Function"
ctor__Cdy = O
End Function
Private Function ctor_Mthln$(PfxPrv$, Udtn$, Mbr() As TUmb)
Const Tp$ = "Function ?(?) As ?"
Dim Pm$: Pm = ctor_Pm(Mbr)
ctor_Mthln = PfxPrv & FmtQQ(Tp, Udtn, Pm, Udtn)
End Function
Private Function ctor_Mbln$(U As TUmb)
Dim PfxSet$
Select Case True
Case Not U.IsAy And IsTynObj(U.Tyn): PfxSet = "Set "
End Select
ctor_Mbln = PfxSet & RplQ(".? = ?", U.Mbn) 'The Udt constructor member line
End Function
Private Function ctor_Pm$(M() As TUmb)
Dim O$()
Dim J%: For J = 0 To UbTUmb(M)
    PushI O, ctor_Arg(M(J))
Next
ctor_Pm = JnCmaSpc(O)
End Function
Private Function ctor_Arg$(M As TUmb)
Dim N$: N = M.Mbn
Dim T$: T = M.Tyn
Dim O$
Dim IsPrim As Boolean: IsPrim = IsTynPrim(M.Tyn)
Select Case True
Case IsPrim And M.IsAy: O = FmtQQ("??()", N, TycTycn(T))
Case IsPrim:            O = N
Case M.IsAy:            O = FmtQQ("?() As ?", N, T)
Case Else:              O = FmtQQ("? As ?", N, T)
End Select
ctor_Arg = O
End Function

Private Sub B_ctor__Cdy()
Const T As Boolean = True
Const F As Boolean = False
GoSub T1
Exit Sub
Dim PfxPrv$, Udtn$, Mbr() As TUmb, GenCtor As Boolean, GenAy As Boolean, GenOpt As Boolean, GenAyAdd, GenPushAy
T1:
    PfxPrv = "Private "
    Udtn = "sdf"
    Erase Mbr
    GenCtor = True
    PushTUmb Mbr, TUmb(IsAy:=T, Mbn:="AA", Tyn:="ABC")
    PushTUmb Mbr, TUmb(IsAy:=F, Mbn:="BB", Tyn:="TUdt")
    PushTUmb Mbr, TUmb(IsAy:=F, Mbn:="CC", Tyn:="Ws")
    PushTUmb Mbr, TUmb(IsAy:=T, Mbn:="DD", Tyn:="Integer")
    
    Ept = SplitCrLf(RplVBar("Private Function sdf(AA() As ABC, BB As TUdt, CC As Ws, DD%()) As sdf|With sdf" & _
        "|    .AA = AA" & _
        "|    .BB = BB" & _
        "|    Set .CC = CC" & _
        "|    .DD = DD" & _
        "|End With|End Function"))
    GoTo Tst
Tst:
    Act = ctor__Cdy(PfxPrv, Udtn, GenCtor, Mbr)
    C
    Return
End Sub
Function SrcoptDvUdt(Src$()) As Lyopt: SrcoptDvUdt = LyoptOldNew(Src, SrcDvUdt(Src)): End Function

