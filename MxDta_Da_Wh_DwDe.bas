Attribute VB_Name = "MxDta_Da_Wh_DwDe"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Wh_DwDe."


Sub AsgColonFfTo(ColonFf$, OFnyA$(), OFnyB$())
Erase OFnyA, OFnyB
Dim F: For Each F In SySs(ColonFf)
    With BrkBoth(F, ":")
        PushI OFnyA, .S1
        PushI OFnyB, .S2
    End With
Next
End Sub

Function DeCeqC(D As Drs, CC$) As Drs
Dim Dr, C1&, C2&
AsgCxapDrs D, CC, C1, C2
For Each Dr In Itr(D.Dy)
    If Dr(C1) <> Dr(C2) Then
        PushI DeCeqC.Dy, Dr
    End If
Next
DeCeqC.Fny = D.Fny
End Function

Function DwCeqC(D As Drs, CC$) As Drs
Dim Dr, C1&, C2&
AsgCxapDrs D, CC, C1, C2
For Each Dr In Itr(D.Dy)
    If Dr(C1) = Dr(C2) Then
        PushI DwCeqC.Dy, Dr
    End If
Next
DwCeqC.Fny = D.Fny
End Function

Function DwCneC(D As Drs, CC$) As Drs
DwCneC = DeCeqC(D, CC)
End Function

Function DySelMay(Dy(), Ciy%()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DySelMay, AwIxyMay(Dr, Ciy)
Next
End Function

Function DyWhKey(Dy(), KeyIxy&(), Key()) As Variant()
'Ret : SubSet-of-row of @Dy for each row has val of %CurKey = @Key
Dim Dr: For Each Dr In Itr(Dy)
    Dim CurK: CurK = AwIxy(Dr, KeyIxy)
    If IsEqAy(CurK, Key) Then         '<- If %CurKey = @Key, select it.
        PushI DyWhKey, Dr
    End If
Next
End Function

Function DyWhKeySel(Dy(), KeyIxy&(), Key(), CiySel%()) As Variant()
DyWhKeySel = DySel(DyWhKey(Dy, KeyIxy, Key), CiySel)
End Function

Function IxOptDyDr(Dy(), Dr) As Lngopt
Dim IDr, Ix&
For Each IDr In Itr(Dy)
    If IsEqAy(IDr, Dr) Then IxOptDyDr = SomLng(Ix): Exit Function
    Ix = Ix + 1
Next
End Function

Function JnDrs(A As Drs, B As Drs, Jn$, Add$, Optional IsLeftJn As Boolean, Optional AnyFld$) As Drs
'@A        :..@Jn-LHS..            It is a drs with col-@Jn-LHS.
'@B        :..@Jn-RHS..@Add-RHS    It is a drs with col-@Jn-RHS & col-@Add-RHS.
'@Jn       :SS-of-:ColonTm       It is :SS-of-:ColTm. :ColTm: is a :Tm with 1-or-0 [:]. :Tm: is a fm :TLn: or :TmLn:  LHS of [:] is for @A and RHS of [:] is for @B
'                                  It is used to jn @A & @B
'@Add      :SS-of-ColonStr-Fld-@B  What col in @B to be added to @A.  It may use new name, if it has colon.
'@IsLeftJn :Bool                   Is it left join, otherwise, it is inner join
'@AnyFld   :Fldn                    It is optional fld to be add to rslt drs stating if any rec in @B according to @Jn.
'                                    It is vdt only when IsLeftJn=True.
'                                    It has bool value.  It will be TRUE if @B has jn rec else FALSE.
'Ret       :..@A..@Add-RHS..@AnyFld  It has all fld from @A and @Add-RHS-fld and optional @AnyFld.
'                                    If @IsLeftJn, it will have at least same rec as @A, and may have if there is dup rec in @B accord to @Jn fld.
'                                    If not @IsLeftJn, only those records fnd in both @A & @B
Dim JnFnyA$(), JnFnyB$()
Dim AddFnyFm$(), AddFnyAs$()
    AsgTmlFldMap Jn, JnFnyA, JnFnyB
    AsgTmlFldMap Add, AddFnyFm, AddFnyAs
    
Dim CiyAdd%(): CiyAdd = InyFnySub(B.Fny, AddFnyFm)
Dim BJnIxy&(): BJnIxy = IxyEley(B.Fny, JnFnyB)
Dim AJnIxy&(): AJnIxy = IxyEley(A.Fny, JnFnyA)

Dim Emp() ' it is for LeftJn and for those rec when @B has no rec joined.  It is for @Add-fld & @AnyFld.
          ' It has sam ele as @Add.  1 more fld is @AnyFld<>""
    If IsLeftJn Then
        ReDim Emp(UB(AddFnyFm))
        If AnyFld <> "" Then PushI Emp, False
    End If
Dim ODy()                       ' Bld %ODy for each %ADr, that mean fld-Add & fld-Any
    Dim Adr: For Each Adr In Itr(A.Dy)
        Dim JnVy():            JnVy = AwIxy(Adr, AJnIxy)                     'JnFld-Vy-Fm-@A
        Dim Bdy():            Bdy = DyWhKeySel(B.Dy, BJnIxy, JnVy, CiyAdd) '@B-Dy-joined
        Dim NoRec As Boolean: NoRec = Si(Bdy) = 0                           'no rec joined
            
        Select Case True
        Case NoRec And IsLeftJn: PushI ODy, AyAdd(Adr, Emp) '<== ODy, Only for NoRec & LeftJn
        Case NoRec
        Case Else
            '
            Dim Bdr: For Each Bdr In Bdy
                If AnyFld <> "" Then
                    Push Bdr, True
                End If
                PushI ODy, AyAdd(Adr, Bdr) '<== ODy, for each %BDr in %BDy, push to %ODy
            Next
        End Select
    Next Adr

Dim O As Drs: O = Drs(SyNB(A.Fny, AddFnyAs, AnyFld), ODy)
JnDrs = O

If False Then
    Erase XX
    XBox "Debug JnDrs"
    X "A-Fny  : " & TmlAy(A.Fny)
    X "B-Fny  : " & TmlAy(B.Fny)
    X "Jn     : " & Jn
    X "Add    : " & Add
    X "IsLefJn: " & IsLeftJn
    X "AnyFld : [" & AnyFld & "]"
    X "O-Fny  : " & TmlAy(O.Fny)
    X "More ..: A-Drs B-Drs Rslt"
    X FmtDrsNmv("A-Drs  : ", A)
    X FmtDrsNmv("B-Drs  : ", B)
    X FmtDrsNmv("Rslt   : ", O)
    Brw XX
    Erase XX
    Stop
End If
End Function

Function LDrsJn(A As Drs, B As Drs, Jn$, Add$, Optional AnyFld$) As Drs
LDrsJn = JnDrs(A, B, Jn, Add, IsLeftJn:=True, AnyFld:=AnyFld)
End Function

Function SelDistFny(D As Drs, Fny$()) As Drs
With GpCntFny(D, Fny)
    SelDistFny = Drs(Fny, .GpDy)
End With
End Function
Function SelDistAllCol(D As Drs) As Drs
With GpCntAllCol(D)
    SelDistAllCol = Drs(D.Fny, .GpDy)
End With
End Function

Function SelDist(D As Drs, FF$) As Drs
With GpCnt(D, FF)
    SelDist = DrsFf(FF, .GpDy)
End With
End Function

Function SelDistCnt(D As Drs, FF$) As Drs
'@D : ..{Gpcc}    ! it has columns-Gpcc
'Ret   : {Gpcc} Cnt  ! each @Gpcc is unique.  Cnt is rec cnt of such gp
Dim GpDy(), Cnt&()
    With GpCnt(D, FF)
        GpDy = .GpDy
        Cnt = .Cnt
    End With
Dim ODy()
    Dim J&, Dr: For Each Dr In Itr(GpDy)
        Push Dr, Cnt(J)
        PushI GpDy, Dr
        J = J + 1
    Next
Dim Fny$(): Fny = SySyEle(D.Fny, "Cnt")
SelDistCnt = Drs(Fny, ODy)
End Function

Function DrsSelFf(D As Drs, FF$) As Drs:          DrsSelFf = DrsSelFny(D, FnyFF(FF)):       End Function
Function DrsSelFfMay(D As Drs, FfMay$) As Drs: DrsSelFfMay = DrsSelFnyMay(D, FnyFF(FfMay)): End Function

Function DrsSelFnyMay(D As Drs, FnyMay$()) As Drs
If IsEqAy(D.Fny, FnyMay) Then DrsSelFnyMay = D: Exit Function
DrsSelFnyMay = Drs(FnyMay, DySelMay(D.Dy, InySubayThw(D.Fny, FnyMay)))
End Function

Function DrsSelFfAs(D As Drs, FfAs$) As Drs
Dim FnyA$(), FnyB$(): AsgTmlFldMap FfAs, FnyA, FnyB
DrsSelFfAs = Drs(FnyB, DrsSelFny(D, FnyA).Dy)
End Function

Function DrsSelAtEndFf(D As Drs, AtEndFf$) As Drs
Dim NewFny$(): NewFny = RseqFnyEnd(D.Fny, SySs(AtEndFf))
DrsSelAtEndFf = DrsSelFny(D, NewFny)
End Function

Function DrsSelInFrontFf(D As Drs, InFrontFf$) As Drs
Dim NewFny$(): NewFny = RseqFnyFront(D.Fny, SySs(InFrontFf))
DrsSelInFrontFf = DrsSelFny(D, NewFny)
End Function

Function DrsSelExlCCLik(D As Drs, ExlCCLik$) As Drs
Stop
Dim LikC: For Each LikC In SySs(ExlCCLik)
'    AyMinus(
Next
End Function

Function DrsSelFny(D As Drs, Fny$()) As Drs
Const CSub$ = CMod & "DrsSelFny"
ChkIsSupAy CSub, D.Fny, Fny
Dim CiySel%(): CiySel = InyFnySub(D.Fny, Fny)
DrsSelFny = Drs(Fny, DySel(D.Dy, CiySel))
End Function

Function DtSelFf(D As Dt, FF$) As Dt
DtSelFf = DtDrs(DrsSelFf(DrsDt(D), FF), D.Dtn)
End Function

Function DrsUpdColV(D As Drs, C$, V) As Drs
Dim I&: I = IxEle(D.Fny, C)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    Dr(I) = V
    PushI Dy, Dr
Next
DrsUpdColV = Drs(D.Fny, Dy)
End Function

Function DrsUpdP12(D As Drs, P12$, V1, V2) As Drs
Dim I1&, I2&: AsgCxapDrs D, P12, I1, I2
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    Dr(I1) = V1
    Dr(I2) = V2
    PushI Dy, Dr
Next
DrsUpdP12 = Drs(D.Fny, Dy)
End Function

Function DrsUpd(A As Drs, B As Drs, Jn$, Upd$, IsLefJn As Boolean) As Drs
'@A  : ..@Jn-LHS..@Upd-LHS.. ! to be updated
'@B  : ..@Jn-RHS..@Upd-RHS.. ! used to update @A.@Upd-LHS
'@Jn : :SS-JnTm            ! :JnTm is :ColonTm.  LHS is @A-fld and RHS is @B-fld
'Fm Upd : :Upd-UpdTm          ! :UpdTer: is :ColTm.  LHS is @A-fld and RHS is @B-fld
'Ret    : sam as @A             ! new Drs from @A with @A.@Upd-LHS updated from @B.@Upd-RHS. @@
Dim C As Dictionary: Set C = DiFmDrs(B)
Dim O As Drs
    O.Fny = A.Fny
    Dim Dr, K
    For Each Dr In A.Dy
        K = Dr(0)
        If C.Exists(K) Then
            Dr(0) = C(K)
        End If
        PushI O.Dy, Dr
    Next
DrsUpd = O
'BrwDrs3 A, B, O, NN:="A B O", Tit:= _
Stop
End Function


Private Sub B_DwDup()
Dim D As Drs, FF$, Act As Drs
GoSub T0
Exit Sub
T0:
    D = DrsFf("D B C", Av(Av(1, 2, "xxx"), Av(1, 2, "eyey"), Av(1, 2), Av(1), Av(Empty, 2)))
    FF = "A B"
    GoTo Tst
Tst:
    Act = DwDup(D, FF)
    VcDrs Act
    Return
End Sub

Private Sub B_SelDist()
'BrwDrs SelDistCnt(PFunDrs, "Mdn Ty")
End Sub

Function DePatn(D As Drs, C$, ExlPatn$) As Drs
If ExlPatn = "" Then DePatn = D: Exit Function
Dim ODy()
Dim R As RegExp: Set R = Rx(ExlPatn)
Dim Ix%: Ix = IxEle(D.Fny, C)
Dim Dr: For Each Dr In Itr(D.Dy)
    Dim V: V = Dr(Ix)
    If Not R.Test(V) Then
        PushI ODy, Dr
    End If
Next
DePatn = Drs(D.Fny, ODy)
End Function


Function DyWhEq(Dy(), C&, Eq) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then PushI DyWhEq, Dr
Next
End Function

Function DyWhEqVy(Dy(), Ixy&(), Vy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If IsEqAy(AwIxy(Dr, Ixy), Vy) Then PushI DyWhEqVy, Dr
Next
End Function

Function DyWhIn(Dy(), C, VyIn) As Variant()
Const CSub$ = CMod & "DyWhIn"
If Not IsArray(VyIn) Then Thw CSub, "Given VyIn is not an array", "Ty-VyIn", TypeName(VyIn)
Dim Dr
For Each Dr In Itr(Dy)
    If HasEle(VyIn, Dr(C)) Then
        PushI DyWhIn, Dr
    End If
Next
End Function

Function DyWhLik(Dy(), C%, Lik) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) Like Lik Then PushI DyWhLik, Dr
Next
End Function

Function DyWhPfx(Dy(), C%, Pfx, Optional Cmp As VbCompareMethod = vbTextCompare) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
   If HasPfx(Dr(C), Pfx, Cmp) Then PushI DyWhPfx, Dr
Next
End Function

Function DyWhSsub(Dy(), C%, Ssub) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If HasSsub(Dr(C), Ssub) Then PushI DyWhSsub, Dr
Next
End Function
