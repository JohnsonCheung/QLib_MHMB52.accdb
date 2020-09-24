Attribute VB_Name = "MxTp_Spec"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_Spec."
#If Doc Then
'Ivl:Cml #Invalid# A-kind-of-error
'Exc:Cml #Excess# A-kind-of-error
'Mis:Cml #Miss# A-kind-of-error
'IndSpecSep::Chr #Indt-Spec-Separator# a Vbar before IndSpec
'IndSpec::TmlAy #Indted-Spec# it a Rst-String of Tm4rst of fst line of a Spec.  In format of Spect+ where + means has Pfx* and/or Sfx-
'TSpeci::Udt #Spec-Item# It is 1-Hdr-N-Chd Ly inside a spec
'Spect::Tyn #Spec-Type# It is
'Specit::Tyn #Spec-Item-Type#
'Specity::Ny #Spec-Item-Type-Array#
'Specin:: #Spec-Item-name#
'Lixy:: Line index array.  The index pointing given TpLy
'Tp::   #template#
'Ly::   #line-array#
'SpecFmt:: #Spec-Format#
'  #0 Fmt. 2 or more HChd.
'  .       Fst-HChd::
'  .          Hdr     is { *Spec [Specit] [Specin] [IndSpec] }
'  .          Chd     is Rmk
'  .       Rst-HChd:: TSpeci
'  .          Hdr     is { Specit Specn ..Rmk }
'  .          Chd     is speciILny
'  #1 No D3.  All D3 will be removed before process.
'  .          D3 is D3ln or D3str
'  .          D3ln will remove the line complete
'  .          D3str will remove and the line is kept
'  .          D3Str is string of {{--- XXXX}}
'  .          Str aft D3 will always be removed
'  #2 Rewrite.  If any error, TSpecFt will be rewritten by
'  .            adding D3str at end of a line
'  .            adding D3ln infront only
'  .            in next cycle D3 is removed before process.
'  #3 Fstln.  Aft D3 is removed,
'  .          { *Spec :Spect Specn | IndSpec }
'  '          Fst term: must be *Spec, otherwise throw error
'  .          Snd term: :specTy, must with pfx :, otherwise throw error.
'  .                    The name aft : will be the Spect
'  .          Third Tm: Specn
'  .          Rest: { | IndSpec }, if missing error
'  #4 IndSpec: it is SS of specTy+
'  .           speciTy+ is speciTy with optional pfx * and/or sfx -
'  .           speci is spec-item
'  .           pfx * means the speciTy is a must
'  .           sfx - means the speciTy is at most one
'  #5 HdrChd   [Spec] is Fst-HdrChd + TSpeci-HdrChd
'  .           HChd are group of lines
'  .               Hdr: with first line is not indent (ChrFst ne space)
'  .               Chd: following lines that having fsrChr eq space
'  #6 FstHdrChd hcHdr is see #3 Fstln. where hc is HdrChd.
'  .            hcChd is all remark
'  #7 RstHdrChd rest of HdrChd are speci
'  .            hcHdr is [Specit] [Specin] [ShtRmk]
'  .            hcChd is ILny.  The ln is Trim and with D3Rmv
'SpecFmt:: #Spec-Format-Specification#  See !EdtSampLnkImp 1
#End If
Private Type HdrChd: HdrIx As Integer: Hdr As String: Chd() As String: End Type ' Deriving(Ctor Ay)
'-- Eu
Private Type EuLvl
    Vdt() As String  ' :Specity which is valid
    Ivl() As Itmxy   ' :TSpeci which is invalid
    End Type
Private Type EuExc
    Sng() As String  ' :Specity which be 0 or 1 speci
    Exc() As Itmxy   ' :TSpeci which is invalid
    End Type
Private Type Eu '#Spec-Error-Udt#  Use to tell what is wrong about the :Spec.  It can convert to :SpecEu by !VVTErSpec together :Ly formated by !FmtSpecEu
    IsLnMis As Boolean
    IsSpeciMis As Boolean
    IsSigMis As Boolean
    IsSpectMis As Boolean
    IsSpecnMis As Boolean
    IsIndSpecMis As Boolean
    EuExc As EuExc  ' Specit is Exc er.  In the spec, there is Specit not found in the IndSpec->Must.  It is the Lixy pointing to such TSpeci-Hdr
    EuIvl As EuLvl  ' Specit is Ivl er.  In the spec, there is Specit not found in the IndSpec
    EuMis() As String     ' Specit is Mis er. Miss tyn error.
    IsEr As Boolean         ' If there is any in above error.
End Type

Sub BrwSpec(S As TSpec): Brw FmtTSpec(S): End Sub
Sub VcSpec(S As TSpec):  Vc FmtTSpec(S):  End Sub
Function FmtTSpec(S As TSpec) As String() ' Fmt @S:Spec Hdr + Speciy
FmtTSpec = WhDr(S)
Dim I() As TSpeci: I = S.Itms
Dim J&: For J = 0 To UbTSpeci(I)
    PushIAy FmtTSpec, WSpeci(I(J))
Next
End Function
Private Function WhDr(S As TSpec) As String() 'Fmt the @S:Spec-Hdr = Hdr-Ln + Rmk-Ly
PushI WhDr, "*Spec " & S.Spect & " " & S.Specn & " " & S.IndSpec
PushIAy WhDr, AmAddPfx(S.Rmk, "  ")
End Function
Private Function WSpeci(I As TSpeci) As String() 'Fmt one @I:TSpeci
PushI WSpeci, WItmHdr(I)
PushIAy WSpeci, WItmLLn(I.IxLny)
End Function
Private Function WItmHdr$(I As TSpeci)  ' Fmt spec-item-header-@I as a TSpeci-Hdr-Ln
WItmHdr = I.Specit & " " & I.Specin & " " & I.Rst
End Function
Private Function WItmLLn(L() As TIxLn) As String() ' Fmt speci @L::ILny with 2 space as pfx
WItmLLn = AmAddPfx(LyILny(L), "  ")
End Function

Function TSpecFt(Ft$) As TSpec ' Load Spec from @TSpecFt.  Thw Er if Er
Dim Ly$(): Ly = LyFt(Ft)
Dim S As TSpec: S = TSpecSrcInd(Ly)
Dim E As Eu:   E = W_Eu(S)
If E.IsEr Then
    BkuFfn Ft
    Stop
    'WrtAy FmtSpecEu(Ly, VVTErSpec(E)), TSpecFt, OvrWrt:=True
    Raise "Edit the TSpecFt in notepad as open.  TSpecFt=[" & Ft & "]"
End If
TSpecFt = S
End Function

Private Function VVTErSpec(E As Eu) As TErSpec
'VVTErSpec.Top = W4Top(E)
'VVTErSpec.LnEnd = W4LnEnd(E)
End Function
Private Function W4Top(E As Eu) As String()
Dim A$(), B$(), C$(), D$()
A = W4Top6(E)
B = W4TopSpeciMis(E.EuMis)
C = W4TopSpeciIvl(E)
D = W4TopSpeciExc(E)
W4Top = SyAddAp(A, B, C, D)
End Function
Private Function W4TopSpeciMis(MisSpecit$()) As String()
Const Exc$ = "--- #Exc:: IndSpec indicate that there are this TSpeci should have only 1 such Specit.  Now it is found that there are more than 1.  So they are Exc."
Const TynEr$ = "--- #Specitn-Invalid:: IndSpec indicate that list of valid Specit, but this Specit in not in the list.  So they are TynEr."
End Function
Private Function W4Top6(E As Eu) As String()
Dim O$()
With E
    If .IsIndSpecMis Then PushI O, "IndSpec is missing."
    If .IsLnMis Then PushI O, "No line in TpLy at all"
    If .IsSigMis Then PushI O, "*Spec is missing"
    If .IsSpeciMis Then PushI O, "TSpeci is missing"
    If .IsSpecnMis Then PushI O, "Specn is missng"
    If .IsSpectMis Then PushI O, "Spect is missing"
End With
W4Top6 = O
End Function
Private Function W4TopSpeciIvl(E As Eu) As String()
Stop
End Function
Private Function W4TopSpeciExc(E As Eu) As String()
Stop
End Function
Private Function W4LnEnd(E As Eu) As TIxLn()
W4LnEnd = W4LnEndi(E.EuExc.Exc, X_2SpecitExc)
'PushTIxLnAy W4D3ILnyE, W4D3ILny(E.SpecitIvl, X_2SpecitIvl)
End Function
Private Function W4LnEndi(I() As Itmxy, M$) As TIxLn() ' It is part of :Eu
Dim Ix: For Each Ix In Itr(DisIxyItmxyAy(I))
    PushTIxLn W4LnEndi, TIxLn(Ix, "--- " & M)
Next
End Function

Private Sub B_TSpecSrcInd()
GoSub T1
Exit Sub
Dim Act As TSpec, Ept As TSpec, SrcInd$()
T1:
    SrcInd = SchmSamp(1)
    GoTo Tst
Tst:
    Act = TSpecSrcInd(SrcInd)
    BrwAy FmtTSpec(Act)
    Return
End Sub
Function TSpecSrcInd(SrcInd$()) As TSpec ' Load @SrcInd as :Spec.  If any error find @SrcInd, thw error
'@SrcInd Hdrl is not no hdr space line
'       Chdl is following with space line
'       any D3Msg will be removed
Const CSub$ = CMod & "TSpecSrcInd"
If Si(SrcInd) = 0 Then Thw CSub, "SrcInd is empty"
Dim A$(): A = W2RmvD3(SrcInd$())
If IsLnInd(A(0)) Then Thw CSub, "First chr of first line of SrcInd must not be blank", "SrcInd aft rmv D3", A
Dim B() As HdrChd: B = W2HChdy(A)
TSpecSrcInd = W2SpecHdr(B(0))  ' Fst HChd-element is Spec-Hdr
TSpecSrcInd.Itms = W2TSpeciHChdy(B) ' Rst HChd-element are Spec-Itm
End Function
Private Function W2HChdy(SrcInd$()) As HdrChd()
Const CSub$ = CMod & "W2HChdy"
Dim M As HdrChd, Fst As Boolean: Fst = True
Dim L, Ix%: For Each L In Itr(SrcInd)
    Dim IsHdr As Boolean: IsHdr = Not IsLnInd(L)
    Select Case True
    Case Fst And IsHdr: Fst = False: M = W2WHdrChd(Ix, L)
    Case Fst:           ThwImposs CSub, "First and Hdr is impossible, due to it has been checed Fst Chr Fst Ln must not be blank"
    Case IsHdr: YPushHChd W2HChdy, M
                M = W2WHdrChd(Ix, L)
    Case Else:  PushI M.Chd, LTrim(L)
    End Select
    Ix = Ix + 1
Next
YPushHChd W2HChdy, M
End Function
Private Function W2WHdrChd(HdrIx, Ln) As HdrChd
With W2WHdrChd
    .HdrIx = HdrIx
    .Hdr = Ln
End With
End Function
Private Function W2TSpeciHChdy(I() As HdrChd) As TSpeci()
Dim J%: For J = 1 To XXHChdUB(I)
    PushTSpeci W2TSpeciHChdy, W2Speci(I(J))
Next
End Function
Private Function W2Speci(I As HdrChd) As TSpeci
W2Speci = W2TSpec(I)
W2Speci.IxLny = W2TIxLny(I)
End Function
Private Function W2TSpec(I As HdrChd) As TSpeci
With W2TSpec
    .Ix = I.HdrIx
    Stop 'AsgT3R I.Hdr, .Specit, .Specin, .Rst
    Dim L: For Each L In Itr(I.Chd)
        PushTIxLn W2TSpec.IxLny, TIxLn(.Ix, L)
    Next
End With
End Function
Private Function W2TIxLny(I As HdrChd) As TIxLn() ' Ret ILny by chd-ly & hdr-ix, where  chd-ly is @I.Chd and hdr-ix is @I.Hdrix.  The Chd-ly starts as hdr-ix + 1
Dim J%: For J = 0 To UB(I.Chd)
    PushTIxLn W2TIxLny, TIxLn(I.HdrIx + J + 1, I.Chd(J))
Next
End Function
Private Function W2RmvD3(TpLy$()) As String() ' Rmv all D3ln and D3Ssub
Dim L: For Each L In Itr(TpLy)
    If HasSsub(L, "---") Then
        With Brk1(L, "---", NoTrim:=True)
            If Trim(.S1) <> "" Then
                PushI W2RmvD3, .S1
            End If
        End With
    Else
        PushI W2RmvD3, L
    End If
Next
End Function
Private Function W2SpecHdr(I As HdrChd) As TSpec ' ret a new Spec with Hdr is set.  Hdr is all except .Itms, ie [Spect Specn] & Rmk
Const CSub$ = CMod & "W2SpecHdr"
With W2SpecHdr
    Dim Sig$
    AsgT3r I.Hdr, Sig, .Spect, .Specn, .IndSpec
    If Sig <> "*Spec" Then Thw CSub, "First term of first lline should be *Spec", "@HChd.Hdr", I.Hdr
    .Rmk = I.Chd
End With
End Function

Private Function W2HdrRmk(SrcInd$()) As String() ' ret Hdr rmk, which is fm snd up to next speci-hdr
Dim J%: For J = 1 To UB(SrcInd)
    If ChrFst(SrcInd(J)) <> " " Then Exit Function
    PushI W2HdrRmk, LTrim(SrcInd(J))
Next
End Function
Private Function W2SpeciHdr(Ix%, ItmHdrLn$) As TSpeci
With W2SpeciHdr
    AsgT2r ItmHdrLn, .Specit, .Specin, .Rst
    .Ix = Ix
End With
End Function

Private Function W_Eu(S As TSpec) As Eu ' return er if IndTp has any error
'@IndSpec is a SS with speciTyx as term.
'         the Specitx is Specit with optional * pfx or - sfx.
'         * pfx means must
'         - sfx means single
'         eg AA *BB- *CC DD-
'            means AA is 0-N
'            means BB is 1
'            means CC is 1-N
'            means DD is 0-1
'            VdtNN  will be AA BB CC DD
'            MustNN will be BB CC       (With * pfx)
'            SngNN  will be BB DD       (With - sfx) @@
'@IndTp is lines with HdrLn and ChdLn.  see @SrcInd
Dim Vdt$(), Must$(), Sng$() ' ele of these array are: Specit
    Dim N$(): N = SySs(S.IndSpec)
    Vdt = W3IndSpec_Vdt(N)    ' W1 is returning Specit
    Must = W3IndSpec_Must(N)
    Sng = W3IndSpec_Sng(N)
    
Dim I() As TSpeci: I = S.Itms
Dim O As Eu
With O
    .EuIvl = W3EuIvl(S.Itms, Must)
    .EuMis = W3EuMis(I, Vdt) 'W1 is return error-of-Ixy%() or Misspeciny
    .EuExc = W3EuExc(I, Sng)

    .IsIndSpecMis = S.IndSpec = ""
    '.IsLnMis = S.IsLnMis
    '.IsSigMis = S.IsSigMis
    '.IsSpeciMis = SpeciSi(S.Itms) = 0
    .IsSpecnMis = S.Specn = ""
    .IsSpectMis = S.Spect = ""
    .IsEr = .IsIndSpecMis Or .IsLnMis Or .IsSigMis Or .IsSpeciMis Or .IsSpecnMis Or .IsSpectMis Or _
        X_3IsEr(.EuIvl.Ivl) Or _
        X_3IsEr(.EuExc.Exc) Or _
        Si(.EuMis) > 0
End With
End Function
Private Function W3IndSpec_Vdt(N$()) As String()
Dim I: For Each I In N
    PushI W3IndSpec_Vdt, W3RmvPfxSfx(I)
Next
End Function
Private Function W3IndSpec_Must(N$()) As String() '#
Dim I: For Each I In N
    If HasPfx(I, "*") Then PushI W3IndSpec_Must, W3RmvPfxSfx(I)
Next
End Function
Private Function W3IndSpec_Sng(N$()) As String()
Dim I: For Each I In N
    If HasSfx(I, "-") Then PushI W3IndSpec_Sng, W3RmvPfxSfx(I)
Next
End Function
Private Function W3EuMis(S() As TSpeci, Must$()) As String() ' Missing speciTy
W3EuMis = SyMinus(Must, Specity(S))
End Function
Private Function W3EuExc(S() As TSpeci, Sng$()) As EuExc
W3EuExc.Sng = Sng
Dim Sngi: For Each Sngi In Itr(Sng) 'Sngi:Cml #single-item# Specit which is should have only single :specin
    PushItmxyAy W3EuExc.Exc, W3EuExci(Sngi, S)
Next
End Function
Private Function W3EuExci(Sngi, S() As TSpeci) As Itmxy() '#excess.spec.item-item#

End Function
Private Function W3EuIvl(S() As TSpeci, Vdt$()) As EuLvl ' #excess.spec.item-Item#
W3EuIvl.Vdt = Vdt
Dim Specit: For Each Specit In Itr(Vdt)
    PushItmxy W3EuIvl.Ivl, W3EuIvli(Specit, S)
Next
End Function
Private Function W3EuIvli(Specit, S() As TSpeci) As Itmxy '#invalid.spec.item-item#

End Function
Private Function W3RmvPfxSfx$(N) ' Rmv Pfx * and Sfx -
W3RmvPfxSfx = RmvSfx(RmvPfx(N, "*"), "-")
End Function

Private Function X_1Mis(Mis$(), Must$()) As String()
Dim O$()
PushI O, "Following Specit must exist, but they are missed:"
PushI O, ". Missed  : " & JnSpc(Mis)
PushI O, ". All Must: " & JnSpc(Must)
X_1Mis = O
End Function
Private Function X_1Exc(E As EuExc) As String()
Dim O$()
With E
PushI O, "Following Specit must be single, but they are found more than one:"
Dim J%: For J = 0 To ItmxyUB(.Exc)
    With .Exc(J)
        PushI O, ". Exc Specit / Lno : " & .Itm & " / " & JnSpc(AmInc(.Ixy, 1))
    End With
Next
PushIAy O, " . All sng specit : " & JnSpc(.Sng)
End With
X_1Exc = O
End Function
Private Function X_1Ivl(IvlSpecity$(), VdtSpecity$()) As String()
Dim O$()
PushI O, "Following Specit are invalid:"
PushI O, ". Invalid : " & JnSpc(IvlSpecity)
PushI O, ". Valid   : " & JnSpc(VdtSpecity)
X_1Ivl = O
End Function
Private Function X_12Exc(Exc() As Itmxy) As String()

End Function

Private Function X_2SpecitIvl$(): X_2SpecitIvl = "#Specit-Invalid#: End Function": End Function
Private Function X_2SpecitExc$(): X_2SpecitExc = "#Specit-Exc#:     End Function": End Function

Private Sub X_3___IsEr():                                                                     End Sub
Private Function X_3IsErSpeciExc(E As Eu) As Boolean: X_3IsErSpeciExc = X_3IsEr(E.EuExc.Exc): End Function
Private Function X_3IsErSpeciIvl(E As Eu) As Boolean: X_3IsErSpeciIvl = X_3IsEr(E.EuIvl.Ivl): End Function
Private Function X_3IsEr(I() As Itmxy) As Boolean:            X_3IsEr = ItmxySI(I) > 0:       End Function

Private Function XXHChd(HdrIx, Hdr, Chd$()) As HdrChd
With XXHChd
    .HdrIx = HdrIx
    .Hdr = Hdr
    .Chd = Chd
End With
End Function
Private Function XXAddHChd(A As HdrChd, B As HdrChd) As HdrChd(): YPushHChd XXAddHChd, A: YPushHChd XXAddHChd, B: End Function
Private Sub YPushHChdAy(O() As HdrChd, A() As HdrChd): Dim J&: For J = 0 To XXHChdUB(A): YPushHChd O, A(J): Next: End Sub
Private Sub YPushHChd(O() As HdrChd, M As HdrChd):     Dim N&: N = XXHChdSI(O): ReDim Preserve O(N): O(N) = M:    End Sub
Function XXHChdSI&(A() As HdrChd): On Error Resume Next: XXHChdSI = UBound(A) + 1: End Function
Function XXHChdUB&(A() As HdrChd): XXHChdUB = XXHChdSI(A) - 1: End Function
