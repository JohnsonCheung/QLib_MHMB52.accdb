Attribute VB_Name = "MxDao_Fea_Lim_LnkImp1"
Option Compare Text
Option Explicit
Const WTpBepr_AMis$ = ""
#If Doc Then
'Cml
' Vinp #V-Lvl-Fmt-Er#
' Eu   #Error-Udt# Use as Sfx-Cml showing that is it a Udt
' Et   #Error-Ty#  Use as Sfx-Enum showing that it is a Enum for Error
' Bepr #Bool-Express# a string can evaluated to boolean
' Er   #Error#     Use as Sfx-Varn showing that is an error of type :Ly
#End If
Const CMod$ = "MxDao_Fea_LnkImp."
'-- Dta
Private Enum eTyLnk: eTyLnkFb: eTyLnkFx: End Enum
'-- Src
Private Type TLnkInp: Ix As Integer: Inpn As String: Ffn As String: End Type 'Deriving(Ay Ctor)
Private Type TLnkFb: Ix As Integer: Inpn As String: Tny() As String: End Type 'Deriving(Ay Ctor)
Private Type TLnkFx: Ix As Integer: Inpn As String: Inpnw As String: Stru As String: End Type 'Deriving(Ay Ctor)
Private Type TLnkBepr: Ix As Integer: Tbn As String: Bepr As String: End Type 'Deriving(Ay Ctor)
Private Type TLnkFld: Ix As Integer: Intn As String: Ty As String: Extn As String: End Type 'Deriving(Ay Ctor)
Private Type TLnkStru: Ix As Integer: Stru As String: Fld() As TLnkFld: End Type 'Deriving(Ay Ctor)
Private Type TLnk
    Inp() As TLnkInp
    Fb() As TLnkFb
    Fx() As TLnkFx
    Bepr() As TLnkBepr
    Stru() As TLnkStru
End Type
Private Type TImp: T As String: Map() As FldMap: Bepr As String: End Type 'Deriving(Ctor Ay)
'-- Er
Private Enum eFxwEt: eWsNFnd: End Enum
Private Enum eFbtEt: eFilNFnd: End Enum
Private Enum eInpEt: eFilNFnd: End Enum
'-- EuFb
Private Type EuFbt_FbnDup: Lix As Integer: End Type 'Deriving(Ay Ctor)
Private Type FbnMisEu: Lix As Integer: End Type 'Deriving(Ay Ctor)
Private Type EuFbt_TblDup: Lix As Integer: End Type 'Deriving(Ay Ctor)
Private Type EuFbt_TblMis: Lix As Integer: End Type 'Deriving(Ay Ctor)
Private Type FbEuStruExc: Lix As Integer: End Type 'Deriving(Ay Ctor)
Private Type EuFbt
    FbnMis() As FbnMisEu
    FbnDup() As EuFbt_FbnDup
    FbTblDup() As EuFbt_TblDup
    FbTblMis() As EuFbt_TblMis
    StruMis() As FbEuStruExc
End Type
'-- EuFxw
Private Type EuFxw_FxnDup: Lix As Integer: End Type
Private Type EuFxw_FxnMis: Lix As Integer: End Type
Private Type EuFxw_FxnExc: Lix As Integer: End Type
Private Type EuFxw
    FxnDup As EuFxw_FxnDup
    FxnMis As EuFxw_FxnMis
    FxnExc As EuFxw_FxnExc
End Type
'-- EuInp
Private Type EuInp_InpnDup: Lix As Integer: End Type
Private Type EuInp_InpnMis: Lix As Integer: End Type
Private Type EuInp_FilMis: Lix As Integer: End Type
Private Type EuInp_FilKdDup: Lix As Integer: End Type
Private Type EuInp
    InpnDup() As EuInp_InpnDup
    InpfDup() As EuInp_InpnMis
    InpfMis() As EuInp_FilMis
End Type
'-- EuStru
Private Type EuStru_ADup: Lix As Integer: End Type
Private Type EuStru_AExc: Lix As Integer: End Type
Private Type EuStru_AMis: Lix As Integer: End Type
Private Type EuStru_FldDup: Lix As Integer: End Type
Private Type EuStru_FldNone: Lix As Integer: End Type
Private Type EuStru_FldMis: Lix As Integer: End Type
Private Type EuStru_TblDup: Lix As Integer: End Type
Private Type EuStru_TblExc: Lix As Integer: End Type
Private Type EuStru_TblMis: Lix As Integer: End Type
Private Type EuStru_TyInvdt: Lix As Integer: End Type
Private Type EuStru
    ADup As EuStru_ADup
    AMis() As EuStru_AMis
    AExc() As EuStru_AExc
    FldDup() As EuStru_FldDup
    FldNone() As EuStru_FldDup
    FldMis() As EuStru_FldMis
    TblDup() As EuStru_TblDup
    TblMis() As EuStru_TblMis
    TyInvdt() As EuStru_TyInvdt
    MustHasRecTbl() As TIxLn
End Type
'-- EuBepr
Private Type EuBepr_TblDup: Lix As Integer: End Type
Private Type EuBepr_AMis: Lix As Integer: End Type
Private Type EuBepr: AMis() As EuBepr_AMis: TblDup() As EuBepr_TblDup: End Type
Private Type EuOth:
    NoFxAndNoFb As Boolean
End Type
Private Type Eu: Inp As EuInp: Fxw As EuFxw: Fbt As EuFbt: Stru As EuStru: Bepr As EuBepr: Oth As EuOth: MusHasRecTbl As Boolean: End Type
Private Sub B_LnkImp()
Dim LnkImpSp$(), D As Database
GoSub T0
Exit Sub
T0:
    LnkImpSp = LnkImpSpSamp
    Set D = DbTmp
    GoTo Tst
Tst:
    LnkImp D, LnkImpSp
    Return
End Sub
Private Sub B_W0_0_TLnk()
GoSub T1
Exit Sub
Dim SpLnk$(), Act As TLnk, Ept As TLnk
T1:
    SpLnk = LnkImpSpSamp
    GoTo Tst
Tst:
    Act = W0_0_TLnk(SpLnk)
    Return
End Sub

Sub LnkImp(D As Database, LnkImpSp$())
Const CSub$ = CMod & "LnkImp"
Dim L As TLnk
Dim Eu As Eu
Dim Er$()
Dim Tbl As TLnkTbl
Dim Sqy$():
     L = W0_0_TLnk(LnkImpSp)
    Eu = W0_1_Eu(L)
    Er = W0_2_Msg(Eu): ChkEry Er, CSub
   Tbl = W0_3_TLnkTbl(L)
   Sqy = W0_4_SqyImp(L)
         LnkTLnkTbl D, Tbl   '<==
         RunqSqy D, Sqy         '<==
End Sub

Private Function W0_1_Eu(S As TLnk) As Eu
With W0_1_Eu
    .Bepr = W01_0_Bepr
    .Fbt = W01_1_Fbt(S)
    .Fxw = W01_2_Fxw
    .Inp = W01_3_Inp
    .Oth = W01_4_Oth(S)
    .Stru = W01_5_Stru
Stop '    .MusHasRecTbl = W01_MusHasRecTbl(S)
End With
End Function
Private Function W01_0_Bepr() As EuBepr

End Function
Private Function W01_1_Fbt(S As TLnk) As EuFbt
With W01_1_Fbt
    .FbnDup = W0EuFbt_FbnDup
    .FbnMis = W0EuFbt_FbnMis
    .FbTblDup = W0EuFbt_TblDup
    .FbTblMis = W0EuFbt_TblMis
    .StruMis = W0EuFbt_StruMis
End With
End Function
Private Function W01_2_Fxw() As EuFxw

End Function
Private Function W01_3_Inp() As EuInp

End Function
Private Function W01_4_Oth(S As TLnk) As EuOth
With W01_4_Oth
    .NoFxAndNoFb = W0EuOth_NoFxAndNoFb()
'    .NoFxAndNoFb = WNoFxAndNoFb(Ipx, Ipb)
End With
End Function
Private Function W0EuOth_NoFxAndNoFb() As Boolean
'If Si(Ipx.Dy) > 0 Then Exit Function
'If Si(Ipb.Dy) > 0 Then Exit Function
W0EuOth_NoFxAndNoFb = True
End Function
Private Function W01_5_Stru() As EuStru
With W01_5_Stru
    .ADup = W0EuStru_ADup ' IpsHdStru
    .AMis = W0EuStru_AMis
    .AExc = W0EuStru_AExc
    .FldDup = W0EuStru_FldDup
    .TyInvdt = W0EuStru_TyInvdt
End With
End Function
Private Function W0_2_Msg(E As Eu) As String()
With E
Dim I$(): I = W0Msg_Inp(.Inp)
Dim X$(): X = W0Msg_Fxw(.Fxw)
Dim B$(): B = W0Msg_Fbt(.Fbt)
Dim S$(): S = W0Msg_Stru(.Stru)
Dim W$(): W = W0Msg_Bepr(.Bepr)
Dim O$(): O = W0Msg_Oth(.Oth)
End With
W0_2_Msg = SyAddAp(I, X, B, S, W, O)
End Function

Private Function W0Msg_Oth(E As EuOth) As String()
Dim Er0$: If E.NoFxAndNoFb Then Er0 = ""
Dim Er1$()
PushI W0Msg_Oth, Er0
PushIAy W0Msg_Oth, Er1
End Function


Private Function W0Msg_Inp(E As EuInp) As String()
Dim A$(): A = W0MsgInp_FilKdDup
Dim B$(): B = W0MsgInp_FfnDup
Dim C$(): C = W0MsgInp_FfnMis
W0Msg_Inp = Sy(A, B, C)
End Function

Private Function W0MsgInp_FilKdDup() As String()
End Function
Private Function W0MsgInp_FfnDup() As String()
End Function
Private Function W0MsgInp_FfnMis() As String()
End Function
Private Function W0Msg_Fbt(Eu As EuFbt) As String()

End Function
Private Function W0Msg_Fxw(Eu As EuFxw) As String()
Dim A$(): A = W0MsgFxw_TblDup
Dim B$(): B = W0MsgFxw_FxnDup
Dim C$(): C = W0MsgFxw_FxnMis
Dim D$(): D = W0MsgFxw_WsMis
Dim E$(): E = W0MsgFxw_WsMisFld
Dim F$(): F = W0MsgFxw_WsMisFldTy
Dim G$(): G = W0MsgFxw_StruMis
W0Msg_Fxw = SyAddAp(A, B, C, D, E, F, G)
End Function

Private Function W0MsgFxw_TblDup() As String()

End Function
Private Function W0MsgFxw_FxnDup() As String()

End Function
Private Function W0MsgFxw_FxnMis() As String()

End Function
Private Function W0MsgFxw_WsMis() As String()

End Function
Private Function W0MsgFxw_WsMisFld() As String()

End Function
Private Function W0MsgFxw_WsMisFldTy() As String()

End Function
Private Function W0MsgFxw_StruMis() As String()

End Function

Private Function W0EuFbt_StruMis() As FbEuStruExc()

End Function
Private Function W0EuFbt_TblMis() As EuFbt_TblMis()

End Function
Private Function W0EuFbt_TblDup() As EuFbt_TblDup()

End Function
Private Function W0EuFbt_FbnMis() As FbnMisEu()

End Function
Private Function W0EuStru_ADup() As EuStru_ADup

End Function
Private Function W0Msg_Stru(Eu As EuStru) As String()
With Eu
Dim A$(): A = W0MsgStru_Dup
Dim B$(): B = W0MsgStru_Mis
Dim C$(): C = W0MsgStru_Exc
Dim D$(): D = W0MsgStru_NoFld
Dim E$(): E = W0MsgStru_FldDup
Dim F$(): F = W0MsgStru_Ty
End With
W0Msg_Stru = SyAddAp(A, B, C, D, E, F)
End Function
Private Function W0MsgStru_Dup() As String()
End Function
Private Function W0MsgStru_Mis() As String()
End Function
Private Function W0MsgStru_Exc() As String()
End Function
Private Function W0MsgStru_NoFld() As String()
End Function
Private Function W0MsgStru_FldDup() As String()
End Function
Private Function W0MsgStru_Ty() As String()
End Function
Private Function W0Msg_Bepr(E As EuBepr) As String()
With E
Dim A$(): A = W0MsgBepr_TblDup(.TblDup)
Dim B$(): B = W0MsgBepr_TblExa               ' tbl.Bepr is more
Dim C$(): C = W0MsgBepr_Emp                   ' with tbl nm but no Bepr
End With
W0Msg_Bepr = SyAddAp(A, B, C)
End Function

Private Function W0MsgBepr_TblDup(E() As EuBepr_TblDup) As String()

End Function
Private Function W0MsgBepr_TblExa() As String()

End Function
Private Function W0MsgBepr_Emp() As String()

End Function

Private Function W0EuFxw_WsMisFld(IpxfMis As Drs, ActWsf As Drs) As String()
If NoRecDrs(IpxfMis) Then Exit Function
Dim Fxo$(), OFxn$(), OWs$(), O$(), Fxn, Fx$, Ws$, Mis As Drs, Act As Drs, J%, O1$()
AsgCol IpxfMis, "Fxn Fx Ws", OFxn, Fxo, OWs
'---=
PushI O, "Some columns in ws is missing"
For Each Fxn In OFxn
    Fxn = OFxn(J)
    Fx = Fxo(J)
    Ws = OWs(J)
    Stop 'Mis = Dw3EqE(IpxfMis, "Fxn Fx Ws", Fxn, Fx, Ws)
    Stop 'Act = Dw3EqE(ActWsf, "Fxn Fx Ws", Fxn, Fx, Ws)
    '-
    
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Ws     : " & Ws
    X FmtDrsNmv("Mis col: ", Mis)
    X FmtDrsNmv("Act col: ", Act)
    'PushIAy O, AmTab(XX)
    J = J + 1
Next
W0EuFxw_WsMisFld = O
'Insp "QDao_Lnk_ErTLnk.ErTLnk", "Inspect", "Oup(W0EuFxw_WsMisFld) ExWsMisFld IpxfMis ActWsf",ExWsMisFld, ExWsMisFld, FmtDrs(IpxfMis), FmtDrs(ActWsf): Stop
End Function
Private Function W0EuFxw_FldTyMis(Ipxf As Drs, ActWsf As Drs) As String()
'Fm IpxFld : Fxn Ws Stru Ipxf Ty Fx ! Where HasFx and HasWs and Not HasFld
'Fm WsActf : Fxn Ws Ipxf Ty @@
'Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsf()
'OFxn = AwDis(DcStrDrs(IpXB, "Fxn"))
''---=
'If Si(OFxn) = 0 Then Exit Function
'PushI W0EuFxw_WsMis, "Some expected ws not found"
'For J = 0 To UB(OFxn)
'    Fxn = OFxn(J)
'    Fx = ValDrs(IpXB, "Fxn", Fxn, "Fx")
'    ActWsf = DwEqSel(IpXB, "Fxn", Fxn, "L Ws").Dy
'    Lno = DcLngDy(ActWsf, 0)
'    Ws = SyDyC(ActWsf, 1)
'
'    Act = AmzRmvT1(AwT1(WsAct, Fxn)) '*WsActPerFxn::Sy{WsAct}
'    PushIAy W0EuFxw_WsMis, XMisWs_OneFx(Fxn, Fx, Lno, Ws, Act)
'Next
'Insp "QDao_Lnk_ErTLnk.ErTLnk", "Inspect", "Oup(W0EuFxw_FldTyMis) ExWsMisFldTy Ipxf ActWsf",ExWsMisFldTy, ExWsMisFldTy, FmtDrs(Ipxf), FmtDrs(ActWsf): Stop
End Function

Private Function W0EuFxw_WsMis(IpxMis As Drs, ActWs As Drs) As String()
'@ActWs : Fxn Ws @@
Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsnn$, IpxMisi As Drs, O$()
OFxn = AwDis(DcStrDrs(IpxMis, "Fxn"))
'---=
If Si(OFxn) = 0 Then Exit Function
PushI O, "Some expected ws not found"
For J = 0 To UB(OFxn)
    Fxn = OFxn(J)
    Fx = ValDrs(IpxMis, "Fxn", Fxn, "Fx")
    Stop 'IpxMisi = DwEqSel(IpxMis, "Fxn", Fxn, "L Ws")
    ActWsnn = TmlAy(DcFstDrs(DwEqDrp(ActWs, "Fxn", Fxn)))
    '-
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Act ws : " & ActWsnn
    X FmtDrsNmv("Mis ws : ", IpxMisi)
    Stop
    'PushIAy O, AmTab(XX)
Next
W0EuFxw_WsMis = O
'Insp "QDao_Lnk_ErTLnk.ErTLnk", "Inspect", "Oup(W0EuFxw_WsMis) ExWsMis IpxMis ActWs",ExWsMis, ExWsMis, FmtDrs(IpxMis), FmtDrs(ActWs): Stop
End Function

Private Sub B_ErSpLnkImp()
Brw ErSpLnkImp(LnkImpSpSamp)
End Sub
Private Function ErSpLnkImp(SampSpLnk$()) As String()
ErSpLnkImp = W0_2_Msg(W0_1_Eu(W0_0_TLnk(LnkImpSpSamp)))
End Function
Private Function W0EuInp_FilDup() As EuInp_InpnMis()

End Function

Private Function W0EuInp_InpnDup() As EuInp_InpnDup()
'@Ipf : L FilKd Ffn IsFx HasFfn @@
'Dim Ffn$(): Ffn = DcStrDrs(Ipf, "Ffn")
'Dim Dup$(): Dup = AwDup(Ffn)
'If Si(Dup) = 0 Then Exit Function
'Dim DupD As Drs: DupD = DwIn(Ipf, "Ffn", Dup)
'XBox "Ffn Duplicated"
'XDrs DupD
'XLn
'Stop
End Function

Private Function W0EuInp_FilMis() As EuInp_FilMis()
'@Ipf : L FilKd Ffn IsFx HasFfn @@
'If NoRecDrs() Then Exit Function
'Dim A As Drs: A = DwEq(Ipf, "HasFfn", True) '! L Inp Ffn IsFx HasFfn
'Dim B As Drs: B = Vinp_DrsAddDc_Pth_Fn(A)
'Dim C As Drs: C = DrsSelFf(B, "L FilKd Pth Fn")
'      VinpFfnMis = NmvzDrsO("File missing: ", C, DrsFmto(MaxWdt:=200))

'Insp "QDao_Lnk_ErTLnk.ErTLnk", "Inspect", "Oup(VinpFfnMis) EiFfnMis Ipf",EiFfnMis, EiFfnMis, FmtDrs(Ipf): Stop
End Function

Private Function W0EuInp_FilKdDup() As EuInp_FilKdDup()
'@Ipf : L FilKd Ffn IsFx HasFfn @@
'Dim FilKd$(): FilKd = DcStrDrs(Ipf, "FilKd")
'Dim Dup$(): Dup = AwDup(FilKd)
'If Si(Dup) = 0 Then Exit Function
'Dim DupD As Drs: DupD = DwIn(Ipf, "FilKd", Dup)
'XBox "FilKd Duplicated"
'XDrs DupD
'XLn
'VinpFilKdDup = XX
'Insp "QDao_Lnk_ErTLnk.ErTLnk", "Inspect", "Oup(VinpFilKdDup) EiFilKdDup Ipf",EiFilKdDup, EiFilKdDup, FmtDrs(Ipf): Stop
End Function

Private Function W0EuFxw_FxnMis() As EuFxw_FxnMis()
End Function
Private Function W0EuFbt_FbnDup() As EuFbt_FbnDup()

End Function

Private Function W0EuStru_FldDup() As EuStru_FldDup()
End Function

Private Function W0EuStru_TyInvdt() As EuStru_TyInvdt()
End Function

Private Function W0EuInp_InpTblMis() As String()

End Function


Private Function W0EuStru_Dup(IpsHdStru$()) As String()
'@IpsHdStru :  ! the stru coming from the Ips hd @@
'Insp "QDao_Lnk_ErTLnk.ErTLnk", "Inspect", "Oup(VstruDup) EsSDup IpsHdStru",EsSDup, EsSDup, IpsHdStru: Stop
End Function
Private Function W0EuStru_AMis() As EuStru_AMis()
End Function
Private Function W0EuStru_AExc() As EuStru_AExc()
End Function
Private Function W0EuBepr_TblExc() As EuStru_TblExc()
End Function
Private Function W0EuBepr_TblDup() As EuBepr_TblDup()
End Function
Private Function W0EuBepr_TblMis(Ipw As Drs, Tny$()) As String()
'Fm:Wh@Ipw::Drs{L T Bepr}
Dim OL&(), OT$(), J%, T, Dr, O$()
For Each Dr In Itr(Ipw.Dy)
    T = Dr(1)
    If Not HasEle(Tny, T) Then
        PushI OL, Dr(0)
        PushI OT, T
    End If
Next
'---=
If Si(OL) = 0 Then Exit Function
For J = 0 To UB(OL)
    PushI O, FmtQQ("L#(?) Tbl(?) is not defined.", OL(J), OT(J))
Next
PushI O, vbTab & "Defined tables are:"

For Each T In Itr(Tny)
    PushI O, vbTab & vbTab & T
Next
W0EuBepr_TblMis = O
End Function
Private Function W0MsgBepr_AMis(E() As EuBepr_AMis) As String()
Dim J%: Stop 'For J = 0 To UB(E)
    'With E(J)
Stop '    PushI W0MsgBepr_AMis, FmtQQ(WTpBepr_AMis, .Lix + 1, .Tbn)
'Next
End Function

Private Function W0_0_TLnk(LnkImpSp$()) As TLnk
Dim S As TSpec: S = TSpecSrcInd(LnkImpSp)
Stop
With W0_0_TLnk
    .Inp = W0TLnk_Inp(S)
    .Fb = W0TLnk_Fb(S)
    .Fx = W0TLnk_Fx(S)
    .Stru = W0TLnk_Stru(S)
    .Bepr = W0TLnk_Bepr(S)
End With
End Function
Private Function W0TLnk_Inp(S As TSpec) As TLnkInp()
Dim ILn() As TIxLn: ILn = X_0_TIxLny(S, "Inp")
Dim J%: For J = 0 To UbTIxLn(ILn)
    PushTLnkInp W0TLnk_Inp, W0TLnkInp_Itm(ILn(J))
Next
End Function
Private Function W0TLnkInp_Itm(L As TIxLn) As TLnkInp

End Function
Private Function W0TLnk_Fb(S As TSpec) As TLnkFb()
Dim L() As TIxLn: L = X_0_TIxLny(S, "Fb")
Dim J%: For J = 0 To UbTIxLn(L)
    PushTLnkFb W0TLnk_Fb, W0TLnkFb_Itm(L(J))
Next
End Function
Private Function W0TLnkFb_Itm(L As TIxLn) As TLnkFb
Dim Inpn$, Tny$(), A$
AsgT1r L.Ln, Inpn, A
Tny = SySs(A)
W0TLnkFb_Itm = TLnkFb(L.Ix, Inpn, Tny)
End Function
Private Function W0TLnk_Fx(S As TSpec) As TLnkFx()
Dim ILn() As TIxLn: ILn = X_0_TIxLny(S, "Fx")
Dim J%: For J = 0 To UbTIxLn(ILn)
    PushTLnkFx W0TLnk_Fx, W0TLnkFx_Itm(ILn(J))
Next
End Function
Private Function W0TLnkFx_Itm(L As TIxLn) As TLnkFx
Dim Inpn$, Inpnw$, Stru$, A$
AsgT1r L.Ln, Inpn, A
AsgT1r A, Inpnw, Stru
W0TLnkFx_Itm = TLnkFx(L.Ix, Inpn, Inpnw, Stru)
End Function
Private Function W0TLnk_Stru(S As TSpec) As TLnkStru()
Dim I() As TSpeci: I = SpeciyT(S, "Stru")
Dim J%: For J = 0 To UbTSpeci(I)
    PushTLnkStru W0TLnk_Stru, W0TLnkStru_Itm(I(J))
Next
End Function
Private Function W0TLnkStru_Itm(I As TSpeci) As TLnkStru
Dim Stru$: Stru = I.Specin
Dim Fld() As TLnkFld: Fld = W0TLnkStruItm_FldFld(I.IxLny)
W0TLnkStru_Itm = TLnkStru(I.Ix, Stru, Fld)
End Function
Private Function W0TLnkStruItm_FldFld(L() As TIxLn) As TLnkFld()
Dim Ix%, Intn$, Ty$, Extn$
Dim J%: For J = 0 To UbTIxLn(L)
    Ix = L(J).Ix
    AsgT2r L(J).Ln, Intn, Ty, RmvBktSq(Trim(Extn))
    PushTLnkFld W0TLnkStruItm_FldFld, TLnkFld(Ix, Intn, Ty, Extn)
Next
End Function

Private Function W0TLnk_Bepr(S As TSpec) As TLnkBepr()
Dim ILny() As TIxLn: ILny = X_0_TIxLny(S, "Table.Where")
Dim J%: For J = 0 To UbTIxLn(ILny)
    PushTLnkBepr W0TLnk_Bepr, W0TLnkBepr_Itm(ILny(J))
Next
End Function
Private Function W0TLnkBepr_Itm(L As TIxLn) As TLnkBepr
Dim Tbn$, Bepr$
AsgS12 BrkSpc(L.Ln), Tbn, Bepr
W0TLnkBepr_Itm = TLnkBepr(L.Ix, Tbn, Bepr)
End Function
Private Function W01_MusHasRecTbl(S As TLnk) As TIxLn()
Stop
'W01_MusHasRecTbl = X_0_TIxLny(S, "MustHasRecTbl")
End Function

Private Function W0_3_TLnkTbl(S As TLnk) As TLnkTbl
With W0_3_TLnkTbl
    .Fb = W0TLnkTbl_Tb
    .Fx = W0TLnkTbl_Ws
End With
End Function
Private Function W0TLnkTbl_Tb() As TLnkTblFb()

End Function
Private Function W0TLnkTbl_Ws() As TLnkTblFx()

End Function
Private Function W0_4_SqyImp(L As TLnk) As String()
Stop '
'Dim J%: For J = 0 To UbTImp(U)
'    PushI W0_4_SqyImp, W0Sqy_Sql(U(J))
'Next
End Function
Private Function W0Sqy_Sql$(U As TImp)
With U
Dim X$:       X = QpSelAs(U.Map)
Dim Into$: Into = "#I" & .T
Dim Fm$:     Fm = ">" & .T
W0Sqy_Sql = SqlIntoSelX(Into, Fm, X, .Bepr)
End With
End Function

Private Function TImpAdd(A As TImp, B As TImp) As TImp(): PushTImp TImpAdd, A: PushTImp TImpAdd, B: End Function
Private Sub PushTImpy(O() As TImp, A() As TImp): Dim J&: For J = 0 To UbTImp(A): PushTImp O, A(J): Next: End Sub
Private Sub PushTImp(O() As TImp, M As TImp): Dim N&: N = SiTImp(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function SiTImp&(A() As TImp): On Error Resume Next: SiTImp = UBound(A) + 1: End Function
Private Function UbTImp&(A() As TImp): UbTImp = SiTImp(A) - 1: End Function
Private Function TImp(T, Map() As FldMap, Bepr) As TImp
With TImp
    .T = T
    .Map = Map
    .Bepr = Bepr
End With
End Function
Private Function TLnkInpAdd(A As TLnkInp, B As TLnkInp) As TLnkInp(): PushTLnkInp TLnkInpAdd, A: PushTLnkInp TLnkInpAdd, B: End Function
Private Sub PushTLnkInpy(O() As TLnkInp, A() As TLnkInp): Dim J&: For J = 0 To UbTLnkInp(A): PushTLnkInp O, A(J): Next: End Sub
Private Sub PushTLnkInp(O() As TLnkInp, M As TLnkInp): Dim N&: N = SiTLnkInp(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function SiTLnkInp&(A() As TLnkInp): On Error Resume Next: SiTLnkInp = UBound(A) + 1: End Function
Private Function UbTLnkInp&(A() As TLnkInp): UbTLnkInp = SiTLnkInp(A) - 1: End Function
Private Function TLnkInp(Ix, Inpn, Ffn) As TLnkInp
With TLnkInp
    .Ix = Ix
    .Inpn = Inpn
    .Ffn = Ffn
End With
End Function
Private Function TLnkFbAdd(A As TLnkFb, B As TLnkFb) As TLnkFb(): PushTLnkFb TLnkFbAdd, A: PushTLnkFb TLnkFbAdd, B: End Function
Private Sub PushTLnkFby(O() As TLnkFb, A() As TLnkFb): Dim J&: For J = 0 To UbTLnkFb(A): PushTLnkFb O, A(J): Next: End Sub
Private Sub PushTLnkFb(O() As TLnkFb, M As TLnkFb): Dim N&: N = SiTLnkFb(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function SiTLnkFb&(A() As TLnkFb): On Error Resume Next: SiTLnkFb = UBound(A) + 1: End Function
Private Function UbTLnkFb&(A() As TLnkFb): UbTLnkFb = SiTLnkFb(A) - 1: End Function
Private Function TLnkFb(Ix, Inpn, Tny$()) As TLnkFb
With TLnkFb
    .Ix = Ix
    .Inpn = Inpn
    .Tny = Tny
End With
End Function
Private Function TLnkFxAdd(A As TLnkFx, B As TLnkFx) As TLnkFx(): PushTLnkFx TLnkFxAdd, A: PushTLnkFx TLnkFxAdd, B: End Function
Private Sub PushTLnkFxy(O() As TLnkFx, A() As TLnkFx): Dim J&: For J = 0 To UbTLnkFx(A): PushTLnkFx O, A(J): Next: End Sub
Private Sub PushTLnkFx(O() As TLnkFx, M As TLnkFx): Dim N&: N = SiTLnkFx(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function SiTLnkFx&(A() As TLnkFx): On Error Resume Next: SiTLnkFx = UBound(A) + 1: End Function
Private Function UbTLnkFx&(A() As TLnkFx): UbTLnkFx = SiTLnkFx(A) - 1: End Function
Private Function TLnkFx(Ix, Inpn, Inpnw, Stru) As TLnkFx
With TLnkFx
    .Ix = Ix
    .Inpn = Inpn
    .Inpnw = Inpnw
    .Stru = Stru
End With
End Function
Private Function TLnkBeprAdd(A As TLnkBepr, B As TLnkBepr) As TLnkBepr(): PushTLnkBepr TLnkBeprAdd, A: PushTLnkBepr TLnkBeprAdd, B: End Function
Private Sub PushTLnkBepry(O() As TLnkBepr, A() As TLnkBepr): Dim J&: For J = 0 To UbTLnkBepr(A): PushTLnkBepr O, A(J): Next: End Sub
Private Sub PushTLnkBepr(O() As TLnkBepr, M As TLnkBepr): Dim N&: N = SiTLnkBepr(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function SiTLnkBepr&(A() As TLnkBepr): On Error Resume Next: SiTLnkBepr = UBound(A) + 1: End Function
Private Function UbTLnkBepr&(A() As TLnkBepr): UbTLnkBepr = SiTLnkBepr(A) - 1: End Function
Private Function TLnkBepr(Ix, Tbn, Bepr) As TLnkBepr
With TLnkBepr
    .Ix = Ix
    .Tbn = Tbn
    .Bepr = Bepr
End With
End Function
Private Function TLnkFldAdd(A As TLnkFld, B As TLnkFld) As TLnkFld(): PushTLnkFld TLnkFldAdd, A: PushTLnkFld TLnkFldAdd, B: End Function
Private Sub PushTLnkFldy(O() As TLnkFld, A() As TLnkFld): Dim J&: For J = 0 To UbTLnkFld(A): PushTLnkFld O, A(J): Next: End Sub
Private Sub PushTLnkFld(O() As TLnkFld, M As TLnkFld): Dim N&: N = SiTLnkFld(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function SiTLnkFld&(A() As TLnkFld): On Error Resume Next: SiTLnkFld = UBound(A) + 1: End Function
Private Function UbTLnkFld&(A() As TLnkFld): UbTLnkFld = SiTLnkFld(A) - 1: End Function
Private Function TLnkFld(Ix, Intn, Ty, Extn) As TLnkFld
With TLnkFld
    .Ix = Ix
    .Intn = Intn
    .Ty = Ty
    .Extn = Extn
End With
End Function
Private Function TLnkStruAdd(A As TLnkStru, B As TLnkStru) As TLnkStru(): PushTLnkStru TLnkStruAdd, A: PushTLnkStru TLnkStruAdd, B: End Function
Private Sub PushTLnkStruy(O() As TLnkStru, A() As TLnkStru): Dim J&: For J = 0 To UbTLnkStru(A): PushTLnkStru O, A(J): Next: End Sub
Private Sub PushTLnkStru(O() As TLnkStru, M As TLnkStru): Dim N&: N = SiTLnkStru(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function SiTLnkStru&(A() As TLnkStru): On Error Resume Next: SiTLnkStru = UBound(A) + 1: End Function
Private Function UbTLnkStru&(A() As TLnkStru): UbTLnkStru = SiTLnkStru(A) - 1: End Function
Private Function TLnkStru(Ix, Stru, Fld() As TLnkFld) As TLnkStru
With TLnkStru
    .Ix = Ix
    .Stru = Stru
    .Fld = Fld
End With
End Function

Private Function X_0_TIxLny(S As TSpec, Specit$) As TIxLn()
Dim J%: For J = 0 To UbTSpeci(S.Itms)
    If S.Itms(J).Specit = Specit Then
        X_0_TIxLny = S.Itms(J).IxLny
        Exit Function
    End If
Next
End Function
