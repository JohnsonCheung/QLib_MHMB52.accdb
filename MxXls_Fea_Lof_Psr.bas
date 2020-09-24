Attribute VB_Name = "MxXls_Fea_Lof_Psr"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Fea_Lof_Psr."
Const TpLonLnMis$ = "No line [Lo Nm xxx]"
Const TpLonNmDup$ = "Lon line [Lo Nm xxx] is duplicated.  Lno#-List(?) "
Const TpBetFmToEq$ = "The from and to fields are equal.  Bet-Ln-(Tot <Sum> <Fm> <To>) Lno#(?)"
Const TpBetFldEr$ = "One of <Sum> <Fm> <To> field is not in Fny.  Lno#(?) ErFld(?) Fny(?)"
':FunPfx:Eo: :FunPfx #Error-Of# ! It returns String or Sy.  All variables are given to determine the condition of the error and the message of error to be return.
'                               ! It will call MoXXX to find the message.  & MoXXX will use Const Tp_XXX_.. to build the message
':FunPfx:Mo: :FunPfx #Tp-Of#   ! It takes variables to bld the er message return string or Sy.
':CnstPfx:Tp_: :CnstPfx
':CnstPfx:Ffo:  :CnstPfx        ! It is public constant
Public Const LofiiTot$ = "Sum Avg Cnt"
Public Const LofiiHdr$ = "Lon Fny" 'Iss:Cml :Ss #Itm-Sng-Spc#
Public Const LofiiFval$ = "Fml Lbl Tit Sum" ' Sngsigle field per line
Public Const LofiiValff$ = "Ali Bdr Tot Wdt Fmt Lvl Cor"                             ' Multiple field per line
Public Const LofiiVal$ = LofiiFval & " " & LofiiValff
Public Const Lofii$ = LofiiHdr & " " & LofiiVal
Const LnoStrDig As Byte = 2
Private Type EuDtaLon: A As String: End Type 'Deriving(Ctor)
Private Type EuDtaFny: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaAli: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaWdt: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaBdr: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaLvl: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaCor: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaTot: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaFmt: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaTit: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaFml: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaLbl: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDtaSum: A As String: End Type 'Deriving(Ctor Ay)
Private Type EuDta
    Lon As String
    Fny() As EuDtaFny
    Ali() As EuDtaAli
    Wdt() As EuDtaWdt
    Bdr() As EuDtaBdr
    Lvl() As EuDtaLvl
    Cor() As EuDtaCor
    Tot() As EuDtaTot
    Fmt() As EuDtaFmt
    Tit() As EuDtaTit
    Fml() As EuDtaFml
    Lbl() As EuDtaLbl
    Sum() As EuDtaSum
End Type
Type LofSrc ' Input from Lofly$
    A As String
End Type
Type LofVdt
    Er() As String
    Src As LofSrc
End Type
Private Type PsrRslt: Lofdta As Lofdta: Er() As String: End Type
Const TpDcDrs_Dup$ = "Lno(?) has Dup-?[?]"
Const TpDcDrs_NotIn$ = "Lno(?) has ?[?] which is invalid.  Valid-?=[?]"
Const TpDcDrs_NotNum$ = "Lno(?) has non-numeric-?[?]"
Const TpColx_Blnk$ = "Lno(?) has a blank [?] value"
Const TpColAy_Empty = "Lno(?) has a value of no-element-ay of a column-which-is-an-array"
Const TpColFldLiky_NotInFny$ = "Lno(?) has FldLik[?] not in Fny[?]"
Const TpColNum_NotBet$ = "Lno(?) has ?[?] not between [?] and [?]"
Private Function W_MsgDcDrs_Dup$(Lnoss$, Valn$, Dup): W_MsgDcDrs_Dup = FmtQQ(TpDcDrs_Dup, Lnoss, Valn, Dup): End Function
Private Function W_MsgDcDrs_NotIn(L&, V$, Valn$, VdtValss$):
Stop 'MsgDcDrs_NotIn = FmtQQ(TpDcDrs_NotIn, LnoStr(L), Valn, V, Valn, VdtValss):  End Function

End Function
Private Function W_MsgDcDrs_NotNum$(L&, Valn$, V$):
Stop 'MsgDcDrs_NotNum = FmtQQ(TpDcDrs_NotNum, LnoStr(L), Valn, V):                 End Function
End Function
Private Function W_MsgColNum_NotBet(L&, Valn$, NumV, FmV, ToV):
Stop 'MsgColNum_NotBet = FmtQQ(TpColNum_NotBet, LnoStr(L), Valn, NumV, FmV, ToV): End Function
End Function

Private Function W_MsgColF_3Er(Wi_L_Colx As Drs, Fny$()) As String()
W_MsgColF_3Er = W_MsgColF3Er(Wi_L_Colx, "F", Fny)
End Function

Private Function W_MsgColF3Er(Wi_L_Colx As Drs, ColxNm$, Vy$()) As String()
Dim D As Drs: D = Wi_L_Colx
Dim A$(), B$(), C$(), VV$
VV = JnSpc(Vy)
A = W_MsgColx_NotIn(D, "F", "Fld", VV)
B = W_MsgColx_Dup(D, "F", "Fld")
C = W_MsgColx_Blnk(D, ColxNm)
W_MsgColF3Er = SyAdd(A, B)
End Function

Private Function W_MsgColx_Blnk(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
Dim Valn$: Valn = IIf(Valn0 = "", ColxNm, Valn0)
Dim IxL%: IxL = IxEle(Wi_L_Colx.Fny, ColxNm)
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    If IsLnBlnk(Dr(IxL)) Then
        Dim L&: L = Dr(IxL)
        PushI W_MsgColx_Blnk, FmtQQ(TpColx_Blnk, LnoStr(L), Valn)
    End If
Next
End Function

Private Function W_MsgColFldLiky_3Er(Wi_L_Liky As Drs, Fny$()) As String()
Dim D As Drs: D = Wi_L_Liky
Stop 'ErColFldLiky_3Er = SyAp( _
    ErColAy_Empty(D, "FldLiky"), _
    W_MsgColx_Dup(D, "FldLiky", "FldLik"), _
    ErColFldLiky_NotInFny(D, Fny))
End Function

Private Function W_MsgColAy_Empty(Wi_L_Ay As Drs, ColAyNm$) As String()
Stop 'MsgColAy_Empty = FmtQQ(TpColAy_Empty, LnoStr(L)):                         End Function
Dim IxL%, IxFny%:
Stop 'AsgCix Wi_L_Ay, "L Fny", IxL, IxFny
Dim Dr: For Each Dr In Itr(Wi_L_Ay.Dy)
    Dim Fny$(): Fny = Dr(IxFny)
    If Si(Fny) = 0 Then
        Dim L&: L = Dr(IxL)
        Stop 'PushI W_MsgColAy_Empty, MsgColAy_Empty(L)
    End If
Next
End Function

Private Function W_MsgColx_NotIn(Wi_L_Colx As Drs, ColxNm$, Valn$, VdtValss$) As String()
Dim IxL%, IxColx%:
Stop 'AsgCix Wi_L_Colx, "L " & ColxNm, IxL, 1
Dim VdtVy$(): VdtVy = SySs(VdtValss)
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    Dim V$: V = Dr(IxColx)
    Dim L&: L = Dr(IxL)
    If Not HasEle(VdtVy, V) Then
        Stop 'PushI W_MsgColx_NotIn, MsgDcDrs_NotIn(L, V, Valn, VdtValss)
    End If
Next
End Function

Private Function W_MsgColFldLiky_NotInFny(Wi_L_Liky As Drs, InFny$()) As String()
Stop 'MsgColFldLiky_NotInFny = FmtQQ(TpColFldLiky_NotInFny, LnoStr(L), F, Ff):         End Function
Dim IxFny%, IxL%:
Stop 'AsgCix Wi_L_Liky, "L Fny", IxL, IxFny
Dim FF$: FF = JnSpc(InFny)
Dim Dr: For Each Dr In Itr(Wi_L_Liky.Dy)
    Dim Fny$(): Fny = Dr(IxFny)
    Dim F: For Each F In Fny
        If Not HasEle(InFny, F) Then
            Dim L&: L = Dr(IxL)
            Stop 'PushI ErColFldLiky_NotInFny, MsgColFldLiky_NotInFny(L, f, Ff)
        End If
    Next
Next
End Function

Private Function W_MsgColx_Dup(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
Dim Valn$: Valn = StrDft(Valn0, ColxNm)
Dim Colx():      Colx = DcDrs(Wi_L_Colx, ColxNm)
Dim LnoCol&(): LnoCol = DcLngDrs(Wi_L_Colx, "L")
Dim AllLik$():          'AllLik = CvSy(AyAyy(FldLikyCol))
Dim DupAy$():            DupAy = AwDup(AllLik)
Dim DupLik: For Each DupLik In Itr(DupAy)
    Dim Lnoss$: 'Lnoss = Lnoss_FmLnoCol_WhSyCol_HasS(LnoCol, FldLikyCol, DupLik)
    Stop 'PushI W_MsgColx_Dup, MsgDcDrs_Dup(Lnoss, Valn, DupLik)
Next
If Si(W_MsgColx_Dup) > 0 Then
    Dmp W_MsgColx_Dup
    Stop
End If
End Function

Private Function W_MsgColx_Dup1(Wi_L_Colx As Drs, ColxNm$, Optional Valn0$) As String()
'@Valn :Nm #Val-Nm-ToBe-Shw-InMsg#
Dim Valn$: Valn = StrDft(ColxNm, Valn0)
Dim U%: U = UB(Wi_L_Colx.Dy)
Dim F$():           F = Wi_L_Colx.Fny
Dim Sy$():         Sy = DcStrDrs(Wi_L_Colx, ColxNm)
Dim LnoCol&(): LnoCol = DcLngDrs(Wi_L_Colx, "L")
Dim DupAy$():   DupAy = AwDup(Sy)
Dim Dup: For Each Dup In Itr(DupAy)
    Dim Lnoss$
    Stop ': Lnoss = Lnoss_FmLnoCol_WhSCol_HasS(LnoCol, Sy, Dup)
    Stop 'PushI W_MsgColx_Dup1, MsgDcDrs_Dup(Lnoss, Valn, Dup) '<==
Next
End Function

Private Function W_MsgColx_NumNotBet(Wi_L_Colx As Drs, NumColxNm$, FmV, ToV) As String()
Dim IxNum%, IxL%:
Stop 'AsgCix Wi_L_Colx, JnSpcAp(NumColxNm, "L"), IxNum, IxL
Dim Dr: For Each Dr In Itr(Wi_L_Colx.Dy)
    Dim Num: Num = Val(Dr(IxNum))
    If Not IsBet(Num, FmV, ToV) Then
        Dim L&: L = Dr(IxL)
        Stop 'PushI W_MsgColx_NumNotBet, MsgColNum_NotBet(L, NumColxNm, Num, FmV, ToV)
    End If
Next
End Function

Private Function W_MsgColx_NotNum(Wi_L_Colx As Drs, ColxNm$) As String()
Dim IxL%, IxColxNm%:
Stop 'AsgCix Wi_L_Colx, "L " & ColxNm, IxL, IxColxNm
Dim Dr: For Each Dr In Wi_L_Colx.Dy
    Dim V$: V = Dr(IxColxNm)
    Dim L&
    If Not IsNumeric(V) Then
        L = Dr(IxL)
        Stop 'PushI W_MsgColx_NotNum, MsgDcDrs_NotNum(L, ColxNm, V)
    End If
Next
End Function


Private Sub B_LofPsr()
Dim A As Lofdta: A = LofUd1(LofSamp)
Stop
End Sub

Function LofUd1(Lof$()) As Lofdta
Const CSub$ = CMod & "LofUd1"
With WPsrRslt(Lof)
    ChkEry .Er, CSub, "There is error(s) in given Lofly"
    LofUd1 = .Lofdta
End With
End Function
Private Function WPsrRslt(SrcLof$()) As PsrRslt

End Function

Private Sub WW()
'Lo  Nm  Er    [Lo Nm] has error
'Lo  Nm  Mis   [Lo Nm] line is missed
'Lo  Nm  Dup   [Lo Nm] is Dup
'Lo  Fny Mis   [Lo Fny] is missed
'Lo  Fny Dup   [Lo Fny] is missed
'Ali Val NLis  [Ali Val] is not in @AliVal
'Ali Fld NLis  [Ali Fld] is not in @LoFny
'Bdr Val NLis  [Bdr Val] is not in @BdrVal
'Tot Val NLis  [Tot Val] is not in @TotVal
'Wdt Val NNum  [Wdt Val] is not number
'Wdt Val Mis   [Wdt Val] is missed
'Wdt Val NBet  [Wdt Val] is not between 3 to 100
'Lvl Val NNum  [Lvl Val] is not a number
'Lvl Val NBet  [Lvl Val] is not between 2 and 8
'Lvl Fld NLis  [Lvl Fld] is not in @LoFny
'Lvl Fld Dup
End Sub

Private Function W_MsgLoffBet_SumNotBet(L&, FmFld$, ToFld$, SumFld$)
': MsgLoffBet_SumNotBet = FmtQQ(TpLofBet_SumNotBet, LnoStr(L), FmFld, ToFld, SumFld)
End Function
Private Function W_MsgLoffLon_NmEr(L&, Nm$)
Stop
'MsgLoffLon_NmEr = FmtQQ(TpLon_NmEr, L, Nm)
End Function
Private Function W_MsgLoflon_LinMis$(Dyo_L_Lon())
Stop
'If Si(Dyo_L_Lon) = 0 Then ErLoflon_LinMis = TpLon_LinMis
End Function

Private Function W_MsgLoflon_NmEr(Dyo_L_Lon()) As String()
Dim Dr: For Each Dr In Itr(Dyo_L_Lon)
    Dim Nm$: Nm = Dr(1)
    If Not IsNm(Nm) Then
        Dim L&: L = Dr(0)
        Stop 'PushI W_MsgLoflon_NmEr, MsgLoffLon_NmEr(L, Nm)
    End If
Next
End Function

Private Function W_MsgLoflon_LinDup(Dyo_L_Lon()) As String()
If Si(Dyo_L_Lon) <= 1 Then Exit Function
'Dim Lnoss: For Each Lnoss In Itr(LnossAy)
'    PushI ErLonDup, FmtQQ(C_Lo_ErNm, Lnoss)
'Next
End Function

Private Function W_MsgLofloFld_LinMis(Dyo_L_LoFny()) As String()
If Si(Dyo_L_LoFny) = 0 Then
End If
End Function

Private Function W_MsgLofloFld_FldMis() As String()
'If Si(Dyo_L_LoFny) = 1 Then
'    Dim Fny$(): Fny = Dyo_L_LoFny(0)(1)
'    If Si(Fny) = 0 Then
'    End If
'End If
End Function

Private Function W_MsgLofloFld_FldDupLin(Dyo_L_Fny()) As String()
If Si(Dyo_L_Fny) > 1 Then
    Stop
End If
End Function

Private Sub Er_Tst()
Dim Lof$(), Fny$()
GoSub YY
Exit Sub
YY:
    Stop 'Brw ErLof(LofSamp, LofSampFny)
    Return
T0:
    Fny = SySs("A B C D E F G")
    Ept = Sy()
    GoTo Tst
Tst:
    Stop 'Act = ErLof(Lof, Fny)
    C
    Return
End Sub
Private Function W_MsgLoFny(L_LoFny As Drs) As String()
Dim Dy(), A$(), B$, C$()
Dy = L_LoFny.Dy
 Stop 'A = ErLofloFld_LinMis(Dy)
 Stop
 'B = ErLofloFld_FldMis(Dy)
 Stop 'C = ErLofloFld_FldDupLin(Dy)
W_MsgLoFny = SyApNB(A, B, C)
End Function

Private Function W_MsgLon(L_Lon As Drs) As String()
Dim Dy(), A$(), B$, C$()
Dy = L_Lon.Dy
 A = W_MsgLoflon_NmEr(Dy)
 B = W_MsgLoflon_LinMis(Dy)
 C = W_MsgLoflon_LinDup(Dy)
W_MsgLon = SyApNB(A, B, C)
End Function

Private Function W_MsgAli(L_Ali_FldLiky As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Ali_FldLiky
Dim A$(), B$()
'A = W_MsgColx_NotIn(L_Ali_FldLiky, "Ali", "Ali", Lofaliss)
B = W_MsgColFldLiky_3Er(Drs, Fny)
W_MsgAli = SyApNB(A, B)
End Function

Private Function W_MsgFmt(L_Fmt_FldLiky As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Fmt_FldLiky
Dim A$(), B$(), C$(), D$(), E$()
W_MsgFmt = SyApNB(A)
End Function

Private Function W_MsgLvl(L_Lvl_FldLiky As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Lvl_FldLiky
Dim A$(), B$(), C$(), D$(), E$()
A = W_MsgColx_NumNotBet(Drs, "Lvl", 2, 8)
B = W_MsgColx_NotNum(Drs, "Lvl")
C = W_MsgColFldLiky_3Er(Drs, Fny)
W_MsgLvl = SyApNB(A, B, C)
End Function

Private Function W_MsgCor(L_Cor_FldLiky As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Cor_FldLiky
Dim A$(), B$(), C$(), D$(), E$()
A = W_MsgColx_NumNotBet(Drs, "Cor", 2, 8)
B = W_MsgColx_NotNum(Drs, "Cor")
C = W_MsgColFldLiky_3Er(Drs, Fny)
W_MsgCor = SyApNB(A, B, C)
End Function

Private Function W_MsgWdt(L_Wdt_FldLiky As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Wdt_FldLiky
Dim A$(), B$(), C$()
A = W_MsgColx_NumNotBet(Drs, "Wdt", 5, 100)
B = W_MsgColx_NotNum(Drs, "Wdt")
C = W_MsgColFldLiky_3Er(Drs, Fny)
W_MsgWdt = SyApNB(A, B, C)
End Function

Private Function W_MsgLbl(L_F_Lbl As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_F_Lbl
Dim A$(), B$()
Dim FF$: FF = JnSpc(Fny)
A = W_MsgColx_NotIn(Drs, "F", Valn:="Fld", VdtValss:=FF)
B = W_MsgColx_Dup(Drs, "F", "Fld")
W_MsgLbl = SyApNB(A, B)
End Function

Private Function W_MsgTot(L_Tot_FldLiky As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Tot_FldLiky
Dim A$(), B$(), C$()
A = W_MsgColFldLiky_3Er(L_Tot_FldLiky, Fny)
B = W_MsgColx_NotIn(L_Tot_FldLiky, "Tot", "Tot", Lofii)
End Function

Private Function W_MsgBet(L_Fm_To_Sum As Drs, Fny$()) As String()
Dim A$(), B$(), C$(), FF$, D As Drs
FF = JnSpc(Fny)
D = L_Fm_To_Sum
A = W_MsgColx_NotIn(D, "Fm", "FmFld", FF)
B = W_MsgColx_NotIn(D, "To", "FmFld", FF)
C = W_MsgColx_NotIn(D, "Sum", "SumFld", FF)
Dim IxL%, IxFm%, IxTo%, IxSum%: Stop ' AsgCix L_Fm_To_Sum, "L Fm To Sum", IxL, IxFm, IxTo, IxSum
Dim L&, PosFm%, PosTo%, PosSum%, FmFld$, ToFld$, SumFld$
Dim Dr: For Each Dr In Itr(L_Fm_To_Sum.Dy)
    L = Dr(IxL)
    FmFld = Dr(IxFm)
    ToFld = Dr(IxTo)
    SumFld = Dr(IxSum)
    PosFm = IxEle(Fny, FmFld)
    PosTo = IxEle(Fny, ToFld)
    PosSum = IxEle(Fny, SumFld)
    Dim M$
    If PosFm = PosTo Then
        'M = FmtQQ(TpBetFmToEq, LnoStr(L), FmToFld)
        PushI W_MsgBet, M
        'W_MsgBet_FmToEq(L, FmFld)
        GoTo Nxt
    End If
    If IsBet(PosSum, PosFm, PosTo) Then
        'M = MsgLoffBet_SumNotBet(L, FmFld, ToFld, SumFld)
        PushI W_MsgBet, M
    End If
Nxt:
Next
End Function

Private Function W_MsgLof(Lofly$(), Fny$()) As String()
':Lof:  :Fmtr #ListObj-Fmtr# !
':Fmtr: :Ly   #Formatter#
Dim E As EuDta: Stop 'Dta = ErDtaLofly(Lofly)
With E
Dim ELon$():  Stop '   ELon = ErLon(.L_Lon)
Dim ELoFld$(): Stop ' ELoFld = ErLon(.L_Lon)
Dim F$():  Stop '     F = .Fny .Fny
Dim eAli$(): Stop 'EAli = ErAli(.L_Ali_FldLiky, F)
Dim EWdt$(): Stop 'EWdt = ErWdt(.L_Wdt_FldLiky, F)
Dim eFmt$(): Stop 'EFmt = ErFmt(.L_Fmt_FldLiky, F)
Dim ELvl$(): Stop 'ELvl = ErLvl(.L_Lvl_FldLiky, F)
Dim ECor$(): Stop 'ECor = ErCor(.L_Cor_FldLiky, F)
Dim ETot$(): Stop 'ETot = ErTot(.L_Tot_FldLiky, F)
Dim EBdr$(): Stop 'EBdr = ErBdr(.L_Bdr_FldLiky, F)
Dim EFml$(): Stop 'EFml = ErFml(.L_F_Fml, F)
Dim ELbl$(): Stop 'ELbl = ErLbl(.L_F_Lbl, F)
Dim ETit$(): Stop 'ETit = ErTit(.L_F_Tit, F)
Dim EBet$(): Stop 'EBet = ErBet(.L_Fm_To_Sum, F)
End With
Dim O$(): O = SyApNB(ELon, ELoFld, eAli, EBdr, ETot, EWdt, eFmt, ELvl, ECor, EFml, ELbl, ETit, EBet)
If Si(O) > 0 Then
    W_MsgLof = SyAp(AySrtQ(O), "-----------", AmAddIxPfx(Lofly))
End If
End Function

Private Function W_MsgBdr(L_Bdr_FldLiky As Drs, Fny$()) As String()
Dim Drs As Drs: Drs = L_Bdr_FldLiky
Dim A$(), B$(), C$()
A = W_MsgColFldLiky_3Er(Drs, Fny)
B = W_MsgColx_NotIn(Drs, "Bdr", "Bdr", LofBdrss)
C = W_MsgColx_Dup(Drs, "Bdr", "Bdr")
W_MsgBdr = SyApNB(A, B, C)
End Function

Private Function W_MsgLofBet_FldCannotBetFmTo() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return Eo of M_Bet_* if any
End Function

Private Function Lnoss_FmLnoCol_WhSCol_HasS$(LnoCol&(), DcStrDrs$(), S)
Dim OLno$()
Dim J&: For J = 0 To UB(LnoCol)
    If DcStrDrs(J) = S Then
        PushI OLno, LnoStr(LnoCol(J))
    End If
Next
Lnoss_FmLnoCol_WhSCol_HasS = JnSpc(OLno)
End Function

Private Function Lnoss_FmLnoCol_WhSyCol_HasS$(LnoCol&(), SyCol(), S)
Dim OLno$()
Dim J&: For J = 0 To UB(LnoCol)
    If HasEle(SyCol(J), S) Then
        PushI OLno, LnoStr(LnoCol(J))
    End If
Next
Lnoss_FmLnoCol_WhSyCol_HasS = JnSpc(OLno)
End Function

Private Function LnoStr$(L&): LnoStr = AliR(L, LnoStrDig): End Function

Private Sub B_ErErSrc()
Dim Src$(), Mdn, ErNy$(), ErMthAet As Dictionary
'GoSub T1
'GoSub YY1
GoSub YY2
Exit Sub
T1:

YY2:
    Src = SrcMdn("MXls_Lof_ErLof")
    GoSub Tst
    Brw Act
    Return
YY1:
    GoSub Set_Src
    Mdn = "XX"
    GoSub Tst
    Brw Act
    Return
Tst:
Stop '    Act = ErErSrc(Src, ErMthAet, Mdn)
    Return
Set_Src:
    Const X$ = "'GenErm-Src-Beg." & _
    "|'Val_NotNum      Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number" & _
    "|'Val_NBet      Lno#{Lno} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})" & _
    "|'Val_NotInLis    Lno#{Lno} is [{T1$}] line having invalid Val({ErVal$}).  See valid-value-{VdtValNm$}" & _
    "|'Val_FmlFld      Lno#{Lno} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErFny$()}).  Valid-Fny are [{VdtFny$()}]" & _
    "|'Val_FmlNotBegEq Lno#{Lno} is [Fml] line having [{Fml$}] which is not started with [=]" & _
    "|'Fld_NotInFny    Lno#{Lno} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]" & _
    "|'Fld_Dup         Lno#{Lno} is [{T1$}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno}" & _
    "|'Fldss_NotSel    Lno#{Lno} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]" & _
    "|'Fldss_DupSel    Lno#{Lno} is [{T1$}] line having" & _
    "|'Lon            Lno#{Lno} is [Lo-Nm] line having value({Val$}) which is not a good name" & _
    "|'Lon_Mis        [Lo-Nm] line is missing" & _
    "|'Lon_Dup        Lno#{Lno} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno}" & _
    "|'Tot_DupSel      Lno#{Lno} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored." & _
    "|'Bet_N3Fld        Lno#{Lno} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]" & _
    "|'Bet_EqFmTo      Lno#{Lno} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal." & _
    "|'Bet_FldSeq      Lno#{Lno} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]" & _
    "|'GenErm-Src-End." & _
    "|Const M_Bet_FldSeq$ = 1"
    Src = SplitVBar(X)
    Return
End Sub

Private Function W_MsgMthny(ErNy$()) As String()
Dim I
For Each I In Itr(ErNy)
'    PushI ErMthny, ErMthn(I)
Next
End Function

Private Function W_MsgCnstn$(ErNm)
W_MsgCnstn = "M_" & ErNm
End Function

Private Sub B_ErMthlny()
Dim ErNy$(), ErmAy$(), ErMthlny$()
'GoSub Z
GoSub T1
Exit Sub
Z:
    Brw ErMthlny
    Return
T1:
    ErNy = Sy("Val_NotNum")
    ErmAy = Sy("Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number")
    Ept = Sy("Function W_MsgVal_NotNum(Lno, T1, Val$) As String(): MsgVal_NotNum = FmtMacro(M_Val_NotNum, Lno, T1, Val): End Function")
    GoTo Tst
Tst:
    Act = ErMthlny
    C
    Return
End Sub

Private Function W_MsgMthlny(ErNy$(), ErmAy$()) As String() 'One ErMth is one-MulStmtLin
Dim J%, O$()
For J = 0 To UB(ErNy)
'    PushI O, ErMthLByNm(ErNy(J), MsgAy(J))
Next
'ErMthlny = FmtMulStmtSrc(O)
End Function

Private Function W_MsgMthLByNm$(ErNm$, Erm$)
Dim CNm$:         CNm = W_MsgCnstn(ErNm)
Dim ErNy$():     ErNy = Macrony(Erm)
Dim Pm$:           Pm = JnCmaSpc(AwDis(ErNy))
Dim Calling$: Stop 'Calling = Jn(AmAddPfx(ItmnyDcl(ErNy), ", "))
Dim Mthn:     'Mthn = ErMthn(ErNm)
W_MsgMthLByNm = FmtQQ("Function ?(?) As String():? = FmtMacro(??):End Function", _
    Mthn, Pm, Mthn, CNm, Calling)
End Function
