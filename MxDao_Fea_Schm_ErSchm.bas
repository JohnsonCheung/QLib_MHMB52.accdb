Attribute VB_Name = "MxDao_Fea_Schm_ErSchm"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Feat_Schm_Er."
     Const MsgDF_DesMis$ = "*DF_DesMis       L#({L}) Des.Fld({F}) has no des"
  Const MsgDF_FldNotUse$ = "*DF_FldNotUse    L#({L}) Des.Fld({F}) has not used"

     Const MsgDT_DesMis$ = "*DesTbl-DesMsg   L#({L}) Des.Tb{T}) has no des"
    Const MsgDT_TblNDef$ = "*DesTbl-TblNDef  L#({L}) Des.Tbl({T}) is not defined"

    Const MsgDTF_DesMis$ = "*DesTFld-DesMis  L#({L}) Des.TblF{{TF}) has no des"
   Const MsgDTF_FldNDef$ = "*DesTFld-FldNDef L#({L}) Des.TblF Tbl({T}) has Fld({F}) is not defined"
   Const MsgDTF_TblNDef$ = "*DesTF-TblNDef   L#({L}) Des.TblF.Tbl({T}) is not defined TnyAll"

 Const MsgE_EleNoEleStr$ = "*Ele_EleNoStr    L#({L}) Ele({E}) has no EleStr "
   Const MsgE_EleNotUse$ = "*Ele-EleNotUse   L#({L}) Ele({E}) is not used *E_EleNotUse"
    Const MsgE_EleStrEr$ = "*Ele-EleStrEr    L#({L}) Ele({E}) has er in EleStr({EleStr}): {Er}"

   Const MsgEF_EleNoFld$ = "*EF-EleNoFld     L#({L}) Ele({E}) has no Fld"
  Const MsgEF_EleNotUse$ = "*EF-EleNotUse    L#({L}) Ele({E}) has all Likf not used.  The line can be deleted. "
Const MsgEF_LikfNotUsed$ = "*EF-LikfNotUsed  L#({L}) Ele({E}) has LikFld({F}) not used"

    Const MsgSk_FldNDef$ = "*Sk-FldNDef      L#({L}) SkTbl({T}) does not has Fld({F})"
     Const MsgSk_FldDup$ = "*Sk-FldDup       L#({L}) SkTbl({T}) has dup Fld({F})"
    Const MsgSk_TblNDef$ = "*Sk-TblNDef      L#({L}) SkTbl({T}) is not defined"
     Const MsgSk_TblDup$ = "*Sk-TblDup       L#({L}) SkTbl({T}) is duplicated"
   Const MsgSk_TblNoFld$ = "*Sk-TblNoFld     L#({L}) SkTbl({T}) does not has Fld."

      Const MsgT_FldDup$ = "*Tbl-FldDup      L#({L}) Tbl({T}) HasDupFld({F})"
    Const MsgT_FldNoEle$ = "*Tbl-FldNoEle    L#({L}) Tbl({T}) Fld({F}) is not def in EleF nor {StdFldLikss}"
Const MsgT_FldDupMulLin$ = "*Tbl-FldDupMulLn L#({L1}) Tbl({T}) HasDupFld({F}) in L#({L1}) " 'L# has multiple Lno#.  Tbl may be multiple TblNm, Fld is single field
      Const MsgT_LinMis$ = "No Tbl Ln"

Private S As SchmSrc
Private D As SchmDta
Private FnyAll$(), TnyAll$()
Function Schm_Er(Src As SchmSrc, Dta As SchmDta) As String()
S = Src
D = Dta
'Dim X_AllFnyWithEle$(): X_AllFnyWithEle = Fnd_AllFnyWithEle(FnyAll, .EF_FldLiky) 'FnyAll in T has ele
'Dim X_InUseEny$(): X_InUseEny = Fnd_InUseEny(X_AllFnyWithEle, .E_E, .EF_FldLiky)  'T use F, F use E, all E is in use
Dim ErDesFld$(): ErDesFld = WErDesFld
Dim ErDesTbl$(): ErDesTbl = WErDesTbl
Dim ErDesTF$():  ErDesTF = WErDesTF
Dim ErEle$():    ErEle = WErEle
Dim ErEleFld$(): ErEleFld = WErEleFld
Dim ErSk$():     ErSk = WErSk
Dim ErTbl$():    ErTbl = WErTbl
Schm_Er = SyAp(ErDesFld, ErDesTbl, ErDesTF, ErEle, ErEleFld, ErSk, ErTbl)
End Function
Private Function WErDesFld() As String()
Dim E0$():  ' E1 = W1_00_ErDesFld(Lnoy, .DF_F, .DF_D)
Dim E1$():  ' E2 = W1_01_ErFldNotUse(Lnoy, .DF_F)
WErDesFld = Sy(E0, E1)
End Function
Private Function WErDesTbl() As String()
Dim E0$():   E0 = W1_10_DesMis '(.DT_L, .DT_T, .DT_D)
Dim E1$():   E1 = W1_11_ErTblNDef '(.DT_L, .DT_T, .T_T)
End Function
Private Function WErDesTF() As String()
Dim E0$(): 'E0 = W1_20_DesMis(.DTF_L, .DTF_TF, .DTF_D)
Dim E1$(): 'E1 = W1_21_FldNDef(.DTF_L, .DTF_T, .DTF_F)
Dim E2$(): 'E2 = W1_22_TblNDef(.DTF_L, .DTF_T, .T_T)
WErDesTF = Sy(E0, E1, E2)
End Function
Private Function WErEle() As String()

End Function
Private Function WErEleFld() As String()
Dim E0$(): 'E0 = W1_40_EleNoFld(.EF_L, .EF_E, .EF_FldLiky)
Dim E1$(): 'E1 = W1_40_EleNotUse(.EF_L, .EF_E, .EF_FldLiky, FnyAll)
Dim E2$(): 'E2 = W1_40_LikfNotUsed(.EF_L, .EF_E, .EF_FldLiky, FnyAll)

End Function
Private Function WErSk() As String()
Dim Lnoy%(), TnySk$(), FnyySk(), FnyAll$(), TnyAll$()
Dim Sk1$():   Sk1 = W1_50_FldNDef(Lnoy, TnySk, FnyySk, FnyAll)
Dim Sk2$():   Sk2 = W1_51_FldDup(Lnoy, TnySk, FnyySk)
Dim Sk3$():   Sk3 = W1_52_TblNDef(Lnoy, TnySk, TnyAll)
Dim Sk4$():   Sk4 = W1_53_TblDup(Lnoy, TnySk)
Dim Sk5$():   Sk5 = W1_54_TblNoFld(Lnoy, TnySk, FnyySk)
End Function
Private Function WErTbl() As String()
With S
Dim T1$():     T1 = W1_61_FldDup ' (.T_L, .T_T, .T_Fny)
Dim T2$():     T2 = W1_62_FldDupMulLin '(.T_L, .T_T, .T_Fny)
Dim T3$():     T3 = W1_63_FldNoEle '(.T_L, .T_T, .T_Fny, X_AllFnyWithEle)
Dim T4$:       T4 = W1_64_LinMis '(.T_T)
End With
End Function

Private Function W1_00_ErDesFld(L%(), F$(), D$()) As String()
'W1_00_ErDesFld = W1_000_ErDesMis(L, F, D, MsgDF_DesMis)
End Function
Private Function W1_000_ErDesMis(L%(), T_or_F_or_TF$(), D$(), FF$, Msg$) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If D(J) = "" Then
        PushI Dy, Array(L(J), T_or_F_or_TF(J))
    End If
Next
'W1_000_ErDesMis = X_MsgDrs(Msg, DrsFf(Ff, Dy))
End Function

Private Function W1_01_ErFldNotUse(L%(), F$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(FnyAll, F(J)) Then
        PushI Dy, Array(L(J), F(J))
    End If
Next
'W1_01_ErFldNotUse = X_MsgDrs(MsgDF_FldNotUse, DrsFf(FfoLF, Dy))
End Function

Private Function W1_10_DesMis() As String() '(L%(), T$(), D$()) As String()
'W1_10_DesMis = W1_000_ErDesMis(L, T, D, FfoLT, MsgDT_DesMis)
End Function

Private Function W1_11_ErTblNDef() As String() '(L%(), T$()) As String()
Dim Dy()
'Dim J%: For J = 0 To UB(L)
'    If Not HasEle(TnyAll, T(J)) Then
'        PushI Dy, Array(L(J), T(J))
'    End If
'Next
'W1_11_ErTblNDef = X_MsgDrs(MsgDT_TblNDef, DrsFf(FfoLT, Dy))
End Function

Private Function W1_20_DesMis(L%(), TF$(), D$()) As String()
'W1_20_DesMis = W1_000_ErDesMis(L, TF, D, Ffo_L_TF, MsgDTF_DesMis)
End Function

Private Function W1_21_FldNDef(L%(), T$(), F$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(FnyAll, F(J)) Then
        PushI Dy, Array(L(J), T(J), F(J))
    End If
Next
'W1_21_FldNDef = X_MsgDrs(MsgDTF_FldNDef, DrsFf(FfoLTF, Dy))
End Function

Private Function W1_22_TblNDef(L%(), T$(), TnyAll$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(TnyAll, T(J)) Then
        PushI Dy, Array(L(J), T(J))
    End If
Next
'W1_22_TblNDef = X_MsgDrs(MsgDTF_TblNDef, DrsFf(FfoLT, Dy))
End Function

Private Function X_ErE_EleNoStr(L%(), E$(), EleStr$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If EleStr(J) = "" Then
        PushI Dy, Array(L(J), E(J))
    End If
Next
'X_ErE_EleNoStr = X_MsgDrs(MsgE_EleNoEleStr, DrsFf(Ffo_L_E_EleStr_Er, Dy))
End Function

Private Function EoE_EleNotUse(L%(), E$(), InUseEny$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    If Not HasEle(InUseEny, E(J)) Then
        PushI Dy, Array(L(J), E(J))
    End If
Next
'EoE_EleNotUse = X_MsgDrs(MsgE_EleNotUse, DrsFf(FfoLE, Dy))
End Function

Private Function W1_ErEleStr(LnoyEle%(), Eley$(), StryEle$()) As String()
Dim Dy()
'Dim J%: For J = 0 To UB(Lnoy)
'    Dim Er$: Er = ErEleStr(StryEle(J))
'    If Er <> "" Then
'        PushI Dy, Array(L(J), E(J), StryEle(J), Er)
'    End If
'Next
'W_EoE_EleStrEr = X_MsgDrs(MsgE_EleStrEr, DrsFf(Ffo_L_E_StryEle_Er, Dy))
End Function

Private Function W1_ErEleNoFld(L%(), E$(), E_FldLiky()) As String()
Dim Dy()
'W1_40_EleNoFld = X_MsgDrs(MsgEF_EleNoFld, DrsFf(FfoLT, Dy))
End Function

Function HasLik(Ay, Lik) As Boolean
Dim V: For Each V In Itr(Ay)
    If V Like Lik Then HasLik = True: Exit Function
Next
End Function

Private Function W1_41_EleNotUse(L%(), E$(), FldLiky(), FnyAll$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(L)
    Dim I%: For I = 0 To UB(FldLiky)
        If HasLik(FnyAll, FldLiky(I)) Then GoTo X
    Next
    PushI Dy, Array(L(J), E(J))
X:
Next
'W1_41_EleNotUse = X_MsgDrs(MsgEF_EleNotUse, DrsFf(FfoLE, Dy))
End Function

Private Function W1_40_LikfNotUsed() As String() '(L%(), E$(), FldLiky(), FnyAll$()) As String()
Dim Dy()
'Dim J%: For J = 0 To UB(L)
'    Dim Liky$(): Liky = FldLiky(J)
'    Dim I%: For I = 0 To UB(Liky)
'        Dim Lik$: Lik = Liky(J)
'        If Not HasLik(FnyAll, Lik) Then
'            PushI Dy, Array(L(J), E(J), Liky(I))
'        End If
'    Next
'Next
'W1_40_LikfNotUsed = X_MsgDrs(MsgEF_LikfNotUsed, DrsFf(FfoLEF, Dy))
End Function

Private Function W1_50_FldNDef(L%(), TnySku$(), FnyySk(), FnyAll$()) As String()
Dim Dy()
'Dim J%: For J = 0 To UB(L)
'    Dim Fny$(): Fny = FnyySk(J)
'    Dim F: For Each F In Itr(Fny)
'        If Not HasEle(FnyAll, F) Then
'            PushI Dy, Array(L(J), TnySku(J), F)
'        End If
'    Next
'Next
'W1_50_FldNDef = X_MsgDrs(MsgSk_FldNDef, DrsFf(FfoLTF, Dy))
End Function

Private Function X_ErFldDup(Lnoy%(), Tny$(), Fnyy(), Msg$) As String()
Dim Dy()
'Dim J%: For J = 0 To UB(Lnoy)
'    Dim TbFny$(): TbFny = Fny(J)
'    Dim Dup$(): Dup = AwDup(TbFny)
'    Dim F: For Each F In Itr(Dup)
'        PushI Dy, Array(Lnoy(J), Tny(J), F)
'    Next
'Next
'X_ErFldDup = X_MsgDrs(Msg, DrsFf(FfoLTF, Dy))
End Function

Private Function W1_51_FldDup(Lnoy%(), TnySk$(), FnySk()) As String()
W1_51_FldDup = X_ErFldDup(Lnoy, TnySk, FnySk, MsgSk_FldDup)
End Function

Private Function W1_52_TblNDef(Lnoy%(), TnySk$(), TnyAll$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(Lnoy)
    If Not HasEle(TnyAll, TnySk(J)) Then
        PushI Dy, Array(Lnoy(J), TnySk(J))
    End If
Next
'W1_52_TblNDef = X_MsgDrs(MsgSk_TblNDef, DrsFf(FfoLT, Dy))
End Function

Private Function W1_53_TblDup(Lnoy%(), TnySk$()) As String()
Dim Dy()
Dim J%: For J = 0 To UB(Lnoy)
    Dim Dup$(): Dup = AwDup(TnySk)
    Dim D: For Each D In Itr(Dup)
        PushI Dy, Array(Lnoy(J), D)
    Next
Next
'W1_53_TblDup = X_MsgDrs(MsgSk_TblDup, DrsFf(FfoLT, Dy))
End Function

Private Function W1_54_TblNoFld(Lnoy%(), TnySk$(), FnySk()) As String()
Dim Dy()
'Dim J%: For J = 0 To UB(Lnoy)
'    Dim Fny$(): Fny = FnySk(J)
'    If Si(Fny) = 0 Then
'        PushI Dy, Array(Lnoy(J), TnySk(J))
'    End If
'Next
'W1_54_TblNoFld = X_MsgDrs(MsgSk_TblNoFld, DrsFf(FfoLT, Dy))
End Function

Private Function W1_61_FldDup() As String()  '(Lnoy%(), T$(), Fny()) As String()
'W1_61_FldDup = X_ErFldDup(Lnoy, T, Fny, MsgT_FldDup)
End Function

Private Function W1_62_FldDupMulLin() As String() '(Lnoy%(), T$(), Fny()) As String()
'Dim Dy()
'Dim J%: For J = 0 To UB(Lnoy)
'    Dim TbFny$(): TbFny = Fny(J)
'    Dim F: For Each F In Itr(TbFny)
'        Dim I%: For I = 0 To UB(Lnoy)
'            If I <> J Then
'                Dim IFny$(): IFny = Fny(I)
'                If HasEle(IFny, F) Then
'                    PushI Dy, Array(Lnoy(J), T(J), F, I)
'                End If
'            End If
'        Next
'    Next
'Next
'W1_62_FldDupMulLin = X_MsgDrs(MsgT_FldDupMulLin, DrsFf(FfoLTFL, Dy))
End Function

Private Function W1_63_FldNoEle() As String() '(Lnoy%(), T$(), Fny(), AllFnyWithEle$()) As String()
Dim Dy()
'Dim J%: For J = 0 To UB(Lnoy)
'    Dim IFny$(): IFny = Fny(J)
'    Dim F: For Each F In Itr(IFny)
'        If Not HasEle(AllFnyWithEle, F) Then
'            If Not IsStdFld(F) Then
'                PushI Dy, Array(Lnoy(J), T(J), F)
'            End If
'        End If
'    Next
'Next
'W1_63_FldNoEle = X_MsgDrs(MsgT_FldNoEle, DrsFf(FfoLTF, Dy))
End Function

Private Function W1_64_LinMis$() '(Tny$())
'If Si(Tny) = 0 Then W1_64_LinMis = MsgT_LinMis
End Function

Function IsItmInLiky(Itm, Liky) As Boolean
Dim Lik: For Each Lik In Itr(Liky)
    If Itm Like Lik Then IsItmInLiky = True: Exit Function
Next
End Function

Private Function Fnd_AllFnyWithEle(FnyAll$(), EF_FldLiky()) As String()
Dim F: For Each F In Itr(FnyAll)
    Dim J%: For J = 0 To UB(EF_FldLiky)
        Dim Liky$(): Liky = EF_FldLiky(J)
        If IsItmInLiky(F, Liky) Then
            PushI Fnd_AllFnyWithEle, F
            GoTo Nxt
        End If
    Next
Nxt:
Next
End Function

Private Function Fnd_InUseEny(AllFnyWithEle$(), E_E$(), EF_FldLiky()) As String()
'T use F, F use E, all E is in use
Dim F: For Each F In Itr(AllFnyWithEle)
    Dim J%: For J = 0 To UB(EF_FldLiky)
        Dim Liky$(): Liky = EF_FldLiky(J)
        If IsItmInLiky(F, Liky) Then
            PushI Fnd_InUseEny, E_E(J)
            GoTo Nxt
        End If
    Next
Nxt:
Next
End Function

Private Function X_MsgDrs(Macro$, MsgDta As Drs) As String()
Dim Fny$(): Fny = MsgDta.Fny
Dim Dr: For Each Dr In MsgDta.Dy
    PushI X_MsgDrs, FmtMacroDi(Macro, DiFmFnyDr(Fny, Dr))
Next
End Function
