Attribute VB_Name = "MxIde_Mthn_Mthny"
Option Compare Text
':Mthn: :Nm ! Rule1-FstVerbBeingDo: the mthn will not return any value
'       ! Rule2-FstVerbBeingDo: tThe Cmls aft Do is a verb
':Dta_MthQVNm: :Nm ! It is a String dervied from Nm.  Q for quoted.  V for verb.  It has 3 Patn: NoVerb-[#xxx], MidVerb-[xxx(vvv)xxx], FstVerb-[(vvv)xxx]."
':Nm: :S ! less that 64 chr.
':FunNm: :Rul ! If there is a Subj in pm, put the Subj as fst CmlTm and return that Subj;
'       ! give a Noun to the subj noun is MulCml.
'       ! Each Mthn must belong to one of these rule:
'       !   Noun | Noun.Verb.Extra | Verb.Variant | Noun.z.Variant
'       ! Pm-Rule
'       !   Subj    : Choose a subj in pm if there is more than one arg"
'       !   MuliNoun: It is Ok to group mul-arg as one subj
'       !   MulNounUseOneCml: Mul-noun as one subj use one Cml
':Noun: :Nm  ! it is 1 or more Cml to form a Noun."
':Cml:  :Nm  ! Tag:Type. P1.NumIsLCase:.  P2.LowDashIsLCase:.  P3.fstChrCanAnyNmChr:.
':Sfxx: :SS !  NmRul means variable or function name.
':VdtVerss: :SS ! P1.Opt: Each module may one DoczVdtSSoVerb.  P2.OneOccurance: "
':NounVerbExtra :SS !Tag: FunNmRule.  Prp1.TakAndRetNoun: Fst Cml is Noun and Return Noun.  Prp2.OneCmlNoun: Noun should be 1 Cml.  " & _
'                ! Prp3.VdtVerb: Snd Cml should be approved/valid noun.  Prp4.OptExtra: Extra is optional."
Option Explicit
Const CMod$ = "MxIde_Mthn_Mthny."

Function MthnnMC$():                        MthnnMC = MthnnM(CMd):              End Function
Function MthnnM$(M As CodeModule):           MthnnM = Mthnn(SrcM(M)):           End Function
Function Mthnn$(Src$()):                      Mthnn = JnSpc(AySrt(Mthny(Src))): End Function
Function MthnyVC() As String():             MthnyVC = MthnyV(CVbe):             End Function
Function MthnaetV() As Dictionary:     Set MthnaetV = AetAy(MthnyVC):           End Function
Function MthnyV(V As VBE) As String():       MthnyV = Mthny(SrcV(V)):           End Function

Function MthnaetPC() As Dictionary:          Set MthnaetPC = AetAy(MthnyPC): End Function
Function MthnyPC() As String():                    MthnyPC = MthnyP(CPj):    End Function
Function MthnyP(P As VBProject) As String():        MthnyP = Mthny(SrcP(P)): End Function

Function MthnyMC() As String():               MthnyMC = MthnyM(CMd):    End Function
Function MthnyM(M As CodeModule) As String():  MthnyM = Mthny(SrcM(M)): End Function
Function Mthny(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB Mthny, MthnL(L)
Next
End Function
Private Sub B_Mthny()
GoSub Z
Exit Sub
Z:
   Brw Mthny(SrcVC)
   Return
End Sub


Function MthnyTstPC() As String():              MthnyTstPC = MthnyTstP(CPj):        End Function
Function MthnyTstP(P As VBProject) As String():  MthnyTstP = MthnyTst(MthnyP(P)):   End Function
Function MthnyTst(Mthny$()) As String():          MthnyTst = AwSfx(Mthny, "__Tst"): End Function

Sub MthnBrw_CCmlFmt(): Brw W1_LyPerPj(CPj): End Sub
Private Function W1_LyPerPj(P As VBProject) As String()
Dim O$()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy O, W1_LyPerCmp(C)
Next
W1_LyPerPj = FmtSsy(O)
End Function

Private Function W1_LyPerCmp(C As VBComponent) As String()
Dim N$: N = C.Name
Dim Mthn: For Each Mthn In Itr(Mthny(SrcCmp(C)))
    PushI W1_LyPerCmp, N & " " & CCmlln(Mthn)
Next
End Function
