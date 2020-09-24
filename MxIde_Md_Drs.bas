Attribute VB_Name = "MxIde_Md_Drs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Drs."
 Public Const FfTMdn4$ = "Pjn CmpTy Mdn NMdLn"
Public Const FfTMdn11$ = FfTMdn4 & " CModv IsCModEr"
  Public Const FfTMd$ = FfTMdn11 & " NMth NPub NPrv NFrd Mthnn"

Private Sub B_MdnDrsP():                                                       BrwDrs DrsTMdn9PC:                   End Sub
Function DrsTMdn9PC(Optional PatnssAndMd$) As Drs:                DrsTMdn9PC = DrsTMdn9P(CPj, PatnssAndMd$):        End Function
Function DrsTMdn9P(P As VBProject, Optional PatnssAndMd$) As Drs:  DrsTMdn9P = DrsFf(FfTMdn4, WDy(P, PatnssAndMd)): End Function
Private Function WDy(P As VBProject, PatnssAndMd$) As Variant()
Dim N: For Each N In AwPatn(Itn(P.VBComponents), PatnssAndMd)
    Push WDy, DrTMdn9(P.VBComponents(N).CodeModule)
Next
End Function
Function DrTMdn9(M As CodeModule) As Variant()
DrTMdn9 = DrMdn(PjnM(M), _
    ShtCmpTyM(M), _
    Mdn(M), _
    M.CountOfLines)
End Function
Function DrMdn(Pjn$, CmpTy$, Mdn$, NMdLn&) As Variant(): DrMdn = Array(Pjn, CmpTy, Mdn, NMdLn): End Function

Private Sub B_DrsTMdn99P():                                     BrwDrs DrsTMdn99P:          End Sub
Function DrsTMdn99P(Optional PatnssAndMd$) As Drs: DrsTMdn99P = DrsTMdn9(CPj, PatnssAndMd): End Function
Function DrsTMdn9(P As VBProject, Optional PatnssAndMd$) As Drs
Const CSub$ = CMod & "DrsTMdn9"
Dim ODy()
    Dim Drs As Drs: Drs = DrsTMdn9P(P, PatnssAndMd)
    Dim IxMdn%: IxMdn = IxEle(Drs.Fny, "Mdn")
    Dim Dr: For Each Dr In Drs.Dy
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim D$(): D = DclM(P.VBComponents(Mdn).CodeModule)
        Dim ICModv$: ICModv = CModv(D)
        Dim IIsCModEr As Boolean: IIsCModEr = ICModv <> Mdn
        'Pjn CmpTy Mdn NMdLn CModv IsCModEr
        PushI ODy, AyAddAp(Dr, ICModv, IIsCModEr)
    Next
Dim Bef As Drs: Bef = DrsFf(FfTMdn11, ODy)
DrsTMdn9 = DrsSrt(Bef)
'Insp CSub, "Before and After sort", "Bef Aft", FmtDrs(Bef), FmtDrs(DrsTMdn9): Stop
End Function

Private Sub B_DrsTMdPC():                                   BrwDrs DrsTMdPC:           End Sub
Function DrsTMdPC(Optional PatnssAndMd$) As Drs: DrsTMdPC = DrsTMdP(CPj, PatnssAndMd): End Function
Function DrsTMdP(P As VBProject, Optional PatnssAndMd$) As Drs
Dim D As Drs: D = DrsTMdn9(P, PatnssAndMd)
Dim Dy()
    Dim IxMdn%: IxMdn = IxEle(D.Fny, "Mdn")
    Dim Dr: For Each Dr In Itr(D.Dy)
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim M As CodeModule: Set M = MdP(P, Mdn)
        Dim S$(): S = SrcM(M)
        With TMthStsSrc(Mdn, S)
            Dim NMth%: NMth = .NPub + .NPrv + .NFrd
            PushI Dy, AyAdd(Dr, Array(NMth, .NPub, .NPrv, .NFrd, Mthnn(S)))
        End With
    Next
DrsTMdP = DrsAddDc(D, "NMth NPub NPrv NFrd Mthnn", Dy)
End Function

