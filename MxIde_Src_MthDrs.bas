Attribute VB_Name = "MxIde_Src_MthDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_MthDrs."
Public Const FfTMth$ = FfTMdn4$ & " L Mdy Ty Mthn Mthln"
Public Const FfTMthc$ = "Pjn CmpTy Mdn NLn L Mdy Ty Mthn Mthln Mthl"

Private Sub B_DrsTMthP(): BrwDrs DrsTMthPC: End Sub
Private Sub B_DrsTMthM(): BrwDrs DrsTMthMC: End Sub

Function DrsTMthPC() As Drs: DrsTMthPC = DrsTMthP(CPj): End Function
Function DrsTMthMC() As Drs: DrsTMthMC = DrsTMthM(CMd): End Function

Function DrsTMthM(M As CodeModule) As Drs: DrsTMthM = Drs(SySs(FfTMth), WDyTMthM(M)): End Function
Private Function WDyTMthM(M As CodeModule) As Variant()
Dim Src$(): Src = SrcM(M)
'Pjn CmpTy Mdn NMdLn L Mdy Ty Mthn Mthln
Dim Pjn$, ShtCmpTy$, Mdn$, NMdLn&
    Pjn = PjnM(M)
    ShtCmpTy = ShtCmpTyM(M)
    NMdLn = M.CountOfLines
Dim Mthix: For Each Mthix In ItrMthix(Src)
    Dim Mthln$: Mthln = ContlnIx(Src, Mthix)
    Dim Lno&: Lno = Mthix + 1
    With TMthL(Mthln)
        PushI WDyTMthM, Array(Pjn, ShtCmpTy, Mdn, NMdLn, Lno, .ShtMdy, .ShtTy, .Mthn, Mthln)
    End With
Next
End Function

Function DrsTMthP(P As VBProject) As Drs
Dim ODy()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy ODy, WDyTMthM(C.CodeModule)
Next
DrsTMthP = DrsFf(FfTMth, ODy)
End Function
