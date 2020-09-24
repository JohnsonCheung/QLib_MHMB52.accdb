Attribute VB_Name = "MxVb_Var_BrwVar"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Var_Brw."

Private Sub LisVal(V, Oup As TOup): LisAy FmtVal(V), Oup:     End Sub
Sub Vc(V, Optional PfxFn$):         LisVal V, TOupVc(PfxFn):  End Sub
Sub Brw(V, Optional PfxFn$):        LisVal V, TOupBrw(PfxFn): End Sub

Function FmtVal(V) As String()
Dim O$()
Select Case True
Case IsLsy(V): O = FmtLsy(CvSy(V))
Case IsStr(V):     O = Sy(V)
Case IsArray(V):   O = SyAy(V)
Case IsAet(V):     O = CvAet(V).Sy
Case IsDi(V):     O = FmtDi(CvDi(V), InlValTy:=True)
Case IsEmpty(V):   O = Sy("#Empty")
Case IsNothing(V): O = Sy("#Nothing")
Case Else:         O = Sy("#TypeName:" & TypeName(V))
End Select
FmtVal = O
End Function
Sub VcAy(Ay, Optional PfxFn$):  LisAy Ay, TOupVc(PfxFn):  End Sub
Sub BrwAy(Ay, Optional PfxFn$): LisAy Ay, TOupBrw(PfxFn): End Sub
Sub LisAy(Ay, O As TOup)
Const CSub$ = CMod & "LisAy"
With O
    Select Case .Oup
    Case eOupBrw, eOupVc
        Dim F$: F = FtTmp("LisAy", .PfxFn)
        WrtAy Ay, F
    End Select
    
    Select Case True
    Case .Oup = eOupBrw: BrwFt F
    Case .Oup = eOupDmp: DmpAy Ay
    Case .Oup = eOupVc:  VcFt F
    Case Else: ThwEnm CSub, .Oup, EnmmOup
    End Select
End With
End Sub


Sub BrwSyIfNB(Sy$())
If Si(Sy) > 0 Then BrwAy Sy
End Sub
Sub BrwStrIfNB(S$)
If Trim(S) <> "" Then BrwStr S
End Sub
Sub BrwStr(S, Optional Fnn$, Optional UseVc As Boolean)
Dim T$: T = FtTmp("BrwStr", Fnn$)
WrtStr S, T
BrwFt T, UseVc
End Sub
Sub VcStr(S, Optional Fnn$)
BrwStr S, Fnn, UseVc:=True
End Sub
