Attribute VB_Name = "MxIde_Mthn_TsyMi8Cmntfbel"
Option Compare Text
Option Explicit

Private Sub TsyMi8CmnffbelModPC__Tst():                           Vc TsyMi8CmnffbelModPC:  End Sub
Function TsyMi8CmnffbelModPC() As String(): TsyMi8CmnffbelModPC = TsyMi8CmnffbelModP(CPj): End Function 'Mntfl = module-mthn-shtty-modifier
Function TsyMi8CmnffbelModP(P As VBProject) As String()  'Method-module-key-name-array
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMod(C) Then
        PushIAy TsyMi8CmnffbelModP, TsyMi8CmntfbelM(C.CodeModule)
    End If
Next
End Function
Private Sub TsyMi8CmnffbelPC__Tst():                        Vc TsyMi8CmnffbelPC:  End Sub
Function TsyMi8CmnffbelPC() As String(): TsyMi8CmnffbelPC = TsyMi8CmnffbelP(CPj): End Function 'Mntfl = module-mthn-shtty-modifier
Function TsyMi8CmnffbelP(P As VBProject) As String()  'Method-module-key-name-array
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy TsyMi8CmnffbelP, TsyMi8CmntfbelM(C.CodeModule)
Next
End Function
Function TsyMi8CmntfbelM(M As CodeModule) As String(): TsyMi8CmntfbelM = TsyMi8CmntfbelS(SrcM(M), Mdn(M), ShtCmpTyM(M)): End Function
Function TsyMi8CmntfbelS(Src$(), Mdn$, ShtCmpTy$) As String()
Dim Bix: For Each Bix In ItrMthix(Src)
    Dim Eix&: Eix = Mtheix(Src, Bix)
    Dim L$: L = Mthln(Src, Bix)
    PushI TsyMi8CmntfbelS, TslMi8Cmntfbel(Bix, Eix, L, Mdn, ShtCmpTy)
Next
End Function
Function TslMi8Cmntfbel$(Bix, Eix&, Mthln, Mdn$, ShtCmpTy$): TslMi8Cmntfbel = JnTabAp(ShtCmpTy, TslMi4Mntf(Mthln, Mdn), Bix, Eix, Mthln): End Function
