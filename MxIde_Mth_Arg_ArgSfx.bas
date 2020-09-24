Attribute VB_Name = "MxIde_Mth_Arg_ArgSfx"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Arg_ArgSfx."


Function ShtArgSfx$(A As TArg): ShtArgSfx = ShtArgSfxTVt(A.Vt) & StrPfxIfNB("=", A.Dft): End Function
Function ShtArgSfxTVt$(T As TVt)
Dim O$
With T
Dim Bkt$: Bkt = StrTrue(.IsAy, "()")
Select Case True
Case .Tyc <> "": O = .Tyc & Bkt
Case Else: O = ":" & .Tyn & Bkt
End Select
End With
ShtArgSfxTVt = O
End Function

Private Sub B_S12yDimPC():                              BrwS12y S12yDimPC:                         End Sub
Function S12yDimPC() As S12():              S12yDimPC = S12yDimP(CPj):                             End Function ' S1 is Dimn and S2 is Vsfx
Function S12yDimP(P As VBProject) As S12():  S12yDimP = S12yDim(SySrtQ(SywDis(ItmyDim(SrcP(P))))): End Function
Function S12yDim(ItmyDim$()) As S12()
Dim I: For Each I In Itr(ItmyDim)
    PushS12 S12yDim, S12(TakNm(I), I)
Next
End Function
