Attribute VB_Name = "MxIde_Dcl_LnoDclln"
Option Compare Text
Option Explicit

Function LnoDclln%(M As CodeModule, Dclln)
Dim J&: For J = 1 To M.CountOfDeclarationLines
   If M.Lines(J, 1) = Dclln Then LnoDclln = J: Exit Function
Next
End Function
Function LnoPfxDclln%(M As CodeModule, PfxDclln)
Dim J&: For J = 1 To M.CountOfDeclarationLines
   If HasPfx(M.Lines(J, 1), PfxDclln) Then LnoPfxDclln = J: Exit Function
Next
End Function

