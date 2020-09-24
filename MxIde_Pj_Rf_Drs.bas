Attribute VB_Name = "MxIde_Pj_Rf_Drs"
Option Compare Text
Const CMod$ = "MxIde_Pj_Rf_Drs."
Option Explicit
Public Const FfTRfLn$ = "Pjn Rfn Guid Mjr Mnr Rff"
Public Const FfTRf$ = FfTRfLn & " Des BuiltIn Type IsBroken"
Private Sub B_DrsRfPC():               BrwDrs DrsTRfPC: End Sub
Function DrsTRfPC() As Drs: DrsTRfPC = DrsTRfP(CPj):    End Function
Function DrsTRfP(P As VBProject) As Drs
Const C1$ = "Collection.Parent.Name Name Guid Major Minor FullPath Description BuiltIn TYpe IsBroken"
DrsTRfP = DrsItp(P.References, C1, FfTRf)
End Function
