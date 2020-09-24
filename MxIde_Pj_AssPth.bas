Attribute VB_Name = "MxIde_Pj_AssPth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_AssPth."

Sub BrwPthAssPC(): BrwPth PthAssPC: End Sub

Function PthAssM$(M As CodeModule): PthAssM = PthAssP(PjM(M)): End Function
Function PthAssP$(P As VBProject)
Static A$
Dim B$: B = PthAss(Pjf(P)): If A <> B Then A = PthEns(B)
PthAssP = A
End Function
Function PthAssPC$()
Static P$: If P = "" Then P = PthAssP(CPj)
PthAssPC = P
End Function
