Attribute VB_Name = "MxVb_OupDmpVal"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_OupDmpVal."

Sub D(V):      DmpAy LyV(V):            End Sub
Sub Dmp(A):    D A:                     End Sub
Sub DmpTyn(V): Debug.Print TypeName(V): End Sub ' Dmp Tyn of @V
Sub DmpAy(Ay, Optional WithIx As Boolean)  'Dmp Ay with Ix
Dim J&: For J = 0 To UB(Ay)
    If WithIx Then Debug.Print J; ": ";
    Debug.Print Ay(J)
Next
End Sub

Sub OupAy(Ay, OupTy As eOup)
Select Case OupTy
Case eOup.eOupDmp: DmpAy Ay
Case eOup.eOupBrw: BrwAy Ay, "OupAy_"
Case eOup.eOupVc:  VcAy Ay, "VcAy_"
Case Else: ThwPm "OupAy", "OupTy", OupTy, "0 1 2"
End Select
End Sub
