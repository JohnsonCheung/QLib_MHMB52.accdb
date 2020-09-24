Attribute VB_Name = "MxIde_Mth_Lsy"
Option Compare Text
Const CMod$ = "MxIde_Mth_Lsy."
Option Explicit
Private Sub B_LyyMthPC():                                 BrwLyy LyyMthPC: End Sub
Function LyyMthPC() As Variant():              LyyMthPC = LyyMthP(CPj):    End Function
Function LyyMthP(P As VBProject) As Variant():  LyyMthP = LyyMth(SrcPC):   End Function
Function LyyMth(Src$()) As Variant()
Dim Mthix: For Each Mthix In Itr(Mthixy(Src))
    PushI LyyMth, MthyIx(Src, Mthix)
Next
End Function
