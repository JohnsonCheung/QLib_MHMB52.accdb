Attribute VB_Name = "MxIde_Pj_Rf_Add"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Rf_Add."

Sub ImpRfP(P As VBProject, PthSrc$)
Dim F$: F = PthEnsSfx(PthSrc) & "Rf.txt"
If NoFfn(F) Then Exit Sub
Dim L: For Each L In Itr(LyFtNB(F))
    WAddRf P, RffLn(L)
Next
End Sub
Private Sub WAddRf(P As VBProject, Rff$)
If HasRff(P, Rff) Then
    Debug.Print Now, "WAddRf", "Rff exists in Pj", P.Name, Rff
    Exit Sub
End If
P.References.AddFromFile Rff
Debug.Print Now, "WAddRf", "Rff added", P.Name, Rff
End Sub
Sub ExpRfP(P As VBProject):            WrtAy SrcRfP(P), FtRfP(P):       End Sub
Sub ExpRfPthP(P As VBProject, PthTo$): WrtAy SrcRfP(P), FtRfPth(PthTo): End Sub
