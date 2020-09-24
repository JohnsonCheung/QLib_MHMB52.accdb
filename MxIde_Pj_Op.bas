Attribute VB_Name = "MxIde_Pj_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Op."
Sub ActPj(P As VBProject): Set P.Collection.VBE.ActiveVBProject = P: End Sub

Sub RmvPj(P As VBProject)
Const CSub$ = CMod & "RmvPj"
On Error GoTo X
Dim Pjn$: Pjn = P.Name
P.Collection.Remove P
Exit Sub
X:
Dim E$: E = Err.Description
Inf CSub, FmtQQ("Cannot remove P[?] Er[?]", Pjn, E)
End Sub
