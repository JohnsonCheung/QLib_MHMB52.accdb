Attribute VB_Name = "MxIde_Mth_Slm_zTool_VcSlmln"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Slmb_SlmLny."
Sub slmzVcSlmlnyPC():                                            VcAy SlmlnyPC:      End Sub
Private Function SlmlnyPC() As String():              SlmlnyPC = SlmlnyP(CPj):       End Function
Private Function SlmlnyP(P As VBProject) As String():  SlmlnyP = SlmlnySrc(SrcP(P)): End Function
Private Function SlmlnySrc(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLnSlm(L) Then PushI SlmlnySrc, L
Next
End Function
