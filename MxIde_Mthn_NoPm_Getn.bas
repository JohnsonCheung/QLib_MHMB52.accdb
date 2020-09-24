Attribute VB_Name = "MxIde_Mthn_NoPm_Getn"
Option Compare Text
Const CMod$ = "MxIde_Mthn_NoPm_Getn."
Option Explicit

Private Sub B_GetnyNoPm()
Brw GetnyNoPm(SrcPC)
End Sub
Function GetnyNoPm(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB GetnyNoPm, GetnNoPm(L)
Next
End Function

Private Sub B_GetnNoPm()
Debug.Assert GetnNoPm("Property Get AA$()") = "AA"
End Sub
Function GetnNoPm$(Ln)
Dim L$: L = RmvMdy(Ln)
If Not ShfPfxSpc(L, "Property Get") Then Exit Function
Dim O$: O = ShfNm(L)
If Left2(RmvTyc(L)) = "()" Then GetnNoPm = O
End Function
