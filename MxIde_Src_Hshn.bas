Attribute VB_Name = "MxIde_Src_Hshn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Hshn."
Function RxcHshn() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("#(\w[:.\-\w]*)#")
Set RxcHshn = X
End Function
Function RxcHshnG() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = RxGlobal(RxcHshn)
Set RxcHshnG = X
End Function
Function Hshn$(S):                 Hshn = SsubRx(S, RxcHshn):   End Function
Function Hshny(S) As String():    Hshny = SsubyRx(S, RxcHshnG): End Function
Function HasHshn(S) As Boolean: HasHshn = RxcHshnG.Test(S):     End Function

Private Sub B_HshnyPC():                               BrwAy HshnyPC:   End Sub
Function HshnyPC() As String():              HshnyPC = HshnyP(CPj):     End Function
Function HshnyP(P As VBProject) As String():  HshnyP = Hshny(SrclP(P)): End Function

Private Sub B_HshnlyPC():                                  VcAy HshnlnyPC:     End Sub
Function HshnlnyPC() As String():              HshnlnyPC = HshnlnyP(CPj):      End Function
Function HshnlnyP(P As VBProject) As String():  HshnlnyP = HshnlnyLy(SrcP(P)): End Function
Function HshnlnyLy(Ly$()) As String()
Dim L: For Each L In Itr(Ly)
    If HasHshn(L) Then PushI HshnlnyLy, L
Next
End Function

Private Function WSTup2or0Ln(Ln) As String()
Dim H$: H = Hshn(Ln): If H = "" Then Exit Function
WSTup2or0Ln = Sy(Memn(Ln), H)
End Function

Function HshnokyPC() As String():              HshnokyPC = HshnokyP(CPj):     End Function
Function HshnokyP(P As VBProject) As String():  HshnokyP = Hshnoky(SrclP(P)): End Function
Function Hshnoky(S) As String()
Dim A$(): A = Hshny(S)
Dim N: For Each N In Itr(A)
    If IsHshnOk(N) Then PushI Hshnoky, N
Next
End Function

Function HshneryPC() As String():              HshneryPC = HshneryP(CPj):     End Function
Function HshneryP(P As VBProject) As String():  HshneryP = Hshnery(SrclP(P)): End Function
Function Hshnery(S) As String()
Dim A$(): A = Hshny(S)
Dim N: For Each N In Itr(A)
    If Not IsHshnOk(N) Then PushI Hshnery, N
Next
End Function
Function IsHshnOk(Hshn) As Boolean: IsHshnOk = HasSsub(Hshn, ":"): End Function
