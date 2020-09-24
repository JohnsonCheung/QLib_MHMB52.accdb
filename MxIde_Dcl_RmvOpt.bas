Attribute VB_Name = "MxIde_Dcl_RmvOpt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_RmvOpt."
Private Sub B_RmvDclOptln()
Dim O() As S12
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushS12 O, W2S12(DclCmp(C))
Next
BrwLsy LsyS12yMix(O)
End Sub
Private Function W2S12(Dcl$()) As S12
Dim Aftl$, Befl$
Aftl = JnCrLf(RmvDclOptln(Dcl))
Befl = JnCrLf(Dcl)
W2S12 = S12(Befl, Aftl)
End Function

Function RmvDclOptln(Dcl$()) As String(): RmvDclOptln = AeBei(Dcl, WBei(Dcl)): End Function
Private Function WBei(Dcl$()) As Bei
Dim B%: B = WBix(Dcl)
WBei = Bei(B, WEix(Dcl, B))
End Function
Private Function WBix%(Dcl$()):       WBix = PfxIx(Dcl, "Option "):    End Function
Private Function WEix%(Dcl$(), Bix%): WEix = NotPfxIx(Dcl, "Option "): End Function
