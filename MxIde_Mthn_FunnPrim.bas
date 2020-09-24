Attribute VB_Name = "MxIde_Mthn_FunnPrim"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_FunnPrim."
Private Sub B_FunPfxSpcRstLyPC(): Vc FunPfxSpcRstLyPC: End Sub
Function FunPfxSpcRstLyPC() As String()
Dim Funn: For Each Funn In FunnyPubPC
Stop '    PushI FunPfxSpcRstLyPC, Funn & " " & FunPfxSpcRstLn_Funn(Funn)
Next
End Function
Function FunnPrimzFunn$(Funn): Stop 'FunnPrimzFunn = WFunnPrimzFunnNsegLas(NsegLas(Funn)): End Function

End Function
Function NsegLas$(Nm)
Dim P%: P = InStrRev(Nm, "_")
If P = 0 Then NsegLas = Nm: Exit Function
NsegLas = Mid(Nm, P + 1)
End Function
Function Nsegy(Nm$) As String()
End Function
Function FunPfx_Funn$(Funn)
FunPfx_Funn = Bef(Funn, "_")
If FunPfx_Funn = "" Then
    FunPfx_Funn = CCmlFst(Funn)
End If
End Function
Private Sub B_FunPfxy_PC(): Vc SySrtQ(FunPfxy_PC): End Sub
Function FunPfxy_PC() As String()
Dim N: For Each N In FunnyPubPC
    PushNoDup FunPfxy_PC, FunPfx_Funn(N)
Next
End Function
