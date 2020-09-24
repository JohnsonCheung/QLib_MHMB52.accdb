Attribute VB_Name = "MxIde_Src_Eix_EixSrcItm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_SrcEixno."
Sub RmvVmk__Tst()
TimFun "TstVmk"
End Sub
Sub TstVmk()
Dim L: For Each L In SrcPC
    RmvVmk L
Next
End Sub
Function EixSrcItm&(Src$(), Ix, SrcItm$) '#Src-End-Ix# :Ix it is an @Src-Ix pointing to the Las-Ln of the mth pointed by @Ix
If Ix < 0 Then EixSrcItm = -1: Exit Function
Const CSub$ = CMod & "EixSrcItm"
Dim Rx As RegExp: Set Rx = WDi_SrcItm_Rx(SrcItm)
Dim O&: For O = Ix To UB(Src)
    Dim L$: L = Src(O)
    If Rx.Test(RmvVmk(L)) Then EixSrcItm = O: Exit Function
Next
Dim Endln$: Endln = "End " & SrcItm
Thw CSub, "Cannot find Endln in @Src from @Ix", "Endln Ix Src", Endln, Ix, AmAddIxPfx(Src, Bix:=0)
End Function
Private Sub WDi_SrcItm_Rx__Tst()
GoSub ZZ
Exit Sub
ZZ:
    Dim O$(), Rx1 As RegExp, Rx2 As RegExp
    Set Rx1 = WDi_SrcItm_Rx("Function")
    Set Rx2 = WDi_SrcItm_Rx("Sub")
    Dim L: For Each L In SrcPC
        Dim Fnd As Boolean
        If Rx1.Test(L) Then
            Fnd = True
        ElseIf Rx2.Test(L) Then
            Fnd = True
        Else
            Fnd = False
        End If
        If Fnd Then PushI O, L
    Next
    VcAy O
    Return
End Sub
Private Function WDi_SrcItm_Rx(SrcItm$) As RegExp
Static Di As New Dictionary
If Not Di.Exists(SrcItm) Then
    Dim Patn$
        Patn = "End " & SrcItm & "$"
    Di.Add SrcItm, Rx(Patn)
End If
Set WDi_SrcItm_Rx = Di(SrcItm)
End Function
Function EnoSrcItm&(M As CodeModule, LnoBeg, SrcItm$)
Const CSub$ = CMod & "EnoSrcItm"
Dim Endln$, O&
Endln = "End " & SrcItm
If HasSsub(M.Lines(LnoBeg, 1), Endln) Then EnoSrcItm = LnoBeg: Exit Function
For O = LnoBeg + 1 To M.CountOfLines
   If HasPfx(M.Lines(O, 1), Endln) Then EnoSrcItm = O: Exit Function
Next
Thw CSub, "Cannot find EndLno", "MthEndLin @LnoBeg @Mdn", Endln, LnoBeg, Mdn(M)
End Function
