Attribute VB_Name = "MxIde_Mth_Mthix_MthIxy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Mthix_MthIxy."

Function MthixyMthn(Src$(), Mthn, Optional ShtMthTy$) As Long()
Dim Ix&: Ix = Mthix(Src, Mthn, ShtMthTy): If Ix = -1 Then Exit Function
PushI MthixyMthn, Ix
If IsLnPrp(Src(Ix)) Then
    PushIx MthixyMthn, Mthix(Src, Mthn, ShtMthTy, Ix + 1)
End If
End Function

Function MthixyMthny(Src$(), Mthny$()) As Long()
Dim Ix: For Each Ix In ItrMthix(Src)
    Dim L$: L = Src(Ix)
    Dim N$: N = MthnL(L)
    If HasEle(Mthny, N) Then PushI MthixyMthny, Ix
Next
End Function

Function Mthixy(Src$()) As Long() ' method index array
Dim Ix&: For Ix = 0 To UB(Src)
    Dim L$: L = Src(Ix)
    If IsLnMth(L) Then
        PushI Mthixy, Ix
    End If
Next
End Function
Function MthixyPub(Src$()) As Long() ' method index array
Dim Ix&: For Ix = 0 To UB(Src)
    Dim L$: L = Src(Ix)
    If IsLnMthPub(L) Then
        PushI MthixyPub, Ix
    End If
Next
End Function

Function ItrMthixPub(Src$()): Asg Itr(MthixyPub(Src)), ItrMthixPub: End Function
Function ItrMthix(Src$()): Asg Itr(Mthixy(Src)), ItrMthix: End Function

Private Sub B_MthIxy()
Dim S$()
GoSub Z
Exit Sub
Z:
    S = SrcMC
    Dim MIxy&(): MIxy = Mthixy(S)
    Brw AwIxy(S, MIxy)
    Return

End Sub
