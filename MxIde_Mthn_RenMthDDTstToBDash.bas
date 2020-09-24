Attribute VB_Name = "MxIde_Mthn_RenMthDDTstToBDash"
Option Compare Text
Const CMod$ = "MxIde_Mthn_RenMthDDTstToBDash."
Option Explicit

Sub RenMthToBDashRptOnly():   RenMthToBDash eRUpdRpt:  End Sub
Sub RenMthToBDashUpdAndRpt(): RenMthToBDash eRUpdBoth: End Sub
Sub RenMthToBDash(Upd As eRUpd)
Dim C As VBComponent: For Each C In CPj.VBComponents
    W00_Ren C, Upd
Next
End Sub
Private Sub W00_Ren(C As VBComponent, Upd As eRUpd)
Dim IsU As Boolean: IsU = IsUpd(Upd)
Dim IsR As Boolean: IsR = IsRpt(Upd)
Dim S$(): S = SrcCmp(C): If Si(S) = 0 Then Exit Sub
Dim Mthix: For Each Mthix In Itr(Mthixy(S))
    Dim Mthln$: Mthln = S(Mthix)
    Dim Nm$: Nm = MthnL(Mthln)
    If HasSfx(Nm, "__Tst") Then
        Dim IsMdnPrinted As Boolean
        If Not IsMdnPrinted Then
            IsMdnPrinted = True
            If IsR Then Debug.Print C.Name
        End If
    
        Dim NmNew$: NmNew = "B_" & RmvSfx(Nm, "__Tst")
        Dim LnNew$: LnNew = MthlnRen(Mthln, Nm, NmNew)
        Dim LnOld$: LnOld = S(Mthix)
        If IsR Then
            Debug.Print , Nm
            Debug.Print , NmNew
            Debug.Print , LnOld
            Debug.Print , LnNew
            Debug.Print
        End If
        If IsU Then C.CodeModule.ReplaceLine Mthix + 1, LnNew
    End If
Next
End Sub
Function MthlnRen$(Mthln$, MthnOld$, MthnNew$)
Const CSub$ = CMod & "MthlnRen"
If Not HasSsub(Mthln, "(") Then Thw CSub, "@Mthln does not have (", "@Mthln", Mthln
With Brk(Mthln, "(")
    If Not HasSsub(.S1, MthnOld) Then Thw CSub, "@Mthn does have @MthnOld", "Mthln MthnOld", Mthln, MthnOld
    MthlnRen = Replace(.S1, MthnOld, MthnNew) & "(" & .S2
End With
End Function
