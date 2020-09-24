Attribute VB_Name = "MxDta_Da_Fmt_Fun"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Fmt_Fun."

Function NoRecMsg(D As Drs, Optional Nm$ = "Drs1") As String()
Dim FF$: FF = JnSpc(D.Fny)
If FF = "" Then FF = " (No Fny)"
FF = FmtQQ("Drs(?) (NoRec) ?", Nm, FF)
NoRecMsg = Sy(FF)
End Function

Function IsEqAyIxy(A, B, Ixy&()) As Boolean
Dim J%: For J = 0 To UB(Ixy)
    If A(Ixy(J)) <> B(Ixy(J)) Then Exit Function
Next
IsEqAyIxy = True
End Function

Function IsLinesDy(Dy()) As Boolean
Dim Dr: For Each Dr In Itr(Dy)
    Dim V: For Each V In Itr(Dr)
        If HasLf(V) Then IsLinesDy = True
    Next
Next
IsLinesDy = False
End Function
