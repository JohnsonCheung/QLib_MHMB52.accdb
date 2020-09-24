Attribute VB_Name = "MxDta_Da_Dy_Op_DyOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy_Op_DyOp."
Function DywDup(Dy(), C) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim Dup As Dictionary: Set Dup = AetAwDup(DcDy(Dy, C))
Dim Dr: For Each Dr In Dy
    If Dup.Exists(Dr(C)) Then PushI DywDup, Dr
Next
End Function
Function DyJn(Dy(), Optional CiySel, Optional Sep$ = " ") As String()
Dim Dy1()
If IsMissing(CiySel) Then
    Dy1 = Dy
Else
    Dy1 = DySel(Dy, CiySel)
End If
Dim Dr: For Each Dr In Itr(Dy1)
    PushI DyJn, RTrim(Jn(Dr, Sep))
Next
End Function
