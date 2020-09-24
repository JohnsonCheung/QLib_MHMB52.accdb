Attribute VB_Name = "MxVb_Ay_Op_IsSupy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Is."
Function IsSupy(Supy, Suby) As Boolean
Dim SubI: For Each SubI In Itr(Suby)
    If NoEle(Supy, SubI) Then Exit Function
Next
IsSupy = True
End Function
