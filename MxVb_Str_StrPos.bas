Attribute VB_Name = "MxVb_Str_StrPos"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_StrPos."
Type TStrpos: Str As String: Pos As Long: End Type
Function TStrpos(Str, Pos) As TStrpos
With TStrpos
    .Str = Str
    .Pos = Pos
End With
End Function
Function PosNxtTstrpos%(A As TStrpos)
If A.Pos = 0 Then Exit Function
PosNxtTstrpos = A.Pos + Len(A.Pos)
End Function
