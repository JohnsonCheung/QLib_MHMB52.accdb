Attribute VB_Name = "MxDta_Da_Dy_DyNw"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy_DyNw."

Function Dy2AyV(Ay, V) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI Dy2AyV, Array(I, V)
Next
End Function

Function Dy2VAy(V, Ay) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI Dy2VAy, Array(V, I)
Next
End Function
