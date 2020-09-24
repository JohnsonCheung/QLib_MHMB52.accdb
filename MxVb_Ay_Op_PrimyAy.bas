Attribute VB_Name = "MxVb_Ay_Op_PrimyAy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Cv_Primy."

Function DteyAy(Ay) As Date():   DteyAy = X_Into(DteyEmp, Ay): End Function
Function DblyAy(Ay) As Double(): DblyAy = X_Into(DblyEmp, Ay): End Function

Private Function X_Into(Into, Ay)
If VarType(Into) = VarType(Ay) Then X_Into = Ay: Exit Function
Dim O: O = Into
Dim I: For Each I In Itr(Ay)
    PushI O, I
Next
X_Into = O
End Function
