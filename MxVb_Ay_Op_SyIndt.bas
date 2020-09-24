Attribute VB_Name = "MxVb_Ay_Op_SyIndt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_SyOp."

Function SyIndt(Sy$(), Optional Indt% = 4) As String()
Dim I, S$
S = Space(Indt)
For Each I In Itr(Sy)
    PushI SyIndt, S & I
Next
End Function
