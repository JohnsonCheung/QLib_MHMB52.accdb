Attribute VB_Name = "MxVb_Ay_Op_HasAyPred"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_HasPred."

Function HasElePredPXTrue(A, PX$, P) As Boolean
Dim X
For Each X In Itr(A)
    If Run(PX, P, X) Then HasElePredPXTrue = True: Exit Function
Next
End Function

Function HasElePredXPTrue(A, XP$, P) As Boolean
If Si(A) = 0 Then Exit Function
Dim X: For Each X In Itr(A)
    If Run(XP, X, P) Then
        HasElePredXPTrue = True
        Exit Function
    End If
Next
End Function
