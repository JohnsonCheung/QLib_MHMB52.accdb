Attribute VB_Name = "MxVb_Ay_Prp_EleFst"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_EleFst."

Function EleFstInAet(Ay, InAet As Dictionary)
Dim I
For Each I In Ay
    If InAet.Exists(I) Then EleFstInAet = I: Exit Function
Next
End Function

Function EleFstLik$(Ay, Lik$)
Dim X: For Each X In Itr(Ay)
    If X Like Lik Then EleFstLik = X: Exit Function
Next
End Function

Function EleFstPredPX(Ay, PX$, P)
Dim X: For Each X In Itr(Ay)
    If Run(PX, P, X) Then
        Asg EleFstPredPX, _
            X
        Exit Function
    End If
Next
End Function

Function EleFstPredXABTrue(Ay, XAB$, A, B)
Dim X
For Each X In Itr(Ay)
    If Run(XAB, X, A, B) Then
        Asg EleFstPredXABTrue, _
            X
        Exit Function
    End If
Next
End Function

Function EleFstPredXP(A, XP$, P)
If Si(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(XP, X, P) Then
        Asg EleFstPredXP, _
            X
        Exit Function
    End If
Next
End Function

Function EleFstwRmvT1$(Sy$(), T1)
EleFstwRmvT1 = RmvA1T(EleFstwT1(Sy, T1))
End Function

Function EleFstwT1$(Ay, T1)
Dim I
For Each I In Itr(Ay)
    If Tm1(I) = T1 Then EleFstwT1 = I: Exit Function
Next
End Function

Function FstPfx$(Pfxy$(), S)
Dim P: For Each P In Pfxy
    If HasPfx(S, P) Then FstPfx = P: Exit Function
Next
End Function

Function EleFstRmvT1$(Sy$(), T1)
EleFstRmvT1 = RmvA1T(EleFstT1(Sy, T1))
End Function

Function EleFstT1$(Sy$(), T1)
Dim S: For Each S In Itr(Sy)
    If HasTmo1(S, T1) Then EleFstT1 = S: Exit Function
Next
End Function

Function EleFstT2$(Sy$(), T2)
Dim S: For Each S In Itr(Sy)
    If HasT2(S, T2) Then EleFstT2 = S: Exit Function
Next
End Function

Function EleFst2T$(Sy$(), T1, T2)
Dim S: For Each S In Itr(Sy)
    If HasTmo2(S, T1, T2) Then EleFst2T = S: Exit Function
Next
End Function
