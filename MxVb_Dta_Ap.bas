Attribute VB_Name = "MxVb_Dta_Ap"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Ap."

Function IntySy(Sy$()) As Integer(): IntySy = IntoyAy(IntyEmp, Sy): End Function
Function LngySy(Sy$()) As Long():    LngySy = IntoyAy(LngyEmp, Sy): End Function

Function IntyLngy(Lngy&()) As Integer()
Dim I: For Each I In Itr(Lngy)
    PushI IntyLngy, I
Next
End Function

Function IntySs(SsInt$) As Integer(): IntySs = IntySy(SySs(SsInt)): End Function
Function LngySs(SsLng$) As Long():    LngySs = LngySy(SySs(SsLng)): End Function

Private Sub B_SyAy()
Dim Act$(): Act = SyAy(Array(1, 2, 3))
Ass Si(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Function SyAy(Ay) As String()
If IsSy(Ay) Then SyAy = Ay: Exit Function
SyAy = IntoyAy(SyEmp, Ay)
End Function

Function SyAyNB(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If I <> "" Then PushI SyAyNB, I
Next
End Function
