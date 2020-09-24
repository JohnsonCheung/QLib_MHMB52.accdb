Attribute VB_Name = "MxVb_Ay__Unclassify_functions"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Ixy."

Private Sub B_AsgCix()
Dim Drs As Drs, FF$, A%, B%, C%, EA%, EB%, Ec%
GoSub T1
Exit Sub
T1:
    Drs.Fny = SySs("A B C")
    FF = "C B A"
    EA = 0
    EB = 1
    Ec = 2
    GoTo Tst
Tst:
    AsgCxapDrs Drs, FF, A, B, C
    Debug.Print A = EA
    Debug.Print B = EB
    Debug.Print C = Ec
    Return
End Sub

Sub AsgCix(Fny$(), FF$, ParamArray OCxap())
Dim Ix, J%: For Each Ix In InySubssThw(Fny, FF)
    OCxap(J) = Ix
    J = J + 1
Next
End Sub

Sub AsgCxapDrs(D As Drs, FF$, ParamArray OCxap())
Dim Ix, J%: For Each Ix In InySubssThw(D.Fny, FF)
    OCxap(J) = Ix
    J = J + 1
Next
End Sub
