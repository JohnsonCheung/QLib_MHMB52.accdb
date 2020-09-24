Attribute VB_Name = "MxVb_Dta_IxGp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_IxGp."
Type IxGp: Ixy() As Long: End Type ' Deriving(Ctor Ay)
Function IxGpUB&(A() As IxGp): IxGpUB = IxGpSI(A) - 1: End Function
Function IxGpSI&(A() As IxGp): On Error Resume Next: IxGpSI = UBound(A) + 1: End Function
Sub PushIxGp(O() As IxGp, M As IxGp): Dim N&: N = IxGpSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function IxGp(Ixy) As IxGp
With IxGp
    .Ixy = Ixy
End With
End Function

Private Sub B_SamEleIxGp()
Dim Act() As IxGp, Ept() As IxGp, Ay, Ixy&(), J&, Ix&
GoSub ZZ
Exit Sub
ZZ:
    Ay = Array(1, 2, 3, 4, 1, 2, 2, 4)
    Act = SamEleIxGp(Ay)
    For J = 0 To UB(Ay)
        Debug.Print Ay(J);
    Next
    Debug.Print
    Debug.Print "==================="
    Stop
    For J = 0 To IxGpUB(Act)
        Ixy = Act(J).Ixy
        For Ix = 0 To UB(Ixy)
            Debug.Print Ay(Ixy(Ix));
        Next
        Debug.Print
    Next
    Stop
    Return
Tst:
    Act = SamEleIxGp(Ay)
    Return
End Sub
Function SamEleIxGp(Ay) As IxGp() ' return IxGpy of Ay with same ele
Dim O() As IxGp, A: A = AyNw(Ay)
Dim J&: For J = 0 To UB(Ay)
    Dim Ix&: Ix = IxEle(A, Ay(J))
    If Ix = -1 Then
        PushI A, Ay(J)
        PushIxGp O, IxGp(Lngy(J))
    Else
        PushI O(Ix).Ixy, J
    End If
Next
SamEleIxGp = O
End Function
