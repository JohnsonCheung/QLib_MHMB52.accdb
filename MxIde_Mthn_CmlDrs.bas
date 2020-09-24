Attribute VB_Name = "MxIde_Mthn_CmlDrs"
':MthCml$ = "NewType:Sy."
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_CmlDrs."

Private Sub B_DrsTMthnCmlPC():                   BrwDrs DrsTMthnCmlPC: End Sub
Function DrsTMthnCmlPC() As Drs: DrsTMthnCmlPC = DrsTMthnCmlP(CPj):    End Function
Function DrsTMthnCmlP(P As VBProject) As Drs
Dim A$()
A = MthnyP(P)
A = AeSfx(A, "__Tst")
A = AePfx(A, "T_")
A = AwDis(A)
A = AySrtQ(A)
DrsTMthnCmlP = DrsTCml(A)
End Function
Function DrsTCml(Ny$()) As Drs
Const CSub$ = CMod & "DrsTCml"
If Si(Ny) = 0 Then Thw CSub, "@Ny cannot be empty"
Dim Dy()
Dim N: For Each N In Itr(Ny)
    PushI Dy, Cmly(N)
Next
Dy = DyReDimDr(Dy)
DrsTCml = Drs(W2Fny(UB(Dy(0))), Dy)
End Function
Private Function W2Fny(UDc%) As String()
PushI W2Fny, "Nm"
Dim J%: For J = 0 To UDc
    PushI W2Fny, "Cml" & J + 1
Next
End Function

Function DyReDimDr(Dy()) As Variant()
If IsEmpAy(Dy) Then Exit Function
Dim UDc%: UDc = WUDc(Dy)
Dim O(): O = Dy
Dim Dr, J&: For Each Dr In Dy
    If UB(Dr) <> UDc Then
        ReDim Preserve Dr(UDc)
        O(J) = Dr
    End If
    J = J + 1
Next
DyReDimDr = O
End Function
Private Function WUDc%(Dy())
Dim O%
Dim Dr: For Each Dr In Dy
    O = Max(O, UB(Dr))
Next
WUDc = O
End Function
Function WsTMthnCmlPC() As Worksheet: Set WsTMthnCmlPC = WsDrs(DrsTMthnCmlPC): End Function
