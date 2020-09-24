Attribute VB_Name = "MxDta_FldVy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_FldVy."
Type TFvy: F As String: Vy() As Variant: End Type 'Deriving(Ctor Ay)
Function TFvy(F, Vy()) As TFvy
With TFvy
    .F = F
    .Vy = Vy
End With
End Function
Function FvyUB&(A() As TFvy)

End Function
Function FvySI&(A() As TFvy)

End Function
Function FnyFvyy(A() As TFvy) As String()
Dim J%: For J = 0 To FvyUB(A)
    PushI FnyFvyy, A(J).F
Next
End Function

Function FfTFvyy$(A() As TFvy): FfTFvyy = TmlAy(FnyFvyy(A)): End Function
Sub PushFvy(O() As TFvy, M As TFvy): Dim N&: N = FvySI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function VyTFvyy(A() As TFvy, F$) As Variant()
Const CSub$ = CMod & "VyFvyy"
Dim J%: For J = 0 To FvyUB(A)
    With A(J)
        If .F = F Then
            VyTFvyy = .Vy
            Exit Function
        End If
    End With
Next
Thw CSub, "@Fld not found in given @FldVyAy", "@Fld Ff-in-@FldVyAy", F, FnyFvyy(A)
End Function
