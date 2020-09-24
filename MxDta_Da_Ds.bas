Attribute VB_Name = "MxDta_Da_Ds"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Ds."
Function DsAddDt(Ds As Ds, Dt As Dt) As Ds ' add ds dt becomes ds
Const CSub$ = CMod & "DsAddDt"
If HasDt(Ds, Dt.Dtn) Then Thw CSub, "@Ds already has @Dt", "Ds Dt", Ds.Dsn, Dt.Dtn
DsAddDt = Ds
PushDt DsAddDt.Dty, Dt
End Function

Function DtDs(D As Ds, Dtn$) As Dt
Const CSub$ = CMod & "DtDs"
Dim Ay() As Dt: Ay = D.Dty
Dim J&: For J = 0 To DtUB(Ay)
    If Ay(J).Dtn = Dtn Then
        DtDs = Ay(J)
        Exit Function
    End If
Next
Thw CSub, "No such Dtn in Ds", "Such-Dtn DtNy-In-Ds", Dtn, TnyDs(D)
End Function
Function HasDt(D As Ds, Dtn$) As Boolean
Dim Ay() As Dt: Ay = D.Dty
Dim J&: For J = 0 To DtUB(D.Dty)
    If Ay(J).Dtn = Dtn Then HasDt = True: Exit Function
Next
End Function
Function DtSI&(A() As Dt): On Error Resume Next: DtSI = UBound(A) + 1: End Function
Function DtUB&(A() As Dt): DtUB = DtSI(A) - 1: End Function
Sub PushDt(O() As Dt, M As Dt): Dim N&: N = DtSI(O): ReDim Preserve O(N): O(N) = M: End Sub

Function TnyDs(D As Ds) As String()
Dim Ay() As Dt: Ay = D.Dty
Dim J&: For J = 0 To DtUB(Ay)
    PushI TnyDs, Ay(J).Dtn
Next
End Function
