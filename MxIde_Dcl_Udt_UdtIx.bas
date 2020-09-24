Attribute VB_Name = "MxIde_Dcl_Udt_UdtIx"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_UdtIx."

Function BeiUdt(Dcl$(), Udtn$) As Bei
Dim B%: B = BiUdtn(Dcl, Udtn): If B = -1 Then BeiUdt = BeiEmp: Exit Function
BeiUdt = Bei(B, EiUdt(Dcl, B))
End Function
Private Function BiUdtn&(Dcl$(), Udtn$) ' Ret -1 if not found
Dim J%: For J = 0 To UB(Dcl)
   If UdtnLn(Dcl(J)) = Udtn Then BiUdtn = J: Exit Function
Next
BiUdtn = -1
End Function
Private Function EiUdt%(Dcl$(), Bi) 'Return -1 if Bi is < 0 or the Eix of [End Type] line
Const CSub$ = CMod & "EiUdt"
If Bi < 0 Then ThwPm CSub, "Bi must >=0", "Bi", Bi
Dim J%: For J = Bi To UB(Dcl)
    If HasSsub(Dcl(J), "End Type", eCasSen) Then
        EiUdt = J
        Exit Function
    End If
Next
Thw CSub, "No End Type is found", "Bi Dcl", Bi, AmAddIxPfx(Dcl)
End Function

Function BeiyUdt(Dcl$()) As Bei()
Dim Biy%()
    Dim J%: For J = 0 To UB(Dcl)
        If IsLnUdt(Dcl(J)) Then PushI Biy, J
    Next

Dim Bi: For Each Bi In Itr(Biy)
    PushBei BeiyUdt, Bei(Bi, EiUdt(Dcl, Bi))
Next
End Function

Private Sub B_IsLnUdt()
Dim O$()
Dim L: For Each L In DclPC
    If IsLnUdt(L) Then PushI O, L
Next
BrwAy O
End Sub

Private Function IsLnUdt(Ln) As Boolean: IsLnUdt = IsShfTm(RmvMdy(Ln), "Type"): End Function
