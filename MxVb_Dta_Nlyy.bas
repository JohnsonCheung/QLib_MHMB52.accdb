Attribute VB_Name = "MxVb_Dta_Nlyy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Nlyy."

Function LstssNlyy(A() As Nly) As Lstss: LstssNlyy = LstssLyy(LyyNlyy(A)): End Function
Function LyyNlyy(A() As Nly) As Variant()
Dim J&: For J = 0 To UbNly(A)
    PushI LyyNlyy, A(J).Ly
Next
End Function
Function NyNlyy(A() As Nly) As String()
Dim J&: For J = 0 To UbNly(A)
    PushI NyNlyy, A(J).Nm
Next
End Function
Function LenNlyy&(Nlyy() As Nly)
Dim J&, O&: For J = 0 To UbNly(Nlyy)
    O = O + LenSy(Nlyy(J).Ly)
Next
LenNlyy = O
End Function


Function LsyNlyy(A() As Nly) As String()
Dim J&: For J = 0 To UbNly(A)
    Dim Nm$: Nm = StsNly(A(J))
    PushI LsyNlyy, Nm & vbCrLf & JnCrLf(A(J).Ly)
Next
End Function
Private Sub B_BrwNlyy(): BrwLsy LsyNlyy(MsrcyPC): End Sub
Function NLnNlyy&(A() As Nly)
Dim O&
Dim J&: For J = 0 To UbNly(A)
    O = O + NLnNly(A(J))
Next
End Function
Sub BrwNlyy(N() As Nly):                            BrwAy FmtNlyy(N):   End Sub
Sub VcNlyy(N() As Nly):                             VcAy FmtNlyy(N):    End Sub
Function FmtNlyy(A() As Nly) As String(): FmtNlyy = FmtLsy(LsyNlyy(A)): End Function
