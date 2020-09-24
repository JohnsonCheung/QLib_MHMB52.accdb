Attribute VB_Name = "MxVb_Dta_Nly"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Nly."
Type Nly: Nm As String: Ly() As String: End Type 'Deriving(Ay Ctor)
Type Nlyopt: Som As Boolean: Nly As Nly: End Type
Function Nly(Nm, Ly) As Nly
With Nly
    .Nm = Nm
    .Ly = Ly
End With
End Function
Function SomNly(A As Nly) As Nlyopt
With SomNly
    .Som = True
    .Nly = A
End With
End Function
Function NlyoptLyopt(A As Lyopt, Nm$) As Nlyopt
If Not A.Som Then Exit Function
NlyoptLyopt = SomNly(Nly(Nm, A.Ly))
End Function
Sub PushNly(O() As Nly, M As Nly): Dim N&: N = SiNly(O): ReDim Preserve O(N): O(N) = M: End Sub
Sub PushNlyopt(O() As Nly, M As Nlyopt)
If M.Som Then PushNly O, M.Nly
End Sub
Function UbNly(A() As Nly): UbNly = SiNly(A) - 1: End Function
Function SiNly(A() As Nly): On Error Resume Next: SiNly = UBound(A) + 1: End Function

Function StsNly$(A As Nly):     StsNly = A.Nm & " " & StsLy(A.Ly):                                           End Function
Function NLnNly&(A As Nly):     NLnNly = Si(A.Ly):                                                           End Function
Function LenNly&(A As Nly):     LenNly = LenSy(A.Ly):                                                        End Function
Function StsNlyy$(A() As Nly): StsNlyy = FmtQQ("N-Nlyy(?) NLn(?) Len(?)", SiNly(A), NLnNlyy(A), LenNlyy(A)): End Function
Sub PushNlyy(O() As Nly, M() As Nly)
Dim J&: For J = 0 To UbNly(M)
    PushNly O, M(J)
Next
End Sub
