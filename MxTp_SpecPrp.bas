Attribute VB_Name = "MxTp_SpecPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_SpecPrp."

Function SpeciHdrLixy(I() As TSpeci, Specit) As Integer() ' Return the Lix of @I which match the @Specit
Dim J%: For J = 0 To UbTSpeci(I)
    If I(J).Specit = Specit Then PushI SpeciHdrLixy, I(J).Ix
Next
End Function

Function LyySpeciy(I() As TSpeci) As Variant()
Dim J%: For J = 0 To UbTSpeci(I)
    PushI LyySpeciy, LyILny(I(J).IxLny)
Next
End Function

Function SpeciyT(S As TSpec, Specit$, Optional Fmix = 0) As TSpeci()
Dim I() As TSpeci: I = S.Itms
Dim J&: For J = Fmix To UbTSpeci(S.Itms)
    Dim M As TSpeci: M = I(J)
    If M.Specit = Specit Then
        PushTSpeci SpeciyT, M
    End If
Next
End Function

Function Specity(I() As TSpeci) As String() '#spec-item-type-array#
Dim J%: For J = 0 To UbTSpeci(I)
    PushI Specity, I(J).Specit
Next
End Function
