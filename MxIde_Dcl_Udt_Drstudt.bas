Attribute VB_Name = "MxIde_Dcl_Udt_Drstudt"
Option Compare Text
Const CMod$ = "MxIde_Dcl_Udt_Drstudt."
Option Explicit
Public Const FFTUdt$ = "Pjn Mdn IsPrv Udtn Mbn Tyn IsAy"
Sub BrwTUdt():                       BrwDrs DrstUdt: End Sub
Function DrstUdt() As Drs: DrstUdt = DrstUdtP(CPj):  End Function
Function DrstUdtP(P As VBProject) As Drs
Dim Dy()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy Dy, WDyCmp(C)
Next
DrstUdtP = DrsFf(FFTUdt, Dy)
End Function
Private Function WDyCmp(C As VBComponent) As Variant()
Dim U() As TUdt: U = TUdtyM(C.CodeModule)
Dim Pjn$: Pjn = PjCmp(C).Name
Dim J%: For J = 0 To UbTUdt(U)
    PushIAy WDyCmp, WDyTUdt(Pjn, C.Name, U(J))
Next
End Function
Private Function WDyTUdt(Pjn$, Mdn$, U As TUdt) As Variant()
Dim J%: For J = 0 To UbTUmb(U.Mbr)
    With U.Mbr(J)
    PushI WDyTUdt, Array(Pjn, Mdn, U.IsPrv, U.Udtn, .Mbn, .Tyn, .IsAy)
    End With
Next
End Function
