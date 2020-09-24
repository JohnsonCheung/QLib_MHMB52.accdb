Attribute VB_Name = "MxAcs_AcCtl_AcCtlOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_AcCtl_SetAcCtl."
Property Let FrmAllCtlPrp(F As Access.Form, P, V)
Dim C As Access.Control: For Each C In F.Controls
    C.Properties(P) = V
Next
End Property
Property Let CtlPrp(C As Access.Control, P, V)
On Error GoTo X
C.Properties(P) = V
Exit Property
X: Dim Er$: Er = Err.Description: InfVbEr CSub, Er, "Ctln P @V-Tyn @V", C.Name, P, TypeName(V), V
End Property
Sub InfVbEr(Fun$, VbMsg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Inf Fun, "VbEr found: [" & VbMsg & "]", Nav
End Sub
Property Let FrmCtlTabStop(F As Access.Form, Ctlnn$, OnOff As Boolean)
If Ctlnn$ = "*All" Then FrmAllCtlPrp(F, "TabStop") = OnOff
Dim C As Access.Control: For Each C In F.Controls
    FrmCtlPrp(F, C, "TabStop") = OnOff
Next
End Property
Property Let FrmCtlPrp(F As Access.Form, C As Access.Control, P, V)
Dim I As AccessObjectProperty: For Each P In C.Properties
    If I.Name = P Then
        I.Value = V
        Exit Property
    End If
Next
End Property
Sub SetTBoxIf(T As Access.TextBox, Txt$)
If T <> "" Then T.Text = Txt
End Sub
