Attribute VB_Name = "MxIde_Dv_Udt_zIntl1_MsrcDvUdt"
Option Compare Database
Option Explicit


Private Sub B_MsrcyDvUdtP()
GoSub ZZ
Exit Sub
ZZ:
    BrwNlyy MsrcyDvUdtP(CPj)
    Return
End Sub

Private Function MsrcoptDvUdt(C As VBComponent) As Nlyopt
Dim Srcopt As Lyopt: Srcopt = SrcoptDvUdt(SrcCmp(C))
MsrcoptDvUdt = NlyoptLyopt(Srcopt, C.Name)
End Function

Function MsrcyDvUdtP(P As VBProject) As Nly()
Dim C As VBComponent: For Each C In P.VBComponents
    With SrcoptDvUdt(SrcCmp(C))
        If .Som Then PushNly MsrcyDvUdtP, Nly(C.Name, .Ly)
    End With
Next
End Function


