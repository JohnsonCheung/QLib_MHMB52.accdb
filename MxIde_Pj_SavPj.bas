Attribute VB_Name = "MxIde_Pj_SavPj"
Option Compare Text
Option Explicit
Const TimMdy As Date = #9/18/2020 9:38:09 PM#
Const CMod$ = "MxIde_Pj_SavPj."

Private Sub B_SavPj()
Dim P As VBProject
Dim Wb As Workbook: Set Wb = WbNw
Dim Fx$: Fx = FxTmp
Wb.SaveAs Fx
Set P = PjWb(Wb)
Dim Cmp As VBComponent: Set Cmp = P.VBComponents.Add(vbext_ct_ClassModule)
Cmp.CodeModule.AddFromString "Sub AA()" & vbCrLf & "End Sub"
SavPj P
End Sub
Sub SavPjC(): SavPj CPj: End Sub
Sub SavPj(P As VBProject)
Const CSub$ = CMod & "SavPj"
Const Shw As Boolean = False
If P.Saved Then
    If Shw Then Debug.Print FmtQQ("SavPj: Pj(?) is already saved", P.Name)
    Exit Sub
End If
Dim C As VBComponent: For Each C In P.VBComponents
    WSavCmp C
Next
'<===== Save
DoEvents
If Not P.Saved Then
    If Shw Then Debug.Print FmtQQ("SavPj: Pj(?) cannot be saved for unknown reason <=================================", P.Name)
Else
    If Shw Then Debug.Print FmtQQ("SavPj: Pj(?) is saved <---------------", P.Name)
End If
End Sub
Function MdnyNotSav() As String(): MdnyNotSav = ItnWhPrpFalse(CCmpy, "Saved"): End Function
Private Sub WSavCmp(C As VBComponent)
DoEvents
Dim V As VBE: Set V = C.Collection.VBE
EnsTimMdyM C.CodeModule
Dim P As VBProject: Set P = PjCmp(C)
    ChkPjHasFfn P
ActPj P
Dim BtnSav As CommandBarButton: Set BtnSav = IBtnSavzV(V) 'Set Sav-Button to B
    If Not HasPfx(BtnSav.Caption, "&Save " & FnnPj(P)) Then
        Stop
        Thw CSub, "Caption is not expected", "Save-Bottun-Caption Expected", BtnSav.Caption, "&Save " & FnnPj(P)
    End If
BtnSav.Execute '<===
End Sub
Function FnnPj$(P As VBProject): FnnPj = Fnn(Pjf(P)): End Function
Sub ChkPjHasFfn(P As VBProject)
Stop '
End Sub
