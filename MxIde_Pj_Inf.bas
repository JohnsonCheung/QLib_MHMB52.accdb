Attribute VB_Name = "MxIde_Pj_Inf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Inf."
Public Const ExtFrm$ = ".frm.txt"

Function PjFxaX(X As Excel.Application, Fxa$) As VBProject
'Ret: Ret :Pj of @Fxa fm @X if exist, else @X.Opn @Fxa
Dim O As VBProject: Set O = PjFxaV(X.VBE, Fxa)
If Not IsNothing(O) Then OpnFxX X, Fxa
Set PjFxaX = PjFxaV(X.VBE, Fxa)
End Function
Function PjFxa(Fxa$) As VBProject: Set PjFxa = PjFxaV(CVbe, Fxa): End Function
Function PjFxaV(V As VBE, Fxa$) As VBProject
Dim P As VBProject: For Each P In V.VBProjects
    If Pjf(P) = Fxa Then Set PjFxaV = P: Exit Function
Next
End Function
Function HasFxa(Fxa$) As Boolean: HasFxa = HasEleStr(PjnyVC, Fn(Fxa)): End Function

Private Sub B_CrtFxa()
GoSub T1
Exit Sub
Dim Fxa$
T1:
    Fxa = FxmTmp


End Sub
Sub CrtFxa(Fxa$)
Const CSub$ = CMod & "CrtFxa"
'Do: crt an emp Fxa with pjn derived from @Fxa
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then Thw CSub, "In Xls, there is Pjn = Fxa", "Fxa AllPj-In-Xls", Fxa, PjnyVC
Dim B As Workbook: Set B = Xls.Workbooks.Add
B.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn 'Must save first, otherwise PjFxa will fail.
PjFxaX(Xls, Fxa).Name = Fnn(Fxa)
B.Close True
End Sub

Function FfnyFrmPth(Pth) As String()
Dim I: For Each I In Itr(Ffny(Pth, "*" & ExtFrm))
    PushI FfnyFrmPth, I
Next
End Function

Function ClsyP(P As VBProject) As CodeModule()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsCls(C) Then
        PushObj ClsyP, C
    End If
Next
End Function


Private Sub B_CmpyP()
Dim Act() As VBComponent
Dim C, T As vbext_ComponentType
For Each C In CmpyP(CPj)
    T = CvCmp(C).Type
    If T <> vbext_ct_StdModule And T <> vbext_ct_ClassModule Then Stop
Next
End Sub

Function IsPjNoClsNoMod(P As VBProject) As Boolean
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
IsPjNoClsNoMod = True
End Function

Function ItrModP(P As VBProject): Asg Itr(ModyP(P)), ItrModP: End Function
Function ModyP(P As VBProject) As CodeModule()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent: For Each C In P.VBComponents
    If C.Type = vbext_ct_StdModule Then
        PushObj ModyP, C.CodeModule
    End If
Next
End Function

Function MdayMdny(Mdny$()) As CodeModule()
Dim P As VBProject: Set P = CPj
Dim N: For Each N In Itr(Mdny)
    PushI MdayMdny, MdP(P, N)
Next
End Function

Function CmpyPC() As VBComponent(): CmpyPC = CmpyP(CPj): End Function
Function CmpyC() As VBComponent(): PushObj CmpyC, CCmp: End Function
Function CmpyP(P As VBProject) As VBComponent()
Dim C As VBComponent: For Each C In P.VBComponents
    PushObj CmpyP, C
Next
End Function

Private Sub B_MdayP()
Dim Mday() As CodeModule
Mday = MdayP(CPj)
Dim M, O$(): For Each M In Itr(Mday)
    PushI O, Mdn(CvMd(M))
Next
BrwAy O
End Sub

Function MdayPC() As CodeModule(): MdayPC = MdayP(CPj): End Function
Function MdayP(P As VBProject) As CodeModule()
MdayP = MdayCmpy(CmpyP(P))
End Function
Function MdayCmpy(C() As VBComponent) As CodeModule()
Dim I: For Each I In Itr(C)
    PushObj MdayCmpy, CvCmp(I).CodeModule
Next
End Function
