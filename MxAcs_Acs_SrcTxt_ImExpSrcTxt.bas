Attribute VB_Name = "MxAcs_Acs_SrcTxt_ImExpSrcTxt"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_SrcTxt_IExpSrcTxt."
Sub ExpAcsRptC(): ExpAcsRpt Acs, PthSrcPC: End Sub
Sub ExpAcsRpt(A As Access.Application, PthTo$)
Dim F As AccessObject: For Each F In A.CurrentProject.AllReports
    Debug.Print Now, "Exp Rpt", F.Name
    A.SaveAsText acReport, F.Name, PthTo & F.Name & ".rpt.txt"
Next
End Sub
Sub ExpAcsFrmC(): ExpAcsFrm Acs, PthSrcPC: End Sub

Sub ExpAcsFrm(A As Access.Application, PthTo$)
Dim F As AccessObject: For Each F In A.CurrentProject.AllForms
    Debug.Print Now, "Exp Frm", F.Name
    A.SaveAsText acForm, F.Name, PthTo & F.Name & ".frm.txt"
Next
End Sub

Sub ImpAcsFrmC(PthFm$): ImpAcsFrm Acs, PthFm: End Sub
Sub ImpAcsFrm(A As Access.Application, PthFm$)
Dim P$: P = PthEnsSfx(PthFm)
Dim F: For Each F In Itr(Fnay(P, "*.rpt.txt"))
    Debug.Print Now, "Imp Frm", F
    A.LoadFromText acForm, Fnn(Fnn(F)), P & F
Next
End Sub
Sub ImpAcsRptC(PthFm$): ImpAcsRpt Acs, PthFm: End Sub
Sub ImpAcsRpt(A As Access.Application, PthFm$)
Dim P$: P = PthEnsSfx(PthFm)
Dim F: For Each F In Itr(Fnay(P, "*.rpt.txt"))
    Debug.Print Now, "Imp Rpt", F
    A.LoadFromText acReport, Fnn(Fnn(F)), P & F
Next
End Sub
