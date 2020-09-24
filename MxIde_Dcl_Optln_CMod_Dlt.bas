Attribute VB_Name = "MxIde_Dcl_Optln_CMod_Dlt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Optln_CMod_Dlt."
Sub CMod_DltPC():                                DltTMdLnoy XTMdLnoy(CPj, "CMod"):          End Sub
Sub CMod_Brw():                                  Brw XLy("CMod"):                           End Sub
Private Function XLy(CnstnL$) As String(): XLy = FmtT1ry(LyTMdLnoy(XTMdLnoy(CPj, CnstnL))): End Function

Private Function XTMdLnoy(P As VBProject, CnstnL$) As TMdLno()
Dim C As VBComponent: For Each C In P.VBComponents
    If C.Name <> "MxVb_Cnst" Then
        Dim Lno&: Lno = xlNo(C, CnstnL)
        If Lno > 0 Then
            PushTMdLno XTMdLnoy, TMdLno(C.CodeModule, Lno)
        End If
    End If
Next
End Function
Private Function xlNo&(C As VBComponent, CnstnL$)
Dim M As CodeModule: Set M = MdCmp(C): If IsNothing(M) Then Exit Function
Dim J%: For J = 1 To M.CountOfDeclarationLines
    If HasCnstnL(M.Lines(J, 1), CnstnL) Then
        xlNo = J
        Exit Function
    End If
Next
End Function
