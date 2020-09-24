Attribute VB_Name = "MxIde_Md_MdPrp_EmpMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Emp."

Function IsMdEmp(M As CodeModule) As Boolean
If M.CountOfLines > 10 Then Exit Function
Dim J&: For J = 1 To M.CountOfLines
    If Not IsLnNonSrc(M.Lines(J, 1)) Then Exit Function
Next
IsMdEmp = True
End Function

Sub RmvEmpCMd()
RmvEmpMd CPj
End Sub
Sub RmvEmpMd(P As VBProject)
Dim N: For Each N In Itr(EmpMdny(P))
    RmvCmp P.VBComponents(N)
Next
End Sub

Function EmpCMdny() As String()
EmpCMdny = EmpMdny(CPj)
End Function

Function EmpMdnyP() As String()
EmpMdnyP = EmpMdny(CPj)
End Function

Function EmpMdny(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMdEmp(C.CodeModule) Then
        PushI EmpMdny, C.Name
    End If
Next
End Function

Private Sub B_IsMdEmp()
Dim M As CodeModule
'GoSub T1
'GoSub T2
GoSub T3
Exit Sub
T3:
    Debug.Assert IsMdEmp(Md("Dic"))
    Return
T2:
    Set M = Md("Module2")
    Ept = True
    GoTo Tst
T1:
    '
    Dim T$, P As VBProject
        Set P = CPj
        T = Tmpn
    '
'    Set M = PjAddMd(P, T)
    Ept = True
    GoSub Tst
    Return
Tst:
    Act = IsMdEmp(M)
    C
    Return
End Sub

Function IsSrcEmp(Src$()) As Boolean
Dim L: For Each L In Itr(Src)
    If Not IsLnNonSrc(L) Then Exit Function
Next
IsSrcEmp = True
End Function
Function MdnyEmpV(V As VBE) As String()
Dim P As VBProject
For Each P In V.VBProjects
    PushIAy MdnyEmpV, EmpMdny(P)
Next
End Function

Function MthnyNoMthPC() As String(): MthnyNoMthPC = AeSfx(MthnyNoMthP(CPj), "__CmlDfn"): End Function
Function MthnyNoMthP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMdNoMth(C.CodeModule) Then PushI MthnyNoMthP, C.Name
Next
End Function

Function IsMdNoMth(M As CodeModule) As Boolean
Dim J&: For J = M.CountOfDeclarationLines + 1 To M.CountOfLines
    If IsLnMth(M.Lines(J, 1)) Then Exit Function
Next
IsMdNoMth = True
End Function
