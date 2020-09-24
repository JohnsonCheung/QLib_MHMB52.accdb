Attribute VB_Name = "MxIde_A_MdTmp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_A_MdTmp."
Public Const PfxTmpMd$ = "ZTmp_"
Function PjCmp(A As VBComponent) As VBProject
Set PjCmp = A.Collection.Parent
End Function

Function HasCmp(Cmpn) As Boolean:                   HasCmp = HasCmpP(CPj, Cmpn):           End Function
Function HasCmpP(P As VBProject, Cmpn) As Boolean: HasCmpP = HasItn(P.VBComponents, Cmpn): End Function

Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function

Function MdayCmpy(Cmpy() As VBComponent) As CodeModule()
Dim I: For Each I In Itr(Cmpy)
    PushObj MdayCmpy, CvCmp(I).CodeModule
Next
End Function

Function MdTmp() As CodeModule
Dim T$: T = Tmpn(PfxTmpMd)
AddMod CPj, T
Set MdTmp = Md(T)
End Function

Function MdnyTmpP(P As VBProject) As String(): MdnyTmpP = AwPfx(MdnyP(P), PfxTmpMd): End Function
Sub DltTmpMdP(P As VBProject)
Dim M: For Each M In Itr(MdnyTmpP(P))
    P.VBComponents.Remove P.VBComponents(M)
Next
End Sub
Sub DltTmpMdPC(): DltTmpMdP CPj: End Sub
