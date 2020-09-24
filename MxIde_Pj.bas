Attribute VB_Name = "MxIde_Pj"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj."

Function CvPj(I) As VBProject:    Set CvPj = I:                      End Function
Function Pj(Pjn) As VBProject:      Set Pj = CVbe.VBProjects(Pjn):   End Function
Function IsPjn(N) As Boolean:        IsPjn = HasEle(PjnyVC, N):      End Function
Function CPjn$():                     CPjn = CPj.Name:               End Function
Function CLnoM&(M As CodeModule):    CLnoM = RrccPne(M.CodePane).R1: End Function
Function CLno&():                     CLno = CLnoM(CMd):             End Function
Function CCmpy() As VBComponent(): PushObj CCmpy, CMd.Parent: End Function
Function CmpyMdn(Mdn$) As VBComponent(): PushObj CmpyMdn, CmpMdn(Mdn): End Function
Function CCmpny() As String(): PushI CCmpny, CMdn: End Function
Function CMdny() As String(): PushI CMdny, CMdn: End Function
Function CCmp() As VBComponent:           Set CCmp = CMd.Parent:           End Function
Function CVbe() As VBE:                   Set CVbe = Application.VBE:      End Function
Function CPj() As VBProject:               Set CPj = CVbe.ActiveVBProject: End Function
Function CPne() As VbIde.CodePane:        Set CPne = CVbe.ActiveCodePane:  End Function
Function PjMainAcsC() As VBProject: Set PjMainAcsC = PjMainAcs(Acs):       End Function
Function PjMainAcs(A As Access.Application) As VBProject ' main project
Dim P As VBProject: For Each P In A.VBE.VBProjects
    If A.CurrentDb.Name = P.FileName Then Set PjMainAcs = P: Exit Function
Next
ThwImposs "PjMainAcs", FmtQQ("Each @Acs with CurrentDb should have a PjMainAcs.  @Acs[?]", A.CurrentDb.Name)
End Function
Sub BrwPthPC():              BrwPthP CPj:    End Sub
Sub BrwPthP(P As VBProject): BrwPth PthP(P): End Sub
