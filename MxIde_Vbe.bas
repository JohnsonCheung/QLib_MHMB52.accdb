Attribute VB_Name = "MxIde_Vbe"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Vbe."
Function CvVbe(A) As VBE
Set CvVbe = A
End Function

Function PjV(V As VBE, Pjn$) As VBProject: Set PjV = V.VBProjects(Pjn): End Function
Function IsPjSav(P As VBProject):          IsPjSav = P.Saved:           End Function

Function IsPjSavC() As Boolean: IsPjSavC = CPj.Saved: End Function
Sub ChkPjSav(P As VBProject)
If Not P.Saved Then Raise "Pj is not saved"
End Sub

Function PjfyVC() As String(): PjfyVC = PjfyV(CVbe): End Function
Function PjfyV(V As VBE) As String()
Dim P As VBProject: For Each P In V.VBProjects
    PushNB PjfyV, Pjf(P)
Next
End Function

Function PjnyVC() As String(): PjnyVC = PjnyV(CVbe): End Function
Function PjnyV(V As VBE) As String()
Dim P As VBProject: For Each P In V.VBProjects:
    PushI PjnyV, Pjf(P)
Next
End Function

Function HasIdeIBarV(V As VBE, IBarn) As Boolean
HasIdeIBarV = HasItn(V.CommandBars, IBarn)
End Function

Function HasPj(V As VBE, Pjn$) As Boolean
HasPj = HasItn(V.VBProjects, Pjn)
End Function

Function HasPjfV(V As VBE, Pjf) As Boolean
Dim P As VBProject: For Each P In V.VBProjects
    If Pjf(P) = Pjf Then HasPjfV = True: Exit Function
Next
End Function

Private Sub B_MthnyV(): Brw MthnyV(CVbe): End Sub
