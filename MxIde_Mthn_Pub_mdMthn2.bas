Attribute VB_Name = "MxIde_Mthn_Pub_mdMthn2"
Option Compare Text
Option Explicit
Function Mi2yMthMdPubPC() As String(): Mi2yMthMdPubPC = Mi2yMthMdP(CPj): End Function
Function Mi2yMthMdPC() As String():       Mi2yMthMdPC = Mi2yMthMdP(CPj): End Function
Function Mi2yMthMdP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy Mi2yMthMdP, AmAddSfx(MthnyPub(SrcCmp(C)), " " & C.Name)
Next
End Function
Function Mi2yMdMthPubP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy Mi2yMdMthPubP, AmAddPfx(MthnyPub(SrcCmp(C)), C.Name & " ")
Next
End Function
Function Mi2yMdMthPubPC() As String(): Mi2yMdMthPubPC = Mi2yMdMthPubP(CPj): End Function
