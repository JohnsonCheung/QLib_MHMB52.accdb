Attribute VB_Name = "MxIde_Mthn_Rel"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Rel."

Function RelMthnPubToMdnPC() As Dictionary
Set RelMthnPubToMdnPC = RelMthnPubToMdnP(CPj)
End Function

Private Sub B_RelMthnPubToMdnP()
BrwRel RelMthnPubToMdnP(CPj)
End Sub

Function RelMthnPubToMdnP(P As VBProject) As Dictionary
Dim O As New Dictionary, Mthn, Mdn$
Dim C As VBComponent: For Each C In P.VBComponents
    Mdn = C.Name
    For Each Mthn In Itr(MthnyPub(SrcM(C.CodeModule)))
        PushParChd O, Mthn, Mdn
    Next
Next
Set RelMthnPubToMdnP = O
End Function

Function RelMthnToMdnP(P As VBProject) As Dictionary
Dim O As New Dictionary, Mthn, Mdn$
Dim C As VBComponent: For Each C In P.VBComponents
    Mdn = C.Name
    For Each Mthn In Itr(Mthny(SrcM(C.CodeModule)))
        PushParChd O, Mthn, Mdn
    Next
Next
Set RelMthnToMdnP = O
End Function

Function MthnRelMdnP() As Dictionary
Set MthnRelMdnP = RelMthnToMdnP(CPj)
End Function
