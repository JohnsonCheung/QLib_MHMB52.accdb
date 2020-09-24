Attribute VB_Name = "MxAcs_Rescu_RescuCpy"
Option Compare Text
Const CMod$ = "MxAcs_Acs_RescuFbByCpy."
Option Explicit
Sub RescuByCpyPthNoMdC(): RescuByCpyPthNoMd PthPC: End Sub
Sub RescuByCpyPthNoMd(Pth$)
Dim FbFm: For Each FbFm In ItrFbCorrudC
    Dim FbTo$: FbTo = FbRescud(FbFm)
    If AskDltFfn(FbTo) Then
        CrtFb FbTo
        CpyAcsObjNoMd FbFm, FbTo
    End If
Next
End Sub

Sub RescuByCpy(): RescuByCpyPth PthPC: End Sub
Sub RescuByCpyPth(Pth$)
Dim FbFm: For Each FbFm In ItrFbCorrudC
    Dim FbTo$: FbTo = FbRescud(FbFm)
    If AskDltFfn(FbTo) Then
        CrtFb FbTo
        CpyAcsObj FbFm, FbTo
    End If
Next
End Sub
Sub DltFbCorrud(): DltFfnIf WFbCorrud: End Sub

Sub CrtFbCorrud()
Dim Fb$: Fb = WFbCorrud
If HasFfn(Fb) Then MsgBox "FbCorrud already exist", vbInformation: Exit Sub
CpyFfn CPjf, Fb
End Sub
Private Function WFbCorrud$(): WFbCorrud = FfnAddFnsfx(CPjf, "(Corrupted)"): End Function

Sub DmpCmpErByCmps()
Dim J%
Dim C As VBComponent: For Each C In CPj.VBComponents
    J = J + 1
    If C.Name = "MxIde_Mth_CSub_Intl_CSubbEns" Then Stop
Next
End Sub
Sub DmpCmpErByDaodoc()
Dim J%
Dim D As Dao.Document: For Each D In CDb.Containers("Modules").Documents
    J = J + 1
    If D.Name = "MxIde_Mth_CSub_Intl_CSubbEns" Then Stop
Next
End Sub
Sub DmpCmpBlnkNm()
Dim J%, N%
Dim C As VBComponent: For Each C In CPj.VBComponents
    J = J + 1
    If C.Name = "" Then
        Debug.Print "InoCmp has blank name"; J;
        Debug.Print "<=== Removed"
        N = N + 1
    End If
Next
Debug.Print N; "Components with blank name"
End Sub
Sub DmpCmpInCntrInPj()
Dim NyPj$(): NyPj = MdnyPC
Dim NyDoc$(): NyDoc = Itn(CntrMdC.Documents)
Dim N, M%: For Each N In Itr(SyMinus(NyDoc, NyPj))
    Debug.Print N
    M = M + 1
Next
Debug.Print "There is"; M; "Cmp excess in CntrMd"
M = 0: For Each N In Itr(SyMinus(NyPj, NyDoc))
    Debug.Print N
    M = M + 1
Next
Debug.Print "There is"; M; "Cmp excess in Pj"
End Sub
Sub RmvCmpBlnkNm()
If Not CfmYes("Start remove blank name component?") Then Exit Sub
Dim J%, N%
Dim C As VBComponent: For Each C In CPj.VBComponents
    J = J + 1
    If C.Name = "" Then
        Debug.Print "InoCmp has blank name"; J;
        CPj.VBComponents.Remove C
        Debug.Print "<=== Removed"
        N = N + 1
    End If
Next
Debug.Print N; "Component with blank name is removed"
End Sub
