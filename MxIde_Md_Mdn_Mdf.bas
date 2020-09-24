Attribute VB_Name = "MxIde_Md_Mdn_Mdf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Mdn_Mdf."
Function MdnyMdfssub(Mdfssub$, Optional C As eCas) As String()
Dim Cmp As VBComponent: For Each Cmp In CPj.VBComponents
    If HasMdfssub(Cmp.Name, Mdfssub, C) Then PushI MdnyMdfssub, Cmp.Name
Next
End Function
Function HasMdfssubSen(Mdn$, Mdfssub) As Boolean:                  HasMdfssubSen = HasSsub(MdfMdn(Mdn), Mdfssub, eCasSen): End Function
Function HasMdfssub(Mdn$, Mdfssub, Optional C As eCas) As Boolean:    HasMdfssub = HasSsub(MdfMdn(Mdn), Mdfssub, C):       End Function
Function MdfMdn$(Mdn):                                                    MdfMdn = AftOrAllRev(Mdn, "_"):                  End Function
Function Mdfy() As String()
Dim O$()
Dim Mdn: For Each Mdn In MdnyPC
    Push O, MdfMdn(Mdn) & " " & Mdn
    'Debug.Print Mdn, MdfMdn(Mdn)
Next
Mdfy = SySrtQ(O)
End Function
Function MdMdf(Mdf$) As CodeModule: Set MdMdf = MdMdn(MdnMdf(Mdf)): End Function
Function MdnMdf$(Mdf$)
Const CSub$ = CMod & "MdnMdf"
Dim A$(): A = MdnyMdf(Mdf)
Select Case Si(A)
Case 0: Debug.Print CSub; "No Mdf[" & Mdf & "]"
Case 1: MdnMdf = A(0)
Case Else: MdnMdf = A(0)
    Debug.Print CSub; "more Mdn with Mdf [" & Mdf & "]:"
    DmpAy LyTab4Spc(A)
End Select
End Function
Function MdnyMdf(Mdf) As String()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If HasSfx(C.Name, "_" & Mdf) Then PushI MdnyMdf, C.Name
Next
End Function
