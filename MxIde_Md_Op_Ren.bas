Attribute VB_Name = "MxIde_Md_Op_Ren"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Op_Ren."
Sub RenMdLis(PatnssAndMd$)
Dim M$(): M = SySrtQ(AmAli(AmQuoDbl(MdnyPC(PatnssAndMd)))) ' Quoted and Aligned
Dim I: For Each I In Itr(M)
    Debug.Print FmtQQ("RenMdFmTo ?, ?:JmpMdn?", I, I, RTrim(I))
Next
Debug.Print "Count: "; Si(M)
End Sub
Sub RenMdFmTo(MdnFm$, MdnTo$)
If MdnFm = MdnTo Then Exit Sub
Cmp(MdnFm).Name = MdnTo
End Sub
Sub RenMdC(Nwn): CCmp.Name = Nwn: End Sub
Sub RenMdPfx(PfxFm$, PfxTo$)
Dim C As VBComponent: For Each C In CPj.VBComponents
    If HasPfx(C.Name, PfxFm) Then
        C.Name = RplPfx(C.Name, PfxFm, PfxTo)
    End If
Next
End Sub

Sub RenMdRmvPfx(Pfx$)
Dim C As VBComponent: For Each C In CPj.VBComponents
    If HasPfx(C.Name, Pfx) Then
        C.Name = RmvPfx(C.Name, Pfx)
    End If
Next
End Sub

Sub RenMdStrSfxDash()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        C.Name = C.Name & "_"
    End If
Next
End Sub
Sub RenMdRmvSfxDash()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        C.Name = RmvSfx(C.Name, "_")
    End If
Next
End Sub
