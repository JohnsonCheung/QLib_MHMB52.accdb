Attribute VB_Name = "MxIde_Src_Msrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Msrc."
Private Sub B_MdnSrcyPC()
MsgBox RepLstss(LstssNlyy(MsrcyPC))
End Sub
Sub BrwMsrcy(M() As Nly):              BrwNlyy M:         End Sub
Sub VcMsrcy(M() As Nly):               VcNlyy M:          End Sub
Function MsrcyPC() As Nly(): MsrcyPC = MsrcyMdny(MdnyPC): End Function
Function MsrcyMdny(Mdny$()) As Nly()
Dim N: For Each N In Itr(Mdny)
    PushNly MsrcyMdny, Nly(N, SrcMdn(N))
Next
End Function

Sub RplMdMsrcy(Msrcy() As Nly)
Dim J%: For J = 0 To UbNly(Msrcy)
    With Msrcy(J)
        If .Nm <> "MxIde_Src_Dta_TMdSrc_Op" Then
            Dim M As CodeModule: Set M = Md(.Nm)
            RplMd M, JnCrLf(.Ly)
            SavPj PjM(M)
        End If
    End With
Next
End Sub
