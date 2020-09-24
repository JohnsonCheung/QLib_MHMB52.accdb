Attribute VB_Name = "MxIde_Md_Op_AddMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Op_AddMd."

Sub AddCls(Clsnn$) 'To CPj
WAddCmpnnP CPj, Clsnn, vbext_ct_ClassModule
JmpMdn Tm1(Clsnn)
End Sub

Private Sub WAddCmp(P As VBProject, Nm, Ty As vbext_ComponentType)
Const CSub$ = CMod & "WAddCmp"
If HasCmpP(P, Nm) Then Inf CSub, "Cmpn exist in Pj", "Cmpn Pjn", Nm, P.Name: Exit Sub
P.VBComponents.Add(Ty).Name = CStr(Nm) ' no CStr will break
End Sub

Private Sub WAddCmpnnP(P As VBProject, Cmpnn$, T As vbext_ComponentType)
Dim N: For Each N In ItrSS(Cmpnn)
    WAddCmp P, N, T
Next
End Sub


Sub AddMd(Mdn$): AddMod CPj, Mdn: End Sub
Sub AddMdnn(Modnn$)
WAddCmpnnP CPj, Modnn, vbext_ct_StdModule
JmpMdn Tm1(Modnn)
End Sub
Sub AddMod(P As VBProject, Modn): WAddCmp P, Modn, vbext_ct_StdModule: End Sub

Sub EnsCls(P As VBProject, Clsn): EnsCmp P, Clsn, vbext_ct_ClassModule: End Sub
Sub EnsMod(P As VBProject, Modn): EnsCmp P, Modn, vbext_ct_StdModule:   End Sub
Sub EnsMd(P As VBProject, Mdn):   EnsCmp P, Mdn, vbext_ct_StdModule:    End Sub
Sub EnsCmp(P As VBProject, Nm, Ty As vbext_ComponentType)
If Not HasCmpP(P, Nm) Then WAddCmp P, Nm, Ty
End Sub

Sub EnsMdCdl(M As CodeModule, Cdl$)
Const CSub$ = CMod & "EnsMdCdl"
If Cdl = LinesEndTrim(SrclM(M)) Then Inf CSub, "Same module lines, no need to replace", "Mdn", Mdn(M): Exit Sub
RplMd M, Cdl
End Sub

Sub RenCmp(A As VBComponent, NewNm$)
Const CSub$ = CMod & "RenCmp"
If HasCmp(NewNm) Then
    Inf CSub, "New cmp exists", "OldCmp NewCmp", A.Name, NewNm
Else
    A.Name = NewNm
End If
End Sub

Sub RmvMdC(): RmvMd CMd: End Sub
Sub RmvCmp(C As VBComponent)
If IsNothing(C) Then Exit Sub
Dim N$: N = MdnNxt(C.Name)
C.Collection.Remove C
JmpMdn N
End Sub
Function MdnNxt$(Mdn)
Dim N$(): N = SySrtQ(MdnyPC)
Dim J%: For J = 0 To UB(N)
    If N(J) = Mdn Then
        If J = UB(N) Then
            MdnNxt = N(0)
        Else
            MdnNxt = N(J + 1)
        End If
    End If
Next
End Function
Sub RmvMdn(Mddn): RmvMd Md(Mddn): End Sub
Sub RmvMd(M As CodeModule)
If IsNothing(M) Then Exit Sub
RmvCmp M.Parent
End Sub

Sub RenModPfx(PfxFm$, PfxTo$): RenModPfxP CPj, PfxFm, PfxTo: End Sub
Sub RenModPfxP(Pj As VBProject, PfxFm$, PfxTo$)
Dim C As VBComponent, N$
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        If HasPfx(C.Name, PfxFm) Then
            RenCmp C, RplPfx(C.Name, PfxFm, PfxTo)
        End If
    End If
Next
End Sub

Function SetCmpn(C As VBComponent, Nm, Optional Fun$ = "SetCmpn") As VBComponent
Dim Pj As VBProject
Set Pj = PjCmp(C)
If HasCmpP(Pj, Nm) Then
    Thw Fun, "Cmp already Has", "Cmp Has-in-Pj", Nm, Pj.Name
End If
If Pj.Name = Nm Then
    Thw Fun, "Cmpn same as Pjn", "Cmpn", Nm
End If
C.Name = Nm
Set SetCmpn = C
End Function
Sub ChgToCls(FmModn$)
Const CSub$ = CMod & "ChgToCls"
If Not HasCmp(FmModn) Then Inf CSub, "Mod not exist", "Mod", FmModn: Exit Sub
If Not IsMod(Md(FmModn)) Then Inf CSub, "It not Mod", "Mod", FmModn: Exit Sub
Dim T$: T = Left(FmModn & "_" & Format(Now, "HHMMDD"), 31)
Md(FmModn).Name = T
Dim Cmp As VBComponent
Dim Cdl$
RmvCmp Cmp
AddCmpCdl CPj, FmModn, vbext_ct_ClassModule, Cdl
End Sub
Sub AddModCdl(P As VBProject, Modn, Cdl$): AddCmpCdl P, Modn, vbext_ct_StdModule, Cdl:   End Sub
Sub AddClsCdl(P As VBProject, Clsn, Cdl$): AddCmpCdl P, Clsn, vbext_ct_ClassModule, Cdl: End Sub
Sub AddCmpCdl(P As VBProject, Cmpn, T As vbext_ComponentType, Cdl$)
WAddCmp P, Cmpn, T
AppCdl MdP(P, Cmpn), Cdl
End Sub
