Attribute VB_Name = "MxVb_ObjPrpDa"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_ObjPrpDa."

Function DrsItp(Itr, SsPrpp$, Optional FF$) As Drs
'@SsPrpp:: :SS ! #Prp-Pth-Spc-Sep#
Dim Prppy$(): Prppy = SySs(SsPrpp)
DrsItp = DrsFf(StrDft(FF, SsPrpp), X_1Dy(Itr, Prppy))
End Function

Function DrsItrPy(Itr, Prppy$()) As Drs
DrsItrPy = Drs(Prppy, X_1Dy(Itr, Prppy))
End Function

Private Function X_1Dy(Itr, Prppy$()) As Variant()
Dim Obj As Object: For Each Obj In Itr
    Push X_1Dy, Opvy(Obj, Prppy)
Next
End Function

Function QuietOpv(Obj, P)
On Error Resume Next
Asg Opv(Obj, P), QuietOpv
End Function

Private Sub B_DrsItrPy()
BrwDrs DrsItp(Excel.Application.AddIns, "Name Installed IsOpen FullName CLSId ")
'BrwDrs DrsItpcc(Fds(Db(MHDutyDtaFb), "Permit"), "Name Type Required")
'BrwDrs ItpDrs(Application.VBE.VBProjects, "Name Type")
'BrwDrs DrsItrPy(CPj.VBComponents, SySs("Name Type CmpTy=ShtCmpTy(Type)"))
End Sub
