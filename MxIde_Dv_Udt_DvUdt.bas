Attribute VB_Name = "MxIde_Dv_Udt_DvUdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dv_Udt_DvUdt."
Private Sub B_DvUdtMC():             DvUdtM CMd:                          End Sub
Private Sub B_DvUdtMdn():            DvUdtM MdMdn("MxIde_Dcl_Enm_TEnm"):  End Sub
Sub DvUdtMC():                       DvUdtM CMd:                          End Sub
Sub DvUdtMdn(Mdn):                   DvUdtM Md(Mdn):                      End Sub
Sub WrtMsrcDvUdtP(P As VBProject):   WrtMsrcy MsrcyDvUdtP(P):             End Sub
Private Sub DvUdtM(M As CodeModule): RplMdSrcopt M, SrcoptDvUdt(SrcM(M)): End Sub
