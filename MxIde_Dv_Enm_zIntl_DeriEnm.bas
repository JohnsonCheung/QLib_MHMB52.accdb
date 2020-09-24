Attribute VB_Name = "MxIde_Dv_Enm_zIntl_DeriEnm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Deri_Enm_DeriEnm."
Function DvenmSrc(Src$()) As String()
Dim Blky(): Stop ' Blky = SrcEnmDvableBlky(Src)
Dim ISrc$(): ISrc = Src
Dim J%: For J = 0 To UB(Blky)
    ISrc = BlkDvenmSrc(Src, EnmBlk:=CvSy(Blky(J)))
Next
End Function
Function BlkDvenmSrc(Src$(), EnmBlk$()) As String()
Dim U As TEnm: U = TEnmBlk(EnmBlk)
Dim ISrc$(): Stop 'ISrc = EnsEnmCnstCdlSrc(Src, EnmCnstCdl(U))
Stop 'BlkDvenmSrc = EnsEnmFunCdlSrc(ISrc, EnmFunCdl(U))
End Function
