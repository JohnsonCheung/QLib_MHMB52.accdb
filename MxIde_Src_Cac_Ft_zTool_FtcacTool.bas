Attribute VB_Name = "MxIde_Src_Cac_Ft_zTool_FtcacTool"
Option Compare Text
Option Explicit

Sub ftcaczDmpMdnyOut(): Dmp MdnyFtcacOutP(CPj): End Sub
Sub ftcaczBrwPth():     BrwPth PthFtcacP(CPj):  End Sub
Sub ftcaczDmpMdnyMthNo()
Dim Ffn: For Each Ffn In Itr(MdnyFtcacMthNoP(CPj))
    Debug.Print FmtQQ("VcFt ""?""", Ffn)
Next
End Sub
Sub ftcaczVcTsyMi8CmntfbelMdn(Mdn$): Vc TsyMi8CmntfbelFtcaczMdn(Mdn):   End Sub
Sub ftcaczVcTsyMi8Cmntfbel():        Vc TsyMi8CmntfbelFtcacPC:          End Sub
Sub ftcaczDmpMD5():                  Debug.Print MD5Ft(FtCacMD5P(CPj)): End Sub
