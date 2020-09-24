Attribute VB_Name = "MxIde_Src_Cac_Db_TblFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcFtcac_Db_TblFun."
Function TbidC&(Tbn$, N$):               TbidC = Tbid(CDb, Tbn, N):                                                        End Function
Function Tbid&(D As Database, Tbn$, N$):  Tbid = ValQ(D, FmtQQ("Select [?Id] from [?] where [?n]='?'", Tbn, Tbn, Tbn, N)): End Function
Function Pjid&(Pjn$):                     Pjid = TbidC("Pj", Pjn):                                                         End Function
