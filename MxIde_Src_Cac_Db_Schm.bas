Attribute VB_Name = "MxIde_Src_Cac_Db_Schm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcDbcac_Db_Schm."
Function FbSrcDbcacPC$():                FbSrcDbcacPC = FbSrcDbcacP(CPj):                                 End Function
Function FbSrcDbcacP$(P As VBProject):    FbSrcDbcacP = PthSrcDbcacP(P) & Fn(Pjf(P)) & ".SrcDbcac.accdb": End Function
Function PthSrcDbcacPC$():              PthSrcDbcacPC = PthSrcDbcacP(CPj):                                End Function
Function PthSrcDbcacP$(P As VBProject):  PthSrcDbcacP = PthAddFdrEns(PthAssP(P), ".SrcDbcac"):            End Function

Sub EnsSrcDbcacPC():              EnsSrcDbcacP CPj:    End Sub
Sub EnsSrcDbcacP(P As VBProject): WEns FbSrcDbcacP(P): End Sub
Private Sub WEns(Fb$)
EnsFb FbSrcDbcacPC
Dim D As Database
Set D = Db(FbSrcDbcacPC)
SchmEns D, SchmSrcDbcac
End Sub

Function DbSrcDbcacP(P As VBProject) As Database:  Set DbSrcDbcacP = Db(FbSrcDbcacP(P)): End Function
Function DbSrcDbcacPC() As Database:              Set DbSrcDbcacPC = DbSrcDbcacP(CPj):   End Function

Sub BrwDbSrcDbcac()
BrwDb DbSrcDbcacPC
End Sub

Function SchmSrcDbcac() As String()
Erase XX
X "Fld"
X " Nm  Md Pj"
X " T50 MchStr"
X " T10 MthPfx"
X " Txt Pjf Prm Ret LinRmk"
X " T3  Ty Mdy"
X " T4  CmpTy"
X " Lng Lno"
X " Mem Lines Mmk"
X "Tbl"
X " Pj  *Id Pjf | Pjn PjDte"
X " Md  *Id PjId Mdn | CmpTy"
X " Mth *Id MdId Mthn ShtTy | ShtMdy Prm Ret LinRmk Mmk Lines Lno"
SchmSrcDbcac = XX
Erase XX
End Function
