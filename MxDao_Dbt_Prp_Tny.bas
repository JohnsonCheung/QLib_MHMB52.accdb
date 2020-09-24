Attribute VB_Name = "MxDao_Dbt_Prp_Tny"
Option Compare Text
Const CMod$ = "MxDao_Dbt_Prp_Tny."
Option Explicit

Function TnyC() As String():             TnyC = Tny(CDb):                        End Function
Function TTC$():                          TTC = TT(CDb):                         End Function
Function TT$(D As Database):               TT = Tml(Tny(D)):                     End Function
Function Tny(D As Database) As String():  Tny = AePfx(Itn(D.TableDefs), "MSys"): End Function

Function TnyLnkC() As String():        TnyLnkC = TnyLnk(CDb):      End Function
Function TnyInpC() As String():        TnyInpC = TnyInp(CDb):      End Function
Function TnyTmpInpC() As String():  TnyTmpInpC = TnyTmpInp(CDb):   End Function
Function TnyTmpC() As String():        TnyTmpC = TnyTmp(CDb):      End Function
Function TnyHshC() As String():        TnyHshC = TnyHsh(CDb):      End Function
Function TnyOupC() As String():        TnyOupC = TnyOup(CDb):      End Function
Function TnyPfxC(Pfx$) As String():    TnyPfxC = TnyPfx(CDb, Pfx): End Function

Function TnyLnk(D As Database) As String():          TnyLnk = ItnWhPrpBlnk(D.TableDefs, "Connect"): End Function
Function TnyInp(D As Database) As String():          TnyInp = TnyPfx(D, ">"):                       End Function
Function TnyTmp(D As Database) As String():          TnyTmp = TnyPfx(D, "$"):                       End Function
Function TnyTmpInp(D As Database) As String():    TnyTmpInp = TnyPfx(D, "#I"):                      End Function
Function TnyHsh(D As Database) As String():          TnyHsh = AePfx(TnyPfx(D, "#"), "#I"):          End Function
Function TnyOup(D As Database) As String():          TnyOup = TnyPfx(D, "@"):                       End Function
Function TnyPfx(D As Database, Pfx$) As String():    TnyPfx = AwPfx(Tny(D), Pfx):                   End Function

Function Tni(D As Database): Asg Itr(Tny(D)), Tni: End Function
Function TnyLcl(D As Database) As String(): TnyLcl = AePfx(ItnWhPrpBlnk(D.TableDefs, "Connect"), "MSys"): End Function

Function Tny1(D As Database) As String()
Dim T As TableDef, O$()
Dim X As Dao.TableDefAttributeEnum
X = Dao.TableDefAttributeEnum.dbHiddenObject Or Dao.TableDefAttributeEnum.dbSystemObject
For Each T In D.TableDefs
    Select Case True
    Case T.Attributes And X
    Case Else
        PushI Tny1, T.Name
    End Select
Next
End Function

Function TnyMSysObj(D As Database) As String(): TnyMSysObj = DcStrQ(D, SqlTbnMSysObj): End Function
