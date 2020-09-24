Attribute VB_Name = "MxDao_Fea_Ccm"
'Ccm:Cml #[C]ir[c]umflex-accent#
'CcmTbl is ^xxx table in Db (pgm-database),
'          which should be same stru as N:\..._Data.accdb @ xxx
'          and   data should be copied from N:\..._Data.accdb for development purpose
'At the same time, in Db, there will be xxx as linked table either
'  1. In production, linking to N:\..._Data.accdb @ xxx
'  2. In development, linking to Db @ ^xxx
'Notes:
'  The TarFb (N:\..._Data.accdb) of each CcmTbl may be diff
'      They are stored in Description of CcmTbl manual, it is edited manually during development.
'  those xxx table in Db will be used in the program.
'  and ^xxx is create manually in development and should be deployed to N:\..._Data.accdb
'  assume Db always have some ^xxx, otherwise throw
'This Sub is to re-link the xxx in given [Db] to
'  1. [Db] if [TarFb] is not given
'  2. [TarFb] if [TarFb] is given.
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Ccm."

Private Sub B_TnyCcm()
Dim D As Database: Set D = MHO.MHORelCst.DbPgm
Ept = SplitSpc("^CurYM ^IniRate ^IniRateH ^InvD ^InvH ^YM ^YMGR ^YMGRnoIR ^YMOH ^YMRate")
GoSub Tst
Exit Sub
Tst:
    Act = TnyCcm(D)
    C
    Return
End Sub
Function TnyCcm(D As Database) As String():  TnyCcm = AwPfx(Tny(D), "^"): End Function
Function TnyCcmC() As String():             TnyCcmC = TnyCcm(CDb):        End Function

Sub ChkCcm(D As Database, Tny$(), FbExt$)
Dim DbExt As Database: Set DbExt = Db(FbExt)
Dim TLcl$(): TLcl = TnyCcm(D)
Dim Text$(): Text = TnyCcm(DbExt)
Dim TMis$()
Dim ErLclMis$()
    TMis = SyMinus(Tny, TLcl)    ' TnyCcm where not found in @D
    ErLclMis = EryHdrInd("", TMis)
Dim ErExtMis$()
    TMis = SyMinus(Tny, Text) ' TnyCcm where not found in *DbExt
    ErExtMis = EryHdrInd("", TMis)
Dim ErFldMis$()
    Dim TBth$(): TBth = AyIntersect(TLcl, Text)
    ErFldMis = WErFldMis(D, DbExt, TBth)
ChkEry SyAddAp(ErLclMis, ErExtMis, ErFldMis)
End Sub
Private Function WErFldMis(DbLcl As Database, DbExt As Database, TnyBth$()) As String()
Stop ''
End Function

Sub LnkCcmExt(D As Database, ExtFb$)
Stop ''
End Sub

Private Sub B_LnkCcmLcl()
Dim D As Database, IsLcl As Boolean
Set D = MHO.MHOMB52.DbDta
IsLcl = True
GoSub Tst
Exit Sub
Tst:
    LnkCcmLcl D
    Return
End Sub
Sub LnkCcmLcl(D As Database)
Const CSub$ = CMod & "LnkCcmLcl"
Dim T$()  ' All ^xxx
    T = TnyCcm(D)
    If Si(T) = 0 Then Thw CSub, "No ^xxx table in [Db]", D.Name 'Assume always
X_Lnk D, D.Name, T
End Sub
Private Sub X_Lnk(D As Database, ExtFb$, TnyCcm$())
Dim FbTar$: FbTar = D.Name
Dim TbnCcm: For Each TbnCcm In TnyCcm
    LnkFbt D, RmvFst(TbnCcm), ExtFb, TbnCcm
Next
End Sub
