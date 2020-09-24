Attribute VB_Name = "MxMH_MHO"
'Public Const MHOPthiLgs$ = MHOPthLgs & "SapData\"
'Public Const PthMHDuty$ = MHOPthLgs & "DutyPrepay7\"
'Public Const PthMHTaxCpr$ = MHOPthLgs & "TaxCmp\"
'Public Const PthMHTaxAlrt$ = MHOPthLgs & "TaxAlert\"
'Public Const PthMHRelCst$ = MHOPthLgs & "RelCst\RelCst 1.0\"
'Public Const MHOMB52Pth$ = MHOPthLgs & "StockHolding8\"
'Public Const PthTpMHMB52$ = MHOMB52Pth & "WorkingDir\Templates\"
''------------------------------------------
'Public Const FbDtaMHStmt$ _
'                            = MHOPthStmt & "ARStmt_Data.accdb"
'Public Const FbDtaMHStmtVert$ _
'                            = MHOPthStmtVert & "ARStmt_Data.mdb"
'Public Const FbPgmMHStmtVert$ _
'                            = MHOPthStmtVert & "ARStmt.accdb"
'Public Const FbDtaMHStmtE$ _
'                            = MHOPthStmtE & "ARStmt_Data.accdb"
'Public Const FbDtaMHDuty$ _
'                            = PthMHDuty & "DutyPrepay7_Data.accdb"
'Public Const FbPgmMHDuty$ _
'                            = PthMHDuty & "DutyPrepay7.accdb"
'Public Const FbPgmMHTaxCpr$ _
'                            = PthMHTaxCpr & "TaxCmp v1.3.accdb"
'Public Const MHORelCst_FbPgm$ _
'                            = PthMHRelCst & "RelCst 1.0.accdb"
'Public Const MHOTaxAlert_FbPgm$ _
'                            = PthMHTaxAlrt & "TaxAlert 1.4\TaxAlert 1.4.accdb"
'Public Const MHOMB52_FbPgm$ _
'                            = MHOMB52Pth & "StockHolding8.accdb"
'Public Const MHOMB52_TpMB52$ _
'                            = PthTpMHMB52 & "On Hand Template.xlsx"
'Public Const MHOMB52_TpFc$ _
'                            = PthTpMHMB52 & "Forecast (Template).xlsx"
'Public Const TpMHMB52$ _
'                            = PthTpMHMB52 & "Stock Holding Template.xlsx"
'Public Const FbDtaMHMB52$ _
'                            = MHOMB52Pth & "StockHolding8_Data.accdb"
'Public Const MHOCRFbPgm$ _
'                            = MHOPthAr & "CrHldRls2\CrHldRls2.accdb"
'Public Const MHOMB52FxiSalTxt$ _
'                            = MHOPthiLgs & "Sales Text.xlsx"
'Public Const MHOMB52FxiGit$ _
'                            = MHOPthiLgs & "Sales Text.xlsx"
'Public Const FxiFcMHMB52$ _
'                            = MHOPthiLgs & "Sales Text.xlsx"
'Function FbPgmTmpMHMB52$(): FbPgmTmpMHMB52 = PthTmp & "StockHolding8Tmp.accdb": End Function
Option Compare Text
Option Explicit
Const CMod$ = "MxMH_MHO."
Type MHOApn
    Apnn As String
    Stmt As String
    StmtE As String
    StmtVert As String
    Aging As String
    CrRel As String
    
    MB52 As String
    TaxCpr As String
    TaxAlert As String
    RelCst As String
    Duty As String
    
    CrRvw As String
    OvrHd As String
End Type
'== AR
Type MHOCrRel
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: Tp As String
    MacroFxo As String
    MacroFxi As String
End Type
Type MHOAging
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    MacroFxo As String
    MacroFxi As String
End Type
Type MHOStmt
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    MacroFxo As String
    MacroFxi As String
End Type
Type MHOStmtE
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    MacroFxo As String
    MacroFxi As String
End Type
Type MHOStmtVert
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    MacroFxi As String
    MacroFxo As String
End Type
'== Ac
Type MHOOvrHd
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    MacroFxo As String
    MacroFxi As String
End Type
Type MHOCrRvw
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    MacroFxi As String
    MacroFxo As String
End Type
'== Lgs
Type MHOTaxCpr
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    MacroFxi As String
    MacroFxo As String
End Type
Type MHOTaxAlert
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    MacroFxi As String
    MacroFxo As String
End Type
Type MHORelCst
    Hom As String: FbPgm As String: FbDta As String: Pthi As String: Ptho As String: CnDta As ADODB.Connection: DbPgm As Database: DbDta As Database
    FxiSalTxt As String
    MacroFxi As String
    MacroFxo As String
    TpMB52 As String
    TpSHld As String
End Type
Function MHOPth()
Static X As Boolean
If Not X Then
    X = True
'    With Y
'        .Hom = "C:\Users\Public\"
'        .Ar = .Hom & "DebtorAging4 and ARStmt\"
'        .Stmt = MHOPthAr & "ARStmt\"
'        .StmtE = MHOPthAr & "ARStmt(eStmt)\"
'        .StmtVert = MHOPthAr & "ARStmt(VerticalFormat)\"
'        .Aging = MHOPthAr & "DebtorAging4\"
'        .CrRel = MHOPthAr & "CrHldRep2\"
'        .CrRvw = MHOPthHom & "CrRvw\"
'        .OvrHd = MHOPthHom & "OvrHd\"
'    End With
End If
End Function
Function MHOFbPgm()
'Public Const FbPgmMHOvrHd$ _
'                            = MHOPth.OvrHd & "ARStmt.accdb"
'Public Const FbDtaMHOvrHd$ _
'                            = MHOPthOvrHd & "ARStmt_Data.accdb"
'Public Const FbPgmMHCrRvw$ _
'                            = MHOPthCrRvw & "ARStmt.accdb"
'Public Const FbDtaMHCrRvw$ _
'                            = MHOPthCrRvw & "ARStmt_Data.accdb"
'Public Const FbPgmMHAging$ _
'                            = MHOPthAging & "ARStmt.accdb"
'Public Const FbDtaMHAging$ _
'                            = MHOPthAging & "ARStmt_Data.accdb"
'Public Const FbPgmMHStmt$ _
'                            = MHOPthStmt & "ARStmt.accdb"
'Public Const FbPgmMHStmtE$ _
'                            = MHOPthStmtE & "ARStmt.accdb"
'
End Function
Function DiMHDta() As Dictionary
'ClrBfr
'BfrV "Stmt     " & FbDtaMHStmt
'BfrV "Aging    " & FbDtaMHAging
'BfrV "CrRel    " & FbDtaMHAging
'BfrV "Duty     " & FbDtaMHDuty
'BfrV "StkHld   " & FbDtaMHMB52
'BfrV "OvrHdr   " & FbDtaMHOvrHd
'BfrV "CrRvw    " & FbDtaMHCrRvw
'Set DiMHDta = Diln(LyBfr)
End Function

Function DiMHPgm() As Dictionary
'ClrBfr
'BfrV "Stmt     " & FbPgmMHStmt
'BfrV "Aging    " & FbPgmMHAging
'BfrV "CrRel    " & FbPgmMHAging
'BfrV "TaxCpr   " & FbPgmMHTaxCpr
'BfrV "Duty     " & FbPgmMHDuty
'BfrV "StkHld   " & MHOMB52_FbPgm
'BfrV "RelCst   " & MHORelCst_FbPgm
'BfrV "TaxAlert " & MHOTaxAlert_FbPgm
'BfrV "OvrHdr   " & FbPgmMHOvrHd
'BfrV "CrRvw    " & FbPgmMHCrRvw
'Set DiMHPgm = Diln(LyBfr)
End Function
Sub MHOMB52_BrwQd(): BrwQd MHO.MHOMB52.DbPgm: End Sub
Function MHOApn() As MHOApn
Static X As Boolean, Y As MHOApn
If Not X Then
    X = True
    With Y
        .Apnn = "Aging Stmt CrRel Duty StkHld TaxAlrt TaxCpr RelCst CrRvw OvrHd"
        
    End With
End If
MHOApn = Y
End Function
