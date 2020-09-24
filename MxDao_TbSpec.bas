Attribute VB_Name = "MxDao_TbSpec"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_TbSpec."

Sub BrwSpecPthD()
BrwPth SpecPth(CDb)
End Sub

Sub CrtTbSpec(D As Database)
'SchmCrt D, SplitVBar(SampSchmVbl)
End Sub

Sub EnsTbSpec(D As Database)
If Not HasT(D, "Spec") Then CrtTbSpec D
End Sub

Sub ExpSpecD(D As Database)
Dim P$: P = SpecPth(D)
DltAllPthFil P
Dim N:  For Each N In Itr(SpeCnoy(D))
    ExpSpec D, N, P
Next
End Sub

Sub ExpSpec(D As Database, Specn, ToPth$)
End Sub

Sub ImpSpec(D As Database, Specn)
Const CSub$ = CMod & "ImpSpec"
Dim Ft$
'    Ft = SpnmFt(Spnm)
    
Dim NoCur As Boolean
Dim NoLas As Boolean
Dim CurOld As Boolean
Dim CurNew As Boolean
Dim SamTim As Boolean
Dim DifSz As Boolean
Dim SamSz As Boolean
Dim DifFt As Boolean
Dim Rs As Dao.Recordset
    Dim Q$
    Q = FmtQQ("Select Specn,Ft,Lines,Tim,Si,LTimStr_Dte from Spec where Specn = '?'", Specn)
    Set Rs = D.OpenRecordset(Q)
    NoCur = NoFfn(Ft)
    'NoLas = HasRec(Rs)
    
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    Dim CurS&, LasS&
    Dim LasFt$, LdTimStr_Dte$
    CurS = FfnLen(Ft)
    CurT = FfnTim(Ft)
    If Not NoLas Then
        With Rs
            LasS = Nz(Rs!Si, -1)
            LasT = Nz(!Tim, 0)
            LasFt = Nz(!Ft, "")
'            LdTimStr_Dte = StrDte(!LTimStr_Dte)
        End With
    End If
    SamTim = CurT = LasT
    CurOld = CurT < LasT
    CurNew = CurT > LasT
    SamSz = CurS = LasS
    DifSz = Not SamSz
    DifFt = Ft <> LasFt
    

Const Imported$ = "***** IMPORTED ******"
Const NoImport$ = "----- no import -----"
Const NoCur______$ = "No Ft."
Const NoLas______$ = "No Last."
Const FtDif______$ = "Ft is dif."
Const SamTimSi___$ = "Sam tim & sz."
Const SamTimDifSz$ = "Sam tim & sz. (Odd!)"
Const CurIsOld___$ = "Cur is old."
Const CurIsNew_$ = "Cur is new."
Const C$ = "|[Specn] [Db] [Cur-Ft] [Las-Ft] [Cur-Tim] [Las-Tim] [Cur-Si] [Las-Si] [Imported-Time]."

Dim Dr()
Dr = Array(Specn, Ft, LinesFt(Ft), CurT, CurS, Now)
Select Case True
Case NoCur, SamTim:
'Case NoLas: InsDrRs Dr, Rs
'Case DifFt, CurNew: Dr_Upd_Rs Dr, Rs
Case Else: Stop
End Select

Dim Av()
Av = Array(Specn, D.Name, Ft, LasFt, CurT, LasT, CurS, LasS, LdTimStr_Dte)
Select Case True
'Case NoCur:            XDmp_Ln_AV CSub, NoImport & NoCur______ & C, Av
'Case NoLas:            XDmp_Ln_AV CSub, Imported & NoLas______ & C, Av
'Case DifFt:            XDmp_Ln_AV CSub, Imported & FtDif______ & C, Av
'Case SamTim And SamSz: XDmp_Ln_AV CSub, NoImport & SamTimSi___ & C, Av
'Case SamTim And DifSz: XDmp_Ln_AV CSub, NoImport & SamTimDifSz & C, Av
'Case CurOld:           XDmp_Ln_AV CSub, NoImport & CurIsOld___ & C, Av
'Case CurNew:           XDmp_Ln_AV CSub, Imported & CurIsNew_ & C, Av
Case Else: Stop
End Select
End Sub

Function SpeCnoy(D As Database) As String()
SpeCnoy = DcStrTF(D, "Spec.Specn")
End Function

Function SpecPth$(D As Database)
SpecPth = PthAss(D.Name) & ".Spec\"
End Function
