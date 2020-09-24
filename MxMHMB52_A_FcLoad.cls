VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcLoad."
Private Type WLgcDta
    Fxi As String
    Wsinn As String
    Fxw As String
End Type
Private Fun As New MxMHMB52_MHFun
Friend Sub A_ShwLgc()
End Sub
Friend Sub A_ShwLgcUD()

End Sub

Friend Sub A_LnkBrwFxwMH():    WLnkFxwMH MH.FcLas.StmYmMH: BrwTC ">Fc":              End Sub
Friend Sub A_LnkBrwFxwUD():    WLnkFxwUD8687 MH.FcLas.StmYmUD: BrwTTC ">Fc86 >Fc87": End Sub
Private Sub B_WTmpIFcMH():     WTmpIFcMH:                                            End Sub
Private Sub B_WTmpIFcUD8687(): WTmpIFcUD8687:                                        End Sub
Private Sub B_WChk__TmpFc():
Call WTmpFc_MH: WChk__TmpFc "M"
Call WTmpFc_UD: WChk__TmpFc "U"
End Sub
Private Function WWSampStmYmUD() As StmYm
Static X As StmYm: If X.Stm = "" Then WWSampStmYmUD = MH.FcLas.StmYmUD
End Function
Private Function WWSampStmYmMH() As StmYm
Static X As StmYm: If X.Stm = "" Then WWSampStmYmMH = MH.FcLas.StmYmUD
End Function
Private Sub B_WInsTb_FcSku()
Dim A  As StmYm
If True Then
    A = WWSampStmYmMH
    WLnkFxwMH A
    WTmpIFcMH
    WTmpFc_MH
    WInsTb_FcSku A
Else
    A = WWSampStmYmUD
    WLnkFxwUD8687 A
    WTmpIFcUD8687
    WTmpFc_UD
    WInsTb_FcSku A
End If
End Sub

Sub LoadFcLasMH():           LoadFcLas "M":                       End Sub
Sub LoadFcLasUD():           LoadFcLas "U":                       End Sub
Private Sub LoadFcLas(Stm$): LoadFcStmYm MH.FcLas.StmYmPthi(Stm): End Sub
Sub LoadFcStmYm(M As StmYm)
'Inp: MH.FcIO.Fxi .WsnMH or .Fxi .WsnUD(Co)]
'Rfh: FcSku: VerYY VerMM Co Stm Sku M01 M02 M03 M04 M05 M06 M07 M08 M09 M10 M11 M12 M13 M14 M15
'Rfh: Fc   : VerYY VerMM Stm Siz Tim C87NSku C87N* C87Sc C86NSku C86N* C86Sc DteLoad DteCrt

Select Case M.Stm
Case "M"
    MH.FcFxwMH.CrtFxw M
    WLnkFxwMH M
    WTmpIFcMH
    WTmpFc_MH
    WChk__TmpFc "M"
    WInsTb_FcSku M    ' Load #IFcSku into FcSku   FcSku   = VerYY VerMM Co Sku M01..M15
    DrpApC ">Fc86", ">Fc87", "#IFc86", "IFc87", "#Fc"
Case "U"
    MH.FcFxwUD.CrtFxw M
    WLnkFxwUD8687 M
    WTmpIFcUD8687
    WTmpFc_UD
    WChk__TmpFc "U"
    WInsTb_FcSku M    ' Load #IFcSku into FcSku   FcSku   = VerYY VerMM Co Sku M01..M15
    DrpApC ">Fc", "#IFc", "#Fc"
End Select
'
MH.TbFcFmTbFcSku.RfhStmYm M
ClrSts
If IsFrmnOpn("LoadFc") Then Form_LoadFc.Requery
Done
End Sub
Private Sub WLnkFxwMH(A As StmYm): LnkFxwC ">Fc", MH.FcIO.Fxw(A), MH.FcFxwMH.Wsn: End Sub
Private Sub WLnkFxwUD8687(A As StmYm)
LnkFxwC ">Fc86", MH.FcIO.Fxw(A), MH.FcIO.WsnUD(86)
LnkFxwC ">Fc87", MH.FcIO.Fxw(A), MH.FcIO.WsnUD(87)
End Sub
Private Sub WTmpIFcMH()
DrpTTC "#IFc"
RunqC SqlIntoSelStar("#IFc", ">Fc") & WhFeq("FcTyMH", MH.FcFxwMH.TyFc)
RunqC SqlColAdd("#IFc", "Co Byte")
RunqC "Update [#IFc] set Co=86 where CoNm='HONG KONG DP'"
RunqC "Update [#IFc] set Co=87 where CoNm='MACAO DP'"
RunqC SqlDrpCol("#IFc", "Market,CoNm,FcTyMH")
End Sub
Private Sub WTmpIFcUD8687()
'Import >IFc86 to #IFc86
'And    >IFc87 to #IFc87
DrpTTC "#IFc86 #IFc87"
Dim Sel$: Sel = "Select Sku," & JnPfxSno("M", 15)
RunqC Sel & QpIntoFm("#IFc86", ">Fc86")
RunqC Sel & QpIntoFm("#IFc87", ">Fc87")
End Sub
Private Sub WTmpFc_MH()
'Inp: >Fc
'Oup: creaet #Fc Co Sku M01..M15
Dim Sel$: Sel = WQp_SelSkuM12 & ",Co"
Dim IntoTmpFc$: IntoTmpFc = QpInto("#Fc")
Dim Fm$: Fm = QpFm("#IFc", "x")
Dim Wh$: Wh = WQp_WhToDlt & ")"
DrpC "#Fc"
RunqC Sel & IntoTmpFc & Fm & Wh
RunqC "alter table [#Fc] add column M13 Double,M14 Double,M15 double"
WWUpd_TmpFc_SetNul
End Sub
Private Sub WTmpFc_UD()
'Inp: #IFc86 & #IFc87
'Oup: #Fc
Const Las3M$ = ",Val(Nz(x.M13,0)) As M13," & _
         "Val(Nz(x.M14,0)) As M14," & _
         "Val(Nz(x.M15,0)) As M15"
Dim Sel86$: Sel86 = WQp_SelSkuM12 & Las3M & ",CByte(86) as Co"
Dim Sel87$: Sel87 = WQp_SelSkuM12 & Las3M & ",CByte(87) as Co"
Dim IntoTmpFc$: IntoTmpFc = QpInto("#Fc")
Dim IntoTmpA$: IntoTmpA = QpInto("#A")
Dim Fm86$: Fm86 = QpFm("#IFc86", "x")
Dim Fm87$: Fm87 = QpFm("IFc87", "x")
Dim Wh$: Wh = WQp_WhToDlt & _
                " or Val(Nz(M13,0))<>0" & _
                " or Val(Nz(M14,0))<>0" & _
                " or Val(Nz(M15,0))<>0)"
Dim InsTmpFc$: InsTmpFc$ = QpInsInto("#Fc")
DoCmd.SetWarnings True
RunqC Sel86 & IntoTmpFc & Fm86 & Wh
RunqC InsTmpFc & " " & Sel87 & Fm87 & Wh
DoCmd.SetWarnings False
WWUpd_TmpFc_SetNul
'Set each fields-of-Mnn to null if it is 0
End Sub

Private Sub WInsTb_FcSku(A As StmYm)
'Inp: #Fc     =             Co Sku M01..15 Assume all Sku is in the Stm as @A
'Oup: FcSku   = VerYY VerMM Co Stm SKu M01..15 So #A is just add
'Stp: FcSku   : Delete where VerYY VerMM Co Stm
'     FcSku   : Append from #A

'-- #Sku -> Delete FcSku
RunqC "Delete * from FcSku" & MH.FcWh.WhStmYm(A)
With A
Dim Sql$: Sql = FmtQQ("insert into FcSku" & _
    "(VerYY,VerMM,Co,  Stm,Sku,M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,M13,M14,M15) Select" & _
    " ?    ,    ?,Co,'?',Sku,M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,M13,M14,M15" & _
    " from [#Fc]", .Y, .M, .Stm)
    End With
DoCmd.SetWarnings True
RunqC Sql
DoCmd.SetWarnings False
End Sub

Private Sub WChk__TmpFc(Stm$)
Dim E1$(): E1 = WChk_ErStmCo(Stm, 86)
Dim E2$(): E2 = WChk_ErStmCo(Stm, 87)
ChkEry SyAdd(E1, E2)
End Sub
Private Function WChk_ErStmCo(Stm$, Co As Byte) As String()
Dim Dup$(), NFnd$(), WrongStm$() ' Skuy
Dim Skuy$(): Skuy = DcStrTFC("#Fc.Sku", "Co=" & Co)
WChk_Set3SkuyEr Stm, Skuy$(), Dup, NFnd, WrongStm
Dim E1$(): E1 = WChk_ErDup(Dup)
Dim E2$(): E2 = WChk_ErNFnd(NFnd)
Dim E3$(): E3 = WChk_ErWrongStm(WrongStm, Stm)
ChkEry SyAddAp(E1, E2, E3)
End Function
Private Sub WChk_Set3SkuyEr(Stm$, Skuy$(), ODup$(), ONFnd$(), OWrongStm$())
Dim SkuyAll$(): SkuyAll = MH.Skuy.SkuyTbSku
ODup = AwDup(Skuy)
Dim SkuyStm$(): SkuyStm = MH.Skuy.SkuyStm(Stm)
ONFnd = SyMinus(Skuy, SkuyStm)
Dim Fnd$(): Fnd = SyMinus(Skuy, ONFnd)
OWrongStm = MH.Skuy.SkuyWrongStm(Fnd, Stm)
End Sub
Private Function WChk_ErNFnd(NFnd$()) As String()
If Si(NFnd) = 0 Then Exit Function
Dim O$()
PushI O, "Following " & Si(NFnd) & " Sku in the Forecast importing Excel files are not found in the Sku table!"
PushI O, "This means, these Skus do not have any record found in [Sales Text.xlsx] file."
PushI O, "So, make sure they are found in [Sales Text.xlsx] and import it.  Then, import the forecast again"
PushIAy O, AyTab(NFnd)
Stop: Exit Function '
WChk_ErNFnd = O
End Function
Private Function WChk_ErWrongStm(WrongStm$(), Stm$) As String()
If Si(WrongStm) = 0 Then Exit Function
Dim O$()
PushI O, "Following " & Si(WrongStm) & " Sku in the Forecast importing Excel files should be [" & Fun.Stm2StmOth(Stm) & "] Sku"
PushI O, "This means, these Skus are in the wrong forecast file which is importing!"
PushI O, "A Sku is MH Sku, if the it's [Topaz Code] in the [sales text.xlsx]<>'UDVPROD'"
PushI O, "A Sku is UD Sku, if the it's [Topaz Code] in the [sales text.xlsx]='UDVPROD'"
PushIAy O, AyTab(WrongStm)
WChk_ErWrongStm = O
End Function
Private Function WChk_ErDup(Dup$()) As String()
If Si(Dup) = 0 Then Exit Function
Dim O$()
PushI O, "Following " & Si(Dup) & " Sku in the Forecast importing Excel files are found to have 2 or more forecast data row"
PushI O, "Please fix to have only one forecast data row and import again"
PushIAy O, AyTab(Dup)
WChk_ErDup = O
End Function

Private Function WQp_WhToDlt$()
WQp_WhToDlt = _
Wh("CStr(Nz(Sku,''))<>''" & _
" and (Val(Nz(M01,0))<>0" & _
  " or Val(Nz(M02,0))<>0" & _
  " or Val(Nz(M03,0))<>0" & _
  " or Val(Nz(M04,0))<>0" & _
  " or Val(Nz(M05,0))<>0" & _
  " or Val(Nz(M06,0))<>0" & _
  " or Val(Nz(M07,0))<>0" & _
  " or Val(Nz(M08,0))<>0" & _
  " or Val(Nz(M09,0))<>0" & _
  " or Val(Nz(M10,0))<>0" & _
  " or Val(Nz(M11,0))<>0" & _
  " or Val(Nz(M12,0))<>0")
End Function

Private Sub WWUpd_TmpFc_SetNul()
Dim J%: For J = 1 To 15
    Dim F$: F = "M" & Format(J, "00")
    Dim E$: E = F & "=0"
    Dim S$: S = "Update [#Fc] set " & F & "=Null Where " & E
    RunqC S
Next
End Sub

Private Function WQp_SelSkuM12$()
WQp_SelSkuM12 = "Select CStr(x.Sku) As Sku," & _
    "Val(Nz(x.M01,0)) As M01," & _
    "Val(Nz(x.M02,0)) As M02," & _
    "Val(Nz(x.M03,0)) As M03," & _
    "Val(Nz(x.M04,0)) As M04," & _
    "Val(Nz(x.M05,0)) As M05," & _
    "Val(Nz(x.M06,0)) As M06," & _
    "Val(Nz(x.M07,0)) As M07," & _
    "Val(Nz(x.M08,0)) As M08," & _
    "Val(Nz(x.M09,0)) As M09," & _
    "Val(Nz(x.M10,0)) As M10," & _
    "Val(Nz(x.M11,0)) As M11," & _
    "Val(Nz(x.M12,0)) As M12"
End Function
