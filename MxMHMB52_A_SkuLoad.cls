VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SkuLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Private Type WEu
    MsgMat() As String
    MsgSkuDes() As String
    MsgTopaz() As String
    MsgPH() As String
    MsgStkUnit() As String
    MsgStkUnit_ValEr() As String 'Must be PCE COL
    MsgUnit_per_AC() As String
    MsgUnit_per_SC() As String
    MsgLitre_per_Btl() As String
    MsgBusArea() As String
End Type
Const LvcNmf = "Material,Sales text language 1,Topaz Code,Product hierarchy,Base Unit of Measure,Unit per case,SC,COL Per Case,Bottle Capacity,Business Area"

Const Extn_BusArea$ = "Business Area"
Const Extn_Liter_per_Btl$ = "Bottle Capacity"
Const Extn_Mat$ = "Material"
Const Extn_PH$ = "Product hierarchy"
Const Extn_SkuDes$ = "Sales text language 1"
Const Extn_StkUnit$ = "Base Unit of Measure"
Const Extn_Topaz$ = "Topaz Code"
Const Extn_Btl_per_AC$ = "COL Per Case"
Const Extn_Unit_per_AC$ = "COL Per Case"
Const Extn_Unit_per_SC$ = "SC"
Const Intn_BusArea_Sap$ = "BusAreaSap"
Const Intn_Liter_per_Btl_Sap$ = "Litre/BtlSap"
Const Intn_Mat$ = "Sku"
Const Intn_PH$ = "ProdHierarchy"
Const Intn_SkuDes$ = "SkuDes"
Const Intn_StkUnit$ = "StkUnit"
Const Intn_Topaz$ = "CdTopaz"
Const Intn_Btl_per_AC$ = "Btl/AC"
Const Intn_Unit_per_AC$ = "Unit/AC"
Const Intn_Unit_per_SC$ = "Unit/SC"
Sub LoadSku()
DoCmd.SetWarnings False
Dim Tim As Date: Tim = Now
W_ChkFxiColTy
W_LnkFxi
W_WarnNullMaterial
W_Tmp8687
W_ChkFxiEr__Er
W_InsSkuTopaz
W_Upd Tim
W_InsSku Tim
W_PhRfhTbPhLBus_FmSku_NewBusArea
MH.TbSku.RfhFldOvr
DrpTTC "#86Sku #87Sku >Sku86 >Sku87"
If Form_LoadSku.CurrentView = 1 Then Form_LoadSku.Requery ' F.CurrentView = 1 means form view
W_PromptRslt Tim
End Sub
Private Sub W_PhRfhTbPhLBus_FmSku_NewBusArea()
'Aim: For those Sku->BusAreaSap is not found in PHLBus, add a new record to [PHLBus] and given notePad message.
'     It is called by SkuLoad subr (Whenever there is new record added to Sku, call this subr
'     PHLBus: BusArea (4chr) PHBus (Des
RunqC "Select Distinct BusArea into [#BusArea] from Sku where Nz(BusArea,'')<>''"
RunqC "Select x.BusArea into [#BusAreaNew] from [#BusArea] x left join PHLBus a on x.BusArea=a.BusArea where a.BusArea is null"
RunqC "Insert into PHLBus (BusArea, PHBus) Select BusArea,BusArea & ' Des' as PHBus from [#BusAreaNew]"
Dim N%: N = NRecTC("#BusAreaNew")
If N > 0 Then
    MsgBox "There are [" & N & "] new business area are found, please go enter their description", vbInformation
End If
DrpTTC "#BusArea #BusAreaNew"
End Sub
Private Function WWFix$()
Static Fxi$: If Fxi = "" Then Fxi = MH.SkuPm.Fxi
WWFix = Fxi
End Function
Private Sub W_LnkFxi()
LnkFxwC ">Sku86", WWFix, "8601"
LnkFxwC ">Sku87", WWFix, "8701"
End Sub

Private Sub W_ChkFxiColTy()
'-- Stp-Lnk-and-Chk-InpFx
Const FldnCsv$ = "Material,Sales text language 1,Topaz Code,Product hierarchy,Base Unit of Measure,Unit per case,SC,COL Per Case,Bottle Capacity,Business Area"
Const TycCsv$ = " T       ,T                    ,T         ,T                ,T                   ,TorN         ,N ,N           ,N              ,T"
ChkFxww WWFix, "8601 8701"
ChkFxwCol WWFix, "8601", FldnCsv, TycCsv
ChkFxwCol WWFix, "8701", FldnCsv, TycCsv
End Sub

Private Sub W_WarnNullMaterial()
Dim N86&: N86 = ValQC("Select Count(*) from [>Sku86] where Trim(Nz(Material,''))=''")
Dim N87&: N87 = ValQC("Select Count(*) from [>Sku87] where Trim(Nz(Material,''))=''")
If N86 > 0 Or N87 > 0 Then
    MsgBox _
    "There are [" & N86 & "] lines in 8601 worksheet with empty [Material]" & vbCrLf & _
    "There are [" & N87 & "] lines in 8701 worksheet with empty [Material]" & vbCrLf & vbCrLf & _
    "These lines are ignored!" & vbCrLf & vbCrLf & _
    "[Ok]=Continue", vbInformation
End If
End Sub
Private Sub W_Tmp8687()
Sts "Running import query ....."
'Crt: #86Sku & #87Sku
'Fm : >Sku86 / >Sku87
DrpTTC "#86Sku #87Sku"
Const Q$ = "Select " & _
"Trim(Nz(Material               ,'')) AS SKU," & _
"Trim(Nz(`Sales text language 1`,'')) as SkuDes," & _
"Trim(Nz([Topaz Code]           ,'')) as CdTopaz," & _
"Trim(Nz([Product hierarchy]    ,'')) as ProdHierarchy," & _
"Trim(Nz([Base Unit of Measure] ,'')) as StkUnit," & _
"Val(Nz([Unit per case]          ,0)) as [Unit/Ac]," & _
"Val(Nz([SC]                     ,0)) as [Unit/SC]," & _
"Val(Nz([COL Per Case]           ,0)) As [Btl/AC]," & _
"Val(Nz([Bottle Capacity]        ,0)) As [Litre/BtlSap]," & _
"Trim(Nz([Business Area]         ,0)) As BusAreaSap," & _
"CLng(0)                              As Topaz" & _
" into [#?Sku] from [>Sku?] where Trim(Nz(Material,''))<>''"
RunqQQC Q, "86"
RunqQQC Q, "87"
End Sub
Private Sub B_W_ChkFxiEr__Er()
W_LnkFxi
W_Tmp8687
W_ChkFxiEr__Er
End Sub
Private Sub W_ChkFxiEr__Er()
If MsgBox("Do data checking?", vbYesNo) = vbNo Then Exit Sub
Dim O$(): O = SyAdd(WEr__Er(86), WEr__Er(87))
If Si(O) = 0 Then Exit Sub
BrwAy O
If MsgBox("There are errors in the Sales Text Excel file." & vbCrLf & vbCrLf & "[Ok]=Continue loading Sales Text with some data missing...!" & vbCrLf & "[Cancel]=Cancel", vbQuestion + vbOKCancel) = vbOK Then Exit Sub
Raise "Canceled"
End Sub
Private Function WEr__Er(Co As Byte) As String()
Dim Eu As WEu: Eu = WEu(Co)
Dim O$()
PushI O, "There are errors in the Sales Text Excel files:"
PushI O, "==============================================="
PushI O, "Excel Files: [" & WWFix & "]"
PushI O, "Worksheet  : [" & WEr_WsnCo(Co) & "]"
With Eu
    PushAy O, WEr_Msg_BlnkFld(.MsgBusArea, Extn_BusArea)
    PushAy O, WEr_Msg_ZerOrNeg(.MsgLitre_per_Btl, Extn_Liter_per_Btl)
    PushAy O, WEr_Msg_BlnkFld(.MsgSkuDes, Extn_Mat)
    PushAy O, WEr_Msg_BlnkFld(.MsgPH, Extn_PH)
    PushAy O, WEr_Msg_BlnkFld(.MsgSkuDes, Extn_SkuDes)
    PushAy O, WEr_Msg_BlnkFld(.MsgStkUnit, Extn_StkUnit)
    PushAy O, WEr_Msg_StkUnitValEr(.MsgStkUnit_ValEr, Extn_StkUnit)
    PushAy O, WEr_Msg_BlnkFld(.MsgTopaz, Extn_Topaz)
    PushAy O, WEr_Msg_ZerOrNeg(.MsgUnit_per_AC, Extn_Unit_per_AC)
    PushAy O, WEr_Msg_ZerOrNeg(.MsgUnit_per_SC, Extn_Unit_per_SC)
    PushI O, ""
End With
WEr__Er = O
End Function
Private Function WEu(Co As Byte) As WEu
Dim T$: T = "#" & Co
With WEu
    .MsgTopaz = WEr_SkuDesy_BlnkVal(Co, Intn_Topaz)
    .MsgBusArea = WEr_SkuDesy_BlnkVal(Co, Intn_BusArea_Sap)
    .MsgLitre_per_Btl = WEr_SkuDesy_FldVal_is_ZerOrNeg(Co, Intn_Liter_per_Btl_Sap)
    .MsgMat = WEr_SkuDesy_BlnkVal(Co, Intn_Mat)
    .MsgPH = WEr_SkuDesy_BlnkVal(Co, Intn_PH)
    .MsgSkuDes = WEr_SkuDesy_BlnkVal(Co, Intn_Mat)
    .MsgStkUnit = WEr_SkuDesy_BlnkVal(Co, Intn_StkUnit)
    .MsgTopaz = WEr_SkuDesy_BlnkVal(Co, Intn_Topaz)
    .MsgUnit_per_AC = WEr_SkuDesy_FldVal_is_ZerOrNeg(Co, Intn_Unit_per_AC)
    .MsgUnit_per_SC = WEr_SkuDesy_FldVal_is_ZerOrNeg(Co, Intn_Unit_per_SC)
    .MsgStkUnit_ValEr = WEr_SkuDesy_FmSql("Select Sku,SkuDes from [#" & Co & "Sku] where Not StkUnit in ('COL','PCE')")
End With
End Function
Private Sub B_WEr_SkuDesy_BlnkVal(): DmpAy WEr_SkuDesy_BlnkVal(86, Intn_Topaz): End Sub
Private Function WEr_SkuDesy_FmSql(Sql) As String()
With RsTblC(Sql)
    While Not .EOF
        PushI WEr_SkuDesy_FmSql, "Sku[" & !SKU & "] Des[" & !SkuDes & "]"
        .MoveNext
    Wend
End With
End Function
Private Function WEr_SkuDesy_BlnkVal(Co As Byte, IntNm$) As String()
WEr_SkuDesy_BlnkVal = WEr_SkuDesy_FmSql(FmtQQ("Select SKu,SkuDes from [#?Sku] where Trim(Nz([?],''))=''", Co, IntNm))
End Function

Private Function WEr_WsnCo$(Co As Byte): WEr_WsnCo = Co & "01": End Function
Private Function WEr_IsEuEr(A As WEu) As Boolean
With A
Select Case True
Case _
Si(.MsgBusArea) <> 0, _
Si(.MsgLitre_per_Btl) <> 0, _
Si(.MsgMat) <> 0, _
Si(.MsgPH) <> 0, _
Si(.MsgSkuDes) <> 0, _
Si(.MsgStkUnit) <> 0, _
Si(.MsgStkUnit_ValEr) <> 0, _
Si(.MsgTopaz) <> 0, _
Si(.MsgUnit_per_AC) <> 0, _
Si(.MsgUnit_per_SC) <> 0
Exit Function
End Select
End With
WEr_IsEuEr = True
End Function
Private Function WEr_Msg_ZerOrNeg(Msg$(), Extn$) As String()
If Si(Msg) = 0 Then Exit Function
PushS WEr_Msg_ZerOrNeg, "Column[" & Extn & "] has [" & Si(Msg) & "] lines with no positive numeric value"
PushAy WEr_Msg_ZerOrNeg, AmAddPfxTab(Msg)
End Function
Private Function WEr_Msg_StkUnitValEr(Msg$(), Extn$) As String()
If Si(Msg) = 0 Then Exit Function
PushS WEr_Msg_StkUnitValEr, "Column[" & Extn & "] has [" & Si(Msg) & "] lines has invalid value. Valid value should be COL PCE:"
PushAy WEr_Msg_StkUnitValEr, AmAddPfxTab(Msg)
End Function
Private Function WEr_Msg_BlnkFld(Msg$(), Extn$) As String()
If Si(Msg) = 0 Then Exit Function
PushS WEr_Msg_BlnkFld, "Column[" & Extn & "] has [" & Si(Msg) & "] lines blank value:"
PushAy WEr_Msg_BlnkFld, AmAddPfxTab(Msg)
End Function

Private Function WEr_SkuDesy_FldVal_is_ZerOrNeg(Co As Byte, IntNm$) As String()
Dim Sql$: Sql = FmtQQ("Select Sku,SkuDes from [#?Sku] where Nz([?],0)<=0", Co, IntNm)
WEr_SkuDesy_FldVal_is_ZerOrNeg = WEr_SkuDesy_FmSql(Sql)
End Function

Private Sub W_InsSkuTopaz()
'== Stp-InsTbl_Topaz
RunqC "Select Distinct CdTopaz into `#A` from `#86Sku`"
RunqC "Select Distinct CdTopaz into `#B` from `#87Sku`"
RunqC "Insert into Topaz Select x.CdTopaz from `#A` x left Join Topaz a on x.CdTopaz=a.CdTopaz where a.CdTopaz is null"
RunqC "Insert into Topaz Select x.CdTopaz from `#B` x left Join Topaz a on x.CdTopaz=a.CdTopaz where a.CdTopaz is null"
RunqC "Drop Table `#A`"
RunqC "Drop Table `#B`"

'== Stp-UpdTmp8687-Fld-Topaz
RunqC "Update `#86Sku` x inner join Topaz a on x.CdTopaz=a.CdTopaz set x.Topaz=a.Topaz"
RunqC "Update `#87Sku` x inner join Topaz a on x.CdTopaz=a.CdTopaz set x.Topaz=a.Topaz"
End Sub
Private Sub B_TstDif()
W_LnkFxi
W_Tmp8687
Const A1$ = "Topaz"
Const A2$ = "ProdHierarchy"
Const A3$ = "SkuDes"
Const A4$ = "StkUnit"
Const A5$ = "Unit/AC"
Const A6$ = "Unit/SC"
Const A7$ = "Btl/AC"
Const A8$ = "Litre/BtlSap"
Const A9$ = "BusAreaSap"
Const A10$ = "SkuDes"

Dim Wh$: Wh = " where x.[" & A6 & "]<>a.[" & A6 & "]"

MsgBox _
ValQC("Select Count(*) from SKU x inner join `#86Sku` a on x.SKU=a.SKU " & Wh) & " " & _
ValQC("Select Count(*) from SKU x inner join `#87Sku` a on x.SKU=a.SKU " & Wh) & " "

BrwQC "Select x.SKu,x.[Unit/SC],a.[Unit/SC] from SKU x inner join `#87Sku` a on x.SKU=a.SKU where a.[Unit/SC]<>x.[Unit/SC]"
BrwQC "Select x.SKu,x.SkuDes,a.SkuDes from SKU x inner join `#87Sku` a on x.SKU=a.SKU where a.SkuDes<>x.SkuDes"
End Sub

Private Sub W_Upd(Tim As Date)
RunqC "Select * into [#Mge] from [#86Sku]"
RunqC "Insert into [#Mge] Select x.* from [#87Sku] x left join [#86Sku] a on a.Sku=x.Sku where a.Sku is null"
RunqC "Update SKU x inner join [#Mge] a on x.SKU=a.SKU set" & _
"x.Topaz=a.Topaz," & _
"x.ProdHierarchy=a.ProdHierarchy," & _
"x.SkuDes=a.SkuDes," & _
"x.StkUnit=a.StkUnit," & _
"x.[Unit/Ac]=a.[Unit/Ac]," & _
"x.[Unit/Sc]=a.[Unit/Sc]," & _
"x.[Btl/Ac]=a.[Btl/Ac]," & _
"x.[Litre/BtlSap]=a.[Litre/BtlSap]," & _
"x.BusAreaSap=a.BusAreaSap," & _
"x.DteRUpdTopaz=#" & Tim & "#" & _
" where x.Topaz     <>a.Topaz" & _
" or x.ProdHierarchy<>a.ProdHierarchy" & _
" or x.SkuDes     <>a.SkuDes" & _
" or x.StkUnit    <>a.StkUnit" & _
" or x.[Unit/AC]  <>a.[Unit/AC]" & _
" or x.[Unit/SC]  <>a.[Unit/SC]" & _
" or x.[Btl/AC]   <>a.[Btl/AC]" & _
" or x.[Litre/BtlSap]<>a.[Litre/BtlSap]" & _
" or x.BusAreaSap    <>a.BusAreaSap"
DrpC "#Mge"
End Sub

Private Sub W_InsSku(Tim As Date) ' Ins Tb-Sku from #Sku87 & #Sku86
Const InsSku$ = "Insert Into Sku (Sku,Topaz,SkuDes,ProdHierarchy,StkUnit,[Unit/SC],[Unit/AC],[Litre/BtlSap],[Btl/AC],BusAreaSap,DteCrt)"
Dim SelInp$: SelInp = "Select X.Sku, X.Topaz," & _
"Trim(Nz(X.SkuDes       ,'')) as SkuDes," & _
"Trim(Nz(X.ProdHierarchy,'')) as ProdHierarchy," & _
"Trim(Nz(X.StkUnit      ,'')) as StkUnit," & _
"Nz(X.[Unit/Sc]     ,0)  As [Unit/SC]," & _
"Nz(X.[Unit/Ac]     ,0)  As [Unit/AC]," & _
"Nz(X.[Litre/BtlSap],0)  As [Litre/BtlSap]," & _
"Nz(X.[Btl/Ac]      ,0)  As [Btl/AC]," & _
"Nz(X.[BusAreaSap]  ,'') As BusAreaSap," & _
"#" & Tim & "# As DteCrt"
RunqC InsSku & SelInp & " from `#86Sku` x left join SKU a on x.SKU=a.SKU where a.SKU is null"
RunqC InsSku & SelInp & " from `#87Sku` x left join SKU a on x.SKU=a.SKU where a.SKU is null"
End Sub

Private Sub B_W_PromptRslt(): W_PromptRslt Now: End Sub
Private Sub W_PromptRslt(Tim As Date)
Dim NNew%
Dim NChg%
    Const SqlChg$ = "Select Count(*) from Sku where DteRUpdTopaz=#?#"
    Const SqlNew$ = "Select Count(*) from Sku where DteCrt     =#?#"
    NNew = ValQC(FmtQQ(SqlNew, Tim))
    NChg = ValQC(FmtQQ(SqlChg, Tim))
Dim N86%: N86 = NRecTC(">Sku86")
Dim N87%: N87 = NRecTC(">Sku87")
Dim Tail$: If NChg <> 0 Or NNew <> 0 Then Tail = vbCrLf & "Check time Stamp [" & Tim & "]"
MsgBox "Done" & vbCrLf & vbCrLf & _
"There are [" & N86 & "] input records from 8601 worksheet." & vbCrLf & _
"There are [" & N87 & "] input records from 8701 worksheet" & vbCrLf & vbCrLf & "[" & NNew & "] Sku created" & vbCrLf & "There are [" & NChg & "] Sku changed." & Tail, vbInformation
Exit Sub
Sql:

End Sub
