VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_MB52Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_MB52Load."
Private MFxi$, MYmd As Ymd, MWsnFxi$, MCslFldTy$, MCslFldn$, MIon As MxMHMB52_A_MB52Io, MWhOHYmd$
Const W2a$ = "Material,Plant,Storage Location,Batch,Base Unit of Measure"
Const W3a$ = "T       ,T    ,T               ,T    ,T                   "
Const W2b$ = ",Unrestricted,Transit and Transfer,In Quality Insp#,Blocked,Value Unrestricted,Val# in Trans#/Tfr,Value in QualInsp#,Value BlockedStock,Value Rets Blocked"
Const W3b$ = ",N           ,N                   ,N               ,N      ,N                 ,N                 ,N                 ,N                 ,N"
Property Get CslFldn$():  CslFldn$ = W2a & W2b:                 End Property
Property Get CslFldTy$(): CslFldTy = W3a & W3b:                 End Property
Sub B_LoadMB52():         LoadMB52 MH.TbOH.YmdLas, NoAsk:=True: End Sub
Sub LoadMB52(Ymd As Ymd, Optional NoAsk As Boolean)
Set MIon = MH.MB52IO
MFxi = MIon.Fxi(Ymd)
MWsnFxi = MIon.WsnFxi
MYmd = Ymd
MWsnFxi = MIon.WsnFxi
MCslFldTy = MH.MB52Load.CslFldTy
MCslFldn = MH.MB52Load.CslFldn
MWhOHYmd = MH.TbOH.WhYmd(Ymd)
ChkFfnExi MFxi, CSub
If Not IsCfm("Start Load MB52?", , NoAsk) Then Exit Sub
Sts "Start loading MB52: " & HypymdYmd(Ymd)
WChkFxi__CI
WLnkFxi
WCrtHIMB52_FmIMB52
WCrtHOH_FmHIMB52
End Sub
Private Sub WCrtHOH_FmHIMB52()
DoCmd.SetWarnings False

RunqC "SELECT Co, SKU, BchNo, Sum(x.Q) AS Q, Sum(x.V) AS V, CByte(0) as YpStk, SLoc" & _
" INTO [#OH]" & _
" FROM [#IMB52] x" & _
" GROUP BY Co,SKU,BchNo,SLoc;"
RunqC "DELETE * FROM [#OH] WHERE Nz(Q,0)=0 AND Nz(V,0)=0;"
RunqC "Update [#OH] x inner join YpStk a on a.Co=x.Co and a.SLoc=x.SLoc set x.YpStk=a.YpStk"
RunqC "Drop Table `#IMB52`"
'---------------------------------------------------------------------------
'Update Q -> Btl  (Q is in StkUnit, which may PCE or COL  (PCE is set.  Required to convert to Btl)
RunqC "Alter Table [#OH] add column Btl Long,[Unit/AC] double, [Btl/AC] integer, [Unit/SC] double"
RunqC "Update [#OH] x inner join [qSku_Main] a on x.Sku=a.Sku set " & _
"x.[Unit/SC]=a.[Unit/SC]," & _
"x.[Btl/AC]=a.[Btl/AC]," & _
"x.[Unit/Ac]=a.[Unit/Ac]"
RunqC "Update [#OH] set Btl = Q / [Unit/AC] * [Btl/AC]"
'---------------------------------------------------------------------------
'Upd: OH
'Fm : #OH
RunqC "DELETE FROM OH" & MWhOHYmd
With MYmd
RunqC FmtStr("INSERT INTO OH (YY, MM, DD, Co, SLoc, SKU, YpStk, BchNo, Btl, Val )" & _
" SELECT {0}, {1}, {2}, Co, SLoc, SKU, YpStk, BchNo, Btl, V FROM [#OH];", .Y, .M, .D)
End With
'---------------------------------------------------------------------------
'Upd: Report->(NRecMB52 TotBtlMB52 TotHKDMB52)
'Fm : #OH                                  ' Drop after use
RunqC "SELECT Count(*) AS NRec, Sum(Q) AS TotBtl, Sum(V) AS TotVal INTO [#Tot] FROM [#OH]"
RunqC "UPDATE Report x, [#Tot] a SET DteMB52=Now(), x.NRecMB52=a.NRec, x.TotBtlMB52=a.TotBtl, x.TotHKDMB52=a.TotVal" & MWhOHYmd
RunqC "Drop Table `#Tot`"

'Upd: Report->(MB52SC MB52AC)
RunqC "Select Distinct Sku,Sum(Q) as Btl into [#A] from [#OH] Group by Sku"
RunqC "Alter Table [#A] add column [Unit/SC] double, [Unit/AC] double, [Btl/AC] double, SC double,AC double"
RunqC "Update [#A] x inner join qSku_Main a on a.Sku=x.Sku set x.[Unit/SC]=a.[Unit/SC],x.[Unit/AC] = a.[Unit/AC], x.[Btl/AC]=a.[Btl/AC]"
RunqC "Update [#A] set AC=Btl/[Btl/AC], SC==Btl/[Btl/AC] * [Unit/AC] / [Unit/SC]"
RunqC "Select Sum(x.SC) as SC, Sum(x.AC) as AC into [#B] from [#A] x"
'DoCmd.RunSQL "Update Report,[#B] set MB52SC=SC,MB52AC=AC," & _
'"GitSC=Null,GitAC=Null,GitHKD=Null,GitBtl=Null,GitNRec=Null,GitLoadDte=Null" & Wh
RunqC "Update Report,[#B] set MB52SC=SC,MB52AC=AC" & MWhOHYmd

'-- Drp temp table
RunqC "Drop Table [#A]"
RunqC "Drop Table [#B]"
RunqC "Drop Table [#OH]"
Done
End Sub
Private Sub WCrtHIMB52_FmIMB52()
Sts "Importing....."
RunqC "SELECT Material AS SKU, Batch as BchNo, CByte(Left(Plant,2)) as Co, [Storage Location] as SLoc," & _
" CDbl(Nz(Unrestricted,0)+ Nz([Transit and Transfer],0)+Nz([In Quality Insp#],0)+Nz(Blocked,0)) As Q," & _
" CCur(Nz([Value Unrestricted],0)+Nz([Val# in Trans#/Tfr],0)+Nz([Value in QualInsp#],0)+Nz([Value BlockedStock],0)) As V" & _
" INTO [#IMB52]" & _
" FROM [>MB52];"
End Sub
Private Sub WChkFxi__CI()
Dim Fx$, W$: Fx = MFxi: W = MWsnFxi
Sts "Validating...":
ChkFxwCol Fx, W, MCslFldn, MCslFldTy
Dim M1$(), M2$()
    M1 = WarnFxwcShdNB(Fx, W, "Material")
    M2 = EryFxwcShdAllInAp(Fx, W, "Plant", "8701", "8601")
Dim M$(): M = AyItmAy(M1, M2)
ChkEry M, CSub, "There is errors in Excel file"
End Sub
Private Sub WLnkFxi(): LnkFxwC ">MB52", MFxi, MWsnFxi: End Sub
