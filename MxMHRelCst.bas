Attribute VB_Name = "MxMHRelCst"
Option Compare Text
Option Explicit
Const CMod$ = "MxMHRelCst."
Sub UomDoc()
#If False Then
InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
Oup : UOM        Sku      SkuUOM                 Des                    Sc_U

Note on [Sales text.xls]
DcDrs  Xls Title            FldName     Means
F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
J    Unit per case        Sc_U        how many unit per AC
K    SC                   SC_U        how many unit per SC   ('no need)
L    COL per case         AC_B        how many bottle per AC
-----
Letter meaning
B = Bottle
AC = act case
SC = standard case
U = Unit  (Bottle(COL) or Set (PCE))

 "SC              as SC_U," & _  no need
 "[COL per case]  as AC_B," & _ no need
#End If
End Sub

Function EryPlnt8687Mis(Fxi$, Wsn$) As String()
If NRecFxw(Fxi, Wsn, "Plant in ('8601','8701')") = 0 Then
    EryPlnt8687Mis = WMsg(Fxi, Wsn)
End If
End Function
Private Function WMsg(Fxi$, Wsn$) As String()
Const M$ = "Column-[Plant] must have value 8601 or 8701"
WMsg = MsgyFMNap("EryPlnt8687Mis", M, "MB52-File Worksheet", Fxi, Wsn)
End Function

Private Sub VVOupRat(D As Database)
WOupRate D
WOupMain D
End Sub
Private Function WOupMain$(D As Database)
'#IUom
'#IMB52
'@IMB52 :Drs-Whs-Sku-QUnRes-QBlk-QInsp
'@IUom  :Sku-Sc_U-Des-StkUom
'Ret      : @@
Drp D, "@Main"

'== Crt @Main fm #IMB52
'   Whs Sku OH Des StkUom Sc_U OH
Runq D, "Select Distinct Whs,Sku,Sum(QUnRes+QBlk+QInsp) As OH into [@Main] from [#IMB52] Group by Whs,Sku"
Runq D, "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10),Sc_U Int, OH_Sc Double"
Runq D, "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"
Runq D, "Update [@Main] set OH_Sc=OH/Sc_U where Sc_U>0"

'== Add DcDrs Stream ProdH F2 M32 M35 M37 Topaz ZHT1 RateSc Z2 Z5 Z7
'   Upd DcDrs ProdH Topaz
'   Upd DcDrs F2 M32 M35 M37
Runq D, "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH text(7), F2 Text(2), M32 text(2), M35 text(5), M37 text(7), ZHT1 Text(7), Z2 text(2), Z5 text(5), Z7 text(7), RateSc Currency, Amt Currency"
Runq D, "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"
Runq D, "Update [@Main] set F2=Left(ProdH,2),M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M37=Mid(ProdH,3,7)"

'== Upd DcDrs ZHT1 RateSc
Runq D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M37=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
Runq D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
Runq D, "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'Stream
'Z2 Z5 Z7
'Amt
Runq D, "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"
Runq D, "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z7=Left(ZHT1,7) where not ZHT1 is null"
Runq D, "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
End Function
Private Function WOupRate$(D As Database)
'VdtFm & VdtTo format DD.MM.YYYY
'1: #IZHT18701 VdtFm VdtTo L3 RateSc
'1: #IZHT18601 VdtFm VdtTo L3 RateSc
'2: #IUom     SKu Sc_U
'O: @Rate  ZHT1 RateSc
DrpTT D, "#Cpy1 #Cpy2 #Cpy @Rate"
Runq D, "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
Runq D, "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

Runq D, "Select * into [#Cpy] from [#Cpy1] where False"
Runq D, "Insert into [#Cpy] select * from [#Cpy1]"
Runq D, "Insert into [#Cpy] select * from [#Cpy2]"

Runq D, "Alter Table [#Cpy] Add Column VdtFmDte Date,VdtToDte Date,IsCur YesNo"
Runq D, "Update [#Cpy] Set" & _
" VdtFmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" VdtToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
Runq D, "Update [#Cpy] set IsCur = true where Now between VdtFmDte and VdtToDte"

Runq D, "Select Whs,ZHT1,RateSc into [@Rate] from [#Cpy]"
DrpTT D, "#Cpy #Cpy1 #Cpy2"
End Function
