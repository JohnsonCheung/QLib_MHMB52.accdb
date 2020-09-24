Attribute VB_Name = "MxMHMB52__Globals"
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52__Globals."
Public MH As New MxMHMB52_A__MH
Public Const FrmnRpt$ = "Rpt"
Public Const FrmnLoadSku$ = "Rpt"
Public Const FrmnLoadFc$ = "Rpt"
Public Const FrmnLoadZHT0$ = "Rpt"
Public Const FrmnLoadPH$ = "LoadPH"
Public Const FrmnMacauCalcRate$ = "MacauCalcRate"
Type StmYm: Stm As String: Y As Byte: M As Byte: End Type
Type CoYm: Co As Byte: Y As Byte: M As Byte: End Type
Type CoYmd: Co As Byte: Ymd As Ymd: End Type
Type CoStmYm: Co As Byte: Stm As String: Y As Byte: M As Byte: End Type
Function UbStmYm&(A() As StmYm): UbStmYm = SiStmYm(A) - 1: End Function
Function SiStmYm&(A() As StmYm): On Error Resume Next: SiStmYm = UBound(A): End Function
Sub PushStmYm(O() As StmYm, M As StmYm): Dim N&: N = SiStmYm(O): ReDim O(N): O(N) = M: End Sub


Property Get SpRseqSku$(): SpRseqSku = "Sku *Atr *OH *Unit *Tax *Dta *Ovr" & _
"|*Atr  Topaz SkuDes ProdHierarchy StkUnit BusArea" & _
"|*OH   WithOHHst WithOHCur" & _
"|*Unit Litre/Btl Btl/Ac Unit/Ac Unit/Sc" & _
"|*Tax  TaxRateHK TaxUOMHK TaxRateMO TaxUOMMO" & _
"|*Dta  DteCrt DteRUpdTopaz DteRUpdTaxRate  " & _
"|*Ovr  BusAreaOvr Litre/BtlOvr Litre/BtlSap BusAreaSap": End Property
Sub OpnTbSkuRepackMulti():   DoCmd.OpenTable "SkuRepackMulti":   End Sub
Sub OpnTbSkuTaxBy3rdParty(): DoCmd.OpenTable "SkuTaxBy3rdParty": End Sub
Sub OpnTbSkuNoLongerTax():   DoCmd.OpenTable "SkuNoLongerTax":   End Sub
Sub OpnTbYpStk():            DoCmd.OpenTable "YpStk":            End Sub
Sub OpnTbBusArea():          DoCmd.OpenTable "PHLBus":           End Sub


Sub OpnTbTarMthsL1():  VVOpnL1:  DoCmd.OpenTable "PHTarMthsL1":  End Sub
Sub OpnTbTarMthsL2():  VVOpnL2:  DoCmd.OpenTable "PHTarMthsL2":  End Sub
Sub OpnTbTarMthsL3():  VVOpnL3:  DoCmd.OpenTable "PHTarMthsL3":  End Sub
Sub OpnTbTarMthsL4():  VVOpnL4:  DoCmd.OpenTable "PHTarMthsL4":  End Sub
Sub OpnTbTarMthsSku(): VVOpnSku: DoCmd.OpenTable "PHTarMthsSku": End Sub
Sub OpnTbTarMthsBus(): VVOpnBus: DoCmd.OpenTable "PHTarMthsBus": End Sub
Sub OpnTbTarMthsStm(): DoCmd.OpenTable "PHTarMthsStm":           End Sub
Private Sub VVOpnSku()
DoCmd.SetWarnings False
RunqC "SELECT Distinct Co,Sku into [#A] from OH"
RunqC "Insert into PHTarMthsSku (Co,Sku) select x.Co,x.Sku" & _
" from [#A] x" & _
" left join [PHTarMthsSku] a on a.Co=x.Co and a.Sku=x.Sku" & _
" where a.Sku is null"
RunqC "Drop table [#A]"
End Sub
Private Sub VVOpnBus()
DoCmd.SetWarnings False
RunqC "Select Distinct Co into [#Co] from CoStm"
RunqC "SELECT Co,Stm,BusArea into [#A] from [#Co],PHLBus"
RunqC "Insert into PHTarMthsBus (Co,Stm,BusArea) select x.Co,x.Stm,x.BusArea" & _
" from [#A] x" & _
" left join [PHTarMthsBus] a on a.Co=x.Co and a.Stm=x.Stm and a.BusArea=x.BusArea" & _
" where a.Co is null"
DrpTTC "#Co #A"
End Sub
Private Sub VVOpnL4()
DoCmd.SetWarnings False
RunqC "Select Co,x.Sku,ProdHierarchy,Left(ProdHierarchy,10) as PHL4 into [#A] from PHTarMthsSku x inner join Sku a on x.Sku=a.Sku"
RunqC "Select Distinct Stm into [#Stm] from CoStm group by Stm"
RunqC "Select Distinct Co,Stm,PHL4 into [#B] from [#A],[#Stm] group by Co,Stm,PHL4"
RunqC "Insert Into PHTarMthsL4 (Co,Stm,PHL4) select x.Co,x.Stm,x.PHL4 from [#B] x" & _
" left join [PHTarMthsL4] a on a.Co=x.Co and a.Stm=x.Stm and a.PHL4=x.PHL4" & _
" where a.Co is null"
'DrpTzTT "#A #B #Stm"
End Sub
Private Sub VVOpnL3()
DoCmd.SetWarnings False
RunqC "Select Co,x.Sku,ProdHierarchy,Left(ProdHierarchy,7) as PHL3 into [#A] from PHTarMthsSku x inner join Sku a on x.Sku=a.Sku"
RunqC "Select Distinct Stm into [#Stm] from CoStm group by Stm"
RunqC "Select Distinct Co,Stm,PHL3 into [#B] from [#A],[#Stm] group by Co,Stm,PHL3"
RunqC "Insert Into PHTarMthsL3 (Co,Stm,PHL3) select x.Co,x.Stm,x.PHL3 from [#B] x" & _
" left join [PHTarMthsL3] a on a.Co=x.Co and a.Stm=x.Stm and a.PHL3=x.PHL3" & _
" where a.Co is null"
End Sub
Private Sub VVOpnL2()
DoCmd.SetWarnings False
RunqC "Select Co,x.Sku,ProdHierarchy,Left(ProdHierarchy,4) as PHL2 into [#A] from PHTarMthsSku x inner join Sku a on x.Sku=a.Sku"
RunqC "Select Distinct Stm into [#Stm] from CoStm group by Stm"
RunqC "Select Distinct Co,Stm,PHL2 into [#B] from [#A],[#Stm] group by Co,Stm,PHL2"
RunqC "Insert Into PHTarMthsL2 (Co,Stm,PHL2) select x.Co,x.Stm,x.PHL2 from [#B] x" & _
" left join [PHTarMthsL2] a on a.Co=x.Co and a.Stm=x.Stm and a.PHL2=x.PHL2" & _
" where a.Co is null"
End Sub
Private Sub VVOpnL1()
DoCmd.SetWarnings False
RunqC "Select Co,x.Sku,ProdHierarchy,Left(ProdHierarchy,2) as PHL1 into [#A] from PHTarMthsSku x inner join Sku a on x.Sku=a.Sku"
RunqC "Select Distinct Stm into [#Stm] from CoStm group by Stm"
RunqC "Select Distinct Co,Stm,PHL1 into [#B] from [#A],[#Stm] group by Co,Stm,PHL1"
RunqC "Insert Into PHTarMthsL1 (Co,Stm,PHL1) select x.Co,x.Stm,x.PHL1 from [#B] x" & _
" left join [PHTarMthsL1] a on a.Co=x.Co and a.Stm=x.Stm and a.PHL1=x.PHL1" & _
" where a.Co is null"
End Sub
