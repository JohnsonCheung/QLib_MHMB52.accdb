VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_PhLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_PhLoad."
Sub LoadPH()
Const Wsn$ = "Sheet1"
  Const CslFldn = "Product hierarchy,Level no#,Description"
Const CslFldnTy = "T                ,T        ,T"
Dim Fxi$: Fxi = MH.PHIo.Fxi
ChkFxwCol Fxi, "Sheet1", CslFldn, CslFldnTy
LnkFxwC Fxi, Wsn, ">PH"
DoCmd.SetWarnings False
'-------------------
Sts "Running import query ....."
'Crt #PH = PH Lvl Des
RunqC "Select [Product hierarchy] as PH,[Level no#] as Lvl,Description as Des into [#PH] from [>PH]"

'Dlt #PH ! For
WDltTmpPHRec
    
'Ins ProdHierarchy
'Upd ProdHierarchy
'Drp #PH
DoCmd.SetWarnings True
RunqC "Insert into ProdHierarchy Select x.PH,CByte(Val(x.Lvl)) as Lvl,x.Des from [#PH] x left Join ProdHierarchy a on x.PH=a.PH where a.PH is null"
RunqC "Update ProdHierarchy x inner join [#PH] a on x.PH=a.PH set x.Des=a.Des, x.Lvl=a.Lvl, DteRUpd=Now where x.Des<>a.Des or x.Lvl<>CByte(Val(a.Lvl))"

'Upd ProdHierarchy->WithOH
'Upd ProdHierarchy->Sno
'Upd ProdHierarchy->Srt
MH.TbPh.RfhFldWithOHXxx
MH.TbPh.RfhFldSno
MH.TbPh.RfhFldSrt
'Drp
DrpTTC ">PH #PH"
End Sub

Private Sub WDltTmpPHRec()
Dim N%

N = ValQC("Select Count(*) from [#PH] where Trim(Nz(PH,''))=''")
If N > 0 Then
    MsgBox "There are [" & N & "] records with [Product Hierarchy] is blank.  They are omitted!", vbInformation
    RunqC "Delete * from [#PH] where Trim(Nz(PH,''))='')"
End If
    
N = ValQC("Select Count(*) from [#PH] where Len(Nz(PH,''))>10")
If N > 0 Then
    MsgBox "There are [" & N & "] records with [Product Hierarchy] length is >10.  They are omitted!", vbInformation
    RunqC "Delete * from [#PH] where Len(Nz(PH,''))>10"
End If
    
N = ValQC("Select Count(*) from [#PH] where Not (CLng(Val(Lvl)) between 1 and 5)")
If N > 0 Then
    MsgBox "There are [" & N & "] records with Level# not between 1 and 5.  They are omitted!", vbInformation
    RunqC "Delete * from [#PH] where Not (CLng(Val(Lvl)) between 1 and 5)"
End If
    
N = CurrentDb.OpenRecordset("Select Count(*) from [#PH] where trim(Nz(Des,''))=''").Fields(0).Value
If N > 0 Then
    MsgBox "There are [" & N & "] records with Blank Description.  They are omitted!", vbInformation
    RunqC "Delete * from [#PH] where trim(Nz(Des,'')=''"
End If

End Sub
