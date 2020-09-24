Attribute VB_Name = "MxDao_Ado_Xny"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Ado_Xny."
Function TnyAdoC() As String():              TnyAdoC = TnyAdo(CDb):                                 End Function
Function TnyAdo(D As Database) As String():   TnyAdo = TnyFb(D.Name):                               End Function
Function TnyCat(A As Catalog) As String():    TnyCat = Itn(A.Tables):                               End Function
Function TnyFb(Fb) As String():                TnyFb = Tny(Db(Fb)):                                 End Function
Function TnyFbAdo(Fb) As String():          TnyFbAdo = AeLikk(TnyCat(CatFb(Fb)), "MSys* f_*_Data"): End Function
Function Wny(B As Workbook) As String():         Wny = Itn(B.Sheets):                               End Function
Function TnyFx(Fx) As String():                TnyFx = Itn(AxTdsFx(Fx)):                            End Function
Function WnyWb(B As Workbook) As String():     WnyWb = Itn(B.Sheets):                               End Function

Private Sub B_WnyFx()
Dim Fx$
GoSub Z
'GoSub T1
'GoSub T2
Exit Sub
Tst:
    Act = WnyFx(Fx)
    C
    Return
T1:
    Fx = MHO.MHOMB52.FxiSalTxt
    Ept = SySs("")
    GoTo Tst
T2:
    Fx = "C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
    Ept = SySs("")
    GoTo Tst
Z:
    DmpAy WnyFx(MHO.MHOMB52.FxiSalTxt)
    Return
End Sub
Private Function B_TnyCatFx()
'Note: TnyCatFx(MH.FcIO.Tp) will given 2 different list if MH.FcIO.Tp is openned or not openned
'      . the not openned will have more items.
'      . Openned (less)
'      .    'Fc Bus$'
'      .    'Fc L1$'
'      .    'Fc L2$'
'      .    'Fc L3$'
'      .    'Fc L4$'
'      .    'Fc Sku$'
'      .    'Fc Stm$'
'      . Not-Openned (more)
'      .    'Fc Bus$'
'      .    'Fc Bus$'FcBus
'      .    'Fc L1$'
'      .    'Fc L1$'_FcL1
'      .    'Fc L2$'
'      .    'Fc L2$'_FcL2
'      .    'Fc L3$'
'      .    'Fc L3$'_FcL3
'      .    'Fc L4$'
'      .    'Fc L4$'_FcL4
'      .    'Fc Sku$'
'      .    'Fc Sku$'FcSku
'      .    'Fc Stm$'
'      .    'Fc Stm$'FcStm

Dim F$: F = MH.FcTp.Tp
MaxvFx F
Dim A$(): A = TnyCatFx(F)
D A
Stop
End Function
Function TnyCatFx(Fx) As String(): TnyCatFx = TnyCat(CatFx(Fx)): End Function

Function WsnFst$(Fx)
Dim T: For Each T In Itr(TnyCat(CatFx(Fx)))
    WsnFst = WsnAxtn(T): Exit Function
Next
End Function
Function WnyFx(Fx) As String()
Dim T: For Each T In Itr(TnyCat(CatFx(Fx)))
    PushNB WnyFx, WsnAxtn(T)
Next
End Function
Private Function WsnAxtn$(Axtn) 'Axtn:Cml :Tbn ! #Cat-Tbl-Nm#
If HasSfx(Axtn, "_xlnm#_FilterDatabase") Then Exit Function
WsnAxtn = RmvSfx(RmvQuoSng(Axtn), "$")
End Function

Function FfFxw$(Fx, Optional W$):          Stop '          FfFxw = TmlAy(FnyFxw(Fx, W)): End Function
End Function
Function FnyArs(A As ADODB.Recordset) As String():  FnyArs = Itn(A.Fields):  End Function
Function FnyAxTd(T As ADOX.Table) As String():     FnyAxTd = Itn(T.Columns): End Function
Function FnyFbt(Fb, T) As String():                 FnyFbt = Fny(Db(Fb), T): End Function
Function FnyFbtAdo(Fb, T) As String()
Dim C As ADOX.Catalog
Set C = CatFb(Fb)
FnyFbtAdo = FnyAxTd(C.Tables(T))
End Function
