Attribute VB_Name = "MxDao_Ado"
Option Compare Text
Option Explicit
#If Doc Then
'Ax:Cml #AdoX#
'Td:Cml #Table-Definition#
#End If
Const CMod$ = "MxDao_Ado."
Function CvCn(A) As ADODB.Connection: Set CvCn = A: End Function

Function AxTd(C As Catalog, T) As ADOX.Table  'Ado.Table definition
Set AxTd = C.Tables(T)
End Function

Function Axtn$(Wsn)  ' #Adox-Table-Name# format for Wsn: When Wsn IsNm, just add Sfx-$, otherwise SngQuo(Add Sfx-$)
Axtn = IIf( _
    IsNm(Wsn), _
        Wsn & "$", _
        QuoSng(Wsn & "$"))
End Function

Function AxTdsFx(Fx) As ADOX.Tables
Dim C As Catalog: Set C = CatFx(Fx) ' it is must.  CatFx(Fx).Tables does not work. Because in CatFx(Fx) will be disposed before it can return .Table
Set AxTdsFx = C.Tables
End Function

Function NRecArs&(R As ADODB.Recordset)
If NoRecArs(R) Then Exit Function
R.MoveLast
NRecArs = R.RecordCount
R.MoveFirst
End Function
Function CnsFxOle$(Fx): CnsFxOle = "OLEDb;" & CnsFxAdo(Fx): End Function
Function CnsFbOle$(Fb): CnsFbOle = "OLEDb;" & CnsFbAdo(Fb): End Function


Function ArsFxq(Fx$, Q$) As ADODB.Recordset:                Set ArsFxq = CnFx(Fx).Execute(Q):             End Function
Function ArsFxw(Fx$, W, Optional Bepr$) As ADODB.Recordset: Set ArsFxw = ArsFxq(Fx, SqlSelStar(Axtn(W))): End Function

Sub RunqFbqCn(Fb, Q): CnFb(Fb).Execute Q: End Sub

Function CatFb(Fb) As Catalog: Set CatFb = X_1Cat(CnFb(Fb)): End Function
Function CatFx(Fx) As Catalog: Set CatFx = X_1Cat(CnFx(Fx)): End Function
Private Function X_1Cat(C As ADODB.Connection) As Catalog
Set X_1Cat = New Catalog
Set X_1Cat.ActiveConnection = C
End Function

Function CnC() As ADODB.Connection: Set CnC = Acs.CurrentProject.Connection: End Function
Function CnIf(OCn As ADODB.Connection, Fb) As ADODB.Connection
If IsNothing(OCn) Then Set OCn = CnFb(Fb)
Set CnIf = OCn
End Function
Function CnFb(Fb) As ADODB.Connection: Set CnFb = X_2Cn(CnsFbAdo(FfnChkExist(Fb))): End Function
Function CnFx(Fx) As ADODB.Connection: Set CnFx = X_2Cn(CnsFxAdo(FfnChkExist(Fx))): End Function
Function CnsFbAdo$(A)
'0000000000000000000000000000000
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;
'User ID=Admin;
'Data Source=?;
'Mode=Share Deny None;
'Jet OLEDB:Engine Type=6;
'Jet OLEDB:Database Locking Mode=0;
'Jet OLEDB:Global Partial Bulk Ops=2;
'Jet OLEDB:Global Bulk Transactions=1;
'Jet OLEDB:Create System Database=False;
'Jet OLEDB:Encrypt Database=False;
'Jet OLEDB:Don't Copy Locale on Compact=False;
'Jet OLEDB:Compact Without Replica Repair=False;
'Jet OLEDB:SFP=False;
'Jet OLEDB:Support Complex Data=False;
'Jet OLEDB:Bypass UserINF Validation=False;
'Jet OLEDB:Limited DB Caching=False;
'Jet OLEDB:Bypass ChoiceField Validation=False"
'
'Locking Mode=1 means page (or record level) according to https://www.spreadsheet1.com/how-to-refresh-pivottables-without-locking-the-source-workbook.html
'The ADO connection object initialization property which controls how the database is locked, while records are being read or modified is: Jet OLEDB:Database Locking Mode
'Please note:
'The first user to open the database determines the locking mode to be used while the database remains open.
'A database can only be opened is a single mode at a time.
'For Page-level locking, set property to 0
'For Row-level locking, set property to 1
'With 'Jet OLEDB:Database Locking Mode = 0', the source spreadshseet is locked, while PivotTables update. If the property is set to 1,
'the source file is not locked. Only individual records (Table rows) are locked sequentially, while data is being read.
CnsFbAdo = FmtQQ(X_4Tp, A)

End Function
Private Function X_3Tp$()
'Provider=Microsoft.ACE.OLEDB.16.0;
'User ID=Admin;
'Data Source=C:\Users\Public\Logistic\StockHolding8\StockHolding8.accdb;
'Mode=Share Deny None;
'Extended Properties="";
'Jet OLEDB:System database="";
'Jet OLEDB:Registry Path="";
'Jet OLEDB:Engine Type=6;
'Jet OLEDB:Database Locking Mode=1;
'Jet OLEDB:Global Partial Bulk Ops=2;
'Jet OLEDB:Global Bulk Transactions=1;
'Jet OLEDB:New Database Password="";
'Jet OLEDB:Create System Database=False;
'Jet OLEDB:Encrypt Database=False;
'Jet OLEDB:Don't Copy Locale on Compact=False;
'Jet OLEDB:Compact Without Replica Repair=False;
'Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;
'Jet OLEDB:Bypass UserInfo Validation=False;
'Jet OLEDB:Limited DB Caching=False;
'Jet OLEDB:Bypass ChoiceField Validation=False
End Function
Private Function X_4Tp$()
X_4Tp = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;User ID=Admin;Jet OLEDB:Database Locking Mode=0" 'Mode=Share Deny None;
End Function
Private Function X_6TpFmOkRfh$()
'Provider=Microsoft.ACE.OLEDB.16.0;
'User ID=Admin;
'Data Source=C:\Users\user\Documents\Projects\Vba\QLib\QLib_StockHolding8.accdb;
'Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=1;
'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;
'Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;
'Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False

End Function

Function CnsFxAdo$(A)
'CnsFxAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?", A) 'Try
CnsFxAdo = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""", A) 'Ok
End Function

Function CvAdoTy(A) As ADODB.DataTypeEnum:   CvAdoTy = A: End Function
Function CvAxt(A) As ADOX.Table:           Set CvAxt = A: End Function ' #Axt:Adox.Table#

Function DftTny(Tny0, Fb) As String()
If IsMissing(Tny0) Then
    DftTny = TnyFb(Fb)
Else
    DftTny = CvSy(Tny0)
End If
End Function

Function Scv(Scl$, Nm$):            Scv = BetSPr(EnsSfx(Scl, ";"), Nm & "=", ";"): End Function
Function ScvDtaSrc(Scl$):     ScvDtaSrc = Scv(Scl, "Data Source"):                 End Function '#Scv:Semi-Colon-Value#
Function ScvDatabase(Scl$): ScvDatabase = Scv(Scl, "Database"):                    End Function

Function FnyFxw(Fx, Optional W$ = "") As String() ' ret Fny of Fx->W.  If W is blnk, fst W.
Dim C As Catalog: Set C = CatFx(Fx)
Dim Wsn$: Wsn = WsnDft(W, Fx)
Dim T As ADOX.Table: Set T = C.Tables(Axtn(Wsn))
FnyFxw = FnyAxTd(T)
End Function

Function NoFbt(Fb, T) As Boolean:     NoFbt = Not HasFbt(Fb, T):      End Function
Function NoFxw(Fx, Wsn) As Boolean:   NoFxw = Not HasFxw(Fx, Wsn):    End Function
Function HasFbt(Fb, T) As Boolean:   HasFbt = HasEle(TnyFb(Fb), T):   End Function
Function HasFxw(Fx, Wsn) As Boolean: HasFxw = HasEle(WnyFx(Fx), Wsn): End Function


Sub RunqCnSqy(Cn As ADODB.Connection, Sqy$())
Dim Q: For Each Q In Itr(Sqy)
   Cn.Execute Q
Next
End Sub

Private Sub B_RunqFbqCn()
Dim Fb$: Fb = MHO.MHODuty.FbPgm
Const Q$ = "Select * into [#a] from Permit"
DrpFbt Fb, "#a"
RunqFbqCn Fb, Q
End Sub

Private Sub B_CnFb()
Dim Cn As ADODB.Connection
Set Cn = CnFb(MHO.MHODuty.FbDta)
Stop
End Sub

Private Sub B_DrsCnq()
Dim Cn As ADODB.Connection: Set Cn = CnFx(MHO.MHOMB52.FxiSalTxt)
Dim Q$: Q = "Select * from [Sheet1$]"
BrwDrs DrsCnq(Cn, Q)
End Sub

Private Sub B_DrsFbqAdo()
GoSub ZZ
Exit Sub
Dim Fb$, Q$
ZZ:
    Fb = MHO.MHODuty.FbDta
    Q = "Select * from Permit"
    GoTo Tst
Tst:
    BrwDrs DrsFbqAdo(Fb, Q)
    Return
End Sub

Function HasRecArs(A As ADODB.Recordset) As Boolean: HasRecArs = Not NoRecArs(A): End Function
Function NoRecArs(A As ADODB.Recordset) As Boolean:   NoRecArs = A.EOF And A.BOF: End Function

Private Function X_5IntoArs(Into, A As ADODB.Recordset, F)
Dim O: O = AyNw(Into)
With A
    While Not .EOF
        PushI O, Nz(A(F))
        .MoveNext
    Wend
End With
End Function

Private Sub B_X_2Cn()
Dim O As ADODB.Connection
Set O = X_2Cn(Cns_ADO_SampSQL_EXPR_NOT_WRK)
Stop
End Sub
Private Function X_2Cn(CnsAdo, Optional IsRO As Boolean) As ADODB.Connection
Set X_2Cn = New ADODB.Connection
If IsRO Then X_2Cn.Mode = adModeShareExclusive
X_2Cn.Open CnsAdo
End Function
