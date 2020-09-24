Attribute VB_Name = "MxDao_Db_Fb"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Fb."


Sub AsgFbtStr(FbtStr$, OFb$, OT$)
If FbtStr = "" Then
    OFb = ""
    OT = ""
    Exit Sub
End If
AsgBrk OFb, OT, _
    FbtStr, "].["
If ChrFst(OFb) <> "[" Then Stop
If ChrLas(OT) <> "]" Then Stop
OFb = RmvFst(OFb)
OT = RmvLas(OT)
End Sub

Sub EnsFb(Fb)
If NoFfn(Fb) Then CrtFb Fb
End Sub
Function CrtFb(Fb) As Database:             Set CrtFb = Dao.DBEngine.CreateDatabase(Fb, dbLangGeneral): End Function
Function ArsFbq(Fb, Q) As ADODB.Recordset: Set ArsFbq = CnFb(Fb).Execute(Q):                            End Function
Sub ArunFbq(Fb, Q):                                     CnFb(Fb).Execute Q:                             End Sub
Function DbFb(Fb) As Database:               Set DbFb = Db(Fb):                                         End Function
Function DbRfhTd(D As Database) As Database: D.TableDefs.Refresh: Set DbRfhTd = D:                                                     End Function
Sub DrpFbt(Fb, T):                                   CatFb(Fb).Tables.Delete T: End Sub
Function DrsFbq(Fb, Q) As Drs:              DrsFbq = DrsRs(Rs(Db(Fb), Q)):      End Function
Function DrsSql(D As Database, Q) As Drs:   DrsSql = DrsRs(Rs(D, Q)):           End Function
Private Sub B_TnyOupFb():                            D TnyOupFb(CFb):           End Sub
Function TnyOupFb(Fb) As String():        TnyOupFb = TnyOup(Db(Fb)):            End Function
Function WsFbq(Fb, Q, Optional Wsn$) As Worksheet: Set WsFbq = WsDrs(DrsFbq(Fb, Q), Wsn:=Wsn):                                  End Function
Private Sub B_BrwFb():  BrwFb MHO.MHODuty.FbDta:                                         End Sub
Private Sub B_HasFbt(): Ass HasFbt(MHO.MHODuty.FbDta, "SkuB"):                           End Sub
Private Sub B_TnyFb():  DmpAy TnyFb(MHO.MHODuty.FbDta):                                  End Sub
Private Sub B_WsFbq():  Maxv WsFbq(MHO.MHODuty.FbDta, "Select * from KE24").Application: End Sub
