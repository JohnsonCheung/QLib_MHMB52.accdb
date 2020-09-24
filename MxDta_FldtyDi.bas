Attribute VB_Name = "MxDta_FldtyDi"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_FldtyDi."
Function DiFqShtDao(FxOrFb$, TorW) As Dictionary
Const CSub$ = CMod & "DiFqShtDao"
Select Case True
Case IsFb(FxOrFb): Set DiFqShtDao = DiFqShtDaotyFbt(FxOrFb, TorW)
Case IsFx(FxOrFb): Set DiFqShtDao = DiFqShtAdoFxw(FxOrFb, TorW)
Case Else: Thw CSub, "FxOrFb should be Fx or Fb", "FxOrFb TorW", FxOrFb, TorW
End Select
End Function

Function DiFqShtDaotyFbt(Fb, T) As Dictionary
Set DiFqShtDaotyFbt = New Dictionary
Dim D As Database: Set D = Db(Fb)
Dim Td As Dao.TableDef: Set Td = D.TableDefs(T)
Dim F As Dao.Field: For Each F In Td.Fields
    DiFqShtDaotyFbt.Add F.Name, ShtDaoty(F.Type)
Next
End Function

Function DiFqShtAdoFxw(Fx, Optional W = "Sheet1") As Dictionary
Set DiFqShtAdoFxw = New Dictionary
Dim Cat As Catalog: Set Cat = CatFx(Fx)
Dim C As ADOX.Column: For Each C In Cat.Tables(Axtn(W)).Columns
    DiFqShtAdoFxw.Add C.Name, ShtAdoTy(C.Type)
Next
End Function

Private Sub B_DiFqShtAdoFxw()
Stop 'BrwDi DiFqShtAdoFxw(MHSalTxtFx)
End Sub
