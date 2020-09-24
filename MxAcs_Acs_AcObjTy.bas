Attribute VB_Name = "MxAcs_Acs_AcObjTy"
Option Compare Text
Option Explicit


Function EnmsAcObjTy$(A As AcObjectType)
Dim O$
Select Case True
Case A = AcObjectType.acDatabaseProperties: O = "acDatabaseProperties"
Case A = AcObjectType.acDefault: O = "acDefault"
Case A = AcObjectType.acDiagram: O = "acDiagram"
Case A = AcObjectType.acForm: O = "acForm"
Case A = AcObjectType.acFunction: O = "acFunction"
Case A = AcObjectType.acMacro: O = "acMacro"
Case A = AcObjectType.acModule: O = "acModule"
Case A = AcObjectType.acQuery: O = "acQuery"
Case A = AcObjectType.acReport: O = "acServerView"
Case A = AcObjectType.acServerView: O = "acServerView"
Case A = AcObjectType.acStoredProcedure: O = "acStoredProcedure"
Case A = AcObjectType.acTable: O = "acTable"
Case A = AcObjectType.acTableDataMacro: O = "acTableDataMacro"
Case Else: Thw CSub, "Invalid @AcObjTyp", "AcObjTyp", A
End Select
EnmsAcObjTy = O
End Function
