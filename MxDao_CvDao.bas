Attribute VB_Name = "MxDao_CvDao"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_CvDao."

Function CvFds(A) As Dao.Fields:            Set CvFds = A: End Function
Function CvTd(A) As Dao.TableDef:            Set CvTd = A: End Function
Function CvIdxFds(A) As Dao.IndexFields: Set CvIdxFds = A: End Function
