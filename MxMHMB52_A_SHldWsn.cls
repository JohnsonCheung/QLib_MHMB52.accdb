VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldWsn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Function IsSHld(Wsn$) As Boolean
Select Case True
Case IsFc(Wsn), IsSd(Wsn), IsSh(Wsn): IsSHld = True
End Select
End Function
Function IsSh(Wsn$) As Boolean: IsSh = HasPfx(Wsn, "StkHld "):  End Function
Function IsFc(Wsn$) As Boolean: IsFc = HasPfx(Wsn, "Fc "):      End Function
Function IsSd(Wsn$) As Boolean: IsSd = HasPfx(Wsn, "StkDays "): End Function
