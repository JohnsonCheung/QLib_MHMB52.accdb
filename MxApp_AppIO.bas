Attribute VB_Name = "MxApp_AppIO"
Option Compare Text
Option Explicit
Const CMod$ = "MxApp_AppIO."
Private Type A
    Root As String
    Nm As String
    Ver As String
    HomOup As String
End Type
Private A As A

Function AppFxmDft$(Optional Tpn$)
Dim TpnDft$: TpnDft = StrDft(Tpn, A.Nm)
AppFxmDft = AppHom & AppFxm(TpnDft)
End Function

Function AppFn$(Tpn$):   AppFn = Tpn & "(Template).xlsx": End Function
Function AppFxm$(Tpn$): AppFxm = Tpn & "(Template).xlsm": End Function

Function AppPth$()
On Error GoTo E
Static O$
With A
If O = "" Then O = PthAddFdrAp(.Root, .Nm, .Ver)
End With
AppPth = O
E:
End Function

Function AppFxo$(): AppFxo = PthOup & A.Nm & ".xlsx": End Function
Function aAppFb$(): aAppFb = AppPth & "AppFb.accdb":  End Function

Function AppHom$()
On Error GoTo E
Static Y$
With A
If Y = "" Then Y = PthAddFdrApEns(.Root, .Nm, .Ver)
End With
AppHom = Y
E:
End Function
