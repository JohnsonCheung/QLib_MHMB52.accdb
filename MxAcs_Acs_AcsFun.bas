Attribute VB_Name = "MxAcs_Acs_AcsFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_Acs_AcsFun."
Function TnyAcs(A As Access.Application) As String(): TnyAcs = Tny(A.CurrentDb): End Function

Function InAcsFb(A As Access.Application, Fb) As Boolean
On Error GoTo X
InAcsFb = A.CurrentDb.Name = Fb: Exit Function
X:
End Function

Function AcsGet() As Access.Application:              Set AcsGet = GetObject(, "Access.Application"): End Function
Function Acs() As Access.Application:                    Set Acs = Access.Application:                End Function
Function AcsDb(Db As Database) As Access.Application:  Set AcsDb = AcsFb(Db.Name):                    End Function

Function IsAcsOk(A As Access.Application) As Boolean
On Error GoTo X
Dim N$: N = A.Name
IsAcsOk = True
Exit Function
X:
End Function

Function AcsFb(Fb, Optional IsExl As Boolean) As Access.Application
If CFb = Fb Then Set AcsFb = Acs: Exit Function
Dim O As Access.Application: Set O = AcsNw
O.OpenCurrentDatabase Fb, IsExl
Set AcsFb = O
End Function

Function DftAcs(A As Access.Application) As Access.Application
'Ret :@A if Not Nothing or :AcsNw
If IsNothing(A) Then
    Set DftAcs = AcsNw
Else
    Set DftAcs = A
End If
End Function

Function FbAcs$(A As Access.Application)
On Error Resume Next
FbAcs = A.CurrentDb.Name
End Function

Function AcsNw() As Access.Application
Dim O As Access.Application: Set O = CreateObject("Access.Application")
MinvAcs O
Set AcsNw = O
End Function

Function PjAcs(A As Access.Application) As VBProject: Set PjAcs = A.VBE.ActiveVBProject: End Function

Sub QuitAcs(A As Access.Application)
If IsNothing(A) Then Exit Sub
On Error Resume Next
Stamp "QuitAcs: Begin"
Stamp "QuitAcs: Cls":         A.CloseCurrentDatabase
Stamp "QuitAcs: Quit":        A.Quit
Stamp "QuitAcs: Set Nothing": Set A = Nothing
Stamp "QuitAcs: End"
End Sub

Sub SavRec(): DoCmd.RunCommand acCmdSaveRecord: End Sub
