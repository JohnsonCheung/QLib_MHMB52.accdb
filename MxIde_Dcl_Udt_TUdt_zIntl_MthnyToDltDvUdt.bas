Attribute VB_Name = "MxIde_Dcl_Udt_TUdt_zIntl_MthnyToDltDvUdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_AA_DvuFun."
Function MthnyDvUdtAll(Udtn$) As String()
Static X As Boolean, Y: If Not X Then X = True: Y = SySs(RplQ("? Si? Ub? Push? ?yAdd Push?y Som? Push?opt", Udtn))
MthnyDvUdtAll = Y
End Function
Function MthnyDvUdtToDlt(U As TUdt) As String()
With U
    Dim N$: N = U.Udtn
    Dim O$()
        With U
        If .GenCtor Then PushI O, N
          If .GenAy Then PushIAy O, SySs(RplQ("Si? Ub? Push?", N))
         If .GenOpt Then PushIAy O, SySs(RplQ("Som? Push?Opt", N))
         If .GenAdd Then PushIAy O, SySs(RplQ("?yAdd", N))
      If .GenPushAy Then PushIAy O, SySs(RplQ("Push?y", N))
        End With
End With
MthnyDvUdtToDlt = O
End Function
Function MthnyDvUdtToGen(U As TUdt) As String(): MthnyDvUdtToGen = SyMinus(MthnyDvUdtAll(U.Udtn), MthnyDvUdtToDlt(U)): End Function
Function MthnyDvUdtToDltDcl(Dcl$()) As String()
Dim Uy() As TUdt: Uy = TUdtyDcl(Dcl)
Dim J%: For J = 0 To UbTUdt(Uy)
    PushIAy MthnyDvUdtToDltDcl, MthnyDvUdtToDlt(Uy(J))
Next
End Function

