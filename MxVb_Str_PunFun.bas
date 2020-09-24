Attribute VB_Name = "MxVb_Str_PunFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Pun."

Function IsAscPun(A%) As Boolean
'  0 1 2 3 4 5 6 7 8 9 A B C D E F
'0                
'1                
'2   ! " # $ % & ' ( ) * + , - . /
'3 0 1 2 3 4 5 6 7 8 9 : ; < = > ?
'4 @ A B C D E F G H I J K L M N O
'5 P Q R S T U V W X Y Z [ \ ] ^ _
'6 ` a b c d e f g h i j k l m n o
'7 p q r s t u v w x y z { | } ~ 
Select Case True
Case WIsPun1(A), WIsPun2(A), WIsPun3(A), WIsPun4(A): IsAscPun = True
End Select
End Function
Private Function WIsPun1(A%) As Boolean: WIsPun1 = (&H21 <= A And A <= &H2F): End Function
Private Function WIsPun2(A%) As Boolean: WIsPun2 = (&H3A <= A And A <= &H40): End Function
Private Function WIsPun3(A%) As Boolean: WIsPun3 = (&H5B <= A And A <= &H60): End Function
Private Function WIsPun4(A%) As Boolean: WIsPun4 = (&H7B <= A And A <= &H7F): End Function

Function IsPun(S$) As Boolean: IsPun = IsAscPun(Asc(S)): End Function

Private Sub B_PunyPure():                    VcAy PunyPure(SrclPC):                                   End Sub
Function PunyPure(S) As String(): PunyPure = AySrtQ(AwNB(DisChry(RmvDblSpc(RplCrLf(RplAlpNum(S)))))): End Function
Function RxPun() As RegExp
Static X As RegExp
If IsNothing(X) Then Set X = Rx("/[!""#$%&'()*+,-/:;<=>?@[\\\]^_`{\|}~]/g")
Set RxPun = X
'  0 1 2 3 4 5 6 7 8 9 A B C D E F
'0                
'1                
'2   ! " # $ % & ' ( ) * + , - . /
'3 0 1 2 3 4 5 6 7 8 9 : ; < = > ?
'4 @ A B C D E F G H I J K L M N O
'5 P Q R S T U V W X Y Z [ \ ] ^ _
'6 ` a b c d e f g h i j k l m n o
'7 p q r s t u v w x y z { | } ~ 
End Function
Private Sub B_RplPun():          VcStr RplPun(SrclPC):  End Sub
Function RplPun$(S):    RplPun = RxPun.Replace(S, " "): End Function ' Replace PunChr by space
