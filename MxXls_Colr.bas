Attribute VB_Name = "MxXls_Colr"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Colr."
Const ColrVbl_1$ = "ActiveBorder -4934476" & _
"|ActiveCaption -6703919" & _
"|ActiveCaptionText -16777216" & _
"|AliceBlue -984833" & _
"|AntiqueWhite -332841" & _
"|AppWorkspace -5526613" & _
"|Aqua -16711681" & _
"|Aquamarine -8388652" & _
"|Azure -983041" & _
"|Beige -657956" & _
"|Bisque -6972" & _
"|Black -16777216" & _
"|BlanchedAlmond -5171" & _
"|Blue -16776961" & _
"|BlueViolet -7722014" & _
"|Brown -5952982" & _
"|BurlyWood -2180985" & _
"|ButtonFace -986896" & _
"|ButtonHighlight -1" & _
"|ButtonShadow -6250336"
Const ColrVbl_2$ = "|CadetBlue -10510688" & _
"|Chartreuse -8388864" & _
"|Chocolate -2987746" & _
"|Control -986896" & _
"|ControlDark -6250336" & _
"|ControlDarkDark -9868951" & _
"|ControlLight -1842205" & _
"|ControlLightLight -1" & _
"|ControlText -16777216" & _
"|Coral -32944" & _
"|CornflowerBlue -10185235" & _
"|Cornsilk -1828" & _
"|Crimson -2354116" & _
"|Cyan -16711681" & _
"|DarkBlue -16777077" & _
"|DarkCyan -16741493" & _
"|DarkGoldenrod -4684277" & _
"|DarkGray -5658199" & _
"|DarkGreen -16751616" & _
"|DarkKhaki -4343957"
Const ColrVbl_3$ = "|DarkMagenta -7667573" & _
"|DarkOliveGreen -11179217" & _
"|DarkOrange -29696" & _
"|DarkOrchid -6737204" & _
"|DarkRed -7667712" & _
"|DarkSalmon -1468806" & _
"|DarkSeaGreen -7357301" & _
"|DarkSlateBlue -12042869" & _
"|DarkSlateGray -13676721" & _
"|DarkTurquoise -16724271" & _
"|DarkViolet -7077677" & _
"|DeepPink -60269" & _
"|DeepSkyBlue -16728065" & _
"|Desktop -16777216" & _
"|DimGray -9868951" & _
"|DodgerBlue -14774017" & _
"|Firebrick -5103070" & _
"|FloralWhite -1296" & _
"|ForestGreen -14513374" & _
"|Fuchsia -65281"
Const ColrVbl_4$ = "|Gainsboro -2302756" & _
"|GhostWhite -460545" & _
"|Gold -10496" & _
"|Goldenrod -2448096" & _
"|GradientActiveCaption -4599318" & _
"|GradiEntitynactiveCaption -2628366" & _
"|Gray -8355712" & _
"|GrayText -9605779" & _
"|Green -16744448" & _
"|GreenYellow -5374161" & _
"|Highlight -16746281" & _
"|HighlightText -1" & _
"|Honeydew -983056" & _
"|HotPink -38476" & _
"|HotTrack -16750900" & _
"|InactiveBorder -722948" & _
"|InactiveCaption -4207141" & _
"|InactiveCaptionText -16777216" & _
"|IndianRed -3318692" & _
"|Indigo -11861886"
Const ColrVbl_5$ = "|INF -31" & _
"|INFText -16777216" & _
"|Ivory -16" & _
"|Khaki -989556" & _
"|Lavender -1644806" & _
"|LavenderBlush -3851" & _
"|LawnGreen -8586240" & _
"|LemonChiffon -1331" & _
"|LightBlue -5383962" & _
"|LightCoral -1015680" & _
"|LightCyan -2031617" & _
"|LightGoldenrodYellow -329006" & _
"|LightGray -2894893" & _
"|LightGreen -7278960" & _
"|LightPink -18751" & _
"|LightSalmon -24454" & _
"|LightSeaGreen -14634326" & _
"|LightSkyBlue -7876870" & _
"|LightSlateGray -8943463" & _
"|LightSteelBlue -5192482"
Const ColrVbl_6$ = "|LightYellow -32" & _
"|Lime -16711936" & _
"|LimeGreen -13447886" & _
"|Linen -331546" & _
"|Magenta -65281" & _
"|Maroon -8388608" & _
"|MediumAquamarine -10039894" & _
"|MediumBlue -16777011" & _
"|MediumOrchid -4565549" & _
"|MediumPurple -7114533" & _
"|MediumSeaGreen -12799119" & _
"|MediumSlateBlue -8689426" & _
"|MediumSpringGreen -16713062" & _
"|MediumTurquoise -12004916" & _
"|MediumVioletRed -3730043" & _
"|Menu -986896" & _
"|MenuBar -986896" & _
"|MenuHighlight -13395457" & _
"|MenuText -16777216" & _
"|MidnightBlue -15132304"
Const ColrVbl_7$ = "|MintCream -655366" & _
"|MistyRose -6943" & _
"|Moccasin -6987" & _
"|NavajoWhite -8531" & _
"|Navy -16777088" & _
"|OldLace -133658" & _
"|Olive -8355840" & _
"|OliveDrab -9728477" & _
"|Orange -23296" & _
"|OrangeRed -47872" & _
"|Orchid -2461482" & _
"|PaleGoldenrod -1120086" & _
"|PaleGreen -6751336" & _
"|PaleTurquoise -5247250" & _
"|PaleVioletRed -2396013" & _
"|PapayaWhip -4139" & _
"|PeachPuff -9543" & _
"|Peru -3308225" & _
"|Pink -16181" & _
"|Plum -2252579"
Const ColrVbl_8$ = "|PowderBlue -5185306" & _
"|Purple -8388480" & _
"|Red -65536" & _
"|RosyBrown -4419697" & _
"|RoyalBlue -12490271" & _
"|SaddleBrown -7650029" & _
"|Salmon -360334" & _
"|SandyBrown -744352" & _
"|ScrollBar -3618616" & _
"|SeaGreen -13726889" & _
"|SeaShell -2578" & _
"|Sienna -6270419" & _
"|Silver -4144960" & _
"|SkyBlue -7876885" & _
"|SlateBlue -9807155" & _
"|SlateGray -9404272" & _
"|Snow -1286" & _
"|SpringGreen -16711809" & _
"|SteelBlue -12156236" & _
"|Tan -2968436"
Const ColrVbl_9$ = "|Teal -16744320" & _
"|Thistle -2572328" & _
"|Tomato -40121" & _
"|Transparent 16777215" & _
"|Turquoise -12525360" & _
"|Violet -1146130" & _
"|Wheat -663885" & _
"|White -1" & _
"|WhiteSmoke -657931" & _
"|Window -1" & _
"|WindowFrame -10197916" & _
"|WindowText -16777216" & _
"|Yellow -256" & _
"|YellowGreen -6632142"
Const ColrVbl$ = ColrVbl_1 & ColrVbl_2 & ColrVbl_3 & ColrVbl_4 & ColrVbl_5 & vbCrLf & ColrVbl_6 & ColrVbl_7 & ColrVbl_8 & ColrVbl_9

Function Colrny() As String()
Static N$()
If Si(N) = 0 Then
    Dim L: For Each L In SplitVBar(ColrVbl)
        PushI N, Tm1(L)
    Next
End If
Colrny = N
End Function

Function SqColr() As Variant()
Dim J%, O(), Ly$(), Nm$, Colr&
Ly = Colrny
ReDim O(1 To Si(Ly), 1 To 2)
For J = 1 To Si(Ly)
    AsgT1r Ly(J - 1), Nm, Colr
    O(J, 1) = Nm
    O(J, 2) = Colr
Next
SqColr = O
End Function

Function Colrn$(Colrix&): Colrn = Colrny()(Colrix): End Function

Function IsColrn(Nm) As Boolean: IsColrn = HasEle(Colrny, Nm):   End Function
Function Colrix&(Colrn$):         Colrix = IxEle(Colrny, Colrn): End Function

Function WbColr() As Workbook
Dim Ws As Worksheet, Sq(), J%
Sq = SqColr
'Set Ws = WsRg(RgSq(SqColr, A1Nw))
For J = 1 To UBound(Sq(), 1)
    RgWsRC(Ws, J, 3).Interia.Colr = Sq(J, 2)
Next
RgWsCC(Ws, 1, 2).EntireColumn.AutoFit
Set WbColr = WbWs(Ws)
Maxv WbColr.Application
End Function

Sub SetColr_ToDo()
'TstStep
'   Call Gen
'   Call FmtTSpec_Brw 'Edt
'       Edit and Save, then Call Gen will auto import
'where to add autoImp?
'   Under FmtWbAllLo
'AutoImp will show msg if import/noImport
'Colrny
'   what is the common Colr name in DotNet Library
'       Use Enums: System.Drawing.KnownColr is no good, because the EnmnLn is in seq, it is not return
'       Use VBA.colrCnstants-module is good, but there is few constant
'       Answer: Use *KnownColr to feed in struct-*Colr, there is *Colr.ToArgb & *KnownColr has name
'               Run the FSharp program.
'               Put the generated file
'                   in
'                       C:\Users\user\Source\Repos\EnumLines\EnumLines\bin\Debug\ColrLines.Const.Txt
'                   Into
'                       C:\Users\user\Desktop\Mhd\SAPAccessReports\StockShipRate\StockShipRate\Spec
'               Run ConstGen: It will addd the Const ColrLines = ".... at end
'               Put Fct-Module
'To find some common values to feed into ColrLines
'
'Colr* 4-functions
'    Colrn_MayColr
'    Colrn
'    Colrny
'    ColrLines
End Sub

Function CDicDrsolrnqColr$()
'Aim            : Use Assembly to build the KnowColor-Vb module
'Assembly       : System.Drawing.dll
'Enum           : System.Drawing.KnownColor
'KnownColorCount: 174 - 1  ! All Alpha-Value = 255, only 1 knownColor has Alphas-Value = 2, which is 'Transparent'
'Target-Vb-Fun  : DiColrnqColr

'open System.Drawing
'open System
'open System.IO
'open System.Windows.Forms
'
'type ss = String list
'type sy = String[]
'type ss = String seq
'type als<'a> = 'a list
'type aay<'a> = 'a[]
'type asq<'a> = 'a seq
'let lines'sl(a:sl) = String.Join("\r\n",a)
'let lines'sy(a:sy) = String.Join("\r\n",a)
'let wrt_str  ft a = File.WriteAllText(ft,a)
'let wrt_strSeq ft (a:ss) = File.WriteAllLines(ft,a)
'let wrt_strList ft a = a|> wrt_sseq ft
'let wrt_mayStr ft a = match a with | Some a -> str_wrt a ft | _ -> ()
'let colrCnstFt = "ColrLines.Txt"
'//let knownColr'Ln colr = colr.ToString() + " " + Colr.FromKnownColr(colr).ToArgb().ToString()
'let known'ColrCnstlin colr = "Const " + a.ToString() + "& = " + Colr.FromKnownColr(a).ToArgb().ToString()
'let wrt_sy ft a = a |> wrt_sseq ft
'let asq'ay<'a>(a:Array) = seq { for i in a -> unbox i }
'let aay'ay<'a>(a:Array) = [|for i in a -> unbox i|]
'let als'ay<'a>(a:Array) = [for i in a -> unbox i]
'let ``known:ColrAy`` = Enum.GetValues(KnownColr.ActiveBorder.GetType())
'let ``known:ColrLs`` = knownColrAy |> al'ay<KnownColr>
'let  ``colr:CnstLs`` = knownColrLs |> List.map knownColr'Ln |> List.sort
'[<EntryPoint>]
'let main argv =
'    wrt_sl colrCnstFt colrCnst`Ls
'    do wrt_colrCnstFt()
'    0 // return an integer exit code
End Function
