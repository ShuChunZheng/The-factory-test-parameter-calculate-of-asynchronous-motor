Attribute VB_Name = "Module1"
'***********************************************************************
'程序实现功能：用遗传算法求函数的最大值
'***********************************************************************
Option Explicit
'用来保存2的N次方的数据
Public N2(30) As Long
'调用其Eval函数
Public Script As Object
'交叉方式
Public Enum CrossOver
    OnePointCrossOver             '单点交叉
    TwoPointCrossOver             '两点交叉
    UniformCrossOver              '平均交叉
End Enum
'选择方式
Public Enum Selection
    RouletteWheelSelection        '轮盘赌选择
    StochasticTourament           '随机竞争选择
    RandomLeagueMatches           '随机联赛选择
    StochasticUniversalSampleing  '随机遍历取样
End Enum
'编码方式
Public Enum EnCoding
    Binary                        '标准二进制编码
    Gray                          '格雷码
End Enum
'自定义类型
Type GAinfo
    Max As Double
    Cordinate() As Double
End Type


'*********************************** 二进制码转格雷码 ***********************************
'
'函 数 名： BinaryToGray
'参    数： Value - 要转换的二进制数的实值
'说    明： 如3对应的二进制表示为0011，而用格雷码表示为0010，这个函数的value为0011代表的实数
'           而返回的是0010所代表的实数（2）
'返 回 值： 返回格雷码对应的二进制数的实值
'源 作 者： laviewpbt
'开发语言： C语言
'修 改 者： zsc
'
'*********************************** 二进制码转格雷码 ***********************************
Public Function BinaryToGray(Value As Long) As Long
    Dim V As Long, Max As Long
    Dim start As Long, mEnd As Long, Temp As Long, Counter As Long
    Dim Flag As Boolean
    V = Value: Max = 1
    While V > 0
        V = V / 2
        Max = Max * 2
    Wend
    If Max = 0 Then Exit Function
    Flag = True
    mEnd = Max - 1
    While start < mEnd
        Temp = (mEnd + start - 1) / 2
        If Value <= Temp Then
            If Not Flag Then
                Counter = Counter + (mEnd - start + 1) / 2
            End If
            mEnd = Temp
            Flag = True
        Else
            If Flag Then
                Counter = Counter + (mEnd - start + 1) / 2
            End If
            Temp = Temp + 1
            start = Temp
            Flag = False
        End If
    Wend
    BinaryToGray = Counter
End Function
'*********************************** 格雷码转二进制码 ***********************************
'
'函 数 名： BinaryToGray
'参    数： Value - 要转换的二进制数的实值
'说    明： 如3对应的二进制表示为0011，而用格雷码表示为0010，这个函数的value为0010代表的实数
'           而返回的是0010所代表的实数（2）
'返 回 值： 返回格雷码对应的二进制数的实值
'源 作 者： laviewpbt
'开发语言： C语言
'修 改 者： zsc
'
'*********************************** 格雷码转二进制码 ***********************************
Public Function GrayToBinary(Value As Long) As Long
    Dim V As Long, Max As Long
    Dim start As Long, mEnd As Long, Temp As Long, Counter As Long
    Dim Flag As Boolean
    V = Value: Max = 1
    While V > 0
        V = V / 2
        Max = Max * 2
    Wend
    Flag = True
    mEnd = Max - 1
    While start < mEnd
        Temp = Counter + (mEnd - start + 1) / 2
        If Flag Xor (Value < Temp) Then
           If Flag Then Counter = Temp
           start = (start + mEnd + 1) / 2
           Flag = False
        Else
           If Not Flag Then Counter = Temp
           mEnd = (start + mEnd - 1) / 2
           Flag = True
        End If
    Wend
    GrayToBinary = start
End Function
'*********************************** 十进制转转二进制码 ***********************************
'
'函 数 名： DecToBinary
'参    数： Value - 要转换的十进制数
'返 回 值： 返回对应的二进制数
'作    者： laviewpbt
'修 改 者： zsc
'
'*********************************** 十进制转转二进制码 ***********************************
Private Function DecToBinary(ByVal Value As Long) As String
    Dim StrTemp As String
    Dim ModNum As Integer
    Do While Value > 0
        ModNum = Value Mod 2
        Value = Value \ 2
        StrTemp = ModNum & StrTemp
    Loop
    DecToBinary = StrTemp
End Function
'************************************* 二十进制转换 **********************************
'
'函 数 名： BinToDec
'参    数： BinCode - 二进制字符串
'返 回 值： 转换后的十进制数
'说    明： 二进制字符串转换位十进制数
'作    者： laviewpbt
'修 改 者： zsc
'
'************************************* 二十进制转换 **********************************
Public Function BinToDec(BinCode As String) As Long
    Dim i As Integer, Dec As Long, Length As Integer
    Length = Len(BinCode)
    For i = 1 To Length
        If Mid(BinCode, i, 1) = "1" Then
            Dec = Dec + N2(Length - i)
        End If
    Next
    BinToDec = Dec
End Function

'************************************* 变量的二进制串位数 **********************************
'
'函 数 名： GetIndex
'参    数： Target - 待求数
'返 回 值： 某一指数
'说    明： 求符合2^(GetIndex-1)<Target<=2^GetIndex的 GetIndex
'作    者： laviewpbt
'修 改 者： zsc
'
'************************************* 变量的二进制串位数 **********************************
Public Function GetIndex(Target As Long) As Integer
    Dim i As Integer
    For i = 0 To 30
        If Target <= N2(i) Then
            GetIndex = i
            Exit Function
        End If
    Next
End Function

'************************************ Eval动态执行一个函数 *********************************
'
'函 数 名： CalcFun
'参    数： Fun    - 函数
'           Script - 一个ScriptControl对象
'           X1     － 第一各自变量
'           X2     － 第二各自变量，可选
'           X3     － 第三各自变量，可选
'           X4     － 第四各自变量，可选
'说    明： 动态执行一个函数，最多这支持四个参数，并且变量的形式只可写为X1/X2/X3/X4,GA函数
'           执行慢主要是这个Eval函数计算需要大量时间
'作    者： laviewpbt
'修 改 者： zsc
'
'************************************ Eval动态执行一个函数 *********************************
Public Function CalcFun(ByVal Fun As String, Script As Object, X1 As Double, Optional X2 As Double, Optional X3 As Double, Optional X4 As Double) As Double
    Fun = Replace(Fun, "X1", CStr(X1))
    If Not IsMissing(X2) Then Fun = Replace(Fun, "X2", CStr(X2))
    If Not IsMissing(X3) Then Fun = Replace(Fun, "X3", CStr(X3))
    If Not IsMissing(X4) Then Fun = Replace(Fun, "X4", CStr(X4))
    CalcFun = Script.Eval(Fun)
End Function




