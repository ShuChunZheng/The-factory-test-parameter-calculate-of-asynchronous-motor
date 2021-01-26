VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "异步电机检查试验合格区"
   ClientHeight    =   5730
   ClientLeft      =   2325
   ClientTop       =   2775
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   13530
   Begin VB.Frame Frame5 
      Caption         =   "出厂试验数据参数"
      Height          =   855
      Left            =   9480
      TabIndex        =   78
      Top             =   4800
      Width           =   3975
      Begin VB.TextBox Text36 
         Height          =   375
         Left            =   960
         TabIndex        =   79
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label29 
         Caption         =   "台数"
         Height          =   375
         Left            =   360
         TabIndex        =   80
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   10560
      ScaleHeight     =   4395
      ScaleWidth      =   2835
      TabIndex        =   77
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame9 
      Caption         =   "电机额定值"
      Height          =   1815
      Left            =   6960
      TabIndex        =   67
      Top             =   2880
      Width           =   3375
      Begin VB.TextBox Text28 
         Height          =   375
         Left            =   600
         TabIndex        =   73
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text27 
         Height          =   375
         Left            =   600
         TabIndex        =   72
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   600
         TabIndex        =   71
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label37 
         Caption         =   "r/min"
         Height          =   375
         Left            =   2520
         TabIndex        =   76
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label36 
         Caption         =   "V"
         Height          =   375
         Left            =   2640
         TabIndex        =   75
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label35 
         Caption         =   "W"
         Height          =   375
         Left            =   2640
         TabIndex        =   74
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "ne"
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "Ue"
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Pe"
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      TabIndex        =   66
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Frame Frame8 
      Caption         =   "型式试验值"
      Height          =   2775
      Left            =   2280
      TabIndex        =   45
      Top             =   2880
      Width           =   4575
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   960
         TabIndex        =   65
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   3000
         TabIndex        =   64
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   3000
         TabIndex        =   63
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   3000
         TabIndex        =   60
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   3000
         TabIndex        =   59
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   3000
         TabIndex        =   58
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   960
         TabIndex        =   57
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   960
         TabIndex        =   52
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   960
         TabIndex        =   51
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   960
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Pt1"
         Height          =   375
         Left            =   2520
         TabIndex        =   62
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "P2"
         Height          =   375
         Left            =   2520
         TabIndex        =   61
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label23 
         Caption         =   "Pt2"
         Height          =   375
         Left            =   2520
         TabIndex        =   56
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "r1"
         Height          =   375
         Left            =   2640
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "i1"
         Height          =   375
         Left            =   2640
         TabIndex        =   54
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "iqt/ie"
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Mqt/Me"
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Mmax/Me"
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "cosΨ"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "η"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "技术条件上的极限值"
      Height          =   2655
      Left            =   7800
      TabIndex        =   25
      Top             =   120
      Width           =   2535
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   1320
         TabIndex        =   44
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   1320
         TabIndex        =   43
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   1320
         TabIndex        =   42
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   1320
         TabIndex        =   41
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1320
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "(iqt/ie)max:"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "(Mqt/Me)min:"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "(Mmax/Me)min:"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "(cosΨ)min:"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "(η)min:"
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "技术条件上的保证值"
      Height          =   2655
      Left            =   5280
      TabIndex        =   23
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1200
         TabIndex        =   39
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   1200
         TabIndex        =   38
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1200
         TabIndex        =   37
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1200
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1200
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "(iqt/ie)e:"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "(Mqt/Me)e:"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "(Mmax/Me)e:"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "(cosΨ)e:"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "(η)e:"
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "GA关键参数"
      Height          =   2655
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "编码精度："
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "变异概率："
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "交叉概率："
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "迭代次数："
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "种群大小："
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "交叉方式"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
      Begin VB.OptionButton Option9 
         Caption         =   "平均"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Caption         =   "两点"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "单点"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "编码方式"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
      Begin VB.OptionButton Option6 
         Caption         =   "格雷码"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "标准二进制"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择方式"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Option4 
         Caption         =   "随机遍历"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "随机联赛"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "随机竞争"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "轮盘赌"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GA关键参数
'allmenber       - 群体的大小（20～100）
'DigitNum   - 影响编码位数的一个参数（1～5）
'Pc      - 交叉概率（0.4～0.99）
'Pm      - 变异概率（0.0001～0.1）
'MaxIter - 最大迭代次数（100～500）
Private allmenber As Integer
Private DigitNum As Integer
Private Pc As Double
Private Pm As Double
Private MaxIter As Integer
'技术条件上的保证值
Private ηe As Single
Private cosΨe As Single
Private MmaxMee As Single
Private MqtMee As Single
Private iqtiee As Single
'技术条件上的极限值
Private ηmin As Single
Private cosΨmin As Single
Private MmaxMemin As Single
Private MqtMemin As Single
Private iqtiemin As Single
'型式试验值
Private η As Single
Private cosΨ As Single
Private MmaxMe As Single
Private MqtMe As Single
Private iqtie As Single
Private i1d As Single
Private r1 As Single
Private P2 As Single
Private Pt1 As Single
Private Pt2 As Single
Private Pe As Single
Private Ue As Single
Private ne As Single
'出厂试验数据
Dim evryhanval1() As Single
Dim evryhanval2() As Single
Dim evryhanval3() As Single
Dim evryhanval4() As Single
Dim evryhanval5() As Single
Dim evryhanval6() As Single
Dim jishuzhi As Integer
Private taishu As Integer
Private han As Integer
Private Uo As Single
Private io As Single
Private Po As Single
Private Ud As Single
Private id As Single
Private Pd As Single
Private rd As Single
'方程1与方程2的输入
Dim bianlian1(1 To 1, 1 To 2) As Double
Dim prioutval As Double
Dim outputcanshu As Double
Dim outputcanshu1 As Double
Dim outputcanshu2 As Double
'方程3的输入
Dim bianlian2(1 To 2, 1 To 2) As Double
Dim prioutval1 As Double
Dim outputcanshu3 As Double


'*********************************** 参数计算 ***********************************
'
'说    明： 将函数中涉及的且与变量无关的参数以function的方式给出
'作    者： zsc
'
'*********************************** 参数计算 ***********************************
'参数idmax 待输入值id
Private Function idmax(Value As Single) As Double
    idmax = iqtiemin / iqtie * Value
End Function

'参数idmin 待输入值id
Private Function idmin(Value As Single) As Double
    idmin = MmaxMemin / MmaxMe * Value
End Function
'参数rdmax 待输入值rd
Private Function rdmax(Value As Single) As Double
        rdmax = ((MmaxMe + Sqr(MmaxMe ^ 2 - 1) - 2 * η * ΔPmax() / P2) / MmaxMe) * (Value - r1 / 1.5) + r1 / 1.5
End Function
'参数Canshu1 待输入值io,id
Private Function Canshu1(Value1 As Single, Value2 As Single) As Double
    Canshu1 = idmax(Value2) * 5 - Value1 - i1d * (cosΨ - cosΨmin) / (cosΨ * Sqr(1 - cosΨ ^ 2))
End Function
'参数Canshu11  待输入值Ud id
Private Function Canshu11(Value1 As Single, Value2 As Single) As Double
    Canshu11 = Value1 / Ue * ((i1d * cosΨ) / Value2) ^ 2 + 3 * (Value1 / Ue) ^ 3 * ((i1d * cosΨ) / Value2) ^ 4
End Function
'参数Canshu2 待输入值Pd id
Private Function Canshu2(Value1 As Single, Value2 As Single) As Double
    Canshu2 = MmaxMemin / MqtMe * (Value1 - 1.5 * Value2 ^ 2 * r1 / 1.5)
End Function
'参数Canshu21
Private Function Canshu21() As Double
    Canshu21 = 1.5 * r1 / 1.5
End Function
'参数k1 待输入值Po
Private Function k1(Value As Single) As Double
    k1 = Value + (1 - 2 * η * (Pt1 + Pt2) / P2) * ΔPmax()
End Function
'参数k2
Private Function k2() As Double
    k2 = 2 * Pt2 / i1d * Sqr(1 - cosΨ ^ 2)
End Function
'参数k3 待输入值rd
Private Function k3(Value As Single) As Double
    k3 = Pt2 / (Value - r1 / 1.5)
End Function
'参数ΔPmax
Private Function ΔPmax() As Double
    ΔPmax = (η - ηmin) * P2 / η ^ 2
End Function
'********************************* 标准遗传算法 **********************************
'
'函 数 名： GA
'参    数： Fun     - 待求的函数（变量的形式位X1,X2....）
'           ST      － 约束条件,第一维指向各个变量，第二维指向相关变量的上限和下限,存储值为相关变量的上限和下限

'           CodingMethod    － 编码的方法，二种可选
'           SelectionMethod － 选择的模式，三种可选
'           CrossOver       － 交叉的模式，三种可选
'返 回 值： 函数的最大值
'说    明： 标准遗传算法求解单目标函数
'作    者： laviewpbt
'修 改 者： zsc
'
'********************************* 标准遗传算法 *************************************

 Private Function GA(ST() As Double, Optional io1 As Single, Optional id1 As Single, Optional Ud1 As Single, Optional Pd1 As Single, Optional Po1 As Single, Optional rd1 As Single, Optional Fun As String, Optional funselect As Integer = 1, Optional CodingMethod As EnCoding = EnCoding.Binary, Optional SelectionMethod As Selection = Selection.RouletteWheelSelection, Optional CrossOverMethod As CrossOver = CrossOver.OnePointCrossOver) As GAinfo
    Dim i As Integer, j As Integer
    Dim Temp As String
    Dim i1 As Integer, j1 As Integer, j2 As Integer
    Dim i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer, i7 As Integer, i8 As Integer, i9 As Integer, i10 As Integer, i11 As Integer
    Dim Temp1 As Integer, Temp2 As Double
    Dim ST_Num As Integer                   '约束的个数,其实就是自由变量的个数
    Dim BitsSum As Integer                  '种群的二进制数的个数和
    Dim F As Double                         '群体总适应度
    Dim IterNum As Integer                  '迭代次数
    ReDim Result(1 To allmenber) As Double          '适应度
    ST_Num = UBound(ST, 1)
    ReDim Bits(1 To ST_Num) As Integer      'Fun函数中每个自由变量用二进制串表示时的位数
    Dim BinGroup() As String
    ReDim BinGroup(1 To allmenber) As String        '初始种群
    ReDim DecGroup(1 To allmenber, 1 To ST_Num) As Double '保存种群二进制所对应的十进制数
    ReDim q(allmenber) As Double                    '累计概率,以0为数组下标，有利于后面的轮盘赌选择
    Dim Parent() As Integer                 '作为父辈并进行交叉的染色体下标
    Dim MaxIndex As Long, Max As Double     '最大值和获得最大值的染色体的下标

    For i = 1 To ST_Num
        Bits(i) = GetIndex((ST(i, 2) - ST(i, 1)) * 10 ^ DigitNum) '每个字符串所需要的二进制串位数
        BitsSum = BitsSum + Bits(i)
    Next
    
'产生随机二进制种群
'*********************************** 编码 ***********************************
'

'参    数： BitsSum     - 需要编码的位数
'           BinGroup - 保存群体编码数据的数组
'说    明： 编码，准确的说应该是初始化种群，对于二进制码和格雷码这个过程一样的
'作    者： laviewpbt
'修 改 者： zsc
'
'*********************************** 编码 ***********************************
    Randomize
    For i1 = 1 To UBound(BinGroup, 1)
        Temp = ""
        For j1 = 1 To BitsSum
            If Rnd >= 0.5 Then
                Temp = Temp & "1"
            Else
                Temp = Temp & "0"
            End If
        Next
        BinGroup(i1) = Temp
    Next
      
    
    Do
        Randomize (Timer)
        IterNum = IterNum + 1
        Decoding Bits, ST, BinGroup, DecGroup, CodingMethod
        For i2 = 1 To allmenber
            If ST_Num = 1 And funselect = 1 Then
               ' Result(i) = CalcFun(Fun, Script, DecGroup(i, 1))       '计算各染色体的适应度
                Result(i2) = Canshu1(io1, id1) - Canshu11(Ud1, id1) * (DecGroup(i2, 1) - id1)
                'Result(i) = -Sin(DecGroup(i, 1)) + 0.5
            ElseIf funselect = 2 Then
                Result(i2) = Canshu2(Pd1, id1) + Canshu21() * DecGroup(i2, 1) ^ 2
                'Result(i) = DecGroup(i, 1) ^ 2 + DecGroup(i, 2) ^ 3
                'Result(i) = CalcFun(Fun, Script, DecGroup(i, 1), DecGroup(i, 2))
            ElseIf ST_Num = 2 Then
                Result(i2) = 5 * Pe - k1(Po1) + k2() * (DecGroup(i2, 1) - io1) + k3(rd1) * (DecGroup(i2, 2) - rd1)
                'Result(i) = CalcFun(Fun, Script, DecGroup(i, 1), DecGroup(i, 2), DecGroup(i, 3))
          '  ElseIf ST_Num = 4 Then
          '      Result(i) = 2 * Sin(DecGroup(i, 1) ^ 2) + DecGroup(i, 2) ^ 3 + 2 * DecGroup(i, 3) + 5 * DecGroup(i, 4) ^ 4
          '      Result(i) = CalcFun(Fun, Script, DecGroup(i, 1), DecGroup(i, 2), DecGroup(i, 3), DecGroup(i, 4))
            End If
        Next
        
        F = 0
        For i3 = 1 To allmenber
            F = F + Result(i3)       '计算群体的总适应度
        Next
        q(1) = Result(1) / F
        For i4 = 2 To allmenber
            q(i4) = q(i4 - 1) + Result(i4) / F   '计算每个染色体的累计概率
        Next
        If SelectionMethod = RouletteWheelSelection Then
            Roulette_Wheel_Selection q, BinGroup
        ElseIf SelectionMethod = StochasticTourament Then
            Stochastic_Tournament q, BinGroup, Result
        ElseIf SelectionMethod = RandomLeagueMatches Then
            Random_League_Matches BinGroup, Result, 4
        Else
            Stochastic_Universal_Sampleing BinGroup, Result, UBound(Result)
        End If
        
       
        Temp1 = 0
        For i5 = 1 To allmenber
            Temp2 = Rnd
            If Temp2 < Pc Then
                Temp1 = Temp1 + 1
                ReDim Preserve Parent(Temp1)        '选择交叉的一个父辈
                Parent(Temp1) = i5
            End If
        Next
        If CrossOverMethod = OnePointCrossOver Then
            For i6 = 1 To (Temp1 \ 2) * 2 Step 2
                OnePoint_CrossOver BinGroup(Parent(i6)), BinGroup(Parent(i6 + 1))
            Next
        ElseIf CrossOverMethod = TwoPointCrossOver Then
            For i7 = 1 To (Temp1 \ 2) * 2 Step 2
                TwoPoint_CrossOver BinGroup(Parent(i7)), BinGroup(Parent(i7 + 1))
            Next
        Else
            For i8 = 1 To (Temp1 \ 2) * 2 Step 2
                Uniform_CrossOver BinGroup(Parent(i8)), BinGroup(Parent(i8 + 1))
            Next
        End If
        
        For i9 = 1 To allmenber
            For j2 = 1 To BitsSum
                Temp2 = Rnd
                If Temp2 < Pm Then
                    Mutation BinGroup(i9), j2    '变异
                End If
            Next
        Next
   
    Loop While IterNum < MaxIter
        Max = -1000000
        For i10 = 1 To allmenber
            If Max < Result(i10) Then
                Max = Result(i10)
                MaxIndex = i10
            End If
        Next
        GA.Max = Max
        ReDim GA.Cordinate(1 To ST_Num)
        For i11 = 1 To ST_Num
            GA.Cordinate(i11) = DecGroup(MaxIndex, i11)
        Next
    End Function


'*********************************** 解码 ***********************************
'
'过 程 名： Decoding
'参    数： Bits     - 需要编码的位数
'           ST       - 约束条件
'           BinGroup - 学要解码的数组
'           DecGroup - 保存解码后的十进制数
'说    明： 解码
'修 改 者： zsc
'
'*********************************** 解码 ***********************************
Private Sub Decoding(Bits() As Integer, ST() As Double, BinGroup() As String, DecGroup() As Double, Method As EnCoding)
    Dim m As Integer, i As Integer, j As Integer, ST_Num As Integer, Temp As Integer
    ST_Num = UBound(Bits, 1)
    m = UBound(BinGroup, 1)
    If Method = Binary Then
        For i = 1 To m
            DecGroup(i, 1) = BinToDec(Left(BinGroup(i), Bits(1)))
            Temp = 1
            For j = 2 To ST_Num
                Temp = Temp + Bits(j - 1)
                DecGroup(i, j) = BinToDec(Mid(BinGroup(i), Temp, Bits(j)))
            Next
        Next
    ElseIf Method = Gray Then
        For i = 1 To m
            DecGroup(i, 1) = BinaryToGray(BinToDec(Left(BinGroup(i), Bits(1))))
            Temp = 1
            For j = 2 To ST_Num
                Temp = Temp + Bits(j - 1)
                DecGroup(i, j) = BinaryToGray(BinToDec(Mid(BinGroup(i), Temp, Bits(j))))
            Next
        Next
    End If
    
    For i = 1 To m
        For j = 1 To ST_Num
            DecGroup(i, j) = ST(j, 1) + DecGroup(i, j) * (ST(j, 2) - ST(j, 1)) / (N2(Bits(j)) - 1)
        Next
    Next
End Sub
'************************************* 轮盘赌选择 **********************************
'
'过 程 名： Roulette_Wheel_Selection
'参    数： Q        - 累计概率
'           BinGroup - 染色体数据
'说    明： 运用轮盘赌方法进行选择
'作    者： laviewpbt
'修 改 者： zsc
'
'************************************* 轮盘赌选择 **********************************
Private Sub Roulette_Wheel_Selection(q() As Double, ByRef BinGroup() As String)
    Dim i As Integer, j As Integer, m As Integer
    Dim DblTemp As Double
    m = UBound(BinGroup)
    ReDim TempBinGroup(1 To m) As String
    For i = 1 To m
        TempBinGroup(i) = BinGroup(i)       '备份原数据
    Next
    For i = 1 To m
        DblTemp = Rnd
        For j = 0 To m - 1
            If DblTemp <= q(j + 1) Then
                BinGroup(i) = TempBinGroup(j + 1)        '运用轮盘赌方法选择新的种群
                Exit For
            End If
        Next
    Next
End Sub
'************************************* 随机竞争选择 **********************************
'
'过 程 名： Stochastic_Tournament
'参    数： Q        - 累计概率
'           BinGroup - 染色体数据
'           Result   - 染色体的适应度数据
'说    明： 运用随机竞争进行选择（是基于轮盘赌选择的）
'作    者： laviewpbt
'修 改 者： zsc
'
'************************************* 随机竞争选择 **********************************
Private Sub Stochastic_Tournament(q() As Double, ByRef BinGroup() As String, Result() As Double)
    Dim i As Integer, j As Integer, m As Integer, Index1 As Integer, Index2 As Integer
    Dim DblTemp As Double
    m = UBound(BinGroup)
    ReDim TempBinGroup(1 To m) As String
    For i = 1 To m
        TempBinGroup(i) = BinGroup(i)       '备份原数据
    Next
    For i = 1 To m
        DblTemp = Rnd
        For j = 0 To m - 1
            If DblTemp <= q(j + 1) Then
                Index1 = j + 1               ' 运用轮盘赌方法得到一个个体
                Exit For
            End If
        Next
        DblTemp = Rnd
        For j = 0 To m - 1
            If DblTemp <= q(j + 1) Then       ' 运用轮盘赌方法得到另外一个个体
                Index2 = j + 1
                Exit For
            End If
        Next
        If Result(Index1) > Result(Index2) Then     '取适应度高的
            BinGroup(i) = TempBinGroup(Index1)        '运用随机竞争方法选择新的种群
        Else
            BinGroup(i) = TempBinGroup(Index2)        '运用轮盘赌方法选择新的种群
        End If
    Next
End Sub
'************************************* 随机联赛选择 **********************************
'
'过 程 名： Random_League_Matches
'参    数： BinGroup - 染色体数据
'           Result   - 染色体的适应度数据
'           N        - 联赛规模，常取2
'说    明： 运用随机联赛选择进行选择，似乎结果非常好,并且可以处理负的适应度
'作    者： laviewpbt
'修 改 者： zsc
'
'************************************* 随机联赛选择 **********************************
Private Sub Random_League_Matches(ByRef BinGroup() As String, Result() As Double, n As Double)
    Dim i As Integer, j As Integer, m As Integer, Index As Integer
    Dim DblTemp As Double, RndTemp As Integer
    m = UBound(BinGroup)
    ReDim TempBinGroup(1 To m) As String
    For i = 1 To m
        TempBinGroup(i) = BinGroup(i)       '备份原数据
    Next
    For i = 1 To m
        DblTemp = -100000000
        For j = 1 To n
            RndTemp = Int(1 + Rnd * m)
            If DblTemp < Result(RndTemp) Then ' 比较N个个体的适应度的大小
                Index = RndTemp
                DblTemp = Result(RndTemp)
            End If
        Next
        BinGroup(i) = TempBinGroup(Index)       '运用随机联赛方法选择新的种群
    Next
End Sub
'************************************* 随机全局取样选择 **********************************
'
'过 程 名： Stochastic_Universal_Sampleing
'参    数： BinGroup - 染色体数据
'           Result   - 染色体的适应度数据
'           N        - 联赛规模，没有考虑到代沟的话就取ubound(Result)
'说    明： 随机全局取样选择，似乎结果非常好,但必须要求待求函数在取值区间内全为正数
'作    者： laviewpbt
'修 改 者： zsc
'
'************************************* 随机全局取样选择 **********************************
Private Sub Stochastic_Universal_Sampleing(ByRef BinGroup() As String, Result() As Double, n As Integer)
    Dim m As Long, i As Integer, j As Integer
    Dim rd() As Double
    m = UBound(Result)
    ReDim CumFit(1 To m) As Double      '累计概率
    ReDim Trials(1 To n) As Double
    ReDim rd(1 To m) As Double
    ReDim Index(1 To n) As Integer
    ReDim TempBinGroup(1 To m) As String
    Dim Temp As Integer
    ReDim a(1 To n) As Integer
    CumFit(1) = Result(1)
    For i = 2 To m
        CumFit(i) = CumFit(i - 1) + Result(i)
    Next
    For i = 1 To n
        Trials(i) = CumFit(m) / n * (Rnd + (i - 1))
    Next
    rd(1) = 0
    For i = 2 To m
        rd(i) = CumFit(i - 1)
    Next
    For i = 1 To n
        For j = 1 To m
            If Trials(i) < CumFit(j) And rd(j) <= Trials(i) Then
                Temp = Temp + 1
                Index(Temp) = j
            End If
        Next
    Next
    
    For i = 1 To m
        TempBinGroup(i) = BinGroup(i)       '备份原数据
    Next
    For i = 1 To n
        a(i) = Int(Rnd * n) + 1
        For j = 1 To i - 1
            If a(i) = a(j) Then
                i = i - 1           '不重复的随机数
                Exit For
            End If
        Next
    Next
    For i = 1 To m
        BinGroup(i) = TempBinGroup(Index(a(i)))
    Next
End Sub
   

'********************************* 单点交叉 *************************************
'
'过 程 名： Cross
'参    数： Chromosome1 - 参与交叉的染色体1
'           Chromosome2 - 参与交叉的染色体2
'说    明： 单点交叉变异，开始交叉的基因位在函数内产生
'作    者： laviewpbt
'修 改 者： zsc
'
'********************************* 单点交叉 *************************************
Private Sub OnePoint_CrossOver(ByRef Chromosome1 As String, ByRef Chromosome2 As String)
    Dim CrossOverBit As Integer
    Dim StrTemp1 As String, StrTemp2 As String
    CrossOverBit = Int(1 + Rnd * (Len(Chromosome1) - 1))
    StrTemp1 = Mid(Chromosome1, CrossOverBit + 1)
    StrTemp2 = Mid(Chromosome2, CrossOverBit + 1)
    Mid(Chromosome2, CrossOverBit + 1) = StrTemp1
    Mid(Chromosome1, CrossOverBit + 1) = StrTemp2
End Sub
'********************************* 两点交叉 *************************************
'
'过 程 名： Cross
'参    数： Chromosome1 - 参与交叉的染色体1
'           Chromosome2 - 参与交叉的染色体2
'说    明： 两点交叉变异，开始交叉的基因位在函数内产生
'作    者： laviewpbt
'修 改 者： zsc
'
'********************************* 两点交叉 *************************************
Private Sub TwoPoint_CrossOver(ByRef Chromosome1 As String, ByRef Chromosome2 As String)
    Dim Index1 As Integer, Index2 As Integer, Length As Integer, IntTemp As Integer
    Dim StrTemp1 As String, StrTemp2 As String
    Length = Len(Chromosome1)
    Index1 = Int(1 + Rnd * (Length - 1))        '生成第一个交叉点
    Index2 = Int(1 + Rnd * (Length - 1))        '生成第二个交叉点
    If Index2 < Index1 Then
        IntTemp = Index1
        Index1 = Index2
        Index2 = IntTemp
    End If
    Index2 = Index2 - Index1              '避免重复计算
    Index1 = Index1 + 1
    StrTemp1 = Mid(Chromosome1, Index1, Index2)
    StrTemp2 = Mid(Chromosome2, Index1, Index2)
    Mid(Chromosome1, Index1, Index2) = StrTemp2
    Mid(Chromosome2, Index1, Index2) = StrTemp1
End Sub
'********************************* 均匀交叉 *************************************
'
'过 程 名： Cross
'参    数： Chromosome1 - 参与交叉的染色体1
'           Chromosome2 - 参与交叉的染色体2
'说    明： 均匀交叉变异，屏蔽字实际上转换位Rnd > 0.5
'作    者： laviewpbt
'修 改 者： zsc
'
'********************************* 均匀交叉 *************************************
Private Sub Uniform_CrossOver(ByRef Chromosome1 As String, ByRef Chromosome2 As String)
    Dim i As Integer, Length As Integer
    Dim StrTemp1 As String, StrTemp2 As String
    Length = Len(Chromosome1)
    Randomize
    For i = 1 To Length
        If Rnd > 0.5 Then '相当于屏蔽字的这一位为1
            StrTemp1 = Mid(Chromosome1, i, 1)
            StrTemp2 = Mid(Chromosome2, i, 1)
            Mid(Chromosome2, i, 1) = StrTemp1
            Mid(Chromosome1, i, 1) = StrTemp2
        End If
    Next
End Sub
'********************************* 变异 *************************************
'
'过 程 名： Mutation
'参    数： Chromosome - 待变异的染色体
'           GeneBit     - 变异的基因位
'说    明： 基本位突变
'作    者： laviewpbt
'修 改 者： zsc
'
'********************************* 变异 *************************************
Private Sub Mutation(ByRef Chromosome As String, GeneBit As Integer)
    Dim Temp As String
    Temp = Mid(Chromosome, GeneBit, 1)
    If Temp = "1" Then
        Mid(Chromosome, GeneBit, 1) = "0"
    Else
        Mid(Chromosome, GeneBit, 1) = "1"
    End If
End Sub
Private Sub Command1_Click()
'按钮选项控制
Dim optionsetchoice3 As CrossOver
Dim optionsetchoice2 As Selection
Dim optionsetchoice1 As EnCoding
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Set xlApp = New Excel.Application
xlApp.Visible = False
Set xlBook = xlApp.Workbooks.Open("E:\异步电机检查试验合格区\Book1.xls")
Set xlSheet = xlBook.Worksheets(1)
'GA关键参数
allmenber = Val(Text1.Text)
DigitNum = Val(Text2.Text)
Pc = Val(Text3.Text)
Pm = Val(Text4.Text)
MaxIter = Val(Text5.Text)
'技术条件上的保证值
ηe = Val(Text6.Text)
cosΨe = Val(Text7.Text)
MmaxMee = Val(Text8.Text)
MqtMee = Val(Text9.Text)
iqtiee = Val(Text10.Text)
'技术条件上的极限值
ηmin = Val(Text11.Text)
cosΨmin = Val(Text12.Text)
MmaxMemin = Val(Text13.Text)
MqtMemin = Val(Text14.Text)
iqtiemin = Val(Text15.Text)
'型式试验值
η = Val(Text16.Text)
cosΨ = Val(Text17.Text)
MmaxMe = Val(Text18.Text)
MqtMe = Val(Text19.Text)
iqtie = Val(Text20.Text)
i1d = Val(Text21.Text)
r1 = Val(Text22.Text)
Pt2 = Val(Text23.Text)
P2 = Val(Text24.Text)
Pt1 = Val(Text25.Text)
'电机额定值
Pe = Val(Text26.Text)
Ue = Val(Text27.Text)
ne = Val(Text28.Text)
'出厂试验数据
'Uo = Val(Text29.Text)
'io = Val(Text30.Text)
'Po = Val(Text31.Text)
'Ud = Val(Text32.Text)
'id = Val(Text33.Text)
'Pd = Val(Text34.Text)
'rd = Val(Text35.Text)
taishu = Val(Text36.Text)
ReDim evryhanval1(1 To taishu)
ReDim evryhanval2(1 To taishu)
ReDim evryhanval3(1 To taishu)
ReDim evryhanval4(1 To taishu)
ReDim evryhanval5(1 To taishu)
ReDim evryhanval6(1 To taishu)
If Option1 Then
    optionsetchoice2 = Selection.RouletteWheelSelection
ElseIf Option2 Then
        optionsetchoice2 = Selection.StochasticTourament
ElseIf Option3 Then
        optionsetchoice1 = Selection.RandomLeagueMatches
Else
        optionsetchoice1 = Selection.StochasticUniversalSampleing
End If
            
If Option5 Then
    optionsetchoice1 = EnCoding.Binary
Else
    optionsetchoice2 = EnCoding.Gray
End If

If Option7 Then
    optionsetchoice3 = CrossOver.OnePointCrossOver
ElseIf Option8 Then
    optionsetchoice2 = CrossOver.TwoPointCrossOver
Else
    optionsetchoice2 = CrossOver.UniformCrossOver
End If
For han = 2 To taishu + 1
 Uo = xlSheet.Cells(han, 2)
 io = xlSheet.Cells(han, 3)
 Po = xlSheet.Cells(han, 4)
 Ud = xlSheet.Cells(han, 5)
 id = xlSheet.Cells(han, 6)
 Pd = xlSheet.Cells(han, 7)
 rd = Pd / (1.5 * id ^ 2)
bianlian1(1, 1) = idmin(id)
bianlian1(1, 2) = idmax(id)
evryhanval2(han - 1) = GA(bianlian1, io, id, Ud, , , , , , optionsetchoice1, optionsetchoice2, optionsetchoice3).Cordinate(1)
evryhanval1(han - 1) = idmax(id) * 5 - GA(bianlian1, io, id, Ud, , , , , , optionsetchoice1, optionsetchoice2, optionsetchoice3).Max
evryhanval3(han - 1) = GA(bianlian1, , id, , Pd, , , , 2, optionsetchoice1, optionsetchoice2, optionsetchoice3).Max
evryhanval4(han - 1) = GA(bianlian1, , id, , Pd, , , , 2, optionsetchoice1, optionsetchoice2, optionsetchoice3).Cordinate(1)
bianlian2(1, 1) = 0
bianlian2(1, 2) = evryhanval1(han - 1)
bianlian2(2, 1) = 0
bianlian2(2, 2) = rdmax(rd)
evryhanval6(han - 1) = GA(bianlian2, io, , , , Po, rd, , , optionsetchoice1, optionsetchoice2, optionsetchoice3).Cordinate(2)
evryhanval5(han - 1) = 5 * Pe - GA(bianlian2, io, , , , Po, rd, , , optionsetchoice1, optionsetchoice2, optionsetchoice3).Max
Next
For jishuzhi = 1 To taishu - 1
evryhanval1(jishuzhi + 1) = evryhanval1(jishuzhi) + evryhanval1(jishuzhi + 1)
evryhanval2(jishuzhi + 1) = evryhanval2(jishuzhi) + evryhanval2(jishuzhi + 1)
evryhanval3(jishuzhi + 1) = evryhanval3(jishuzhi) + evryhanval3(jishuzhi + 1)
evryhanval4(jishuzhi + 1) = evryhanval4(jishuzhi) + evryhanval4(jishuzhi + 1)
evryhanval5(jishuzhi + 1) = evryhanval5(jishuzhi) + evryhanval5(jishuzhi + 1)
evryhanval6(jishuzhi + 1) = evryhanval6(jishuzhi) + evryhanval6(jishuzhi + 1)
Next
prioutval = evryhanval1(taishu) / taishu
outputcanshu = evryhanval2(taishu) / taishu
outputcanshu1 = evryhanval3(taishu) / taishu
outputcanshu2 = evryhanval4(taishu) / taishu
prioutval1 = evryhanval5(taishu) / taishu
outputcanshu3 = evryhanval6(taishu) / taishu
'方程1的输出
Picture1.Print "io上限"; prioutval
Picture1.Print "id下限"; outputcanshu
'方程2的输出
Picture1.Print "Pd下限"; outputcanshu1
Picture1.Print "id上限"; outputcanshu2
'方程3的输出
Picture1.Print "Po上限"; prioutval1
Picture1.Print "rd上限"; outputcanshu3
End Sub
Private Sub Form_Load()
Dim i As Integer
N2(0) = 1
    For i = 0 To 29
        N2(i + 1) = N2(i) * 2
    Next
End Sub


