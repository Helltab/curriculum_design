VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "并流多效蒸发系统优化设计（具有冷凝水闪蒸）"
   ClientHeight    =   10650
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   15975
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   20
      Left            =   13200
      TabIndex        =   30
      Text            =   "1.3"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   19
      Left            =   10440
      TabIndex        =   29
      Text            =   "1.0"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   18
      Left            =   7560
      TabIndex        =   28
      Text            =   "0.54"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   17
      Left            =   4560
      TabIndex        =   27
      Text            =   "645"
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   16
      Left            =   2040
      TabIndex        =   26
      Text            =   "0.15"
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   15
      Left            =   11280
      TabIndex        =   25
      Text            =   "0.000146"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   14
      Left            =   11280
      TabIndex        =   24
      Text            =   "0.018"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   13
      Left            =   11280
      TabIndex        =   23
      Text            =   "7200"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   12
      Left            =   11280
      TabIndex        =   22
      Text            =   "0.5"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   11
      Left            =   8640
      TabIndex        =   21
      Text            =   "1.1"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   10
      Left            =   8640
      TabIndex        =   20
      Text            =   "0.6"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   9
      Left            =   8640
      TabIndex        =   19
      Text            =   "1000"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   8
      Left            =   8640
      TabIndex        =   18
      Text            =   "0.98"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   7
      Left            =   5640
      TabIndex        =   17
      Text            =   "3969.45"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   6
      Left            =   5640
      TabIndex        =   16
      Text            =   "4187"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   5
      Left            =   5640
      TabIndex        =   15
      Text            =   "55"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   4
      Left            =   5640
      TabIndex        =   14
      Text            =   "135"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   3
      Left            =   2520
      TabIndex        =   13
      Text            =   "0.45"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   2
      Left            =   2520
      TabIndex        =   12
      Text            =   "65"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   1
      Left            =   2520
      TabIndex        =   11
      Text            =   "0.08"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "显示"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Text            =   "8000"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   12960
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "水的比热容"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   17
      Left            =   3840
      TabIndex        =   45
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "硫酸铵的比热容"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   16
      Left            =   3840
      TabIndex        =   44
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "热利用率"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   15
      Left            =   7080
      TabIndex        =   43
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "水的密度"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   14
      Left            =   7080
      TabIndex        =   42
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "泵的效率"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   13
      Left            =   7080
      TabIndex        =   41
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "附加费用系数  c1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   12
      Left            =   7080
      TabIndex        =   40
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "电费"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   11
      Left            =   10200
      TabIndex        =   39
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "年使用时间   θ"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   10
      Left            =   9960
      TabIndex        =   38
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "回归系数        a"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   9
      Left            =   9960
      TabIndex        =   37
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "回归系数        b"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   8
      Left            =   9960
      TabIndex        =   36
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "年折旧率       Fc"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   7
      Left            =   480
      TabIndex        =   35
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "回归系数        a1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   6
      Left            =   3360
      TabIndex        =   34
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "回归系数       b2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   5
      Left            =   6240
      TabIndex        =   33
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "压力校正系数  f1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   4
      Left            =   8880
      TabIndex        =   32
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "材质校正系数  f2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   3
      Left            =   11760
      TabIndex        =   31
      Top             =   4440
      Width           =   1335
   End
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      Class           =   "Excel.Sheet.8"
      Height          =   4335
      Left            =   330
      OleObjectBlob   =   "优化设计.frx":0000
      TabIndex        =   9
      Top             =   5490
      Width           =   14295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "原料液进料温度 t0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   0
      Left            =   3795
      TabIndex        =   7
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "浓缩液温度     xn"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "加热蒸汽温度  T0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "末效冷凝器的温度  T0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   495
      TabIndex        =   4
      Top             =   2340
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "料液浓度 x0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   495
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "处理量   F0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   495
      TabIndex        =   2
      Top             =   300
      Width           =   1575
   End
   Begin VB.Menu s1 
      Caption         =   "参数设置"
   End
   Begin VB.Menu s2 
      Caption         =   "结果显示"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Wf(20000), Tf1 As Double, Tf2 As Double, I0!, K0, Af, bf, Mf, vw
Dim A(50, 50)
Dim D(50), G(50), W(50), V(50)
Dim a1(50), b1(50), w1(50), x1(50), e(50), t1(50), T(50), K(3), r(50), ct(50), ct1(50), ct3(50), H1(50), h0(50)
Dim sH(50), sA(50), sG, sD
Dim n, m, st, h, c0, sDr, sct, s, sW, Wn, a0, b, bb, tstop, f1, f2, sJ, J1, J2, J3, J4, bean, ff1, ff2

Static Function y(Tf)

For i = 0 To 20
    If Val(Text1(i) & 1) = 1 Then
        MsgBox ("有参数未设置！")
        Exit Function
    End If
Next
n = 3: m = 2 * n
F0 = Text1(0) / 3600: x1(0) = Text1(1): Tk = Text1(2): x1(n) = Text1(3): T(0) = Tf
t1(0) = Text1(5): h = Text1(6): c0 = Text1(7): s = 10 ^ -5
K(1) = 750: K(2) = 680: K(3) = 600
Wn = F0 * (1 - x1(0) / x1(n))
st = T(0) - Tk - 3
For i = 1 To n
    a1(i) = 1: b1(i) = 2 * 10 ^ -5: w1(i) = 2 * 10 ^ -5: r(i) = 2220000
    x1(i) = x1(0) / (1 - i * (1 - x1(0) / x1(n)) / n)
Next i
b1(1) = b1(1)
Text1(4) = T(0)
Do
    For i = 1 To n
         e(i) = Text1(8)
    Next
    For i = 1 To 2 * n
        For j = 1 To 2 * n + 1
            A(i, j) = 0
    Next j, i
                                                           
    For i = 1 To n
      A(i, i) = a1(i) * e(i) - h * b1(i) * e(i)
      A(i, i + 1) = -1
    Next i
    A(1, 1) = a1(1) * e(1)
    For i = 3 To n
      For j = 2 To i - 1
        A(i, j) = -h * b1(i) * e(i)
    Next j, i
    For j = 2 To n + 1
      A(n + 1, j) = 1
    Next j
                                                           
    For i = n + 2 To 2 * n
      For j = 1 To i - (n + 1)
        A(i, j) = h * w1(i - n - 1)
    Next j, i
                                                            
    For i = 2 To n
      j = n + i
      A(i, j) = a1(i) * e(i)
    Next i
                                                            
    For i = n + 2 To 2 * n
      j = i
      A(i, j) = -1
    Next i
                                                          
    For i = 1 To n + 1
      j = 2 * n + 1
      A(i, j) = -F0 * c0 * b1(i) * e(i)
    Next i
    A(n + 1, 2 * n + 1) = Wn
                                            
     For i = 1 To 2 * n
      For j = 2 * n + 1 To i Step -1
        A(i, j) = A(i, j) / A(i, i)
      Next j
      For p = i + 1 To 2 * n
        For j = 2 * n + 1 To i Step -1
          A(p, j) = A(p, j) - A(i, j) * A(p, i)
    Next j, p, i
    For i = 2 * n To 2 Step -1
      For p = i - 1 To 1 Step -1
        For j = 2 * n + 1 To 1 Step -1
          A(p, j) = A(p, j) - A(i, j) * A(p, i)
    Next j, p, i

    
    D(1) = A(1, 2 * n + 1)
    For i = 1 To n
        W(i) = A(i + 1, 2 * n + 1)
    Next
    For i = 1 To n - 1
        G(i) = A(i + n + 1, 2 * n + 1)
    Next i
    For i = 2 To n
        D(i) = W(i - 1) + G(i - 1)
    Next
    sW = 0
    For i = 1 To n - 1
        sW = sW + W(i)
         x1(i) = F0 * x1(0) / (F0 - sW)
    Next
   
    sDr = 0
    For i = 1 To n
        sDr = sDr + (D(i) * r(i)) / K(i)
    Next
    sct = 0
    For i = 1 To n
        ct1(i) = 10.9 * x1(i) + 3.78 * x1(i) ^ 2
        ct3(i) = 1
        sct = sct + ct1(i) + ct3(i)
    Next
    st = T(0) - Tk - sct
    
    For i = 1 To n
        ct(i) = ((D(i) * r(i) / K(i)) * st) / sDr
    Next
      
    
    T(n) = Tk + 1
    For i = n To 2 Step -1
        t1(i) = T(i) + ct1(i)
        T(i - 1) = t1(i) + ct(i) + ct3(i - 1)
    Next
    t1(1) = T(1) + ct1(1)
    For i = 0 To n
        H1(i) = 2474771# + 2410.2 * T(i) - 3.83 * T(i) ^ 2
    Next
    
    For i = 1 To n
        a1(i) = (H1(i - 1) - h * T(i - 1)) / (H1(i) - h * t1(i))
        b1(i) = (t1(i - 1) - t1(i)) / (H1(i) - h * t1(i))
        r(i) = 2466904.92 - 1584.27 * T(i - 1) - 4.93 * T(i - 1) ^ 2
    Next
 
    For i = 1 To n - 1
        w1(i) = (T(i - 1) - T(i)) / (H1(i) - h * T(i))
    Next
    
  
    
    For i = 1 To n
        sA(i) = D(i) * r(i) / K(i) / ct(i)
    Next
    Max = sA(1): Min = sA(n)
    For i = 1 To n
        If sA(i) > Max Then Max = sA(i)
        If sA(i) < Min Then Min = sA(i)
    Next
    
 Loop Until Abs(Max - Min) < s
J1 = 3600# * Text1(13) * D(1) * (Text1(14) + Text1(15) * T(0))
Sum = 0
Tdd = D(1)

For i = 1 To n
    If sA(i) <= 100 Then h0(i) = 1
    If sA(i) > 100 And sA(i) <= 200 Then h0(i) = 1.2
    If sA(i) > 200 And sA(i) <= 400 Then h0(i) = 1.5

    Sum = Sum + 43780 * 1.2 * (0.667 + 0.0287 * sA(i)) * h0(i)
Next
bb = Text1(16)
J2 = bb * Sum
a0 = Text1(17): b = Text1(18): f1 = Text1(19): f2 = Text1(20)
For i = 1 To n - 1
    sD = 0
    For p = 1 To i
        sD = sD + D(p)
    Next p
    sG = 0
    For j = 1 To i - 1
        sG = sG + G(j)
    Next j
    tstop = 300
    pp = 1000
    V(i) = 2 * (sD - sG) * tstop / pp
Next
J3 = 0
For i = 1 To n - 1
    J3 = J3 + bb * a0 * V(i) ^ b * f1 * f2
Next
rk = 2466904.92 - 1584.27 * Tk - 4.93 * (Tk) ^ 2
pp = Text1(9): ep = Text1(10): c1 = Text1(11): c2 = Text1(12)
vw = W(n) * (rk + h * 5) / (pp * h * (Tk - 5 - 35))
J4 = c1 * c2 * 7200 * (21 + 1450 * vw ^ 2) * vw * pp / 102 / ep

sJ = J1 + J2 + J3 + J4
y = sJ

End Function
Function Y2(XXX)
        Y2 = (XXX - 3) ^ 2 + 7
        
End Function


Private Sub Command1_Click()


Af = 100: bf = 180: Wf(1) = 1: Wf(2) = 2: I0 = 1
For I0 = 1 To 100000
    Wf(I0 + 2) = Wf(I0 + 1) + Wf(I0)
    If Wf(I0 + 2) <= (bf - Af) / 0.0001 Then
   
    Else
       
        Tf1 = Af + (bf - Af) * Wf(I0) / Wf(I0 + 2): ff1 = y(Tf1): nf = I0 + 2: Mf = 0: K0 = 1
        Do
            If Mf = 0 Then
                Tf2 = Af + (bf - Af) * Wf(nf - K0) / Wf(nf - K0 + 1)
                ff2 = y(Tf2)
            Else
                Tf1 = Af + (bf - Af) * Wf(nf - K0 - 1) / Wf(nf - K0 + 1)
                ff1 = y(Tf1)
            End If
            If ff1 < ff2 Then
                bf = Tf2: Tf2 = Tf1: ff2 = ff1: Mf = 1
            Else
                Af = Tf1: Tf1 = Tf2: ff1 = ff2: Mf = 0
            End If
            K0 = K0 + 1
        Loop Until K0 = nf - 1
        Exit For
    End If

Next
    
   T0 = (Af + bf) / 2
   y (T0)
   Text1(4) = T0
  End Sub


Private Sub Command2_Click()
Set ob1 = OLE1.object
If bean = 0 Then
    xx = MsgBox("是否清除表格里的内容？", vbOKCancel)
    If xx = 1 Then
        For i = 2 To 4
            For j = 2 To 8
             ob1.activesheet.cells(i, j) = ""
        Next j, i
    End If
Else
    xx = MsgBox("是否清除常用项内容？", vbOKCancel)
    If xx = 1 Then
        For i = 0 To 5
            Text1(i) = ""
        Next
    End If
End If
End Sub

Private Sub Command3_Click()
Set ob = OLE1.object
For i = 1 To n
    ob.activesheet.cells(i + 1, 2) = sA(i)
    ob.activesheet.cells(i + 1, 3) = K(i)
    ob.activesheet.cells(i + 1, 4) = ct(i)
    ob.activesheet.cells(i + 1, 5) = t1(i)
    ob.activesheet.cells(i + 1, 6) = T(i)
    ob.activesheet.cells(i + 1, 7) = W(i)
    ob.activesheet.cells(i + 1, 8) = x1(i)
    ob.activesheet.cells(i + 1, 9) = D(i)
    ob.activesheet.cells(i + 1, 10) = G(i)
Next
ob.activesheet.cells(5, 4) = J1
ob.activesheet.cells(6, 4) = J2
ob.activesheet.cells(7, 4) = J3
ob.activesheet.cells(8, 4) = J4
ob.activesheet.cells(9, 4) = sJ
End Sub

Private Sub Form_Load()
For i = 0 To 20
    Text1(i).Enabled = False
Next
OLE1.Visible = False
OLE1.Enabled = False
Command3.Visible = False
Command1.Enabled = False
Command2.Enabled = False
s2.Enabled = False
End Sub

Private Sub s1_Click()
For i = 0 To 20
    Text1(i).Enabled = True
Next
s1.Enabled = False
s2.Enabled = True
Command1.Enabled = True
Command3.Enabled = False
Command2.Enabled = True
bean = 1
End Sub

Private Sub s2_Click()
For i = 0 To 20
    Text1(i).Enabled = False
Next
s1.Enabled = True
s2.Enabled = False
OLE1.Visible = True
Command3.Visible = True
Command3.Enabled = True
bean = 0
End Sub

