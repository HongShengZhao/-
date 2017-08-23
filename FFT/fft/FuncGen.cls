VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FunctionGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'<CSCC>
'--------------------------------------------------------------------------------
' 工程名称   : prjXYDLL
' 模块名称   : ExToolTip
' 功能描述   : 波形信号产生模块
'
'
' 创建者     : HotSun
' 创建时间   : 2011-11-3-0:06:22
'
' 公司名称   :
' 网址       :
' 作者E-Mail : HotSun@vip.126.com
'
' 修改说明   :
'
'
'
' 修改者     : HotSun
' 修改时间   :
'
' 版本       : V6.0.0
'********************************对外接口全局变量**************************************
'*****************************************************************************
'--------------------------------过程(Sub)-------------------------------------
'1、
'2、
'3、
'4、
'5、
'6、
'--------------------------------函数(Function)---------------------------------
'1、
'2、
'3、
'4、
'5、
'6、
'--------------------------------属性(Property)--------------------------------
'1、
'2、
'3、
'4、
'5、
'6、
'7
'8
'9
'10
'11
'12
'13
'--------------------------------全局变量Variable、Type、Emum--------------------------------
'1、
'2、
'3、
'4、
'5、
'6、
''*************************************************************************
'==================================模块常数或变量========================
'----------------------------------常数(Const)--------------------------
'

'--------------------------------------------------------------------------------
'
'----------------------------------私有变量------------------------------
'
'
'
'========================================================================
'
'----------------------------------------Windows API接口定义------------------------
'
'
'========================================================================
'</CSCC>

Dim tmFunGen    As Timer

Dim sample_rate As Long

Dim samp_buf(n) As Double

Const nsamples = n

Dim ampl As Double

'采样频率
Public Property Get SamplingRate() As Variant
    Sampling_Rate = sample_rate

End Property

Public Property Let SamplingRate(vNewValue As Variant)
    sample_rate = vNewValue

End Property

'读取采样缓冲区内数据
Public Property Get Samples() As Variant
    Samples = samp_buf

End Property

Public Property Get NumSamples() As Variant
    NumSamples = nsamples

End Property

'Public Property Let NumSamples(ByVal vNewValue As Variant)
'nsamples = vNewValue
'End Property

'<CSCM>
'--------------------------------------------------------------------------------
' 工程名称    : fft
' 名称        : GenSine
' 类型        : Function
' 描述        : 产生正弦信号
' 创建者      : Hotsun
' 创建时间    : 2011-11-16-14:06:54
'
' 修改者      :
' 修改说明    :
'
' 修改时间    :
' 参数        : Frequency (Long) 信号的频率
' 返回        :
' 引用全局变量:
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GenSine(Frequency As Long)
    Dim n As Single
    n = 0
    For cnt = 1 To nsamples
        samp_buf(cnt) = samp_buf(cnt) + (ampl * Sin(2 * 3.141592658 * Frequency * (cnt / sample_rate)))
        If samp_buf(cnt) > n Then
          n = samp_buf(cnt)
        End If
    Next cnt
   'Debug.Print n
End Function

Public Sub Clear()

    For cnt = 1 To nsamples
        samp_buf(cnt) = 0
    Next cnt

End Sub

'振幅
Public Property Get amplitude() As Variant
Attribute amplitude.VB_Description = "Sets Amplitude of the wave"
    amplitude = ampl

End Property

Public Property Let amplitude(ByVal vNewValue As Variant)
    ampl = vNewValue

End Property

'<CSCM>
'--------------------------------------------------------------------------------
' 工程名称    : fft
' 名称        : GenCos
' 类型        : Function
' 描述        : 产生余弦信号
' 创建者      : Hotsun
' 创建时间    : 2011-11-16-14:09:23
'
' 修改者      :
' 修改说明    :
'
' 修改时间    :
' 参数        : Frequency (Long)
' 返回        :
' 引用全局变量:
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GenCos(Frequency As Long)

    For cnt = 1 To nsamples
        samp_buf(cnt) = samp_buf(cnt) + (ampl * Cos(2 * 3.141592658 * Frequency * (cnt / sample_rate)))
    Next cnt

End Function

Public Sub GenImpulse(position As Integer)
    samp_buf(position) = samp_buf(position) + ampl

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' 工程名称    : fft
' 名称        : GenSquare
' 类型        : Sub
' 描述        : 产生方波信号
' 创建者      : Hotsun
' 创建时间    : 2011-11-16-14:10:18
'
' 修改者      :
' 修改说明    :
'
' 修改时间    :
' 参数        : Frequency (Long)
' 返回        :
' 引用全局变量:
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub GenSquare(Frequency As Long)

    On Error Resume Next

    polarity = 0
    ontime = n / (2 * Frequency)
   
    For cnt = 0 To (n / ontime)
   
        For cnt2 = 0 To ontime

            If (polarity = 0) Then
                samp_buf((cnt * ontime) + cnt2) = ampl
            Else
                samp_buf((cnt * ontime) + cnt2) = -ampl

            End If

        Next cnt2

        If polarity = 0 Then
            polarity = 1
        Else
            polarity = 0

        End If

    Next cnt
    
End Sub

Public Sub ApplyHamming()

    For cnt = 1 To n
        samp_buf(cnt) = samp_buf(cnt) * (0.54 - (0.46 * Cos(2 * 3.14159265 * cnt / 512)))
    Next cnt

End Sub

Public Sub ApplyHanning()

    For cnt = 1 To n
        samp_buf(cnt) = samp_buf(cnt) * (0.5 - (0.5 * Cos(2 * 3.14159265 * cnt / 512)))
    Next cnt

End Sub

Public Sub ApplyBlackman()

    For cnt = 1 To n
        samp_buf(cnt) = samp_buf(cnt) * (0.42 - (0.5 * Cos(2 * 3.14159265 * cnt / 512)) + (0.08 * Cos(4 * 3.14159265 * cnt / 512)))
    Next cnt

End Sub
