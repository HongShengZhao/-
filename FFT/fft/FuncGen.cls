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
' ��������   : prjXYDLL
' ģ������   : ExToolTip
' ��������   : �����źŲ���ģ��
'
'
' ������     : HotSun
' ����ʱ��   : 2011-11-3-0:06:22
'
' ��˾����   :
' ��ַ       :
' ����E-Mail : HotSun@vip.126.com
'
' �޸�˵��   :
'
'
'
' �޸���     : HotSun
' �޸�ʱ��   :
'
' �汾       : V6.0.0
'********************************����ӿ�ȫ�ֱ���**************************************
'*****************************************************************************
'--------------------------------����(Sub)-------------------------------------
'1��
'2��
'3��
'4��
'5��
'6��
'--------------------------------����(Function)---------------------------------
'1��
'2��
'3��
'4��
'5��
'6��
'--------------------------------����(Property)--------------------------------
'1��
'2��
'3��
'4��
'5��
'6��
'7
'8
'9
'10
'11
'12
'13
'--------------------------------ȫ�ֱ���Variable��Type��Emum--------------------------------
'1��
'2��
'3��
'4��
'5��
'6��
''*************************************************************************
'==================================ģ�鳣�������========================
'----------------------------------����(Const)--------------------------
'

'--------------------------------------------------------------------------------
'
'----------------------------------˽�б���------------------------------
'
'
'
'========================================================================
'
'----------------------------------------Windows API�ӿڶ���------------------------
'
'
'========================================================================
'</CSCC>

Dim tmFunGen    As Timer

Dim sample_rate As Long

Dim samp_buf(n) As Double

Const nsamples = n

Dim ampl As Double

'����Ƶ��
Public Property Get SamplingRate() As Variant
    Sampling_Rate = sample_rate

End Property

Public Property Let SamplingRate(vNewValue As Variant)
    sample_rate = vNewValue

End Property

'��ȡ����������������
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
' ��������    : fft
' ����        : GenSine
' ����        : Function
' ����        : ���������ź�
' ������      : Hotsun
' ����ʱ��    : 2011-11-16-14:06:54
'
' �޸���      :
' �޸�˵��    :
'
' �޸�ʱ��    :
' ����        : Frequency (Long) �źŵ�Ƶ��
' ����        :
' ����ȫ�ֱ���:
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

'���
Public Property Get amplitude() As Variant
Attribute amplitude.VB_Description = "Sets Amplitude of the wave"
    amplitude = ampl

End Property

Public Property Let amplitude(ByVal vNewValue As Variant)
    ampl = vNewValue

End Property

'<CSCM>
'--------------------------------------------------------------------------------
' ��������    : fft
' ����        : GenCos
' ����        : Function
' ����        : ���������ź�
' ������      : Hotsun
' ����ʱ��    : 2011-11-16-14:09:23
'
' �޸���      :
' �޸�˵��    :
'
' �޸�ʱ��    :
' ����        : Frequency (Long)
' ����        :
' ����ȫ�ֱ���:
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
' ��������    : fft
' ����        : GenSquare
' ����        : Sub
' ����        : ���������ź�
' ������      : Hotsun
' ����ʱ��    : 2011-11-16-14:10:18
'
' �޸���      :
' �޸�˵��    :
'
' �޸�ʱ��    :
' ����        : Frequency (Long)
' ����        :
' ����ȫ�ֱ���:
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