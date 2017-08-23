Attribute VB_Name = "Module1"
Option Explicit

'*ģ��********************************************************
'FFT0 �����±���0��ʼ FFT1 �����±���1��ʼ
'AR() ����ʵ��         AI() �����鲿
'N ���ݵ�����Ϊ2����������
'NI �任���� 1Ϊ���任��-1Ϊ���任
'***************************************************************

Public Const Pi = 3.1415926

'<CSCM>
'--------------------------------------------------------------------------------
' ��������    : Project1
' ����        : FFT0
' ����        : Function
' ����        : ���ٸ���Ҷ�任,����������±���0��ʼ
' ������      : Hotsun
' ����ʱ��    : 2011-11-16-11:20:05
'
' �޸���      :
' �޸�˵��    :
'
' �޸�ʱ��    :
' ����        : AR() (Double) ʵ��
'             AI() (Double) �鲿
'             N (Integer) ���鳤��,����Ϊ2��������
'             ni (Integer) 1���任,-1��任
' ����        :
' ����ȫ�ֱ���:
'--------------------------------------------------------------------------------
'</CSCM>
Public Function FFT0(AR() As Double, AI() As Double, N As Integer, ni As Integer)
    Dim i As Integer, j As Integer, k As Integer, L As Integer, M As Integer
    Dim IP As Integer, LE As Integer
    Dim L1 As Integer, N1 As Integer, N2 As Integer
    Dim SN As Double, TR As Double, TI As Double, WR As Double, WI As Double
    Dim UR As Double, UI As Double, US As Double
    M = NTOM(N)
   ' Debug.Print N, M
    N2 = N / 2 '�԰�
    N1 = N - 1
    SN = ni
    j = 1
    '���е��α任
    For i = 1 To N1
        If i < j Then
            TR = AR(j - 1)
            AR(j - 1) = AR(i - 1)
            AR(i - 1) = TR
            TI = AI(j - 1)
            AI(j - 1) = AI(i - 1)
            AI(i - 1) = TI
        End If
        k = N2
        While (k < j)
            j = j - k
            k = k / 2
        Wend
        j = j + k
    Next i
    
    '���м���
    For L = 1 To M
        LE = 2 ^ L
        L1 = LE / 2
        UR = 1#
        UI = 0#
        WR = Cos(Pi / L1)
        WI = SN * Sin(Pi / L1)
        For j = 1 To L1
            For i = j To N Step LE
                IP = i + L1
                TR = AR(IP - 1) * UR - AI(IP - 1) * UI
                TI = AI(IP - 1) * UR + AR(IP - 1) * UI
                AR(IP - 1) = AR(i - 1) - TR
                AI(IP - 1) = AI(i - 1) - TI
                AR(i - 1) = AR(i - 1) + TR
                AI(i - 1) = AI(i - 1) + TI
            Next i
            US = UR
            UR = US * WR - UI * WI
            UI = UI * WR + US * WI
        Next j
    Next L
    
    If SN <> -1 Then '���任
        For i = 1 To N
            AR(i - 1) = AR(i - 1) / N
            AI(i - 1) = AI(i - 1) / N
        Next i
    End If
End Function
'<CSCM>
'--------------------------------------------------------------------------------
' ��������    : Project1
' ����        : FFT1
' ����        : Function
' ����        : ���ٸ���Ҷ�任,����������±���1��ʼ
' ������      : Hotsun
' ����ʱ��    : 2011-11-16-11:20:05
'
' �޸���      :
' �޸�˵��    :
'
' �޸�ʱ��    :
' ����        : AR() (Double) ʵ��
'             AI() (Double) �鲿
'             N (Integer) ���鳤��,����Ϊ2��������
'             ni (Integer) 1���任,-1��任
' ����        :
' ����ȫ�ֱ���:
'--------------------------------------------------------------------------------
'</CSCM>
Public Function FFT1(AR() As Double, AI() As Double, N As Integer, ni As Integer)
    Dim i As Integer, j As Integer, k As Integer, L As Integer, M As Integer
    Dim IP As Integer, LE As Integer
    Dim L1 As Integer, N1 As Integer, N2 As Integer
    Dim SN As Double, TR As Double, TI As Double, WR As Double, WI As Double
    Dim UR As Double, UI As Double, US As Double
    M = NTOM(N)
    N2 = N / 2
    N1 = N - 1
    SN = ni
    j = 1
    For i = 1 To N1
        If i < j Then
            TR = AR(j)
            AR(j) = AR(i)
            AR(i) = TR
            TI = AI(j)
            AI(j) = AI(i)
            AI(i) = TI
        End If
        k = N2
        While (k < j)
            j = j - k
            k = k / 2
        Wend
        j = j + k
    Next i
    For L = 1 To M
        LE = 2 ^ L
        L1 = LE / 2
        UR = 1#
        UI = 0#
        WR = Cos(Pi / L1)
        WI = SN * Sin(Pi / L1)
        For j = 1 To L1
            For i = j To N Step LE
                IP = i + L1
                TR = AR(IP) * UR - AI(IP) * UI
                TI = AI(IP) * UR + AR(IP) * UI
                AR(IP) = AR(i) - TR
                AI(IP) = AI(i) - TI
                AR(i) = AR(i) + TR
                AI(i) = AI(i) + TI
            Next i
            US = UR
            UR = US * WR - UI * WI
            UI = UI * WR + US * WI
        Next j
    Next L
    If SN <> -1 Then
        For i = 1 To N
            AR(i) = AR(i) / N
            AI(i) = AI(i) / N
        Next i
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' ��������    : Project1
' ����        : NTOM
' ����        : Function
' ����        : �����鳤�ȵ�2������
' ������      : Hotsun
' ����ʱ��    : 2011-11-16-11:40:07
'
' �޸���      :
' �޸�˵��    :
'
' �޸�ʱ��    :
' ����        : N (Integer)���鳤��,2^7=128,2^8=256
' ����        :
' ����ȫ�ֱ���:
'--------------------------------------------------------------------------------
'</CSCM>
Private Function NTOM(N As Integer) As Integer
    Dim ND As Double
    ND = N
    NTOM = 0
    While (ND > 1)
        ND = ND / 2
        NTOM = NTOM + 1
    Wend
End Function
