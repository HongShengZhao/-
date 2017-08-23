Imports System
Imports System.Math
Module Module1
    Const N As Integer = 16
    Const M As Integer = 4
    Structure Complex
        Dim real As Double
        Dim img As Double
    End Structure
    Function reverse(ByVal p1 As Integer) As Integer
        Dim i As Integer
        Dim b = 0
        Dim p2 = 0
        For i = 1 To M
            p2 = Floor(p1 / 2)
            b = b * 2 + (p1 - 2 * p2)
            p1 = p2
        Next i
        Return b
    End Function
    '复数相乘
    Function comp_multy(ByVal a As Complex, ByVal b As Complex) As Complex
        Dim c As Complex
        c.real = a.real * b.real - a.img * b.img
        c.img = a.real * b.img + a.img * b.real
        Return c
    End Function
    '复数相加
    Function comp_add(ByVal a As Complex, ByVal b As Complex) As Complex
        Dim c As Complex
        c.real = a.real + b.real
        c.img = a.img + b.img
        Return c
    End Function
    '复数相减
    Function comp_minus(ByVal a As Complex, ByVal b As Complex) As Complex
        Dim c As Complex
        c.real = a.real - b.real
        c.img = a.img - b.img
        Return c
    End Function
    'comp_adjoint复数共轭
    Function comp_adjoint(ByVal a As Complex) As Complex
        Dim c As Complex
        c.real = a.real
        c.img = -a.img
        Return c
    End Function
    'FFT蝶形算法
    'xt为原数据,cur_x为最后所求的数据,nex_x为中间变量
    Sub FFT(ByVal xt() As Complex, ByVal cur_x() As Complex, ByVal nex_x() As Complex)
        Dim i, j, k As Integer
        Dim t, wn As Integer
        Dim x As Double
        Dim W, sec As Complex
        '调整数据顺序
        For i = 0 To N - 1
            j = reverse(i) '反向进位法
            Console.WriteLine("{0}, {1}", i, j)
            cur_x(j) = xt(i)
        Next

        'FFT
        For i = 0 To M - 1
            t = 1 << i '间隔
            wn = t * 2 '计算W变量
            For j = 0 To N / wn - 1
                For k = 0 To t - 1 '连续加号个数
                    x = 2 * PI * k / wn
                    W.real = Cos(x)
                    W.img = -Sin(x)
                    sec = comp_multy(W, cur_x(j * wn + t + k))
                    nex_x(j * wn + k) = comp_add(cur_x(j * wn + k), sec)
                    nex_x(j * wn + t + k) = comp_minus(cur_x(j * wn + k), sec)
                Next
            Next
            For j = 0 To N - 1
                cur_x(j) = nex_x(j)
            Next
        Next
    End Sub
    'IFFT:调用FFT
    'xt为原数据,cur_x为最后所求的数据,nex_x为中间变量
    Sub IFFT(ByVal xt() As Complex, ByVal cur_x() As Complex, ByVal nex_x() As Complex)
        For i As Integer = 0 To N - 1
            nex_x(i) = comp_adjoint(xt(i))
        Next
        FFT(nex_x, cur_x, nex_x)
        For i = 0 To N - 1
            cur_x(i) = comp_adjoint(cur_x(i))
            cur_x(i).real /= N
            cur_x(i).img /= N
        Next
    End Sub
    Sub Main()
        Dim xt(N - 1), cur_x(N - 1), nex_x(N - 1) As Complex
        'comp_num sec'蝶形运算的第二部分
        '初始化原始信号数据
        For i As Integer = 0 To N - 1
            xt(i).real = Cos(2 * PI * i / 100)
            xt(i).img = 0.0
        Next
        Console.WriteLine("原数据：")
        For i = 0 To N - 1
            If (xt(i).img >= 0) Then
                Console.WriteLine("{0}+{1}i", xt(i).real, xt(i).img)
            Else
                Console.WriteLine("{0}{1}i", xt(i).real, xt(i).img)
            End If
        Next
        'FFT变换
        FFT(xt, cur_x, nex_x)
        Console.WriteLine("FFT结果：")
        For i = 0 To N - 1
            If (cur_x(i).img >= 0) Then
                Console.WriteLine("{0}+{1}i", cur_x(i).real, cur_x(i).img)
            Else
                Console.WriteLine("{0}{1}i", cur_x(i).real, cur_x(i).img)
            End If

            '为IFFT铺垫,保存
            nex_x(i) = cur_x(i)
        Next
        'FFT逆变换
        IFFT(nex_x, cur_x, nex_x)
        Console.WriteLine("IFFT还原结果")
        For i = 0 To N - 1
            If (cur_x(i).img >= 0) Then
                Console.WriteLine("{0}+{1}i", cur_x(i).real, cur_x(i).img)
            Else
                Console.WriteLine("{0}{1}i", cur_x(i).real, cur_x(i).img)
            End If
        Next
        Console.ReadLine()
    End Sub

End Module

