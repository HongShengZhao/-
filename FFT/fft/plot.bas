Attribute VB_Name = "plot"

'<CSCM>
'--------------------------------------------------------------------------------
' 工程名称    : fft
' 名称        : plotip
' 类型        : Sub
' 描述        : 显示采样信号波形
' 创建者      : Hotsun
' 创建时间    : 2011-11-16-14:16:18
'
' 修改者      :
' 修改说明    :
'
' 修改时间    :
' 参数        :
' 返回        :
' 引用全局变量:
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub plotip()

    On Error Resume Next

    Dim maxvalue, oldval As Double

    mainfrm.ipplot.Line (0, 1450 / Screen.TwipsPerPixelY)-(10000, 1450 / Screen.TwipsPerPixelY), vbRed

    For cnt = 0 To 512

        If (maxvalue < FuncGen.Samples(cnt)) Then
            maxvalue = FuncGen.Samples(cnt)

        End If

    Next cnt

    oldval = 2900 / (2 * Screen.TwipsPerPixelY)

    For cnt = 1 To 1024
        'mainfrm.ipplot.PSet (cnt, (2900 / (2 * Screen.TwipsPerPixelY)) - (FuncGen.Samples(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue)))
        mainfrm.ipplot.Line (cnt - 1, oldval)-(cnt, (2900 / (2 * Screen.TwipsPerPixelY)) - (FuncGen.Samples(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue))), vbYellow
        oldval = (2900 / (2 * Screen.TwipsPerPixelY)) - (FuncGen.Samples(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue))
    Next cnt

End Sub

