Attribute VB_Name = "plot"

'<CSCM>
'--------------------------------------------------------------------------------
' ��������    : fft
' ����        : plotip
' ����        : Sub
' ����        : ��ʾ�����źŲ���
' ������      : Hotsun
' ����ʱ��    : 2011-11-16-14:16:18
'
' �޸���      :
' �޸�˵��    :
'
' �޸�ʱ��    :
' ����        :
' ����        :
' ����ȫ�ֱ���:
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

