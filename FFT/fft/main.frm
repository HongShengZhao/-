VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form mainfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fast Fourier Transform (Chap.12)"
   ClientHeight    =   9030
   ClientLeft      =   1725
   ClientTop       =   1485
   ClientWidth     =   17565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507.589
   ScaleMode       =   0  'User
   ScaleWidth      =   744.911
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   16650
      TabIndex        =   40
      Top             =   3720
      Width           =   1185
   End
   Begin VB.ListBox List2 
      Height          =   8835
      Left            =   14250
      TabIndex        =   39
      Top             =   3390
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   8835
      Left            =   11910
      TabIndex        =   38
      Top             =   3360
      Width           =   2265
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   210
      Left            =   10560
      TabIndex        =   35
      Top             =   15
      Width           =   1125
   End
   Begin VB.Frame Frame5 
      Caption         =   "Output"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   120
      TabIndex        =   18
      Top             =   5730
      Width           =   11595
      Begin VB.PictureBox output 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   2900
         Left            =   90
         ScaleHeight     =   191
         ScaleMode       =   0  'User
         ScaleWidth      =   250
         TabIndex        =   19
         Top             =   240
         Width           =   11415
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fast Fourier Transform"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   120
      TabIndex        =   14
      Top             =   3375
      Width           =   11595
      Begin VB.TextBox txtMax 
         Height          =   375
         Left            =   420
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txtResult 
         Height          =   825
         Left            =   300
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   1380
         Width           =   2895
      End
      Begin VB.Frame Frame7 
         Caption         =   "Window"
         Height          =   1845
         Left            =   3360
         TabIndex        =   30
         Top             =   285
         Width           =   3390
         Begin VB.OptionButton Optblack 
            Caption         =   "Blackman window"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   34
            Top             =   1515
            Width           =   2955
         End
         Begin VB.OptionButton Opthann 
            Caption         =   "Hanning window"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   33
            Top             =   1120
            Width           =   2820
         End
         Begin VB.OptionButton Opthamm 
            Caption         =   "Hamming window"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   315
            TabIndex        =   32
            Top             =   695
            Width           =   2475
         End
         Begin VB.OptionButton Optnowin 
            Caption         =   "No window"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   31
            Top             =   360
            Value           =   -1  'True
            Width           =   2670
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Output Display"
         Height          =   1860
         Left            =   7230
         TabIndex        =   22
         Top             =   270
         Width           =   4125
         Begin VB.CommandButton cmdDisplayOutPut 
            Caption         =   "Display"
            Height          =   525
            Left            =   2265
            TabIndex        =   29
            Top             =   390
            Width           =   1665
         End
         Begin VB.CommandButton cmdClearOutput 
            Caption         =   "Clear"
            Height          =   510
            Left            =   2250
            TabIndex        =   28
            Top             =   1140
            Width           =   1695
         End
         Begin VB.OptionButton Optreal 
            Caption         =   "实部"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   27
            Top             =   300
            Width           =   1800
         End
         Begin VB.OptionButton Optpower 
            Caption         =   "Power spectrum"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   135
            TabIndex        =   26
            Top             =   1575
            Width           =   2115
         End
         Begin VB.OptionButton Optphase 
            Caption         =   "Phase"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   135
            TabIndex        =   25
            Top             =   1275
            Width           =   1890
         End
         Begin VB.OptionButton Optmag 
            Caption         =   "magnitude"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   135
            TabIndex        =   24
            Top             =   930
            Value           =   -1  'True
            Width           =   2145
         End
         Begin VB.OptionButton Optimg 
            Caption         =   "Ima"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   135
            TabIndex        =   23
            Top             =   562
            Width           =   2145
         End
      End
      Begin VB.CommandButton cmdTransform 
         Caption         =   "Transform"
         Height          =   495
         Left            =   1950
         TabIndex        =   21
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   720
         TabIndex        =   20
         Top             =   990
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17055
      Begin MSComctlLib.Slider Slideramp 
         Height          =   210
         Left            =   45
         TabIndex        =   13
         Top             =   2460
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   370
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CommandButton btn_generate 
         Caption         =   "Generate"
         Height          =   315
         Left            =   1845
         TabIndex        =   12
         Top             =   2760
         Width           =   1770
      End
      Begin VB.CommandButton btn_clr 
         Caption         =   "Clear"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1650
      End
      Begin VB.PictureBox ipplot 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   2900
         Left            =   3750
         ScaleHeight     =   191
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   880
         TabIndex        =   10
         Top             =   210
         Width           =   13230
      End
      Begin VB.Frame Frame3 
         Caption         =   "Parameters"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   1845
         TabIndex        =   5
         Top             =   315
         Width           =   1770
         Begin VB.TextBox txt_impulsepos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   90
            TabIndex        =   16
            Text            =   "10"
            Top             =   1665
            Width           =   870
         End
         Begin VB.TextBox txtfreq 
            Height          =   300
            Left            =   90
            TabIndex        =   9
            Text            =   "10"
            Top             =   1020
            Width           =   765
         End
         Begin VB.TextBox txtsamp 
            Height          =   300
            Left            =   105
            TabIndex        =   8
            Text            =   "1000"
            Top             =   510
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Position"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   1365
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Frequency(Hz)"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   7
            Top             =   795
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Sampling Rate"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   6
            Top             =   270
            Width           =   1635
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Wave"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   120
         TabIndex        =   1
         Top             =   315
         Width           =   1665
         Begin VB.OptionButton opt_impulse 
            Caption         =   "Impulse"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   61
            TabIndex        =   15
            Top             =   1740
            Width           =   1230
         End
         Begin VB.OptionButton Opt_sqr 
            Caption         =   "Square"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   61
            TabIndex        =   4
            Top             =   1260
            Width           =   1170
         End
         Begin VB.OptionButton opt_cos 
            Caption         =   "Cos"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   61
            TabIndex        =   3
            Top             =   810
            Width           =   1080
         End
         Begin VB.OptionButton opt_sin 
            Caption         =   "Sin"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   61
            TabIndex        =   2
            Top             =   300
            Value           =   -1  'True
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const NX = 8192

Dim REX(NX) As Double  'REX[ ] holds the real part of the frequency domain

Dim IMX(NX) As Double  'IMX[ ] holds the imaginary part of the frequency domain
 
'<CSCM>
'--------------------------------------------------------------------------------
' 工程名称    : fft
' 名称        : fft
' 类型        : Sub
' 描述        : 快速傅里叶变换
' 创建者      : Hotsun
' 创建时间    : 2011-11-16-14:25:02
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
Public Sub fft()

        Pi = 3.14159265 'Set constants

1000    'THE FAST FOURIER TRANSFORM
        'copyright ?1997-1999 by California Technical Publishing
        'published with  permission from Steven W Smith, www.dspguide.com
        'GUI by logix4u , www.logix4u.net
        'modified by logix4u, www.logix4.net
1010    'Upon entry, N% contains the number of points in the DFT, REX[ ] and
1020    'IMX[ ] contain the real and imaginary parts of the input. Upon return,
1030    'REX[ ] and IMX[ ] contain the DFT output. All signals run from 0 to N%-1.
1060    NM1% = N% - 1
1070    ND2% = N% / 2
1080    M% = CInt(Log(N%) / Log(2))
1090    j% = ND2%

1100    '
1110    For i% = 1 To N% - 2 'Bit reversal sorting

1120        If i% >= j% Then GoTo 1190
1130        TR = REX(j%)
1140        TI = IMX(j%)
1150        REX(j%) = REX(i%)
1160        IMX(j%) = IMX(i%)
1170        REX(i%) = TR
1180        IMX(i%) = TI
1190        k% = ND2%

1200        If k% > j% Then GoTo 1240
1210        j% = j% - k%
1220        k% = k% / 2
1230        GoTo 1200
1240        j% = j% + k%
1250    Next i%

1260    '
1270    For L% = 1 To M% 'Loop for each stage
1280        LE% = CInt(2 ^ L%)
1290        LE2% = LE% / 2
1300        UR = 1
1310        UI = 0
1320        SR = Cos(Pi / LE2%) 'Calculate sine & cosine values
1330        SI = -Sin(Pi / LE2%)

1340        For j% = 1 To LE2% 'Loop for each sub DFT
1350            JM1% = j% - 1

1360            For i% = JM1% To NM1% Step LE% 'Loop for each butterfly
1370                IP% = i% + LE2%
1380                TR = REX(IP%) * UR - IMX(IP%) * UI 'Butterfly calculation
1390                TI = REX(IP%) * UI + IMX(IP%) * UR
1400                REX(IP%) = REX(i%) - TR
1410                IMX(IP%) = IMX(i%) - TI
1420                REX(i%) = REX(i%) + TR
1430                IMX(i%) = IMX(i%) + TI
1440            Next i%

1450            TR = UR
1460            UR = TR * SR - UI * SI
1470            UI = TR * SI + UI * SR
1480        Next j%
1490    Next L%

1500    '
End Sub

Private Sub btn_clr_Click()
    ipplot.Cls
    FuncGen.Clear

End Sub

Private Sub btn_generate_Click()

    FuncGen.SamplingRate = CInt(txtsamp.Text)
    FuncGen.amplitude = Slideramp.Value   '振幅

    If opt_sin.Value = True Then
        FuncGen.GenSine CInt(txtfreq.Text)

    End If

    If opt_cos.Value = True Then
        FuncGen.GenCos CInt(txtfreq.Text)

    End If

    If opt_impulse.Value = True Then
        FuncGen.GenImpulse Val(txt_impulsepos.Text)

    End If

    If Opt_sqr.Value = True Then
    
        FuncGen.GenSquare CInt(txtfreq.Text)
 
    End If

    ipplot.Cls
    plotip

End Sub

Private Sub cmdClearOutput_Click()
    output.Cls

End Sub

Private Sub cmdDisplayOutPut_Click()

    On Error Resume Next

    Const N_POINT = N / 2 '

    If Optreal.Value = True Then '显示实部
        output.ForeColor = vbRed

        For cnt = 0 To N_POINT
            outputarray(cnt) = REX(cnt)
        Next cnt

    End If

    If Optimg.Value = True Then '显示虚部
        output.ForeColor = vbWhite

        For cnt = 0 To N_POINT
            outputarray(cnt) = IMX(cnt)
        Next cnt

    End If

    If Optmag.Value = True Then '显示震幅
        output.ForeColor = vbYellow

        For cnt = 0 To N_POINT
            outputarray(cnt) = Sqr((IMX(cnt) * IMX(cnt)) + (REX(cnt) * REX(cnt)))
        Next cnt

    End If

    If Optphase.Value = True Then '显示相位
        output.ForeColor = vbCyan

        For cnt = 0 To N_POINT
            outputarray(cnt) = Atn(IMX(cnt) / REX(cnt))
        Next cnt

    End If

    If Optpower.Value = True Then '功率谱
        output.ForeColor = vbGreen

        For cnt = 0 To N_POINT
            outputarray(cnt) = IMX(cnt) '(IMX(cnt) * IMX(cnt)) + (REX(cnt) * REX(cnt))
        Next cnt

    End If

    '进行分析
    Dim dblTmp   As Double

    Dim maxvalue As Double

    Dim nIndex   As Integer
   
    maxvalue = 0

    For cnt = 0 To N_POINT '找最大值

        If (maxvalue < Abs(outputarray(cnt))) Then
            maxvalue = Abs(outputarray(cnt))
            
            nIndex = cnt

            '频率为，最大值的点位*采样频率/采样点数
        End If

    Next cnt
    
    
    txtResult.Text = nIndex * (Val(txtsamp.Text) / N) & "Hz  ,Max=" & maxvalue
    dblTmp = Sqr((IMX(nIndex) * IMX(nIndex)) + (REX(nIndex) * REX(nIndex)))  '幅值
            
    txtResult.Text = txtResult.Text & ",幅值=" & Format(dblTmp * 2, "#0.0####")
    dblTmp = 180 * outputarray(nIndex) / Pi 'Atan2(IMX(nIndex), REX(nIndex))
     txtResult.Text = txtResult.Text & ",角度=" & Format(dblTmp, "#0.0####")
     
     
    oldval = (2600 / (2 * Screen.TwipsPerPixelY)) - (outputarray(1) * ((950 / Screen.TwipsPerPixelY) / maxvalue))

    Dim sngZ As Single

    sngZ = (2600 / (2 * Screen.TwipsPerPixelY)) ' - (maxvalue / 2) * ((950 / Screen.TwipsPerPixelY) / maxvalue)
    output.Line (0, 2600)-(0, 0)

    For cnt = 1 To N_POINT
    
        ' output.Line (cnt - 1, oldval)-(cnt, (2600 / (2 * Screen.TwipsPerPixelY)) - (outputarray(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue)))
        ' oldval = (2600 / (2 * Screen.TwipsPerPixelY)) - (outputarray(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue))
        oldval = (2600 / (2 * Screen.TwipsPerPixelY)) - (outputarray(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue))
        output.Line (cnt - 1, oldval)-(cnt - 1, sngZ)
        
    Next cnt

End Sub

Private Sub Command1_Click()
  
   Dim nFreq As Single
   nFreq = Val(txtfreq.Text)
   Dim cnt As Integer
    cnt = nFreq * N / Val(txtsamp.Text)
    Dim i As Integer
    Const DIS = 6
    For i = cnt - DIS To cnt + DIS
     REX(i) = 0
     IMX(i) = 0
    Next
    
     For i = (N - (cnt - DIS)) To (N - (cnt + DIS)) Step -1
     REX(i) = 0
     IMX(i) = 0
    Next
    
    FFT0 REX, IMX, NX, -1
   '去除指定频率
   For i = 1 To NX
      List2.AddItem REX(i)
   Next
   redraw
End Sub

Private Sub Command3_Click()
    MsgBox "Developed and published by : " + vbNewLine + "LOGIX4U" + vbNewLine + "www.logix4u.net" + vbNewLine + vbNewLine + "Algorithms by : " + vbNewLine + "Steven W Smith " + vbNewLine + "www.dspguide.com", vbInformation, "About"

End Sub

Private Sub cmdTransform_Click()

    On Error Resume Next

    Label5.Visible = True
    DoEvents

    If Opthamm.Value = True Then
        FuncGen.ApplyHamming
        ipplot.Cls
        plotip

    End If

    If Opthann.Value = True Then
        FuncGen.ApplyHanning
        ipplot.Cls
        plotip

    End If

    If Optblack.Value = True Then
        FuncGen.ApplyBlackman
        ipplot.Cls
        plotip

    End If

    Dim sngMax As Single

    sngMax = -99999
    List1.Clear
    For cnt = 1 To NX

        If cnt <= N Then
            REX(cnt) = FuncGen.Samples(cnt)
        Else
            REX(cnt) = 0

        End If
        List1.AddItem Trim$(REX(cnt))
        
        IMX(cnt) = 0

        If sngMax < REX(cnt) Then
            sngMax = REX(cnt)

        End If

    Next cnt
  
    txtMax = sngMax
    'fft
    FFT0 REX, IMX, NX, 1

    If Optreal.Value = True Then '实部
        output.ForeColor = vbRed

        For cnt = 0 To N / 2
            outputarray(cnt) = REX(cnt)
        Next cnt

    End If

    If Optimg.Value = True Then '虚部
        output.ForeColor = vbWhite

        For cnt = 0 To N / 2
            outputarray(cnt) = IMX(cnt)
        Next cnt

    End If

    If Optmag.Value = True Then '震幅
        output.ForeColor = vbYellow

        For cnt = 0 To N / 2
            outputarray(cnt) = Sqr((IMX(cnt) * IMX(cnt)) + (REX(cnt) * REX(cnt)))
        Next cnt

    End If

    If Optphase.Value = True Then '相位
        output.ForeColor = vbCyan

        For cnt = 0 To N / 2
            outputarray(cnt) = Atn(IMX(cnt) / REX(cnt))
        Next cnt

    End If

    If Optpower.Value = True Then '功率谱
        output.ForeColor = vbGreen

        For cnt = 0 To N / 2
            outputarray(cnt) = Sqr((IMX(cnt) * IMX(cnt)) + (REX(cnt) * REX(cnt)))
        Next cnt

    End If

    For cnt = 0 To N / 2

        If (maxvalue < outputarray(cnt)) Then
            maxvalue = outputarray(cnt)

        End If

    Next cnt
    
    oldval = (2600 / (2 * Screen.TwipsPerPixelY)) - (outputarray(1) * ((950 / Screen.TwipsPerPixelY) / maxvalue))

    For cnt = 1 To N / 2
        output.Line (cnt - 1, oldval)-(cnt, (2600 / (2 * Screen.TwipsPerPixelY)) - (outputarray(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue)))
        oldval = (2600 / (2 * Screen.TwipsPerPixelY)) - (outputarray(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue))
    Next cnt

    Label5.Visible = False

End Sub

Private Sub Form_Activate()

    ipplot.Line (0, 1450 / Screen.TwipsPerPixelY)-(10000, 1450 / Screen.TwipsPerPixelY), vbRed

End Sub

Private Sub Form_Paint()

    On Error Resume Next

    Dim maxvalue, oldval As Double

    ipplot.Cls
    plotip

End Sub

Private Sub List1_DblClick()
   List1.Clear
End Sub

Private Sub List2_Click()
   List2.Clear
End Sub

Private Sub opt_impulse_Click()

    If opt_impulse.Value = True Then
        txt_impulsepos.Enabled = True

    End If

End Sub
Public Sub redraw()

    On Error Resume Next

    Dim maxvalue, oldval As Double

     ipplot.Line (0, 1450 / Screen.TwipsPerPixelY)-(10000, 1450 / Screen.TwipsPerPixelY), vbRed

'    For cnt = 0 To 1024
'
'        If (maxvalue < REX(cnt)) Then
'            maxvalue = REX(cnt)
'
'        End If
'
'    Next cnt
    maxvalue = Val(txtMax.Text)
    oldval = 2900 / (2 * Screen.TwipsPerPixelY)

    For cnt = 1 To 1024
        'mainfrm.ipplot.PSet (cnt, (2900 / (2 * Screen.TwipsPerPixelY)) - (FuncGen.Samples(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue)))
         ipplot.Line (cnt - 1, oldval)-(cnt, (2900 / (2 * Screen.TwipsPerPixelY)) - (REX(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue))), vbRed
        oldval = (2900 / (2 * Screen.TwipsPerPixelY)) - (REX(cnt) * ((950 / Screen.TwipsPerPixelY) / maxvalue))
    Next cnt

End Sub
