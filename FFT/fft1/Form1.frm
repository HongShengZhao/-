VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "FFT1"
      Height          =   360
      Left            =   5520
      TabIndex        =   2
      Top             =   7590
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FFT0"
      Height          =   360
      Left            =   3540
      TabIndex        =   1
      Top             =   7590
      Width           =   990
   End
   Begin VB.PictureBox picI_FFT 
      AutoRedraw      =   -1  'True
      Height          =   7245
      Left            =   120
      ScaleHeight     =   7185
      ScaleWidth      =   11175
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   11235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'*使用**********

    Const fftIn = 128
    Dim i As Integer
    Dim xr(fftIn) As Double
    Dim xi(fftIn) As Double
    Dim sngMax As Single
    Dim nMax As Integer
    
    Dim sngMin As Single
    Dim nMin As Integer
    
'赋值，IaIn(i)是采得的数据。
    For i = 0 To fftIn
    xr(i) = 100 * Sin(i)
    xi(i) = 0
    Next
 picI_FFT.Scale (0, 100)-(fftIn - 1, -100)
    picI_FFT.DrawWidth = 1
     
'FFT变换
    Call FFT0(xr(), xi(), fftIn, 1)

'绘图
   sngMax = 0
   sngMin = 99999
   
    For i = 0 To fftIn / 2 '- 1
      'picI_FFT.Line (i, Abs(xr(i)))-(i + 1, Abs(xr(i + 1))), vbBlue
       picI_FFT.Line (i, (xr(i)))-(i + 1, (xr(i + 1))), vbBlue
       If Abs(xr(i)) > sngMax Then
          sngMax = xr(i)
          nMax = i
       End If
    Next i
    picI_FFT.Line (nMax, 100)-(nMax, -100), vbRed
    picI_FFT.ForeColor = vbRed
    picI_FFT.CurrentX = 0
    picI_FFT.CurrentY = 0
    picI_FFT.Print nMax
    Call FFT0(xr(), xi(), fftIn, -1)
     For i = 0 To fftIn - 1

       picI_FFT.Line (i, (xr(i)))-(i + 1, (xr(i + 1))), vbRed
     Next i

    
End Sub

Private Sub Command2_Click()
   '*使用**********

    Const fftIn = 128
    Dim i As Integer
    Dim xr(1 To fftIn) As Double
    Dim xi(1 To fftIn) As Double
    Dim sngMax As Single
    Dim sngMin As Single
    
'赋值，IaIn(i)是采得的数据。
    For i = 1 To fftIn
    xr(i) = 100 * Sin(i)
    xi(i) = 0
    Next
 picI_FFT.Scale (0, 100)-(fftIn, -100)
    picI_FFT.DrawWidth = 1
   
'FFT变换
    Call FFT1(xr(), xi(), fftIn, 1)

'绘图
   
    For i = 1 To fftIn - 1
      'picI_FFT.Line (i, Abs(xr(i)))-(i + 1, Abs(xr(i + 1))), vbBlue
       picI_FFT.Line (i, (xr(i)))-(i + 1, (xr(i + 1))), vbBlue
    Next i
    
    Call FFT1(xr(), xi(), fftIn, -1)
     For i = 1 To fftIn - 1
      'picI_FFT.Line (i, Abs(xr(i)))-(i + 1, Abs(xr(i + 1))), vbBlue
       picI_FFT.Line (i, (xr(i)))-(i + 1, (xr(i + 1))), vbRed
     Next i
End Sub
