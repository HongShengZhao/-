VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim hpict As PictureBox
Public Property Get Picturehandle() As PictureBox
SetPicturehandle = hpict
End Property

Public Property Let Picturehandle(ByVal vNewValue As PictureBox)
    hpict = vNewValue
End Property

Public Function SetPixel(x As Long, y As Long)
    hpict.PSet (x, y)
End Function
