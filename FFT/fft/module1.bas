Attribute VB_Name = "module1"

Public FuncGen              As New FunctionGenerator

Public impulse_response(32) As Double

Public Const N = 8192

Public outputarray(N + 20)
Public Function Atan2(numberX As Double, numberY As Double) As Double
 
  If numberY < 0 Then
    Atan2 = Atn(numberX / numberY) + PI
  Else
     Atan2 = Atn(numberX / numberY)
  End If
End Function
