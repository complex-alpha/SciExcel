Attribute VB_Name = "CalcOfRotation"
Function Rotation(x As Double, y As Double, z As Double, roll As Double, pitch As Double, yaw As Double, sign As String) As Double

    'Roll Convert
    x = RotTransU(x, y, roll)
    y = RotTransL(x, y, roll)
    
    'Pitch Convert
    y = RotTransU(y, z, pitch)
    z = RotTransL(y, z, pitch)
    
    'Yaw Convert
    z = RotTransU(z, x, yaw)
    x = RotTransL(z, x, yaw)
    
    If sign = "x" Then

        Rotation = x

    ElseIf sign = "y" Then

        Rotation = y

    ElseIf sign = "z" Then

        Rotation = z

    End If

End Function

Private Function RotTransU(x As Double, y As Double, theta As Double) As Double

    RotTransU = x * Math.Cos(theta) - y * Math.Sin(theta)
    
End Function

Private Function RotTransL(x As Double, y As Double, theta As Double) As Double

    RotTransL = x * Math.Sin(theta) + y * Math.Cos(theta)
    
End Function
