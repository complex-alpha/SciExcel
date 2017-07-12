Attribute VB_Name = "CalcOfRotation"
Function Rotation(x As Double, y As Double, z As Double, roll As Double, pitch As Double, yaw As Double, sign As String) As Double

    'Temporary Variable
    Dim cx As Double
    Dim cy As Double
    Dim cz As Double

    'Roll Convert
    cx = x
    cy = y
    x = RotTransU(cx, cy, roll)
    y = RotTransL(cx, cy, roll)

    'Pitch Convert
    cy = y
    cz = z
    y = RotTransU(cy, cz, pitch)
    z = RotTransL(cy, cz, pitch)
    
    'Yaw Convert
    cz = z
    cx = x
    z = RotTransU(cz, cx, yaw)
    x = RotTransL(cz, cx, yaw)
    
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
