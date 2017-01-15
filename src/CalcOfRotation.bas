Attribute VB_Name = "CalcOfRotation"
Function Rotation(x As Double, y As Double, z As Double, roll As Double, pitch As Double, yaw As Double, sign As String) As Double

    Dim cX As Double
    Dim cY As Double
    Dim cZ As Double

    'Roll Convert
    cX = Math.Cos(roll) * x - Math.Sin(roll) * y
    cY = Math.Sin(roll) * x + Math.Cos(roll) * y
    cZ = z

    x = cX
    y = cY
    z = cZ


    'Pitch Convert
    cY = Math.Cos(pitch) * y - Math.Sin(pitch) * z
    cZ = Math.Sin(pitch) * y + Math.Cos(pitch) * z
    cX = x

    x = cX
    y = cY
    z = cZ


    'Yaw Convert
    cZ = Math.Cos(yaw) * z - Math.Sin(yaw) * x
    cX = Math.Sin(yaw) * z + Math.Cos(yaw) * x
    cY = y

    x = cX
    y = cY
    z = cZ



    If sign = "x" Then

        Rotation = x

    ElseIf sign = "y" Then

        Rotation = y

    ElseIf sign = "z" Then

        Rotation = z

    End If

End Function
