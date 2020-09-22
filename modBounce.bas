Attribute VB_Name = "modBounce"
'
'
'
Option Explicit

Public Type Vec2D
    X As Long
    Y As Long
End Type

Public Type AnimBall
    Vec As Vec2D
    dx As Double
    dy As Double
    Img As Image
End Type

Dim nBalls As Integer
Dim Xpos, Ypos
Dim DeltaT As Double
Dim SegLen
Dim SpringK
Dim Mass
Dim Gravity
Dim Resistance
Dim StopVel As Double
Dim StopAcc As Double
Dim DotSize As Long
Dim Bounce As Double
Dim bFollowM As Boolean
Dim balls() As AnimBall

Function InitVal()
' Some of the variables are still unknown to me
    nBalls = 7          ' numbers of ball
    Xpos = Ypos = 0     ' evaluate position
    DeltaT = 0.01       '
    SegLen = 10#        ' it seem like the distance between the
                        ' mouse pointer and the ball
                        ' it's quite intersting to change the value
                        ' and see the effect
    SpringK = 11       ' spring constant,
                       ' if large, the longer and higher the tail
                       ' will swing
    Mass = 1            'mass of the ball
    Gravity = 40        ' gravity coeff,
                        ' if large, the balls are more difficult
                        ' to move upward
    Resistance = 9     ' resistivity of the ball to move itself
                        ' from a location, the larger the more difficult to
                        ' move
    StopVel = 0.1
    StopAcc = 0.1
    DotSize = 11        ' the size of the ball in pixel
    Bounce = 0.95       ' bouncing coeff,
    bFollowM = True     ' animation flag
End Function


' must only be called after load all imgBall
Function InitBall()
    Dim i As Integer
    ReDim balls(nBalls)

    For i = 0 To nBalls
        balls(i) = BallSet(frmBounce.ImgBall(i))
    Next i

    For i = 0 To nBalls
        balls(i).Img.Left = balls(i).Vec.X
        balls(i).Img.Top = balls(1).Vec.Y
    Next i
End Function

' initialize a ball
Function BallSet(Img As Image) As AnimBall
    BallSet.Vec.X = Xpos
    BallSet.Vec.Y = Ypos
    BallSet.dx = BallSet.dy = 0
    Set BallSet.Img = Img
End Function

' initialize a vector variable
Function VecSet(X As Long, Y As Long) As Vec2D
    VecSet.X = X
    VecSet.Y = Y
End Function

' update position when mouse move
Function MoveHandler(X As Long, Y As Long)
    Xpos = X
    Ypos = Y
End Function

' calculate the spring force of the balls chain
Function SpringForce(i As Integer, j As Integer, ByRef spring As Vec2D)
    Dim tempdx, tempdy, tempLen, springF
    tempdx = balls(i).Vec.X - balls(j).Vec.X
    tempdy = balls(i).Vec.Y - balls(j).Vec.Y
    tempLen = Sqr(tempdx * tempdx + tempdy * tempdy)
    If (tempLen > SegLen) Then
        springF = SpringK * (tempLen - SegLen)
        spring.X = spring.X + (tempdx / tempLen) * springF
        spring.Y = spring.Y + (tempdy / tempLen) * springF
    End If
End Function

' main routine of this animated balls
' call on mouse move or every 20ms
Function Animate()
    Dim iH, iW
    Dim start As Integer
    Dim i As Integer
    Dim spring As Vec2D
    Dim resist As Vec2D
    Dim accel As Vec2D
    ' enable the animation
    If (bFollowM) Then
        balls(0).Vec.X = Xpos
        balls(0).Vec.Y = Ypos
        start = 1
    End If
    
    For i = start To nBalls
        spring = VecSet(0, 0)
        
        If (i > 0) Then
            Call SpringForce(i - 1, i, spring)
        End If
        
        If (i < (nBalls - 1)) Then
            Call SpringForce(i + 1, i, spring)
        End If
        resist = VecSet(-balls(i).dx * Resistance, -balls(i).dy * Resistance)
        accel = VecSet((spring.X + resist.X) / Mass, _
                        (spring.Y + resist.Y) / Mass + Gravity)

        balls(i).dx = balls(i).dx + DeltaT * accel.X
        balls(i).dy = balls(i).dy + DeltaT * accel.Y

        If (Abs(balls(i).dx) < StopVel And _
            Abs(balls(i).dy) < StopVel And _
            Abs(accel.X) < StopAcc And _
            Abs(accel.Y) < StopAcc) Then
            balls(i).dx = 0
            balls(i).dy = 0
        End If

        balls(i).Vec.X = balls(i).Vec.X + balls(i).dx
        balls(i).Vec.Y = balls(i).Vec.Y + balls(i).dy

        ' checking for boundary conditions
        iW = frmBounce.ScaleWidth
        iH = frmBounce.ScaleHeight

        ' check bottom
        If (balls(i).Vec.Y >= iH - DotSize - 1) Then
            If (balls(i).dy > 0) Then
                balls(i).dy = Bounce * (-balls(i).dy)
            End If
            balls(i).Vec.Y = iH - DotSize - 1
        End If
        
        ' check right
        If (balls(i).Vec.X >= iW - DotSize) Then
            If (balls(i).dx > 0) Then
                balls(i).dx = Bounce * (-balls(i).dx)
            End If
            balls(i).Vec.X = iW - DotSize - 1
        End If

        ' check left
        If (balls(i).Vec.X < 0) Then
            If (balls(i).dx < 0) Then
                balls(i).dx = Bounce * (-balls(i).dx)
            End If
            balls(i).Vec.X = 0
        End If
        ' check top
        If (balls(i).Vec.Y < 0) Then
            If (balls(i).dy < 0) Then
                balls(i).dy = Bounce * (-balls(i).dy)
            End If
            balls(i).Vec.Y = 0
        End If

        balls(i).Img.Left = balls(i).Vec.X
        balls(i).Img.Top = balls(i).Vec.Y
    Next i
End Function
