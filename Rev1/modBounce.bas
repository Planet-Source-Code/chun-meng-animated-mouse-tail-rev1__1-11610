Attribute VB_Name = "modBounce"
'
'
'
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Type Vec2D
    x As Long
    y As Long
End Type

Public Type AnimBall
    Vec As Vec2D
    dx As Double
    dy As Double
    Img As Image
End Type
'
' variables used
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
    SegLen = 10         ' it seem like the distance between the
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
        balls(i).Img.Left = balls(i).Vec.x
        balls(i).Img.Top = balls(1).Vec.y
    Next i
End Function

Function BallSet(Img As Image) As AnimBall
    BallSet.Vec.x = Xpos
    BallSet.Vec.y = Ypos
    BallSet.dx = BallSet.dy = 0
    Set BallSet.Img = Img
End Function

Function VecSet(x As Long, y As Long) As Vec2D
    VecSet.x = x
    VecSet.y = y
End Function

Function MoveHandler(x As Long, y As Long)
    Xpos = x
    Ypos = y
End Function

Function SpringForce(i As Integer, j As Integer, ByRef spring As Vec2D)
    Dim tempdx, tempdy, tempLen, springF
    tempdx = balls(i).Vec.x - balls(j).Vec.x
    tempdy = balls(i).Vec.y - balls(j).Vec.y
    tempLen = Sqr(tempdx * tempdx + tempdy * tempdy)

    If (tempLen > SegLen) Then
        springF = SpringK * (tempLen - SegLen)
        spring.x = spring.x + (tempdx / tempLen) * springF
        spring.y = spring.y + (tempdy / tempLen) * springF
    End If
End Function

Function Animate()
    Dim iH, iW
    Dim start As Integer
    Dim i As Integer
    Dim spring As Vec2D
    Dim resist As Vec2D
    Dim accel As Vec2D
    
    If (bFollowM) Then
        balls(0).Vec.x = Xpos
        balls(0).Vec.y = Ypos
        start = 1
    End If

    For i = start To nBalls
        ' before apply the new x, y
        ' clear the previous image
        BitBlt frmBounce.hDC, balls(i).Vec.x, _
            balls(i).Vec.y, DotSize, DotSize, frmBounce.picBg.hDC, _
            0, 0, vbSrcCopy

        spring = VecSet(0, 0)

        If (i > 0) Then
            Call SpringForce(i - 1, i, spring)
        End If

        If (i < (nBalls - 1)) Then
            Call SpringForce(i + 1, i, spring)
        End If
        resist = VecSet(-balls(i).dx * Resistance, -balls(i).dy * Resistance)
        accel = VecSet((spring.x + resist.x) / Mass, _
                        (spring.y + resist.y) / Mass + Gravity)

        balls(i).dx = balls(i).dx + DeltaT * accel.x
        balls(i).dy = balls(i).dy + DeltaT * accel.y

        If (Abs(balls(i).dx) < StopVel And _
            Abs(balls(i).dy) < StopVel And _
            Abs(accel.x) < StopAcc And _
            Abs(accel.y) < StopAcc) Then
            balls(i).dx = 0
            balls(i).dy = 0
        End If

        balls(i).Vec.x = balls(i).Vec.x + balls(i).dx
        balls(i).Vec.y = balls(i).Vec.y + balls(i).dy

        ' checking for boundary conditions
        iW = frmBounce.ScaleWidth
        iH = frmBounce.ScaleHeight

        ' check bottom
        If (balls(i).Vec.y >= iH - DotSize - 1) Then
            If (balls(i).dy > 0) Then
                balls(i).dy = Bounce * (-balls(i).dy)
            End If
            balls(i).Vec.y = iH - DotSize - 1
        End If
        
        ' check right
        If (balls(i).Vec.x >= iW - DotSize) Then
            If (balls(i).dx > 0) Then
                balls(i).dx = Bounce * (-balls(i).dx)
            End If
            balls(i).Vec.x = iW - DotSize - 1
        End If

        ' check left
        If (balls(i).Vec.x < 0) Then
            If (balls(i).dx < 0) Then
                balls(i).dx = Bounce * (-balls(i).dx)
            End If
            balls(i).Vec.x = 0
        End If
        ' check top
        If (balls(i).Vec.y < 0) Then
            If (balls(i).dy < 0) Then
                balls(i).dy = Bounce * (-balls(i).dy)
            End If
            balls(i).Vec.y = 0
        End If
' draw the ball using bitblt, before this is call
' the picball should be process with some transparent
' bitmap routine. (not implement here)
        'balls(i).Img.Left = balls(i).Vec.x
        'balls(i).Img.Top = balls(i).Vec.y
        'frmBounce.picBall.Picture = balls(i).Img.Picture
        BitBlt frmBounce.hDC, balls(i).Vec.x, _
            balls(i).Vec.y, DotSize, DotSize, frmBounce.picBall.hDC, _
            0, 0, vbSrcCopy
'       frmBounce.PaintPicture frmBounce.picBall.Picture, _
            balls(i).Vec.x, balls(i).Vec.y, _
            DotSize, DotSize, , , , , vbSrcCopy

    Next i
End Function
