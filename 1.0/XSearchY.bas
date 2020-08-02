Attribute VB_Name = "Module2"
Option Explicit
Public num As Integer
Public Control_X() As Single
Public LeftHand_X() As Single
Public RightHand_X() As Single
Public Control_Y() As Single
Public LeftHand_Y() As Single
Public RightHand_Y() As Single
Public Numin As Boolean
Public isaddpoint As Boolean

Public Function XSearchY(Calculate_X As Single) As Single
Dim i As Integer
Dim i_index
Dim ax As Double
Dim bx As Double
Dim cx As Double
Dim dx As Double
Dim ay As Single
Dim by As Single
Dim cy As Single
Dim dy As Single
Dim x1r As Double
Dim x1i As Double
Dim x2r As Double
Dim x2i As Double
Dim x3r As Double
Dim x3i As Double
Dim root As String
Dim t As Single     '0-1的解'
For i = 0 To num - 2
If (Calculate_X >= Control_X(i) And Calculate_X <= Control_X(i + 1)) Or (Calculate_X <= Control_X(i) And Calculate_X >= Control_X(i + 1)) Then
i_index = i
Exit For
End If
Next i
If i <> num - 1 Then    '在某一个区间范围内，区间号为i_index+1
    cx = 3 * (RightHand_X(i_index) - Control_X(i_index))
    bx = 3 * (LeftHand_X(i_index) - RightHand_X(i_index)) - cx
    ax = Control_X(i_index + 1) - Control_X(i_index) - cx - bx
    dx = Control_X(i_index) - Calculate_X
    'x(t) = ax * t ^ 3 + bx * t ^ 2 + cx * t + dx=calculate_x'
    If ax <> 0 Then '三次方程’
    root = Module1.CubicEquation(ax, bx, cx, dx, x1r, x1i, x2r, x2i, x3r, x3i)
    If x1i = 0 And (x1r >= 0 And x1r <= 1) Then t = x1r
    If x2i = 0 And (x2r >= 0 And x2r <= 1) Then t = x2r
    If x3i = 0 And (x3r >= 0 And x3r <= 1) Then t = x3r
    Else
    If bx <> 0 Then '二次方程'
    x1r = (-cx + (cx * cx - 4 * bx * dx) ^ (1 / 2)) / (2 * bx)
    x2r = (-cx - (cx * cx - 4 * bx * dx) ^ (1 / 2)) / (2 * bx)
    If (x1r >= 0 And x1r <= 1) Then t = x1r
    If (x2r >= 0 And x2r <= 1) Then t = x2r
    Else            '一次方程'
    t = -dx / cx
    End If
    End If
    cy = 3 * (RightHand_Y(i_index) - Control_Y(i_index))
    by = 3 * (LeftHand_Y(i_index) - RightHand_Y(i_index)) - cy
    ay = Control_Y(i_index + 1) - Control_Y(i_index) - cy - by
    dy = Control_Y(i_index)
    XSearchY = (ay * t ^ 3 + by * t ^ 2 + cy * t + dy)
    Numin = True
Else
MsgBox "拟合的线不包括散点的X范围！", vbOKOnly, "出错报告！"

Numin = False
End If
End Function
