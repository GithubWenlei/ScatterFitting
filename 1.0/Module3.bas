Attribute VB_Name = "Module3"
Option Explicit
Public num As Integer
Public Control_X() As Single
Public LeftHand_X() As Single
Public RightHand_X() As Single
Public Control_Y() As Single
Public LeftHand_Y() As Single
Public RightHand_Y() As Single
Public Function YSearchX(Calculate_Y As Single) As Single
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
Dim t As Single     '0-1�Ľ�'
For i = 0 To num - 2
If (Calculate_Y >= Control_Y(i) And Calculate_Y <= Control_Y(i + 1)) Or (Calculate_Y <= Control_Y(i) And Calculate_Y >= Control_Y(i + 1)) Then
i_index = i
Exit For
End If
Next i
If i <> num - 1 Then    '��ĳһ�����䷶Χ�ڣ������Ϊi_index+1
    cx = 3 * (RightHand_Y(i_index) - Control_Y(i_index))
    bx = 3 * (LeftHand_Y(i_index) - RightHand_Y(i_index)) - cx
    ax = Control_Y(i_index + 1) - Control_Y(i_index) - cx - bx
    dx = Control_Y(i_index) - Calculate_Y
    'x(t) = ax * t ^ 3 + bx * t ^ 2 + cx * t + dx=calculate_x'
    If ax <> 0 Then '���η��̡�
    root = Module1.CubicEquation(ax, bx, cx, dx, x1r, x1i, x2r, x2i, x3r, x3i)
    If x1i = 0 And (x1r >= 0 And x1r <= 1) Then t = x1r
    If x2i = 0 And (x2r >= 0 And x2r <= 1) Then t = x2r
    If x3i = 0 And (x3r >= 0 And x3r <= 1) Then t = x3r
    Else
    If bx <> 0 Then '���η���'
    x1r = (-cx + (cx * cx - 4 * bx * dx) ^ (1 / 2)) / (2 * bx)
    x2r = (-cx - (cx * cx - 4 * bx * dx) ^ (1 / 2)) / (2 * bx)
    If (x1r >= 0 And x1r <= 1) Then t = x1r
    If (x2r >= 0 And x2r <= 1) Then t = x2r
    Else            'һ�η���'
    t = -dx / cx
    End If
    End If
    cy = 3 * (RightHand_X(i_index) - Control_X(i_index))
    by = 3 * (LeftHand_X(i_index) - RightHand_X(i_index)) - cy
    ay = Control_X(i_index + 1) - Control_X(i_index) - cy - by
    dy = Control_X(i_index)
    YSearchX = (ay * t ^ 3 + by * t ^ 2 + cy * t + dy)
Else
MsgBox "Ҫ��ϵĵ��ڶ��ߵķ�Χ֮�⣡", vbOKOnly, "�����棡"
End If
End Function
