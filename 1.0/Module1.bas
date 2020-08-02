Attribute VB_Name = "Module1"
Option Explicit
'a = InputBox("请输入a", eq)
'b = InputBox("请输入b", eq)
'c = InputBox("请输入c", eq)
'd = InputBox("请输入d", eq)
'ret = CubicEquation(a, b, c, d, x1r, x1i, x2r, x2i, x3r, x3i)    '5x^3+4x^2+3x-12=0
'Debug.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & ret
'Debug.Print x1r; " + "; x1i; " i"
'Debug.Print x2r; " + "; x2i; " i"
'Debug.Print x3r; " + "; x3i; " i"
Public Function CubicEquation _
(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, _
 x1r As Double, x1i As Double, x2r As Double, x2i As Double, x3r As Double, x3i As Double) As String
'Cubic equation(v2.2), coded by www.dayi.net btef (please let this line remain)
Dim e As Double, f As Double, g As Double, h As Double, delta As Double
Dim r As Double, sita As Double, pi As Double, rr As Double, ri As Double

If a = 0 Then
    CubicEquation = "Not a cubic equation: a = 0"
    Exit Function
End If

'pi = 3.14159265358979
pi = 4 * Atn(1)
b = b / a                           'simplify to a=1: x^3+bx^2+cx+d=0
c = c / a
d = d / a
e = -b ^ 2 / 3 + c              'substitute x=y-b/3: y^3+ey+f=0
f = (2 * b ^ 2 - 9 * c) * b / 27 + d

If e = 0 And f = 0 Then
    x1r = -b / 3
    x2r = x1r
    x3r = x1r
    CubicEquation = "3 same real roots:"
ElseIf e = 0 Then              'need to deal with e = 0, or it will cause z = 0 later.
    r = -f                           'y^3+f=0, y^3=-f
    r = Cur(r)
    x1r = r - b / 3               'a real root
    If r > 0 Then                 'r never = 0 since g=f/2, f never = 0 there
    sita = 2 * pi / 3
    x2r = r * Cos(sita) - b / 3
    x2i = r * Sin(sita)
    Else
    sita = pi / 3
    x2r = -r * Cos(sita) - b / 3
    x2i = -r * Sin(sita)
    End If
    x3r = x2r
    x3i = -x2i
    CubicEquation = "1 real root and 2 image roots:"
Else                                 'substitute y=z-e/3/z: (z^3)^2+fz^3-(e/3)^3=0, z^3=-g+sqr(delta)
    g = f / 2                       '-q-sqr(delta) is ignored
    h = e / 3
    delta = g ^ 2 + h ^ 3
    If delta < 0 Then
        r = Sqr(g ^ 2 - delta)
        sita = Argument(-g, Sqr(-delta))           'z^3=r(con(sita)+isin(sita))
        r = Cur(r)
        rr = r - h / r
        sita = sita / 3                                     'z1=r(cos(sita)+isin(sita))
        x1r = rr * Cos(sita) - b / 3                  'y1=(r-h/r)cos(sita)+i(r+h/r)sin(sita), x1=y1-b/3
        sita = sita + 2 * pi / 3                        'no image part since r+h/r = 0
        x2r = rr * Cos(sita) - b / 3
        sita = sita + 2 * pi / 3
        x3r = rr * Cos(sita) - b / 3
        CubicEquation = "3 real roots:"
      Else                                                   'delta >= 0
        r = -g + Sqr(delta)
        r = Cur(r)
        rr = r - h / r
        ri = r + h / r
        If ri = 0 Then
        CubicEquation = "3 real roots:"
        Else
        CubicEquation = "1 real root and 2 image roots:"
        End If
        x1r = rr - b / 3                             'a real root
        If r > 0 Then                                'r never = 0 since g=f/2, f never = 0 there
        sita = 2 * pi / 3
        x2r = rr * Cos(sita) - b / 3
        x2i = ri * Sin(sita)
        Else                                            'r < 0
        sita = pi / 3
        x2r = -rr * Cos(sita) - b / 3
        x2i = -ri * Sin(sita)
        End If
        x3r = x2r
        x3i = -x2i
      End If
End If

End Function

Private Function Cur(v As Double) As Double

If v < 0 Then
    Cur = -(-v) ^ (1 / 3)
Else
    Cur = v ^ (1 / 3)
End If

End Function

Private Function Argument(a As Double, b As Double) As Double
Dim sita As Double, pi As Double

'pi = 3.14159265358979
pi = 4 * Atn(1)
If a = 0 Then
    If b >= 0 Then
    Argument = pi / 2
    Else
    Argument = -pi / 2
    End If
Else

    sita = Atn(Abs(b / a))
   
    If a > 0 Then
        If b >= 0 Then
        Argument = sita
        Else
        Argument = -sita
        End If
    ElseIf a < 0 Then
        If b >= 0 Then
        Argument = pi - sita
        Else
        Argument = pi + sita
        End If
    End If

End If

End Function
Public Function maxnum(a() As Double) As Double
Dim ww As Long
Dim temp As Double
Dim iii As Long
temp = a(1)
For iii = 2 To UBound(a)
If a(iii) > temp Then
temp = a(iii)
End If
Next iii
maxnum = temp
End Function
Public Function minnum(a() As Double) As Double
Dim ww As Long
Dim temp As Double
Dim iii As Long
temp = a(1)
For iii = 2 To UBound(a)
If a(iii) < temp Then
temp = a(iii)
End If
Next iii
minnum = temp
End Function
Public Function NSE(obsX() As Double, obsY() As Double) As Double
Dim wll As Long
Dim Ne1 As Double
Dim NE2 As Double
Dim averageObs As Double
Ne1 = 0
NE2 = 0
For wll = 1 To UBound(obsY)
averageObs = averageObs + obsY(wll)
Next wll
averageObs = averageObs / (UBound(obsY) - 1)
For wll = 1 To UBound(obsY)
If Numin = True Then
Ne1 = Ne1 + (obsY(wll) - XSearchY(Val(obsX(wll)))) ^ 2
NE2 = NE2 + (obsY(wll) - averageObs) ^ 2
Else
Exit For
End If
Next wll
If NE2 > 0 Then
NSE = 1 - Ne1 / NE2
Else
MsgBox ("模拟的线X范围未包含数据点的范围！请重新定线！")
End If

End Function

