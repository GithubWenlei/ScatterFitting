VERSION 5.00
Object = "{EB7A6012-79A9-4A1A-91AF-F2A92FCA3406}#1.0#0"; "TeeChart8.ocx"
Begin VB.Form Form1 
   Caption         =   "散点拟合查值软件"
   ClientHeight    =   8925
   ClientLeft      =   1875
   ClientTop       =   2175
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   13230
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   11400
      TabIndex        =   3
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Command9 
         Caption         =   "批量查询Y值"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Text            =   "100"
         Top             =   5281
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   6777
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "计算Y"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   6029
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "导入数据"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   7525
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "退出"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   8280
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "导出离散数据"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "导入曲线"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   4533
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Text            =   "1000"
         Top             =   3037
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "单值检验"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2289
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "已定好"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1541
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "显示"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   793
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Text            =   "定线控制点数"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X值："
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   5400
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "离散点"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   540
      End
   End
   Begin TeeChart.TeeCommander TeeCommander1 
      Height          =   735
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   1
      Top             =   8160
      Width           =   11175
   End
   Begin TeeChart.TChart TChart1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Base64          =   $"Form1.frx":004D
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   7815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
Dim on_off As Boolean
Dim number As Integer
'Dim num As Single
'Dim Control_X() As Single
'Dim LeftHand_X() As Single
'Dim RightHand_X() As Single
'Dim Control_Y() As Single
'Dim LeftHand_Y() As Single
'Dim RightHand_Y() As Single
Dim r() As Double
Dim p() As Double
Dim tp As String * 16
Dim tr As String * 16
Dim x() As Single
Dim y() As Single
Dim curveNum As Integer
Private Sub Command1_Click()
Dim i As Integer

Dim ss As Single
num = Val(Combo1.Text)
number = num + 2 * num - 2
If p(2) = 0 Then
MsgBox ("请先导入数据")
Else
ss = UBound(p)
ReDim x(number - 1): ReDim y(number - 1)
x(0) = minnum(r) 'r(2)
x(number - 1) = maxnum(r)
For i = 1 To UBound(p)
If x(0) = r(i) Then
y(0) = p(i)
ElseIf r(i) = x(number - 1) Then
y(number - 1) = p(i)
End If
Next i

For i = 1 To number - 2
x(i) = x(0) + (maxnum(r) - x(0)) * i / number 'r(i * ss / number + 1)
y(i) = p(i * ss / number + 1)
Next i
TChart1.Series(0).Clear
TChart1.Series(0).AddArray number, y, x
TChart1.Legend.Hide
'TChart1.Legend.Show
TChart1.Series(0).asBezier.BezierStyle = bsBezier4
For i = 0 To number - 1
If i Mod 3 = 0 Then
TChart1.Series(0).PointColor(i) = vbRed
Else
If (i + 1) Mod 3 = 0 Then
TChart1.Series(0).PointColor(i) = vbBlue
Else
TChart1.Series(0).PointColor(i) = vbMagenta
End If
End If
Next i
TChart1.Tools.Active = True
End If

isaddpoint = True
End Sub

Private Sub Command2_Click()
If isaddpoint = True Then
If on_off = True Then   '确定
    TChart1.Tools.Active = False
    ReDim Control_X(num - 1) As Single
    ReDim LeftHand_X(num - 2) As Single
    ReDim RightHand_X(num - 2) As Single
    ReDim Control_Y(num - 1) As Single
    ReDim LeftHand_Y(num - 2) As Single
    ReDim RightHand_Y(num - 2) As Single
    Dim i As Integer
    Dim i_Co As Integer: i_Co = 0
    Dim i_Le As Integer: i_Le = 0
    Dim i_Ri As Integer: i_Ri = 0
    For i = 0 To number - 1
    If i Mod 3 = 0 Then
    Control_X(i_Co) = TChart1.Series(0).XValues.Value(i)
    Control_Y(i_Co) = TChart1.Series(0).YValues.Value(i)
    i_Co = i_Co + 1
    Else
    If (i + 1) Mod 3 = 0 Then
    LeftHand_X(i_Le) = TChart1.Series(0).XValues.Value(i)
    LeftHand_Y(i_Le) = TChart1.Series(0).YValues.Value(i)
    i_Le = i_Le + 1
    Else
    RightHand_X(i_Ri) = TChart1.Series(0).XValues.Value(i)
    RightHand_Y(i_Ri) = TChart1.Series(0).YValues.Value(i)
    i_Ri = i_Ri + 1
    End If
    End If
    Next i
    Command3.Enabled = True
    Command2.Caption = "重新定"
    Command7.Enabled = True
    Command4.Enabled = True
    Command1.Enabled = False
    Label4.Caption = "NSE: " & Format(NSE(r, p), "0.000")
    curveNum = curveNum + 1
Dim outi As Integer
Dim outax() As Double
Dim outbx() As Double
Dim outcx() As Double
Dim outdx() As Double
Dim outay() As Single
Dim outby() As Single
Dim outcy() As Single
Dim outdy() As Single
ReDim outax(num - 1), outbx(num - 1), outcx(num - 1), outdx(num - 1), outay(num - 1), outby(num - 1), outcy(num - 1), outdy(num - 1)
Open App.Path & "\out_curve" & curveNum & ".csv" For Output As #3
Open App.Path & "\out_curve" & curveNum & ".txt" For Output As #4

For outi = 0 To num - 2
  '在某一个区间范围内，区间号为i_index+1
    outcx(outi) = 3 * (RightHand_X(outi) - Control_X(outi))
    outbx(outi) = 3 * (LeftHand_X(outi) - RightHand_X(outi)) - outcx(outi)
    outax(outi) = Control_X(outi + 1) - Control_X(outi) - outcx(outi) - outbx(outi)
    outdx(outi) = Control_X(outi)
    
    outcy(outi) = 3 * (RightHand_Y(outi) - Control_Y(outi))
    outby(outi) = 3 * (LeftHand_Y(outi) - RightHand_Y(outi)) - outcy(outi)
    outay(outi) = Control_Y(outi + 1) - Control_Y(outi) - outcy(outi) - outby(outi)
    outdy(outi) = Control_Y(outi)
    Write #3, "第" & outi + 1 & "段曲线参数", "x:", Control_X(outi), "to", Control_X(outi + 1), "y:", Control_Y(outi), "to", Control_Y(outi + 1)
    Write #3, "x=a*t^3+b*t^2+c*t+d   t:0--1"
    Write #3, "y=e*t^3+f*t^2+g*t+h   t:0--1"
    Write #3, "a=", outax(outi), "b=", outbx(outi), "c=", outcx(outi), "d=", outdx(outi)
    Write #3, "e=", outay(outi), "f=", outby(outi), "g=", outcy(outi), "h=", outdy(outi)
    Write #3,
    Write #4, Control_X(outi), Control_Y(outi)
    Write #4, LeftHand_X(outi); LeftHand_Y(outi)
    Write #4, RightHand_X(outi), RightHand_Y(outi)
Next outi
Write #4, Control_X(outi), Control_Y(outi)
Close #3
Close #4
Else
    TChart1.Tools.Active = True
    Command2.Caption = "已定好"
    Command3.Enabled = False
    Command7.Enabled = False
    Command1.Enabled = True
    
End If
Command5.Enabled = False
'Command1.Enabled = False

on_off = Not on_off
Else
MsgBox ("请先添加模拟曲线")
End If

End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim ax As Single
Dim bx As Single
Dim cx As Single
Dim t As Single
For i = 0 To num - 2
    cx = 3 * (RightHand_X(i) - Control_X(i))
    bx = 3 * (LeftHand_X(i) - RightHand_X(i)) - cx
    ax = Control_X(i + 1) - Control_X(i) - cx - bx
    If ax <> 0 Then
    If ax > 0 Then
        If (-bx / (3 * ax)) < 0 Then
        If cx < 0 Then MsgBox "第" & Str(i + 1) & "段不是单值关系！重新定线！", vbOKOnly, "提示信息"
        Else
            If (-bx / (3 * ax)) <= 1 Then
            t = -bx / (3 * ax)
            If 3 * ax * t * t + 2 * bx * t + cx < 0 Then MsgBox "第" & Str(i + 1) & "段不是单值关系！重新定线！", vbOKOnly, "提示信息"
            Else
            If 3 * ax + 2 * bx + cx < 0 Then MsgBox "第" & Str(i + 1) & "段不是单值关系！重新定线！", vbOKOnly, "提示信息"
            End If
        End If
    Else
        If (-bx / (3 * ax)) < 0.5 Then
        If 3 * ax + 2 * bx + cx < 0 Then MsgBox "第" & Str(i + 1) & "段不是单值关系！重新定线！", vbOKOnly, "提示信息"
        Else
        If cx < 0 Then MsgBox "第" & Str(i + 1) & "段不是单值关系！重新定线！", vbOKOnly, "提示信息"
        End If
    End If
    Else
        If bx <> 0 Then
        If bx > 0 Then
        If cx < 0 Then MsgBox "第" & Str(i + 1) & "段不是单值关系！重新定线！", vbOKOnly, "提示信息"
        Else
        If 2 * bx + cx < 0 Then MsgBox "第" & Str(i + 1) & "段不是单值关系！重新定线！", vbOKOnly, "提示信息"
        End If
        Else
        If cx < 0 Then MsgBox "第" & Str(i + 1) & "段不是单值关系！重新定线！", vbOKOnly, "提示信息"
        End If
    End If
Next i
End Sub

Private Sub Command4_Click()
Dim Calculate_X As Single
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
Calculate_X = Val(Text1)
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
    Text2 = Str(ay * t ^ 3 + by * t ^ 2 + cy * t + dy)
    Numin = True
Else
MsgBox "要查找的点在定线的范围之外！", vbOKOnly, "出错报告！"
Numin = False
End If
End Sub

Private Sub Command5_Click()
Dim i As Integer
i = 1
Dim dataFile As String
dataFile = InputBox("请输入数据文件名：", "导入数据：", "data.csv")
If dataFile = "" Then
MsgBox ("请输入数据文件名称！")
'With CommonDialog1
'.DialogTitle = "打开数据文件"
'.InitDir = App.Path
'.Filter = "逗号分割文件(*.csv) |*.csv|文本文件(*.txt) |*.txt|所有文件(*.*) |*.*"
'.FilterIndex = 1
'.ShowOpen '或使用CommonDialog1.Action=1
'Open .FileName For Input As #1
'End With
Else
TChart1.Series(0).Clear
'Open App.Path & "\data.csv" For Input As #1
Open App.Path & "\" & dataFile For Input As #1
Input #1, tp, tr
Do While Not EOF(1)
ReDim Preserve r(i), p(i)
Input #1, r(i), p(i)
i = i + 1
Loop
Close #1

For i = 1 To UBound(r)
'p(i) = DateDiff("h", tp(1), tp(i))
TChart1.Series(1).AddXY r(i), p(i), "", vbGreen
Next i
End If
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
Dim wl As Long
Dim XgetY As Single
Open App.Path & "\out.csv" For Output As #2
Write #2, tp, tr
Dim outn As Long
outn = Val(Text3.Text)
For wl = 0 To outn
'w = wl * (x(number - 1) - x(0)) / 1000 + x(0)
XgetY = XSearchY(wl * (x(number - 1) - x(0)) / outn + x(0))
'XgetY = XSearchY(1# * wl)
If Numin = False Then
Exit For
Else
Write #2, wl * (x(number - 1) - x(0)) / outn + x(0), XgetY
'Write #2, wl, XgetY
End If
Next wl
Close #2
End Sub

Private Sub Command8_Click()
Dim inx() As Double
Dim iny() As Double
Dim ini As Integer
Dim curveFile As String
'Open App.Path & "\out_curve" & curveNum & ".txt" For Input As #5
curveFile = InputBox("请输入导入曲线文件名：", "导入曲线参数：", "out_curve1.txt")
If curveFile = "" Then
MsgBox ("请输入数据文件名称！")
Else
Open App.Path & "\" & curveFile For Input As #5

Do While Not EOF(5)
ReDim Preserve inx(ini), iny(ini)
Input #5, inx(ini), iny(ini)
ini = ini + 1
Loop
Close #5
TChart1.AddSeries (scBezier)
Dim nCv As Integer: nCv = TChart1.SeriesCount - 1
Dim nNumberC As Integer: nNumberC = UBound(inx) + 1

TChart1.Series(nCv).AddArray nNumberC, iny, inx
TChart1.Series(nCv).asBezier.BezierStyle = bsBezier4
For ini = 0 To nNumberC - 1
If ini Mod 3 = 0 Then
TChart1.Series(nCv).PointColor(ini) = vbRed

Else
If (ini + 1) Mod 3 = 0 Then
TChart1.Series(nCv).PointColor(ini) = vbBlue
Else
TChart1.Series(nCv).PointColor(ini) = vbMagenta
End If
End If
Next ini

TChart1.Tools.Active = True
End If
End Sub

Private Sub Command9_Click()
Dim Calculate_X As Single
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
Dim wl As Integer
Dim root As String
Dim t As Single     '0-1的解'
i = 1
Dim dataFile As String
dataFile = InputBox("请输入数据文件名：", "导入批量查询的X数据：", "Xdata.csv")
If dataFile = "" Then
MsgBox ("请输入数据文件名称！")
'With CommonDialog1
'.DialogTitle = "打开数据文件"
'.InitDir = App.Path
'.Filter = "逗号分割文件(*.csv) |*.csv|文本文件(*.txt) |*.txt|所有文件(*.*) |*.*"
'.FilterIndex = 1
'.ShowOpen '或使用CommonDialog1.Action=1
'Open .FileName For Input As #1
'End With
Else

'Open App.Path & "\data.csv" For Input As #1
Open App.Path & "\" & dataFile For Input As #1
Input #1, tp, tr
Do While Not EOF(1)
ReDim Preserve r(i)
Input #1, r(i)
i = i + 1
Loop
Close #1
End If
Open App.Path & "\Yout.csv" For Output As #2

For wl = 1 To UBound(r)
Calculate_X = Val(r(wl))
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
    Text2 = Str(ay * t ^ 3 + by * t ^ 2 + cy * t + dy)
    Write #2, r(wl), ay * t ^ 3 + by * t ^ 2 + cy * t + dy
    Numin = True
Else
Write #2, r(wl), 1
MsgBox "要查找的点在定线的范围之外！", vbOKOnly, "出错报告！"
Numin = False
End If
Next wl
Close #2

End Sub

Private Sub Form_Load()
curveNum = 0
Numin = True
isaddpoint = False
ReDim p(2)
TeeCommander1.Chart = TChart1
Dim i As Integer
For i = 2 To 10
Combo1.AddItem Str(i)
Next i
Combo1.ListIndex = 2
TChart1.Tools.Active = False
on_off = True
Command3.Enabled = False
End Sub


