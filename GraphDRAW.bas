Attribute VB_Name = "GraphDRAW"
Option Explicit
'微分係数計算幅
Public Const h = 0.00000001
'媒介変数範囲
Public Const t_min = 0
Public Const t_max = 5
' 分割数
Public Const n = 5

Function s(t As Double) As Double
    '速度媒介変数
    s = t
End Function

Function f(t As Double) As Double
    '曲線のx成分 x=f(s)
    f = s(t)
End Function
Function g(t As Double) As Double
    '曲線のy成分 y=g(s)
    g = B(3, 2, s(t))
End Function

Function Df(t As Double) As Double
    ' 導凾数 df/dt
    '記号微分
    Df = 1
    '数値微分
    'Df = (f(t + h) - f(t - h)) / (2 * h)
End Function
Function Dg(t As Double) As Double
    ' 導凾数 dg/dt
    '記号微分
    'g = B(0, 2, s(t))
    '数値微分
    'Dg = (g(t + h) - g(t - h)) / (2 * h)
    Dg = (g(t + h) - g(t)) / h
End Function

Sub GraphDRAW0()
    Dim i As Integer
    Dim t1 As Double
    
    Dim crv As Curve
    Set crv = ActiveDocument.CreateCurve
    With crv.CreateSubPath(f(t_min), g(t_min))
        For i = 0 To n - 1
            t1 = t_min + (i + 1) * (t_max - t_min) / n
            .AppendCurveSegment f(t1), g(t1)
        Next i
    End With
    Dim s1 As Shape
    Set s1 = ActiveLayer.CreateCurve(crv)
    s1.Fill.ApplyNoFill
    s1.Outline.SetProperties 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
End Sub

Sub GraphDRAW1()
    Dim i As Integer
    Dim k As Double
    Dim t1 As Double
    Dim t0 As Double
    Dim tc As Double
    Dim a1 As Double
    Dim a2 As Double
    Dim a3 As Double
    Dim b1 As Double
    Dim b2 As Double
    Dim b3 As Double
    
    Dim crv As Curve
    Set crv = ActiveDocument.CreateCurve
    With crv.CreateSubPath(f(t_min), g(t_min))
        For i = 0 To n - 1
            t0 = t_min + i * (t_max - t_min) / n
            t1 = t_min + (i + 1) * (t_max - t_min) / n
            tc = (t0 + t1) / 2
            If ((Df(t0) - Df(t1)) ^ 2 + (Dg(t0) - Dg(t1)) ^ 2 = 0) Then
                a1 = f(t0) + (f(t1) - f(t0)) / 3
                b1 = g(t0) + (g(t1) - g(t0)) / 3
                a2 = f(t1) - (f(t1) - f(t0)) / 3
                b2 = g(t1) - (g(t1) - g(t0)) / 3
            Else
               ' ベジエ曲線最適化係数
                k = -4 * ((Df(t0) - Df(t1)) * (f(t0) + f(t1) - 2 * f(tc)) + (Dg(t0) - Dg(t1)) * (g(t0) + g(t1) - 2 * g(tc))) / ((Df(t0) - Df(t1)) ^ 2 + (Dg(t0) - Dg(t1)) ^ 2) / 3
                a1 = f(t0) + Df(t0) * k
                b1 = g(t0) + Dg(t0) * k
                a2 = f(t1) - Df(t1) * k
                b2 = g(t1) - Dg(t1) * k
            End If
            a3 = f(t1)
            b3 = g(t1)
            
            
            .AppendCurveSegment2 a3, b3, a1, b1, a2, b2
        Next i
    End With
    Dim s1 As Shape
    Set s1 = ActiveLayer.CreateCurve(crv)
    s1.Fill.ApplyNoFill
    s1.Outline.SetProperties 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
End Sub

Sub GraphDRAW3()
    Dim i As Integer
    Dim k As Double
    Dim l As Double
    Dim t0 As Double
    Dim t1 As Double
    Dim tc As Double
    Dim a1 As Double
    Dim a2 As Double
    Dim a3 As Double
    Dim b1 As Double
    Dim b2 As Double
    Dim b3 As Double
    
    Dim crv As Curve
    Set crv = ActiveDocument.CreateCurve
    With crv.CreateSubPath(f(t_min), g(t_min))
        For i = 0 To n - 1
            t0 = t_min + i * (t_max - t_min) / n
            t1 = t_min + (i + 1) * (t_max - t_min) / n
            tc = (t0 + t1) / 2
            ' ベジエ曲線最適化係数
            k = 4 * (-Dg(t1) * (f(t0) + f(t1) - 2 * f(tc)) + Df(t1) * (g(t0) + g(t1) - 2 * g(tc))) / (Df(t0) * Dg(t1) - Df(t1) * Dg(t0)) / 3
            l = 4 * (-Dg(t0) * (f(t0) + f(t1) - 2 * f(tc)) + Df(t0) * (g(t0) + g(t1) - 2 * g(tc))) / (Df(t0) * Dg(t1) - Df(t1) * Dg(t0)) / 3
            a1 = f(t0) + Df(t0) * k
            b1 = g(t0) + Dg(t0) * k
            a2 = f(t1) - Df(t1) * l
            b2 = g(t1) - Dg(t1) * l
            a3 = f(t1)
            b3 = g(t1)
            
            .AppendCurveSegment2 a3, b3, a1, b1, a2, b2
        Next i
    End With
    Dim s1 As Shape
    Set s1 = ActiveLayer.CreateCurve(crv)
    s1.Fill.ApplyNoFill
    s1.Outline.SetProperties 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
End Sub


