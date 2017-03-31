﻿Imports System.IO
Imports System.Text
Imports System.Math

'Based on American Bureau of Shipping (ABS)
'Guide for Buckling and Ultimate strength assessment for offshore structures
'Updated February 2014

Public Class Form1
    Public _Ψ As Double = 1     'Adjustment factor (psi)
    Public _σxmax As Double     'maximum compressive stress in longitudinal direction
    Public _σymax As Double     'maximum compressive stress in transverse direction
    Public _σCx As Double       'maximum compressive stress in longitudinal direction
    Public _σCy As Double       'maximum compressive stress in transverse direction

    '-------- yield points ----------------
    Public _σ0 As Double        'specified minimum yield point of plate [N/cm2]
    Public _τ0 As Double        'shear strength plate

    '-------- shear stress -------
    Public _τ As Double         'Edge shear stress
    Public _τC As Double        'Edge shear stress

    Public _η As Double = 1     'max length utilizing factor

    '------- loads--------
    Public _q As Double         'Uniform lateral load [N/cm2]
    Public _t As Double         'Plate thickness [mm]
    Public _S As Double         'Stiffeners distance [mm]
    Public _L As Double         'stiffeners length [mm]
    Public _LG As Double        'girder length [mm]

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Read_material()
    End Sub

    Private Sub Read_material()
        _L = NumericUpDown1.Value   'Length
        _S = NumericUpDown2.Value   'stiffener distance
        _t = NumericUpDown3.Value   'plate thicknes

    End Sub
    Private Sub Read_loads()
        _q = NumericUpDown6.Value   'Uniform lateral load [N/cm2]
        _S = NumericUpDown5.Value   'Comp stress in longit direction
        _t = NumericUpDown4.Value   'plate thicknes

    End Sub
    Private Sub Read_properties()
        _σ0 = NumericUpDown14.Value
        _τ0 = _σ0 / Sqrt(3)
    End Sub

    'See page 27
    Private Sub Calc_chaper3_1()
        Dim strength_criterium As Double

        strength_criterium = (_σxmax / (_η * _σCx)) ^ 2 + (_σymax / (_η * _σCy)) ^ 2 + (_τ / (_η * _τC)) ^ 2
        Label112.Text = IIf(strength_criterium <= 1, "Plate buckling state limit  OK", "Plate buckling state limit NOK")
        Label112.BackColor = IIf(strength_criterium <= 1, Color.Green, Color.Red)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, TabPage4.Enter
        Calc_sequence()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, GroupBox1.Enter
        Calc_sequence()
    End Sub

    Private Sub Calc_sequence()
        Read_material()
        Read_properties()
        Read_loads()
        Calc_chaper3_1()
    End Sub

End Class
