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

    Public _σxmin As Double     'mimimum compressive stress in longitudinal direction
    Public _σymin As Double     'minimum compressive stress in transverse direction

    Public _σCx As Double       'maximum compressive stress in longitudinal direction
    Public _σCy As Double       'maximum compressive stress in transverse direction

    '-------- Material ----------------
    Public _σ0 As Double        'specified minimum yield point of plate [N/cm2]
    Public _τ0 As Double        'shear strength plate
    Public _E As Double         'Elasticity
    Public _v As Double         'Poissons ratio
    Public _Pr As Double        'Proportional lineair elastic limit for structure

    '-------- shear stress -------
    Public _τC As Double        'Edge shear critical stress
    Public _η As Double = 1     'max length utilizing factor

    '------- loads--------
    Public _σax, _σay As Double     'Uniforn in-plane compression
    Public _σbx, _σby As Double     'Uniforn in-plane bending
    Public _q As Double             'Uniform lateral load [N/cm2]
    Public _τ As Double             'Edge shear stress
    Public _kx As Double            'Ratio of edge stresses X direction
    Public _ky As Double            'Ratio of edge stresses Y direction
    Public _α As Double             'aspect ratio

    '------- dimensions --------
    Public _t As Double         'Plate thickness [mm]
    Public _S As Double         'Stiffeners distance [mm]
    Public _L As Double         'stiffeners length [mm]
    Public _LG As Double        'girder length [mm]
    Public _β As Double         'slenderness ratio [-]


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Calc_sequence()
    End Sub

    Private Sub Read_dimensions()
        _L = NumericUpDown1.Value   'Length
        _S = NumericUpDown2.Value   'stiffener distance
        _t = NumericUpDown3.Value   'plate thicknes

        _α = _L / _S                 'aspect ratio
        TextBox11.Text = Round(_α, 3).ToString
    End Sub
    Private Sub Read_loads()
        _q = NumericUpDown6.Value   'Uniform lateral load [N/cm2]
        _σax = NumericUpDown5.Value 'Comp stress in longit direction
        _σay = NumericUpDown4.Value 'Comp stress in travesre direction

        _σbx = NumericUpDown9.Value 'Bending stress in longit direction
        _σby = NumericUpDown8.Value 'Bending stress in traverse direction

        _τ = NumericUpDown7.Value   'Edge shear stress

        _σxmax = _σax + _σbx 'maximum compressive stress in longitudinal direction
        _σymax = _σay + _σby 'maximum compressive stress in transverse direction

        _σxmin = _σax - _σbx 'minimum compressive stress in longitudinal direction
        _σymin = _σay - _σby 'minimum compressive stress in transverse direction

        _kx = _σxmin / _σxmax
        _ky = _σymin / _σymax

        TextBox5.Text = Round(_σxmax, 0).ToString
        TextBox6.Text = Round(_σymax, 0).ToString

        TextBox7.Text = Round(_σxmin, 0).ToString
        TextBox8.Text = Round(_σymin, 0).ToString

        TextBox9.Text = Round(_kx, 1).ToString
        TextBox20.Text = Round(_kx, 1).ToString
        TextBox10.Text = Round(_ky, 1).ToString
        TextBox21.Text = Round(_ky, 1).ToString
    End Sub
    Private Sub Read_properties()
        _σ0 = NumericUpDown14.Value
        _τ0 = _σ0 / Sqrt(3)
        _E = NumericUpDown13.Value      'Elasticity
        _v = NumericUpDown12.Value      'Poissons
        _Pr = NumericUpDown16.Value     'Proportional lineair elastic limit of structure
    End Sub
    'See page 28
    Private Sub Calc_chaper3_1_1()
        Dim Ks, τE, C1 As Double

        Select Case True
            Case RadioButton1.Checked
                C1 = 1.1
            Case RadioButton2.Checked
                C1 = 1.0
            Case Else
                C1 = 1.0
        End Select

        Ks = (4.0 * (_S / _L) ^ 2 + 5.34) * C1

        τE = Ks * PI ^ 2 * _E / (12 * (1 - _v ^ 2))
        τE *= (_t / _S) ^ 2

        If (τE < _Pr * _τ0) Then
            _τC = τE
        Else
            _τC = _τ0 * (1 - _Pr * (1 - _Pr) * _τ0 / τE)
        End If

        TextBox1.Text = Round(τE, 0).ToString
        TextBox2.Text = Round(_τC, 0).ToString
        TextBox3.Text = Round(Ks, 2).ToString
        TextBox4.Text = Round(_τ0, 0).ToString
    End Sub
    'See page 29
    Private Sub Calc_chaper3_1_2()
        Dim σCix, σEix As Double
        Dim σCiy, σEiy As Double
        Dim Ksx_short, Ksx_long As Double
        Dim Ksy_short, Ksy_long As Double
        Dim C1, C2 As Double

        Select Case True
            Case RadioButton1.Checked
                C1 = 1.1
                C2 = 1.2
            Case RadioButton2.Checked
                C1 = 1.0
                C2 = 1.1
            Case Else
                C1 = 1.0
                C2 = 1.0
        End Select

        '============== X DIRECTION============================================
        '==============Loading applied along short edge========================
        If (_kx >= 0 And _kx <= 1) Then
            Ksx_short = C1 * 8.4 / (_kx + 1.1)
        Else
            Ksx_short = C1 * (7.6 - 6.4 * _kx + 10 * _kx ^ 2)
        End If

        '==============Loading applied along short edge========================
        If (_kx < 1 / 3) Then
            If (_α >= 1 And _α <= 2) Then
                Ksx_long = 24 / _α ^ 2
                Ksx_long += (1.0875 * (1 + 1 / _α ^ 2) ^ 2 - 18 / _α ^ 2) * (1 + _kx)
                Ksx_long *= C2
            Else
                Ksx_long = 12 / _α ^ 2
                Ksx_long += (1.0875 * (1 + 1 / _α ^ 2) ^ 2 - 9 / _α ^ 2) * (1 + _kx)
                Ksx_long *= C2
            End If
        Else
            Ksx_long = (1 + 1 / _α ^ 2) ^ 2 * (1.675 - 0.675 * _kx)
            Ksx_long *= C2
        End If

        '==============Elastic buckling stress=============================

        σEix = Ksx_long * (_t / _S) ^ 2 * (PI ^ 2 * _E) / (12 * (1 - _v ^ 2))

        '==============Critical buckling stress=============================
        If (σEix < _Pr * _σ0) Then
            σCix = σEix
        Else
            σCix = _σ0 * (1 - _Pr * (1 - _Pr) * _σ0 / σEix)
        End If


        TextBox18.Text = Round(Ksx_long, 2).ToString
        TextBox13.Text = Round(σEix, 0).ToString
        TextBox14.Text = Round(σCix, 0).ToString

        '============== Y DIRECTION============================================
        '==============Loading applied along short edge========================
        If (_ky >= 0 And _ky <= 1) Then
            Ksy_short = C1 * 8.4 / (_kx + 1.1)
        Else
            Ksy_short = C1 * (7.6 - 6.4 * _kx + 10 * _kx ^ 2)
        End If

        '==============Loading applied along short edge========================
        If (_kx < 1 / 3) Then
            If (_α >= 1 And _α <= 2) Then
                Ksy_long = 24 / _α ^ 2
                Ksy_long += (1.0875 * (1 + 1 / _α ^ 2) ^ 2 - 18 / _α ^ 2) * (1 + _kx)
                Ksy_long *= C2
            Else
                Ksy_long = 12 / _α ^ 2
                Ksy_long += (1.0875 * (1 + 1 / _α ^ 2) ^ 2 - 9 / _α ^ 2) * (1 + _kx)
                Ksy_long *= C2
            End If
        Else
            Ksy_long = (1 + 1 / _α ^ 2) ^ 2 * (1.675 - 0.675 * _kx)
            Ksy_long *= C2
        End If

        '==============Elastic buckling stress=============================
        σEiy = Ksy_short * (_t / _S) ^ 2 * (PI ^ 2 * _E) / (12 * (1 - _v ^ 2))

        '==============Critical buckling stress=============================
        If (σEiy < _Pr * _σ0) Then
            σCiy = σEiy
        Else
            σCiy = _σ0 * (1 - _Pr * (1 - _Pr) * _σ0 / σEiy)
        End If

        TextBox17.Text = Round(Ksy_short, 2).ToString
        TextBox16.Text = Round(σEiy, 0).ToString
        TextBox15.Text = Round(σCiy, 0).ToString
    End Sub

    'See page 30
    Private Sub Calc_chaper3_3()
        Dim τu, σux, σuy, Cy, Cx As Double

        _β = _S / _t * Sqrt(_σ0 / _E)   'Slenderness ratio
        _η = 0.6                        'See page 2

        'Utimate strength
        τu = _τC + 0.5 * (_σ0 - Sqrt(3 * _τC)) / (1 + _α + _α ^ 2) ^ 0.5
        If τu < _τC Then τu = _τC

        '-------------- Cx-----------------
        If _β > 1 Then
            Cx = 2 / _β - (1 / _β ^ 2)
        Else
            Cx = 1
        End If

        '-------------- Cy-----------------
        Cy = Cx * _S / _L + 0.1 * (1 - _S / _L) * (1 + 1 / _β ^ 2) ^ 2
        If Cy > 1 Then Cy = 1

        '-------- σux -------------
        σux = Cx * _σ0
        If σux < _σCx Then σux = _σCx

        '-------- σuy -------------
        σuy = Cy * _σ0
        If σuy < _σCy Then σux = _σCy

        '-------- φ (phi)-------------


        TextBox12.Text = Round(_β, 2).ToString
        TextBox19.Text = Round(_η, 2).ToString
        TextBox22.Text = Round(_η, 2).ToString
        TextBox23.Text = Round(Cx, 2).ToString
        TextBox24.Text = Round(Cy, 2).ToString
        TextBox25.Text = Round(τu, 0).ToString
        TextBox26.Text = Round(σux, 0).ToString
        TextBox27.Text = Round(σuy, 0).ToString
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.Enter, NumericUpDown8.Enter, NumericUpDown7.Enter, NumericUpDown6.Enter, NumericUpDown5.Enter, NumericUpDown4.Enter, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged
        Read_loads()
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, GroupBox1.Enter, NumericUpDown3.Enter, NumericUpDown3.Click, NumericUpDown2.Enter, NumericUpDown2.Click, NumericUpDown1.Enter, NumericUpDown1.Click
        Calc_sequence()
    End Sub

    Private Sub Calc_sequence()
        Read_dimensions()
        Read_properties()
        Read_loads()
        Calc_chaper3_1_1()
        Calc_chaper3_1_2()
        Calc_chaper3_3()
        Calc_chaper3_1()
    End Sub

End Class
