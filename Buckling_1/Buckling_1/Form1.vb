Imports System.IO
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
    Public _k As Double             'Ratio of edge stresses
    Public _α As Double             'aspect ratio

    '------- dimensions --------
    Public _t As Double         'Plate thickness [mm]
    Public _S As Double         'Stiffeners distance [mm]
    Public _L As Double         'stiffeners length [mm]
    Public _LG As Double        'girder length [mm]

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Calc_sequence()
    End Sub

    Private Sub Read_dimensions()
        _L = NumericUpDown1.Value   'Length
        _S = NumericUpDown2.Value   'stiffener distance
        _t = NumericUpDown3.Value   'plate thicknes

        _α = _L / _S                 'aspect ratio
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

        TextBox5.Text = Round(_σxmax, 0).ToString
        TextBox6.Text = Round(_σymax, 0).ToString
    End Sub
    Private Sub Read_properties()
        _σ0 = NumericUpDown14.Value
        _τ0 = _σ0 / Sqrt(3)
        _E = NumericUpDown13.Value
        _v = NumericUpDown12.Value
        _Pr = NumericUpDown16.Value
    End Sub
    'See page 29
    Private Sub Calc_chaper3_1_2()
        Dim σCi, σEi, Ks As Double

        _k = 1    'Ratio of edge stresses see page 29

        Ks = 1



        σEi = 1


        If (σEi < _Pr * _σ0) Then
            σCi = 1
        Else
            σCi = 1
        End If



        'TextBox1.Text = Round(τE, 0).ToString
        'TextBox2.Text = Round(_τC, 0).ToString
        'TextBox3.Text = Round(Ks, 0).ToString
        'TextBox4.Text = Round(_τ0, 0).ToString
    End Sub

    'See page 28
    Private Sub Calc_chaper3_1_1()
        Dim Ks, τE, C1 As Double

        C1 = 1.0    'For plate panels
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
        TextBox3.Text = Round(Ks, 0).ToString
        TextBox4.Text = Round(_τ0, 0).ToString
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, GroupBox1.Enter
        Calc_sequence()
    End Sub

    Private Sub Calc_sequence()
        Read_dimensions()
        Read_properties()
        Read_loads()
        Calc_chaper3_1_1()
        Calc_chaper3_1()
    End Sub

End Class
