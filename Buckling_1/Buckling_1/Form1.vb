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
    Public _qu As Double            'Ultimate lateral load [N/cm2]
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

    Public _Iw As Double    'Moment of inertia of stiffener and effective plating sw
    Public _Ie As Double    'Moment Of inertia Of stiffener And effective plating se

    Public _se As Double    'Plate effective width (plate buckling limit must be satiesfied)
    Public _sw As Double    'Plate effective breadth

    Public _re As Double    'Radius of giration od Area Ae
    Dim _φ As Double        'Interaction between longitudinal and transverse sresses

    Public _σUx As Double   'Ultimate stress 
    Public _σUy As Double   'Ultimate stress 
    Public _smw As Double      'Effective section modulus of longitidunal at flange

    Public _A As Double     'Total plate + stiffener (in ABS called "A")
    Public _Ast As Double   'Area stiffener (in ABS called "As")
    Public _Ae As Double    'Effective area plate + area stiffener
    Public _Aw As Double    'Effective breadth area plate + area stiffener

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
        _σbx = NumericUpDown9.Value 'Bending stress in longit direction

        _σay = NumericUpDown4.Value 'Comp stress in tranvesre direction
        _σby = NumericUpDown8.Value 'Bending stress in traverse direction

        _τ = NumericUpDown7.Value   'Edge shear stress

        '------- X direction---------
        _σxmax = _σax + _σbx 'maximum compressive stress in longitudinal direction
        _σxmin = _σax - _σbx 'minimum compressive stress in longitudinal direction

        '------- Y direction---------
        _σymax = _σay + _σby 'maximum compressive stress in transverse direction
        _σymin = _σay - _σby 'minimum compressive stress in transverse direction

        _kx = _σxmin / _σxmax
        _ky = _σymin / _σymax

        TextBox5.Text = Round(_σxmax, 0).ToString
        TextBox6.Text = Round(_σymax, 0).ToString

        TextBox7.Text = Round(_σxmin, 0).ToString
        TextBox8.Text = Round(_σymin, 0).ToString

        TextBox9.Text = Round(_kx, 2).ToString("0.00")
        TextBox20.Text = Round(_kx, 2).ToString("0.00")
        TextBox10.Text = Round(_ky, 2).ToString("0.00")
        TextBox21.Text = Round(_ky, 2).ToString("0.00")
    End Sub
    Private Sub Read_properties()
        _σ0 = NumericUpDown14.Value
        _τ0 = _σ0 / Sqrt(3)
        _E = NumericUpDown13.Value      'Elasticity
        _v = NumericUpDown12.Value      'Poissons
        _Pr = NumericUpDown16.Value     'Proportional lineair elastic limit of structure
        _η = IIf(RadioButton4.Checked, 0.6, 0.8)    'See page 2
    End Sub
    'See page 28 of the ABS guide for Buckling
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
    'See page 29 of the ABS guide for Buckling
    Private Sub Calc_chaper3_1_2()
        Dim σCix, σEix As Double
        Dim σCiy, σEiy As Double
        Dim Ksx_long As Double
        Dim Ksy_short As Double
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

        '==============Loading applied along long edge========================
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

        TextBox18.Text = Round(Ksx_long, 2).ToString("0.00")
        TextBox13.Text = Round(σEix, 0).ToString
        TextBox14.Text = Round(σCix, 0).ToString

        '============== Y DIRECTION============================================
        '==============Loading applied along short edge========================
        If (_ky >= 0 And _ky <= 1) Then
            Ksy_short = C1 * 8.4 / (_kx + 1.1)
        Else
            Ksy_short = C1 * (7.6 - 6.4 * _kx + 10 * _kx ^ 2)
        End If

        '==============Elastic buckling stress=============================
        σEiy = Ksy_short * (_t / _S) ^ 2 * (PI ^ 2 * _E) / (12 * (1 - _v ^ 2))

        '==============Critical buckling stress=============================
        If (σEiy < _Pr * _σ0) Then
            σCiy = σEiy
        Else
            σCiy = _σ0 * (1 - _Pr * (1 - _Pr) * _σ0 / σEiy)
        End If

        TextBox17.Text = Round(Ksy_short, 2).ToString("0.00")
        TextBox16.Text = Round(σEiy, 0).ToString
        TextBox15.Text = Round(σCiy, 0).ToString
    End Sub

    'See page 30 of the ABS guide for Buckling
    Private Sub Calc_chaper3_3()
        Dim τu, Cy, Cx As Double
        Dim strength_criterium As Double

        _β = _S / _t * Sqrt(_σ0 / _E)   'Slenderness ratio

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
        _σUx = Cx * _σ0
        If _σUx < _σCx Then _σUx = _σCx

        '-------- σuy -------------
        _σUy = Cy * _σ0
        If _σUy < _σCy Then _σUx = _σCy

        '-------- φ (phi)-------------
        _φ = 1 - _β / 2

        '--------- strength criterium----------------
        strength_criterium = (_σxmax / (_η * _σUx)) ^ 2
        strength_criterium -= _φ * (_σxmax / (_η * _σUx)) * (_σymax / (_η * _σUy))
        strength_criterium += (_σymax / (_η * _σUx)) ^ 2
        strength_criterium += (_τ / (_η * τu)) ^ 2

        Label112.Text = IIf(strength_criterium <= 1, "Plate buckling state limit  OK", "Plate buckling state limit NOK")
        Label112.BackColor = IIf(strength_criterium <= 1, Color.LightGreen, Color.Coral)

        TextBox12.Text = Round(_β, 2).ToString
        TextBox19.Text = Round(_η, 2).ToString
        TextBox22.Text = Round(_η, 2).ToString
        TextBox23.Text = Round(Cx, 2).ToString
        TextBox24.Text = Round(Cy, 2).ToString
        TextBox25.Text = Round(τu, 0).ToString
        TextBox26.Text = Round(_σUx, 0).ToString
        TextBox27.Text = Round(_σUy, 0).ToString
        TextBox28.Text = Round(_φ, 2).ToString
        TextBox29.Text = Round(strength_criterium, 4).ToString
    End Sub
    'See page 31 of the ABS guide for Buckling
    Private Sub Calc_chaper3_5()
        Dim σe As Double

        σe = Sqrt(_σxmax ^ 2 - _σxmax * _σymax + _σymax ^ 2 + 3 * _τ ^ 2) 'Von misses
        _qu = _η * 4.0 * _σ0 * (_t / _S) ^ 2 * (1 + 1 / _α ^ 2) * Sqrt(1 - (σe / _σ0) ^ 2)

        TextBox31.Text = Round(σe, 1).ToString
        TextBox30.Text = Round(_qu, 3).ToString
        TextBox30.BackColor = IIf(_qu >= _q, Color.LightGreen, Color.Coral)
        NumericUpDown6.BackColor = TextBox30.BackColor
    End Sub
    'See page 45 of the ABS guide for Buckling
    Private Sub Calc_chaper13_1()
        Dim dw, tw, bf, tf, b1, Zep, Zwp As Double

        bf = NumericUpDown18.Value
        tf = NumericUpDown15.Value
        b1 = NumericUpDown17.Value
        dw = NumericUpDown11.Value
        tw = NumericUpDown10.Value

        _se = _S     'Page 32, Buckling state limit is satisfied
        _sw = _se    'Effective breadth figure 8.

        _Ast = dw * tw + bf * tf    'Area stiffener
        _Ae = _Ast + _se * _t       'Area stiffener + Area plate
        _Aw = _Ast + _sw * _t       'Area stiffener + Area plate
        _A = _Ast + _S * _t         'Area stiffener + Area plate

        Zep = (0.5 * ((_t + dw) * dw * tw) + (0.5 * _t + dw + 0.5 * tf) * bf * tf) / _Ae

        _Ie = _t ^ 3 * _se / 12
        _Ie += dw ^ 3 * tw / 12
        _Ie += tf ^ 3 * bf / 12
        _Ie += 0.25 * (_t + dw) ^ 2 * dw * tw
        _Ie += bf * tf * (0.5 * _t + dw + 0.5 * tf) ^ 2
        _Ie -= _Ae * Zep ^ 2

        _re = Sqrt(_Ie / _Ae)

        Zwp = (0.5 * ((_t + dw) * dw * tw) + (0.5 * _t + dw + 0.5 * tf) * bf * tf) / _Aw

        _Iw = _t ^ 3 * _se / 12
        _Iw += dw ^ 3 * tw / 12
        _Iw += tf ^ 3 * bf / 12
        _Iw += 0.25 * (_t + dw) ^ 2 * dw * tw
        _Iw += bf * tf * (0.5 * _t + dw + 0.5 * tf) ^ 2
        _Iw -= _Ae * Zep ^ 2

        _smw = _Iw / ((0.5 * _t + dw + tf) - Zwp)

        TextBox32.Text = Round(_se, 1).ToString
        TextBox33.Text = Round(_sw, 1).ToString
        TextBox34.Text = Round(_Ast, 1).ToString
        'TextBox35.Text = Round(A_tot, 1).ToString
        TextBox36.Text = Round(_Ae, 1).ToString
        TextBox37.Text = Round(_Aw, 1).ToString
        TextBox38.Text = Round(Zep, 1).ToString
        TextBox39.Text = Round(_Ie, 1).ToString
        TextBox40.Text = Round(_re, 1).ToString
        TextBox41.Text = Round(Zwp, 1).ToString
        TextBox42.Text = Round(_Iw, 1).ToString
        TextBox43.Text = Round(_smw, 1).ToString
    End Sub
    'See page 32 of the ABS guide for Buckling
    Private Sub Calc_chaper5_1()
        Dim σa, σEC, σCA, Cm As Double
        Dim σb, M As Double
        Dim Cy, Cx, Cxy As Double
        Dim bsl_crit As Double  'Buckling state limit criterium  

        '-----------Maximum bending moment-----------
        M = _q * _S * _L ^ 2 / 12
        σb = M / _smw

        '-------- Cx, Cy and Cxy------------
        Cxy = Sqrt(1 - (_τ / _τ0) ^ 2)
        Cy = 0.5 * _φ * (_σymax / _σUy) + Sqrt(1 - (1 - 0.25 * _φ ^ 2) * (_σymax / _σUy) ^ 2)

        If _β > 1 Then
            Cx = 2 / _β - 1 / _β ^ 2
        Else
            Cx = 1
        End If

        '------------- σE(C)----------
        σEC = PI ^ 2 * _E * _re ^ 2 / _L ^ 2

        '-------------σCA-------------
        If (σEC < _Pr * _σ0) Then
            σCA = σEC
        Else
            σCA = _σ0 * (1 - _Pr * (1 - _Pr) * (_σ0 / σEC))
        End If

        '-----------P----------------
        σa = _σax
        Cm = 0.75

        bsl_crit = σa / (_η * σCA * (_Ae / _A))
        bsl_crit += Cm * σb / (_η * _σ0 * (1 - σa / (_η * σEC)))

        TextBox44.Text = Round(σa, 0).ToString
        TextBox45.Text = Round(σEC, 0).ToString
        TextBox46.Text = Round(_Pr, 1).ToString
        TextBox47.Text = Round(σCA, 0).ToString
        TextBox48.Text = Round(M, 0).ToString
        TextBox49.Text = Round(σb, 0).ToString
        TextBox50.Text = Round(Cm, 2).ToString
        TextBox51.Text = Round(_η, 1).ToString
        TextBox52.Text = Round(bsl_crit, 2).ToString
        'TextBox53.Text = Round(_re, 1).ToString
        'TextBox54.Text = Round(_re, 1).ToString
        'TextBox55.Text = Round(_re, 1).ToString

        TextBox56.Text = Round(Cx, 1).ToString
        TextBox57.Text = Round(Cy, 1).ToString
        TextBox58.Text = Round(Cxy, 1).ToString

        '----- checks-----------
        TextBox52.BackColor = IIf(bsl_crit <= 1, Color.LightGreen, Color.Coral)
    End Sub
    'See page 35 of the ABS guide for Buckling
    Private Sub Calc_chaper5_3()
        Dim σcL, Ixf, Iz, Iy, Γ, Co, uf, m, n, Io As Double


        uf = 1

        '--------------
        n = 1
        σcL = PI ^ 2 * _E * (n / _α + _α / n) ^ 2 * (_t / _S) ^ 2
        σcL /= (12 * (1 - _v ^ 2))


        Ixf = 1

        Γ = 1
        Co = 1

        m = 1
        Io = 1


        TextBox59.Text = Round(σcL, 0).ToString
        TextBox60.Text = Round(Ixf, 2).ToString
        TextBox61.Text = Round(Γ, 1).ToString
        TextBox62.Text = Round(Co, 2).ToString
        TextBox63.Text = Round(uf, 1).ToString
        TextBox64.Text = Round(m, 1).ToString
        TextBox65.Text = Round(Iy, 1).ToString
        TextBox66.Text = Round(Iz, 1).ToString
        TextBox67.Text = Round(Io, 1).ToString
        TextBox68.Text = Round(n, 1).ToString
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, TabPage4.Enter
        Calc_sequence()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.Enter, NumericUpDown8.Enter, NumericUpDown7.Enter, NumericUpDown6.Enter, NumericUpDown5.Enter, NumericUpDown4.Enter, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged
        Calc_sequence()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown18.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged
        Calc_sequence()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabPage9.Enter
        Calc_sequence()
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
        Calc_chaper3_5()
        Calc_chaper13_1()
        Calc_chaper5_1()
        Calc_chaper5_3()
    End Sub

End Class
