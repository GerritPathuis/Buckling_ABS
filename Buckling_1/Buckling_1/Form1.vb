Imports System.IO
Imports System.Text
Imports System.Math
Imports Word = Microsoft.Office.Interop.Word

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

    '------- dimensions overall (figure1) --------
    Public _t As Double         'Plate thickness [mm]
    Public _S As Double         'Stiffeners distance [mm]
    Public _L As Double         'stiffeners length [mm]
    Public _LG As Double        'girder length [mm]
    Public _β As Double         'slenderness ratio [-]

    '------- dimensions stiffener (figure 2)--------
    Public _dw As Double        'Web depth
    Public _tw As Double
    Public _bf As Double
    Public _tf As Double
    Public _b1 As Double
    Public _b2 As Double

    Public _y0 As Double        'Centriod to center line web [cm]
    Public _z0 As Double        'Centroid to plate [cm]

    Public _Iw As Double    'Moment of inertia of stiffener and effective plating sw
    Public _Ie As Double    'Moment Of inertia Of stiffener And effective plating se
    Public _Iy As Double    'Moment of Inertia of stiffener, trough centroid, excl plating [cm4]
    Public _Iz As Double    'Moment of Inertia of stiffener, trough centroid, excl plating [cm4]

    Public _se As Double    'Plate effective width (plate buckling limit must be satiesfied)
    Public _sw As Double    'Plate effective breadth

    Public _re As Double    'Radius of giration od Area Ae
    Dim _φ As Double        'Interaction between longitudinal and transverse sresses

    Public _σUx As Double   'Ultimate stress 
    Public _σUy As Double   'Ultimate stress 
    Public _smw As Double      'Effective section modulus of longitidunal at flange

    Public _A As Double     'Total plate + stiffener (in ABS named "A")
    Public _As As Double    'Area stiffener (in ABS called "As")
    Public _Ae As Double    'Effective area plate + area stiffener
    Public _Aw As Double    'Effective breadth area plate + area stiffener

    Dim all_textbox As New List(Of Control)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Calc_sequence()
    End Sub

    Private Sub Read_dimensions()

        _L = NumericUpDown1.Value   'Length
        _S = NumericUpDown2.Value   'stiffener distance
        _t = NumericUpDown3.Value   'plate thicknes

        _α = _L / _S                 'aspect ratio
        TextBox11.Text = Round(_α, 3).ToString

        _bf = NumericUpDown18.Value
        _tf = NumericUpDown15.Value
        _b1 = NumericUpDown17.Value
        _dw = NumericUpDown11.Value 'Web depth
        _tw = NumericUpDown10.Value

        Select Case True
            Case RadioButton9.Checked   'Angle iron
                _b1 = _tw / 2
                '---- minimale breedte flange 20mm---
                If _bf <= 2 Then
                    _bf = 2
                End If
                If _b2 <= _b1 Then
                    _b2 = _b1
                End If
            Case RadioButton10.Checked  'Tee iron
                '---- minimale breedte flange 40mm---
                If _bf <= 4 Then
                    _bf = 4
                End If
                _b1 = _bf / 2
            Case RadioButton11.Checked  'Flat bar
                _b1 = _tw / 2
                _bf = 0
                '_tf = 0

        End Select
        NumericUpDown15.Value = CDec(_tf)
        NumericUpDown17.Value = CDec(_b1)
        NumericUpDown18.Value = CDec(_bf)

        _b2 = _bf - _b1 - _tw / 2       'Larger outstanding flange
        TextBox90.Text = Round(_b2, 2).ToString("0.00")

        _As = _dw * _tw + _bf * _tf     'Area stiffener

        _y0 = Abs((_b1 - 0.5 * _bf) * _bf * _tf / _As)
        _z0 = (0.5 * _dw ^ 2 * _tw + (_dw + 0.5 * _tf) * _bf * _tf) / _As

        _Iy = _dw ^ 3 * _tw / 12
        _Iy += _tf ^ 3 * _bf / 12
        _Iy += 0.25 * _dw ^ 3 * _tw
        _Iy += _bf * _tf * (_dw + 0.5 * _tf) ^ 2
        _Iy -= _As * _z0 ^ 2

        _Iz = _tw ^ 3 * _dw / 12
        _Iz += _bf ^ 3 * _tf / 12
        _Iz += _bf * _tf * (_b1 - 0.5 * _bf) ^ 2
        _Iz -= _As * _z0 ^ 2

        TextBox72.Text = Round(_As, 2).ToString
        TextBox73.Text = Round(_y0, 2).ToString
        TextBox74.Text = Round(_z0, 2).ToString

        TextBox75.Text = Round(_Iy, 0).ToString
        TextBox65.Text = Round(_Iy, 0).ToString

        TextBox76.Text = Round(_Iz, 0).ToString
        TextBox66.Text = Round(_Iz, 0).ToString

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
        TextBox78.Text = Round(_q * 100, 1).ToString("0.0")    '[mbar]

        TextBox84.Text = Round(_σax / 100, 1).ToString         '[N/mm2]
        TextBox85.Text = Round(_σay / 100, 1).ToString         '[N/mm2]
    End Sub
    Private Sub Read_properties()
        _σ0 = NumericUpDown14.Value
        _τ0 = _σ0 / Sqrt(3)
        _E = NumericUpDown13.Value      'Elasticity
        _v = NumericUpDown12.Value      'Poissons
        _Pr = NumericUpDown16.Value     'Proportional lineair elastic limit of structure
        _η = CDbl(IIf(RadioButton4.Checked, 0.6, 0.8))    'See page 2
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

        '----------- Check----------------------
        TextBox1.BackColor = CType(IIf(τE > _τ, Color.LightGreen, Color.Coral), Color)
        TextBox2.BackColor = CType(IIf(_τC > _τ, Color.LightGreen, Color.Coral), Color)
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
        '----------- Check----------------------
        TextBox13.BackColor = CType(IIf(σEix > _σxmax, Color.LightGreen, Color.Coral), Color)
        TextBox14.BackColor = CType(IIf(σCix > _σxmax, Color.LightGreen, Color.Coral), Color)

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
        '----------- Check----------------------
        TextBox16.BackColor = CType(IIf(σEiy > _σymax, Color.LightGreen, Color.Coral), Color)
        TextBox15.BackColor = CType(IIf(σCiy > _σymax, Color.LightGreen, Color.Coral), Color)
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

        TextBox12.Text = Round(_β, 2).ToString("0.00")
        TextBox19.Text = Round(_η, 2).ToString("0.00")
        TextBox22.Text = Round(_η, 2).ToString("0.00")
        TextBox23.Text = Round(Cx, 2).ToString("0.00")
        TextBox24.Text = Round(Cy, 2).ToString("0.00")
        TextBox25.Text = Round(τu, 0).ToString
        TextBox26.Text = Round(_σUx, 0).ToString
        TextBox27.Text = Round(_σUy, 0).ToString
        TextBox28.Text = Round(_φ, 2).ToString
        TextBox29.Text = Round(strength_criterium, 4).ToString
        '----------- Check----------------------
        TextBox29.BackColor = CType(IIf(strength_criterium <= 1, Color.LightGreen, Color.Coral), Color)
    End Sub
    'See page 31 of the ABS guide for Buckling, Uniform Lateral pressure
    Private Sub Calc_chaper3_5()
        Dim σe, crit_lateral_press As Double

        σe = Sqrt(_σxmax ^ 2 - _σxmax * _σymax + _σymax ^ 2 + 3 * _τ ^ 2) 'Von misses
        _qu = _η * 4.0 * _σ0 * (_t / _S) ^ 2 * (1 + 1 / _α ^ 2) * Sqrt(1 - (σe / _σ0) ^ 2)

        crit_lateral_press = _q / _qu

        TextBox31.Text = Round(σe, 1).ToString
        TextBox30.Text = Round(_qu, 3).ToString
        TextBox82.Text = Round(_q, 3).ToString
        TextBox83.Text = Round(crit_lateral_press, 2).ToString

        TextBox83.BackColor = CType(IIf(crit_lateral_press < 1, Color.LightGreen, Color.Coral), Color)
        NumericUpDown6.BackColor = TextBox30.BackColor
    End Sub
    'See page 45 of the ABS guide for Buckling
    Private Sub Calc_chaper13_1()
        Dim Zep, Zwp As Double

        _se = _S     'Page 32, Buckling state limit is satisfied
        _sw = _se    'Effective breadth figure 8.

        _Ae = _As + _se * _t       'Area stiffener + Area plate
        _Aw = _As + _sw * _t       'Area stiffener + Area plate
        _A = _As + _S * _t         'Area stiffener + Area plate

        Zep = 0.5 * (_t + _dw) * _dw * _tw
        Zep += (0.5 * _t + _dw + 0.5 * _tf) * _bf * _tf
        Zep /= _Ae

        _Ie = _t ^ 3 * _se / 12
        _Ie += _dw ^ 3 * _tw / 12
        _Ie += _tf ^ 3 * _bf / 12
        _Ie += 0.25 * (_t + _dw) ^ 2 * _dw * _tw
        _Ie += _bf * _tf * (0.5 * _t + _dw + 0.5 * _tf) ^ 2
        _Ie -= _Ae * Zep ^ 2

        _re = Sqrt(_Ie / _Ae)

        Zwp = 0.5 * (_t + _dw) * _dw * _tw
        Zwp += 0.5 * _t + _dw + 0.5 * _tf * _bf * _tf
        Zwp /= _Aw

        _Iw = _t ^ 3 * _se / 12
        _Iw += _dw ^ 3 * _tw / 12
        _Iw += _tf ^ 3 * _bf / 12
        _Iw += 0.25 * (_t + _dw) ^ 2 * _dw * _tw
        _Iw += _bf * _tf * (0.5 * _t + _dw + 0.5 * _tf) ^ 2
        _Iw -= _Ae * Zep ^ 2

        _smw = _Iw / ((0.5 * _t + _dw + _tf) - Zwp)

        TextBox32.Text = Round(_se, 1).ToString
        TextBox33.Text = Round(_sw, 1).ToString
        TextBox34.Text = Round(_As, 1).ToString
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
        TextBox51.Text = Round(_η, 1).ToString("0.00")
        TextBox52.Text = Round(bsl_crit, 2).ToString

        TextBox56.Text = Round(Cx, 1).ToString
        TextBox57.Text = Round(Cy, 1).ToString
        TextBox58.Text = Round(Cxy, 1).ToString

        '----- checks-----------
        TextBox52.BackColor = CType(IIf(bsl_crit <= 1, Color.LightGreen, Color.Coral), Color)
    End Sub
    'See page 35 of the ABS guide for Buckling
    Private Sub Calc_chaper5_3()
        Dim σa, σcL, σET, σCT, Ixf, Γ, Co, uf, m, n, Io, K, flex_crit As Double

        uf = 1 - 2 * _b1 / _bf                  'Unsymatrical factor
        m = 1 - uf * (0.7 - 0.1 * _dw / _bf)

        '--------------
        n = 1       'No half waves
        σcL = PI ^ 2 * _E * (n / _α + _α / n) ^ 2 * (_t / _S) ^ 2
        σcL /= (12 * (1 - _v ^ 2))

        Ixf = _tf * _bf ^ 3 / 12 * (1.0 + 3.0 * uf ^ 2 * _dw * _tw / _As)
        Γ = m * Ixf * _dw ^ 2 + _dw ^ 3 * _tw ^ 2 / 36.0
        Co = _E * _t ^ 3 / (3.0 * _S)
        Io = _Iy + m * _Iz + _As * (_y0 ^ 2 + _z0 ^ 2)
        K = (_bf * _tf ^ 3 + _dw * _tw ^ 3) / 3

        σET = K / 2.6 + (n * PI / _L) ^ 2 * Γ
        σET += Co / _E * (_L / (n * PI)) ^ 2
        σET *= _E
        σET /= Io + Co / σcL * (_L / (n * PI)) ^ 2

        If σET <= _Pr * _σ0 Then
            σCT = σET
        Else
            σCT = _σ0 * (1 - _Pr * (1 - _Pr) * _σ0 / σET)
        End If

        σa = _σax
        flex_crit = σa / (_η * σCT)

        TextBox59.Text = Round(σcL, 0).ToString
        TextBox60.Text = Round(Ixf, 2).ToString
        TextBox61.Text = Round(Γ, 1).ToString
        TextBox62.Text = Round(Co, 0).ToString
        TextBox63.Text = Round(uf, 1).ToString
        TextBox64.Text = Round(m, 1).ToString
        TextBox67.Text = Round(Io, 1).ToString
        TextBox68.Text = Round(n, 1).ToString("0.00")
        TextBox77.Text = Round(K, 1).ToString
        TextBox69.Text = Round(σET, 0).ToString
        TextBox70.Text = Round(σCT, 0).ToString
        TextBox71.Text = Round(flex_crit, 2).ToString
        '----- checks-----------
        TextBox71.BackColor = CType(IIf(flex_crit <= 1, Color.LightGreen, Color.Coral), Color)
    End Sub
    'See page 37 of the ABS guide for Buckling, local buckling of web
    Private Sub Calc_chaper5_5_1()
        Dim Ksw, τEw, Cs, τCw As Double

        'Web depth is _dw
        Select Case True
            Case RadioButton9.Checked Or RadioButton10.Checked
                Cs = 1.0    'Angle of Tee bar
            Case RadioButton11.Checked
                Cs = 0.11   'Flat bar
            Case Else
                Cs = 0.00099999 'Will never happen
        End Select

        Ksw = 4.0 * Cs

        τEw = Ksw * PI ^ 2 * _E / (12 * (1 - _v ^ 2))
        τEw *= (_t / _S) ^ 2

        If (τEw < _Pr * _τ0) Then
            τCw = τEw
        Else
            τCw = _τ0 * (1 - _Pr * (1 - _Pr) * _τ0 / τEw)
        End If

        TextBox91.Text = Round(Cs, 0).ToString
        TextBox94.Text = Round(_τ0, 0).ToString
        TextBox95.Text = Round(Ksw, 2).ToString("0.00")
        TextBox96.Text = Round(τEw, 0).ToString
        TextBox97.Text = Round(τCw, 0).ToString

        '-------------- checks------------
        TextBox96.BackColor = CType(IIf(τEw > _τ, Color.LightGreen, Color.Coral), Color)
        TextBox97.BackColor = CType(IIf(τCw > _τ, Color.LightGreen, Color.Coral), Color)
    End Sub
    Private Sub Calc_chaper5_5_2()
        Dim Ks, σC_local, σE_local As Double

        '==============Loading applied along long edge========================
        Ks = 0.44

        '==============Elastic buckling stress=============================
        σE_local = Ks * (_t / _b2) ^ 2 * (PI ^ 2 * _E) / (12 * (1 - _v ^ 2))

        '==============Critical buckling stress=============================
        If (σE_local < _Pr * _σ0) Then
            σC_local = σE_local
        Else
            σC_local = _σ0 * (1 - _Pr * (1 - _Pr) * _σ0 / σE_local)
        End If

        TextBox86.Text = Round(Ks, 2).ToString("0.00")
        TextBox88.Text = Round(σE_local, 0).ToString
        TextBox87.Text = Round(σC_local, 0).ToString
        TextBox89.Text = Round(_b2, 1).ToString("0.0")

        '----------- Check----------------------
        TextBox88.BackColor = CType(IIf(σE_local > _σxmax, Color.LightGreen, Color.Coral), Color)
        TextBox87.BackColor = CType(IIf(σC_local > _σxmax, Color.LightGreen, Color.Coral), Color)
    End Sub
    Private Sub Calc_chaper5_7()
        Dim Ix, Iy As Double
        Dim Ipx, Ipy As Double
        Dim Asx, Asy As Double

        Ix = _Iz
        Iy = _Iy

        Ipx = 3
        Ipy = 4

        Asx = 5
        Asy = 6

        TextBox93.Text = Round(Ix, 0).ToString
        TextBox108.Text = Round(Iy, 0).ToString
        TextBox92.Text = Round(Ipx, 0).ToString
        TextBox109.Text = Round(Ipy, 0).ToString
        TextBox98.Text = Round(Asx, 0).ToString
        TextBox110.Text = Round(Asy, 0).ToString
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, TabPage4.Enter
        Calc_sequence()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged
        Calc_sequence()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown18.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, RadioButton9.CheckedChanged, RadioButton11.CheckedChanged, RadioButton10.CheckedChanged
        Calc_sequence()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Calc_sequence()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Calc_sequence()
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        Write_to_word()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click, TabPage3.Enter
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
        Calc_chaper5_5_1()
        Calc_chaper5_5_2()
        Calc_chaper5_7()
        Check_for_problems()
    End Sub

    'Write data to Word 
    'see https://msdn.microsoft.com/en-us/library/office/aa192495(v=office.11).aspx
    Private Sub Write_to_word()
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        ' Dim ufilename As String
        Dim row As Integer

        Try
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            'Insert a paragraph at the beginning of the document. 
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = 14
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 1                '1 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = CInt(False)
            oPara2.Range.Text = "ABS Buckling and Utlimate Strength of stiffened panels" & vbCrLf
            oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

            row = 1
            oTable.Cell(row, 1).Range.Text = "Project Name"
            oTable.Cell(row, 2).Range.Text = TextBox79.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Item number"
            oTable.Cell(row, 2).Range.Text = TextBox80.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Filter type"
            oTable.Cell(row, 2).Range.Text = TextBox81.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author"
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date"
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            row += 1
            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '------------------ Panel data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Dimensions main plate (3/1.1)"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Length of long plate edge (And Or stiffener length)"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown1.Value, 1).ToString
            oTable.Cell(row, 3).Range.Text = "[cm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Length of short plate edge"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown2.Value, 1).ToString
            oTable.Cell(row, 3).Range.Text = "[cm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Thickness of plating"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown3.Value, 1).ToString
            oTable.Cell(row, 3).Range.Text = "[cm]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(4.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.1)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Material data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Material properties"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Specified Min yield point of plate"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown14.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[N/cm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Modulus of elasticity"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown13.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[N/cm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Poisson 's rate"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown12.Value, 2).ToString
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(4.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.1)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '------------------ Loads data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Applied loads(3 / 1.3) Figure 5"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Uniform lateral pressure"
            'oTable.Cell(row, 2).Range.Text = ComboBox2.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Compression stress in longitudinal direction, σax"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown11.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[N/cm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Compression stress in transverse direction, σay"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown9.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[N/cm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Compression stress in transverse direction, σbx"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown7.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[N/cm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "In-plane bending stress in transverse direction, σby"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown10.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[N/cm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Edge shear stress	[N/cm2] τ"
            oTable.Cell(row, 2).Range.Text = TextBox20.Text
            oTable.Cell(row, 3).Range.Text = "[N/cm2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(4.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.1)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Results data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Results"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Plate Ultimate Strength under combined in-plane stresses (<=1)"
            oTable.Cell(row, 2).Range.Text = TextBox29.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Uniform Lateral Pressure /Ultimate strength (<=1)"
            oTable.Cell(row, 2).Range.Text = TextBox83.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Beam colums buckling state limit stiffeners FIXED (<=1)"
            oTable.Cell(row, 2).Range.Text = TextBox52.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Flexural-torsional buckling state limit criterium (<=1)"
            oTable.Cell(row, 2).Range.Text = TextBox71.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(4.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.1)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        Catch ex As Exception
            'MessageBox.Show(ex.Message & "Problem storing file to" & dirpath_Rap)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Check_for_problems()

        '-------- find all numeric, combobox, checkbox and radiobutton controls -----------------
        FindControlRecursive(all_textbox, Me, GetType(TextBox))   'Find the control
        all_textbox = all_textbox.OrderBy(Function(x) x.Name).ToList()      'Alphabetical order
        For i = 0 To all_textbox.Count - 1
            Dim grbx As TextBox = CType(all_textbox(i), TextBox)
            If grbx.BackColor = Color.Coral Then
                Button9.Visible = CBool(vbTrue)
                Exit For
            Else
                Button9.Visible = CBool(vbFalse)
            End If
        Next

    End Sub

    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

End Class
