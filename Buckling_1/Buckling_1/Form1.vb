Imports System.IO
Imports System.Text
Imports System.Math

Public Class Form1
    Public _Ψ As Double         'Adjustment factor (psi)
    Public _σxmax As Double     'maximum compressive stress in longitudinal direction
    Public _σymax As Double     'maximum compressive stress in transverse direction
    Public _σCx As Double     'maximum compressive stress in longitudinal direction
    Public _σCy As Double     'maximum compressive stress in transverse direction
    Public _τ As Double         'Edge shear stress
    Public _τC As Double         'Edge shear stress
    Public _η As Double         'max length utilizing factor

    Public _t As Double         'Plate thickness [mm]
    Public _s As Double         'Stiffeners distance [mm]
    Public _l As Double         'stiffeners length [mm]
    Public _LG As Double        'girder length [mm]



    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter
        'Do somthing
    End Sub

End Class
