Attribute VB_Name = "Module1"
Option Explicit

Public Pi As Single
Public Vp As Single
Public ActiveType As Integer
Public VcapBelowVreg As Integer
Public t1, t2, t3, t4 As Single
Public DataArray(1 To 5, 1 To 5) As Single
Public LeftSpace, RightSpace, TopSpace, BottomSpace As Integer
Public GraphLeft, GraphTop, GraphWidth, GraphHeight As Integer

Function ArcCos(X!)

'Domain ArcCos: [-1, 1]
ArcCos = -Atn(X! / Sqr(-X! * X! + 1)) + Pi / 2

End Function

Function ArcSin(X!) As Single
ArcSin = Atn(X / Sqr(-X * X + 1))
End Function

Sub CalculateVp()

Select Case ActiveType
 Case Is = 1, 2, 3, 5    'one diode drop in voltage
   Vp = DataArray(ActiveType, 1) * Sqr(2) - Val(Form1.DiodeDrp.Text)
 Case Is = 4             'two diode drops in voltage   sing full 2
   Vp = DataArray(ActiveType, 1) * Sqr(2) - 2 * Val(Form1.DiodeDrp.Text)
End Select

End Sub

Sub DoDiodeWatts()
    Dim Edd As Double, RatioVt As Double, farads As Double
    Dim f As Double, W As Double, Watts As Double
    
    'Watts = Iavg over t1->t3  * Edd
    
    Edd = Val(Form1.DiodeDrp.Text)
    RatioVt = (t3 - t2) / (t3 - t1)
    farads = DataArray(ActiveType, 2) * 10 ^ -6
    f = DataArray(ActiveType, 5)
    W = 2 * Pi * f
    
    Select Case ActiveType
     Case Is = 1, 2   'halfwaves
      Watts = (farads * Vp * (Sin(W * t3) - Sin(W * t2)) / (t3 - t2) * Edd) * RatioVt
     Case Else    'fullwaves  **each diode conducts half the time**
      Watts = (-farads * Vp * (Sin(W * t3) - Sin(W * t2)) / (t3 - t2) * Edd) * RatioVt / 2
    End Select
    
    Form1.WattsDiodeLbl.Caption = Str$(SigDigits(Watts, 3))

End Sub

Sub DoDoideSurgeAmps()
    Dim farads As Double, f As Double, W As Double
    Dim It As Double
    
    farads = DataArray(ActiveType, 2) * 10 ^ -6
    f = DataArray(ActiveType, 5)
    W = 2 * Pi * f
    
    Select Case ActiveType
     Case Is = 1, 2   'halfwaves
       It = farads * W * Vp * Cos(W * t2)
     Case Else        'fullwaves
       It = -farads * W * Vp * Cos(W * t2)
    End Select
    
    Form1.AmpsDiodeLbl.Caption = Str$(SigDigits(It, 3))
    
End Sub

Sub DoRegWatts()
Dim Ireg As Double, Vreg As Double, RatioSVt As Double, RatioLin As Double
Dim Watts As Double

'Watts = Ireg * [avg Vt over t1->t3  - Vreg]

Ireg = DataArray(ActiveType, 4)
Vreg = DataArray(ActiveType, 3)
RatioSVt = (t3 - t2) / (t3 - t1)
RatioLin = (t2 - t1) / (t3 - t1)

Select Case ActiveType
 Case Is = 1, 2   'halfwaves
  Watts = Ireg * (SVt(t2, t3) / (t3 - t2) * RatioSVt + (Vt(t1) + Vt(t2)) / 2 * RatioLin - Vreg)
 Case Else        'fullwaves **we take negative of definate integral**
  Watts = Ireg * (-SVt(t2, t3) / (t3 - t2) * RatioSVt + (Vt(t1) + Vt(t2)) / 2 * RatioLin - Vreg)
End Select
  

Form1.WattsRegLbl.Caption = Str$(SigDigits(Watts, 3))

End Sub

Function SVt(ta, tb) As Single
    Dim f As Double, W As Double
    
    'This function returns the definite integral of the function
    '     Vt(t)=Vp sin(wt) over the interval [ta ,tb]
    'Returns SVt(ta, tb)= -(Vp/w)*( cos(w tb) - cos(w ta) ) ...
    '*******  WITHOUT  ******* fullwave/halfwave modifications
    'Where the units of the [ta, tb] are understood to be "seconds"
    'Domain t <> [ 0 , 1.5/f ]  one and a half periods
    'Global variables used- Pi ....................................... 3.1415.....
    '                       DataArray( ActiveType, 5 ) ...............f(frequency)
    '                       Vp ..................................Vp(Peak Voltage)
    
    f = DataArray(ActiveType, 5)             'frequency
    W = 2 * Pi * f                           'angular frequency
    
    If ta < 0 Then
        MsgBox "Parameter (t) of function SVt(t) is < 0. " + Chr$(13) + Chr$(10) + "t <> [0, 1.5/frequency]", 0, "Function Call Error"
        End
    End If
    If tb > (1.5 / f) Then
        MsgBox "Parameter (t) of function SVt(t) is > 1.5/f. " + Chr$(13) + Chr$(10) + "t <> [0, 1.5/frequency]", 0, "Function Call Error"
        End
    End If
    't is in range, so do it
    SVt = (-Vp / W) * (Cos(W * tb) - Cos(W * ta))

End Function

Function Vt(t) As Single
Dim f As Double, W As Double, V As Double

'Returns Vt(t)= Vp sin(wt) with fullwave/halfwave modifications
'Where the units of "t" are understood to be "seconds"
'Domain t <> [ 0 , 1.5/f ]  one and a half periods
'Global variables used- ActiveType ... which one of the five pwr. supply types
'                       Pi ....................................... 3.1415.....
'                       DataArray( ActiveType, 5 ) ...............f(frequency)
'                       Vp ..................................Vp(Peak Voltage)

f = DataArray(ActiveType, 5)             'frequency
W = 2 * Pi * f                           'angular frequency

Select Case t
 Case Is < 0
  MsgBox "Parameter (t) of function Vt(t) is < 0. " + Chr$(13) + Chr$(10) + "t <> [0, 1.5/frequency]", 0, "Function Call Error"
  End
 Case Is > (1.5 / f)
  MsgBox "Parameter (t) of function Vt(t) is > 1.5/f. " + Chr$(13) + Chr$(10) + "t <> [0, 1.5/frequency]", 0, "Function Call Error"
  End
 Case Else  't is in range, so do it
  V = Vp * Sin(W * t)
  If V < 0 Then    'sine is returning negative values
    Select Case ActiveType
     Case 1, 2     'halfwaves
       V = 0       'make it = 0  on halfwaves
     Case Else     'fullwaves
       V = -V      'flip it upside down on fullwaves
    End Select
  End If
  Vt = V
End Select

End Function

Function Ya(Xa) As Single
    Dim GH As Integer, B As Double, Y As Double
    
    'Returns Ya(Xa)= GH sin(b Xa) with fullwave/halfwave modifications
    'Where the units of "Xa" are understood to be "twips"
    'GH = GraphHeight, GW = GraphWidth, b = 3*Pi/GW
    'Domain Xa <> [ 0 , GW ]  one and a half periods, due to b
    'Global variables used- ActiveType ... which one of the five pwr. supply types
    '                       Pi ....................................... 3.1415.....
    '                       GraphHeight ................. Height of graph in twips
    '                       GraphWidth ................... Width of graph in twips
    
    GH = GraphHeight
    B = 3 * Pi / GraphWidth
    
    Select Case Xa
     Case Is < 0
      MsgBox "Parm. (Xa) of function Ya(Xa) is < 0. " + Chr$(13) + Chr$(10) + "Xa <> [0, GraphWidth]", 0, "Function Call Error"
      End
     Case Is > GraphWidth
      MsgBox "Parm. (Xa) of function Ya(Xa) is > GraphWidth." + Chr$(13) + Chr$(10) + "Xa <> [0, GraphWidth]", 0, "Function Call Error"
      End
     Case Else  'Xa is in range, so do it
      Y = GH * Sin(B * Xa)
      If Y < 0 Then    'sine is returning negative values
        Select Case ActiveType
         Case 1, 2     'halfwaves
           Y = 0       'make it = 0  on halfwaves
         Case Else     'fullwaves
           Y = -Y      'flip it upside down on fullwaves
        End Select
      End If
      Ya = Y
    End Select
    
End Function

