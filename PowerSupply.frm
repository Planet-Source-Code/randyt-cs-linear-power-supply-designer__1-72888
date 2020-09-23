VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Linear Power Supply Designer"
   ClientHeight    =   5790
   ClientLeft      =   720
   ClientTop       =   1680
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "PowerSupply.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5790
   ScaleMode       =   0  'User
   ScaleWidth      =   8295
   Begin VB.PictureBox DataBox 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   8025
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txtHertz 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Text            =   "txtHertz"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtAmpsFL 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Text            =   "txtAmpsFL"
         Top             =   1665
         Width           =   1215
      End
      Begin VB.TextBox txtVoltsReg 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Text            =   "txtVoltsReg"
         Top             =   1290
         Width           =   1215
      End
      Begin VB.TextBox txtUF 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Text            =   "txtUF"
         Top             =   915
         Width           =   1215
      End
      Begin VB.TextBox txtVoltsCoil 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Text            =   "txtVoltsCoil"
         Top             =   540
         Width           =   1215
      End
      Begin VB.ComboBox DiodeDrp 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   540
         Width           =   1155
      End
      Begin VB.CommandButton UpdateGraph 
         Appearance      =   0  'Flat
         Caption         =   "UPDATE"
         Height          =   315
         Left            =   3240
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hertz"
         Height          =   195
         Left            =   1440
         TabIndex        =   26
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amps full-load each regulator"
         Height          =   195
         Left            =   1440
         TabIndex        =   25
         Top             =   1713
         Width           =   2490
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volts +/- each regulator"
         Height          =   195
         Left            =   1440
         TabIndex        =   24
         Top             =   1342
         Width           =   2040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MicroFarads each C1, C2"
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   975
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volts RMS each secondary coil. *"
         Height          =   195
         Left            =   1440
         TabIndex        =   22
         Top             =   600
         Width           =   2865
      End
      Begin VB.Label AmpsDiodeWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Amps diode surge"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   10
         Top             =   1710
         Width           =   1935
      End
      Begin VB.Label AmpsDiodeLbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label WattsDiodeWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Watts each diode"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   1350
         Width           =   1935
      End
      Begin VB.Label WattsDiodeLbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   6
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label WattsRegWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Watts each regulator"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         TabIndex        =   8
         Top             =   990
         Width           =   1935
      End
      Begin VB.Label WattsRegLbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label DiodedrpLbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Diode drop Si-.7 Ge-.3"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5880
         TabIndex        =   19
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label GraphBoxMsjLbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4440
         TabIndex        =   21
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label XfmrMsjLbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.PictureBox DualFullBmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      Picture         =   "PowerSupply.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox SingFul2Bmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      Picture         =   "PowerSupply.frx":16D4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox SingFullBmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      Picture         =   "PowerSupply.frx":2A9E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox DualHalfBmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      Picture         =   "PowerSupply.frx":3E68
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox FwaveBmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      Picture         =   "PowerSupply.frx":5272
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   780
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox SingHalfBmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      Picture         =   "PowerSupply.frx":6D98
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox DisplayBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox HwaveBmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      Picture         =   "PowerSupply.frx":81C2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox GraphBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   3840
      ScaleHeight     =   2865
      ScaleWidth      =   4305
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.ComboBox PSType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim formActive As Boolean

Private Sub BoxFill(X1, Y1, H, W, culr&)
    Dim delta As Integer
    
    'draw a black boarder box from (X1,Y1) to (X1+W,Y2+H)
    'GraphBox.Line (X1, Y1)-(X1 + W, Y1 + H), RGB(0, 0, 0), B
    
    delta = 40
    'fill a colored box from (X1+delta,Y1+delta) to (X+W-delta,Y1+H-delta)
    GraphBox.Line (X1 + delta, Y1 + delta)-(X1 + W - delta, Y1 + H - delta), culr&, BF

End Sub

Private Sub DBox(X, Y, A$, C)
    DisplayBox.CurrentX = X
    DisplayBox.CurrentY = Y
    DisplayBox.ForeColor = C
    DisplayBox.Print A$
    DisplayBox.ForeColor = 0
End Sub

Private Function DrawCapForm() As Boolean
    Dim amps As Double, farads As Double, f As Double, W As Double, stp As Double
    Dim Xa1 As Double, Xa2 As Double, Xa3 As Double, Xa4 As Double
    Dim Ya1 As Double, Ya2 As Double, Ya3 As Double
    Dim Xalg As Double, Yalg As Double
    Dim Vt3 As Double, Vt4 As Double, Ya4 As Double
    Dim Xa3a As Double, Xa3b As Double
    Dim t3b As Double, Vt3b As Double
    
    If Not tn(t1, t2, t3, t4) Then  'gets the real-time values on the waveform
        'error
        DrawCapForm = False
        Exit Function
    End If
    
    amps = DataArray(ActiveType, 4)
    farads = DataArray(ActiveType, 2) * 10 ^ -6
    f = DataArray(ActiveType, 5)
    W = 2 * Pi * f
    stp = GraphWidth / 1000
    
    'check for Vcap < Vreg at t2
    If Vt(t2) < DataArray(ActiveType, 3) Then
      VcapBelowVreg = -1
    Else
      VcapBelowVreg = 0
    End If
    
    'calculate the algebraic equivalents of t1-t4
    Xa1 = f * GraphWidth * t1 / 1.5: Ya1 = Ya(Xa1)
    Xa2 = f * GraphWidth * t2 / 1.5: Ya2 = Ya(Xa2)
    Xa3 = f * GraphWidth * t3 / 1.5: Ya3 = Ya1
    Xa4 = GraphWidth
    
    GraphBox.DrawWidth = 2
    'plot the function from 0 to Xa1
    For Xalg = 0 To Xa1 Step stp
      Yalg = Ya(Xalg)
      GraphBox.PSet (Xalg + GraphLeft, GraphTop + GraphHeight - Yalg), vbRed
    Next Xalg
    
    'draw the line from Xa1 to Xa2
    GraphBox.Line (Xa1 + GraphLeft, GraphTop + GraphHeight - Ya1)-(Xa2 + GraphLeft, GraphTop + GraphHeight - Ya2), vbRed
    
    'plot the function from Xa2 to Xa3
    For Xalg = Xa2 To Xa3 Step stp
      Yalg = Ya(Xalg)
      GraphBox.PSet (Xalg + GraphLeft, GraphTop + GraphHeight - Yalg), vbRed
    Next Xalg
    
    'plot the rest of the graph depending on full/half wavetype
    Select Case ActiveType
     Case Is = 1, 2  'halfwaves
       'draw the line from Xa3 to Xa4
        Vt3 = Vt(t3)
        Vt4 = Vt3 - (amps / farads) * (1.5 / f - t3)
        Ya4 = Vt4 * GraphHeight / Vp     'Peak Voltage
        GraphBox.Line (Xa3 + GraphLeft, GraphTop + GraphHeight - Ya3)-(Xa4 + GraphLeft, GraphTop + GraphHeight - Ya4), vbRed
     Case Else       'fullwaves
       'draw the line from Xa3 to Xa3a
        Xa3a = Xa2 + GraphWidth / 3
        GraphBox.Line (Xa3 + GraphLeft, GraphTop + GraphHeight - Ya3)-(Xa3a + GraphLeft, GraphTop + GraphHeight - Ya2), vbRed
       
       'plot the function from Xa3a to Xa3b
        Xa3b = Xa3 + GraphWidth / 3
        For Xalg = Xa3a To Xa3b Step stp
          Yalg = Ya(Xalg)
          GraphBox.PSet (Xalg + GraphLeft, GraphTop + GraphHeight - Yalg), vbRed
        Next Xalg
    
        'draw the line from Xa3b to Xa4
        t3b = t3 + 0.5 / f
        Vt3b = Vt(t3b)
        Vt4 = Vt3b - (amps / farads) * (1.5 / f - t3b)
        Ya4 = Vt4 * GraphHeight / Vp    'Peak Voltage
        GraphBox.Line (Xa3b + GraphLeft, GraphTop + GraphHeight - Ya3)-(Xa4 + GraphLeft, GraphTop + GraphHeight - Ya4), vbRed
    
    End Select

    DrawCapForm = True
End Function

Private Sub DrawRegVolt()
    Dim f As Double, W As Double, stp As Double
    Dim Xa1 As Double, Ya1 As Double, Xa2 As Double, Ya2 As Double
    Dim Xalg As Double, Yalg As Double
    
    f = DataArray(ActiveType, 5)
    W = 2 * Pi * f
    stp = GraphWidth / 1000
    
    'check for Vreg >= Vpeak
    If DataArray(ActiveType, 3) >= Vp Then   'We have an error condition
      Xa1 = 0
      Ya1 = GraphHeight + 30                 'draw Vreg visibly above Vpeak
      Xa2 = GraphWidth
      Ya2 = Ya1
    Else
      'calculate the algebraic Xa1 of VReg.
      Xa1 = f * GraphWidth * ArcSin(DataArray(ActiveType, 3) / Vp) / (1.5 * W)
      Xa2 = GraphWidth
    
      'plot the function from 0 to Xa1
      For Xalg = 0 To Xa1 Step stp
        Yalg = Ya(Xalg)
        GraphBox.PSet (Xalg + GraphLeft, GraphTop + GraphHeight - Yalg), vbBlue
      Next Xalg
    
      Ya1 = Ya(Xa1)
      Ya2 = Ya1
    End If
    
    'draw the line from Xa1 to Xa2
    GraphBox.Line (Xa1 + GraphLeft, GraphTop + GraphHeight - Ya1)-(Xa2 + GraphLeft, GraphTop + GraphHeight - Ya2), vbBlue
    
End Sub

Private Sub ErrorCondition(msj$)
    
    GraphBoxMsjLbl.ForeColor = &HFF&      'red
    GraphBoxMsjLbl.Caption = msj$
    WattsRegLbl.Caption = ""      'clear the label
    WattsDiodeLbl.Caption = ""
    AmpsDiodeLbl.Caption = ""
    WattsRegLbl.BackColor = &HFFFF&      'yellow
    WattsDiodeLbl.BackColor = &HFFFF&
    AmpsDiodeLbl.BackColor = &HFFFF&
    
End Sub


Private Sub Form_Load()
    
    PSType.AddItem "Halfwave Single polarity"
    PSType.AddItem "Halfwave Dual   polarity"
    PSType.AddItem "Fullwave Single polarity 1"
    PSType.AddItem "Fullwave Single polarity 2"
    PSType.AddItem "Fullwave Dual   polarity"
    
    DiodeDrp.AddItem ".7"
    DiodeDrp.AddItem ".3"
    
    Pi = 4 * Atn(1)
    
    InnitalizeData
    
    'GraphBox parameters
    LeftSpace = 100: RightSpace = 125
    TopSpace = 450: BottomSpace = 150
    GraphLeft = LeftSpace
    GraphTop = TopSpace
    GraphWidth = GraphBox.Width - LeftSpace - RightSpace
    GraphHeight = GraphBox.Height - TopSpace - BottomSpace
    
    DisplayBox.AutoRedraw = -1
    GraphBox.AutoRedraw = -1
    
    DiodeDrp.ListIndex = 0
    PSType.ListIndex = 0

End Sub

Private Sub Form_Activate()
    Static oneShot As Boolean
    If oneShot = False Then
        UpdateGraph_Click
        formActive = True
        oneShot = True
    End If
End Sub

Private Sub GBox(X!, Y!, A$, C&)
    GraphBox.CurrentX = X
    GraphBox.CurrentY = Y
    GraphBox.ForeColor = C
    GraphBox.Print A$
    GraphBox.ForeColor = 0
End Sub

Private Sub GBoxLegend()
    BoxFill 330, 2040, 125, 125, vbRed    'Vcap legend
    GBox 510, 2010, "Vcap", vbBlack
    BoxFill 330, 2280, 125, 125, vbBlue     'Vreg legend
    GBox 510, 2250, "Vreg", vbBlack
End Sub

Private Sub GraphDataUpdate()
   
    Form1.MousePointer = 11 'hourglass
    
     CalculateVp
    
    Select Case ActiveType
      Case Is = 1  'singhalf
       GraphBox.Cls
       GraphBox.Picture = HwaveBmp.Picture
       GBox 1695, 75, "Waveform at A (+Volts)", vbRed
       GBox 45, 75, "Vp:" + Str$(SigDigits(Vp, 3)), vbBlack
       GBoxLegend
       If Not DrawCapForm() Then
        'error
        MsgBox "Terminating program", vbCritical, "Bad Input Data"
        End
       End If
       DrawRegVolt
       If ThereIsNoError() Then
         DoRegWatts
         DoDiodeWatts
         DoDoideSurgeAmps
       End If
    
      Case Is = 2  'dualhalf
       GraphBox.Cls
       GraphBox.Picture = HwaveBmp.Picture
       GBox 1295, 60, "Waveform at A(+Volts) or B(-Volts)", vbRed
       GBox 45, 75, "Vp:" + Str$(SigDigits(Vp, 3)), vbBlack
       GBoxLegend
       If Not DrawCapForm() Then
        'error
        MsgBox "Terminating program", vbCritical, "Bad Input Data"
        End
       End If
       DrawRegVolt
       If ThereIsNoError() Then
         DoRegWatts
         DoDiodeWatts
         DoDoideSurgeAmps
       End If
    
      Case Is = 3  'singfull1
       GraphBox.Cls
       GraphBox.Picture = FwaveBmp.Picture
       GBox 1695, 75, "Waveform at A (+Volts)", vbRed
       GBox 45, 75, "Vp:" + Str$(SigDigits(Vp, 3)), vbBlack
       GBoxLegend
       If Not DrawCapForm() Then
        'error
        MsgBox "Terminating program", vbCritical, "Bad Input Data"
        End
       End If
       DrawRegVolt
       If ThereIsNoError() Then
         DoRegWatts
         DoDiodeWatts
         DoDoideSurgeAmps
       End If
    
      Case Is = 4  'singfull2
       GraphBox.Cls
       GraphBox.Picture = FwaveBmp.Picture
       GBox 1695, 75, "Waveform at A (+Volts)", vbRed
       GBox 45, 75, "Vp:" + Str$(SigDigits(Vp, 3)), vbBlack
       GBoxLegend
       If Not DrawCapForm() Then
        'error
        MsgBox "Terminating program", vbCritical, "Bad Input Data"
        End
       End If
       DrawRegVolt
       If ThereIsNoError() Then
         DoRegWatts
         DoDiodeWatts
         DoDoideSurgeAmps
       End If
    
      Case Is = 5  'dualfull
       GraphBox.Cls
       GraphBox.Picture = FwaveBmp.Picture
       GBox 1295, 60, "Waveform at A(+Volts) or B(-Volts)", vbRed
       GBox 45, 75, "Vp:" + Str$(SigDigits(Vp, 3)), vbBlack
       GBoxLegend
       If Not DrawCapForm() Then
        'error
        MsgBox "Terminating program", vbCritical, "Bad Input Data"
        End
       End If
       DrawRegVolt
       If ThereIsNoError() Then
         DoRegWatts
         DoDiodeWatts
         DoDoideSurgeAmps
       End If
    
    End Select
 
    Form1.MousePointer = 0   'default
    
End Sub


Private Sub InnitalizeData()
Dim cnt As Integer
Dim A As Double, B As Double, C As Double, D As Double, E As Double

    For cnt = 1 To 5
      Select Case cnt
      '      secvolt,     uf,  regvolt,regfla,  hertz
        Case Is = 1
          A = 12.6: B = 4000: C = 12: D = 1: E = 60    'singhalf
        Case Is = 2
          A = 18:  B = 2000: C = 15: D = 1: E = 60   'dualhalf
        Case Is = 3
          A = 6.3:  B = 2200: C = 5: D = 1: E = 60   'singfull
        Case Is = 4
          A = 12.6: B = 2000: C = 12: D = 1: E = 60  'singful2
        Case Is = 5
          A = 18:  B = 1000: C = 15: D = 1: E = 60   'dualfull
      End Select
      DataArray(cnt, 1) = A
      DataArray(cnt, 2) = B
      DataArray(cnt, 3) = C
      DataArray(cnt, 4) = D
      DataArray(cnt, 5) = E
    
    Next cnt
  
End Sub

Private Sub PSType_Click()
   
    DisplayBox.Visible = -1
    GraphBox.Visible = -1
    DataBox.Visible = -1
    
    Select Case PSType.Text
     Case "Halfwave Single polarity"
      ActiveType = 1
      DisplayBox.Picture = SingHalfBmp.Picture
      DBox 1825, 215, "A", vbRed: DBox 1370, 1165, "C1", vbBlue
      XfmrMsjLbl.Caption = " *  FULLY-LOADED  secondary"
     Case "Halfwave Dual   polarity"
      ActiveType = 2
      DisplayBox.Picture = DualHalfBmp.Picture
      DBox 1985, 225, "A", vbRed: DBox 1985, 2085, "B", vbRed
      DBox 2370, 745, "C1", vbBlue: DBox 2355, 1595, "C2", vbBlue
      DBox 905, 1000, "CT", vbGreen
      XfmrMsjLbl.Caption = "*FULL-LOAD ONE secondary coil"
     Case "Fullwave Single polarity 1"
      ActiveType = 3
      DisplayBox.Picture = SingFullBmp.Picture
      DBox 1750, 235, "A", vbRed: DBox 2230, 1145, "C1", vbBlue
      DBox 925, 1005, "CT", vbGreen
      XfmrMsjLbl.Caption = "*FULL-LOAD ONE secondary coil"
     Case "Fullwave Single polarity 2"
      ActiveType = 4
      DisplayBox.Picture = SingFul2Bmp.Picture
      DBox 2160, 195, "A", vbRed: DBox 2480, 1140, "C1", vbBlue
      XfmrMsjLbl.Caption = " *  FULLY-LOADED  secondary"
     Case "Fullwave Dual   polarity"
      ActiveType = 5
      DisplayBox.Picture = DualFullBmp.Picture
      DBox 2175, 225, "A", vbRed: DBox 2175, 2055, "B", vbRed
      DBox 2430, 730, "C1", vbBlue: DBox 2430, 1565, "C2", vbBlue
      DBox 770, 1050, "CT", vbGreen
      XfmrMsjLbl.Caption = "*FULL-LOAD ONE secondary coil"
    End Select
    
    txtVoltsCoil.Text = Str$(DataArray(ActiveType, 1))
    txtUF.Text = Str$(DataArray(ActiveType, 2))
    txtVoltsReg.Text = Str$(DataArray(ActiveType, 3))
    txtAmpsFL.Text = Str$(DataArray(ActiveType, 4))
    txtHertz.Text = Str$(DataArray(ActiveType, 5))
    
    UpdateGraph_Click
    'SendKeys "{TAB}"   'Removes highlight from listbox
    'UpdateGraph.SetFocus
    
End Sub

Private Function ThereIsNoError()
    Dim AOK As Integer
    
    AOK = -1
    If VcapBelowVreg Then
       ErrorCondition ("           Vcap goes below Vreg")
       AOK = 0
    End If
    
    If DataArray(ActiveType, 3) >= Vp Then  'If Vpeak < regulator voltage
       ErrorCondition ("            Vreg is above Vpeak")
       AOK = 0
    End If
    
    'If there are not any errors then disable timer and reset colors etc.
    'ie. undo what ErrorCondition() does.
    If AOK Then
     GraphBoxMsjLbl.ForeColor = vbBlue       'blue
     GraphBoxMsjLbl.Caption = "   Should work with indicated values"
     WattsRegLbl.BackColor = &HFFFF80          'light blue
     WattsDiodeLbl.BackColor = &HFFFF80
     AmpsDiodeLbl.BackColor = &HFFFF80
    End If
    
    ThereIsNoError = AOK    'return error status

End Function


Private Function tn(t1, t2, t3, t4) As Boolean
    Dim amps As Double, farads As Double, f As Double, W As Double
    Dim Vt1 As Double, hi As Double, lo As Double
    Dim cnt As Integer, n As Double
    Dim v_a As Double, v_b As Double
    Dim arcCosArg As Single
    
    't1 - where first linear-discharge-line begins
    't2 - where first linear-discharge-line hits next [-sine or sine] hump
    't3 - where second linear-discharge-line begins
    't4 - end of graph = 1.5/f ... one and a half periods.
    
    't1 calculation
    'slope = volts/sec = -amps/farads , for linear discharge line
    'slope = w Vp cos(w t)            , for Vp sin(w t)
    'w Vp cos(w t) = -amps/farads     , simultaneous solution ... for "t"
    't = arccos(-amps/(farads w Vp) )/ w
    'ArcCos = -Atn(x! / Sqr(-x! * x! + 1)) + Pi / 2
    ' range= [0,Pi]  domain= [-1,1]
    
    amps = DataArray(ActiveType, 4)
    farads = DataArray(ActiveType, 2) * 10 ^ -6
    f = DataArray(ActiveType, 5)
    W = 2 * Pi * f
    
    arcCosArg = (-amps / (farads * W * Vp))
    If arcCosArg < -1 Or arcCosArg > 1 Then
        'error condition:
        tn = False
        Exit Function
    Else
        t1 = ArcCos(arcCosArg) / W
    End If
    
    t4 = 1.5 / f
    
    Vt1 = Vt(t1)
    
    Select Case ActiveType  'set t3s, and high side intersect test values
     Case Is = 1, 2  'halfwaves
       t3 = t1 + 1 / f
       hi = 1.25 / f
     Case Else       'fullwaves
       t3 = t1 + 0.5 / f
       hi = 0.75 / f
    End Select
       
    lo = 0.25 / f
    
    For cnt = 1 To 40  'close-in walls, from both sides, finds t2 ... fast!
     n = (lo + hi) / 2
     v_a = Vt1 - (amps / farads) * (n - t1)  'linear-discharge-line volts @ n
     v_b = Vt(n)                             'Vt - sine hump @ n
     If v_a > v_b Then
       lo = n
     Else
       hi = n
     End If
    Next cnt
    
    t2 = n    'note- this is a t2 approximation, but it is very accurate.
    
    tn = True
    
End Function

Private Sub txtAmpsFL_Change()
    txtAmpsFL.BackColor = vbYellow
End Sub

Private Sub txtHertz_Change()
    txtHertz.BackColor = vbYellow
End Sub

Private Sub txtUF_Change()
    txtUF.BackColor = vbYellow
End Sub

Private Sub txtVoltsCoil_Change()
    txtVoltsCoil.BackColor = vbYellow
End Sub

Private Sub txtVoltsReg_Change()
    txtVoltsReg.BackColor = vbYellow
End Sub

Private Sub UpdateGraph_Click()
    
    If Val(txtVoltsCoil.Text) > 0 Then
        DataArray(ActiveType, 1) = Val(txtVoltsCoil.Text)
        txtVoltsCoil.BackColor = vbWindowBackground
    Else
        txtVoltsCoil.Text = DataArray(ActiveType, 1)
        Beep
    End If
    
    If Val(txtUF.Text) > 0 Then
        txtUF.BackColor = vbWindowBackground
        DataArray(ActiveType, 2) = Val(txtUF.Text)
    Else
        txtUF.Text = DataArray(ActiveType, 2)
        Beep
    End If
    
    If Val(txtVoltsReg.Text) > 0 Then
        txtVoltsReg.BackColor = vbWindowBackground
        DataArray(ActiveType, 3) = Val(txtVoltsReg.Text)
    Else
        txtVoltsReg.Text = DataArray(ActiveType, 3)
        Beep
    End If
    
    If Val(txtAmpsFL.Text) > 0 Then
        txtAmpsFL.BackColor = vbWindowBackground
        DataArray(ActiveType, 4) = Val(txtAmpsFL.Text)
    Else
        txtAmpsFL.Text = DataArray(ActiveType, 4)
        Beep
    End If
    
    If Val(txtHertz.Text) > 0 Then
        txtHertz.BackColor = vbWindowBackground
        DataArray(ActiveType, 5) = Val(txtHertz.Text)
    Else
        txtHertz.Text = DataArray(ActiveType, 5)
        Beep
    End If
    
    GraphDataUpdate
    
    If formActive Then DataBox.SetFocus

End Sub


