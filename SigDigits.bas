Attribute VB_Name = "SigDigits_Module"
Option Explicit

'Returns "Number" rounded to "NumberOfSignificantDigits"
Public Function SigDigits(Number, NumberOfSignificantDigits As Integer) As Double
Dim NegativeFlag As Integer, TensPower As Integer, NumA As Double
Dim NumB As Double, SD  As Double, Xtmp As Double
Dim n As Integer

    n = NumberOfSignificantDigits 'Integer
    'Here, Number must not be 0 [zero]:
    If Number <> 0 Then
        'Check for sign of Number:
        Select Case Number
           Case Is > 0
            NegativeFlag = 0
            Xtmp = Number
           Case Is < 0
            NegativeFlag = -1
            Xtmp = Number * -1
        End Select
        TensPower = -Int(Log(Xtmp) / Log(10)) - 1
        'NumA = .########
        NumA = Xtmp * 10 ^ TensPower
        NumB = Round(NumA, n)
        SD = NumB / 10 ^ TensPower
        'Correct for sign if necessary:
        If NegativeFlag Then SD = -SD
    Else  'Number = 0
        SD = 0
    End If
    'Return the rounded value:
    SigDigits = SD
End Function
