Attribute VB_Name = "modFormat"
Option Explicit

Sub RemoveBackgroundFillColor()
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub ApplyNumberFormat(TargetRange As Range, numberFormat As String)
    TargetRange.numberFormat = numberFormat
End Sub

Sub ApplyFontColorRGB(TargetRange As Range, red As Integer, green As Integer, blue As Integer)
    TargetRange.Font.Color = RGB(red, green, blue)
End Sub

Sub ApplyBackgroundColorRGB(TargetRange As Range, red As Integer, green As Integer, blue As Integer, Optional TintAndShade As Double = 0)
    TargetRange.Interior.Color = RGB(red, green, blue)
    TargetRange.Interior.TintAndShade = TintAndShade
End Sub
