Attribute VB_Name = "MSpeed"

' Copyright https://github.com/ftrippel

Option Explicit

Private isSpeed As Boolean

Public Sub Speed(b As Boolean)
    If b And Not isSpeed Then
        isSpeed = True
        Application.screenUpdating = False
        Application.calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
    ElseIf Not b And isSpeed Then
        isSpeed = False
        Application.screenUpdating = True
        Application.calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
    End If
End Sub

