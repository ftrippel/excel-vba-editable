Attribute VB_Name = "MSpeed"
' Copyright https://github.com/ftrippel

Private isSpeed As Boolean
Private speedCalculation As XlCalculation
Private speedScreenUpdating As Boolean
Private speedEnableEvents As Boolean
Private speedDisplayStatusBar As Boolean

Public Sub Speed(b As Boolean)
    If b And Not isSpeed Then
        isSpeed = True
        speedScreenUpdating = Application.screenUpdating
        speedCalculation = Application.calculation
        speedEnableEvents = Application.EnableEvents
        speedDisplayStatusBar = Application.DisplayStatusBar
        
        Application.screenUpdating = False
        Application.calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
    ElseIf Not b And isSpeed Then
        isSpeed = False
        Application.screenUpdating = speedScreenUpdating
        Application.calculation = speedCalculation
        Application.EnableEvents = speedEnableEvents
        Application.DisplayStatusBar = speedDisplayStatusBar
    End If
End Sub

