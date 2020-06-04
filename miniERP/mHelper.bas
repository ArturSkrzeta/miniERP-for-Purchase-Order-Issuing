Attribute VB_Name = "mHelper"
Public Sub TurnOnExcleFeatures()
    Excel.Application.EnableEvents = True
    Excel.Application.ScreenUpdating = True
    Excel.Application.DisplayAlerts = True
End Sub

Public Sub TurnOFFExcleFeatures()
    Excel.Application.DisplayAlerts = False
    Excel.Application.ScreenUpdating = False
    Excel.Application.EnableEvents = False
End Sub

