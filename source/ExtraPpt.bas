Attribute VB_Name = "ExtraPpt"
Option Explicit

'this function helps to position shapes in Powerpoint
Public Function cm2points(pValueCm As Double) As Long

    #If Mac Then
        cm2points = CLng(pValueCm * (72 / 2.54))
    #Else
        cm2points = CLng(pValueCm * (96 / 2.54))
    #End If
End Function
