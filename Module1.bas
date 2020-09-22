Attribute VB_Name = "Module1"
Sub Pause(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub
