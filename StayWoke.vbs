Sub StayWoke()

Dim Switch As String, i As Integer
Switch = ActiveWorkbook.Worksheets(1).Range("A2").Value

If Switch = "On" Then
    ActiveWorkbook.Worksheets(1).Range("A2").Value = "Off"
    Exit Sub
Else
    ActiveWorkbook.Worksheets(1).Range("A2").Value = "On"
    
Do While ActiveWorkbook.Worksheets(1).Range("A2").Value = "On"
        
 i = 0
        
 If ActiveWorkbook.Worksheets(1).Range("A2").Value = "Off" Then
    ActiveWorkbook.Worksheets(1).Range("A3").Value = ""
    End
 End If
        
     Do While i < 15
        Pause (1)
        
        If ActiveWorkbook.Worksheets(1).Range("A2").Value = "Off" Then
            ActiveWorkbook.Worksheets(1).Range("A3").Value = ""
            End
        End If
        
            If ActiveWorkbook.Worksheets(1).Range("A3").Value = "Woke" Then
               ActiveWorkbook.Worksheets(1).Range("A3").Value = ""
            Else
               ActiveWorkbook.Worksheets(1).Range("A3").Value = "Woke"
            End If
            
            i = i + 1
            'ActiveWorkbook.Worksheets(1).Range("A4").Value = i
     Loop
     
 Application.SendKeys ("+{F12}")
 ActiveWorkbook.Worksheets(1).Range("A5").Value = "Sent"
    
Loop
    
End If

End Sub
Public Function Pause(NumberOfSeconds As Variant)
    On Error GoTo Error_GoTo

    Dim PauseTime As Variant
    Dim Start As Variant
    Dim Elapsed As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Elapsed = 0
    Do While Timer < Start + PauseTime
        Elapsed = Elapsed + 1
        If Timer = 0 Then
            ' Crossing midnight
            PauseTime = PauseTime - Elapsed
            Start = 0
            Elapsed = 0
        End If
        DoEvents
    Loop

Exit_GoTo:
    On Error GoTo 0
    Exit Function
Error_GoTo:
    Debug.Print Err.Number, Err.Description, Erl
    GoTo Exit_GoTo
End Function
