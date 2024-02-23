Option Explicit
'==============================================================================
' Timer Class Module
'==============================================================================
' Description: Simple tool to detect areas in your own code that are slow
' Author: Matthew Snow
' Version: 1.0
' Last Modified: 15 Feb 2024
'==============================================================================

'==============================================================================
' Types
'==============================================================================
' Define a custom Type to store timer information
Private Type TimerInfo
    title As String
    Time As Double
End Type

' Array to store multiple timers
Private Timers() As TimerInfo

'==============================================================================
' Public Subs
'==============================================================================
Public Sub Start()
    ' Start method sets a beginning point for the timer
    ReDim Timers(1 To 1)
    Timers(1).Time = Timer
    Timers(1).title = "Total Completion Time"
End Sub

Public Sub StoreTime(title As String)
    ' StoreTime method records the time and title for a specific point in the code
    Dim timerIndex As Integer
    timerIndex = UBound(Timers) + 1
    ReDim Preserve Timers(1 To timerIndex)
    Timers(timerIndex).Time = Timer
    Timers(timerIndex).title = title
End Sub

Public Sub Report()
    ' Report method prints the time differences between each timer and the start time
    ' by default this outputs to the console
    Dim i As Integer
    
    For i = 2 To UBound(Timers)
        Debug.Print Timers(i).title & ": ", Format((Timers(i).Time - Timers(1).Time) * 1000, "0.000") & " ms"
    Next i
    
    Debug.Print Timers(1).title & ": ", Format((Timer - Timers(1).Time) * 1000, "0.000") & " ms"
End Sub

