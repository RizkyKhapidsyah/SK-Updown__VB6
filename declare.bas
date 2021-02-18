Attribute VB_Name = "Module1"
Dim start_sec
Public rs As Recordset
Public con As Connection
Public Sub Delay(d As Integer)
start_sec = Second(Time)
If start_sec + d > 59 Then
d = d - 59
End If
While Second(Time) - start_sec <> d
Wend
End Sub

