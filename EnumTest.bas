Attribute VB_Name = "EnumTest"
'@Folder("Module")
'@ModuleDescription("Enumerator Test.")

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

Public Sub TestForEachEarlyBinding()

    Dim EnumTest As EnumTestEarlyClass
    Set EnumTest = New EnumTestEarlyClass

    EnumTest.count = 25

    Dim v As Variant
    For Each v In EnumTest
        Debug.Print v
    Next

End Sub


Private Sub TestForEachNestedEarlyBinding()
    Dim EnumTest As EnumTestEarlyClass
    Set EnumTest = New EnumTestEarlyClass

    EnumTest.count = 10

    Dim u As Variant, v As Variant, w As Variant
    For Each u In EnumTest
        Debug.Print u
        If u = EnumTest.count \ 2 Then
            For Each v In EnumTest
                Debug.Print u, v
                If v = EnumTest.count \ 2 Then
                    For Each w In EnumTest
                        Debug.Print u, v, w
                    Next
                End If
            Next
        End If
    Next

End Sub


Private Sub TestTimerEarlyBinding()
    Dim EnumTest As EnumTestEarlyClass
    Dim v As Variant

    Dim n As Long: n = 10000
    Debug.Print "Timings for enumerators for n = "; n

    Dim i As Long
    Dim Seconds As Double

    Set EnumTest = New EnumTestEarlyClass
    EnumTest.count = n

    Stopwatch.Reset
    Stopwatch.Start
    For Each v In EnumTest
        ' Do as little as possible, just measure the loop.
        v = 0
    Next
    Seconds = Stopwatch.Halt
    Debug.Print "Custom enum (one by one)", Format$(1000 * Seconds, "Standard") & " ms"


    Dim a() As Variant: ReDim a(n - 1)
    For i = 0 To n - 1
        a(i) = i
    Next

    Stopwatch.Reset
    Stopwatch.Start
    For Each v In a
        ' Do as little as possible, just measure the loop.
        v = 0
    Next
    Seconds = Stopwatch.Halt
    Debug.Print "VB Array enumerator", Format$(1000 * Seconds, "Standard") & " ms"

    Dim c As Collection
    Set c = New Collection
    For i = 1 To n
        c.Add i
    Next

    Stopwatch.Reset
    Stopwatch.Start
    For Each v In c
        ' Do as little as possible, just measure the loop.
        v = 0
    Next
    Seconds = Stopwatch.Halt
    Debug.Print "VB Collection enumerator", Format$(1000 * Seconds, "Standard") & " ms"

    Stopwatch.Reset
    Stopwatch.Start
    For i = 0 To n - 1
        ' Do as little as possible, just measure the loop.
        a(i) = 0
    Next
    Seconds = Stopwatch.Halt
    Debug.Print "VB Array For i", Format$(1000 * Seconds, "Standard") & " ms"

    ' Timings include on how long it takes to obtain EnumTest.Items
    Stopwatch.Reset
    Stopwatch.Start
    For Each v In EnumTest.Items
    ' Don't do anything, just measure the loop.
    Next
    Seconds = Stopwatch.Halt
    Debug.Print "Items         ", Format$(1000 * Seconds, "Standard") & " ms"

End Sub
