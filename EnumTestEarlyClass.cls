VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed, UseMeaningfulName
'@Folder("Class")
'@ModuleDescription("Enumerator Test Class.")

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

' Wrapper for private data.
Private Type TENUMTEST
    n As Long
    a() As Variant
End Type
Private this As TENUMTEST

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Interface methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Implements IEnumerator


'@Description "Returns the index of the first enumerated item."
Private Function IEnumerator_First() As Long

    IEnumerator_First = 1

End Function


'@Description "Returns the index of the last enumerated item."
Private Function IEnumerator_Last() As Long

    IEnumerator_Last = this.n

End Function


'@Description "Returns an item by its index."
Private Function IEnumerator_Item(ByVal index As Long) As Variant

    If VBA.IsObject(this.a(index)) Then
        Set IEnumerator_Item = this.a(index)
    Else
        IEnumerator_Item = this.a(index)
    End If

End Function


'@Enumerator
Public Function Enumerate() As IEnumVARIANT

    Set Enumerate = Enumerator.Enumerate(Me)

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get count() As Long
    count = this.n
End Property
Public Property Let count(ByVal RHS As Long)
    this.n = RHS
    ReDim this.a(1 To this.n)
    Dim i As Long
    For i = 1 To this.n
        this.a(i) = i
    Next
End Property


'@DefaultMember
Public Property Get Item(ByVal index As Long) As Variant
    Item = this.a(index)
End Property
Public Property Let Item(ByVal index As Long, ByVal RHS As Variant)
    this.a(index) = RHS
End Property


Public Property Get Items() As Variant

    Items = this.a

End Property
