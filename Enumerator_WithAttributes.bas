Attribute VB_Name = "Enumerator"

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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Enumerator Module using Early Binding
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.10.10
'
' Methods
' Enumerate(iterable)       Sets IEnumVARIANT interface to an iterable object
'
' The enumeration works correctly for nested loops with mixed objects as well
' as for nested loops with mixed enumerators.
'
' The iterable object needs to set its IEnumVARIANT interface by the Enumerate
' method and is required to implement the IEnumerator interface for the early
' binding of the method for getting an indexed item from the iterable object.
'
' Code to be included in the iterable object:
'
' Implements IEnumerator
'
' see IEnumerator for functions to implement.
'
' '@Enumerator
' Public Function Enumerate() As IEnumVARIANT
'    Set Enumerate = Enumerator.Enumerate(Me)
' End Function
'
' Timings (ms) for n = 10.000
' Iterable object with IEnumerator interface (For Each)     2.69 (API 15 ms)
' VB Collection - VB enumerator (For Each)                  0.21
' VB Array - VB enumerator (For Each)                       0.14
' VB Array - VB loop (For)                                  0.08
'
' The original ideas for a custom enumerator using a typelib and redefining
' the IEnumVARIANT interface routines in a standard module originate from
' Hardcore Visual Basic 5.0 by Bruce McKinney.
'
' The implementation without using a typelib is based on work by Dexter
' Freivald who's original code was for 32 bits and was using late binding.
' https://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib
'
' An alternative method is to define an enumeration procedure in the iterable
' object by copying its items to an embedded VB Collection and subsequently
' exposing the IEnumVARIANT interface of this VB Collection. Another alternative
' is to let the object export the items to an array. The latter is what is used
' by the VB Dictionary. Both of these alternative methods export the items
' "at once" whereas the Enumerate exports the enumerated items "one by one".

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compiler Directives
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The IEnumVARIANT_Next function is critical to the overall performance and it
' scales with the number of items enumerated. IEnumVARIANT_Next needs to copy a
' variant to a memory address. Depending on the API compiler directive it uses
' the VarCopyToPtr API or the 5x faster CopyVarByRef Variant ByRef method.
' The early binding approach calls the Me.IEnum.Item(index) Function.

#Const API = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private API Declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlmovememory
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal NBytes As Long _
)

' https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cotaskmemalloc
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32.dll" ( _
    ByVal cb As Long _
) As LongPtr

' https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cotaskmemfree
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" ( _
    ByVal pv As LongPtr _
)

#If API Then
' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-variantcopy
Private Declare PtrSafe Function VarCopyToPtr Lib "oleaut32.dll" Alias "VariantCopy" ( _
    ByVal pvDest As LongPtr, _
    ByVal pvSrc As Variant _
) As Long
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The vbNullPtr constant is the null pointer (not VB defined).
Private Const vbNullPtr As LongPtr = 0

' The memory size of intrinsic data types.
Private Enum VARSIZE
    vbSizeInteger = 2
    vbSizeLong = 4
#If Win64 Then
    vbSizeLongPtr = 8
    vbSizeVariant = 24
#Else
    vbSizeLongPtr = 4
    vbSizeVariant = 16
#End If
End Enum

' Selected HRESULT constants.
Private Enum HRESULT
    S_OK = &H0                          ' Operation successful, returns True
    S_FALSE = &H1                       ' Operation successful, returns False
    E_NOTIMPL = &H80004001              ' Not implemented
    E_NOINTERFACE = &H80004002          ' No such interface supported
    E_POINTER = &H80004003              ' Pointer that is not valid
    E_ABORT = &H80004004
    E_FAIL = &H80004005
    E_OUTOFMEMORY = &H8007000E          ' Failed to allocate necessary memory
    E_INVALIDARG = &H80070057           ' One of the arguments is not valid
End Enum

' Selected VBA errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
    vbErrorOutOfMemory = 7
    vbErrorObjectRequired = 424
End Enum

' GUID is the UDT for the Global Unique Identifier.
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' The IEnumVARIANT status is captured in an UDT.
Private Type TENUM
    pvTable As LongPtr
    IEnum As IEnumerator
    nRef As Long
    First As Long
    Last As Long
    Current As Long
    Step As Long
End Type

#If API = False Then
' Variant ByRef construct for memory access by address.
Private Const VT_BYREF As Integer = &H4000
Private Type CONSTRUCT
    vt As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Enumerate(ByVal iterable As Object) As IEnumVARIANT

    If iterable Is Nothing Then Err.Raise vbErrorObjectRequired

    ' Initialize the vTable with the redefined IUnknown/IEnumVARIANT functions.
    Static vTable(0 To 6) As LongPtr
    If vTable(0) = vbNullPtr Then
        vTable(0) = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
        vTable(1) = VBA.CLngPtr(AddressOf IUnknown_AddRef)
        vTable(2) = VBA.CLngPtr(AddressOf IUnknown_Release)
        vTable(3) = VBA.CLngPtr(AddressOf IEnumVARIANT_Next)
        vTable(4) = VBA.CLngPtr(AddressOf IEnumVARIANT_Skip)
        vTable(5) = VBA.CLngPtr(AddressOf IEnumVARIANT_Reset)
        vTable(6) = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
#If API = False Then
        ' Initialize the Variant ByRef construct.
        InitializeVarByRef
#End If
    End If

    ' Construct the synthetic IEnumVARIANT object.
    Dim obj As TENUM
    With obj
        .pvTable = VarPtr(vTable(0))
        If TypeOf iterable Is IEnumerator Then
            Set .IEnum = iterable
            .First = .IEnum.First
            .Last = .IEnum.Last
        Else
            Err.Raise vbErrorInvalidProcedureCall, , "IEnumerator not implemented"
        End If
        .nRef = 1
        .Current = .First
        ' The indices can be either in ascending or descending order.
        .Step = VBA.IIf(.First <= .Last, 1, -1)
    End With

    Dim MemoryBlock As LongPtr: MemoryBlock = CoTaskMemAlloc(LenB(obj))
    If MemoryBlock = vbNullPtr Then Err.Raise vbErrorOutOfMemory
    CopyMemory ByVal MemoryBlock, obj, LenB(obj)
    CopyMemory ByVal VarPtr(Enumerate), MemoryBlock, vbSizeLongPtr

    ' When obj goes out of scope nref for iterable is decreased.
    ' KeepAlive compensates by increasing the nref for iterable.
    Set KeepAlive(MemoryBlock) = obj.IEnum

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IUnknown_QueryInterface( _
    ByRef obj As TENUM, _
    ByRef riid As GUID, _
    ByVal ppvObj As LongPtr _
) As Long

    If ppvObj = vbNullPtr Then
        IUnknown_QueryInterface = E_POINTER
        Exit Function
    End If

    If IsIID_IUnknown(riid) Or IsIID_IEnumVARIANT(riid) Then
        CopyMemory ByVal ppvObj, VarPtr(obj), vbSizeLongPtr
        IUnknown_AddRef obj
        IUnknown_QueryInterface = S_OK
    Else
        IUnknown_QueryInterface = E_NOINTERFACE
    End If

End Function


Private Function IUnknown_AddRef(ByRef obj As TENUM) As Long

    obj.nRef = obj.nRef + 1
    IUnknown_AddRef = obj.nRef

End Function


Private Function IUnknown_Release(ByRef obj As TENUM) As Long

    obj.nRef = obj.nRef - 1
    IUnknown_Release = obj.nRef

    If obj.nRef = 0 Then
        Set KeepAlive(VarPtr(obj)) = Nothing
        CoTaskMemFree VarPtr(obj)
    End If

End Function


Private Function IEnumVARIANT_Next( _
    ByRef obj As TENUM, _
    ByVal celt As Long, _
    ByVal rgVar As LongPtr, _
    ByVal pceltFetched As LongPtr _
) As Long

    If rgVar = vbNullPtr Then
        IEnumVARIANT_Next = E_INVALIDARG
        Exit Function
    End If

    ' Set pceltFetched to 0 if the pointer is provided.
    If pceltFetched <> vbNullPtr Then
#If API Then
        CopyMemory ByVal pceltFetched, 0, vbSizeLong
#Else
        CopyLngByRef pceltFetched, 0, VarByRef.vt, VarByRef.ref
#End If
    End If

    ' Only continue with loop for celt > 0.
    If celt <= 0 Then
        If celt < 0 Then
            IEnumVARIANT_Next = E_INVALIDARG
        Else
            IEnumVARIANT_Next = S_OK
        End If
        Exit Function
    End If

    ' Get the next item(s) from the iterable object.
    Dim i As Long, NumberFetched As Long
    For i = obj.Current To obj.Last Step obj.Step
#If API Then
        If VarCopyToPtr(rgVar, obj.IEnum.Item(i)) <> S_OK Then
            IEnumVARIANT_Next = E_FAIL
            Exit Function
        End If
#Else
        CopyVarByRef rgVar, obj.IEnum.Item(i), VarByRef.vt, VarByRef.ref
#End If
        NumberFetched = NumberFetched + 1
        If NumberFetched = celt Then Exit For
        ' Advance the pointer to the next element in the destination array.
        rgVar = rgVar + vbSizeVariant
    Next
    obj.Current = obj.Current + NumberFetched * obj.Step

    ' Set pceltFetched to NumberFetched if the pointer is provided.
    If pceltFetched <> vbNullPtr Then
#If API Then
        CopyMemory ByVal pceltFetched, NumberFetched, vbSizeLong
#Else
        CopyLngByRef pceltFetched, NumberFetched, VarByRef.vt, VarByRef.ref
#End If
    End If

    ' Return S_OK if the number of fetched items matches the requested number.
    If NumberFetched = celt Then
        IEnumVARIANT_Next = S_OK
    Else
        IEnumVARIANT_Next = S_FALSE
    End If

End Function


Private Function IEnumVARIANT_Skip(ByRef obj As TENUM, ByVal celt As Long) As Long

    Select Case True
    Case celt = 0
        IEnumVARIANT_Skip = S_OK
    Case celt < 0
        IEnumVARIANT_Skip = E_INVALIDARG
    Case celt <= (obj.Step * (obj.Last - obj.Current) + 1)
        obj.Current = obj.Current + celt * obj.Step
        IEnumVARIANT_Skip = S_OK
    Case Else
        ' For overshoot set one-past-end.
        obj.Current = obj.Last + obj.Step
        IEnumVARIANT_Skip = S_FALSE
    End Select

End Function


Private Function IEnumVARIANT_Reset(ByRef obj As TENUM) As Long

    obj.Current = obj.First
    IEnumVARIANT_Reset = S_OK

End Function


Private Function IEnumVARIANT_Clone(ByRef obj As TENUM, ByVal ppEnum As LongPtr) As Long

    If ppEnum = vbNullPtr Then
        IEnumVARIANT_Clone = E_INVALIDARG
        Exit Function
    End If

    Dim Copy As TENUM: Copy = obj
    Set Copy.IEnum = obj.IEnum
    Copy.nRef = 1

    Dim MemoryBlock As LongPtr: MemoryBlock = CoTaskMemAlloc(LenB(obj))
    If MemoryBlock = vbNullPtr Then
        IEnumVARIANT_Clone = E_OUTOFMEMORY
        Exit Function
    End If
    CopyMemory ByVal MemoryBlock, Copy, LenB(obj)
    CopyMemory ByVal ppEnum, MemoryBlock, vbSizeLongPtr
    IEnumVARIANT_Clone = S_OK

    ' When Copy goes out of scope nref for iterable is decreased.
    ' KeepAlive compensates by increasing the nref for iterable.
    Set KeepAlive(MemoryBlock) = Copy.IEnum

End Function


Private Function IsIID_IUnknown(ByRef id As GUID) As Boolean

'    Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
    IsIID_IUnknown = _
        (id.Data1 = &H0) And _
        (id.Data2 = &H0) And _
        (id.Data3 = &H0) And _
        (id.Data4(0) = &HC0) And _
        (id.Data4(1) = &H0) And _
        (id.Data4(2) = &H0) And _
        (id.Data4(3) = &H0) And _
        (id.Data4(4) = &H0) And _
        (id.Data4(5) = &H0) And _
        (id.Data4(6) = &H0) And _
        (id.Data4(7) = &H46)

End Function


Private Function IsIID_IEnumVARIANT(ByRef id As GUID) As Boolean

'    Const IID_IEnumVARIANT As String = "{00020404-0000-0000-C000-000000000046}"
    IsIID_IEnumVARIANT = _
        (id.Data1 = &H20404) And _
        (id.Data2 = &H0) And _
        (id.Data3 = &H0) And _
        (id.Data4(0) = &HC0) And _
        (id.Data4(1) = &H0) And _
        (id.Data4(2) = &H0) And _
        (id.Data4(3) = &H0) And _
        (id.Data4(4) = &H0) And _
        (id.Data4(5) = &H0) And _
        (id.Data4(6) = &H0) And _
        (id.Data4(7) = &H46)

End Function


Private Property Set KeepAlive(ByVal block As LongPtr, ByVal RHS As Object)
' Increase or decrease reference count for the iterable object.

    Static Table As Collection
    If Table Is Nothing Then Set Table = New Collection

    ' The key is the address of the allocated memory block.
    Dim Key As String: Key = CStr(block)

    If RHS Is Nothing Then
        On Error Resume Next
        ' Ignore if not found.
        Table.Remove Key
        On Error GoTo 0
    Else
        On Error Resume Next
        ' Replace if present.
        Table.Remove Key
        Err.Clear
        Table.Add RHS, Key
        On Error GoTo 0
    End If

End Property


#If API = False Then
Private Sub InitializeVarByRef()

    VarByRef.vt = VarPtr(VarByRef.ref)
    CopyMemory VarByRef.vt, VBA.vbInteger Or VT_BYREF, vbSizeInteger

End Sub
#End If


#If API = False Then
Private Sub CopyVarByRef( _
    ByVal address As LongPtr, _
    ByVal value As Variant, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
)

    VarByRef.ref = address
    vt = VBA.vbVariant Or VT_BYREF
    If VBA.IsObject(value) Then
        Set ref = value
    Else
        ref = value
    End If

End Sub
#End If


#If API = False Then
Private Sub CopyLngByRef( _
    ByVal address As LongPtr, _
    ByVal value As Long, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
)

    VarByRef.ref = address
    vt = VBA.vbLong Or VT_BYREF
    ref = value

End Sub
#End If
