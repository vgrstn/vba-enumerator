Attribute VB_Name = "Enumerator"
'@IgnoreModule MultipleDeclarations, HungarianNotation, UseMeaningfulName, AssignedByValParameter, FunctionReturnValueDiscarded, UnassignedVariableUsage, VariableNotAssigned, IntegerDataType, UDTMemberNotUsed
'@Folder("Module")
'@ModuleDescription("Enumerator Module.")

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
' Version: 2026.11.08
'
' Methods
' Enumerate(iterable)       Sets IEnumVARIANT interface to an iterable object
'
' The enumeration works correctly for nested loops with mixed objects as well
' as for nested loops with mixed enumerators. An error raised in the Item
' function of the iterable object or in the body of the loop reaches the caller
' as usual. The iterable object stays alive until the enumeration ends.
'
' The iterable object needs to set its IEnumVARIANT interface by the Enumerate
' method and is required to implement the IEnumerator interface for the early
' binding of the method for getting an indexed item from the iterable object.
'
' The indices First and Last can be in ascending or descending order and both
' belong to the enumeration, so First = Last is one item. An empty object
' returns Null for both First and Last.
'
' Code to be included in the iterable object:
'
' Implements IEnumerator
'
' Private Function IEnumerator_First() As Variant
'    If this.Count = 0 Then
'        IEnumerator_First = Null
'    Else
'        IEnumerator_First = 1
'    End If
' End Function
'
' Private Function IEnumerator_Last() As Variant
'    If this.Count = 0 Then
'        IEnumerator_Last = Null
'    Else
'        IEnumerator_Last = this.Count
'    End If
' End Function
'
' Private Function IEnumerator_Item(ByVal index As Long) As Variant
'    (return the item at index, with Set for an object)
' End Function
'
' '@Enumerator
' Public Function Enumerate() As IEnumVARIANT
'    Set Enumerate = Enumerator.Enumerate(Me)
' End Function
'
' Timings (ms) for n = 10.000, see BenchmarkEnumerator
' Iterable object with IEnumerator interface (For Each)     3.01 (API 16.04 ms)
' Iterable object - Item(index) (For)                       0.83
' Iterable object - Items array (For Each)                  0.32
' VB Collection - VB enumerator (For Each)                  0.26
' VB Array - VB enumerator (For Each)                       0.12
' VB Array - VB loop (For)                                  0.08
'
' The enumeration is about 10 times slower than For Each over a VB Collection:
' the VB Collection enumerator is native code, whereas here every item takes a
' call of IEnumVARIANT_Next and of the Item function, both VBA code. Just
' calling Item already takes 3 times as long as a VB Collection step. For
' speed, enumerate the Items array of the object instead.
'
' Starting a For Each takes about 4 µs (API 8 µs), against 0.2 µs for a VB
' Collection. The memory blocks of the IEnumVARIANT objects are reused through
' a stack, so a new block is only allocated when more enumerations are active
' at the same time than ever before. The blocks are never freed.
'
' The unit tests are in TestEnumEarlyBinding and the timings in
' BenchmarkEnumerator.
'
' The original ideas for a custom enumerator using a typelib and redefining
' the IEnumVARIANT interface routines in a standard module originate from
' Hardcore Visual Basic 5.0 by Bruce McKinney.
'
' The implementation without using a typelib is based on work by Dexter
' Freivald whose original code was for 32 bits and was using late binding.
' https://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib
'
' An alternative method is to define an enumeration procedure in the iterable
' object by copying its items to an embedded VB Collection and subsequently
' exposing the IEnumVARIANT interface of this VB Collection. Another alternative
' is to let the object export the items to an array. The latter is what is used
' by the VB Dictionary. Both of these alternative methods export the items
' "at once" whereas the Enumerate exports the enumerated items "one by one".
'
' Using the VarCopyToPtr API to copy a variant value to a variant address is
' relatively slow. APIs are best avoided in the hot zone for raw performance:
' every Declare call takes about 1.4 µs. The memory blocks are therefore reused
' instead of allocated and freed for every enumeration (CoTaskMemAlloc and
' CoTaskMemFree together take about 6 µs), and pointers are written with the
' Variant ByRef construct. On 64 bits a LongLong can't be written by reference
' (error 458), so a pointer is written as a Currency with the same 8 bytes.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compiler Directives
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The IEnumVARIANT_Next function is critical to the overall performance and it
' scales with the number of items enumerated. IEnumVARIANT_Next needs to copy a
' variant to a memory address. Depending on the API compiler directive it uses
' the VarCopyToPtr API or the 5x faster CopyVarByRef Variant ByRef method.
' The same directive chooses CopyMemory or the Variant ByRef construct for
' writing pointers. The early binding approach calls the IEnum.Item(index)
' Function.

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

' The memory blocks of released IEnumVARIANT objects are kept for reuse in a
' stack (top at position 1): Enumerate and Clone take a block from the stack and
' only allocate a new one when it is empty, IUnknown_Release puts it back.
Private FreeBlocks As Collection

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

'@Ignore NonReturningFunction
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
        Set FreeBlocks = New Collection
    End If

    ' Construct the synthetic IEnumVARIANT object.
    Dim obj As TENUM
    With obj
        .pvTable = VarPtr(vTable(0))
        ' Check if the IEnumerator interface is defined for the iterable object.
        If TypeOf iterable Is IEnumerator Then
            Set .IEnum = iterable
        Else
            Err.Raise vbErrorInvalidProcedureCall, , "IEnumerator not implemented"
        End If
        ' First and Last are both Null for an empty object.
        Dim FirstIndex As Variant: FirstIndex = .IEnum.First
        Dim LastIndex As Variant: LastIndex = .IEnum.Last
        If VBA.IsNull(FirstIndex) And VBA.IsNull(LastIndex) Then
            ' An empty range, as for an empty array.
            .First = 1
            .Last = 0
            .Step = 1
        ElseIf VBA.IsNull(FirstIndex) Or VBA.IsNull(LastIndex) Then
            Err.Raise vbErrorInvalidProcedureCall, , "IEnumerator First and Last are inconsistent"
        Else
            .First = FirstIndex
            .Last = LastIndex
            ' The indices can be either in ascending or descending order.
            .Step = VBA.IIf(.First <= .Last, 1, -1)
        End If
        .nRef = 1
        .Current = .First
    End With

    Dim MemoryBlock As LongPtr: MemoryBlock = TakeBlock(LenB(obj))
    If MemoryBlock = vbNullPtr Then Err.Raise vbErrorOutOfMemory
    CopyMemory ByVal MemoryBlock, obj, LenB(obj)

    ' The memory block now holds the reference to the iterable object: clear it
    ' in obj without a release, IUnknown_Release releases it.
#If API Then
    CopyMemory ByVal VarPtr(Enumerate), MemoryBlock, vbSizeLongPtr
    CopyMemory ByVal VarPtr(obj.IEnum), vbNullPtr, vbSizeLongPtr
#Else
    '@Ignore ValueRequired
    CopyPtrByRef VarPtr(Enumerate), MemoryBlock, VarByRef.vt, VarByRef.ref
    '@Ignore ValueRequired
    CopyPtrByRef VarPtr(obj.IEnum), vbNullPtr, VarByRef.vt, VarByRef.ref
#End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Queries a COM object for a pointer to one of its interfaces."
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
#If API Then
        CopyMemory ByVal ppvObj, VarPtr(obj), vbSizeLongPtr
#Else
        CopyPtrByRef ppvObj, VarPtr(obj), VarByRef.vt, VarByRef.ref
#End If
        IUnknown_AddRef obj
        IUnknown_QueryInterface = S_OK
    Else
        IUnknown_QueryInterface = E_NOINTERFACE
    End If

End Function


'@Description "Increments the reference count for an interface pointer to a COM object."
Private Function IUnknown_AddRef(ByRef obj As TENUM) As Long

    obj.nRef = obj.nRef + 1
    IUnknown_AddRef = obj.nRef

End Function


'@Description "Decrements the reference count for an interface pointer to a COM object."
Private Function IUnknown_Release(ByRef obj As TENUM) As Long

    obj.nRef = obj.nRef - 1
    IUnknown_Release = obj.nRef

    If obj.nRef = 0 Then
        Set obj.IEnum = Nothing
        ' Put the memory block on the stack for reuse.
        If FreeBlocks.count = 0 Then
            FreeBlocks.Add VarPtr(obj)
        Else
            FreeBlocks.Add VarPtr(obj), Before:=1
        End If
    End If

End Function


'@Description "Retrieves the next item in the enumeration sequence."
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


'@Description "Skips over a number of elements in the enumeration sequence."
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
        obj.Current = obj.Last + VBA.Sgn(obj.Step)
        IEnumVARIANT_Skip = S_FALSE
    End Select

End Function


'@Description "Resets the enumeration sequence to the beginning."
Private Function IEnumVARIANT_Reset(ByRef obj As TENUM) As Long

    obj.Current = obj.First
    IEnumVARIANT_Reset = S_OK

End Function


'@Description "Creates a copy of the current state of enumeration."
Private Function IEnumVARIANT_Clone(ByRef obj As TENUM, ByVal ppEnum As LongPtr) As Long

    If ppEnum = vbNullPtr Then
        IEnumVARIANT_Clone = E_INVALIDARG
        Exit Function
    End If

    ' UDT assignment AddRefs IEnum — no Set needed
    Dim Copy As TENUM: Copy = obj
    Copy.nRef = 1

    Dim MemoryBlock As LongPtr: MemoryBlock = TakeBlock(LenB(obj))
    If MemoryBlock = vbNullPtr Then
        IEnumVARIANT_Clone = E_OUTOFMEMORY
        Exit Function
    End If
    CopyMemory ByVal MemoryBlock, Copy, LenB(obj)
    IEnumVARIANT_Clone = S_OK

    ' The memory block now holds the reference to the iterable object: clear it
    ' in Copy without a release, IUnknown_Release releases it.
#If API Then
    CopyMemory ByVal ppEnum, MemoryBlock, vbSizeLongPtr
    CopyMemory ByVal VarPtr(Copy.IEnum), vbNullPtr, vbSizeLongPtr
#Else
    CopyPtrByRef ppEnum, MemoryBlock, VarByRef.vt, VarByRef.ref
    '@Ignore ValueRequired
    CopyPtrByRef VarPtr(Copy.IEnum), vbNullPtr, VarByRef.vt, VarByRef.ref
#End If

End Function


'@Description "Returns a free memory block from the stack, or a new one if the stack is empty."
Private Function TakeBlock(ByVal NBytes As Long) As LongPtr

    If FreeBlocks.count = 0 Then
        TakeBlock = CoTaskMemAlloc(NBytes)
    Else
        TakeBlock = FreeBlocks.Item(1)
        FreeBlocks.Remove 1
    End If

End Function


'@Description "Returns True if id is IID_IUnknown GUID."
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


'@Description "Returns True if id is IID_IEnumVARIANT GUID."
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


#If API = False Then
'@Description "Initializes the Variant ByRef construct."
Private Sub InitializeVarByRef()

    VarByRef.vt = VarPtr(VarByRef.ref)
    CopyMemory VarByRef.vt, VBA.vbInteger Or VT_BYREF, vbSizeInteger

End Sub
#End If


#If API = False Then
'@Description "Copies a Variant to a memory address using the Variant ByRef construct."
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
'@Description "Copies a Long to a memory address using the Variant ByRef construct."
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


#If API = False Then
'@Description "Copies a pointer to a memory address using the Variant ByRef construct."
Private Sub CopyPtrByRef( _
    ByVal address As LongPtr, _
    ByVal value As LongPtr, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
)

#If Win64 Then
    ' A LongLong by reference is not supported (error 458). The 8 bytes of the
    ' LongLong are read as a Currency by changing the type only, and that
    ' Currency is written by reference: the bits are copied, nothing is calculated.
    VarByRef.ref = value
    vt = VBA.vbCurrency
    Dim Bits As Currency: Bits = ref

    VarByRef.ref = address
    vt = VBA.vbCurrency Or VT_BYREF
    ref = Bits
#Else
    VarByRef.ref = address
    vt = VBA.vbLong Or VT_BYREF
    ref = value
#End If

End Sub
#End If
