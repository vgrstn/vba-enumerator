# vba-enumerator
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for `IEnumVARIANT` interface implementation — no typelib required, early binding via `IEnumerator`.

Implements the full `IEnumVARIANT` interface (`Next`, `Skip`, `Reset`, `Clone`) in a standard module using `AddressOf` and a heap-allocated vtable. Items are retrieved one by one via the `IEnumerator` interface, which the iterable Class must implement.

> For the late-binding variant (no interface required, uses `CallByName`) see [vba-enumerator-late-binding](https://github.com/vgrstn/vba-enumerator-late-binding).

---

## 📦 Features

- **`For Each` without a typelib** — pure VBA, no external dependencies
- **Early binding** — items retrieved via `IEnumerator.Item(index)`, direct vtable call, no dispatch overhead
- **Nested loops** — works correctly for nested `For Each` with mixed objects and mixed enumerators
- **Ascending and descending** — `First` and `Last` can be in either order; `First = Last` is one item
- **Empty objects** — an empty object returns `Null` for both `First` and `Last`, and the loop body never runs
- **Errors reach the caller** — an error raised in `Item` or in the loop body propagates as usual
- **Fast variant copy** — uses a Variant ByRef construct (~5× faster than `VariantCopy` API); switch to API mode with `#Const API = True`
- **Reused memory blocks** — released enumerators go on a stack, so after the first `For Each` no memory is allocated
- **Full COM lifecycle** — `QueryInterface`, `AddRef`, `Release`, `Clone` all correctly implemented
- x86 / x64 compatible via `LongPtr` and `#If Win64`
- Pure VBA, zero dependencies, Rubberduck-friendly annotations

---

## 📁 Files

| File | Type | Description |
|---|---|---|
| `IEnumerator.cls` | Interface | Defines `First`, `Last`, `Item` — implement this in your iterable Class |
| `Enumerator.bas` | Module | `Enumerate(iterable)` — the main entry point |
| `CEnumTestEarly.cls` | Example | Simple iterable Class implementing `IEnumerator` |
| `EnumTestEarly.bas` | Example | `For Each` examples and performance timings |

Each file has a corresponding `_WithAttributes` version (e.g. `Enumerator_WithAttributes.bas`) with [Rubberduck](https://rubberduckvba.com/) annotations removed and VB attributes baked in. Import the `_WithAttributes` files if you are not using Rubberduck.

> **Note:** `EnumTestEarly.bas` uses a `Stopwatch` module for timing measurements. Remove or replace those calls if you do not have it.

---

## ⚙️ Public Interface

### `Enumerator` module

| Member | Description |
|---|---|
| `Enumerate(iterable)` | Returns a synthetic `IEnumVARIANT` for the iterable object. Raises error 424 if `iterable` is `Nothing`, and error 5 if it does not implement `IEnumerator` or if only one of `First` and `Last` is `Null`. |

### `IEnumerator` interface

| Method | Description |
|---|---|
| `First()` | Returns the index of the first item as a `Variant`, `Null` for an empty object |
| `Last()` | Returns the index of the last item as a `Variant`, `Null` for an empty object |
| `Item(index)` | Returns the item at the given index |

`First` and `Last` both belong to the enumeration and can be in ascending or descending order. An empty object returns `Null` for both; one `Null` without the other breaks the rules of the interface.

---

## 🚀 Quick Start

**1. Implement `IEnumerator` in your Class:**

```vb
Implements IEnumerator

Private Function IEnumerator_First() As Variant
    If this.Count = 0 Then
        IEnumerator_First = Null
    Else
        IEnumerator_First = 1
    End If
End Function

Private Function IEnumerator_Last() As Variant
    If this.Count = 0 Then
        IEnumerator_Last = Null
    Else
        IEnumerator_Last = this.Count
    End If
End Function

Private Function IEnumerator_Item(ByVal index As Long) As Variant
    If VBA.IsObject(this.Items(index)) Then
        Set IEnumerator_Item = this.Items(index)
    Else
        IEnumerator_Item = this.Items(index)
    End If
End Function
```

**2. Add the `Enumerate` function:**

```vb
'@Enumerator
Public Function Enumerate() As IEnumVARIANT
    Set Enumerate = Enumerator.Enumerate(Me)
End Function
```

**3. Use `For Each`:**

```vb
Dim obj As MyClass
Set obj = New MyClass
' ... populate obj ...

Dim v As Variant
For Each v In obj
    Debug.Print v
Next
```

---

## ⏱️ Performance

Timings for `n = 10,000` items:

| Method | Time (ms) |
|---|---|
| Custom enumerator — `For Each` (VarByRef) | 3.01 |
| Custom enumerator — `For Each` (API) | 16.04 |
| Iterable object — `For i`, `Item(i)` | 0.83 |
| Iterable object — `For Each` over an `Items` array | 0.32 |
| VB Collection — `For Each` | 0.26 |
| VB Array — `For Each` | 0.12 |
| VB Array — `For i` | 0.08 |

The custom enumerator is about 10 times slower than `For Each` over a VB Collection. The VB Collection enumerator is native code, whereas here every item takes a call of `IEnumVARIANT_Next` and of the `Item` function, both VBA code: just calling `Item` already takes 3 times as long as a VB Collection step. For speed, enumerate an `Items` array instead. The value of the custom enumerator is `For Each` over your own Class without copying the items first.

Starting a `For Each` takes about 4 µs (API 8 µs), against 0.2 µs for a VB Collection. It used to take 13.6 µs, before the memory blocks were reused and the pointer writes moved from `CopyMemory` to the Variant ByRef construct.

---

## ⚙️ Compiler directive

| Directive | Default | Description |
|---|---|---|
| `#Const API` | `False` | Copy Variants and pointers via the Variant ByRef construct (fast) or via the `VariantCopy` and `CopyMemory` APIs (slow) |

`#Const API = True` uses `VarCopyToPtr` (`VariantCopy` from oleaut32). `#Const API = False` uses the Variant ByRef construct — approximately 5× faster. See performance table above.

---

## 🧠 How it works

VBA's `For Each` requires the iterable object to expose `IEnumVARIANT` via `_NewEnum`. Rather than using a typelib to define `IEnumVARIANT`, this module:

1. Takes a block of heap memory, large enough to hold the enumerator state (`TENUM` UDT), from a stack of free blocks — or allocates one with `CoTaskMemAlloc` when the stack is empty
2. Builds a vtable of function pointers (`AddressOf`) for the seven COM methods: `QueryInterface`, `AddRef`, `Release`, `Next`, `Skip`, `Reset`, `Clone`
3. Writes the vtable pointer into the first field of `TENUM` — making it a valid COM object
4. Overwrites the return value of `Enumerate` with the heap pointer — returning the synthetic object as `IEnumVARIANT`
5. Hands the reference to the iterable object over to the heap block, which keeps it alive until the enumeration is released

Based on work by Dexter Freivald (32-bit, late binding) and ideas from *Hardcore Visual Basic 5.0* by Bruce McKinney.

---

## 🧠 Implementation notes

### `TENUM` — synthetic COM object layout

```vb
Private Type TENUM
    pvTable  As LongPtr     ' MUST be first — COM reads vtable pointer at offset 0
    IEnum    As IEnumerator ' early-bound reference to the iterable object
    nRef     As Long        ' COM reference count
    First    As Long        ' index of first item
    Last     As Long        ' index of last item
    Current  As Long        ' index of current position
    Step     As Long        ' +1 (ascending) or -1 (descending)
End Type
```

`pvTable` is first because COM requires a pointer to the vtable at offset 0 of any COM object. `IEnum As IEnumerator` is early-bound — the compiler emits a direct vtable call to `.Item(i)` rather than a late-bound `IDispatch.Invoke`, which is the dominant cost in `IEnumVARIANT_Next`. `Step` is derived from `First`/`Last` at construction time; ascending and descending enumeration share the same code paths throughout. An empty object gets the empty range `First = 1`, `Last = 0`, `Step = 1`, as for an empty array.

### vtable — built once per session

```vb
Static vTable(0 To 6) As LongPtr
If vTable(0) = vbNullPtr Then
    vTable(0) = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
    ...
    vTable(6) = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
End If
```

The `Static` array persists for the lifetime of the VBA session. `vTable(0) = vbNullPtr` is the once-only sentinel — subsequent calls to `Enumerate` reuse the same vtable without rebuilding it. Slots 0–2 are the three `IUnknown` methods; slots 3–6 are the four `IEnumVARIANT` methods.

### Return value trick

```vb
CopyPtrByRef VarPtr(Enumerate), MemoryBlock, VarByRef.vt, VarByRef.ref
```

`Enumerate` returns `IEnumVARIANT`. VBA stores the return value as an object pointer at `VarPtr(Enumerate)`. Overwriting those bytes with the heap block address makes VBA believe that address is a valid COM object — which it is, because the vtable pointer sits at offset 0. This is why `Enumerate` carries `'@Ignore NonReturningFunction`; the return value is set by raw memory write, not by a `Set Enumerate = ...` assignment.

### Reference counting — the heap block owns the reference

`CopyMemory ByVal MemoryBlock, obj, LenB(obj)` copies the `TENUM` struct to the heap as raw bytes — it copies the `IEnum` pointer without calling `AddRef`. The local `obj` holds the one counted reference, taken by `Set .IEnum = iterable`. Clearing that pointer in `obj` without a `Release` hands the reference over to the heap block:

```vb
CopyPtrByRef VarPtr(obj.IEnum), vbNullPtr, VarByRef.vt, VarByRef.ref
```

When `IUnknown_Release` sees `nRef = 0`, it releases the reference held by the block and puts the block on the stack of free blocks:

```vb
Set obj.IEnum = Nothing                ' release — iterable may now be destroyed
FreeBlocks.Add VarPtr(obj), Before:=1  ' (or Add when the stack is empty)
```

The same pattern applies in `IEnumVARIANT_Clone`.

> Earlier versions kept the iterable object alive with a `KeepAlive` Collection keyed by the block address. Its `On Error` handling cleared the `Err` object when an enumerator was released while an error was being raised, so an error in `Item` or in the body of a `For Each` reached the caller one level up as `Err.Number = 0`. The handover above needs no error handling.

### Memory blocks — a stack for reuse

```vb
Private FreeBlocks As Collection

Private Function TakeBlock(ByVal NBytes As Long) As LongPtr
    If FreeBlocks.count = 0 Then
        TakeBlock = CoTaskMemAlloc(NBytes)
    Else
        TakeBlock = FreeBlocks.Item(1)
        FreeBlocks.Remove 1
    End If
End Function
```

A `Declare` call takes about 1.4 µs, and `CoTaskMemAlloc` with `CoTaskMemFree` about 6 µs. Released blocks are therefore never freed but kept in a Collection used as a stack (top at position 1, no key). Only as many blocks are allocated as the most enumerations that were active at the same time — usually one or two.

### `IEnumVARIANT_Next` — hot path

```vb
For i = obj.Current To obj.Last Step obj.Step
    CopyVarByRef rgVar, obj.IEnum.Item(i), VarByRef.vt, VarByRef.ref
    NumberFetched = NumberFetched + 1
    If NumberFetched = celt Then Exit For
    rgVar = rgVar + vbSizeVariant
Next
obj.Current = obj.Current + NumberFetched * obj.Step
```

Per iteration: one early-bound `.Item(i)` call and one Variant copy to the destination address. `rgVar` is advanced by `vbSizeVariant` (16 bytes x86 / 24 bytes x64) for multi-element fetches; VBA's `For Each` always requests one item at a time (`celt = 1`), so the inner `Exit For` fires immediately. `obj.Current` is updated in one step after the loop.

### `IEnumVARIANT_Skip` — direction-aware, `Select Case True`

```vb
Select Case True
Case celt = 0:  IEnumVARIANT_Skip = S_OK
Case celt < 0:  IEnumVARIANT_Skip = E_INVALIDARG
Case celt <= (obj.Step * (obj.Last - obj.Current) + 1)
    obj.Current = obj.Current + celt * obj.Step
    IEnumVARIANT_Skip = S_OK
Case Else
    obj.Current = obj.Last + VBA.Sgn(obj.Step)
    IEnumVARIANT_Skip = S_FALSE
End Select
```

`obj.Step * (obj.Last - obj.Current) + 1` gives the remaining item count for both ascending (`Step = 1`) and descending (`Step = -1`) without a branch. `obj.Current` is only mutated after the bounds check. For overshoot, `obj.Current = obj.Last + VBA.Sgn(obj.Step)` places current one step past the end — the same post-exhaustion state that `Next` leaves.

### `IEnumVARIANT_Clone` — state snapshot

```vb
' UDT assignment AddRefs IEnum — no Set needed
Dim Copy As TENUM: Copy = obj
Copy.nRef = 1
```

`Copy = obj` copies all fields. VBA automatically calls `AddRef` on the embedded `IEnum As IEnumerator` during the UDT assignment, and that reference is handed over to the clone's heap block as in `Enumerate`. The clone starts at `nRef = 1` regardless of the original's count, and captures the enumeration position at the moment of cloning.

### Variant ByRef construct

```vb
Private Type CONSTRUCT
    vt  As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
```

`InitializeVarByRef` sets `vt` to `VT_INTEGER | VT_BYREF` pointing at `ref`. `CopyVarByRef` writes a Variant to any address without an API call; `CopyLngByRef` uses the same mechanism (`vbLong | VT_BYREF`) for `pceltFetched`. Both helpers receive `vt` and `ref` as `ByRef Variant` parameters rather than accessing `VarByRef` directly, avoiding a global UDT indirection per call.

`CopyPtrByRef` writes a pointer the same way. On x64 a `LongLong` by reference is not supported (error 458), so the 8 bytes are copied as a `Currency`:

```vb
VarByRef.ref = value            ' the pointer as a LongLong
vt = VBA.vbCurrency             ' the same 8 bytes, now read as a Currency
Dim Bits As Currency: Bits = ref
VarByRef.ref = address
vt = VBA.vbCurrency Or VT_BYREF
ref = Bits                      ' the 8 bytes written as they are
```

Only the type is changed, so the bits are copied and nothing is calculated. (An earlier version wrote `CCur(value) * 0.0001@`; that is exact for every user-mode address, but silently wrong above 922,337,203,685,477.)

### GUID comparison — field-by-field

```vb
' Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
IsIID_IUnknown = (id.Data1 = &H0) And (id.Data2 = &H0) And ... And (id.Data4(7) = &H46)
```

Direct integer comparison on the `GUID` UDT fields — no string parsing. The commented `Const` documents the expected GUID string without runtime cost.

### `VARSIZE` — x86/x64 portability

```vb
#If Win64 Then
    vbSizeLongPtr = 8
    vbSizeVariant = 24
#Else
    vbSizeLongPtr = 4
    vbSizeVariant = 16
#End If
```

A `Variant` is 16 bytes on x86 and 24 bytes on x64. All `CopyMemory` sizes and pointer arithmetic use these constants — no hard-coded values appear in the code.

---

## 📝 Changes

**2026-11-08**
- Fixed: errors raised in `Item` or in the body of a `For Each` were lost one level up (`Err.Number = 0`); the `KeepAlive` Collection is gone.
- Changed: `IEnumerator.First` and `Last` return a `Variant`; `Null` for both means an empty object. Before, an empty object with `First = 1` and `Last = 0` was enumerated as two items in descending order.
- Faster start of a `For Each` (13.6 → 4 µs): memory blocks are reused through a stack, and pointers are written with the Variant ByRef construct.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
