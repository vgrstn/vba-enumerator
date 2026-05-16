# vba-enumerator
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for IEnumVARIANT interface implementation — no typelib required early binding.

Implements the full `IEnumVARIANT` interface (`Next`, `Skip`, `Reset`, `Clone`) in a standard module using `AddressOf` and a heap-allocated vtable. Items are retrieved one by one via the `IEnumerator` interface, which the iterable Class must implement.

---

## 📦 Features

- **`For Each` without a typelib** — pure VBA, no external dependencies
- **Early binding** — items are retrieved via `IEnumerator.Item(index)`, avoiding the overhead of late-bound dispatch
- **Nested loops** — works correctly for nested `For Each` with mixed objects and mixed enumerators
- **Ascending and descending** — `First` and `Last` can be in either order
- **Fast variant copy** — uses a Variant ByRef construct (5× faster than `VariantCopy` API); switch to API mode with `#Const API = True`
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
| `EnumTestEarly.bas` | Example | `For Each` tests and performance timings |

Each file has a corresponding `_WithAttributes` version (e.g. `Enumerator_WithAttributes.bas`) with [Rubberduck](https://rubberduckvba.com/) annotations removed and VB attributes baked in. Import the `_WithAttributes` files if you are not using Rubberduck.

> **Note:** `EnumTestEarly.bas` uses a `Stopwatch` module for timing measurements. Remove or replace those calls if you do not have it.

---

## ⚙️ Public Interface

### `Enumerator` module

| Member | Description |
|---|---|
| `Enumerate(iterable)` | Returns a synthetic `IEnumVARIANT` for the iterable object. Raises an error if `iterable` is `Nothing` or does not implement `IEnumerator`. |

### `IEnumerator` interface

| Method | Description |
|---|---|
| `First()` | Returns the index of the first item |
| `Last()` | Returns the index of the last item |
| `Item(index)` | Returns the item at the given index |

---

## 🚀 Quick Start

**1. Implement `IEnumerator` in your Class:**

```vb
Implements IEnumerator

Private Function IEnumerator_First() As Long
    IEnumerator_First = 1
End Function

Private Function IEnumerator_Last() As Long
    IEnumerator_Last = this.Count
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

Timings for `n = 10,000` items (Immediate Window):

| Method | Time (ms) |
|---|---|
| Custom enumerator — `For Each` (VarByRef) | 2.89 |
| Custom enumerator — `For Each` (API) | 16.38 |
| VB Collection — `For Each` | 0.24 |
| VB Array — `For Each` | 0.13 |
| VB Array — `For i` | 0.07 |

The overhead versus a native VB Collection is primarily the early-bound `IEnumerator.Item(i)` call per element. The VarByRef variant copy is ~5× faster than the `VariantCopy` API.

---

## ⚙️ Compiler directive

| Directive | Default | Description |
|---|---|---|
| `#Const API` | `False` | Copy Variants via Variant ByRef construct (fast) or `VariantCopy` API (slow) |

`#Const API = True` uses `VarCopyToPtr` (`VariantCopy` from oleaut32). `#Const API = False` uses the Variant ByRef construct — approximately 5× faster. See performance table above.

---

## 🧠 How it works

VBA's `For Each` requires the iterable object to expose `IEnumVARIANT` via `_NewEnum`. Rather than using a typelib to define `IEnumVARIANT`, this module:

1. Allocates a block of heap memory (`CoTaskMemAlloc`) large enough to hold the enumerator state (`TENUM` UDT)
2. Builds a vtable of function pointers (`AddressOf`) for the seven COM methods: `QueryInterface`, `AddRef`, `Release`, `Next`, `Skip`, `Reset`, `Clone`
3. Writes the vtable pointer into the first field of `TENUM` — making it a valid COM object
4. Overwrites the return value of `Enumerate` with the heap pointer — returning the synthetic object as `IEnumVARIANT`
5. Keeps the iterable object alive via a `Static Collection` keyed by the heap address, compensating for reference count changes when the local `TENUM` goes out of scope

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

`pvTable` is the first field because COM requires a pointer to the vtable at offset 0 of any COM object. `IEnum As IEnumerator` is early-bound — the compiler emits a direct vtable call to `.Item(i)` rather than a late-bound `IDispatch.Invoke`, which is the dominant cost in `IEnumVARIANT_Next`. `Step` is derived from `First`/`Last` at construction time; ascending and descending enumeration share the same code paths throughout.

### vtable — built once per session

```vb
Static vTable(0 To 6) As LongPtr
If vTable(0) = vbNullPtr Then
    vTable(0) = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
    ...
    vTable(6) = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
End If
```

The `Static` array persists for the lifetime of the VBA session. `vTable(0) = vbNullPtr` is the once-only sentinel — subsequent calls to `Enumerate` reuse the same vtable without rebuilding it. Slots 0–2 are the three `IUnknown` methods (required first by all COM interfaces); slots 3–6 are the four `IEnumVARIANT` methods.

### Return value trick

```vb
CopyMemory ByVal VarPtr(Enumerate), MemoryBlock, vbSizeLongPtr
```

`Enumerate` returns `IEnumVARIANT`. VBA stores the return value as an object pointer at `VarPtr(Enumerate)`. Overwriting those bytes with the heap block address makes VBA believe that address is a valid COM object — which it is, because the vtable pointer sits at offset 0. This is why `Enumerate` carries `'@Ignore NonReturningFunction`; the return value is set by raw memory write, not by a `Set Enumerate = ...` assignment.

### `KeepAlive` — reference count management

`CopyMemory ByVal MemoryBlock, obj, LenB(obj)` copies the `TENUM` struct to the heap as raw bytes — it copies the `IEnum` pointer without calling `AddRef`. When `obj` goes out of scope at function exit, VBA calls `Release` on `obj.IEnum`, which could destroy the iterable even though the heap block still holds a raw (untracked) copy of the pointer.

`KeepAlive` compensates:

```vb
Set KeepAlive(MemoryBlock) = obj.IEnum   ' hold one tracked reference
```

A `Static Collection` inside the property holds the reference, keyed by the heap block address. When `IUnknown_Release` sees `nRef = 0`, it removes the entry and frees the block:

```vb
Set KeepAlive(VarPtr(obj)) = Nothing   ' release — iterable may now be destroyed
CoTaskMemFree VarPtr(obj)              ' free heap block
```

The same pattern applies in `IEnumVARIANT_Clone`: the cloned enumerator's `IEnum` reference is also registered.

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

Per iteration: one early-bound `.Item(i)` call and one Variant copy to the destination address. `rgVar` is advanced by `vbSizeVariant` (16 bytes x86 / 24 bytes x64) for multi-element fetches; VBA's `For Each` always requests one item at a time (`celt = 1`), so the inner `Exit For` fires immediately and `rgVar` is never advanced. `obj.Current` is updated in one step after the loop. `pceltFetched` is written only when the caller provides a non-null pointer — the COM spec allows `NULL` when `celt = 1`.

### Variant ByRef construct

```vb
Private Type CONSTRUCT
    vt  As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
```

`InitializeVarByRef` sets `vt` to `VT_INTEGER | VT_BYREF` pointing to `ref`. After this, `CopyVarByRef` writes a Variant to any address without an API call:

```vb
VarByRef.ref = address        ' point vt at the target address
vt = VBA.vbVariant Or VT_BYREF
If VBA.IsObject(value) Then
    Set ref = value           ' VBA writes the object Variant to address
Else
    ref = value               ' VBA writes the value Variant to address
End If
```

VBA performs the Variant copy through the pointer without knowing it is writing to unmanaged memory. `CopyLngByRef` uses the same mechanism (`vbLong | VT_BYREF`) to write a `Long` to a target address — used for updating `pceltFetched`. Both helpers receive `vt` and `ref` as `ByRef Variant` parameters rather than accessing `VarByRef` directly, avoiding a global UDT indirection per call.

### `IEnumVARIANT_Skip` — direction-aware

```vb
Case celt <= (obj.Step * (obj.Last - obj.Current) + 1)
    obj.Current = obj.Current + celt * obj.Step
```

`obj.Step * (obj.Last - obj.Current) + 1` gives the remaining item count for both ascending (`Step = 1`) and descending (`Step = -1`) without a branch. For overshoot, `obj.Current = obj.Last + obj.Step` places current one step past the end — the same post-exhaustion state that `Next` leaves after the sequence is consumed.

### `IEnumVARIANT_Clone` — state snapshot

```vb
Dim Copy As TENUM: Copy = obj      ' UDT copy — VBA AddRefs IEnum
Copy.nRef = 1
```

`Copy = obj` copies all fields including the raw `IEnum` pointer; VBA automatically calls `AddRef` on the embedded object member during the UDT assignment, so `KeepAlive` has exactly the one tracked reference it needs for the clone. The clone starts at `nRef = 1` regardless of the original's count, and captures the enumeration position at the moment of cloning.

### GUID comparison — field-by-field

```vb
' Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
IsIID_IUnknown = (id.Data1 = &H0) And (id.Data2 = &H0) And ... And (id.Data4(7) = &H46)
```

`QueryInterface` checks for `IID_IUnknown` and `IID_IEnumVARIANT`. Comparing the `GUID` UDT field-by-field avoids string parsing and is direct integer comparison. The commented-out `Const` documents the expected GUID string without runtime cost.

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

A `Variant` is 16 bytes on x86 and 24 bytes on x64 due to pointer-size alignment in the data union. All `CopyMemory` sizes and pointer arithmetic use these constants — no hard-coded values appear in the code.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
