# vba-enumerator
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module that adds `For Each` support to any class using early binding and a synthetic `IEnumVARIANT` COM object — no typelib required.

Implements the full `IEnumVARIANT` interface (`Next`, `Skip`, `Reset`, `Clone`) in a standard module using `AddressOf` and a heap-allocated vtable. Items are retrieved one by one via the `IEnumerator` interface, which the iterable class must implement.

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
| `IEnumerator.cls` | Interface | Defines `First`, `Last`, `Item` — implement this in your iterable class |
| `Enumerator.bas` | Module | `Enumerate(iterable)` — the main entry point |
| `EnumTestEarlyClass.cls` | Example | Simple iterable class implementing `IEnumerator` |
| `EnumTest.bas` | Example | `For Each` tests and performance timings |

> **Note:** `EnumTest.bas` uses a `Stopwatch` module for timing measurements. Remove or replace those calls if you do not have it.

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

**1. Implement `IEnumerator` in your class:**

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
| Custom enumerator — `For Each` (VarByRef) | 2.69 |
| Custom enumerator — `For Each` (API) | 15.0 |
| VB Collection — `For Each` | 0.21 |
| VB Array — `For Each` | 0.14 |
| VB Array — `For i` | 0.08 |

The overhead versus a native VB Collection is primarily the early-bound `IEnumerator.Item(i)` call per element. The VarByRef variant copy is ~5× faster than the `VariantCopy` API.

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

## 📄 License

MIT © 2025 Vincent van Geerestein
