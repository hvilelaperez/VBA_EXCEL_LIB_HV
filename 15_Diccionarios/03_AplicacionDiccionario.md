# Aplicaciones del Diccionario en VBA

Múltiples ejemplos de uso del diccionario VBA para obtener valores únicos y realizar consolidación de datos con suma.

## Código VBA

```vba
Option Explicit

Sub VALORES_UNICOS()
    Dim sh1 As Worksheets, rng As Range
    Dim i As Integer, key As Variant, items As Variant
    
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.CompareMode = vbBinaryCompare

    Set rng = Sheet1.Range("A1:A" & Sheet1.Range("A" & Rows.Count).End(xlUp).Row)

    On Error Resume Next
    For Each key In rng
        Dic.Add key.Value, ""
    Next
    On Error GoTo 0

    Sheet1.Range("C1").Resize(Dic.Count, 1) = Application.Transpose(Dic.Keys)
End Sub


Sub VALORES_UNICOS2()
    Dim sh1 As Worksheets, rng As Range
    Dim i As Integer, key As Variant, items As Variant
    
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.CompareMode = vbBinaryCompare

    Set rng = Sheet1.Range("A1:A" & Sheet1.Range("A" & Rows.Count).End(xlUp).Row)

    For Each key In rng
        Dic.Item(key.Value) = key.Value
    Next key
    items = Dic.items

    Sheet1.Range("D1").Resize(Dic.Count, 1) = Application.Transpose(Dic.Keys)
End Sub



Sub VALORES_UNICOS3()
    Dim sh1 As Worksheets, rng As Range
    Dim i As Integer, key As Variant, items As Variant
    
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.CompareMode = vbBinaryCompare

    Set rng = Sheet1.Range("A1:A" & Sheet1.Range("A" & Rows.Count).End(xlUp).Row)

    On Error Resume Next
    For Each key In rng
        Dic.Add key.Value, ""
    Next
    On Error GoTo 0

    Sheet1.Range("C1").Resize(Dic.Count, 1) = Application.Transpose(Dic.Keys)
End Sub


Sub CONSOLIDAR_CON_DIC()
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dim sh As Worksheet, rng As Range, arr As Variant, key As Variant
    Dim i As Integer, j As Integer
    Dim ultimaFila As Integer
    
    Set sh = Worksheets(Sheet2.Name)
    ultimaFila = sh.Range("A" & Rows.Count).End(xlUp).Row
    Set rng = sh.Range("A2:B" & ultimaFila)
    arr = rng.Value

    For i = 1 To UBound(arr, 1)
        If Not Dic.Exists(arr(i, 1)) Then
            Dic.Add arr(i, 1), arr(i, 2)
        Else
            Dic.Item(arr(i, 1)) = Dic.Item(arr(i, 1)) + arr(i, 2)
        End If
    Next i
    
    Sheet2.Range("C2").Resize(Dic.Count, 1) = Application.Transpose(Dic.Keys)
    Sheet2.Range("D2").Resize(Dic.Count, 1) = Application.Transpose(Dic.items)
End Sub
```

## Descripción

| Sub | Descripción |
|-----|-------------|
| **VALORES_UNICOS** | Obtiene valores únicos usando `Dic.Add` con manejo de errores para duplicados |
| **VALORES_UNICOS2** | Obtiene valores únicos usando `Dic.Item()` que sobrescribe duplicados automáticamente |
| **VALORES_UNICOS3** | Similar al primero, usa `On Error Resume Next` para ignorar duplicados |
| **CONSOLIDAR_CON_DIC** | Consolida datos sumando valores por categoría usando el diccionario |

| Método | Ventaja |
|--------|---------|
| `Dic.Add` + `On Error` | Ignora duplicados silenciosamente |
| `Dic.Item()` | Sobrescribe valores existentes automáticamente |
| **Consolidación** | Acumula sumas por clave, ideal para reportes resumen |

---

*Módulo: 09 - VBA Dictionary | Archivo: 02 - Ung dung Dic*
