# Insertar Fila por Condición

Inserta filas automáticamente para cada código único entre fechas de inicio y fin.

## Código VBA

```vba
Option Explicit

Sub insertarFilaYRellenarFormula()

Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
Dim fechaInicio As Date: fechaInicio = ws.Range("C1").Value
Dim fechaFin As Date: fechaFin = ws.Range("D1").Value

Dim contadorFecha As Date: contadorFecha = fechaInicio
Dim i As Long
Dim UltimaFila As Long
UltimaFila = ws.Range("A" & ws.Rows.Count).End(xlUp).Row

Dim dicMa As Object
Set dicMa = CreateObject("Scripting.Dictionary")
Dim xKey As Variant

For i = 4 To UltimaFila
    xKey = ws.Range("A" & i).Value2
    If Not dicMa.exists(xKey) Then
        dicMa.Add xKey, ""
    End If
Next i

i = 4
Dim k As Variant
For Each k In dicMa.Keys
    Do
        ws.Range("A" & i).Value2 = k
        ws.Range("B" & i).Value = contadorFecha
        contadorFecha = contadorFecha + 1
        i = i + 1
    Loop Until contadorFecha > fechaFin
    contadorFecha = fechaInicio
Next k

UltimaFila = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
ws.Range("C4", "D" & UltimaFila).FillDown

End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `fechaInicio` | Fecha de inicio del rango (celda C1) |
| `fechaFin` | Fecha de fin del rango (celda D1) |
| `Scripting.Dictionary` | Almacena códigos únicos para evitar duplicados |
| `FillDown` | Copia fórmulas hacia abajo |
| Columna A | Código único |
| Columna B | Fecha (itera entre inicio y fin) |

---

*Módulo: 31_CodigosGenerales | Archivo: 21_InsertarFilaCondicion.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
