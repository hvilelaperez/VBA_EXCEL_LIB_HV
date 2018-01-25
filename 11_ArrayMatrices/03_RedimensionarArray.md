# Redimensionar Array con ReDim

Código VBA para reordenar columnas de un array usando `Application.Index` y un array de índices de columna.

## Código VBA

```vba
Sub RESIZE_ARRAY()
    Dim Array1 As Variant, Array2 As Variant
    Dim ColExport As Variant, ColVar As Variant
    Dim I As Long
    Application.ScreenUpdating = False
    
    Array1 = Sheet2.Range("Array1").Value
    ColExport = [{1,3,5,7,6,2}]
    ReDim ColVar(1 To UBound(Array1, 1), 1 To 1)
    
    For I = LBound(Array1, 1) To UBound(Array1, 1)
        ColVar(I, 1) = I
    Next
    Sheet2.Range("K3").Resize(UBound(Array1, 1), 7).ClearContents
    Sheet2.Range("K3").Resize(UBound(Array1, 1), UBound(ColExport, 1)) = Application.Index(Array1, ColVar, ColExport)
    
    Application.ScreenUpdating = True
End Sub
```

## Descripción

| Variable | Función |
|----------|---------|
| `Array1` | Array bidimensional con datos origen del rango "Array1" |
| `ColExport` | Array de constantes con orden de columnas a exportar: `[1,3,5,7,6,2]` |
| `ColVar` | Array de índices de fila (1 a N) |

### Proceso

1. Lee datos del rango nombrado "Array1" en Sheet2
2. Define el orden de columnas: `ColExport = [{1,3,5,7,6,2}]`
3. Crea array de índices de fila (`ColVar`)
4. Limpia el área destino (K3 con 7 columnas)
5. Usa `Application.Index` para reordenar y extraer columnas
6. Vuelca el resultado en K3

### Ejemplo de reordenación

| Columna origen | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
|----------------|---|---|---|---|---|---|---|
| **Columna destino** | 1 | - | 3 | - | 5 | 7 | 6 |

---

*Módulo: 05 - VBA Array | Archivo: 02 - RESIZE MANG*