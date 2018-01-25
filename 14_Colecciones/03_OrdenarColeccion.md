# Ordenar Colección A-Z

Procedimiento recursivo para ordenar los elementos de una colección de forma ascendente (A-Z) utilizando el algoritmo QuickSort.

## Código VBA

```vba
'// Ordenar los Items de la Collection de A-Z
Public Sub SortingCollection(myCol As Collection, firstIndex As Long, lastIndex As Long)
  Dim valCentre As Variant, vTemp As Variant
  Dim valMin As Long
  Dim valMax As Long
  valMin = firstIndex
  valMax = lastIndex
  valCentre = myCol((firstIndex + lastIndex) \ 2)
  Do While valMin <= valMax
    Do While myCol(valMin) < valCentre And valMin < lastIndex
      valMin = valMin + 1
    Loop
    Do While valCentre < myCol(valMax) And valMax > firstIndex
      valMax = valMax - 1
    Loop
    If valMin <= valMax Then
      ' Intercambiar valores
      vTemp = myCol(valMin)
      myCol.Add myCol(valMax), After:=valMin
      myCol.Remove valMin
      myCol.Add vTemp, before:=valMax
      myCol.Remove valMax + 1
      ' Mover a las siguientes posiciones
      valMin = valMin + 1
      valMax = valMax - 1
    End If
  Loop
  If firstIndex < valMax Then SortingCollection myCol, firstIndex, valMax
  If valMin < lastIndex Then SortingCollection myCol, valMin, lastIndex
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `SortingCollection` | Sub que implementa QuickSort recursivo sobre una Collection |
| `myCol` | Colección de elementos a ordenar |
| `firstIndex` / `lastIndex` | Índices que definen el rango a ordenar |
| `valCentre` | Elemento pivote seleccionado del medio del rango |
| `valMin` / `valMax` | Punteros que se mueven desde los extremos hacia el centro |
| Intercambio | Usa `Add` y `Remove` para intercambiar elementos en la colección |
| Recursión | Se llama a sí misma para ordenar las sub-particiones |

---

*Módulo: 08 - VBA Collections | Archivo: 02 - Sort A-Z key trong collection.txt*
