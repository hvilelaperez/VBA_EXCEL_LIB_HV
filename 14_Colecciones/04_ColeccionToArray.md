# Colección a Array

Función para transferir todos los elementos de una Collection a un array bidimensional.

## Código VBA

```vba
'// Transferir los Items de una Collection a un Array (2 dimensiones)
Public Function CollectionToArray(myCol As Collection) As Variant
    Dim arr() As Variant, i As Long
    ReDim arr(1 To myCol.Count, 1 To 1)
    For i = 1 To myCol.Count
        arr(i, 1) = myCol.Item(i)
    Next i
    CollectionToArray = arr
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `CollectionToArray` | Función que retorna un array con los elementos de la colección |
| `myCol` | Colección de origen cuyos elementos se transferirán |
| `arr()` | Array bidimensional dimensionado al tamaño de la colección |
| `ReDim arr(1 To myCol.Count, 1 To 1)` | Redimensiona el array con filas = cantidad de elementos, 1 columna |
| `myCol.Item(i)` | Accede al elemento en la posición `i` de la colección |
| Retorno | El array bidimensional con todos los elementos de la colección |

---

*Módulo: 08 - VBA Collections | Archivo: 03 - Dua cac item trong collection vao array.txt*
