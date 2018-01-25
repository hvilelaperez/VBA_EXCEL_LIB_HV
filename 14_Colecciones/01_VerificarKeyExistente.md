# Verificar Key Existente en Collection

Función para verificar si una clave (key) existe dentro de una colección (Collection).

## Código VBA

```vba
'// Verificar la existencia de una clave en una Collection
Public Function KeyExists(myCol As Collection, ByVal keyCheck As String) As Boolean
    KeyExists = False
    On Error GoTo EndFunction
    myCol.Item keyCheck
    KeyExists = True
EndFunction:
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `KeyExists` | Función pública que retorna `True` si la clave existe |
| `myCol` | Parámetro de tipo Collection que representa la colección a verificar |
| `keyCheck` | Cadena con la clave que se desea buscar |
| `On Error GoTo EndFunction` | Si `myCol.Item` falla (clave no existe), salta a `EndFunction` |
| `myCol.Item keyCheck` | Intenta acceder al elemento con la clave especificada |
| `EndFunction` | Etiqueta de salida donde termina la función |

---

*Módulo: 08 - VBA Collections | Archivo: 00 - Kiem tra su ton tai cua key.txt*
