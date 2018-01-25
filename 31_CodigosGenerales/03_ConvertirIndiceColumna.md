# Convertir Índice de Columna a Nombre

Función que convierte el número índice de una columna en su letra correspondiente (ej: 1 → "A", 28 → "AB").

## Código VBA

```vba
Public Function Coleter(ColIndex As Long) As String
    Coleter = Replace(Cells(1, ColIndex).Address(0, 0), 1, "")
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `Coleter` | Función que recibe el índice numérico de columna y retorna su letra |
| `Cells(1, ColIndex)` | Referencia a la celda en la fila 1 de la columna indicada |
| `.Address(0, 0)` | Obtiene la dirección en formato relativo (sin `$`) |
| `Replace(..., 1, "")` | Elimina el número de fila de la dirección para dejar solo la columna |

**Ejemplos de uso:**

| Entrada | Salida |
|---------|--------|
| `Coleter(1)` | "A" |
| `Coleter(26)` | "Z" |
| `Coleter(27)` | "AA" |
| `Coleter(100)` | "CV" |

---

*Módulo: 31_CodigosGenerales | Archivo: 03_ConvertirIndiceColumna.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
