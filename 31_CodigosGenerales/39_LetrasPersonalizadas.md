# Letras Personalizadas en Excel

Función para convertir un número de columna en su letra correspondiente (A, B, C...).

## Código VBA

```vba
Function ConvertirLetra(k As Long) As String
    Dim vArr As Variant
    vArr = Split(Cells(1, k).Address(True, False), "$")
    ConvertirLetra = vArr(0)
End Function


'O usando la fórmula:
'=CHAR(64+ROW())
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `ConvertirLetra` | Función que convierte un número de columna a su letra |
| `k` | Parámetro: número de columna (Long) |
| `Cells(1, k).Address(True, False)` | Obtiene la dirección de la celda en formato relativo (ej: "$A$1") |
| `Split(..., "$")` | Divide la dirección por el carácter "$" |
| `vArr(0)` | Primera parte de la dirección (la letra de columna) |

### Ejemplos

| Entrada (k) | Salida |
|-------------|--------|
| 1 | A |
| 2 | B |
| 26 | Z |
| 27 | AA |

### Alternativa con fórmula

La fórmula `=CHAR(64+ROW())` también puede usarse:
- `CHAR(65)` = "A" (cuando ROW()=1)
- `CHAR(66)` = "B" (cuando ROW()=2)

> **Nota:** La función VBA es más flexible ya que puede usarse en cualquier celda, mientras que la fórmula depende de la fila actual.

---

*Módulo: 96 - VBA Codes general | Archivo: 27 - Chu cai tu ding trong Excel.txt*
