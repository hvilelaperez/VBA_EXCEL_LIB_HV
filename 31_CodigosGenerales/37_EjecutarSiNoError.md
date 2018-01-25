# Ejecutar Si No Hay Error

Función para seleccionar la primera celda vacía en una columna específica, verificando que la columna exista.

## Código VBA

```vba
Private Function SeleccionarPrimeraFilaVaciaEnColumnaConEncabezado(ByVal hoja As Worksheet, Optional ByVal encabezado As String = "Style") As Long
    Dim col As Variant
    With hoja
        col = Application.Match(encabezado, .Rows(1), 0)
        If Not IsError(col) Then
            .Activate '<--| Debe seleccionar la hoja para activar una celda
            .Cells(.Rows.Count, col).End(xlUp).Offset(1).Select
        End If
    End With
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `SeleccionarPrimeraFilaVaciaEnColumnaConEncabezado` | Función privada que retorna un valor Long |
| `hoja` | Parámetro: hoja de cálculo donde buscar |
| `encabezado` | Parámetro opcional: nombre del encabezado de columna (default: "Style") |
| `Application.Match` | Busca el encabezado en la primera fila de la hoja |
| `IsError` | Verifica si `Match` no encontró el encabezado |
| `End(xlUp)` | Busca la última celda con datos en la columna |
| `Offset(1)` | Se desplaza una fila hacia abajo (primera celda vacía) |

### Lógica

1. Busca el encabezado proporcionado en la primera fila de la hoja
2. Si el encabezado existe (no hay error), activa la hoja
3. Encuentra la última celda con datos en esa columna
4. Selecciona la celda inmediatamente debajo (primera vacía)
5. Si el encabezado no existe, la función termina sin seleccionar nada

### Nota importante

> La hoja debe estar activa para poder seleccionar una celda de ella. Por eso se llama `.Activate` antes de `.Select`.

---

*Módulo: 96 - VBA Codes general | Archivo: 18 - Case khong bi error thi thuc hien lenh.txt*
