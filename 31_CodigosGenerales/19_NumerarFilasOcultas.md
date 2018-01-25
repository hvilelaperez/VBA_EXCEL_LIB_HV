# Numerar Filas con Datos (Ocultar vacías)

Numeración automática de filas y ocultamiento de filas sin datos.

## Código VBA

```vba
Sub AnDongKhiKhongCoDuLieu()
 Dim Filas As Long, Conteo As Integer
 Dim WF As Object, Rng As Range, STT As Long
 
 Set WF = Application.WorksheetFunction
 Filas = [B2].CurrentRegion.Rows.Count
 Rows("1:" & Filas).Hidden = False
 Application.ScreenUpdating = False
 For J = 2 To Filas
    Set Rng = Range(Cells(J, "D"), Cells(J, "Y"))
    Conteo = WF.CountA(Rng)
    If Conteo Then
        STT = STT + 1:              Cells(J, "A").Value = STT
    Else
        Rows(J & ":" & J).Hidden = True
    End If
 Next J
 Application.ScreenUpdating = True
 MsgBox "¡Completado!", , "GPE.COM ¡Bienvenido!"
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `CountA` | Cuenta celdas con contenido en el rango D a Y |
| `Hidden` | Oculta filas sin datos |
| `STT` | Contador de secuencia (número de fila) |
| `ScreenUpdating` | Desactiva actualización de pantalla para optimizar |
| Rango evaluado | Columnas D a Y de cada fila |

---

*Módulo: 31_CodigosGenerales | Archivo: 19_NumerarFilasOcultas.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
