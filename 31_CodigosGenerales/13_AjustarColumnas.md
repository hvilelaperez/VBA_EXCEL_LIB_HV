# Ajustar Columnas Automáticamente

Ajusta el ancho de las columnas automáticamente cuando se modifican datos en la hoja.

## Código VBA

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Columns("F:G").AutoFit
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| Evento | `Worksheet_Change` - Se ejecuta cuando cambia el contenido de una celda |
| Objetivo | Columnas F y G |
| Acción | AutoFit - Ajusta el ancho al contenido |

---

*Módulo: 31_CodigosGenerales | Archivo: 13_AjustarColumnas.md*
*Autor: Henry Vilela - hvilelaperez@gmail.com*
