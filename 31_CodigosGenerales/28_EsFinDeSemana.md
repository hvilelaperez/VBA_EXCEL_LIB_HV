# Verificar si es Fin de Semana

Función pública que determina si una fecha corresponde a un fin de semana (sábado o domingo).

## Código VBA

```vba
' Autor: Henry Vilela - hvilelaperez@gmail.com
Public Function EsFinDeSemana(FechaEntrada As Date) As Boolean
    Select Case Weekday(FechaEntrada)
        Case vbSaturday, vbSunday
            EsFinDeSemana = True
        Case Else
            EsFinDeSemana = False
    End Select
End Function
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función** | `EsFinDeSemana` - Determina si una fecha es sábado o domingo |
| **Parámetro** | `FechaEntrada` - Fecha a evaluar (tipo Date) |
| **Retorno** | `Boolean` - `True` si es fin de semana, `False` si no |
| **Constantes** | `vbSaturday` y `vbSunday` representan el sábado y domingo respectivamente |
| **Uso** | `=EsFinDeSemana(HOY())` para verificar si la fecha actual es fin de semana |

---

*Módulo: 31_CodigosGenerales | Archivo: 37 - Is weekend.txt*
