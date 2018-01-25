# Autofilter vs VBA

Ejemplos de cómo aplicar y eliminar filtros autofilter usando VBA en Excel, incluyendo filtros simples, múltiples columnas y uso de arrays.

## Código VBA

```vba
Sub APPLY_FILTER_DATA()
    Set Sh = Sheet3
    Set rng = Sheet3.Range("Data_filter")
    
    'Filtro por 1 columna
    '
    '************************************************************************************************************
    'rng.AutoFilter field:=2, Criteria1:=Sheet3.Range("D9").Value, Criteria2:=Sheet3.Range("D10"), Operator:=xlOr
    '
    '************************************************************************************************************
    
    
    'Filtro por múltiples columnas
    '
    '************************************************************************************************************
    'rng.AutoFilter field:=2, Criteria1:=Sheet3.Range("D9").Value
    'rng.AutoFilter field:=4, Criteria1:=Sheet3.Range("D10").Value
    'rng.AutoFilter field:=3, Criteria1:=Sheet3.Range("D10").Value & "*"
    '
    '************************************************************************************************************
    
    'Filtro usando array
    rng.AutoFilter field:=2, Criteria1:=Array("West", "East"), Operator:=xlFilterValues
    
End Sub

Sub RESET_AUTOFILTER()
    Sheet3.Range("Data_filter").AutoFilter
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función** | Ejemplos de filtros autofilter con diferentes configuraciones |
| **Filtro simple** | Filtra una columna con dos criterios usando operador OR |
| **Múltiples columnas** | Aplica filtros en varias columnas simultáneamente |
| **Filtro con array** | Usa un array de valores para filtrar múltiples opciones en una columna |
| **Reset** | Elimina todos los filtros aplicados al rango |

### Tipos de filtro

| Método | Descripción |
|--------|-------------|
| `xlOr` | Operador OR para combinar dos criterios en una columna |
| `xlFilterValues` | Filtra usando un array de valores permitidos |
| Comodín `*` | Usa asterisco para coincidencias parciales |

### Uso

- Descomentar la línea del filtro deseado
- El rango "Data_filter" debe contener los datos a filtrar
- Las celdas D9 y D10 contienen los criterios de filtrado

---

*Módulo: 14 - VBA Autofilter | Archivo: 00 - Sd autofilter vs vba.txt*
