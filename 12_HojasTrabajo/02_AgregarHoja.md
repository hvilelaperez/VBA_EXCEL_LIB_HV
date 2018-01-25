# Agregar Hojas de Trabajo

Procedimiento para crear nuevas hojas de trabajo a partir de los valores de un rango llamado `NameSheet`.

## Código VBA

```vba
Sub ADD_SHEET()
Dim sh As Worksheet
Dim i As Variant
For Each i In Feuil1.Range("NameSheet")
    Set sh = Sheets.Add(After:=Worksheets(Worksheets.Count))
    sh.Name = i.Value
Next

End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `ADD_SHEET` | Procedimiento que agrega hojas nuevas al final del libro |
| `sh` | Variable de tipo Worksheet que representa la hoja recién creada |
| `i` | Variable iteradora que recorre cada celda del rango `NameSheet` |
| `Feuil1.Range("NameSheet")` | Rango en la hoja `Feuil1` que contiene los nombres de las hojas a crear |
| `Sheets.Add(After:=...)` | Agrega la nueva hoja después de la última hoja existente |

---

*Módulo: 06 - VBA Worksheets | Archivo: 01 - Add Sheet.txt*
