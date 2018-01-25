# Crear Workbook con Autofilter

Macro para crear múltiples libros de trabajo (workbooks) separados según los valores únicos de una columna usando Autofilter.

## Código VBA

```vba
Sub TAO_FILE_OPTI()
    Dim Dic As Object
    Dim i&, j&, k&, n As Long, t As Single
    Dim rngKey As Range, wb As Workbook, rngi As Range, rngTotal As Range
    Dim arr(), keyDic()
    Dim strFilter As String
    
    FastWB True: t = Timer
    Set Dic = CreateObject("Scripting.Dictionary")
    
    Set rngTotal = Sheet4.Range("A1:F" & Sheet4.[A65536].End(xlUp).Row)
    Set rngKey = Sheet4.Range("F2:F" & Sheet4.[A65536].End(xlUp).Row)
    
    For Each rngi In rngKey
        If Not Dic.Exists(rngi.Value) Then Dic.Add rngi.Value, rngi.Value
    Next rngi
    
    keyDic = Dic.keys
    For i = 0 To UBound(keyDic)
        With rngTotal
            .AutoFilter field:=6, Criteria1:=keyDic(i)
            .Copy
        End With
        Set wb = Workbooks.Add
        With wb
            .Worksheets(1).Range("A2").PasteSpecial xlPasteColumnWidths
            .Worksheets(1).Range("A2").PasteSpecial xlPasteAll
            .Worksheets(1).Range("A2").Select
            .Worksheets(1).Range("A2").Copy
            '.Worksheets(1).Columns("A:F").AutoFit
            .Worksheets(1).Columns("A:F").AutoFit
            .SaveAs ThisWorkbook.Path & "\" & keyDic(i) & ".xlsx"
            .Close
        End With

        rngTotal.AutoFilter
    Next i

    FastWB False: MsgBox "Completado " & Round((Timer - t), 2) & " segundos"
End Sub
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| **Función** | Crea un workbook separado para cada valor único de la columna F |
| **Entrada** | Rango A1:F de la hoja Sheet4 con datos |
| **Salida** | Archivos .xlsx individuales nombrados según el valor de la columna F |
| **Método** | Usa Dictionary para obtener valores únicos y Autofilter para separar datos |

### Pasos del proceso

1. Optimiza el rendimiento con `FastWB`
2. Crea un diccionario con los valores únicos de la columna F
3. Para cada valor único:
   - Aplica Autofilter en la columna F
   - Copia los datos filtrados
   - Crea un nuevo workbook
   - Pega los datos (anchos de columna y valores)
   - Guarda el archivo con el nombre del valor
   - Cierra el workbook
4. Restaura el rendimiento

### Notas

- La columna F (campo 6) es la columna de filtrado
- Los archivos se guardan en la misma ruta que el libro original
- Usa formato .xlsx para los archivos generados

---

*Módulo: 14 - VBA Autofilter | Archivo: 02 - Ung dung Autofilter tao workbook.txt*
